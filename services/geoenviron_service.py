"""
GeoEnviron service utilities for retrieving case documents.
"""
import zipfile
import httpx
import re

from datetime import datetime, timezone
from html import escape
from io import BytesIO
from automation_server_client import Credential
from geoenviron.client import GeoEnvironClient
from pydantic_ai import Agent
from pydantic import BaseModel
from pydantic_ai.models.openai import OpenAIChatModel
from pydantic_ai.providers.openai import OpenAIProvider
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from typing import Literal

# System prompt for screening byggesagsmateriale for personoplysninger
# and other potentially sensitive content before sharing.
SYSTEM_PROMPT = """
Du er en faglig assistent for byggesager i en dansk kommune.
Din opgave er at identificere oplysninger, der kan vaere personfoelsomme
eller uegnede at dele i en aktindsigt uden manuel vurdering.

Vurder altid dokumentindhold ud fra et forsigtighedsprincip.

Marker kun oplysninger, der reelt kan vaere foelsomme eller kraever manuel vurdering:
1) CPR-numre (hele eller delvise), sundhedskortnumre og lignende entydige ID-oplysninger.
2) Helbredsoplysninger, fx sygdom, diagnoser, behandlingsforloeb, psykiske forhold
    eller beskrivelser af borgere, der er syge pga. skimmelsvamp.
3) Oplysninger om sociale/private forhold, fx familieforhold, økonomiske vanskeligheder,
    saarbare livsomstaendigheder eller andre tydeligt private detaljer.
4) De foelgende ord, naar de optraeder som selvstaendige ord eller tydelige begreber
    i dokumentet. Et tilfældigt delord eller et almindeligt navn er ikke nok:
    Anonym
    Børneværn 
    Børneværnet
    CPR
    Fortroligt
    Fødselsdato
    Fødselsnummer
    Indeklimaproblemer
    Pengeautomat
    Personnr.
    Personnummer
    Skifteretsattest
    Sygdom
    Skimmelsvamp
    Syg
    Politi

Se altid bort fra følgende, medmindre de indgår i en konkret foelsom sammenhaeng:
1) Almindelige navne på ansøgere, ejere, rådgivere, entreprenører og kontaktpersoner.
2) Adresser, telefonnumre, e-mailadresser og CVR-numre. Disse oplysninger er forventelige
    i byggesager og er ikke i sig selv sensitive.
3) Projektnavne, tegningsnavne, sagsnumre og almindelige oplysninger om en byggesag.

Et dokument skal derfor vurderes som ikke-sensitivt, hvis det kun indeholder almindelige
navne, adresser, telefonnumre, e-mailadresser eller CVR-numre uden helbredsoplysninger,
sociale/private forhold, CPR-oplysninger eller en tydelig manuel vurderingsgrund.


Svar kun i det aftalte outputformat med felterne:
- reasoning: kort begrundelse for vurderingen.
- sensitive_information: konkret opsummering af de fundne sensitive oplysninger.
    Hvis der ikke findes sensitive oplysninger, skal denne vaere praecist "" (tom streng).
- confidence: high, medium eller low.
- needs_human_review: true ved tvivl, ved helbredsoplysninger, ved CPR-relaterede fund,
  eller hvis konteksten ikke er tydelig.

Hvis du er i tvivl, saet needs_human_review til true.
"""

class Verdict(BaseModel):
    reasoning: str
    sensitive_information: str = ""
    confidence: Literal["high", "medium", "low"]
    needs_human_review: bool

class GeoEnvironService:
    """Service wrapper around GeoEnviron document operations."""

    def __init__(self, client: GeoEnvironClient, litellm: Credential, geoenviron_mount: str, output_mount: str):
        self.client = client
        self.litellm = litellm
        self.agent = self._build_agent(self.litellm.data["api"], self.litellm.password)
        self.geoenviron_mount = geoenviron_mount
        self.output_mount = output_mount

    @staticmethod
    def _read_pdf_text(file_path: Path) -> str:
        """Extract text from a PDF file."""
        try:
            from pypdf import PdfReader
        except ImportError:
            return ""

        try:
            reader = PdfReader(str(file_path))
        except (OSError, ValueError):
            return ""

        values: list[str] = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            page_text = page_text.strip()
            if page_text:
                values.append(page_text)

        return " ".join(values)
    

    @staticmethod
    def _read_word_text(file_path: Path) -> str:
        """Extract text from Word documents (.docx, with fallback for .doc)."""
        if file_path.suffix.lower() == ".doc":
            return file_path.read_text(encoding="utf-8", errors="ignore")

        try:
            from docx import Document
        except ImportError:
            return ""

        try:
            document = Document(str(file_path))
        except (OSError, ValueError, zipfile.BadZipFile):
            return ""

        values: list[str] = []
        for paragraph in document.paragraphs:
            paragraph_text = paragraph.text.strip()
            if paragraph_text:
                values.append(paragraph_text)

        return " ".join(values)
    

    @staticmethod
    def _read_excel_text(file_path: Path) -> str:
        """Extract cell text from an Excel file using EPPlus-like traversal."""
        try:
            from openpyxl import load_workbook
        except ImportError:
            return ""

        try:
            workbook = load_workbook(filename=file_path, read_only=True, data_only=True)
        except (OSError, ValueError, zipfile.BadZipFile):
            return ""

        values: list[str] = []
        try:
            for worksheet in workbook.worksheets:
                # Mirror EPPlus behavior: skip worksheets in page layout view.
                if getattr(getattr(worksheet, "sheet_view", None), "view", None) == "pageLayout":
                    continue

                max_row = worksheet.max_row or 0
                max_col = worksheet.max_column or 0
                for row in worksheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                    for cell in row:
                        if cell.value is None:
                            continue

                        cell_text = str(cell.value).strip()
                        if cell_text:
                            values.append(cell_text)
        finally:
            workbook.close()

        return " ".join(values)
    

    @staticmethod
    def _read_html_text(file_path: Path) -> str:
        """Extract text from an HTML file."""
        try:
            from html.parser import HTMLParser
        except ImportError:
            return ""

        class TextExtractor(HTMLParser):
            def __init__(self):
                super().__init__()
                self.text_parts: list[str] = []
                self.skip_script = False
                self.skip_style = False

            def handle_starttag(self, tag: str, attrs: list) -> None:
                if tag in ("script", "style"):
                    if tag == "script":
                        self.skip_script = True
                    else:
                        self.skip_style = True

            def handle_endtag(self, tag: str) -> None:
                if tag == "script":
                    self.skip_script = False
                elif tag == "style":
                    self.skip_style = False

            def handle_data(self, data: str) -> None:
                if not self.skip_script and not self.skip_style:
                    text = data.strip()
                    if text:
                        self.text_parts.append(text)

        try:
            content = file_path.read_text(encoding="utf-8", errors="ignore")
            parser = TextExtractor()
            parser.feed(content)
            return " ".join(parser.text_parts)
        except (OSError, ValueError):
            return ""

    @staticmethod
    def _read_email_text(file_path: Path) -> str:
        """Extract text from RFC email (.eml) and Outlook MSG files."""
        if file_path.suffix.lower() == ".msg":
            try:
                import extract_msg

                message = extract_msg.Message(str(file_path))
                text_parts = [
                    str(value).strip()
                    for value in (message.subject, message.body)
                    if value
                ]
                message.close()
                return " ".join(text_parts)
            except (OSError, ValueError):
                return ""

        try:
            from email import message_from_file, policy
        except ImportError:
            return ""

        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                msg = message_from_file(f, policy=policy.default)

            text_parts: list[str] = []
            
            # Extract subject
            if "subject" in msg:
                subject = msg["subject"]
                if subject:
                    text_parts.append(str(subject).strip())
            
            # Extract body
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            text = payload.decode(encoding="utf-8", errors="ignore").strip()
                            if text:
                                text_parts.append(text)
                    elif content_type == "text/html":
                        # Extract text from HTML part if no plain text available
                        payload = part.get_payload(decode=True)
                        if payload and not any(p.startswith("text") for p in text_parts):
                            html_text = payload.decode(encoding="utf-8", errors="ignore")
                            # Simple HTML strip
                            import re
                            text = re.sub(r"<[^>]+>", "", html_text).strip()
                            if text:
                                text_parts.append(text)
            else:
                # Single-part message
                payload = msg.get_payload()
                if payload:
                    text = str(payload).strip()
                    if text:
                        text_parts.append(text)

            return " ".join(text_parts)
        except (OSError, ValueError):
            return ""

    @staticmethod
    def _parse_reg_date(value: object) -> datetime:
        if isinstance(value, datetime):
            return value if value.tzinfo is not None else value.replace(tzinfo=timezone.utc)

        if isinstance(value, str):
            try:
                parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
                return parsed if parsed.tzinfo is not None else parsed.replace(tzinfo=timezone.utc)
            except ValueError:
                return datetime.min.replace(tzinfo=timezone.utc)

        return datetime.min.replace(tzinfo=timezone.utc)
    

    @staticmethod
    def _document_path(doc_path: str, geoenviron_mount: str) -> Path:
        normalized_path = doc_path.replace("\\", "/")
        relative_path = re.sub(
            r"^//[^/]+/GeoEnviron/",
            "",
            normalized_path,
            count=1,
            flags=re.IGNORECASE,
        )

        return Path(geoenviron_mount) / relative_path.lstrip("/")

    @staticmethod
    def _record_path(self, document: dict) -> Path|None:
        path_value = document.get("link_path") or document.get("doc_path")
        if not path_value:
            return None
        mount_path = str(self.geoenviron_mount)
        return self._document_path(path_value, mount_path)
    

    @staticmethod
    def _iter_documents_with_bilag(dokumenter: list[dict]):
        for dokument in dokumenter:
            yield dokument
            yield from dokument.get("bilag", [])


    @staticmethod
    def _zip_filename(description: object, suffix: str, used_names: set[str]) -> str:
        filename = str(description or "document").strip()
        filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename)
        filename = filename.rstrip(" .") or "document"

        if filename.upper().split(".", 1)[0] in {
            "CON", "PRN", "AUX", "NUL",
            *(f"COM{number}" for number in range(1, 10)),
            *(f"LPT{number}" for number in range(1, 10)),
        }:
            filename = f"_{filename}"

        if suffix and not filename.lower().endswith(suffix.lower()):
            filename = f"{filename}{suffix}"

        filename = filename[:240].rstrip(" .") or "document"
        candidate = filename
        counter = 2
        while candidate.casefold() in used_names:
            stem, extension = Path(filename).stem, Path(filename).suffix
            candidate = f"{stem} ({counter}){extension}"
            counter += 1
        used_names.add(candidate.casefold())
        return candidate
    

    @staticmethod
    def _document_identifier(document: dict) -> str | None:
        for key in ("doc_id", "doc_seq_no", "doc_no", "link_id", "link_seq_no"):
            value = document.get(key)
            if value not in (None, ""):
                return str(value).strip()
        return None
    

    @staticmethod
    def _document_title_with_id(self, document: dict, default: str = "Dokument") -> str:
        title = str(document.get("description") or document.get("comments") or default).strip() or default
        identifier = self._document_identifier(document)
        return f"{title} ({identifier})" if identifier else title
    

    @staticmethod
    def _build_agent(base_url: str, api_key: str, system_prompt: str = SYSTEM_PROMPT) -> Agent[None, Verdict]:
        http_client = httpx.AsyncClient(verify=False)
        provider = OpenAIProvider(base_url=base_url, api_key=api_key, http_client=http_client)
        model = OpenAIChatModel("odense-large", provider=provider)    
        return Agent(model, output_type=Verdict, system_prompt=system_prompt)


    @staticmethod
    def _text_to_pdf_bytes(title: str, text: str) -> bytes | None:
        text = text.strip()
        if not text:
            return None

        styles = getSampleStyleSheet()
        title_style = styles["Heading3"]
        body_style = styles["BodyText"]
        body_style.leading = 14

        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            lines = [text]

        story = [Paragraph(escape(title), title_style), Spacer(1, 4 * mm)]
        for line in lines:
            story.append(Paragraph(escape(line), body_style))
            story.append(Spacer(1, 2 * mm))

        pdf_buffer = BytesIO()
        SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4,
            rightMargin=18 * mm,
            leftMargin=18 * mm,
            topMargin=18 * mm,
            bottomMargin=18 * mm,
        ).build(story)
        return pdf_buffer.getvalue()
    

    @staticmethod
    def _merge_pdf_bytes(pdf_files: list[bytes]) -> bytes:
        from pypdf import PdfReader, PdfWriter

        writer = PdfWriter()
        for pdf_content in pdf_files:
            reader = PdfReader(BytesIO(pdf_content))
            for page in reader.pages:
                writer.add_page(page)

        output = BytesIO()
        writer.write(output)
        return output.getvalue()

    @staticmethod
    def _convert_docx_to_pdf_bytes(file_path: Path) -> bytes | None:
        if file_path.suffix.lower() not in {".docx", ".doc"}:
            return None

        try:
            with file_path.open("rb") as file_handle:
                response = httpx.post(
                    "http://rpa-ats.odknet.dk:8331/convert",
                    files={"file": (file_path.name, file_handle)},
                    timeout=120,
                )
                response.raise_for_status()
        except Exception:
            return None

        return response.content

    def hent_dokumenter_og_bilag(
        self,
        sags_id: str,        
    ) -> list[dict]:
        sag = self.client.hent_sager(
            params={
                "$top": 1,
                "$filter": f"case_no eq '{sags_id}'",
            }
        )

        sag = sag[0] if sag else None

        if not sag:
            return []

        return self.client.hent_dokumenter(sag, inkluder_bilag=True)


    async def scan_dokumenter(self, dokumenter: list[dict]) -> str:
        findings: list[str] = []

        for dokument in self._iter_documents_with_bilag(dokumenter):
            finding = await self.scan_document(dokument)
            if finding:
                findings.append(finding)

        if not findings:
            return ""

        return "<ul>" + "".join(f"<li>{finding}</li>" for finding in findings) + "</ul>"

    async def scan_document(self, dokument: dict) -> str:
        text = ""
        file_path = self._record_path(self, dokument)

        if file_path is None:            
            text = dokument.get("description", "")
        elif file_path.exists():        
            if dokument["File@odata.mediaContentType"] == "application/pdf":
                text = self._read_pdf_text(file_path)
            elif dokument["File@odata.mediaContentType"] == "application/msword" or dokument["File@odata.mediaContentType"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                text = self._read_word_text(file_path)
            elif dokument["File@odata.mediaContentType"] == "application/msexcel" or dokument["File@odata.mediaContentType"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":                                    
                text = self._read_excel_text(file_path)                        
            elif dokument["File@odata.mediaContentType"] == "text/html":
                text = self._read_html_text(file_path)
            elif (
                dokument["File@odata.mediaContentType"] in ["message/rfc822", "application/vnd.ms-outlook"]
                or file_path.suffix.lower() == ".msg"
            ):
                text = self._read_email_text(file_path)
            elif dokument["File@odata.mediaContentType"] == "text/plain" and dokument["doc_type_name"].lower() not in ["notat"]:                                    
                text = file_path.read_text(encoding="utf-8", errors="ignore")
            elif dokument["File@odata.mediaContentType"] == "text/plain" and dokument["doc_type_name"].lower() in ["notat", "dpost ind", "dpost ud"]:
                text =  dokument["description"]                

        if text:
            try:
                result = await self.agent.run(text)
                if result.output.sensitive_information:
                    description = escape(self._document_title_with_id(self, dokument))
                    information = escape(result.output.sensitive_information)                    
                    return f"<strong>{description}</strong>: {information}"
            except Exception as e:
                # Handle any exceptions that occur during scanning
                return f"<strong>Kunne ikke scanne dokument {escape(self._document_title_with_id(self, dokument))}:</strong>"

        return ""
    

    def forbered_sagsnotater(
        self,
        dokumenter: list[dict],
        sag: dict | None = None,
    ) -> list[dict[str, str | bytes]]:
        """Create in-memory PDF summaries for the supported note document types."""
        dokumenttyper = [
            {
                "doc_type": "Notat",
                "Titel i skabelon": "Sagsnoter",
                "Filnavn": "Sagsnoter.pdf"
            },
            {
                "doc_type": "DPost ind",
                "Titel i skabelon": "Indgående Digital Post noter",
                "Filnavn": "Indgående Digital Post noter.pdf"
            },
            {
                "doc_type": "DPost ud",
                "Titel i skabelon": "Udgående Digital Post noter",
                "Filnavn": "Udgående Digital Post noter.pdf"
            }
        ]

        styles = getSampleStyleSheet()
        title_style = styles["Heading3"]
        body_style = styles["BodyText"]
        body_style.leading = 14

        all_documents = [
            document
            for dokument in dokumenter
            for document in [dokument, *dokument.get("bilag", [])]
        ]

        def value(*keys: str) -> str:
            sources = [sag or {}, *(dokumenter[:1] or [{}])]
            for source in sources:
                for key in keys:
                    candidate = source.get(key)
                    if candidate not in (None, ""):
                        return str(candidate)
            return ""

        case_name = value("case_name", "case_title", "Sagstitel")
        address = value("site_address", "Adresse")
        postal_code = value("postal_codes_id", "postal_code")
        postal_name = value("postal_codes_name", "city", "postal_city")
        case_id = value("case_no", "sags_id", "task_seq_no")
        address_parts = [part for part in [address, postal_code, postal_name] if part]
        case_address = ", ".join(address_parts[:1])
        if len(address_parts) > 1:
            case_address = f"{case_address}, {' '.join(address_parts[1:])}"

        generated_files: list[dict[str, str | bytes]] = []
        for dokumenttype in dokumenttyper:
            note_type = dokumenttype["doc_type"].lower()
            notes = [
                document
                for document in all_documents
                if document.get("File@odata.mediaContentType") == "text/plain"
                and str(document.get("doc_type_name", "")).lower() == note_type
            ]
            if not notes:
                continue

            story = [
                Paragraph("Sagsinfo", title_style),
                Paragraph(f"Sagsnavn: {escape(case_name)}<br/>"
                          f"Adresse: {escape(case_address)}<br/>"
                          f"Id: {escape(case_id)}", body_style),
                Spacer(1, 8 * mm),
                Paragraph(escape(dokumenttype["Titel i skabelon"]), title_style),
            ]

            for index, document in enumerate(notes, start=1):
                document_date = self._parse_reg_date(document.get("doc_date"))
                registration_date = self._parse_reg_date(document.get("reg_date"))
                selected_date = document_date if document_date != datetime.min.replace(tzinfo=timezone.utc) else registration_date
                story.extend([
                    Spacer(1, 5 * mm),
                    Paragraph(f"Note {index}", title_style),
                    Paragraph(
                        f"Dokumentdato: {selected_date.strftime('%d-%m-%Y')}<br/>"
                        f"Registreringsdato: {registration_date.strftime('%d-%m-%Y')}<br/>"
                        f"Dokumenttype: {escape(str(document.get('doc_type_name', '')))}<br/>"
                        f"Beskrivelse: {escape(str(document.get('description', '')))}",
                        body_style,
                    ),
                ])

            pdf_buffer = BytesIO()
            SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=18 * mm,
                leftMargin=18 * mm,
                topMargin=18 * mm,
                bottomMargin=18 * mm,
            ).build(story)
            generated_files.append({
                "name": dokumenttype["Filnavn"],
                "content": pdf_buffer.getvalue(),
                "content_type": "application/pdf",
            })

        return generated_files




    def saml_pdf_filer(
        self,
        dokumenter: list[dict],        
        output_mount: str | Path,
        sags_id: str,
        extra_files: list[dict[str, str | bytes]] | None = None,
    ) -> list[Path]:
        """Create a raw-file folder and merged PDF bundles for the case."""
        output_path = Path(output_mount) / sags_id
        output_path.mkdir(parents=True, exist_ok=True)

        raw_files_path = output_path / "Rå filer"
        raw_files_path.mkdir(parents=True, exist_ok=True)

        pdf_files_path = output_path / "Pdf filer"
        pdf_files_path.mkdir(parents=True, exist_ok=True)

        created_paths: list[Path] = []
        raw_used_names: set[str] = set()
        pdf_used_names: set[str] = set()
        max_size = 60 * 1024 * 1024

        def add_raw_copy(file_path: Path, title: str, suffix: str) -> None:
            raw_name = self._zip_filename(title, suffix, raw_used_names)
            destination = raw_files_path / raw_name
            destination.write_bytes(file_path.read_bytes())
            created_paths.append(destination)

        for dokument in dokumenter:
            documents = [dokument, *dokument.get("bilag", [])]
            for document in documents:
                file_path = self._record_path(self, document)
                if file_path is None or not file_path.exists():
                    continue

                title = self._document_title_with_id(self, document, default=file_path.stem)
                add_raw_copy(file_path, title, file_path.suffix)

        pdf_entries: list[dict[str, object]] = []
        for dokument in dokumenter:
            documents = [dokument, *dokument.get("bilag", [])]
            for document in documents:
                file_path = self._record_path(self, document)
                if file_path is None or not file_path.exists():
                    continue

                media_type = str(document.get("File@odata.mediaContentType") or "")
                suffix = file_path.suffix.lower()
                title = self._document_title_with_id(self, document, default=file_path.stem)

                pdf_bytes: bytes | None = None
                if media_type == "application/pdf" or suffix == ".pdf":
                    pdf_bytes = file_path.read_bytes()
                    if len(pdf_bytes) > max_size:
                        pdf_name = self._zip_filename(title, file_path.suffix, pdf_used_names)
                        destination = pdf_files_path / pdf_name
                        destination.write_bytes(pdf_bytes)
                        created_paths.append(destination)
                        continue
                elif media_type == "text/plain" or suffix == ".txt":
                    text = file_path.read_text(encoding="utf-8", errors="ignore")
                    pdf_bytes = self._text_to_pdf_bytes(title, text)
                elif media_type == "application/msword" or media_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    converted_pdf = self._convert_docx_to_pdf_bytes(file_path)
                    if converted_pdf is not None:
                        pdf_bytes = converted_pdf
                    else:
                        pdf_bytes = self._text_to_pdf_bytes(title, self._read_word_text(file_path))
                elif media_type == "application/msexcel" or media_type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                    pdf_bytes = self._text_to_pdf_bytes(title, self._read_excel_text(file_path))
                elif media_type == "text/html":
                    pdf_bytes = self._text_to_pdf_bytes(title, self._read_html_text(file_path))
                elif media_type in ["message/rfc822", "application/vnd.ms-outlook"] or suffix == ".msg":
                    pdf_bytes = self._text_to_pdf_bytes(title, self._read_email_text(file_path))

                if pdf_bytes is None:
                    continue

                document_date = self._parse_reg_date(document.get("doc_date"))
                if document_date == datetime.min.replace(tzinfo=timezone.utc):
                    document_date = self._parse_reg_date(document.get("reg_date"))
                pdf_entries.append({
                    "date": document_date,
                    "title": title,
                    "content": pdf_bytes,
                })

        for extra_file in extra_files or []:
            content = extra_file.get("content")
            name = extra_file.get("name")
            if not isinstance(content, (bytes, bytearray)) or not isinstance(name, str):
                continue

            extra_content = bytes(content)
            name_path = Path(name)
            if name_path.suffix.lower() == ".pdf" and len(extra_content) > max_size:
                pdf_name = self._zip_filename(name_path.stem, name_path.suffix, pdf_used_names)
                destination = pdf_files_path / pdf_name
                destination.write_bytes(extra_content)
                created_paths.append(destination)
                continue

            if name_path.suffix.lower() == ".pdf":
                pdf_entries.append({
                    "date": datetime.max.replace(tzinfo=timezone.utc),
                    "title": name_path.stem,
                    "content": extra_content,
                })
            else:
                raw_name = self._zip_filename(name_path.stem, name_path.suffix, raw_used_names)
                raw_path = output_path / raw_name
                raw_path.write_bytes(extra_content)
                created_paths.append(raw_path)

        pdf_entries.sort(key=lambda item: item["date"])

        def flush_chunk(chunk: list[dict[str, object]], chunk_number: int) -> None:
            if not chunk:
                return

            writer = __import__('pypdf').PdfWriter()

            for entry in chunk:
                content = entry["content"]
                source_reader = __import__('pypdf').PdfReader(BytesIO(content))
                for page in source_reader.pages:
                    writer.add_page(page)

            chunk_name = f"{sags_id}_{chunk_number}.pdf"
            destination = pdf_files_path / chunk_name
            destination.write_bytes(_render_pdf_bytes(writer))
            created_paths.append(destination)

        def _render_pdf_bytes(writer) -> bytes:
            output = BytesIO()
            writer.write(output)
            return output.getvalue()

        current_chunk: list[dict[str, object]] = []
        current_chunk_size = 0
        chunk_counter = 1

        for entry in pdf_entries:
            content = entry["content"]
            if not isinstance(content, (bytes, bytearray)):
                continue
            candidate_size = current_chunk_size + len(content)
            if current_chunk and candidate_size > max_size:
                flush_chunk(current_chunk, chunk_counter)
                chunk_counter += 1
                current_chunk = []
                current_chunk_size = 0

            if not current_chunk and len(content) > max_size:
                pdf_name = self._zip_filename(str(entry["title"]), ".pdf", pdf_used_names)
                chunk_path = pdf_files_path / pdf_name
                chunk_path.write_bytes(content)
                created_paths.append(chunk_path)
                continue

            current_chunk.append(entry)
            current_chunk_size += len(content)

        if current_chunk:
            flush_chunk(current_chunk, chunk_counter)

        return created_paths