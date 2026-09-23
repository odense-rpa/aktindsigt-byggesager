import asyncio
import logging
import os
import re
import sys
from html import escape
from pathlib import Path

from automation_server_client import AutomationServer, Workqueue, WorkItemError, Credential, WorkItemStatus
from geoenviron.client import GeoEnvironClient
from odk_tools.tracking import Tracker
from services.geoenviron_service import GeoEnvironService
from services.mail_service import MailService
from msgraph.generated.models.body_type import BodyType
from typing import Optional

proces_navn = "Aktindsigt Byggesager"

geoenviron: Optional[GeoEnvironClient] = None
geoenviron_service: Optional[GeoEnvironService] = None
mail_service: Optional[MailService] = None


def _is_windows_metadata_entry(path: Path) -> bool:
    name = path.name.lower()
    ignored_names = {
        ".ds_store",
        "desktop.ini",
        "thumbs.db",
        "ehthumbs.db",
        "ehthumbs_vista.db",
        "$recycle.bin",
        "system volume information",
    }
    return name in ignored_names or name.startswith(".~") or name.startswith("~$") or name.startswith("._")


async def populate_queue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    if geoenviron_service is None:
        raise RuntimeError("GeoEnviron service is not initialized.")

    output_mount = Path(geoenviron_service.output_mount)
    if not output_mount.exists():
        logger.warning("Output mount does not exist: %s", output_mount)
        return

    for folder in sorted(output_mount.iterdir(), key=lambda path: path.name):
        if not folder.is_dir():
            continue

        has_visible_content = False
        for child in folder.iterdir():
            if child.name.startswith(".") or _is_windows_metadata_entry(child):
                continue
            has_visible_content = True
            break

        if has_visible_content:
            continue

        eksisterende_kødata = workqueue.get_item_by_reference(folder.name)
        if eksisterende_kødata:
            continue

        workqueue.add_item(
            reference=folder.name,
            data={"sags_id": folder.name},
        )
        logger.info("Queued empty case folder: %s", folder.name)


async def process_workqueue(workqueue: Workqueue) -> None:
    logger = logging.getLogger(__name__)

    if mail_service is None:
        raise RuntimeError("Mail service is not initialized.")

    if geoenviron_service is None:
        raise RuntimeError("GeoEnviron service is not initialized.")

    for item in workqueue:
        with item:
            data = item.data  # Item data deserialized from json as dict
            
            if data.get("kind") != "folder" or not re.match(r"^\d{4}-\d{4}$", data.get("name", "")):
                logger.warning("Skipping item with unexpected kind or name format: %s", data)
                continue
 
            try:
                sag = geoenviron.hent_sager(params={"$top": 1, "$filter": f"case_no eq '{data['name']}'"})
                if not sag:
                    raise WorkItemError(f"Sag med ID {data['name']} blev ikke fundet i GeoEnviron.")

                
                dokumenter = geoenviron_service.hent_dokumenter_og_bilag(sags_id=data["name"])
                konklusion = await geoenviron_service.scan_dokumenter(dokumenter=dokumenter)
                sagsnoter = geoenviron_service.forbered_sagsnotater(
                    dokumenter=dokumenter,
                    sag=sag[0],
                )
                geoenviron_service.saml_pdf_filer(
                    dokumenter=dokumenter,
                    extra_files=sagsnoter,
                    output_mount=geoenviron_service.output_mount,
                    sags_id=data["name"],
                )

                body_parts = [
                    "<html><body>",
                    f"<p>TYRA har klargjort en pakke med filer til sag <strong>{escape(str(data['name']))}</strong>.</p>",
                ]                
                if konklusion:
                    body_parts.extend([
                        "<h3>Fund til manuel vurdering</h3>",
                        konklusion,
                    ])
                else:
                    body_parts.append("<p>Der blev ikke fundet oplysninger, som kræver manuel vurdering.</p>")

                body_parts.append(f"<p>Pakken kan findes på G:\\RPA\\Leverancer\\Aktindsigt Byggesager\\{data['name']}</p>")
                body_parts.append("</body></html>")
                
                await mail_service.send_mail(
                    to_address="byggesag@odense.dk",
                    subject=f"TYRA har klargjort en pakke til sag: {data['name']}",
                    body="".join(body_parts),
                    attachments=[],
                    body_type=BodyType.Html,
                )

                tracker.track_task(proces_navn)
            except WorkItemError as e:
                # A WorkItemError represents a soft error that indicates the item should be passed to manual processing or a business logic fault
                logger.error(f"Error processing item: {data}. Error: {e}")
                item.fail(str(e))


def initialize_sync_services() -> tuple[AutomationServer, Tracker, Credential]:
    """Initialize synchronous services before entering async context."""
    global geoenviron, geoenviron_service

    ats = AutomationServer.from_environment()

    geoenviron_credential = Credential.get_credential("GeoEnviron")
    tracking_credential = Credential.get_credential("Odense SQL Server")
    robot_credential = Credential.get_credential("RoboA")

    if not geoenviron_credential.username or not geoenviron_credential.password:
        raise RuntimeError("GeoEnviron credential is missing username or password.")

    if not tracking_credential.username or not tracking_credential.password:
        raise RuntimeError("Tracking credential is missing username or password.")

    litellm = Credential.get_credential(
        "LiteLLM"
    )

    mount_arg = next((sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == "--mount" and i + 1 < len(sys.argv)), None)
    geoenviron_mount = mount_arg if mount_arg else os.getenv("GEOENVIRON_MOUNT")

    mount_arg = next((sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == "--output-mount" and i + 1 < len(sys.argv)), None)
    output_mount = mount_arg if mount_arg else os.getenv("OUTPUT_MOUNT")

    geoenviron = GeoEnvironClient(
        base_url=geoenviron_credential.data["url"],
        username=geoenviron_credential.username,
        password=geoenviron_credential.password,
    )
    geoenviron_service = GeoEnvironService(geoenviron, litellm, geoenviron_mount, output_mount)

    tracker = Tracker(
        username=tracking_credential.username,
        password=tracking_credential.password,
    )
    
    return ats, tracker, robot_credential


async def main(ats: AutomationServer, tracker: Tracker, robot_credential: Credential) -> None:
    """Main async entry point."""
    logger = logging.getLogger(__name__)
    logger.info("Starting: %s", proces_navn)

    global mail_service

    workqueue = ats.workqueue()

    mail_service = MailService(robot_credential)
    await mail_service.initialize()

    if "--queue" in sys.argv:
        await populate_queue(workqueue)
        return

    await process_workqueue(workqueue)


if __name__ == "__main__":
    ats, tracker, robot_credential = initialize_sync_services()
    asyncio.run(main(ats, tracker, robot_credential))
