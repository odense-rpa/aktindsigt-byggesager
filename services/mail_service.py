"""
Mail service utilities for working with Microsoft Graph API.
"""

import logging

from pathlib import Path
from typing import Optional
from azure.identity import UsernamePasswordCredential
from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
from automation_server_client import Credential

logger = logging.getLogger(__name__)

class MailService:
    """Utility service for email operations using Microsoft Graph API."""

    def __init__(self, credential: Credential):
        """
        Initialize the mail service.

        Args:
            roboc_credential: Automation Server credential for email access
        """
        self.credential_obj = credential
        self.graph_client: Optional[GraphServiceClient] = None
        self.credential: Optional[UsernamePasswordCredential] = None

    @staticmethod
    def _resolve_attachment_content(attachment: dict) -> bytes:
        """Read bytes from attachment payload or file path."""
        content = attachment.get("content")
        if isinstance(content, bytes):
            return content

        file_path = attachment.get("path") or attachment.get("file_path")
        if isinstance(file_path, str) and file_path:
            return Path(file_path).read_bytes()

        raise ValueError(
            "Attachment must include either bytes in 'content' or a file path in 'path'/'file_path'."
        )

    async def initialize(self) -> None:
        """Initialize the Microsoft Graph client."""
        await self._initialize_graph_client()

    async def _initialize_graph_client(self) -> None:
        """Initialize Microsoft Graph client with username/password credential."""

        try:
            # Extract Graph API configuration from credential data
            credential_data = self.credential_obj.data if hasattr(self.credential_obj, 'data') else {}
            
            logger.debug(f"Credential data keys: {credential_data.keys()}")
            logger.debug(f"Username: {self.credential_obj.username}")
            logger.debug(f"Tenant ID: {credential_data.get('tenant_id', 'NOT SET')}")
            logger.debug(f"Client ID: {credential_data.get('client_id', 'NOT SET')}")
            
            # Create credential for delegated authentication
            if not self.credential_obj.password:
                raise ValueError("Mail credential password is missing.")

            username_with_domain = f"{self.credential_obj.username}@odense.dk"
            self.credential = UsernamePasswordCredential(
                client_id=credential_data.get('client_id', ''),
                username=username_with_domain,
                password=self.credential_obj.password,
                tenant_id=credential_data.get('tenant_id', '')
            )

            # Create Graph service client
            scope = credential_data.get('graph_scope', 'https://graph.microsoft.com/.default')
            self.graph_client = GraphServiceClient(
                credentials=self.credential,
                scopes=[scope]
            )

            # Test authentication by getting user info
            await self._test_authentication()

        except Exception as e:
            logger.error(f"Failed to initialize Graph client: {e}")
            raise

    async def _test_authentication(self) -> None:
        """Test authentication by making a simple Graph API call."""
        try:
            if self.graph_client is None:
                raise RuntimeError("Graph client is not initialized.")

            # Get current user to verify authentication works
            user = await self.graph_client.me.get()
            if user and user.display_name:
                logger.debug(
                    f"Authenticated successfully as: {user.display_name} ({user.user_principal_name})")
            else:
                raise Exception(
                    "Authentication test failed - no user data returned")

        except Exception as e:
            logger.error(f"Authentication test failed: {e}")
            raise Exception(
                f"Failed to authenticate with Microsoft Graph: {e}") from e

    async def send_mail(
        self,
        to_address: str,
        subject: str,
        body: str,
        attachments: list[dict],
        body_type: BodyType = BodyType.Text,
    ) -> None:
        """
        Send an email from the authenticated user.

        Args:
            to_address:  Recipient email address.
            subject:     Email subject line.
            body:        Email body content.
            attachments: List of attachment dicts, each with:
                           - "name"         (str)   filename shown to recipient
                           - "content"      (bytes) raw file content
                           - "content_type" (str, optional) MIME type,
                                            defaults to "application/octet-stream"
            body_type:   BodyType.Text (default) or BodyType.Html.
        """
        if self.graph_client is None:
            raise RuntimeError("Graph client is not initialized. Call initialize() first.")

        

        graph_attachments = [
            
            FileAttachment(
                odata_type="#microsoft.graph.fileAttachment",
                name=attachment["name"],
                content_type=attachment.get("content_type", "application/octet-stream"),
                content_bytes=attachment["content"],
            )
            for attachment in attachments
        ]

        message = Message(
            subject=subject,
            body=ItemBody(content_type=body_type, content=body),
            to_recipients=[
                Recipient(email_address=EmailAddress(address=to_address))
            ],
            attachments=graph_attachments,
        )

        request_body = SendMailPostRequestBody(message=message, save_to_sent_items=True)

        await self.graph_client.me.send_mail.post(request_body)
        logger.info("Mail sent to %s (subject: %r, attachments: %d)", to_address, subject, len(attachments))