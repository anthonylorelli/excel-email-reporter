"""Outlook Draft Manager: populates draft emails directly in Microsoft Outlook via COM or IMAP."""

import imaplib
import os
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class OutlookDraftManager:
    """Manages populating draft emails directly into Outlook."""

    def __init__(self, config: Any):
        self.config = config
        self._com_app = None
        self._imap_client = None
        self._imap_drafts_folder = None

        # IMAP configuration with fallbacks to SMTP credentials
        self.imap_host = os.getenv("IMAP_HOST", getattr(config, "imap_host", "outlook.office365.com"))
        self.imap_port = int(os.getenv("IMAP_PORT", str(getattr(config, "imap_port", 993))))
        self.imap_user = os.getenv("IMAP_USERNAME", os.getenv("SMTP_USERNAME", getattr(config, "smtp_user", "")))
        self.imap_password = os.getenv("IMAP_PASSWORD", os.getenv("SMTP_PASSWORD", getattr(config, "smtp_password", "")))

    def is_com_available(self) -> Tuple[bool, Optional[str]]:
        """Checks if Classic Outlook COM automation (win32com.client) is functional."""
        try:
            import win32com.client  # type: ignore
            app = win32com.client.Dispatch("Outlook.Application")
            _ = app.GetNamespace("MAPI")
            self._com_app = app
            return True, None
        except Exception as ex:
            return False, str(ex)

    def get_or_create_com_app(self):
        """Retrieves or instantiates the Outlook COM application."""
        if self._com_app is None:
            import win32com.client  # type: ignore
            self._com_app = win32com.client.Dispatch("Outlook.Application")
        return self._com_app

    def create_draft_via_com(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> bool:
        """Populates and saves a draft email directly in Outlook's local Drafts folder via COM."""
        outlook = self.get_or_create_com_app()
        # 0 = olMailItem
        mail = outlook.CreateItem(0)

        # Set headers & body
        mail.To = f'"{recipient_name}" <{recipient_email}>' if recipient_name else recipient_email
        mail.Subject = subject
        mail.HTMLBody = html_body

        if attachments:
            for att in attachments:
                if att.exists():
                    mail.Attachments.Add(str(att.resolve()))

        # Save to Outlook's Drafts folder
        mail.Save()
        return True

    def _get_imap_client(self) -> imaplib.IMAP4_SSL:
        """Establishes or returns an authenticated IMAP SSL connection."""
        if self._imap_client is None:
            if not self.imap_user or not self.imap_password:
                raise ValueError("IMAP credentials not configured. Please set IMAP_USERNAME and IMAP_PASSWORD in .env.")
            
            client = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
            client.login(self.imap_user, self.imap_password)
            self._imap_client = client

            # Auto-detect Drafts mailbox name
            self._imap_drafts_folder = self._find_imap_drafts_folder(client)

        return self._imap_client

    def _find_imap_drafts_folder(self, client: imaplib.IMAP4_SSL) -> str:
        """Finds the drafts folder name by searching mailbox flags or common names."""
        status, folder_list = client.list()
        if status != "OK" or not folder_list:
            return "Drafts"

        for f in folder_list:
            decoded = f.decode("utf-8", errors="ignore")
            # Look for standard \Drafts flag
            if "\\Drafts" in decoded:
                # Name is typically the last quoted or unquoted token
                parts = decoded.split(' "/" ')
                if len(parts) > 1:
                    return parts[-1].strip('"').strip()
            # Fallback to case-insensitive name match
            if "draft" in decoded.lower():
                parts = decoded.split(' "/" ')
                if len(parts) > 1:
                    return parts[-1].strip('"').strip()

        return "Drafts"

    def create_draft_via_imap(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        text_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> bool:
        """Appends a draft email to the server's Drafts folder via IMAP SSL (syncs into Outlook)."""
        client = self._get_imap_client()
        folder = self._imap_drafts_folder or "Drafts"

        # Build MIME message
        msg = MIMEMultipart("mixed")
        msg["Subject"] = subject
        sender_header = getattr(self.config, "email_sender", self.imap_user)
        msg["From"] = sender_header
        to_header = f'"{recipient_name}" <{recipient_email}>' if recipient_name else recipient_email
        msg["To"] = to_header
        msg["X-Unsent"] = "1"

        alt_part = MIMEMultipart("alternative")
        alt_part.attach(MIMEText(text_body, "plain", "utf-8"))
        alt_part.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alt_part)

        if attachments:
            import mimetypes
            for filepath in attachments:
                if not filepath.exists():
                    continue
                ctype, encoding = mimetypes.guess_type(str(filepath))
                if ctype is None or encoding is not None:
                    ctype = "application/octet-stream"
                maintype, subtype = ctype.split("/", 1)
                with open(filepath, "rb") as fp:
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(fp.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", "attachment", filename=filepath.name)
                msg.attach(part)

        # Append to Drafts folder with \Draft flag
        internal_date = imaplib.Time2Internaldate(time.time())
        status, _ = client.append(folder, "\\Draft", internal_date, msg.as_bytes())
        if status != "OK":
            # Retry with quoted folder name
            status, _ = client.append(f'"{folder}"', "\\Draft", internal_date, msg.as_bytes())

        return status == "OK"

    def save_draft_to_outlook(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        text_body: str,
        attachments: Optional[List[Path]] = None,
    ) -> Dict[str, Any]:
        """Attempts to save a draft directly into Outlook (trying COM first, then IMAP)."""
        # 1. Try Classic Outlook COM
        com_ok, com_err = self.is_com_available()
        if com_ok:
            try:
                self.create_draft_via_com(
                    recipient_email=recipient_email,
                    recipient_name=recipient_name,
                    subject=subject,
                    html_body=html_body,
                    attachments=attachments,
                )
                return {"success": True, "method": "com", "error": None}
            except Exception as ex:
                com_err = str(ex)

        # 2. Try IMAP Server Drafts Folder (Universal sync to New Outlook / Outlook.com / Exchange)
        if self.imap_user and self.imap_password:
            try:
                self.create_draft_via_imap(
                    recipient_email=recipient_email,
                    recipient_name=recipient_name,
                    subject=subject,
                    html_body=html_body,
                    text_body=text_body,
                    attachments=attachments,
                )
                return {"success": True, "method": "imap", "error": None}
            except Exception as imap_ex:
                return {
                    "success": False,
                    "method": "imap",
                    "error": f"Outlook COM failed ({com_err}) and IMAP failed ({imap_ex})",
                }

        # 3. Neither method available
        return {
            "success": False,
            "method": None,
            "error": (
                f"Classic Outlook COM unavailable ({com_err}). "
                "For New Outlook / IMAP sync, provide IMAP_USERNAME and IMAP_PASSWORD in .env."
            ),
        }

    def close(self):
        """Clean up connections."""
        if self._imap_client:
            try:
                self._imap_client.logout()
            except Exception:
                pass
            self._imap_client = None
