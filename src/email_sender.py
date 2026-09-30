"""Email sender module supporting SMTP, MIME multipart, and file attachments."""

import mimetypes
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, List, Optional


class EmailSender:
    """Handles SMTP connection, email assembly, and sending."""

    def __init__(self, config: Any):
        self.config = config

    def build_message(
        self,
        subject: str,
        html_body: str,
        text_body: str,
        recipients: List[str],
        attachments: Optional[List[Path]] = None,
    ) -> MIMEMultipart:
        """Constructs a MIME multi-part email with both HTML and plain-text alternatives, plus attachments."""
        msg = MIMEMultipart("mixed")
        msg["Subject"] = subject
        msg["From"] = self.config.email_sender
        msg["To"] = ", ".join(recipients)

        if self.config.email_cc:
            msg["Cc"] = ", ".join(self.config.email_cc)

        # Alternative body part (text + html)
        alt_part = MIMEMultipart("alternative")
        alt_part.attach(MIMEText(text_body, "plain", "utf-8"))
        alt_part.attach(MIMEText(html_body, "html", "utf-8"))
        msg.attach(alt_part)

        # Attachments
        if attachments:
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
                part.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=filepath.name,
                )
                msg.attach(part)

        return msg

    def create_draft_eml(
        self,
        recipient_email: str,
        recipient_name: str,
        subject: str,
        html_body: str,
        text_body: str,
        output_eml_path: Path,
        attachments: Optional[List[Path]] = None,
    ) -> Path:
        """Saves an email draft as an RFC 822 .eml file with X-Unsent: 1 header.
        
        Double-clicking this .eml file on Windows will open it in Outlook, Windows Mail,
        or Thunderbird directly in compose/draft mode ready for review and sending.
        """
        output_eml_path.parent.mkdir(parents=True, exist_ok=True)
        
        to_header = f'"{recipient_name}" <{recipient_email}>' if recipient_name else recipient_email
        msg = self.build_message(
            subject=subject,
            html_body=html_body,
            text_body=text_body,
            recipients=[to_header],
            attachments=attachments,
        )
        
        # Mark as draft for Outlook / Windows Mail / Thunderbird
        msg["X-Unsent"] = "1"
        
        with open(output_eml_path, "wb") as f:
            f.write(msg.as_bytes())
            
        return output_eml_path


    def send(
        self,
        subject: str,
        html_body: str,
        text_body: str,
        recipients: Optional[List[str]] = None,
        attachments: Optional[List[Path]] = None,
        dry_run: bool = False,
    ) -> bool:
        """Sends the email or performs a dry-run output preview.
        
        Args:
            subject: Email subject line.
            html_body: Rendered HTML body.
            text_body: Rendered plain-text body.
            recipients: Optional override of recipient list.
            attachments: List of file paths to attach.
            dry_run: If True, previews email without sending over network.
            
        Returns:
            True if successful.
        """
        target_recipients = recipients or self.config.email_recipients
        if not target_recipients and not dry_run:
            raise ValueError("No recipients configured. Set EMAIL_RECIPIENTS in .env or config.yaml.")

        all_recipients = list(target_recipients)
        if self.config.email_cc:
            all_recipients.extend(self.config.email_cc)
        if self.config.email_bcc:
            all_recipients.extend(self.config.email_bcc)

        if dry_run:
            print("\n[DRY RUN] Simulating email delivery:")
            print(f"  From:        {self.config.email_sender}")
            print(f"  To:          {', '.join(target_recipients)}")
            if self.config.email_cc:
                print(f"  Cc:          {', '.join(self.config.email_cc)}")
            if self.config.email_bcc:
                print(f"  Bcc:         {', '.join(self.config.email_bcc)}")
            print(f"  Subject:     {subject}")
            print(f"  Attachments: {[p.name for p in attachments] if attachments else 'None'}")
            print(f"  HTML Size:   {len(html_body)} bytes")
            print(f"  Text Size:   {len(text_body)} bytes\n")
            return True

        # Send via SMTP
        msg = self.build_message(
            subject=subject,
            html_body=html_body,
            text_body=text_body,
            recipients=target_recipients,
            attachments=attachments,
        )

        host = self.config.smtp_host
        port = self.config.smtp_port
        user = self.config.smtp_user
        password = self.config.smtp_password

        if self.config.smtp_use_ssl:
            server = smtplib.SMTP_SSL(host, port)
        else:
            server = smtplib.SMTP(host, port)
            if self.config.smtp_use_tls:
                server.starttls()

        try:
            if user and password:
                server.login(user, password)
            server.sendmail(self.config.email_sender, all_recipients, msg.as_string())
            print(f"Email successfully sent to {', '.join(target_recipients)} via {host}:{port}")
            return True
        finally:
            server.quit()
