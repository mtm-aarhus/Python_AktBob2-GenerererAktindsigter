"""This module contains a single function for sending emails using the SMTP protocol."""

from typing import Sequence
from email.message import EmailMessage
import smtplib
from io import BytesIO
from dataclasses import dataclass
import mimetypes


@dataclass
class EmailAttachment:
    """A simple dataclass representing an email attachment."""
    file: BytesIO
    file_name: str


def send_email(receiver: str | list[str], sender: str, subject: str, body: str, smtp_server: str, smtp_port: int,
               html_body: bool = False, attachments: Sequence[EmailAttachment] | None = None) -> None:

    msg = EmailMessage()
    msg['to'] = receiver
    msg['from'] = sender
    msg['subject'] = subject

    # Set body
    if html_body:
        msg.set_content("Please enable HTML to view this message.")
        msg.add_alternative(body, subtype='html')
    else:
        msg.set_content(body)

    # Attach files
    if attachments:
        for attachment in attachments:
            mime = mimetypes.guess_type(attachment.file_name)[0]
            main, sub = mime.split("/") if mime else ("application", "octet-stream")
            attachment.file.seek(0)
            msg.add_attachment(attachment.file.read(), maintype=main, subtype=sub, filename=attachment.file_name)

    # Send message
    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
        smtp.starttls()
        smtp.send_message(msg)