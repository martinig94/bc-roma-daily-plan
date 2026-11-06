"""Send PNG images by email over SMTP."""

import os
import smtplib
from email.message import EmailMessage

REQUIRED_SETTINGS = ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD", "MAIL_FROM")
SSL_PORT = 465
DEFAULT_PORT = 587


class MailConfigError(RuntimeError):
    """Raised when SMTP settings are missing from the environment."""


def send_images(
    images: dict[str, bytes], recipient: str, subject: str, body: str
) -> None:
    """Email PNG images as attachments.

    SMTP settings come from the environment (``SMTP_HOST``, ``SMTP_PORT``,
    ``SMTP_USER``, ``SMTP_PASSWORD``, ``MAIL_FROM``).

    Args:
        images: Mapping of section name to PNG bytes.
        recipient: Destination email address.
        subject: Email subject.
        body: Plain text body.

    Raises:
        MailConfigError: If a required setting is missing.
        smtplib.SMTPException: If the server rejects the message.
    """
    missing = [key for key in REQUIRED_SETTINGS if not os.environ.get(key)]
    if missing:
        raise MailConfigError(f"Missing settings: {', '.join(missing)}")

    message = EmailMessage()
    message["From"] = os.environ["MAIL_FROM"]
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    for name, data in images.items():
        message.add_attachment(
            data, maintype="image", subtype="png", filename=f"{name}.png"
        )

    host = os.environ["SMTP_HOST"]
    port = int(os.environ.get("SMTP_PORT", DEFAULT_PORT))
    user, password = os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"]
    if port == SSL_PORT:
        with smtplib.SMTP_SSL(host, port) as server:
            server.login(user, password)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(message)
