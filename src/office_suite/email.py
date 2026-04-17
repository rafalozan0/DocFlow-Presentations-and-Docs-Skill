import os
import smtplib
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any, Dict, List, Optional


EMAIL_PASSWORD_ENV = "OFFICE_EMAIL_PASSWORD"


def _resolve_password(password: Optional[str]) -> str:
    if password:
        return password
    env_password = os.getenv(EMAIL_PASSWORD_ENV)
    if env_password:
        return env_password
    raise RuntimeError(
        f"Email password not provided. Pass `password=...` or set {EMAIL_PASSWORD_ENV} environment variable."
    )


def send_email(
    to: List[str],
    subject: str,
    body: str,
    smtp_server: str,
    smtp_port: int,
    username: str,
    password: Optional[str] = None,
    attachments: Optional[List[str]] = None,
    cc: Optional[List[str]] = None,
    bcc: Optional[List[str]] = None,
    content_type: str = "plain",
    display_name: Optional[str] = None,
    **kwargs,
) -> Dict[str, Any]:
    """Send an email with optional attachments."""
    _ = kwargs
    resolved_password = _resolve_password(password)

    msg = MIMEMultipart()

    if display_name:
        msg["From"] = f"{Header(display_name).encode()} <{username}>"
    else:
        msg["From"] = username

    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if bcc:
        # Standard behavior: keep BCC out of visible headers
        pass

    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(body, content_type, "utf-8"))

    attachment_count = 0
    if attachments:
        for file_path in attachments:
            if not os.path.exists(file_path):
                continue

            filename = os.path.basename(file_path)
            with open(file_path, "rb") as f:
                part = MIMEApplication(f.read())

            part.add_header(
                "Content-Disposition",
                "attachment",
                filename=Header(filename, "utf-8").encode(),
            )
            msg.attach(part)
            attachment_count += 1

    all_recipients = list(to)
    if cc:
        all_recipients.extend(cc)
    if bcc:
        all_recipients.extend(bcc)

    try:
        if smtp_port in (465, 994):
            with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30) as server:
                server.login(username, resolved_password)
                server.sendmail(username, all_recipients, msg.as_string())
        else:
            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
                server.ehlo()
                if server.has_extn("starttls"):
                    server.starttls()
                    server.ehlo()
                server.login(username, resolved_password)
                server.sendmail(username, all_recipients, msg.as_string())

        return {
            "success": True,
            "recipients": len(all_recipients),
            "attachments": attachment_count,
            "message_id": msg.get("Message-ID"),
        }
    except Exception as e:
        raise RuntimeError(f"Email sending failed: {e}") from e
