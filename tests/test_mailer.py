"""Tests for the SMTP mailer using a fake server."""

import smtplib
from email.message import EmailMessage
from types import TracebackType

import pytest

from daily_plan.mailer import MailConfigError, send_images

SETTINGS = {
    "SMTP_HOST": "smtp.test",
    "SMTP_PORT": "587",
    "SMTP_USER": "user",
    "SMTP_PASSWORD": "secret",
    "MAIL_FROM": "plan@test",
}


class FakeSMTP:
    """Records what the mailer does instead of opening a connection."""

    sent: list[EmailMessage] = []
    calls: list[str] = []

    def __init__(self, host: str, port: int) -> None:
        self.calls.append(f"connect {host}:{port}")

    def __enter__(self) -> "FakeSMTP":
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        return None

    def starttls(self) -> None:
        self.calls.append("starttls")

    def login(self, user: str, password: str) -> None:
        self.calls.append(f"login {user}")

    def send_message(self, message: EmailMessage) -> None:
        self.sent.append(message)


def test_send_images_attaches_pngs(monkeypatch: pytest.MonkeyPatch) -> None:
    """Every image becomes a png attachment sent over STARTTLS."""
    for key, value in SETTINGS.items():
        monkeypatch.setenv(key, value)
    FakeSMTP.sent, FakeSMTP.calls = [], []
    monkeypatch.setattr(smtplib, "SMTP", FakeSMTP)

    send_images({"Watson_2026-11-06": b"png1", "team": b"png2"}, "a@b.c", "S", "B")

    message = FakeSMTP.sent[0]
    assert message["To"] == "a@b.c"
    names = [part.get_filename() for part in message.iter_attachments()]
    assert names == ["Watson_2026-11-06.png", "team.png"]
    assert FakeSMTP.calls == ["connect smtp.test:587", "starttls", "login user"]


def test_missing_settings_raise(monkeypatch: pytest.MonkeyPatch) -> None:
    """Missing SMTP settings give a clear error."""
    for key in SETTINGS:
        monkeypatch.delenv(key, raising=False)
    with pytest.raises(MailConfigError, match="SMTP_HOST"):
        send_images({"a": b"x"}, "a@b.c", "S", "B")
