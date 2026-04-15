from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def load_env_file(path: str | Path = ".env") -> None:
    env_path = Path(path)
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


def _as_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass(slots=True)
class OutlookSettings:
    provider: str
    auth_mode: str
    tenant_id: str
    client_id: str
    client_secret: str | None
    user_id: str | None
    mailbox_name: str | None
    folder: str
    all_folders: bool
    max_messages: int
    save_attachments: bool
    output_dir: Path
    historical_unread_catchup_enabled: bool
    historical_unread_catchup_batch_size: int
    graph_base_url: str = "https://graph.microsoft.com/v1.0"

    @classmethod
    def from_env(cls) -> "OutlookSettings":
        provider = os.getenv("OUTLOOK_PROVIDER", "desktop").strip().lower()
        auth_mode = os.getenv("OUTLOOK_AUTH_MODE", "device_code").strip().lower()
        tenant_id = os.getenv("OUTLOOK_TENANT_ID", "").strip()
        client_id = os.getenv("OUTLOOK_CLIENT_ID", "").strip()
        client_secret = os.getenv("OUTLOOK_CLIENT_SECRET", "").strip() or None
        user_id = os.getenv("OUTLOOK_USER_ID", "").strip() or None
        mailbox_name = os.getenv("OUTLOOK_MAILBOX_NAME", "").strip() or user_id
        folder = os.getenv("OUTLOOK_MAIL_FOLDER", "inbox").strip().lower()
        all_folders = _as_bool(os.getenv("OUTLOOK_ALL_FOLDERS"), default=False)
        max_messages = int(os.getenv("OUTLOOK_MAX_MESSAGES", "25"))
        save_attachments = _as_bool(os.getenv("OUTLOOK_SAVE_ATTACHMENTS"), default=True)
        output_dir = Path(os.getenv("OUTLOOK_OUTPUT_DIR", "raw/mail"))
        historical_unread_catchup_enabled = _as_bool(os.getenv("OUTLOOK_HISTORICAL_UNREAD_CATCHUP_ENABLED"), default=True)
        historical_unread_catchup_batch_size = int(os.getenv("OUTLOOK_HISTORICAL_UNREAD_CATCHUP_BATCH_SIZE", "100"))

        settings = cls(
            provider=provider,
            auth_mode=auth_mode,
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            user_id=user_id,
            mailbox_name=mailbox_name,
            folder=folder,
            all_folders=all_folders,
            max_messages=max_messages,
            save_attachments=save_attachments,
            output_dir=output_dir,
            historical_unread_catchup_enabled=historical_unread_catchup_enabled,
            historical_unread_catchup_batch_size=historical_unread_catchup_batch_size,
        )
        settings.validate()
        return settings

    def validate(self) -> None:
        if self.provider not in {"desktop", "graph"}:
            raise ValueError("OUTLOOK_PROVIDER must be desktop or graph.")
        if self.provider == "graph":
            if not self.tenant_id:
                raise ValueError("OUTLOOK_TENANT_ID is required for graph provider.")
            if not self.client_id:
                raise ValueError("OUTLOOK_CLIENT_ID is required for graph provider.")
            if self.auth_mode not in {"device_code", "client_credentials"}:
                raise ValueError("OUTLOOK_AUTH_MODE must be device_code or client_credentials.")
            if self.auth_mode == "client_credentials":
                if not self.client_secret:
                    raise ValueError("OUTLOOK_CLIENT_SECRET is required for client_credentials auth.")
                if not self.user_id:
                    raise ValueError("OUTLOOK_USER_ID is required for client_credentials auth.")
        if self.historical_unread_catchup_batch_size < 0:
            raise ValueError("OUTLOOK_HISTORICAL_UNREAD_CATCHUP_BATCH_SIZE must be zero or greater.")
