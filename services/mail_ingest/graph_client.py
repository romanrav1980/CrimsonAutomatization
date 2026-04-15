from __future__ import annotations

import time
from typing import Any

import msal
import requests

from services.mail_ingest.config import OutlookSettings


class GraphApiError(RuntimeError):
    """Raised when Microsoft Graph returns an error."""


class OutlookGraphClient:
    def __init__(self, settings: OutlookSettings) -> None:
        self.settings = settings
        self._token: str | None = None
        self._token_expires_at: float = 0

    def list_messages(self, *, top: int | None = None, since_iso: str | None = None) -> list[dict[str, Any]]:
        endpoint = self._messages_endpoint()
        params: dict[str, str] = {
            "$top": str(top or self.settings.max_messages),
            "$orderby": "receivedDateTime desc",
            "$select": ",".join(
                [
                    "id",
                    "subject",
                    "sender",
                    "from",
                    "toRecipients",
                    "ccRecipients",
                    "bccRecipients",
                    "conversationId",
                    "internetMessageId",
                    "receivedDateTime",
                    "sentDateTime",
                    "bodyPreview",
                    "body",
                    "hasAttachments",
                    "importance",
                    "isRead",
                    "categories",
                    "webLink",
                ]
            ),
        }
        if since_iso:
            params["$filter"] = f"receivedDateTime ge {since_iso}"
        response = self._request("GET", endpoint, params=params)
        items = response.get("value", [])
        source_store_name = self.settings.mailbox_name or self.settings.user_id or ""
        source_folder_name = self.settings.folder
        source_folder_path = self.settings.folder
        for item in items:
            item.setdefault("sourceFolderName", source_folder_name)
            item.setdefault("sourceFolderPath", source_folder_path)
            item.setdefault("sourceStoreName", source_store_name)
        return items

    def list_attachments(self, message_id: str) -> list[dict[str, Any]]:
        endpoint = f"{self._message_item_endpoint(message_id)}/attachments"
        params = {
            "$select": ",".join(
                [
                    "id",
                    "name",
                    "contentType",
                    "size",
                    "@odata.type",
                    "contentBytes",
                    "isInline",
                    "lastModifiedDateTime",
                ]
            )
        }
        response = self._request("GET", endpoint, params=params)
        return response.get("value", [])

    def _messages_endpoint(self) -> str:
        base = self.settings.graph_base_url
        if self.settings.auth_mode == "device_code":
            return f"{base}/me/mailFolders/{self.settings.folder}/messages"
        return f"{base}/users/{self.settings.user_id}/mailFolders/{self.settings.folder}/messages"

    def _message_item_endpoint(self, message_id: str) -> str:
        base = self.settings.graph_base_url
        if self.settings.auth_mode == "device_code":
            return f"{base}/me/messages/{message_id}"
        return f"{base}/users/{self.settings.user_id}/messages/{message_id}"

    def _request(self, method: str, url: str, **kwargs: Any) -> dict[str, Any]:
        token = self._get_access_token()
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"
        headers["Accept"] = "application/json"
        response = requests.request(method, url, headers=headers, timeout=60, **kwargs)
        if response.status_code >= 400:
            raise GraphApiError(f"Graph API error {response.status_code}: {response.text}")
        if not response.text:
            return {}
        return response.json()

    def _get_access_token(self) -> str:
        now = time.time()
        if self._token and now < self._token_expires_at - 60:
            return self._token

        authority = f"https://login.microsoftonline.com/{self.settings.tenant_id}"
        if self.settings.auth_mode == "device_code":
            app = msal.PublicClientApplication(
                client_id=self.settings.client_id,
                authority=authority,
            )
            flow = app.initiate_device_flow(scopes=["Mail.Read"])
            if "user_code" not in flow:
                raise RuntimeError(f"Failed to create device-code flow: {flow}")
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
        else:
            app = msal.ConfidentialClientApplication(
                client_id=self.settings.client_id,
                authority=authority,
                client_credential=self.settings.client_secret,
            )
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        access_token = result.get("access_token")
        if not access_token:
            raise RuntimeError(f"Could not acquire Graph token: {result}")

        expires_in = int(result.get("expires_in", 3600))
        self._token = access_token
        self._token_expires_at = now + expires_in
        return access_token
