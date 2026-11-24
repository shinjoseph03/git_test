"""Microsoft Graph helpers for SharePoint file download/upload."""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from urllib.parse import quote

import msal
import requests

log = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"


class GraphError(Exception):
    """Graph API / auth failure."""


@dataclass
class GraphCredentials:
    tenant_id: str
    client_id: str
    client_secret: str

    @classmethod
    def from_env(cls) -> GraphCredentials:
        tenant = os.environ.get("AZURE_TENANT_ID", "").strip()
        client_id = os.environ.get("AZURE_CLIENT_ID", "").strip()
        secret = os.environ.get("AZURE_CLIENT_SECRET", "").strip()
        missing = [
            name
            for name, val in (
                ("AZURE_TENANT_ID", tenant),
                ("AZURE_CLIENT_ID", client_id),
                ("AZURE_CLIENT_SECRET", secret),
            )
            if not val
        ]
        if missing:
            raise GraphError(f"Missing Azure secrets: {', '.join(missing)}")
        return cls(tenant_id=tenant, client_id=client_id, client_secret=secret)


def acquire_token(creds: GraphCredentials) -> str:
    app = msal.ConfidentialClientApplication(
        creds.client_id,
        authority=f"https://login.microsoftonline.com/{creds.tenant_id}",
        client_credential=creds.client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise GraphError(
            f"Token acquisition failed: {result.get('error')} — {result.get('error_description')}"
        )
    return result["access_token"]


class GraphClient:
    def __init__(self, token: str, session: requests.Session | None = None):
        self._session = session or requests.Session()
        self._token = token

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._token}"}

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        resp = self._session.request(method, url, headers=self._headers(), timeout=60, **kwargs)
        if resp.status_code >= 400:
            raise GraphError(f"{method} {url} -> {resp.status_code}: {resp.text[:500]}")
        return resp

    def resolve_site_id(self, hostname: str, site_path: str) -> str:
        """Resolve site id from hostname + server-relative path (/teams/...)."""
        path = site_path if site_path.startswith("/") else f"/{site_path}"
        url = f"{GRAPH}/sites/{hostname}:{path}"
        data = self._request("GET", url).json()
        site_id = data.get("id")
        if not site_id:
            raise GraphError(f"No site id in response for {hostname}:{path}")
        return site_id

    def download_drive_item(self, site_id: str, item_id: str) -> bytes:
        # Prefer sites/{site-id}/drive/items/{id}/content
        url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/content"
        resp = self._request("GET", url)
        return resp.content

    def upload_drive_item(self, site_id: str, item_id: str, content: bytes) -> None:
        url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/content"
        resp = self._session.put(
            url,
            headers={
                **self._headers(),
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
            data=content,
            timeout=120,
        )
        if resp.status_code >= 400:
            raise GraphError(
                f"Upload failed {resp.status_code}: {resp.text[:500]} "
                "(file may be locked/blocked — treat as skip)"
            )

    def download_by_path(self, site_id: str, item_path: str) -> bytes:
        encoded = quote(item_path.lstrip("/"))
        url = f"{GRAPH}/sites/{site_id}/drive/root:/{encoded}:/content"
        return self._request("GET", url).content
