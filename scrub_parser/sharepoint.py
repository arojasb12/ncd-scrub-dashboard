"""
Microsoft Graph API client for SharePoint file discovery and download.

Uses MSAL for authentication and requests for API calls.
Expects these environment variables:
  - GRAPH_TENANT_ID
  - GRAPH_CLIENT_ID
  - GRAPH_CLIENT_SECRET
  - SHAREPOINT_SITE_ID   (or SHAREPOINT_SITE_URL for auto-resolution)
  - SHAREPOINT_DRIVE_ID   (optional — resolved from site if not set)
"""

from __future__ import annotations

import io
import logging
import os
import re
from dataclasses import dataclass
from typing import Optional

import requests

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


@dataclass
class SharePointFile:
    """Represents a file found on SharePoint."""
    name: str
    web_url: str
    download_url: str
    item_id: str
    size: int = 0


class GraphClient:
    """
    Thin wrapper around Microsoft Graph REST API for SharePoint file ops.

    Usage:
        client = GraphClient.from_env()
        files = client.list_folder_files("Scrubs/Billing Date Alignment Scrub/2026/")
        content = client.download_file(files[0])
    """

    def __init__(self, tenant_id: str, client_id: str, client_secret: str,
                 site_id: str, drive_id: Optional[str] = None):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_id = site_id
        self.drive_id = drive_id
        self._token: Optional[str] = None

    @classmethod
    def from_env(cls) -> GraphClient:
        """Construct from environment variables."""
        return cls(
            tenant_id=os.environ["GRAPH_TENANT_ID"],
            client_id=os.environ["GRAPH_CLIENT_ID"],
            client_secret=os.environ["GRAPH_CLIENT_SECRET"],
            site_id=os.environ["SHAREPOINT_SITE_ID"],
            drive_id=os.environ.get("SHAREPOINT_DRIVE_ID"),
        )

    # ── Auth ───────────────────────────────────────────────────────────────

    def _get_token(self) -> str:
        """Acquire an app-only access token via client credentials flow."""
        if self._token:
            return self._token

        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        resp = requests.post(token_url, data={
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }, timeout=30)
        resp.raise_for_status()
        self._token = resp.json()["access_token"]
        return self._token

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._get_token()}"}

    def invalidate_token(self):
        """Force re-auth on next request."""
        self._token = None

    # ── Drive resolution ───────────────────────────────────────────────────

    def _resolve_drive_id(self) -> str:
        """Get the default document library drive ID for the site."""
        if self.drive_id:
            return self.drive_id

        url = f"{GRAPH_BASE}/sites/{self.site_id}/drive"
        resp = requests.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        self.drive_id = resp.json()["id"]
        logger.info("Resolved drive ID: %s", self.drive_id)
        return self.drive_id

    # ── File discovery ─────────────────────────────────────────────────────

    def list_folder_files(self, folder_path: str,
                          file_ext: str = ".xlsx") -> list[SharePointFile]:
        """
        List all files in a SharePoint folder (by path relative to the
        document library root).

        Returns files filtered to the given extension.
        """
        drive_id = self._resolve_drive_id()
        # URL-encode the folder path (colons and spaces)
        encoded = folder_path.strip("/")
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}:/children"

        files: list[SharePointFile] = []
        while url:
            resp = requests.get(url, headers=self._headers(),
                                params={"$top": 200}, timeout=30)
            if resp.status_code == 404:
                logger.warning("Folder not found: %s", folder_path)
                return []
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("value", []):
                name = item.get("name", "")
                if "folder" in item:
                    # recurse into subfolders (e.g. month subfolders)
                    subfolder = f"{folder_path.rstrip('/')}/{name}"
                    files.extend(self.list_folder_files(subfolder, file_ext))
                elif name.lower().endswith(file_ext):
                    dl_url = item.get("@microsoft.graph.downloadUrl", "")
                    files.append(SharePointFile(
                        name=name,
                        web_url=item.get("webUrl", ""),
                        download_url=dl_url,
                        item_id=item.get("id", ""),
                        size=item.get("size", 0),
                    ))

            url = data.get("@odata.nextLink")

        logger.info("Found %d files in %s", len(files), folder_path)
        return files

    def search_files(self, query: str,
                     file_ext: str = ".xlsx") -> list[SharePointFile]:
        """
        Search for files across the site's document library using Graph
        search API. Useful when folder structure is uncertain.
        """
        drive_id = self._resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root/search(q='{query}')"

        files: list[SharePointFile] = []
        while url:
            resp = requests.get(url, headers=self._headers(),
                                params={"$top": 200}, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            for item in data.get("value", []):
                name = item.get("name", "")
                if name.lower().endswith(file_ext):
                    dl_url = item.get("@microsoft.graph.downloadUrl", "")
                    files.append(SharePointFile(
                        name=name,
                        web_url=item.get("webUrl", ""),
                        download_url=dl_url,
                        item_id=item.get("id", ""),
                        size=item.get("size", 0),
                    ))

            url = data.get("@odata.nextLink")

        logger.info("Search '%s' returned %d files", query, len(files))
        return files

    # ── Download ───────────────────────────────────────────────────────────

    def download_file(self, sp_file: SharePointFile) -> io.BytesIO:
        """Download file content into an in-memory BytesIO buffer."""
        url = sp_file.download_url
        if not url:
            # fallback: construct download URL from item ID
            drive_id = self._resolve_drive_id()
            url = f"{GRAPH_BASE}/drives/{drive_id}/items/{sp_file.item_id}/content"

        resp = requests.get(url, headers=self._headers(), timeout=120,
                            stream=True)
        resp.raise_for_status()

        buf = io.BytesIO()
        for chunk in resp.iter_content(chunk_size=8192):
            buf.write(chunk)
        buf.seek(0)

        logger.debug("Downloaded %s (%d bytes)", sp_file.name, buf.getbuffer().nbytes)
        return buf

    def download_file_by_path(self, file_path: str) -> io.BytesIO:
        """Download a file by its full path relative to the library root."""
        drive_id = self._resolve_drive_id()
        encoded = file_path.strip("/")
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{encoded}:/content"

        resp = requests.get(url, headers=self._headers(), timeout=120,
                            stream=True)
        resp.raise_for_status()

        buf = io.BytesIO()
        for chunk in resp.iter_content(chunk_size=8192):
            buf.write(chunk)
        buf.seek(0)
        return buf
