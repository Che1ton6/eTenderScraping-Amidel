"""
SharePoint client for the eTender scraper (Azure App Service deployment).

# =============================================================================
# TODO (Siyabonga / Azure setup) - before this module can succeed at runtime:
# =============================================================================
# Pick ONE of two auth paths and configure it on the etenderwebapp App Service.
#
# PATH A - Managed Identity (recommended)
#   1. On the etenderwebapp App Service, Identity -> System-assigned -> On.
#   2. Grant that identity Microsoft Graph API permission
#      `Sites.ReadWrite.All` (application), OR use SharePoint site-level
#      permissions via `Sites.Selected` + PowerShell grant.
#   3. Set env var USE_MANAGED_IDENTITY=1 on the App Service.
#   4. Leave SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET unset.
#
# PATH B - App registration + client secret (fallback)
#   1. Azure AD -> App registrations -> New registration.
#   2. API permissions -> Microsoft Graph -> Application -> Sites.ReadWrite.All.
#      Grant admin consent.
#   3. Certificates & secrets -> New client secret. Copy the value.
#   4. On the App Service, set env vars:
#         SHAREPOINT_TENANT_ID     = <tenant guid>
#         SHAREPOINT_CLIENT_ID     = <app registration client id>
#         SHAREPOINT_CLIENT_SECRET = <secret value>
#      Leave USE_MANAGED_IDENTITY unset.
#
# ALWAYS required (either path):
#   SHAREPOINT_SITE_URL      = https://amidel.sharepoint.com/sites/<SITE>
#   SHAREPOINT_FOLDER_PATH   = <server-relative path inside the site's default
#                              document library, e.g. "TenderAutomation/master">
#   SHAREPOINT_MASTER_FILENAME = master_tenders.xlsx   (optional; this is default)
# =============================================================================
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from urllib.parse import quote

import httpx

log = logging.getLogger("etender.sharepoint")

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"

MASTER_FILENAME_DEFAULT = "master_tenders.xlsx"


def _acquire_token() -> str:
    """Get a Graph access token via Managed Identity or client-credentials."""
    if os.environ.get("USE_MANAGED_IDENTITY", "").strip().lower() in ("1", "true", "yes"):
        try:
            from azure.identity import ManagedIdentityCredential
        except ImportError as e:
            raise RuntimeError("azure-identity is required for Managed Identity auth") from e
        return ManagedIdentityCredential().get_token(GRAPH_SCOPE).token

    tenant = os.environ.get("SHAREPOINT_TENANT_ID")
    client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
    if not (tenant and client_id and client_secret):
        raise RuntimeError(
            "SharePoint auth not configured. Set USE_MANAGED_IDENTITY=1, or "
            "set SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET."
        )
    try:
        import msal
    except ImportError as e:
        raise RuntimeError("msal is required for client-credentials auth") from e

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=[GRAPH_SCOPE])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error_description')}")
    return result["access_token"]


def _parse_site_url(site_url: str) -> tuple[str, str]:
    m = re.match(r"^https?://([^/]+)(/.*)$", site_url.rstrip("/"))
    if not m:
        raise ValueError(f"Unexpected SHAREPOINT_SITE_URL: {site_url!r}")
    return m.group(1), m.group(2)


def _get_site_and_drive(client: httpx.Client, token: str) -> tuple[str, str]:
    site_url = os.environ["SHAREPOINT_SITE_URL"]
    hostname, site_path = _parse_site_url(site_url)
    headers = {"Authorization": f"Bearer {token}"}
    r = client.get(f"{GRAPH}/sites/{hostname}:{site_path}", headers=headers)
    r.raise_for_status()
    site_id = r.json()["id"]
    r = client.get(f"{GRAPH}/sites/{site_id}/drive", headers=headers)
    r.raise_for_status()
    return site_id, r.json()["id"]


def _folder_path() -> str:
    return os.environ.get("SHAREPOINT_FOLDER_PATH", "").strip("/")


def _item_path(*parts: str) -> str:
    folder = _folder_path()
    joined = "/".join(p.strip("/") for p in parts if p)
    return "/".join(p for p in (folder, joined) if p)


def download_master(local_dest: str) -> bool:
    """Pull master from SharePoint. Returns False if not yet present."""
    filename = os.environ.get("SHAREPOINT_MASTER_FILENAME", MASTER_FILENAME_DEFAULT)
    token = _acquire_token()
    with httpx.Client(timeout=120.0) as client:
        _, drive_id = _get_site_and_drive(client, token)
        headers = {"Authorization": f"Bearer {token}"}
        url = f"{GRAPH}/drives/{drive_id}/root:/{quote(_item_path(filename))}:/content"
        r = client.get(url, headers=headers, follow_redirects=True)
        if r.status_code == 404:
            log.info("SharePoint master not found (first run?): %s", filename)
            return False
        r.raise_for_status()
        Path(local_dest).parent.mkdir(parents=True, exist_ok=True)
        Path(local_dest).write_bytes(r.content)
        log.info("Downloaded SharePoint master -> %s (%d bytes)", local_dest, len(r.content))
        return True


def upload_master(local_src: str) -> None:
    filename = os.environ.get("SHAREPOINT_MASTER_FILENAME", MASTER_FILENAME_DEFAULT)
    _upload_file(local_src, _item_path(filename))
    log.info("Uploaded master -> SharePoint: %s", filename)


def upload_batch_folder(local_folder: str) -> int:
    base = Path(local_folder)
    if not base.is_dir():
        raise FileNotFoundError(f"Batch folder not found: {local_folder}")
    remote_root = _item_path("batches", base.name)
    count = 0
    for path in base.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(base).as_posix()
        _upload_file(str(path), f"{remote_root}/{rel}")
        count += 1
    log.info("Uploaded batch folder %s -> SharePoint (%d files)", base.name, count)
    return count


_UPLOAD_SMALL_LIMIT = 4 * 1024 * 1024


def _upload_file(local_src: str, remote_path: str) -> None:
    token = _acquire_token()
    with httpx.Client(timeout=300.0) as client:
        _, drive_id = _get_site_and_drive(client, token)
        headers = {"Authorization": f"Bearer {token}"}
        size = Path(local_src).stat().st_size
        item_path = quote(remote_path)

        if size <= _UPLOAD_SMALL_LIMIT:
            with open(local_src, "rb") as f:
                data = f.read()
            url = f"{GRAPH}/drives/{drive_id}/root:/{item_path}:/content"
            r = client.put(
                url,
                headers={**headers, "Content-Type": "application/octet-stream"},
                content=data,
            )
            r.raise_for_status()
            return

        url = f"{GRAPH}/drives/{drive_id}/root:/{item_path}:/createUploadSession"
        r = client.post(url, headers=headers, json={
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        r.raise_for_status()
        upload_url = r.json()["uploadUrl"]

        chunk_size = 5 * 1024 * 1024
        with open(local_src, "rb") as f:
            offset = 0
            while offset < size:
                chunk = f.read(chunk_size)
                end = offset + len(chunk) - 1
                cr = client.put(
                    upload_url,
                    headers={
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {offset}-{end}/{size}",
                    },
                    content=chunk,
                )
                if cr.status_code not in (200, 201, 202):
                    cr.raise_for_status()
                offset += len(chunk)
