"""
SharePoint client for the eTender scraper (Azure App Service deployment).

Handles three files under SHAREPOINT_FOLDER_PATH:
    auto_tenders.xlsx           <- scraper writes; download to append, upload to persist
    master_tenders.xlsx         <- derived (merge of auto + manual); Power BI reads
    Manual_Scrapes/manual_tenders.xlsx  <- humans write; scraper only reads

Plus the batch folder uploads to Auto_Scrapes/<batch-name>/.

Env vars (auth):
    USE_MANAGED_IDENTITY=1                (path A) OR
    SHAREPOINT_TENANT_ID / _CLIENT_ID / _CLIENT_SECRET  (path B)

Env vars (paths):
    SHAREPOINT_SITE_URL                   required
    SHAREPOINT_FOLDER_PATH                required, e.g. "etenders.gov.za (Tender Scraper)"
    SHAREPOINT_AUTO_FILENAME              default: auto_tenders.xlsx
    SHAREPOINT_MASTER_FILENAME            default: master_tenders.xlsx
    SHAREPOINT_MANUAL_SUBPATH             default: Manual_Scrapes/manual_tenders.xlsx
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

AUTO_FILENAME_DEFAULT = "auto_tenders.xlsx"
MASTER_FILENAME_DEFAULT = "master_tenders.xlsx"
MANUAL_SUBPATH_DEFAULT = "Manual_Scrapes/manual_tenders.xlsx"


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


def _download(remote_subpath: str, local_dest: str) -> bool:
    """Generic download helper. remote_subpath is relative to SHAREPOINT_FOLDER_PATH."""
    token = _acquire_token()
    with httpx.Client(timeout=120.0) as client:
        _, drive_id = _get_site_and_drive(client, token)
        headers = {"Authorization": f"Bearer {token}"}
        url = f"{GRAPH}/drives/{drive_id}/root:/{quote(_item_path(remote_subpath))}:/content"
        r = client.get(url, headers=headers, follow_redirects=True)
        if r.status_code == 404:
            log.info("SharePoint file not found: %s", remote_subpath)
            return False
        r.raise_for_status()
        Path(local_dest).parent.mkdir(parents=True, exist_ok=True)
        Path(local_dest).write_bytes(r.content)
        log.info("Downloaded %s -> %s (%d bytes)", remote_subpath, local_dest, len(r.content))
        return True


def download_auto(local_dest: str) -> bool:
    filename = os.environ.get("SHAREPOINT_AUTO_FILENAME", AUTO_FILENAME_DEFAULT)
    return _download(filename, local_dest)


def download_master(local_dest: str) -> bool:
    filename = os.environ.get("SHAREPOINT_MASTER_FILENAME", MASTER_FILENAME_DEFAULT)
    return _download(filename, local_dest)


def download_manual(local_dest: str) -> bool:
    subpath = os.environ.get("SHAREPOINT_MANUAL_SUBPATH", MANUAL_SUBPATH_DEFAULT)
    return _download(subpath, local_dest)


def upload_auto(local_src: str) -> None:
    filename = os.environ.get("SHAREPOINT_AUTO_FILENAME", AUTO_FILENAME_DEFAULT)
    _upload_file(local_src, _item_path(filename))
    log.info("Uploaded auto -> SharePoint: %s", filename)


def upload_master(local_src: str) -> None:
    filename = os.environ.get("SHAREPOINT_MASTER_FILENAME", MASTER_FILENAME_DEFAULT)
    _upload_file(local_src, _item_path(filename))
    log.info("Uploaded master -> SharePoint: %s", filename)


def upload_batch_folder(local_folder: str) -> int:
    """
    Upload the per-scrape output to SharePoint under Auto_Scrapes/<batch-name>/.

    Only three things are uploaded — the rest of what BatchProcessor writes
    locally (batches/, Tender Analysis, PowerBI export) stays local:
        - end product/**
        - Display Equation/**
        - Tender Summary.xlsx  (top-level file)
    """
    base = Path(local_folder)
    if not base.is_dir():
        raise FileNotFoundError(f"Batch folder not found: {local_folder}")

    remote_root = _item_path("Auto_Scrapes", base.name)
    keep_dirs = ("end product", "Display Equation")
    keep_files_root = ("Tender Summary.xlsx",)

    count = 0
    for path in base.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(base)
        parts = rel.parts
        keep = False
        if len(parts) == 1 and parts[0] in keep_files_root:
            keep = True
        elif len(parts) >= 2 and parts[0] in keep_dirs:
            keep = True
        if not keep:
            continue
        _upload_file(str(path), f"{remote_root}/{rel.as_posix()}")
        count += 1

    log.info("Uploaded batch folder %s -> SharePoint/Auto_Scrapes (%d files)", base.name, count)
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
