"""
box_client.py - Box REST API integration for proposal file management.

Uses requests library (no boxsdk dependency).
Auth token is stored in box_config.json next to this file.

Box folder structure (Salesforce Sync):
  Salesforce Altenergy Sync / 案件進捗 / {商談名} / 03_提案資料
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional

import requests

logger = logging.getLogger(__name__)

CONFIG_PATH = Path(__file__).parent / "box_config.json"
BOX_API_BASE = "https://api.box.com/2.0"
BOX_UPLOAD_BASE = "https://upload.box.com/api/2.0"
DEALS_FOLDER_ID = "164168692319"  # 案件進捗 folder


def _load_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def _headers() -> dict:
    cfg = _load_config()
    token = cfg.get("access_token", "")
    if not token:
        raise BoxAuthError("access_tokenが未設定です。box_config.jsonを確認してください。")
    return {"Authorization": f"Bearer {token}"}


class BoxAuthError(Exception):
    pass


class BoxAPIError(Exception):
    pass


def is_available() -> bool:
    """Check if Box integration is configured."""
    cfg = _load_config()
    return bool(cfg.get("access_token"))


def search_deal_folder(deal_name: str) -> Optional[dict]:
    """Search for a deal folder under 案件進捗 by name.

    Returns {"id": str, "name": str} or None.
    """
    params = {
        "query": deal_name,
        "type": "folder",
        "ancestor_folder_ids": DEALS_FOLDER_ID,
        "limit": 10,
    }
    resp = requests.get(f"{BOX_API_BASE}/search", headers=_headers(), params=params)
    resp.raise_for_status()
    entries = resp.json().get("entries", [])
    for entry in entries:
        if entry.get("type") == "folder" and deal_name in entry.get("name", ""):
            return {"id": entry["id"], "name": entry["name"]}
    return None


def find_proposal_folder(deal_folder_id: str) -> Optional[str]:
    """Find 03_提案資料 subfolder inside a deal folder. Returns folder ID."""
    params = {"fields": "id,name,type", "limit": 100}
    resp = requests.get(
        f"{BOX_API_BASE}/folders/{deal_folder_id}/items",
        headers=_headers(),
        params=params,
    )
    resp.raise_for_status()
    for entry in resp.json().get("entries", []):
        if entry.get("type") == "folder" and "03_提案資料" in entry.get("name", ""):
            return entry["id"]
    return None


def list_files(folder_id: str) -> list[dict]:
    """List files in a Box folder. Returns [{"id", "name", "modified_at"}]."""
    params = {"fields": "id,name,type,modified_at,size", "limit": 200}
    resp = requests.get(
        f"{BOX_API_BASE}/folders/{folder_id}/items",
        headers=_headers(),
        params=params,
    )
    resp.raise_for_status()
    return [
        {
            "id": e["id"],
            "name": e["name"],
            "modified_at": e.get("modified_at", ""),
            "size": e.get("size", 0),
        }
        for e in resp.json().get("entries", [])
        if e.get("type") == "file"
    ]


def upload_file(folder_id: str, file_path: Path, filename: str = None) -> dict:
    """Upload a file to Box. Returns {"id", "name"}."""
    fname = filename or file_path.name
    attributes = json.dumps({"name": fname, "parent": {"id": folder_id}})
    with open(file_path, "rb") as f:
        resp = requests.post(
            f"{BOX_UPLOAD_BASE}/files/content",
            headers=_headers(),
            data={"attributes": attributes},
            files={"file": (fname, f)},
        )
    if resp.status_code == 409:
        # File already exists - upload new version
        existing_id = resp.json()["context_info"]["conflicts"]["id"]
        return upload_new_version(existing_id, file_path, fname)
    resp.raise_for_status()
    entry = resp.json()["entries"][0]
    return {"id": entry["id"], "name": entry["name"]}


def upload_new_version(file_id: str, file_path: Path, filename: str = None) -> dict:
    """Upload a new version of an existing Box file."""
    fname = filename or file_path.name
    with open(file_path, "rb") as f:
        resp = requests.post(
            f"{BOX_UPLOAD_BASE}/files/{file_id}/content",
            headers=_headers(),
            files={"file": (fname, f)},
        )
    resp.raise_for_status()
    entry = resp.json()["entries"][0]
    return {"id": entry["id"], "name": entry["name"]}


def download_file(file_id: str, save_path: Path) -> Path:
    """Download a file from Box."""
    resp = requests.get(
        f"{BOX_API_BASE}/files/{file_id}/content",
        headers=_headers(),
        allow_redirects=True,
    )
    resp.raise_for_status()
    with open(save_path, "wb") as f:
        f.write(resp.content)
    return save_path


def get_deal_proposal_folder(deal_name: str) -> Optional[str]:
    """Convenience: find deal folder → find 03_提案資料 subfolder.

    Returns folder_id of 03_提案資料 or None.
    """
    deal = search_deal_folder(deal_name)
    if not deal:
        return None
    return find_proposal_folder(deal["id"])
