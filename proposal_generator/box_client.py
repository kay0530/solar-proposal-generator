"""
box_client.py - Box REST API integration for proposal file management.

Uses OAuth 2.0 with refresh token for persistent authentication.
Tokens are stored in box_config.json (local) or st.secrets (Cloud).

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
BOX_TOKEN_URL = "https://api.box.com/oauth2/token"
DEALS_FOLDER_ID = "164168692319"  # 案件進捗 folder


class BoxAuthError(Exception):
    pass


class BoxAPIError(Exception):
    pass


def _load_config() -> dict:
    """Load Box config from session_state cache, box_config.json, or st.secrets."""
    # 1. Try session_state cache (refreshed tokens on Cloud)
    try:
        import streamlit as st
        cached = st.session_state.get("_box_config_cache")
        if cached:
            return dict(cached)
    except Exception:
        pass
    # 2. Try local file (local dev)
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    # 3. Try Streamlit secrets (Cloud deployment, initial config)
    try:
        import streamlit as st
        if "box" in st.secrets:
            return dict(st.secrets["box"])
    except Exception:
        pass
    return {}


def _save_config(cfg: dict) -> None:
    """Save config to file if possible, otherwise cache in session_state."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except OSError:
        # Streamlit Cloud: can't write to disk, cache in session_state
        try:
            import streamlit as st
            st.session_state["_box_config_cache"] = cfg
        except Exception:
            logger.warning("Could not persist Box config")


def _refresh_access_token() -> str:
    """Use refresh_token to obtain a new access_token from Box OAuth2.

    Box refresh tokens are single-use: each refresh returns a new
    refresh_token that must be saved for the next cycle.
    """
    cfg = _load_config()
    refresh_token = cfg.get("refresh_token", "")
    client_id = cfg.get("client_id", "")
    client_secret = cfg.get("client_secret", "")

    if not all([refresh_token, client_id, client_secret]):
        raise BoxAuthError(
            "box_config.json または st.secrets[box] に "
            "client_id / client_secret / refresh_token を設定してください"
        )

    resp = requests.post(BOX_TOKEN_URL, data={
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
    })

    if resp.status_code != 200:
        err = resp.json().get("error_description", resp.text)
        raise BoxAuthError(f"Token refresh failed: {err}")

    data = resp.json()
    cfg["access_token"] = data["access_token"]
    cfg["refresh_token"] = data["refresh_token"]  # new single-use token
    _save_config(cfg)
    logger.info("Box access token refreshed successfully")
    return data["access_token"]


def _get_access_token() -> str:
    """Get a valid access token, refreshing if needed."""
    cfg = _load_config()
    token = cfg.get("access_token", "")
    if token:
        return token
    return _refresh_access_token()


def _headers() -> dict:
    token = _get_access_token()
    return {"Authorization": f"Bearer {token}"}


def _request(method: str, url: str, **kwargs) -> requests.Response:
    """Make a Box API request with automatic token refresh on 401."""
    headers = kwargs.pop("headers", {})
    headers.update(_headers())

    resp = requests.request(method, url, headers=headers, **kwargs)

    if resp.status_code == 401:
        # Token expired — refresh and retry once
        logger.info("Access token expired, refreshing...")
        new_token = _refresh_access_token()
        headers["Authorization"] = f"Bearer {new_token}"
        resp = requests.request(method, url, headers=headers, **kwargs)

    return resp


def is_available() -> bool:
    """Check if Box integration is configured."""
    cfg = _load_config()
    return bool(cfg.get("refresh_token") or cfg.get("access_token"))


def test_connection() -> dict:
    """Test Box connection by fetching current user info."""
    resp = _request("GET", f"{BOX_API_BASE}/users/me")
    resp.raise_for_status()
    user = resp.json()
    return {"name": user.get("name", ""), "login": user.get("login", "")}


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
    resp = _request("GET", f"{BOX_API_BASE}/search", params=params)
    resp.raise_for_status()
    entries = resp.json().get("entries", [])
    for entry in entries:
        if entry.get("type") == "folder" and deal_name in entry.get("name", ""):
            return {"id": entry["id"], "name": entry["name"]}
    return None


def find_proposal_folder(deal_folder_id: str) -> Optional[str]:
    """Find 03_提案資料 subfolder inside a deal folder. Returns folder ID."""
    params = {"fields": "id,name,type", "limit": 100}
    resp = _request(
        "GET",
        f"{BOX_API_BASE}/folders/{deal_folder_id}/items",
        params=params,
    )
    resp.raise_for_status()
    for entry in resp.json().get("entries", []):
        name = entry.get("name", "")
        if entry.get("type") == "folder" and name.startswith("03"):
            return entry["id"]
    return None


def list_files(folder_id: str) -> list[dict]:
    """List files in a Box folder. Returns [{"id", "name", "modified_at"}]."""
    params = {"fields": "id,name,type,modified_at,size", "limit": 200}
    resp = _request(
        "GET",
        f"{BOX_API_BASE}/folders/{folder_id}/items",
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
        resp = _request(
            "POST",
            f"{BOX_UPLOAD_BASE}/files/content",
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
        resp = _request(
            "POST",
            f"{BOX_UPLOAD_BASE}/files/{file_id}/content",
            files={"file": (fname, f)},
        )
    resp.raise_for_status()
    entry = resp.json()["entries"][0]
    return {"id": entry["id"], "name": entry["name"]}


def download_file(file_id: str, save_path: Path) -> Path:
    """Download a file from Box."""
    resp = _request(
        "GET",
        f"{BOX_API_BASE}/files/{file_id}/content",
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
