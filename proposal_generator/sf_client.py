"""
sf_client.py - simple-salesforce backend for Streamlit Community Cloud.

Supports two authentication methods (checked in order):

1. **Refresh Token** (recommended for SSO/CloudGate environments):
       [salesforce]
       instance_url = "https://altenergyinc.my.salesforce.com"
       refresh_token = "xxxx"
       client_id = "PlatformCLI"

2. **Username/Password** (for direct-login environments):
       [salesforce]
       username = "user@example.com"
       password = "password"
       security_token = "token"
       domain = "login"
"""

from __future__ import annotations

import logging

import streamlit as st

logger = logging.getLogger(__name__)

_HAS_SIMPLE_SF = False
try:
    from simple_salesforce import Salesforce  # type: ignore
    _HAS_SIMPLE_SF = True
except ImportError:
    pass


def _is_configured() -> bool:
    """Return True if any Salesforce auth secrets are present."""
    try:
        sec = st.secrets.get("salesforce", {})
        if sec.get("refresh_token") and sec.get("instance_url"):
            return True
        if sec.get("username") and sec.get("password") and sec.get("security_token"):
            return True
        return False
    except Exception:
        return False


def _auth_via_refresh_token(sec) -> "Salesforce | None":
    """Authenticate using OAuth2 refresh token flow."""
    import requests

    client_id = sec.get("client_id", "PlatformCLI")
    token_url = sec.get("token_url", "https://login.salesforce.com/services/oauth2/token")

    resp = requests.post(token_url, data={
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": sec["refresh_token"],
    })

    if resp.status_code != 200:
        err = resp.json() if resp.text else {}
        raise RuntimeError(
            f"Token refresh failed ({resp.status_code}): "
            f"{err.get('error', '?')} - {err.get('error_description', resp.text[:200])}"
        )

    token_data = resp.json()
    access_token = token_data["access_token"]
    instance_url = token_data.get("instance_url", sec["instance_url"])

    sf = Salesforce(instance_url=instance_url, session_id=access_token)
    return sf


def _auth_via_password(sec) -> "Salesforce | None":
    """Authenticate using username/password."""
    sf = Salesforce(
        username=sec["username"],
        password=sec["password"],
        security_token=sec.get("security_token", ""),
        domain=sec.get("domain", "login"),
    )
    return sf


@st.cache_resource(show_spinner="Salesforce接続中...")
def _get_connection() -> "Salesforce | None":
    """Create and cache a Salesforce connection."""
    if not _HAS_SIMPLE_SF:
        logger.warning("simple-salesforce is not installed.")
        return None
    if not _is_configured():
        logger.info("Salesforce secrets not configured.")
        return None
    try:
        sec = st.secrets["salesforce"]

        # Method 1: Refresh Token (for SSO / CloudGate)
        if sec.get("refresh_token") and sec.get("instance_url"):
            logger.info("Attempting refresh token auth...")
            return _auth_via_refresh_token(sec)

        # Method 2: Username/Password
        if sec.get("username") and sec.get("password"):
            logger.info("Attempting username/password auth...")
            return _auth_via_password(sec)

    except Exception as e:
        logger.error("Failed to connect to Salesforce: %s", e)
        st.error(f"Salesforce接続エラー: {e}")
        return None
    return None


def sf_query(soql: str) -> list[dict]:
    """Execute a SOQL query and return records list."""
    conn = _get_connection()
    if conn is None:
        return []
    try:
        result = conn.query_all(soql)
        records = result.get("records", [])
        cleaned = []
        for rec in records:
            clean = {k: v for k, v in rec.items() if k != "attributes"}
            for k, v in clean.items():
                if isinstance(v, dict) and "attributes" in v:
                    clean[k] = {kk: vv for kk, vv in v.items() if kk != "attributes"}
            cleaned.append(clean)
        return cleaned
    except Exception as e:
        logger.error("sf_query failed: %s", e)
        return []


def is_available() -> bool:
    """Return True if simple-salesforce connection is usable."""
    return _get_connection() is not None
