"""
sf_client.py - simple-salesforce backend for Streamlit Community Cloud.

Authenticates via st.secrets["salesforce"] and provides sf_query() that
returns records in the same format as the CLI-based _sf_query() in app.py.

Expected secrets (in .streamlit/secrets.toml or Streamlit Cloud settings):

    [salesforce]
    username = "user@example.com"
    password = "password"
    security_token = "token"
    domain = "login"          # optional, defaults to "login" (production)
"""

from __future__ import annotations

import logging

import streamlit as st

logger = logging.getLogger(__name__)

# Flag: whether simple-salesforce is installed
_HAS_SIMPLE_SF = False
try:
    from simple_salesforce import Salesforce  # type: ignore
    _HAS_SIMPLE_SF = True
except ImportError:
    pass


def _is_configured() -> bool:
    """Return True if Salesforce secrets are present in st.secrets."""
    try:
        sec = st.secrets.get("salesforce", {})
        return bool(sec.get("username") and sec.get("password") and sec.get("security_token"))
    except Exception:
        return False


@st.cache_resource(show_spinner="Salesforce接続中...")
def _get_connection() -> "Salesforce | None":
    """Create and cache a Salesforce connection via simple-salesforce."""
    if not _HAS_SIMPLE_SF:
        logger.warning("simple-salesforce is not installed; skipping API connection.")
        return None
    if not _is_configured():
        logger.info("Salesforce secrets not configured; skipping API connection.")
        return None
    try:
        sec = st.secrets["salesforce"]
        sf = Salesforce(
            username=sec["username"],
            password=sec["password"],
            security_token=sec["security_token"],
            domain=sec.get("domain", "login"),
        )
        return sf
    except Exception as e:
        logger.error("Failed to connect to Salesforce via simple-salesforce: %s", e)
        return None


def sf_query(soql: str) -> list[dict]:
    """Execute a SOQL query and return records list.

    Returns the same format as the CLI _sf_query: a list of dicts where each
    dict contains the queried fields. The 'attributes' key is stripped from
    each record for consistency.

    Returns an empty list on any failure (no secrets, import error, etc.).
    """
    conn = _get_connection()
    if conn is None:
        return []
    try:
        result = conn.query_all(soql)
        records = result.get("records", [])
        # Strip 'attributes' metadata that simple-salesforce includes
        cleaned = []
        for rec in records:
            clean = {k: v for k, v in rec.items() if k != "attributes"}
            # Also strip 'attributes' from nested relationship objects
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
