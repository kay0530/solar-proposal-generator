"""Interactive compass component for azimuth angle selection."""

import streamlit.components.v1 as components
from pathlib import Path

_FRONTEND_DIR = Path(__file__).parent / "frontend"
_component_func = components.declare_component("compass", path=str(_FRONTEND_DIR))


def compass_input(value: int = 0, key: str = None) -> int:
    """Render an interactive compass for azimuth angle selection.

    Args:
        value: Initial angle in degrees (0=North, 90=East, 180=South, 270=West).
        key: Streamlit widget key.

    Returns:
        Selected angle in degrees (0-359).
    """
    result = _component_func(value=value, key=key, default=value)
    return int(result) if result is not None else value
