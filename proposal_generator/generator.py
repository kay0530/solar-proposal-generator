"""
generator.py
Main entry point: takes customer data + slide list, assembles PPTX.
"""

from __future__ import annotations

import importlib
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

# Default paths (adjust as needed)
BASE_DIR   = Path(__file__).parent.parent
TEMPLATE   = BASE_DIR / "templates" / "セールスツールテンプレート（白抜き）.pptx"
LOGO_PATH  = None  # Set to actual logo path when available


def generate_proposal(
    slide_ids: list[str],
    data: dict,
    output_path: Path,
    template_path: Path = TEMPLATE,
    logo_path: Path = None,
) -> Path:
    """
    Generate a PPTX proposal from slide IDs and data dict.

    Args:
        slide_ids: Ordered list of slide IDs (e.g. ["PP0","PP1","NEW_ff"])
        data: Output dict from excel_runner (or manually constructed)
        output_path: Where to save the generated PPTX
        template_path: Base template PPTX
        logo_path: Company logo PNG path

    Returns:
        output_path
    """
    # Always start from a fresh presentation to avoid template residue issues.
    # The design (logo, header bar) is drawn programmatically by each slide generator.
    prs = Presentation()
    prs.slide_width = Inches(11.69)   # A4 landscape width
    prs.slide_height = Inches(8.27)   # A4 landscape height

    effective_logo = logo_path or LOGO_PATH

    for slide_id in slide_ids:
        generator_fn = _load_generator(slide_id)
        if generator_fn is None:
            # Placeholder slide for unimplemented generators
            slide = _add_blank(prs)
            from proposal_generator.utils import add_textbox, C_SUB, FONT_BODY
            add_textbox(slide,
                        Inches(0.5), Inches(5.0),
                        Inches(7.0), Inches(0.5),
                        f"[{slide_id}] — スライド実装予定",
                        font_name=FONT_BODY, font_size_pt=14,
                        font_color=C_SUB)
        else:
            slide = _add_blank(prs)
            generator_fn(slide, data, effective_logo)

    prs.save(str(output_path))
    return output_path


def _add_blank(prs: Presentation):
    """Add a blank slide (layout index 6 = blank)."""
    # Find or use last layout
    layout = prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]
    return prs.slides.add_slide(layout)


def _load_generator(slide_id: str):
    """
    Dynamically import and return the generate() function for a slide ID.
    Module path convention:
        PP0  -> proposal_generator.slides.ppa.pp0
        EP0  -> proposal_generator.slides.epc.ep0
        NEW_ff -> proposal_generator.slides.new.new_ff
    """
    try:
        if slide_id.startswith("PP"):
            n = slide_id[2:].lower()
            module_path = f"proposal_generator.slides.ppa.pp{n}"
        elif slide_id.startswith("EP"):
            n = slide_id[2:].lower()
            module_path = f"proposal_generator.slides.epc.ep{n}"
        elif slide_id.startswith("NEW_"):
            name = slide_id[4:].lower()
            module_path = f"proposal_generator.slides.new.new_{name}"
        else:
            return None

        module = importlib.import_module(module_path)
        return module.generate
    except (ImportError, AttributeError):
        return None
