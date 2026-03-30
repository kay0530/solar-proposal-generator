"""
pp5a.py - 効果シミュレーション（セクション区切り）

PDF P6: Section divider page before the simulation slides.
Shows customer name and "効果シミュレーション" in large text.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    C_DARK, C_LIGHT_GRAY, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "【効果シミュレーション】"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    # Full-page dark background effect
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_DARK)

    # Orange accent bar at top
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.08), C_ORANGE)

    # Logo
    if logo_path and Path(logo_path).exists():
        slide.shapes.add_picture(
            str(logo_path),
            MARGIN, Inches(0.3), Inches(2.5),
        )

    # Section title
    add_textbox(slide, MARGIN, Inches(2.5), SLIDE_W - MARGIN * 2, Inches(1.0),
                "【効果シミュレーション】",
                font_name=FONT_BLACK, font_size_pt=36,
                font_color=C_ORANGE, bold=True,
                align=PP_ALIGN.CENTER)

    # Customer name
    company = data.get("company_name", "")
    office = data.get("office_name", "")
    if office:
        customer_text = f"{company}様\n{office}のケース"
    elif company:
        customer_text = f"{company}様"
    else:
        customer_text = ""

    add_textbox(slide, MARGIN, Inches(3.8), SLIDE_W - MARGIN * 2, Inches(1.0),
                customer_text,
                font_name=FONT_BODY, font_size_pt=18,
                font_color=C_WHITE, align=PP_ALIGN.CENTER)

    # Bottom accent
    add_rect(slide, 0, SLIDE_H - Inches(0.08), SLIDE_W, Inches(0.08), C_ORANGE)

    add_footer(slide)
