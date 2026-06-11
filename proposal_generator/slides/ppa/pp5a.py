"""
pp5a.py - 効果シミュレーション（セクション区切り）(Design v2)

Full-bleed navy section divider before the simulation slides.
Center-left (cols 1-9): 'SIMULATION' eyebrow -> section title ->
customer case line -> short orange rule. 'altenergy' wordmark
bottom-right (no white-logo image dependency).
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    C_NAVY, C_ORANGE, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    SIZE_CAPTION,
    add_line, add_rect, add_textbox, grid_x, grid_w,
)

TITLE = "効果シミュレーション"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    # Full-bleed navy canvas
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, C_NAVY)

    x = grid_x(1)
    w = grid_w(9)

    # Customer case line: company + office + のケース
    company = data.get("company_name", "") or ""
    office = data.get("office_name", "") or ""
    if company and office:
        case_text = f"{company}　{office}のケース"
    elif company:
        case_text = f"{company}のケース"
    elif office:
        case_text = f"{office}のケース"
    else:
        case_text = ""

    y0 = Inches(3.00)

    # Eyebrow
    add_textbox(slide, x, y0, w, Inches(0.22),
                "SIMULATION",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_ORANGE, bold=True, tracking_pt=1.2)

    # Section title
    add_textbox(slide, x, y0 + Inches(0.34), w, Inches(0.60),
                TITLE,
                font_name=FONT_BLACK, font_size_pt=32,
                font_color=C_WHITE, bold=True)

    # Customer case line
    if case_text:
        add_textbox(slide, x, y0 + Inches(1.08), w, Inches(0.30),
                    case_text,
                    font_name=FONT_BODY, font_size_pt=14,
                    font_color=C_WHITE)

    # Short orange rule below
    rule_y = y0 + Inches(1.58)
    add_line(slide, x, rule_y, x + Inches(2.0), rule_y,
             C_ORANGE, width_pt=0.75)

    # Wordmark bottom-right (no white-logo image dependency)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(1.6),
                SLIDE_H - Inches(0.48),
                Inches(1.6), Inches(0.24),
                "altenergy",
                font_name=FONT_BLACK, font_size_pt=10,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.RIGHT)
