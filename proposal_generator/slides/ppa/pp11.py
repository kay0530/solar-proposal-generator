"""
pp11.py - 導入スケジュールスライド

Layout (A4 landscape):
  - Orange header bar with "導入スケジュール"
  - 4-phase visual timeline using colored rectangles
  - Phase descriptions with duration estimates
  - Total period summary
"""

from __future__ import annotations

from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_H, CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE,
    C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect,
    add_section_header, add_textbox,
)

TITLE = "導入スケジュール"

# Phase definitions: (label, duration, description, color)
PHASES = [
    (
        "STEP 1",
        "ご契約・設計",
        "1〜2ヶ月",
        "現地調査・電力需要分析・システム設計\nPPA契約締結・補助金申請手続き",
        C_ORANGE,
    ),
    (
        "STEP 2",
        "機器調達",
        "2〜3ヶ月",
        "太陽光パネル・PCS等の機器発注\n架台・配線部材の手配",
        RGBColor(0xF0, 0x7D, 0x48),  # lighter orange
    ),
    (
        "STEP 3",
        "施工・設置",
        "1〜2ヶ月",
        "屋根防水処理・架台設置\nパネル設置・電気工事・系統連系",
        RGBColor(0xF5, 0xA6, 0x23),  # amber
    ),
    (
        "STEP 4",
        "運転開始",
        "—",
        "試運転・検査完了後に発電開始\n遠隔監視システム稼働",
        RGBColor(0x4C, 0xAF, 0x50),  # green
    ),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP11 (installation schedule) onto an already-added blank slide.

    data keys used: (none required - static content)
    """
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.1)

    # ---- Total period banner ----
    banner_w = SLIDE_W - MARGIN * 2
    banner_h = Inches(0.55)
    add_rounded_rect(slide, MARGIN, y, banner_w, banner_h, C_LIGHT_ORANGE)
    add_textbox(slide,
                MARGIN + Inches(0.15), y + Inches(0.1),
                banner_w - Inches(0.3), Inches(0.35),
                "ご契約から運転開始まで：約4〜7ヶ月（目安）",
                font_name=FONT_BLACK, font_size_pt=15, font_color=C_ORANGE, bold=True,
                align=PP_ALIGN.CENTER)
    y += banner_h + Inches(0.3)

    # ---- Visual timeline bar ----
    timeline_x = MARGIN + Inches(0.3)
    timeline_w = SLIDE_W - MARGIN * 2 - Inches(0.6)
    bar_h = Inches(0.5)

    # Proportional widths for each phase (1-2, 2-3, 1-2, marker)
    proportions = [0.25, 0.35, 0.25, 0.15]
    phase_xs = []
    px = timeline_x
    for prop in proportions:
        phase_xs.append(px)
        px += timeline_w * prop

    # Draw timeline bars
    for i, (step, label, duration, desc, color) in enumerate(PHASES):
        pw = timeline_w * proportions[i]
        add_rect(slide, phase_xs[i], y, pw - Inches(0.02), bar_h, color)
        # Step label on bar
        add_textbox(slide,
                    phase_xs[i] + Inches(0.05), y + Inches(0.05),
                    pw - Inches(0.1), Inches(0.2),
                    step,
                    font_name=FONT_BODY, font_size_pt=8, font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)
        # Phase name on bar
        add_textbox(slide,
                    phase_xs[i] + Inches(0.05), y + Inches(0.22),
                    pw - Inches(0.1), Inches(0.25),
                    label,
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

    y += bar_h + Inches(0.3)

    # ---- Phase detail cards ----
    card_cols = 4
    gap = Inches(0.15)
    card_w = (SLIDE_W - MARGIN * 2 - gap * (card_cols - 1)) / card_cols
    card_h = Inches(2.8)

    for i, (step, label, duration, desc, color) in enumerate(PHASES):
        cx = MARGIN + i * (card_w + gap)

        # Card background
        add_rounded_rect(slide, cx, y, card_w, card_h, C_LIGHT_GRAY)

        # Colored top accent
        add_rect(slide, cx, y, card_w, Inches(0.07), color)

        # Step number
        add_textbox(slide,
                    cx + Inches(0.08), y + Inches(0.18),
                    card_w - Inches(0.16), Inches(0.22),
                    step,
                    font_name=FONT_BODY, font_size_pt=9, font_color=C_SUB,
                    align=PP_ALIGN.CENTER)

        # Phase name
        add_textbox(slide,
                    cx + Inches(0.08), y + Inches(0.42),
                    card_w - Inches(0.16), Inches(0.35),
                    label,
                    font_name=FONT_BLACK, font_size_pt=14, font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)

        # Duration badge
        badge_w = Inches(1.2)
        badge_h = Inches(0.3)
        badge_x = cx + (card_w - badge_w) / 2
        badge_y = y + Inches(0.85)
        add_rounded_rect(slide, badge_x, badge_y, badge_w, badge_h, color)
        add_textbox(slide,
                    badge_x, badge_y + Inches(0.03),
                    badge_w, Inches(0.24),
                    duration,
                    font_name=FONT_BODY, font_size_pt=11, font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Description
        add_textbox(slide,
                    cx + Inches(0.1), y + Inches(1.3),
                    card_w - Inches(0.2), Inches(1.3),
                    desc,
                    font_name=FONT_BODY, font_size_pt=9, font_color=C_SUB)

    add_footer(slide)
