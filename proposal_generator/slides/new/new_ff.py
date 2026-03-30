"""
new_ff.py - FF振り返り（前回ヒアリング結果）スライド

FF = Fact Findings. Shows what was learned in the previous customer visit:
- Current situation & challenges
- Person-in-charge needs
- Management appeal points
- Constraints / concerns
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header,
)

TITLE = "ヒアリング結果（FF振り返り）"

SECTIONS = [
    ("ff_current_situation", "現状・課題",          C_ORANGE,     "電気代・設備状況・運用課題など"),
    ("ff_customer_needs",    "担当者ニーズ",         C_LIGHT_ORANGE, "担当者が上司・経営者に訴えたいこと"),
    ("ff_mgmt_needs",        "経営者へのアピール",   C_LIGHT_GRAY, "経営者が関心を持つポイント（ROI・リスク・環境）"),
    ("ff_constraints",       "制約・懸念事項",       C_LIGHT_GRAY, "屋根強度・予算・タイムライン・補助金期限など"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)
    company = data.get("company_name", "") or ""

    # Header: customer name + proposal date
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                f"{company}　様　|　ヒアリング実施日: {data.get('proposal_date', '—')}",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_SUB)
    y += Inches(0.32)

    # Section boxes (2x2) - adjusted for landscape height
    box_w = (SLIDE_W - MARGIN * 2 - Inches(0.2)) / 2
    available_h = SLIDE_H - y - Inches(0.35)  # leave room for footer
    box_h = (available_h - Inches(0.12)) / 2   # 2 rows with gap

    for i, (key, section_title, bg_color, placeholder) in enumerate(SECTIONS):
        col = i % 2
        row = i // 2
        bx = MARGIN + col * (box_w + Inches(0.2))
        by = y + row * (box_h + Inches(0.12))

        add_rounded_rect(slide, bx, by, box_w, box_h, bg_color)
        # Orange left accent bar
        add_rect(slide, bx, by, Inches(0.06), box_h, C_ORANGE)

        # Section title
        add_textbox(slide,
                    bx + Inches(0.12), by + Inches(0.08),
                    box_w - Inches(0.18), Inches(0.28),
                    section_title,
                    font_name=FONT_BODY, font_size_pt=12,
                    font_color=C_DARK, bold=True)

        # Divider
        add_rect(slide,
                 bx + Inches(0.12), by + Inches(0.38),
                 box_w - Inches(0.24), Inches(0.02),
                 C_ORANGE)

        # Content
        content = data.get(key, "") or ""
        display_text = content if content.strip() else f"（{placeholder}）"
        text_color = C_DARK if content.strip() else C_SUB

        add_textbox(slide,
                    bx + Inches(0.12), by + Inches(0.46),
                    box_w - Inches(0.24), box_h - Inches(0.55),
                    display_text,
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=text_color,
                    word_wrap=True)

    add_footer(slide)
