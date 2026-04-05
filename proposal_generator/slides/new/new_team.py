"""
new_team.py - ご支援体制図 (Support Team Structure)

Dummy placeholder slide showing department columns with photo placeholders.
Departments: 営業, 設計・開発, 施工管理, O&M・保守, モニタリング
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_NAVY, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "ご支援体制図"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Subtitle
    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.45),
                "マーケティング～販売～支援～設計～開発～施工～分析～メンテナンス",
                font_name=FONT_BODY, font_size_pt=12, font_color=C_SUB,
                align=PP_ALIGN.CENTER)
    y += Inches(0.55)

    # Department columns
    departments = [
        ("営業", "Sales"),
        ("設計・開発", "Engineering"),
        ("施工管理", "Construction"),
        ("O&M・保守", "Maintenance"),
        ("モニタリング", "Monitoring"),
    ]

    n_cols = len(departments)
    col_gap = Inches(0.15)
    total_gap = col_gap * (n_cols - 1)
    col_w = (SLIDE_W - MARGIN * 2 - total_gap) / n_cols
    col_h = Inches(4.8)

    for i, (dept_ja, dept_en) in enumerate(departments):
        cx = MARGIN + i * (col_w + col_gap)

        # Column background card
        add_rounded_rect(slide, cx, y, col_w, col_h, C_LIGHT_GRAY)

        # Department header bar (navy)
        add_rect(slide, cx, y, col_w, Inches(0.45), C_NAVY)
        add_textbox(slide, cx, y + Inches(0.05), col_w, Inches(0.35),
                    dept_ja,
                    font_name=FONT_BLACK, font_size_pt=12,
                    font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)

        # English sub-label
        add_textbox(slide, cx, y + Inches(0.50), col_w, Inches(0.25),
                    dept_en,
                    font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

        # Photo placeholder circle (gray oval shape)
        circle_size = Inches(1.0)
        circle_x = cx + (col_w - circle_size) / 2
        circle_y = y + Inches(1.0)
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(circle_x), int(circle_y), int(circle_size), int(circle_size),
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_SUB
        shape.line.fill.background()

        # Photo icon text (person silhouette placeholder)
        add_textbox(slide, circle_x, circle_y + Inches(0.25),
                    circle_size, Inches(0.5),
                    "PHOTO",
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_WHITE, align=PP_ALIGN.CENTER)

        # Name placeholder
        add_textbox(slide, cx, y + Inches(2.2), col_w, Inches(0.35),
                    "担当者名",
                    font_name=FONT_BLACK, font_size_pt=13,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)

        # Role placeholder
        add_textbox(slide, cx, y + Inches(2.55), col_w, Inches(0.30),
                    "役職",
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

        # Orange accent line under role
        line_w = col_w * 0.6
        line_x = cx + (col_w - line_w) / 2
        add_rect(slide, line_x, y + Inches(2.90), line_w, Inches(0.03), C_ORANGE)

        # Responsibility description placeholder
        add_textbox(slide, cx + Inches(0.1), y + Inches(3.1),
                    col_w - Inches(0.2), Inches(1.5),
                    "担当業務の説明を\nここに記入",
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

    # Footer note
    note_y = y + col_h + Inches(0.15)
    add_textbox(slide, MARGIN, note_y, SLIDE_W - MARGIN * 2, Inches(0.30),
                "※ 詳細は後日記入",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_SUB, align=PP_ALIGN.RIGHT)

    add_footer(slide)
