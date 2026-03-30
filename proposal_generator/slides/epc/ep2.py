"""
ep2.py - EPCモデルとは（static slide）

Explains that the customer purchases and owns the solar system.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "EPCモデルとは"

FLOW_STEPS = [
    ("①", "設備購入",   "お客様が設備を\n購入・所有"),
    ("②", "施工・設置", "EPC事業者が\n設計・施工を実施"),
    ("③", "自家消費",   "発電電力を\nお客様が100%利用"),
    ("④", "電気代削減", "電力会社への支払いが\n大幅に減少"),
]

FEATURES = [
    ("設備はお客様の資産", "太陽光パネル・PCS等の設備はお客様が購入し、自社資産として計上できます。"),
    ("減価償却メリット",   "法定耐用年数17年で減価償却。税制優遇で即時償却も可能です。"),
    ("長期コスト削減",     "PPA単価の支払いが不要。発電した電力は全て無料で利用できます。"),
    ("契約期間の自由度",   "PPAのような長期契約期間の縛りがなく、自由に運用できます。"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    # Flow diagram (horizontal arrow steps)
    step_w = (SLIDE_W - MARGIN * 2 - Inches(0.1) * 3) / 4
    step_h = Inches(1.3)
    for i, (num, label, desc) in enumerate(FLOW_STEPS):
        cx = MARGIN + i * (step_w + Inches(0.1))
        add_rounded_rect(slide, cx, y, step_w, step_h, C_ORANGE)
        add_textbox(slide,
                    cx, y + Inches(0.05),
                    step_w, Inches(0.38),
                    num,
                    font_name=FONT_BLACK, font_size_pt=20,
                    font_color=C_LIGHT_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide,
                    cx, y + Inches(0.4),
                    step_w, Inches(0.38),
                    label,
                    font_name=FONT_BLACK, font_size_pt=13,
                    font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide,
                    cx + Inches(0.08), y + Inches(0.75),
                    step_w - Inches(0.16), Inches(0.5),
                    desc,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_DARK,
                    align=PP_ALIGN.CENTER)

        # Arrow between steps
        if i < len(FLOW_STEPS) - 1:
            arrow_x = cx + step_w + Inches(0.01)
            add_textbox(slide,
                        arrow_x, y + step_h / 2 - Inches(0.15),
                        Inches(0.08), Inches(0.3),
                        "▶",
                        font_name=FONT_BODY, font_size_pt=12,
                        font_color=C_ORANGE,
                        align=PP_ALIGN.CENTER)

    y += step_h + Inches(0.25)

    # Feature rows
    for feat_title, feat_desc in FEATURES:
        add_rect(slide, MARGIN, y, Inches(0.06), Inches(0.55), C_ORANGE)
        add_textbox(slide,
                    MARGIN + Inches(0.14), y,
                    Inches(1.8), Inches(0.28),
                    feat_title,
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_DARK, bold=True)
        add_textbox(slide,
                    MARGIN + Inches(0.14), y + Inches(0.27),
                    SLIDE_W - MARGIN * 2 - Inches(0.14), Inches(0.28),
                    feat_desc,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB)
        y += Inches(0.62)

    add_footer(slide)
