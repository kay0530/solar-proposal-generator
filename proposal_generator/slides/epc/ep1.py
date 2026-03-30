"""
ep1.py - なぜ今EPCか（static slide）

Content focuses on asset ownership benefits, tax depreciation,
and long-term cost savings for EPC model.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header,
)

TITLE = "なぜ今「EPC（設備購入）」なのか？"

BODY_MAIN = (
    "電力コストの上昇が続く中、自家消費型太陽光発電の導入は企業の競争力強化に直結します。\n"
    "EPC（設備購入）モデルでは、お客様自身が設備を所有することで、長期にわたる電気代削減効果を\n"
    "最大限に享受できます。加えて、税制優遇（即時償却・税額控除）や補助金の活用により、\n"
    "実質的な投資負担を大幅に軽減することが可能です。"
)

BODY_SUB1 = "設備を「所有」することで得られる、PPAにはないメリットがあります。"
BODY_SUB2 = (
    "減価償却による節税効果、契約期間の制約なし、長期的なコスト削減効果の最大化。\n"
    "自社資産として太陽光発電を活用する時代です。"
)

POINTS = [
    ("資産所有",     "設備はお客様の資産に。\n減価償却で節税効果を享受"),
    ("長期コスト削減", "PPA単価の支払いが不要。\n発電分は全て電気代削減に"),
    ("税制優遇",     "中小企業経営強化税制等で\n即時償却・税額控除が可能"),
    ("補助金活用",   "国・自治体の補助金で\n初期投資を大幅に圧縮"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.1)

    # Main body text
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(1.4),
                BODY_MAIN,
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK)
    y += Inches(1.5)

    # Sub texts
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.4),
                BODY_SUB1,
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_ORANGE, bold=True)
    y += Inches(0.45)

    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.55),
                BODY_SUB2,
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_DARK)
    y += Inches(0.65)

    # Point cards (2x2 grid)
    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.15)) / 2
    card_h = Inches(1.35)
    for i, (point_title, point_body) in enumerate(POINTS):
        col = i % 2
        row = i // 2
        cx = MARGIN + col * (card_w + Inches(0.15))
        cy = y + row * (card_h + Inches(0.12))
        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE)
        # Orange top border accent
        add_rect(slide, cx, cy, card_w, Inches(0.06), C_ORANGE)
        add_textbox(slide,
                    cx + Inches(0.12), cy + Inches(0.1),
                    card_w - Inches(0.24), Inches(0.35),
                    point_title,
                    font_name=FONT_BLACK, font_size_pt=14,
                    font_color=C_ORANGE, bold=True)
        add_textbox(slide,
                    cx + Inches(0.12), cy + Inches(0.48),
                    card_w - Inches(0.24), Inches(0.75),
                    point_body,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK)

    add_footer(slide)
