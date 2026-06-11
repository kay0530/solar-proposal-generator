"""
pp1.py - なぜ今「オンサイトPPA」なのか？（static スライド）

Design v2 "Institutional Trust Grid":
  Band 1: standfirst lead paragraph (12.5pt, line_spacing 1.4)
  Band 2: 3 white benefit cards (grid cols 0-3 / 4-7 / 8-11)
  Band 3: conclusion band (C_PANEL + orange left bar)
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_ORANGE, C_PANEL,
    FONT_BLACK, FONT_BODY, GAP_IN_CARD, MARGIN, SLIDE_W,
    LINE_SPACING_BODY, LINE_SPACING_LEAD, SIZE_BODY, SIZE_CAPTION, SIZE_LEAD,
    add_card_with_accent, add_footer, add_header_bar, add_rect, add_textbox,
    grid_x, grid_w, vstack,
)

TITLE = "なぜ今「オンサイトPPA」なのか？"
EYEBROW = "01｜導入の背景"

STANDFIRST = (
    "2015年に採択されたパリ協定では、地球の気温上昇を産業革命前と比べて"
    "2℃未満に抑える目標が掲げられました。温室効果ガスの排出を減らす"
    "「低炭素化」にとどまらず、排出量実質ゼロを目指す「脱炭素化」への動きが"
    "世界で加速するなか、「脱炭素経営」はいまや企業の責務となりつつあります。"
    "自ら電気をつくり、賢く使う——その第一歩がオンサイトPPAです。"
)

CARDS = [
    ("01", "電気代削減",
     "再エネ電気を長期固定単価で利用でき、市場価格の変動に左右されない"
     "安定した電力調達と停電リスクの軽減が可能になります。"),
    ("02", "初期費用ゼロ",
     "太陽光パネルなど全設備の設置費用・維持管理費用はPPA事業者が負担。"
     "お客様の初期投資は不要です。"),
    ("03", "CO₂削減・脱炭素経営",
     "再生可能エネルギーの自家消費でCO₂排出量を削減し、"
     "カーボンニュートラル・脱炭素経営への取り組みを前進させます。"),
]

CONCLUSION = "いま導入する経済合理性があります"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    band1_h = Inches(0.95)
    band2_h = Inches(2.05)
    band3_h = Inches(0.90)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [band1_h, band2_h, band3_h])

    # --- Band 1: standfirst lead paragraph ---
    add_textbox(slide, MARGIN, ys[0], content_w, band1_h,
                STANDFIRST,
                font_name=FONT_BODY, font_size_pt=SIZE_LEAD,
                font_color=C_DARK, line_spacing=LINE_SPACING_LEAD,
                anchor=MSO_ANCHOR.TOP)

    # --- Band 2: 3 benefit cards ---
    for i, (num, card_title, body) in enumerate(CARDS):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, ys[1], w, band2_h)

        ey_y = cy + Inches(0.06)
        add_textbox(slide, cx, ey_y, cw, Inches(0.20),
                    num,
                    font_name=FONT_BLACK, font_size_pt=SIZE_CAPTION,
                    font_color=C_ORANGE, bold=True, tracking_pt=1.2)

        title_y = ey_y + Inches(0.20) + GAP_IN_CARD
        add_textbox(slide, cx, title_y, cw, Inches(0.28),
                    card_title,
                    font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                    font_color=C_DARK, bold=True)

        body_y = title_y + Inches(0.28) + GAP_IN_CARD
        add_textbox(slide, cx, body_y, cw, ch - Inches(0.78),
                    body,
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_DARK, line_spacing=LINE_SPACING_BODY)

    # --- Band 3: conclusion band ---
    add_rect(slide, MARGIN, ys[2], content_w, band3_h, C_PANEL)
    add_rect(slide, MARGIN, ys[2], Inches(0.06), band3_h, C_ORANGE)
    add_textbox(slide, MARGIN + Inches(0.30), ys[2],
                content_w - Inches(0.60), band3_h,
                CONCLUSION,
                font_name=FONT_BLACK, font_size_pt=12,
                font_color=C_NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    add_footer(slide)
