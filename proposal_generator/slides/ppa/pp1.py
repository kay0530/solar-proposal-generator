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
    add_card_with_accent, add_footer, add_header_bar, add_icon, add_rect,
    add_textbox,
    grid_x, grid_w, vstack,
)

TITLE = "なぜ今「オンサイトPPA」なのか？"
EYEBROW = "01｜導入の背景"

STANDFIRST = (
    "パリ協定の「1.5℃目標」を起点に、日本も2050年カーボンニュートラル、"
    "2035年度の温室効果ガス60％削減（2013年度比）を掲げ、脱炭素への移行は"
    "もはや後戻りしない潮流となりました。一方で燃料費・再エネ賦課金・"
    "容量拠出金などにより電気料金は上昇基調が続き、取引先からの"
    "サプライチェーン排出削減要請も強まっています。"
    "自ら電気をつくり、賢く使う——その第一歩がオンサイトPPAです。"
)

CARDS = [
    ("01", "yen", "電気代削減",
     "再エネ電気を長期固定単価で利用でき、市場価格の変動に左右されない"
     "安定した電力調達と停電リスクの軽減が可能になります。"),
    ("02", "zero_yen", "初期費用ゼロ",
     "太陽光パネルなど全設備の設置費用・維持管理費用はPPA事業者が負担。"
     "お客様の初期投資は不要です。"),
    ("03", "leaf", "CO₂削減・脱炭素経営",
     "再エネの自家消費でCO₂排出量を削減。取引先からのScope2排出削減要請や"
     "CDP・SBT等の開示要求にも、実需ベースの再エネ調達で応えられます。"),
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
    for i, (num, icon_name, card_title, body) in enumerate(CARDS):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, ys[1], w, band2_h)

        ey_y = cy + Inches(0.06)
        add_textbox(slide, cx, ey_y, cw, Inches(0.20),
                    num,
                    font_name=FONT_BLACK, font_size_pt=SIZE_CAPTION,
                    font_color=C_ORANGE, bold=True, tracking_pt=1.2)
        # flat line icon, top-right of the card
        add_icon(slide, icon_name,
                 cx + cw - Inches(0.46), ey_y, Inches(0.40))

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
