"""
pp1.py - なぜ今「オンサイトPPA」なのか？（static スライド）

Layout (A4 landscape): 4 point cards in a single horizontal row.
Content is fixed (no customer-specific data).
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

TITLE = "なぜ今「オンサイトPPA」なのか？"

BODY_MAIN = (
    "温室効果ガスとは気候変動（地球温暖化）の主要な原因となる二酸化炭素です。\n"
    "2015年に採択されたパリ協定では、これらを出来るだけ食い止めるために、地球の気温上昇を産業革命前と比べて2℃未満に\n"
    "抑えるという目標が設定されました。しかしながら、温室効果ガスの排出を削減する「低炭素化」だけでは目標達成が難しい\n"
    "ことから、温室効果ガスの排出量実質ゼロを目指す「脱炭素化」の動きが世界で加速しています。"
)

BODY_SUB1 = "今、わたしたちにできることは何か？「脱炭素経営」が企業にも求められています。"
BODY_SUB2 = (
    "まずはできるところから。自分たちで電気を作り、不足したときにお互いに融通する時代へ\n"
    "わたしたちは一歩を踏み出します。"
)

POINTS = [
    ("電気代削減",   "再エネ電気を安定した単価で\n長期利用できます"),
    ("CO₂削減",      "再生可能エネルギーで\nカーボンニュートラルに貢献"),
    ("初期費用ゼロ", "設備設置・維持管理費用は\nPPA事業者が負担"),
    ("電力安定供給", "自家消費で停電リスクを\n軽減できます"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Main body text
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(1.1),
                BODY_MAIN,
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK)
    y += Inches(1.15)

    # Sub texts
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.35),
                BODY_SUB1,
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_ORANGE, bold=True)
    y += Inches(0.38)

    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.45),
                BODY_SUB2,
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_DARK)
    y += Inches(0.5)

    # Point cards (1x4 horizontal row - landscape has more width)
    card_cols = 4
    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.15) * (card_cols - 1)) / card_cols
    card_h = Inches(1.35)
    for i, (point_title, point_body) in enumerate(POINTS):
        col = i % card_cols
        cx = MARGIN + col * (card_w + Inches(0.15))
        cy = y
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
