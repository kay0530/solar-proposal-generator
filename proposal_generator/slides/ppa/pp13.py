"""
pp13.py - FAQスライド

Layout (A4 landscape):
  - Orange header bar with "よくあるご質問"
  - 5 Q&A cards with orange accent
"""

from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_H, CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE,
    C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "よくあるご質問"

FAQ_ITEMS = [
    {
        "q": "初期費用はかかりますか？",
        "a": "PPAモデルでは初期費用ゼロでご導入いただけます。設備の設置費用・メンテナンス費用は全てPPA事業者が負担します。",
    },
    {
        "q": "メンテナンスは必要ですか？",
        "a": "設備の保守・点検・修理は全て当社が対応いたします。24時間遠隔監視により、異常の早期発見・迅速対応を行います。",
    },
    {
        "q": "契約期間中に移転したらどうなりますか？",
        "a": "移転先への設備の移設対応が可能です。移転先の条件を確認の上、最適なプランをご提案いたします。",
    },
    {
        "q": "停電時はどうなりますか？",
        "a": "蓄電池を併設することで、停電時にも非常用電源としてご利用いただけます。BCP対策としても有効です。",
    },
    {
        "q": "屋根が劣化しませんか？",
        "a": "設置前に防水処理を実施し、定期点検で屋根の状態を確認します。むしろパネルが直射日光を遮り、屋根の劣化を抑える効果もあります。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP13 (FAQ) onto an already-added blank slide.

    data keys used: (none required - static content)
    """
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.1)

    card_w = SLIDE_W - MARGIN * 2
    card_h = Inches(1.1)
    card_gap = Inches(0.12)

    q_badge_w = Inches(0.4)
    q_badge_h = Inches(0.35)
    a_badge_w = Inches(0.4)
    a_badge_h = Inches(0.35)

    for i, item in enumerate(FAQ_ITEMS):
        cy = y + i * (card_h + card_gap)

        # Card background
        add_rounded_rect(slide, MARGIN, cy, card_w, card_h, C_LIGHT_GRAY)

        # Orange left accent bar
        add_rect(slide, MARGIN, cy, Inches(0.07), card_h, C_ORANGE)

        # Q badge
        q_x = MARGIN + Inches(0.2)
        q_y = cy + Inches(0.12)
        add_rounded_rect(slide, q_x, q_y, q_badge_w, q_badge_h, C_ORANGE)
        add_textbox(slide,
                    q_x, q_y + Inches(0.04),
                    q_badge_w, Inches(0.27),
                    "Q",
                    font_name=FONT_BLACK, font_size_pt=16, font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Question text
        add_textbox(slide,
                    q_x + q_badge_w + Inches(0.12), q_y + Inches(0.04),
                    card_w - q_badge_w - Inches(0.6), Inches(0.3),
                    item["q"],
                    font_name=FONT_BODY, font_size_pt=12, font_color=C_DARK, bold=True)

        # A badge
        a_x = MARGIN + Inches(0.2)
        a_y = cy + Inches(0.55)
        add_rounded_rect(slide, a_x, a_y, a_badge_w, a_badge_h, C_LIGHT_ORANGE)
        add_textbox(slide,
                    a_x, a_y + Inches(0.04),
                    a_badge_w, Inches(0.27),
                    "A",
                    font_name=FONT_BLACK, font_size_pt=16, font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Answer text
        add_textbox(slide,
                    a_x + a_badge_w + Inches(0.12), a_y + Inches(0.02),
                    card_w - a_badge_w - Inches(0.6), Inches(0.4),
                    item["a"],
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB)

    add_footer(slide)
