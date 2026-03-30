"""
pp4a.py - なぜオルテナジーが選ばれるのか (Why Altenergy is chosen)

PDF P5: Three strengths with explanation cards.
- 競争力のある単価
- ワンストップサービス
- 豊富な実績
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "なぜオルテナジーが選ばれるのか"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Intro text
    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.50),
                "PPAにおいて実績を伸ばしている企業は実はそれほど多くはありません。\n"
                "オルテナジーが実績を伸ばすことが出来ている理由は３つの強みがあるからです。",
                font_name=FONT_BODY, font_size_pt=11, font_color=C_DARK)
    y += Inches(0.60)

    # Three strength cards
    strengths = [
        {
            "num": "1",
            "title": "競争力のある単価",
            "body": (
                "単価を決める要素は大きく外的要因と内的要因の２つがあります。\n\n"
                "外的要因：リース金利・パネル原価等の仕入れコスト\n"
                "内的要因：施工コスト・管理コスト等の社内コスト\n\n"
                "オルテナジーは自社施工体制と効率的なオペレーションにより、"
                "両方のコストを最適化し、競争力のある単価を実現しています。"
            ),
        },
        {
            "num": "2",
            "title": "ワンストップサービス",
            "body": (
                "設計・施工・メンテナンス・発電事業運営までを一貫して自社グループで実施。\n\n"
                "お客様の窓口は一つ。設計変更や障害対応も迅速に行えます。\n\n"
                "EPC（設計・調達・建設）事業者としての実績と、"
                "発電事業者としてのノウハウの両方を持つ強みがあります。"
            ),
        },
        {
            "num": "3",
            "title": "豊富な実績",
            "body": (
                "累計設置容量100MW以上、導入企業数200社以上の実績。\n\n"
                "工場・倉庫・商業施設・オフィスビルなど、"
                "多様な建物への導入実績があります。\n\n"
                "全国対応可能なネットワークで、"
                "北海道から沖縄まで日本全国のお客様にサービスを提供しています。"
            ),
        },
    ]

    card_gap = Inches(0.15)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * 2) / 3
    card_h = Inches(4.8)

    for i, s in enumerate(strengths):
        cx = MARGIN + i * (card_w + card_gap)

        # Card background
        add_rounded_rect(slide, cx, y, card_w, card_h, C_LIGHT_GRAY)
        # Orange top bar
        add_rect(slide, cx, y, card_w, Inches(0.06), C_ORANGE)

        # Number circle
        num_size = Inches(0.55)
        add_rounded_rect(slide, cx + (card_w - num_size) / 2, y + Inches(0.20),
                         num_size, num_size, C_ORANGE)
        add_textbox(slide, cx + (card_w - num_size) / 2, y + Inches(0.22),
                    num_size, num_size,
                    s["num"],
                    font_name=FONT_BLACK, font_size_pt=24,
                    font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)

        # Title
        add_textbox(slide, cx + Inches(0.1), y + Inches(0.85),
                    card_w - Inches(0.2), Inches(0.35),
                    s["title"],
                    font_name=FONT_BLACK, font_size_pt=13,
                    font_color=C_ORANGE, bold=True, align=PP_ALIGN.CENTER)

        # Body
        add_textbox(slide, cx + Inches(0.15), y + Inches(1.25),
                    card_w - Inches(0.3), card_h - Inches(1.45),
                    s["body"],
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_DARK)

    add_footer(slide)
