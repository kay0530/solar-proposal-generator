"""
pp3.py - 導入メリット（static スライド）
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_kpi_card,
)

TITLE = "導入メリット"

MERITS = [
    {
        "number": "¥0",
        "unit": "初期費用",
        "label": "設置費用ゼロ",
        "detail": "太陽光パネル・PCSなど全設備の設置費用をPPA事業者が負担します。",
    },
    {
        "number": "固定",
        "unit": "PPA単価",
        "label": "長期安定コスト",
        "detail": "契約期間中のPPA単価が固定。電力市場の変動に左右されません。",
    },
    {
        "number": "CO₂",
        "unit": "削減",
        "label": "脱炭素・ESG対応",
        "detail": "再生可能エネルギーによる発電でCO₂排出量を大幅に削減できます。",
    },
    {
        "number": "24h",
        "unit": "モニタリング",
        "label": "維持管理不要",
        "detail": "設備の保守・点検・保険対応は全てPPA事業者が行います。",
    },
    {
        "number": "☀",
        "unit": "再エネ",
        "label": "BCP対応強化",
        "detail": "蓄電池との組み合わせで停電時の電力確保にも対応できます。",
    },
    {
        "number": "↓",
        "unit": "電気代",
        "label": "電力コスト削減",
        "detail": "現行電気料金よりも安い単価で電力を使用できます。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    card_cols = 3
    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.2) * (card_cols - 1)) / card_cols
    card_h = Inches(2.0)

    for i, merit in enumerate(MERITS):
        col = i % card_cols
        row = i // card_cols
        cx = MARGIN + col * (card_w + Inches(0.2))
        cy = y + row * (card_h + Inches(0.12))

        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE)
        add_rect(slide, cx, cy, card_w, Inches(0.07), C_ORANGE)

        # Big number/symbol
        add_textbox(slide,
                    cx, cy + Inches(0.12),
                    card_w, Inches(0.55),
                    merit["number"],
                    font_name=FONT_BLACK, font_size_pt=32,
                    font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        # Unit
        add_textbox(slide,
                    cx, cy + Inches(0.65),
                    card_w, Inches(0.2),
                    merit["unit"],
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB,
                    align=PP_ALIGN.CENTER)
        # Label
        add_textbox(slide,
                    cx + Inches(0.08), cy + Inches(0.88),
                    card_w - Inches(0.16), Inches(0.28),
                    merit["label"],
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)
        # Detail
        add_textbox(slide,
                    cx + Inches(0.1), cy + Inches(1.18),
                    card_w - Inches(0.2), Inches(0.72),
                    merit["detail"],
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB)

    add_footer(slide)
