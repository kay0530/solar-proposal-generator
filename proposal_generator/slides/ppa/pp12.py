"""
pp12.py - 導入実績・事例紹介スライド

Layout (A4 landscape):
  - Orange header bar with "導入実績・事例紹介"
  - Company stats KPI cards (cumulative installs, MW, satisfaction)
  - 3 case study cards with details
"""

from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_H, CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE,
    C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_rect, add_rounded_rect,
    add_section_header, add_textbox,
)

TITLE = "導入実績・事例紹介"

STATS = [
    {"number": "500+", "unit": "件", "label": "累計導入実績"},
    {"number": "50+", "unit": "MW", "label": "累計設置容量"},
    {"number": "98", "unit": "%", "label": "顧客満足度"},
]

CASE_STUDIES = [
    {
        "company": "製造業A社",
        "capacity": "100kW",
        "saving": "年間150万円削減",
        "co2": "CO2 48t/年 削減",
        "detail": "工場屋根に太陽光パネルを設置。\n昼間の電力需要を自家消費でカバーし、\nデマンドカットにも成功。",
    },
    {
        "company": "物流倉庫B社",
        "capacity": "200kW",
        "saving": "年間300万円削減",
        "co2": "CO2 96t/年 削減",
        "detail": "大型倉庫の広い屋根を活用。\n冷蔵設備の電力を太陽光で補い、\n電気代の大幅な削減を実現。",
    },
    {
        "company": "商業施設C社",
        "capacity": "150kW",
        "saving": "年間200万円削減",
        "co2": "CO2 72t/年 削減",
        "detail": "ショッピングモール屋上に設置。\n来店客へのESGアピール効果も高く、\n企業イメージの向上に貢献。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP12 (track record / case studies) onto an already-added blank slide.

    data keys used: (none required - static content)
    """
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # ---- Company stats KPI cards ----
    add_section_header(slide, MARGIN, y, Inches(5.0), "当社の実績")
    y += Inches(0.4)

    stat_cols = 3
    gap = Inches(0.2)
    stat_w = (SLIDE_W - MARGIN * 2 - gap * (stat_cols - 1)) / stat_cols
    stat_h = Inches(1.15)

    for i, stat in enumerate(STATS):
        sx = MARGIN + i * (stat_w + gap)
        add_kpi_card(slide, sx, y, stat_w, stat_h,
                     stat["number"], stat["unit"], stat["label"],
                     bg_color=C_LIGHT_ORANGE,
                     number_size_pt=36)

    y += stat_h + Inches(0.3)

    # ---- Case study cards ----
    add_section_header(slide, MARGIN, y, Inches(5.0), "導入事例")
    y += Inches(0.45)

    card_cols = 3
    card_gap = Inches(0.18)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * (card_cols - 1)) / card_cols
    card_h = Inches(3.2)

    for i, case in enumerate(CASE_STUDIES):
        cx = MARGIN + i * (card_w + card_gap)

        # Card background
        add_rounded_rect(slide, cx, y, card_w, card_h, C_LIGHT_GRAY)

        # Orange top accent
        add_rect(slide, cx, y, card_w, Inches(0.07), C_ORANGE)

        # Company name
        add_textbox(slide,
                    cx + Inches(0.12), y + Inches(0.2),
                    card_w - Inches(0.24), Inches(0.3),
                    case["company"],
                    font_name=FONT_BLACK, font_size_pt=15, font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)

        # Capacity badge
        badge_w = Inches(1.4)
        badge_h = Inches(0.3)
        badge_x = cx + (card_w - badge_w) / 2
        add_rounded_rect(slide, badge_x, y + Inches(0.58), badge_w, badge_h, C_ORANGE)
        add_textbox(slide,
                    badge_x, y + Inches(0.61),
                    badge_w, Inches(0.24),
                    case["capacity"],
                    font_name=FONT_BODY, font_size_pt=12, font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Saving highlight
        add_textbox(slide,
                    cx + Inches(0.08), y + Inches(1.0),
                    card_w - Inches(0.16), Inches(0.3),
                    case["saving"],
                    font_name=FONT_BLACK, font_size_pt=14, font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)

        # CO2 reduction
        add_textbox(slide,
                    cx + Inches(0.08), y + Inches(1.3),
                    card_w - Inches(0.16), Inches(0.25),
                    case["co2"],
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB,
                    align=PP_ALIGN.CENTER)

        # Description
        add_textbox(slide,
                    cx + Inches(0.12), y + Inches(1.7),
                    card_w - Inches(0.24), Inches(1.3),
                    case["detail"],
                    font_name=FONT_BODY, font_size_pt=9, font_color=C_SUB)

    add_footer(slide)
