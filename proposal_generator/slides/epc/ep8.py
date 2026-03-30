"""
ep8.py - 実績・導入事例（EPC）

Shows track record and case studies for EPC installations.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_rect, add_rounded_rect,
    add_section_header, add_textbox,
    fmt_num,
)

TITLE = "実績・導入事例（EPC）"

CASE_STUDIES = [
    {
        "industry": "製造業",
        "capacity": "150 kW",
        "effect": "年間約180万円の電気代削減",
        "detail": "工場屋根に太陽光パネルを設置。補助金活用により投資回収7年を実現。"
                  "デマンドカット効果と合わせて大幅なコスト削減を達成。",
    },
    {
        "industry": "物流倉庫",
        "capacity": "300 kW",
        "effect": "年間約350万円の電気代削減",
        "detail": "広大な屋根面積を活用した大規模設置。即時償却を適用し初年度に"
                  "全額費用計上。CO₂排出量も年間150t削減。",
    },
    {
        "industry": "商業施設",
        "capacity": "80 kW",
        "effect": "年間約100万円の電気代削減",
        "detail": "店舗屋根への設置。RE100対応の一環として導入。"
                  "お客様への環境訴求にも活用されています。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP8 (track record / case studies for EPC) onto an already-added blank slide.

    data keys used: (none - static content with example cases)
    """
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    # ---- Summary KPI cards ----
    card_cols = 3
    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.18) * (card_cols - 1)) / card_cols
    card_h = Inches(1.0)

    summary_kpis = [
        ("500+", "件", "累計導入実績"),
        ("50MW+", "", "累計設置容量"),
        ("98%", "", "お客様満足度"),
    ]

    for i, (number, unit, label) in enumerate(summary_kpis):
        cx = MARGIN + i * (card_w + Inches(0.18))
        add_kpi_card(slide, cx, y, card_w, card_h,
                     number, unit, label,
                     number_size_pt=30)

    y += card_h + Inches(0.25)

    # ---- Case studies ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "導入事例")
    y += Inches(0.38)

    case_w = SLIDE_W - MARGIN * 2
    for case in CASE_STUDIES:
        case_h = Inches(1.2)
        add_rounded_rect(slide, MARGIN, y, case_w, case_h, C_LIGHT_GRAY)
        add_rect(slide, MARGIN, y, Inches(0.07), case_h, C_ORANGE)

        # Industry + capacity badge
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.08),
                    Inches(2.0), Inches(0.28),
                    case["industry"],
                    font_name=FONT_BLACK, font_size_pt=13,
                    font_color=C_DARK, bold=True)
        add_textbox(slide,
                    MARGIN + Inches(2.2), y + Inches(0.1),
                    Inches(1.5), Inches(0.22),
                    case["capacity"],
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_ORANGE, bold=True)

        # Effect highlight
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.38),
                    case_w - Inches(0.3), Inches(0.28),
                    case["effect"],
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_ORANGE, bold=True)

        # Detail
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.68),
                    case_w - Inches(0.3), Inches(0.48),
                    case["detail"],
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB)

        y += case_h + Inches(0.1)

    add_footer(slide)
