"""
pp12.py - 導入実績・事例紹介 (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - v2 white header, eyebrow '02｜オルテナジーの強み'
  - 当社の実績: 3 KPI cards (累計導入実績 / 累計設置容量 / 顧客満足度)
  - 導入事例: 3 accent-top case cards, savings as 20pt number+unit pairs,
    capacity / CO2 merged into one spec caption per card
"""

from __future__ import annotations

from pathlib import Path

from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_SUB,
    FONT_BLACK, FONT_BODY, LINE_SPACING_BODY, MARGIN,
    SIZE_CAPTION, SIZE_LEAD, SLIDE_W,
    add_card_with_accent, add_footer, add_header_bar, add_kpi_card,
    add_number_unit, add_section_header, add_textbox,
    grid_w, grid_x, vstack,
)

TITLE = "導入実績・事例紹介"
EYEBROW = "02｜オルテナジーの強み"

STATS = [
    ("500+", "件", "累計導入実績"),
    ("50+", "MW", "累計設置容量"),
    ("98", "%", "顧客満足度"),
]

CASE_STUDIES = [
    {
        "company": "製造業A社",
        "saving_num": "150",
        "saving_unit": "万円",
        "spec": "設備容量 100kW ｜ CO₂削減 48t/年",
        "detail": "工場屋根に太陽光パネルを設置。昼間の電力需要を自家消費で"
                  "カバーし、デマンドカットにも成功。",
    },
    {
        "company": "物流倉庫B社",
        "saving_num": "300",
        "saving_unit": "万円",
        "spec": "設備容量 200kW ｜ CO₂削減 96t/年",
        "detail": "大型倉庫の広い屋根を活用。冷蔵設備の電力を太陽光で補い、"
                  "電気代の大幅な削減を実現。",
    },
    {
        "company": "商業施設C社",
        "saving_num": "200",
        "saving_unit": "万円",
        "spec": "設備容量 150kW ｜ CO₂削減 72t/年",
        "detail": "ショッピングモール屋上に設置。来店客へのESGアピール効果も"
                  "高く、企業イメージの向上に貢献。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP12 (track record / case studies) onto a blank slide.

    data keys used: (none required - static content)
    """
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    # ---- Block heights for vertical justify ----
    kpi_h = Inches(1.00)
    block1_h = int(Inches(0.40)) + int(kpi_h)
    case_h = Inches(2.70)
    block2_h = int(Inches(0.40)) + int(case_h)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [block1_h, block2_h])

    # ---- Company stats: 3 KPI cards (28pt forced) ----
    add_section_header(slide, MARGIN, ys[0], content_w, "当社の実績")
    kpi_y = int(ys[0]) + int(Inches(0.40))
    for i, (number, unit, label) in enumerate(STATS):
        add_kpi_card(slide, grid_x(i * 4), kpi_y, grid_w(4), kpi_h,
                     number, unit, label)

    # ---- Case study cards ----
    add_section_header(slide, MARGIN, ys[1], content_w, "導入事例")
    cards_y = int(ys[1]) + int(Inches(0.40))

    for i, case in enumerate(CASE_STUDIES):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, cards_y, w, case_h)

        # Company name
        add_textbox(slide, cx, int(cy) + int(Inches(0.02)),
                    cw, Inches(0.26),
                    case["company"],
                    font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                    font_color=C_DARK, bold=True)

        # Annual saving (label + number/unit baseline pair)
        add_textbox(slide, cx, int(cy) + int(Inches(0.38)),
                    cw, Inches(0.18),
                    "年間削減額",
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, bold=True)
        add_number_unit(slide, cx, int(cy) + int(Inches(0.56)),
                        cw, Inches(0.42),
                        case["saving_num"], case["saving_unit"],
                        number_size_pt=20)

        # Spec caption (capacity + CO2)
        add_textbox(slide, cx, int(cy) + int(Inches(1.08)),
                    cw, Inches(0.20),
                    case["spec"],
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB)

        # Description
        add_textbox(slide, cx, int(cy) + int(Inches(1.38)),
                    cw, int(ch) - int(Inches(1.38)),
                    case["detail"],
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, line_spacing=LINE_SPACING_BODY)

    add_footer(slide)
