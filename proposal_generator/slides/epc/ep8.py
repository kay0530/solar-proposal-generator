"""
ep8.py - 実績・導入事例（EPC） (design v2: Institutional Trust Grid)

Layout (A4 landscape), eyebrow '07｜導入実績':
  - Three summary KPI cards (28pt): cumulative installs / MW / satisfaction
  - Section '導入事例': three full-width left-accent cards, each with
    industry name, capacity (number+unit pair), annual-saving figure
    (20pt number+unit) and a wrapped detail paragraph
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_SUB,
    FONT_BLACK, FONT_BODY, GAP_CARD, MARGIN, SIZE_CAPTION, SLIDE_W,
    add_card_with_accent, add_footer, add_header_bar, add_kpi_card,
    add_number_unit, add_section_header, add_textbox, grid_w, grid_x, vstack,
)

TITLE = "実績・導入事例（EPC）"
EYEBROW = "07｜導入実績"

SUMMARY_KPIS = [
    ("500+", "件", "累計導入実績"),
    ("50+", "MW", "累計設置容量"),
    ("98", "%", "お客様満足度"),
]

# Case studies: capacity (kW), annual saving (万円/年), detail paragraph
CASE_STUDIES = [
    {
        "industry": "製造業",
        "capacity": "150",
        "saving": "約180",
        "detail": "工場屋根に太陽光パネルを設置。補助金活用により投資回収7年を実現。"
                  "デマンドカット効果と合わせて大幅なコスト削減を達成。",
    },
    {
        "industry": "物流倉庫",
        "capacity": "300",
        "saving": "約350",
        "detail": "広大な屋根面積を活用した大規模設置。即時償却を適用し初年度に"
                  "全額費用計上。CO₂排出量も年間150t削減。",
    },
    {
        "industry": "商業施設",
        "capacity": "80",
        "saving": "約100",
        "detail": "店舗屋根への設置。RE100対応の一環として導入。"
                  "お客様への環境訴求にも活用されています。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render EP8 (EPC track record / case studies). data keys used: none."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    # ---- Block heights for vertical justify ----
    kpi_h = Inches(1.00)
    case_h = Inches(1.25)
    n_cases = len(CASE_STUDIES)
    sect_h = (int(Inches(0.34)) + int(case_h) * n_cases
              + int(GAP_CARD) * (n_cases - 1))

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [kpi_h, sect_h],
                min_gap=GAP_CARD)

    # ---- Summary KPI cards (28pt, fixed by add_kpi_card) ----
    kpi_y = ys[0]
    for i, (number, unit, label) in enumerate(SUMMARY_KPIS):
        add_kpi_card(slide, grid_x(i * 4), kpi_y, grid_w(4), kpi_h,
                     number, unit, label)

    # ---- Case studies ----
    sect_y = ys[1]
    add_section_header(slide, MARGIN, sect_y, content_w, "導入事例")

    card_y = int(sect_y) + int(Inches(0.34))
    for case in CASE_STUDIES:
        cx, cy, cw, ch = add_card_with_accent(
            slide, MARGIN, card_y, content_w, case_h,
            accent_position="left")

        # Col 1: industry + system capacity
        add_textbox(slide, cx, cy + int(Inches(0.04)),
                    Inches(1.9), Inches(0.28),
                    case["industry"],
                    font_name=FONT_BLACK, font_size_pt=12.5,
                    font_color=C_DARK, bold=True)
        add_number_unit(slide, cx, int(cy) + int(ch) - int(Inches(0.46)),
                        Inches(1.9), Inches(0.42),
                        case["capacity"], "kW", number_size_pt=16)

        # Col 2: annual electricity-cost saving
        eff_x = int(cx) + int(Inches(2.1))
        add_textbox(slide, eff_x, cy + int(Inches(0.04)),
                    Inches(2.4), Inches(0.20),
                    "年間電気代削減（概算）",
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, bold=True)
        add_number_unit(slide, eff_x, int(cy) + int(ch) - int(Inches(0.52)),
                        Inches(2.4), Inches(0.48),
                        case["saving"], "万円/年", number_size_pt=20)

        # Col 3: detail paragraph (wrapped, no manual breaks)
        det_x = int(cx) + int(Inches(4.8))
        det_w = int(cx) + int(cw) - det_x
        add_textbox(slide, det_x, cy, det_w, ch,
                    case["detail"],
                    font_name=FONT_BODY, font_size_pt=9.5,
                    font_color=C_SUB, line_spacing=1.35,
                    anchor=MSO_ANCHOR.MIDDLE)

        card_y += int(case_h) + int(GAP_CARD)

    add_footer(slide)
