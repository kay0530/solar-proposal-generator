"""
pp7.py - ご利用料金に関して (PPA pricing)

Faithfully reproduces the Excel ご利用料金 slide:
- Title: ご利用料金に関して
- PPA explanation text
- Two large KPI: 初期ご負担金額 ¥0.- / 20年目までの一律単価 ¥XX.XX/kWh
- Four merit boxes: 再エネ賦課金, 燃料費等調整, 環境価値, 炭素税
- Footnotes
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, C_NAVY, C_LIGHT_TEAL,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, fmt_yen, fmt_num,
)

TITLE = "ご利用料金に関して"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP7 (PPA pricing) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)
    tax_display = data.get("tax_display", "税抜")
    ppa_price = data.get("ppa_unit_price", 0)
    years = int(data.get("contract_years", 20) or 20)

    # ---- PPA explanation ----
    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.22),
                "PPA（Power Purchase Agreement）",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_ORANGE, bold=True)
    y += Inches(0.28)

    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.40),
                "太陽光発電システムによる電力供給契約を言います。"
                "使用した電力量のみお支払いいただく契約です。",
                font_name=FONT_BODY, font_size_pt=10, font_color=C_DARK)
    y += Inches(0.40)

    # Tax note (right-aligned)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(2.5), y - Inches(0.05),
                Inches(2.5), Inches(0.20),
                f"金額は全て{tax_display}",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_SUB, align=PP_ALIGN.RIGHT)

    # ---- Two large KPI boxes ----
    kpi_gap = Inches(0.3)
    kpi_w = (SLIDE_W - MARGIN * 2 - kpi_gap) / 2
    kpi_h = Inches(1.8)

    # Left: 初期ご負担金額 (navy background, white text)
    add_rounded_rect(slide, MARGIN, y, kpi_w, kpi_h, C_NAVY)
    add_textbox(slide, MARGIN, y + Inches(0.20), kpi_w, Inches(0.28),
                "初期ご負担金額",
                font_name=FONT_BODY, font_size_pt=14,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, MARGIN, y + Inches(0.60), kpi_w, Inches(0.80),
                "¥0.-",
                font_name=FONT_BLACK, font_size_pt=36,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)

    # Right: 一律単価 (teal background, white text - wider)
    rx = MARGIN + kpi_w + kpi_gap
    add_rounded_rect(slide, rx, y, kpi_w, kpi_h, C_LIGHT_TEAL)
    add_textbox(slide, rx, y + Inches(0.20), kpi_w, Inches(0.28),
                f"{years}年目までの一律単価",
                font_name=FONT_BODY, font_size_pt=16,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    price_str = f"¥{ppa_price:.2f}" if ppa_price else "—"
    add_textbox(slide, rx, y + Inches(0.55), kpi_w, Inches(0.80),
                price_str,
                font_name=FONT_BLACK, font_size_pt=36,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, rx, y + Inches(1.30), kpi_w, Inches(0.28),
                "/kWh",
                font_name=FONT_BODY, font_size_pt=26,
                font_color=C_WHITE, align=PP_ALIGN.CENTER)

    y += kpi_h + Inches(0.25)

    # ---- Four merit boxes ----
    merits = [
        ("再エネ賦課金の上昇対策",
         "基本的には上昇傾向の再エネ賦課金。"
         "上がれば上がるほど自家消費によるメリットは大きくなります。"),
        ("燃料費等調整単価の上昇対策",
         "燃料費が高騰すればプラスに振れる調整額。"
         "再エネ賦課金同様、支払う必要がなくなります。"),
        ("環境価値（CO2排出量抑制）",
         "RE100達成には環境クレジットの購入が必須である企業も。"
         "市場調達する量を削減可能。"),
        ("炭素税対策",
         "2050年までの脱炭素化に向けて段階的に炭素税率を"
         "引き上げていくという計画があります。"),
    ]

    merit_gap = Inches(0.12)
    merit_w = (SLIDE_W - MARGIN * 2 - merit_gap * 3) / 4
    merit_h = Inches(1.8)

    for i, (title, desc) in enumerate(merits):
        mx = MARGIN + i * (merit_w + merit_gap)
        add_rounded_rect(slide, mx, y, merit_w, merit_h, C_LIGHT_GRAY)
        add_rect(slide, mx, y, merit_w, Inches(0.04), C_ORANGE)

        add_textbox(slide, mx + Inches(0.08), y + Inches(0.12),
                    merit_w - Inches(0.16), Inches(0.40),
                    title,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_ORANGE, bold=True)
        add_textbox(slide, mx + Inches(0.08), y + Inches(0.55),
                    merit_w - Inches(0.16), merit_h - Inches(0.65),
                    desc,
                    font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_DARK)

    y += merit_h + Inches(0.15)

    # ---- Footnotes ----
    notes = [
        "※ 通常の電力供給契約で発生するような、基本料金はかかりません。"
        "再生可能エネルギー促進賦課金および燃料費調整額のお支払いが不要となります。",
        "環境価値については貴社に帰属するものとし、上記ご提案単価は環境価値を含めた金額となっています。",
        "ご提案単価の有効期限はご提案日より1ヶ月間となります。",
    ]
    for note in notes:
        add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.18),
                    note,
                    font_name=FONT_BODY, font_size_pt=7, font_color=C_SUB)
        y += Inches(0.16)

    add_footer(slide)
