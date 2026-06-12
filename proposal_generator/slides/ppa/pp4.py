"""
pp4.py - 現状の電気代分析 (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - v2 white header, eyebrow '01｜導入の背景'
  - Lead line: company context
  - 4 KPI cards (契約電力 / 月間使用量 / 月間電気代 / 年間電気代) at 28pt
  - Unit-price band (C_PANEL + orange bar + yen icon + 28pt number pair)
  - Trend card (accent-left) with cost-trend body copy
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_ORANGE, C_PANEL, C_SUB,
    FONT_BLACK, FONT_BODY, LINE_SPACING_BODY, MARGIN,
    SIZE_BODY, SIZE_CAPTION, SIZE_LEAD, SLIDE_W,
    add_card_with_accent, add_footer, add_header_bar, add_icon,
    add_kpi_card, add_multiline_textbox, add_number_unit, add_rect,
    add_section_header, add_textbox, fmt_num, fmt_yen,
    grid_w, grid_x, vstack,
)

TITLE = "現状の電気代分析"
EYEBROW = "01｜導入の背景"

TREND_NOTES = [
    "電気料金は過去10年で約30〜40%上昇しており、今後も上昇傾向が続く見込みです。",
    "再エネ賦課金・燃料費調整額の変動により、企業の電力コストは不安定化しています。",
    "固定単価のPPAモデルなら、将来の電力コスト上昇リスクを回避できます。",
    "自家消費型太陽光発電で系統電力への依存度を下げ、エネルギー自給率を高めます。",
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP4 (current electricity cost analysis) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    company = data.get("company_name", "") or ""
    contract_kw = data.get("contract_kw")
    monthly_kwh = data.get("monthly_kwh")
    monthly_cost = data.get("monthly_cost")
    annual_cost = data.get("annual_cost")
    unit_price = data.get("current_unit_price")

    # ---- Block heights for vertical justify ----
    lead_h = Inches(0.30)
    kpi_h = Inches(1.05)
    price_h = Inches(0.95)
    trend_card_h = Inches(1.40)
    trend_h = int(Inches(0.36)) + int(trend_card_h)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, kpi_h, price_h, trend_h])

    # ---- Lead: company context ----
    lead = f"{company}　様の現状電力コスト" if company else "現状の電力コスト"
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h,
                lead,
                font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                font_color=C_DARK, bold=True)

    # ---- 4 KPI cards (28pt, grid-snapped) ----
    kpis = [
        (fmt_num(contract_kw, 0) if contract_kw else "—", "kW", "契約電力"),
        (fmt_num(monthly_kwh, 0) if monthly_kwh else "—", "kWh/月", "月間使用量"),
        (fmt_yen(monthly_cost, "") if monthly_cost else "—", "円/月", "月間電気代"),
        (fmt_yen(annual_cost, "") if annual_cost else "—", "円/年", "年間電気代"),
    ]
    for i, (number, unit, label) in enumerate(kpis):
        add_kpi_card(slide, grid_x(i * 3), ys[1], grid_w(3), kpi_h,
                     number, unit, label)

    # ---- Current unit-price band (panel + orange bar + yen icon) ----
    band_y = ys[2]
    add_rect(slide, MARGIN, band_y, content_w, price_h, C_PANEL)
    add_rect(slide, MARGIN, band_y, Inches(0.05), price_h, C_ORANGE)

    icon_s = Inches(0.40)
    add_icon(slide, "yen",
             int(MARGIN) + int(Inches(0.25)),
             int(band_y) + (int(price_h) - int(icon_s)) // 2,
             icon_s)

    label_x = int(MARGIN) + int(Inches(0.90))
    price_label = "現在の電力単価" + ("" if unit_price else "（データ未入力）")
    add_textbox(slide, label_x, int(band_y) + int(Inches(0.14)),
                Inches(4.5), Inches(0.20),
                price_label,
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, bold=True)
    add_number_unit(slide, label_x, int(band_y) + int(Inches(0.32)),
                    Inches(4.5), int(price_h) - int(Inches(0.46)),
                    fmt_num(unit_price, 1) if unit_price else "—", "円/kWh")

    # ---- Electricity cost trend card ----
    sect_y = ys[3]
    add_section_header(slide, MARGIN, sect_y, content_w, "電気料金の動向")
    cx, cy, cw, ch = add_card_with_accent(
        slide, MARGIN, int(sect_y) + int(Inches(0.36)),
        content_w, trend_card_h, accent_position="left")

    lines = [(note, FONT_BODY, SIZE_BODY, C_DARK, False, PP_ALIGN.LEFT)
             for note in TREND_NOTES]
    text_h = Inches(0.95)
    ty = int(cy) + max(0, (int(ch) - int(text_h)) // 2)
    add_multiline_textbox(slide, cx, ty, cw, text_h,
                          lines, line_spacing=LINE_SPACING_BODY)

    add_footer(slide)
