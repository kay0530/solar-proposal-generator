"""
pp4.py - 現状電気代分析 (Current electricity cost analysis)

Layout: A4 landscape
- Header bar
- 4 KPI cards: 契約電力, 月間使用量, 月間電気代, 年間電気代
- Note about rising electricity costs
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, fmt_yen, fmt_num,
)

TITLE = "現状の電気代分析"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP4 (current electricity cost analysis) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    company = data.get("company_name", "") or ""
    y = CONTENT_TOP + Inches(0.1)

    # Company context
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.32),
                f"{company}　様の現状電力コスト",
                font_name=FONT_BODY, font_size_pt=13,
                font_color=C_DARK, bold=True)
    y += Inches(0.42)

    # ---- KPI cards (1 row x 4 cols) ----
    contract_kw = data.get("contract_kw")
    monthly_kwh = data.get("monthly_kwh")
    monthly_cost = data.get("monthly_cost")
    annual_cost = data.get("annual_cost")

    kpis = [
        (fmt_num(contract_kw, 0) if contract_kw else "—", "kW", "契約電力"),
        (fmt_num(monthly_kwh, 0) if monthly_kwh else "—", "kWh/月", "月間使用量"),
        (fmt_yen(monthly_cost, "") if monthly_cost else "—", "円/月", "月間電気代"),
        (fmt_yen(annual_cost, "") if annual_cost else "—", "円/年", "年間電気代"),
    ]

    card_cols = 4
    gap = Inches(0.15)
    card_w = (SLIDE_W - MARGIN * 2 - gap * (card_cols - 1)) / card_cols
    card_h = Inches(1.5)

    for i, (number, unit, label) in enumerate(kpis):
        cx = MARGIN + i * (card_w + gap)
        cy = y

        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE)
        add_rect(slide, cx, cy, card_w, Inches(0.06), C_ORANGE)
        # Number
        add_textbox(slide, cx, cy + Inches(0.15), card_w, Inches(0.55),
                    number,
                    font_name=FONT_BLACK, font_size_pt=30,
                    font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        # Unit
        add_textbox(slide, cx, cy + Inches(0.68), card_w, Inches(0.22),
                    unit,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)
        # Label
        add_textbox(slide, cx + Inches(0.06), cy + card_h - Inches(0.32),
                    card_w - Inches(0.12), Inches(0.28),
                    label,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)

    y += card_h + Inches(0.25)

    # ---- Current unit price section ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "現在の電力単価", font_size_pt=12)
    y += Inches(0.38)

    unit_price_info = data.get("current_unit_price")
    price_text = f"現在の電力単価：{fmt_num(unit_price_info, 1)} 円/kWh" if unit_price_info else "現在の電力単価：データ未入力"
    add_textbox(slide, MARGIN + Inches(0.1), y,
                SLIDE_W - MARGIN * 2 - Inches(0.1), Inches(0.3),
                price_text,
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_DARK)
    y += Inches(0.45)

    # ---- Electricity cost trend note ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "電気料金の動向", font_size_pt=12)
    y += Inches(0.38)

    # Warning box about rising costs
    box_w = SLIDE_W - MARGIN * 2
    box_h = Inches(1.8)
    add_rounded_rect(slide, MARGIN, y, box_w, box_h, C_LIGHT_GRAY)

    notes = [
        "■  電気料金は過去10年で約30〜40%上昇しており、今後も上昇傾向が続く見込みです。",
        "■  再エネ賦課金・燃料費調整額の変動により、企業の電力コストは不安定化しています。",
        "■  固定単価のPPAモデルを導入することで、将来の電力コスト上昇リスクを回避できます。",
        "■  自家消費型太陽光発電により、系統電力への依存度を下げ、エネルギー自給率を向上させます。",
    ]
    for i, note in enumerate(notes):
        add_textbox(slide, MARGIN + Inches(0.15), y + Inches(0.12) + i * Inches(0.38),
                    box_w - Inches(0.3), Inches(0.35),
                    note,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK)

    add_footer(slide)
