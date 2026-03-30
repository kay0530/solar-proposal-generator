"""
new_fip.py - FIP制度の活用スライド

Explains the Feed-in Premium (FIP) scheme and shows projected revenue
from surplus electricity sold under FIP.

Layout: A4 landscape
- Header bar
- What is FIP section (explanation)
- KPI cards: FIPプレミアム単価, 想定売電量, 年間FIP収入
- Comparison bar: 自家消費メリット vs FIP売電収入
- Notes: FIP認定要件, バランシングコスト
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_TOP,
    C_DARK,
    C_LIGHT_GRAY,
    C_LIGHT_ORANGE,
    C_ORANGE,
    C_SUB,
    C_TEAL,
    C_WHITE,
    FONT_BLACK,
    FONT_BODY,
    MARGIN,
    SLIDE_H,
    SLIDE_W,
    add_footer,
    add_header_bar,
    add_rect,
    add_rounded_rect,
    add_section_header,
    add_textbox,
    fmt_num,
    fmt_yen,
)

TITLE = "FIP制度の活用"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render the FIP slide onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # ------------------------------------------------------------------
    # Section 1: What is FIP
    # ------------------------------------------------------------------
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "FIP（フィードインプレミアム）制度とは", font_size_pt=13)
    y += Inches(0.35)

    # Explanation box
    explanation_lines = [
        "FIP制度は、再エネ電気を市場で売電する際に、市場価格に一定のプレミアム（補助額）を",
        "上乗せして収入を得られる制度です。2022年4月に開始され、FIT制度の後継として位置づけ",
        "られています。自家消費で賄いきれない余剰電力をFIPで売電することで、収益を最大化できます。",
    ]
    explanation_text = "\n".join(explanation_lines)

    add_rounded_rect(slide, MARGIN, y,
                     SLIDE_W - MARGIN * 2, Inches(0.88),
                     C_LIGHT_GRAY)
    add_textbox(slide, MARGIN + Inches(0.15), y + Inches(0.08),
                SLIDE_W - MARGIN * 2 - Inches(0.3), Inches(0.78),
                explanation_text,
                font_name=FONT_BODY, font_size_pt=10.5,
                font_color=C_DARK)
    y += Inches(1.0)

    # ------------------------------------------------------------------
    # Section 2: Key metrics (3 KPI cards)
    # ------------------------------------------------------------------
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "FIP売電試算", font_size_pt=13)
    y += Inches(0.35)

    # Extract data
    fip_premium = data.get("fip_premium_yen_per_kwh", 0)
    market_price = data.get("fip_market_price", 12.0)
    surplus_kwh = data.get("surplus_kwh", 0)
    fip_gross = data.get("fip_gross_revenue", 0)
    fip_balancing = data.get("fip_balancing_cost", 0)
    fip_net = data.get("fip_net_revenue", 0)

    # If we have the inputs but not calculated values, compute them
    if surplus_kwh > 0 and fip_premium > 0 and fip_gross == 0:
        fip_gross = surplus_kwh * (market_price + fip_premium)
        fip_balancing = surplus_kwh * 1.0  # default balancing rate
        fip_net = fip_gross - fip_balancing

    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.18) * 2) / 3
    card_h = Inches(1.15)
    kpi_data = [
        (fmt_num(fip_premium, 1), "円/kWh", "FIPプレミアム単価"),
        (f"{surplus_kwh:,.0f}" if surplus_kwh else "—", "kWh/年", "想定売電量（余剰分）"),
        (fmt_yen(fip_net, ""), "円/年", "年間FIP収入（税引前）"),
    ]

    for i, (number, unit, label) in enumerate(kpi_data):
        cx = MARGIN + i * (card_w + Inches(0.18))
        cy = y

        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE)
        add_rect(slide, cx, cy, card_w, Inches(0.06), C_ORANGE)
        add_textbox(slide, cx, cy + Inches(0.1), card_w, Inches(0.42),
                    number,
                    font_name=FONT_BLACK, font_size_pt=26,
                    font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, cx, cy + Inches(0.52), card_w, Inches(0.2),
                    unit,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, cx + Inches(0.08), cy + card_h - Inches(0.28),
                    card_w - Inches(0.16), Inches(0.24),
                    label,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)

    y += card_h + Inches(0.15)

    # ------------------------------------------------------------------
    # Section 3: Self-consumption vs FIP comparison
    # ------------------------------------------------------------------
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "自家消費メリット vs FIP売電収入", font_size_pt=13)
    y += Inches(0.35)

    self_consumption_saving = data.get("annual_cost_saving", 0) or 0
    total_benefit = self_consumption_saving + (fip_net or 0)

    bar_w = SLIDE_W - MARGIN * 2
    bar_h = Inches(0.75)

    # Two-column comparison
    half_w = (bar_w - Inches(0.15)) / 2

    # Left: self-consumption
    add_rounded_rect(slide, MARGIN, y, half_w, bar_h, C_LIGHT_ORANGE)
    add_textbox(slide, MARGIN + Inches(0.1), y + Inches(0.06),
                half_w - Inches(0.2), Inches(0.2),
                "自家消費メリット",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK, bold=True)
    add_textbox(slide, MARGIN + Inches(0.1), y + Inches(0.3),
                half_w - Inches(0.2), Inches(0.35),
                fmt_yen(self_consumption_saving) + "/年",
                font_name=FONT_BLACK, font_size_pt=20,
                font_color=C_ORANGE, bold=True)

    # Right: FIP revenue
    right_x = MARGIN + half_w + Inches(0.15)
    add_rounded_rect(slide, right_x, y, half_w, bar_h, C_LIGHT_GRAY)
    add_rect(slide, right_x, y, half_w, Inches(0.05), C_TEAL)
    add_textbox(slide, right_x + Inches(0.1), y + Inches(0.06),
                half_w - Inches(0.2), Inches(0.2),
                "FIP売電収入（余剰分）",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK, bold=True)
    add_textbox(slide, right_x + Inches(0.1), y + Inches(0.3),
                half_w - Inches(0.2), Inches(0.35),
                fmt_yen(fip_net) + "/年",
                font_name=FONT_BLACK, font_size_pt=20,
                font_color=C_TEAL, bold=True)

    y += bar_h + Inches(0.1)

    # Total benefit
    add_rounded_rect(slide, MARGIN, y, bar_w, Inches(0.45), C_ORANGE)
    add_textbox(slide, MARGIN + Inches(0.15), y + Inches(0.05),
                Inches(3.5), Inches(0.35),
                "合計年間メリット",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_WHITE, bold=True)
    add_textbox(slide, MARGIN + Inches(4.0), y + Inches(0.02),
                bar_w - Inches(4.5), Inches(0.4),
                fmt_yen(total_benefit) + "/年",
                font_name=FONT_BLACK, font_size_pt=22,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.RIGHT)
    y += Inches(0.6)

    # ------------------------------------------------------------------
    # Section 4: Notes
    # ------------------------------------------------------------------
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "留意事項", font_size_pt=11)
    y += Inches(0.3)

    notes = [
        "・FIP認定の取得が必要です（経済産業省への申請手続き）",
        "・バランシングコスト（発電予測と実績の差分ペナルティ）が発生します"
        f"（本試算では {data.get('fip_balancing_rate', 1.0):.1f} 円/kWh で概算）",
        "・市場価格は変動するため、実際の収入は上記試算と異なる場合があります"
        f"（本試算の想定市場価格: {market_price:.1f} 円/kWh）",
    ]
    for note in notes:
        add_textbox(slide, MARGIN + Inches(0.1), y,
                    SLIDE_W - MARGIN * 2 - Inches(0.2), Inches(0.22),
                    note,
                    font_name=FONT_BODY, font_size_pt=8.5,
                    font_color=C_SUB)
        y += Inches(0.2)

    add_footer(slide)
