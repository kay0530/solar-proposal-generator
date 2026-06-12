"""
new_fip.py - FIP制度の活用スライド (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - Section: what FIP is (body lead, no manual line wrapping)
  - Metric band: 3 KPI cards 28pt (premium / surplus volume / net revenue)
    + calc detail strip (gross − balancing cost = net) when computable
  - Benefit row: self-consumption KPI + FIP revenue KPI + total panel band
    (C_PANEL + orange left bar + 28pt number, pp9-style)
  - Notes: FIP certification, balancing cost, market price assumption
All fip_* bindings are preserved; gross / balancing now surface in the
calc strip instead of staying invisible.
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_ORANGE, C_PANEL, C_SUB,
    FONT_BODY, GAP_CARD, MARGIN, SIZE_BODY, SIZE_CAPTION, SIZE_SMALL,
    SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_multiline_textbox,
    add_number_unit, add_rect, add_section_header, add_textbox,
    fmt_num, fmt_yen, grid_w, grid_x, vstack,
)

TITLE = "FIP制度の活用"
EYEBROW = "03｜効果シミュレーション"

_EXPLANATION = (
    "FIP制度は、再エネ電気を市場で売電する際に、市場価格に一定のプレミアム"
    "（補助額）を上乗せして収入を得られる制度です。2022年4月にFIT制度の後継"
    "として開始されました。自家消費で賄いきれない余剰電力をFIPで売電する"
    "ことで、収益を最大化できます。"
)


def _f(data: dict, key: str, default: float = 0.0) -> float:
    """Read a numeric value from data, tolerating None / bad types."""
    try:
        v = data.get(key, default)
        return float(v) if v is not None else float(default)
    except (TypeError, ValueError):
        return float(default)


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render the FIP slide onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)
    content_w = SLIDE_W - MARGIN * 2

    # ---- Data (all original bindings preserved) ----
    fip_premium = _f(data, "fip_premium_yen_per_kwh", 0)
    market_price = _f(data, "fip_market_price", 12.0)
    surplus_kwh = _f(data, "surplus_kwh", 0)
    fip_gross = _f(data, "fip_gross_revenue", 0)
    fip_balancing = _f(data, "fip_balancing_cost", 0)
    fip_net = _f(data, "fip_net_revenue", 0)
    balancing_rate = _f(data, "fip_balancing_rate", 1.0)

    # If we have the inputs but not calculated values, compute them
    if surplus_kwh > 0 and fip_premium > 0 and fip_gross == 0:
        fip_gross = surplus_kwh * (market_price + fip_premium)
        fip_balancing = surplus_kwh * balancing_rate
        fip_net = fip_gross - fip_balancing

    self_saving = _f(data, "annual_cost_saving", 0)
    total_benefit = self_saving + fip_net
    has_detail = fip_gross > 0

    # ---- Block heights for vertical justify ----
    explain_h = int(Inches(0.36)) + int(Inches(0.56))
    kpi_block_h = (int(Inches(0.36)) + int(Inches(1.00))
                   + (int(Inches(0.30)) if has_detail else 0))
    benefit_block_h = int(Inches(0.36)) + int(Inches(1.05))
    notes_block_h = int(Inches(0.32)) + int(Inches(0.60))

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [explain_h, kpi_block_h, benefit_block_h, notes_block_h],
                min_gap=GAP_CARD)

    # ---- Section 1: What is FIP ----
    add_section_header(slide, MARGIN, ys[0], content_w,
                       "FIP（フィードインプレミアム）制度とは")
    add_textbox(slide, MARGIN, int(ys[0]) + int(Inches(0.36)),
                content_w, Inches(0.56),
                _EXPLANATION,
                font_size_pt=SIZE_BODY, line_spacing=1.35)

    # ---- Section 2: FIP KPI band (28pt cards) ----
    add_section_header(slide, MARGIN, ys[1], content_w, "FIP売電試算")
    card_y = int(ys[1]) + int(Inches(0.36))
    card_h = Inches(1.00)
    kpi_items = [
        (fmt_num(fip_premium, 1) if fip_premium else "—",
         "円/kWh", "FIPプレミアム単価"),
        (f"{surplus_kwh:,.0f}" if surplus_kwh else "—",
         "kWh/年", "想定売電量（余剰分）"),
        (fmt_yen(fip_net, "") if fip_net else "—",
         "円/年", "年間FIP収入（税引前）"),
    ]
    for i, (number, unit, label) in enumerate(kpi_items):
        add_kpi_card(slide, grid_x(i * 4), card_y, grid_w(4), card_h,
                     number, unit, label)

    if has_detail:
        add_textbox(slide, MARGIN,
                    card_y + int(card_h) + int(Inches(0.08)),
                    content_w, Inches(0.20),
                    f"内訳：売電収入 {fmt_yen(fip_gross)}"
                    f"（市場価格 {market_price:.1f}円 ＋ プレミアム "
                    f"{fmt_num(fip_premium, 1)}円/kWh） − バランシングコスト "
                    f"{fmt_yen(fip_balancing)} ＝ {fmt_yen(fip_net)}/年",
                    font_size_pt=SIZE_SMALL, font_color=C_SUB)

    # ---- Section 3: self-consumption + FIP = total benefit ----
    add_section_header(slide, MARGIN, ys[2], content_w,
                       "年間メリットの合計（自家消費＋FIP売電）")
    row_y = int(ys[2]) + int(Inches(0.36))
    row_h = Inches(1.05)
    add_kpi_card(slide, grid_x(0), row_y, grid_w(4), row_h,
                 fmt_yen(self_saving, "") if self_saving else "—", "円/年",
                 "自家消費メリット（電気代削減）")
    add_kpi_card(slide, grid_x(4), row_y, grid_w(4), row_h,
                 fmt_yen(fip_net, "") if fip_net else "—", "円/年",
                 "FIP売電収入（余剰分）")
    # Total band: C_PANEL + orange left bar + 28pt number (pp9 style)
    band_x = grid_x(8)
    band_w = grid_w(4)
    add_rect(slide, band_x, row_y, band_w, row_h, C_PANEL)
    add_rect(slide, band_x, row_y, Inches(0.05), row_h, C_ORANGE)
    add_textbox(slide, int(band_x) + int(Inches(0.18)),
                row_y + int(Inches(0.12)),
                int(band_w) - int(Inches(0.36)), Inches(0.20),
                "合計年間メリット",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, bold=True)
    add_number_unit(slide, int(band_x) + int(Inches(0.18)),
                    row_y + int(Inches(0.30)),
                    int(band_w) - int(Inches(0.36)),
                    int(row_h) - int(Inches(0.44)),
                    fmt_yen(total_benefit, "") if total_benefit else "—",
                    "円/年")

    # ---- Section 4: Notes ----
    add_section_header(slide, MARGIN, ys[3], content_w, "留意事項",
                       font_size_pt=11)
    note_lines = [
        ("・FIP認定の取得が必要です（経済産業省への申請手続き）",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        ("・バランシングコスト（発電予測と実績の差分ペナルティ）が発生します"
         f"（本試算では {balancing_rate:.1f} 円/kWh で概算）",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        ("・市場価格は変動するため、実際の収入は上記試算と異なる場合があります"
         f"（本試算の想定市場価格：{market_price:.1f} 円/kWh）",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
    ]
    add_multiline_textbox(slide, MARGIN, int(ys[3]) + int(Inches(0.32)),
                          content_w, Inches(0.60),
                          note_lines, line_spacing=1.5)

    add_footer(slide)
