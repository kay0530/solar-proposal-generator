"""
pp_estimate.py - PPA概算費用参考資料 (design v2: Institutional Trust Grid)

For PPA proposals the customer pays nothing upfront, so this slide shows:
  - Estimated facility value (設備概算価額) for reference — 3 KPI cards
  - PPA unit price + monthly fee estimate as an inline metric band
  - Contract-period cost comparison as an audited-figures table
    (savings column tinted, cumulative row as total row)

This is a reference document, not a binding quote.
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_HAIR, C_NAVY, C_ORANGE,
    C_PANEL, C_SUB,
    FONT_BLACK, FONT_BODY, MARGIN, SIZE_BODY, SIZE_CAPTION, SIZE_H2,
    SIZE_SMALL, SLIDE_W, TABLE_ROW_H,
    add_footer, add_header_bar, add_kpi_card, add_line,
    add_multiline_textbox, add_number_unit, add_rect,
    add_section_header, add_table, add_textbox,
    fmt_num, fmt_yen, grid_w, grid_x, vstack,
)

TITLE = "PPA概算費用参考資料"
EYEBROW = "04｜ご契約条件"

DEGRADATION = 0.005  # 0.5% annual degradation


def _fmt_date(val) -> str:
    """Format date string."""
    if not val:
        return (date.today().strftime("%Y年%m月%d日")
                .replace("年0", "年").replace("月0", "月"))
    s = str(val).split(" ")[0]
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    return s


def _yen_parts(v: float) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for KPI display."""
    if v >= 1_0000_0000:
        return f"{v / 1_0000_0000:.2f}", "億円"
    if v >= 10_000:
        return f"{v / 10_000:,.0f}", "万円"
    return f"{v:,.0f}", "円"


def _tbl_yen(v: float) -> str:
    """Table cell: comma-grouped yen or em-dash."""
    return f"{int(v):,}円" if v > 0 else "—"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP_ESTIMATE (PPA cost reference) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    company = data.get("company_name", "") or ""
    office = data.get("office_name", "") or ""
    proposal_date = _fmt_date(data.get("proposal_date"))
    tax_display = data.get("tax_display", "税抜")

    # PPA data
    ppa_price = float(data.get("ppa_unit_price", 0) or 0)
    years = int(data.get("contract_years", 20) or 20)
    self_kwh = float(data.get("self_consumption_kwh", 0) or 0)
    selling_price = float(data.get("selling_price", 0) or 0)
    annual_saving = float(data.get("annual_saving", 0) or 0)
    annual_cost = float(data.get("annual_cost", 0) or 0)
    capacity = float(data.get("system_capacity_kw", 0) or 0)
    annual_kwh = float(data.get("annual_kwh", 0) or 0)

    # ---- Derived figures ----
    kw_price = (int(selling_price / capacity)
                if capacity > 0 and selling_price > 0 else 0)
    monthly_ppa = (int(self_kwh * ppa_price / 12)
                   if self_kwh > 0 and ppa_price > 0 else 0)

    # Contract-period totals with degradation
    total_ppa_cost = 0.0
    total_grid_cost = 0.0
    if self_kwh > 0 and ppa_price > 0 and annual_cost > 0:
        avg_rate = annual_cost / annual_kwh if annual_kwh > 0 else 0
        for yr in range(years):
            decay = (1 - DEGRADATION) ** yr
            yr_self_kwh = self_kwh * decay
            total_ppa_cost += yr_self_kwh * ppa_price
            total_grid_cost += yr_self_kwh * avg_rate
    total_saving = total_grid_cost - total_ppa_cost

    # ---- Block heights for vertical justify ----
    cust_h = Inches(0.32)
    notice_h = Inches(0.46)
    kpi_h = Inches(0.95)
    block_kpi_h = int(Inches(0.36)) + int(kpi_h)
    fee_h = Inches(0.86)
    block_fee_h = int(Inches(0.36)) + int(fee_h)
    table_h = int(TABLE_ROW_H) * 4
    block_tbl_h = int(Inches(0.36)) + table_h
    notes_h = Inches(0.56)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [cust_h, notice_h, block_kpi_h, block_fee_h,
                 block_tbl_h, notes_h])

    # ---- Customer + date row ----
    customer_label = company
    if office:
        customer_label += f"  {office}"
    customer_label += "  御中"
    add_textbox(slide, MARGIN, ys[0], Inches(6.5), cust_h,
                customer_label,
                font_name=FONT_BLACK, font_size_pt=SIZE_H2,
                font_color=C_DARK, bold=True)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(3.0), ys[0],
                Inches(3.0), Inches(0.22),
                f"参考資料  {proposal_date}",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, align=PP_ALIGN.RIGHT)

    # ---- Notice band (panel + orange bar) ----
    add_rect(slide, MARGIN, ys[1], content_w, notice_h, C_PANEL)
    add_rect(slide, MARGIN, ys[1], Inches(0.05), notice_h, C_ORANGE)
    add_textbox(slide, int(MARGIN) + int(Inches(0.25)), ys[1],
                int(content_w) - int(Inches(0.50)), notice_h,
                "PPAモデルでは設備費用のお客様負担はございません。"
                "本資料は設備の概算価額と電力料金の参考情報です。",
                font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                font_color=C_NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    # ---- Section 1: facility value KPI cards ----
    add_section_header(slide, MARGIN, ys[2], content_w,
                       "設備概算価額（ご参考）")
    kpi_y = int(ys[2]) + int(Inches(0.36))
    price_num, price_unit = (_yen_parts(selling_price)
                             if selling_price > 0 else ("—", ""))
    kpis = [
        (fmt_num(capacity, 1) if capacity > 0 else "—", "kW", "システム容量"),
        (price_num, price_unit, "設備概算価額"),
        (f"{kw_price:,}" if kw_price else "—", "円/kW", "kW単価（参考）"),
    ]
    for i, (number, unit, label) in enumerate(kpis):
        add_kpi_card(slide, grid_x(i * 4), kpi_y, grid_w(4), kpi_h,
                     number, unit, label)

    # ---- Section 2: PPA fee metric band ----
    add_section_header(slide, MARGIN, ys[3], content_w,
                       "PPA電力料金お見積り")
    fee_y = int(ys[3]) + int(Inches(0.36))
    fee_items = [
        (f"PPA電力単価（{tax_display}）",
         f"{ppa_price:.2f}" if ppa_price > 0 else "—", "円/kWh",
         f"{years}年間一律単価（契約期間中変動なし）"),
        ("月額電力料金（概算）",
         f"{monthly_ppa:,}" if monthly_ppa > 0 else "—", "円/月",
         "税別・初年度概算"),
    ]
    for i, (label, number, unit, caption) in enumerate(fee_items):
        bx = grid_x(i * 6)
        bw = grid_w(6) - Inches(0.20)
        add_textbox(slide, bx, fee_y, bw, Inches(0.18),
                    label,
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB, bold=True)
        add_number_unit(slide, bx, fee_y + int(Inches(0.20)),
                        bw, Inches(0.42),
                        number, unit)
        add_textbox(slide, bx, fee_y + int(Inches(0.66)),
                    bw, Inches(0.16),
                    caption,
                    font_size_pt=SIZE_SMALL, font_color=C_SUB)
        if i > 0:
            sep_x = bx - Inches(0.12)
            add_line(slide, sep_x, fee_y + int(Inches(0.04)),
                     sep_x, fee_y + int(fee_h) - int(Inches(0.10)),
                     C_HAIR, width_pt=0.5)

    # ---- Section 3: contract-period cost comparison table ----
    add_section_header(slide, MARGIN, ys[4], content_w,
                       f"{years}年間コスト比較")
    tbl_y = int(ys[4]) + int(Inches(0.36))

    yr1_ppa = self_kwh * ppa_price if (self_kwh > 0 and ppa_price > 0) else 0
    rows = [
        ["", "現在の電力料金", "PPA導入後", "削減効果"],
        ["年間電力料金（初年度）",
         _tbl_yen(annual_cost), _tbl_yen(yr1_ppa), _tbl_yen(annual_saving)],
        ["初期費用", "—", "0円（無料）", "—"],
        [f"{years}年間累計",
         _tbl_yen(total_grid_cost), _tbl_yen(total_ppa_cost),
         _tbl_yen(total_saving)],
    ]
    col_widths = [Inches(2.9), Inches(2.6), Inches(2.6),
                  int(content_w) - int(Inches(8.1))]
    add_table(slide, MARGIN, tbl_y, content_w, rows, col_widths,
              highlight_col=3, total_row=3)

    # ---- Notes (8pt floor) ----
    note_lines = [
        ("※ 本資料はPPA電力供給契約の参考資料であり、正式な見積書ではありません。",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        ("※ 設備概算価額はお客様の負担額ではなく、PPA事業者が負担する設備費用の参考値です。",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        (f"※ 電力料金は{tax_display}表記。発電量は年率0.5%低減で試算。"
         "電力料金単価はPPA契約期間を通じて変動しません。",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
    ]
    add_multiline_textbox(slide, MARGIN, ys[5], content_w, notes_h,
                          note_lines, line_spacing=1.35)

    add_footer(slide)
