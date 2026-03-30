"""
pp_estimate.py - PPA estimate reference slide

For PPA proposals, the customer doesn't pay upfront, so this slide shows:
- Estimated facility value (設備概算価額) for reference
- Monthly PPA fee estimate based on usage
- 20-year total cost comparison (PPA vs grid)

This is a reference document, not a binding quote.
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB,
    C_WHITE, C_NAVY, C_TEAL, C_LIGHT_TEAL, C_BORDER,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_line, add_section_header, add_kpi_card, _set_cell_bg,
    fmt_yen, fmt_num,
)

TITLE = "PPA概算費用参考資料"

DEGRADATION = 0.005  # 0.5% annual degradation


def _fmt_date(val) -> str:
    """Format date string."""
    if not val:
        return date.today().strftime("%Y年%m月%d日").replace("年0", "年").replace("月0", "月")
    s = str(val).split(" ")[0]
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    return s


def _fmt_comma(val) -> str:
    """Format number with commas."""
    if val is None or val == 0:
        return "-"
    try:
        return f"{int(val):,}"
    except (TypeError, ValueError):
        return str(val)


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP_ESTIMATE (PPA cost reference) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

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

    y = CONTENT_TOP

    # ---- Header: customer + date ----
    customer_label = f"{company}"
    if office:
        customer_label += f"  {office}"
    customer_label += "  御中"

    add_textbox(slide, MARGIN, y, Inches(5.5), Inches(0.30),
                customer_label,
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_DARK, bold=True)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(3.0), y, Inches(3.0), Inches(0.22),
                f"参考資料  {proposal_date}",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_SUB, align=PP_ALIGN.RIGHT)

    y += Inches(0.40)

    # ---- Notice box ----
    notice_h = Inches(0.40)
    add_rounded_rect(slide, MARGIN, y, SLIDE_W - MARGIN * 2, notice_h,
                     C_LIGHT_ORANGE)
    add_textbox(slide, MARGIN + Inches(0.15), y + Inches(0.05),
                SLIDE_W - MARGIN * 2 - Inches(0.3), notice_h - Inches(0.1),
                "PPAモデルでは設備費用のお客様負担はございません。"
                "本資料は設備の概算価額と電力料金の参考情報です。",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_ORANGE, bold=True)

    y += notice_h + Inches(0.20)

    # ---- Section 1: Facility overview ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "設備概算価額（ご参考）", font_size_pt=12)
    y += Inches(0.32)

    # 3 KPI cards: system capacity, facility value, kW unit price
    card_gap = Inches(0.15)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * 2) / 3
    card_h = Inches(0.90)

    kw_price = int(selling_price / capacity) if capacity > 0 and selling_price > 0 else 0

    kpis = [
        (fmt_num(capacity, 1), "kW", "システム容量"),
        (fmt_yen(selling_price), "", "設備概算価額"),
        (f"{kw_price:,}" if kw_price else "-", "円/kW", "kW単価（参考）"),
    ]
    for i, (number, unit, label) in enumerate(kpis):
        cx = MARGIN + i * (card_w + card_gap)
        add_kpi_card(slide, cx, y, card_w, card_h,
                     number, unit, label,
                     bg_color=C_LIGHT_GRAY, number_size_pt=22)

    y += card_h + Inches(0.20)

    # ---- Section 2: PPA fee estimate ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "PPA電力料金お見積り", font_size_pt=12)
    y += Inches(0.32)

    # Two large boxes: PPA unit price + Monthly estimate
    box_gap = Inches(0.25)
    box_w = (SLIDE_W - MARGIN * 2 - box_gap) / 2
    box_h = Inches(1.30)

    # Left: PPA unit price
    add_rounded_rect(slide, MARGIN, y, box_w, box_h, C_NAVY)
    add_textbox(slide, MARGIN, y + Inches(0.10), box_w, Inches(0.22),
                f"PPA電力単価（{tax_display}）",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    price_str = f"\\{ppa_price:.2f}" if ppa_price > 0 else "-"
    add_textbox(slide, MARGIN, y + Inches(0.38), box_w, Inches(0.55),
                price_str,
                font_name=FONT_BLACK, font_size_pt=32,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, MARGIN, y + Inches(0.90), box_w, Inches(0.22),
                f"/kWh  x  {years}年間一律",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=RGBColor(0xCC, 0xCC, 0xCC), align=PP_ALIGN.CENTER)

    # Right: Monthly estimate
    rx = MARGIN + box_w + box_gap
    monthly_ppa = 0
    if self_kwh > 0 and ppa_price > 0:
        monthly_ppa = int(self_kwh * ppa_price / 12)

    add_rounded_rect(slide, rx, y, box_w, box_h, C_LIGHT_TEAL)
    add_textbox(slide, rx, y + Inches(0.10), box_w, Inches(0.22),
                "月額電力料金（概算）",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    monthly_str = f"\\{monthly_ppa:,}" if monthly_ppa > 0 else "-"
    add_textbox(slide, rx, y + Inches(0.38), box_w, Inches(0.55),
                monthly_str,
                font_name=FONT_BLACK, font_size_pt=32,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, rx, y + Inches(0.90), box_w, Inches(0.22),
                "/月（税別・初年度概算）",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=RGBColor(0xCC, 0xCC, 0xCC), align=PP_ALIGN.CENTER)

    y += box_h + Inches(0.20)

    # ---- Section 3: 20-year cost comparison table ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       f"{years}年間コスト比較", font_size_pt=12)
    y += Inches(0.30)

    # Build comparison table
    table_w = SLIDE_W - MARGIN * 2
    n_cols = 4
    col_widths = [
        Inches(2.8),   # Label
        Inches(2.8),   # Current (grid only)
        Inches(2.8),   # PPA
        Inches(2.6),   # Savings
    ]

    # Calculate 20-year totals with degradation
    total_ppa_cost = 0
    total_grid_cost = 0
    if self_kwh > 0 and ppa_price > 0 and annual_cost > 0:
        annual_kwh = float(data.get("annual_kwh", 0) or 0)
        avg_rate = annual_cost / annual_kwh if annual_kwh > 0 else 0

        for yr in range(years):
            decay = (1 - DEGRADATION) ** yr
            yr_self_kwh = self_kwh * decay
            yr_ppa_cost = yr_self_kwh * ppa_price
            yr_grid_cost = yr_self_kwh * avg_rate
            total_ppa_cost += yr_ppa_cost
            total_grid_cost += yr_grid_cost

    total_saving = total_grid_cost - total_ppa_cost

    rows = [
        ["", "現在の電力料金", "PPA導入後", "削減効果"],
        ["年間電力料金（初年度）",
         f"\\{int(annual_cost):,}" if annual_cost > 0 else "-",
         f"\\{int(self_kwh * ppa_price):,}" if (self_kwh > 0 and ppa_price > 0) else "-",
         f"\\{int(annual_saving):,}" if annual_saving > 0 else "-"],
        [f"{years}年間累計",
         f"\\{int(total_grid_cost):,}" if total_grid_cost > 0 else "-",
         f"\\{int(total_ppa_cost):,}" if total_ppa_cost > 0 else "-",
         f"\\{int(total_saving):,}" if total_saving > 0 else "-"],
        ["初期費用", "-", "\\0（無料）", "-"],
    ]

    n_rows = len(rows)
    row_h = Inches(0.32)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, MARGIN, y, table_w, row_h * n_rows)
    tbl = tbl_shape.table

    for c, cw in enumerate(col_widths):
        tbl.columns[c].width = int(cw)

    for r, row_data in enumerate(rows):
        is_header = (r == 0)
        is_total_row = (r == 2)  # 20-year total row

        for c, cell_text in enumerate(row_data):
            cell = tbl.cell(r, c)
            cell.text = str(cell_text)
            cell.margin_left = Pt(6)
            cell.margin_right = Pt(6)
            cell.margin_top = Pt(3)
            cell.margin_bottom = Pt(3)

            for para in cell.text_frame.paragraphs:
                if c == 0:
                    para.alignment = PP_ALIGN.LEFT
                else:
                    para.alignment = PP_ALIGN.RIGHT if not is_header else PP_ALIGN.CENTER

                for run in para.runs:
                    run.font.name = FONT_BODY
                    run.font.size = Pt(10) if not is_total_row else Pt(11)
                    run.font.bold = is_header or is_total_row
                    if is_header:
                        run.font.color.rgb = C_WHITE
                    elif c == 3 and not is_header:
                        # Savings column in orange
                        run.font.color.rgb = C_ORANGE
                        run.font.bold = True
                    else:
                        run.font.color.rgb = C_DARK

            # Background
            if is_header:
                _set_cell_bg(cell, C_NAVY)
            elif is_total_row:
                _set_cell_bg(cell, RGBColor(0xFF, 0xF0, 0xE0))
            elif r % 2 == 0:
                _set_cell_bg(cell, C_WHITE)
            else:
                _set_cell_bg(cell, RGBColor(0xFA, 0xFA, 0xFA))

    y += row_h * n_rows + Inches(0.15)

    # ---- Notes ----
    notes = [
        "本資料はPPA電力供給契約の参考資料であり、正式な見積書ではありません。",
        "設備概算価額はお客様の負担額ではなく、PPA事業者が負担する設備費用の参考値です。",
        f"電力料金は{tax_display}表記。発電量は年率0.5%低減で試算。",
        "電力料金単価はPPA契約期間を通じて変動しません。",
    ]

    for note in notes:
        add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.17),
                    f"※ {note}",
                    font_name=FONT_BODY, font_size_pt=7, font_color=C_SUB)
        y += Inches(0.16)

    add_footer(slide)
