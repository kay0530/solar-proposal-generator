"""
ep_estimate.py - EPC estimate / quotation slide

Professional Japanese estimate form for EPC (purchase) proposals.
Shows a clean quotation table with line items, subtotal, tax, and total.

Only included when estimate data is provided (estimate_items in customer_data).
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
    C_WHITE, C_NAVY, C_BORDER,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_line, _set_cell_bg,
)

TITLE = "概算費用お見積書"

# Slightly lighter navy for table header
C_TABLE_HEADER = RGBColor(0x00, 0x30, 0x70)
C_TOTAL_BG = RGBColor(0xFF, 0xF0, 0xE0)  # warm highlight for total row


def _fmt_date(val) -> str:
    """Format date for estimate header."""
    if not val:
        return date.today().strftime("%Y年%m月%d日").replace("年0", "年").replace("月0", "月")
    s = str(val).split(" ")[0]
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    return s


def _fmt_yen_comma(val) -> str:
    """Format number with comma separators and yen symbol."""
    if val is None or val == 0:
        return "-"
    try:
        return f"\\{int(val):,}"
    except (TypeError, ValueError):
        return str(val)


def _fmt_comma(val) -> str:
    """Format number with commas, no yen symbol."""
    if val is None or val == 0:
        return "-"
    try:
        return f"{int(val):,}"
    except (TypeError, ValueError):
        return str(val)


def _build_estimate_items(data: dict) -> list[dict]:
    """Build estimate line items from customer_data.

    If data contains 'estimate_items' (list of dicts), use that directly.
    Otherwise, auto-build from equipment data.

    Each item: {name, spec, qty, unit, unit_price, amount}
    """
    # Use explicit estimate items if provided
    items = data.get("estimate_items")
    if items:
        return items

    # Auto-build from equipment data
    auto_items = []
    panels = data.get("panels", [])
    for p in panels:
        model = p.get("model", "太陽電池モジュール")
        count = p.get("count", 0) or 0
        watt = p.get("watt_per_unit", 0) or 0
        unit_price = p.get("selling_unit_price", 0) or 0
        if count > 0:
            auto_items.append({
                "name": "太陽電池モジュール",
                "spec": f"{model} ({watt}W)" if watt else model,
                "qty": count,
                "unit": "枚",
                "unit_price": unit_price,
                "amount": int(unit_price * count) if unit_price else None,
            })

    pcs_list = data.get("pcs_list", [])
    for pcs in pcs_list:
        model = pcs.get("model", "パワーコンディショナ")
        count = pcs.get("count", 0) or 0
        kw = pcs.get("kw_per_unit", 0) or 0
        unit_price = pcs.get("selling_unit_price", 0) or 0
        if count > 0:
            auto_items.append({
                "name": "パワーコンディショナ",
                "spec": f"{model} ({kw}kW)" if kw else model,
                "qty": count,
                "unit": "台",
                "unit_price": unit_price,
                "amount": int(unit_price * count) if unit_price else None,
            })

    # Frame & installation
    frame_cost = data.get("estimate_frame_cost", 0) or 0
    if frame_cost > 0:
        auto_items.append({
            "name": "架台・施工費",
            "spec": "",
            "qty": 1,
            "unit": "式",
            "unit_price": frame_cost,
            "amount": frame_cost,
        })

    # Electrical work
    elec_cost = data.get("estimate_electrical_cost", 0) or 0
    if elec_cost > 0:
        auto_items.append({
            "name": "電気工事費",
            "spec": "",
            "qty": 1,
            "unit": "式",
            "unit_price": elec_cost,
            "amount": elec_cost,
        })

    # Battery
    batteries = data.get("batteries", [])
    for bat in batteries:
        model = bat.get("model", "蓄電池")
        count = bat.get("count", 0) or 0
        kwh = bat.get("kwh_per_unit", 0) or 0
        unit_price = bat.get("selling_unit_price", 0) or 0
        if count > 0:
            auto_items.append({
                "name": "蓄電池システム",
                "spec": f"{model} ({kwh}kWh)" if kwh else model,
                "qty": count,
                "unit": "台",
                "unit_price": unit_price,
                "amount": int(unit_price * count) if unit_price else None,
            })

    # Additional items
    extra_items = data.get("estimate_extra_items", [])
    for ex in extra_items:
        auto_items.append(ex)

    return auto_items


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render EP_ESTIMATE (EPC quotation page) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    company = data.get("company_name", "") or ""
    office = data.get("office_name", "") or ""
    proposal_date = _fmt_date(data.get("proposal_date"))
    tax_display = data.get("tax_display", "税抜")

    # Generate estimate number from date + opp_id
    opp_id = data.get("opp_id", "")
    est_number = data.get("estimate_number", "")
    if not est_number:
        date_part = str(data.get("proposal_date", "")).replace("-", "")[:8]
        est_number = f"EST-{date_part}-001"

    y = CONTENT_TOP

    # ---- Header info: left = customer, right = date/number ----
    # Left: customer name
    customer_label = f"{company}"
    if office:
        customer_label += f"  {office}"
    customer_label += "  御中"

    add_textbox(slide, MARGIN, y, Inches(5.5), Inches(0.35),
                customer_label,
                font_name=FONT_BLACK, font_size_pt=16,
                font_color=C_DARK, bold=True)

    # Underline under customer name
    add_line(slide, MARGIN, y + Inches(0.38),
             MARGIN + Inches(5.5), y + Inches(0.38),
             C_DARK, width_pt=1.5)

    # Right: date and estimate number
    right_x = SLIDE_W - MARGIN - Inches(3.5)
    add_textbox(slide, right_x, y, Inches(3.5), Inches(0.22),
                f"見積日: {proposal_date}",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK, align=PP_ALIGN.RIGHT)
    add_textbox(slide, right_x, y + Inches(0.22), Inches(3.5), Inches(0.22),
                f"見積番号: {est_number}",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK, align=PP_ALIGN.RIGHT)

    y += Inches(0.55)

    # ---- Subject line ----
    capacity = data.get("system_capacity_kw", 0) or 0
    subject = f"太陽光発電設備工事  {capacity:.1f}kW" if capacity else "太陽光発電設備工事"
    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.26),
                f"件名: {subject}",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_DARK, bold=True)
    y += Inches(0.32)

    # ---- Estimate table ----
    items = _build_estimate_items(data)

    # Calculate subtotal from items that have amounts
    subtotal = 0
    for item in items:
        amt = item.get("amount")
        if amt and amt > 0:
            subtotal += amt

    # If no item amounts but selling_price exists, use that
    selling_price = data.get("selling_price", 0) or 0
    if subtotal == 0 and selling_price > 0:
        subtotal = selling_price

    tax_rate = 0.10
    tax_amount = int(subtotal * tax_rate)
    grand_total = subtotal + tax_amount

    # Build table rows
    table_w = SLIDE_W - MARGIN * 2
    n_cols = 6  # No., Item, Spec, Qty, Unit Price, Amount
    col_widths = [
        Inches(0.45),   # No.
        Inches(2.2),    # Item name
        Inches(3.8),    # Spec/details
        Inches(0.9),    # Qty
        Inches(1.8),    # Unit price
        Inches(1.85),   # Amount
    ]

    # Header row
    rows_data = [["No.", "項目", "仕様・詳細", "数量", "単価（円）", f"金額（円）"]]

    for i, item in enumerate(items, 1):
        qty_str = f"{item['qty']}{item.get('unit', '')}" if item.get('qty') else ""
        rows_data.append([
            str(i),
            item.get("name", ""),
            item.get("spec", ""),
            qty_str,
            _fmt_comma(item.get("unit_price")),
            _fmt_comma(item.get("amount")),
        ])

    # Add empty rows if too few items (minimum 6 for professional look)
    while len(rows_data) < 7:
        rows_data.append(["", "", "", "", "", ""])

    # Summary rows
    rows_data.append(["", "", "", "", "小計", _fmt_comma(subtotal)])
    rows_data.append(["", "", "", "", "消費税（10%）", _fmt_comma(tax_amount)])
    rows_data.append(["", "", "", "", "合計（税込）", _fmt_comma(grand_total)])

    # Render table manually for better control
    row_h = Inches(0.28)
    n_rows = len(rows_data)
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, MARGIN, y, table_w, row_h * n_rows)
    tbl = tbl_shape.table

    # Set column widths
    for c, cw in enumerate(col_widths):
        tbl.columns[c].width = int(cw)

    # Right-align columns for numbers (qty, unit_price, amount)
    right_align_cols = {3, 4, 5}

    for r, row in enumerate(rows_data):
        is_header = (r == 0)
        is_summary = (r >= n_rows - 3)
        is_total = (r == n_rows - 1)

        for c, cell_text in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(cell_text) if cell_text is not None else ""

            # Margins
            cell.margin_left = Pt(4)
            cell.margin_right = Pt(4)
            cell.margin_top = Pt(2)
            cell.margin_bottom = Pt(2)

            for para in cell.text_frame.paragraphs:
                if is_header:
                    para.alignment = PP_ALIGN.CENTER
                elif c in right_align_cols:
                    para.alignment = PP_ALIGN.RIGHT
                elif c == 0:
                    para.alignment = PP_ALIGN.CENTER
                else:
                    para.alignment = PP_ALIGN.LEFT

                for run in para.runs:
                    run.font.name = FONT_BODY
                    run.font.size = Pt(9) if not is_header else Pt(9)
                    run.font.bold = is_header or is_total
                    if is_header:
                        run.font.color.rgb = C_WHITE
                    elif is_total:
                        run.font.color.rgb = C_NAVY
                        run.font.size = Pt(11)
                    else:
                        run.font.color.rgb = C_DARK

            # Cell background
            if is_header:
                _set_cell_bg(cell, C_TABLE_HEADER)
            elif is_total:
                _set_cell_bg(cell, C_TOTAL_BG)
            elif is_summary:
                _set_cell_bg(cell, C_LIGHT_GRAY)
            elif r % 2 == 0:
                _set_cell_bg(cell, C_WHITE)
            else:
                _set_cell_bg(cell, RGBColor(0xFA, 0xFA, 0xFA))

    y += row_h * n_rows + Inches(0.18)

    # ---- Grand total highlight box ----
    total_box_w = Inches(4.5)
    total_box_h = Inches(0.55)
    total_box_x = SLIDE_W - MARGIN - total_box_w
    add_rounded_rect(slide, total_box_x, y, total_box_w, total_box_h, C_NAVY)
    add_textbox(slide, total_box_x + Inches(0.2), y + Inches(0.05),
                Inches(1.8), total_box_h - Inches(0.1),
                f"お見積り合計（税込）",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_WHITE, bold=True)
    add_textbox(slide, total_box_x + Inches(2.0), y + Inches(0.03),
                Inches(2.3), total_box_h - Inches(0.06),
                _fmt_yen_comma(grand_total),
                font_name=FONT_BLACK, font_size_pt=20,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.RIGHT)

    y += total_box_h + Inches(0.20)

    # ---- Notes section ----
    validity = data.get("estimate_validity", "本見積書発行日より1ヶ月間")
    delivery = data.get("estimate_delivery", "ご発注後、別途ご相談")
    subsidy_note = data.get("estimate_subsidy_note", "補助金申請費用は別途お見積り")

    notes = [
        f"見積有効期限: {validity}",
        f"納期目安: {delivery}",
        f"備考: {subsidy_note}",
        f"金額表記: {tax_display}（消費税は税込合計に含む）",
    ]

    add_textbox(slide, MARGIN, y, Inches(1.0), Inches(0.22),
                "備考・条件",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_ORANGE, bold=True)
    y += Inches(0.22)

    for note in notes:
        add_textbox(slide, MARGIN + Inches(0.1), y,
                    SLIDE_W - MARGIN * 2 - Inches(0.1), Inches(0.18),
                    f"  {note}",
                    font_name=FONT_BODY, font_size_pt=8, font_color=C_SUB)
        y += Inches(0.17)

    # ---- Issuer info (right-aligned) ----
    issuer_y = SLIDE_H - Inches(0.85)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(3.5), issuer_y,
                Inches(3.5), Inches(0.20),
                "株式会社オルテナジー",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_DARK, bold=True, align=PP_ALIGN.RIGHT)
    add_textbox(slide, SLIDE_W - MARGIN - Inches(3.5), issuer_y + Inches(0.18),
                Inches(3.5), Inches(0.18),
                "https://altenergy.co.jp/",
                font_name=FONT_BODY, font_size_pt=8,
                font_color=C_SUB, align=PP_ALIGN.RIGHT)

    add_footer(slide)
