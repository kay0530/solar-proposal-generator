"""
ep3.py - システム概要（EPC）

Design v2 "Institutional Trust Grid":
  Top: KPI card row (容量 / 販売価格 / 実質投資額 or 投資額 / kW単価)
  Then: audited spec tables (パネル / PCS / 蓄電池) with section headers,
  laid out via vstack with exact TABLE_ROW_H advancing. 8pt note at bottom.
"""
from __future__ import annotations

from pathlib import Path

from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_SUB,
    FONT_BODY, GAP_CARD, MARGIN, SLIDE_W,
    SIZE_BODY, SIZE_SMALL, SIZE_TABLE, TABLE_ROW_H,
    add_footer, add_header_bar, add_kpi_card, add_section_header,
    add_table, add_textbox,
    fmt_num, vstack,
)

TITLE = "システム概要（EPC）"
EYEBROW = "02｜システム概要"


def _yen_parts(v) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for add_kpi_card."""
    if v is None:
        return "—", ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v), ""
    if abs(f) >= 1_0000_0000:
        return f"{f / 1_0000_0000:.2f}", "億円"
    if abs(f) >= 10_000:
        return f"{f / 10_000:,.0f}", "万円"
    return f"{f:,.0f}", "円"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP3 (EPC system overview) onto an already-added blank slide.

    data keys used:
        system_capacity_kw, panel_total_kw, selling_price, subsidy_amount,
        kw_unit_cost, panels (list of {model, watt_per_unit, count}),
        pcs_list (list of PCS specs),
        batteries (list of {model, kwh_per_unit, count, total_kwh}),
        battery_total_kwh
    """
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    capacity       = data.get("system_capacity_kw")
    panel_kw       = data.get("panel_total_kw")
    selling_price  = data.get("selling_price")
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    kw_unit_cost   = data.get("kw_unit_cost", 0) or 0
    panels         = data.get("panels", []) or []
    pcs_list       = data.get("pcs_list", []) or []
    batteries      = data.get("batteries") or []
    battery_kwh    = data.get("battery_total_kwh", 0) or 0

    # Net investment = selling_price - subsidy (if subsidy exists)
    net_price = None
    if selling_price is not None:
        net_price = (selling_price - subsidy_amount
                     if subsidy_amount else selling_price)

    content_w = SLIDE_W - MARGIN * 2

    # ---- KPI card row ----
    has_kw_cost = bool(kw_unit_cost) and kw_unit_cost > 0
    sp_num, sp_unit = _yen_parts(selling_price)
    kpis = [
        (fmt_num(capacity, 1), "kW", "システム容量"),
        (sp_num, sp_unit, "販売価格（税別）"),
    ]
    if subsidy_amount > 0:
        np_num, np_unit = _yen_parts(net_price)
        kpis.append((np_num, np_unit, "実質投資額（補助金控除後）"))
    else:
        kpis.append((sp_num, sp_unit, "投資額"))
    if has_kw_cost:
        kpis.append((fmt_num(kw_unit_cost, 0), "円/kW", "kW単価"))

    kpi_h = Inches(1.00)
    card_w = (content_w - GAP_CARD * (len(kpis) - 1)) / len(kpis)

    def draw_kpis(y):
        for i, (number, unit, label) in enumerate(kpis):
            cx = MARGIN + i * (card_w + GAP_CARD)
            add_kpi_card(slide, cx, y, card_w, kpi_h, number, unit, label)

    # ---- Spec section block builder (header + audited table / fallback) ----
    head_h = Inches(0.34)

    def _section_block(sect_title, rows, col_widths, fallback):
        """Return (height, draw_fn). Table height advances by exact
        TABLE_ROW_H * len(rows)."""
        if rows:
            h = head_h + TABLE_ROW_H * len(rows)

            def draw(y, _t=sect_title, _rows=rows, _cw=col_widths):
                add_section_header(slide, MARGIN, y, content_w, _t)
                add_table(slide, MARGIN, y + head_h, content_w, _rows, _cw,
                          font_size_pt=SIZE_TABLE)
        else:
            h = head_h + Inches(0.22)

            def draw(y, _t=sect_title, _f=fallback):
                add_section_header(slide, MARGIN, y, content_w, _t)
                add_textbox(slide, MARGIN + Inches(0.20), y + head_h,
                            content_w - Inches(0.20), Inches(0.22),
                            _f,
                            font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                            font_color=C_SUB)
        return h, draw

    col_widths_4 = [content_w - Inches(2.0) * 3] + [Inches(2.0)] * 3
    col_widths_3 = [content_w - Inches(2.4) * 2] + [Inches(2.4)] * 2

    blocks: list[tuple] = [(kpi_h, draw_kpis)]

    # Panel specs
    panel_rows = []
    if panels:
        panel_rows.append(["メーカー / 型式", "出力 (W)", "枚数", "合計 (kW)"])
        for p in panels:
            model = p.get("model", "—")
            watt  = p.get("watt_per_unit", 0) or 0
            count = p.get("count", 0) or 0
            try:
                total = f"{float(watt) * float(count) / 1000:.1f}"
            except (TypeError, ValueError):
                total = "—"
            panel_rows.append([model, str(watt), str(count), total])
    blocks.append(_section_block(
        "太陽光パネル仕様", panel_rows, col_widths_4,
        f"パネル合計出力：{fmt_num(panel_kw, 1)} kW"))

    # PCS specs
    pcs_rows = []
    if pcs_list:
        pcs_rows.append(["メーカー / 型式", "容量 (kW)", "台数"])
        for pcs in pcs_list:
            model = pcs.get("model", "—")
            cap   = pcs.get("capacity_kw", "—")
            count = pcs.get("count", 1)
            pcs_rows.append([model, str(cap), str(count)])
    blocks.append(_section_block(
        "パワーコンディショナ（PCS）仕様", pcs_rows, col_widths_3,
        "PCS仕様：詳細は別途ご案内"))

    # Battery specs (only when data exists)
    if batteries or battery_kwh:
        bat_rows = []
        if batteries:
            bat_rows.append(["メーカー / 型式", "容量 (kWh)", "台数", "合計 (kWh)"])
            for bat in batteries:
                model = bat.get("model", "—")
                kwh   = bat.get("kwh_per_unit", 0) or 0
                count = bat.get("count", 1) or 1
                try:
                    total_v = bat.get("total_kwh")
                    if total_v is None:
                        total_v = float(kwh) * float(count)
                    total = f"{float(total_v):.1f}"
                except (TypeError, ValueError):
                    total = "—"
                bat_rows.append([model, str(kwh), str(count), total])
        blocks.append(_section_block(
            "蓄電池仕様", bat_rows, col_widths_4,
            f"蓄電池合計容量：{fmt_num(battery_kwh, 1)} kWh"))

    # Note (8pt)
    note_h = Inches(0.22)

    def draw_note(y):
        add_textbox(slide, MARGIN, y, content_w, note_h,
                    "※ 上記価格には設計・施工・申請費用を含みます。表示価格は税別です。",
                    font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                    font_color=C_SUB)

    blocks.append((note_h, draw_note))

    # ---- Vertical justify ----
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [h for h, _ in blocks])
    for (h, fn), y in zip(blocks, ys):
        fn(y)

    add_footer(slide)
