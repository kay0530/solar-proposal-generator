"""
pp5.py - 導入システム概要 (System overview / specifications)

Layout: A4 landscape
- Header bar
- KPI cards: システム容量, パネル合計, PCS合計
- Equipment table: パネル, PCS, 蓄電池
- Site info: 設置場所, 積雪量
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, add_table, add_kpi_card, fmt_num,
)

TITLE = "導入システム概要"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP5 (system overview) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.1)

    # ---- KPI cards (3 across) ----
    capacity = data.get("system_capacity_kw")
    panel_kw = data.get("panel_total_kw")
    pcs_kw = data.get("pcs_total_kw")
    battery_kwh = data.get("battery_total_kwh")

    kpis = [
        (fmt_num(capacity, 1) if capacity else "—", "kW", "システム容量"),
        (fmt_num(panel_kw, 1) if panel_kw else "—", "kW", "パネル合計出力"),
        (fmt_num(pcs_kw, 1) if pcs_kw else "—", "kW", "PCS合計出力"),
    ]

    card_cols = 3
    gap = Inches(0.15)
    card_w = (SLIDE_W - MARGIN * 2 - gap * (card_cols - 1)) / card_cols
    card_h = Inches(1.15)

    for i, (number, unit, label) in enumerate(kpis):
        cx = MARGIN + i * (card_w + gap)
        add_kpi_card(slide, cx, y, card_w, card_h,
                     number, unit, label,
                     bg_color=C_LIGHT_ORANGE, number_size_pt=30)

    y += card_h + Inches(0.2)

    # ---- Equipment table ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "設備仕様", font_size_pt=12)
    y += Inches(0.35)

    # Build rows
    rows = [["区分", "型式", "出力", "数量", "合計"]]

    panels = data.get("panels") or []
    for p in panels:
        rows.append([
            "太陽光パネル",
            p.get("model", "—"),
            f"{p.get('watt_per_unit', 0)}W",
            str(p.get("count", 0)),
            f"{p.get('total_kw', 0):.1f} kW",
        ])

    pcs_list = data.get("pcs_list") or []
    for pcs in pcs_list:
        rows.append([
            "PCS",
            pcs.get("model", "—"),
            f"{pcs.get('kw_per_unit', 0)} kW",
            str(pcs.get("count", 0)),
            f"{pcs.get('total_kw', 0):.1f} kW",
        ])

    batteries = data.get("batteries") or []
    for bat in batteries:
        rows.append([
            "蓄電池",
            bat.get("model", "—"),
            f"{bat.get('kwh_per_unit', 0)} kWh",
            str(bat.get("count", 0)),
            f"{bat.get('total_kwh', 0):.1f} kWh",
        ])

    # Fallback if no equipment data
    if len(rows) == 1:
        rows.append(["太陽光パネル", "—", "—", "—", f"{panel_kw or '—'} kW"])
        rows.append(["PCS", "—", "—", "—", f"{pcs_kw or '—'} kW"])
        if battery_kwh:
            rows.append(["蓄電池", "—", "—", "—", f"{battery_kwh} kWh"])

    table_w = SLIDE_W - MARGIN * 2
    col_widths = [
        Inches(1.5), Inches(3.5), Inches(1.6), Inches(1.0), Inches(1.6),
    ]
    # Adjust last col to fill remaining width
    col_widths[-1] = table_w - sum(col_widths[:-1])

    add_table(slide, MARGIN, y, table_w, rows, col_widths, font_size_pt=9)
    y += Inches(0.28) * len(rows) + Inches(0.2)

    # ---- Site information ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "設置情報", font_size_pt=12)
    y += Inches(0.35)

    address = data.get("address", "") or "—"
    snow_load = data.get("snow_load", "") or ""

    add_textbox(slide, MARGIN + Inches(0.1), y,
                Inches(1.2), Inches(0.28),
                "設置場所：",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_SUB, bold=True)
    add_textbox(slide, MARGIN + Inches(1.3), y,
                SLIDE_W - MARGIN * 2 - Inches(1.3), Inches(0.28),
                address,
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK)
    y += Inches(0.3)

    if snow_load:
        add_textbox(slide, MARGIN + Inches(0.1), y,
                    Inches(1.2), Inches(0.28),
                    "積雪量：",
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB, bold=True)
        add_textbox(slide, MARGIN + Inches(1.3), y,
                    SLIDE_W - MARGIN * 2 - Inches(1.3), Inches(0.28),
                    snow_load,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK)

    add_footer(slide)
