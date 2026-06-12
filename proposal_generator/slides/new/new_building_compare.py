"""
new_building_compare.py - 建物別 経済効果比較スライド (design v2)

  - Audited table: per-building KPIs (容量 / 発電量 / 削減額 / CO2)
  - Proportional navy bars comparing annual savings per building
  - Four KPI cards (28pt) with the combined totals (合計効果)
Uses data["buildings"] (>=2 dicts) if available, otherwise placeholder.
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_SUB,
    GAP_CARD, MARGIN, SIZE_BODY, SIZE_TABLE, SLIDE_W, TABLE_ROW_H,
    add_footer, add_header_bar, add_kpi_card, add_number_unit, add_rect,
    add_section_header, add_table, add_textbox, fmt_num, fmt_yen,
    grid_w, grid_x, vstack,
)

TITLE = "建物別 経済効果比較"
EYEBROW = "03｜効果シミュレーション"

# KPI definitions: (key, table row label incl. unit, format_fn)
KPIS = [
    ("capacity_kw",     "設置容量（kW）",         lambda v: fmt_num(v, 1)),
    ("annual_gen_kwh",  "年間発電量（kWh/年）",   lambda v: fmt_num(v, 0)),
    ("annual_saving",   "年間削減額（円/年）",    lambda v: fmt_yen(v, "")),
    ("co2_reduction_t", "CO₂削減量（t-CO₂/年）",  lambda v: fmt_num(v, 1)),
]

# Default placeholder data
DEFAULT_BUILDINGS = [
    {
        "name": "A棟",
        "capacity_kw": 150,
        "annual_gen_kwh": 165000,
        "annual_saving": 3300000,
        "co2_reduction_t": 78.5,
    },
    {
        "name": "B棟",
        "capacity_kw": 100,
        "annual_gen_kwh": 110000,
        "annual_saving": 2200000,
        "co2_reduction_t": 52.3,
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)
    content_w = SLIDE_W - MARGIN * 2

    company = data.get("company_name", "") or ""

    # Get building data (first two buildings)
    buildings = data.get("buildings")
    if (not buildings or not isinstance(buildings, list)
            or len(buildings) < 2
            or not all(isinstance(b, dict) for b in buildings[:2])):
        buildings = DEFAULT_BUILDINGS
    bldg_a = buildings[0]
    bldg_b = buildings[1]
    name_a = str(bldg_a.get("name") or "A棟")
    name_b = str(bldg_b.get("name") or "B棟")

    # ---- Block heights for vertical justify ----
    n_rows = len(KPIS) + 1
    lead_h = Inches(0.26)
    table_block_h = int(Inches(0.36)) + int(TABLE_ROW_H) * n_rows
    bar_row_h = Inches(0.36)
    bars_block_h = int(Inches(0.36)) + int(bar_row_h) * 2
    kpi_block_h = int(Inches(0.36)) + int(Inches(1.00))

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, table_block_h, bars_block_h, kpi_block_h],
                min_gap=GAP_CARD)

    # ---- Lead ----
    lead = (f"{company} 様｜建物別シミュレーション結果" if company
            else "建物別シミュレーション結果")
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h, lead,
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- Audited per-building table ----
    add_section_header(slide, MARGIN, ys[1], content_w, "建物別の経済効果")
    rows = [["比較項目", name_a, name_b]]
    for key, label, fmt_fn in KPIS:
        va = bldg_a.get(key)
        vb = bldg_b.get(key)
        rows.append([label,
                     fmt_fn(va) if va is not None else "—",
                     fmt_fn(vb) if vb is not None else "—"])
    label_w = int(Inches(2.4))
    col_w = (int(content_w) - label_w) // 2
    add_table(slide, MARGIN, int(ys[1]) + int(Inches(0.36)), content_w,
              rows, [label_w, col_w, col_w], font_size_pt=SIZE_TABLE)

    # ---- Proportional bars: annual saving per building ----
    bars_y = int(ys[2])
    add_section_header(slide, MARGIN, bars_y, content_w, "年間削減額の比較")
    row_y = bars_y + int(Inches(0.36))
    name_w = int(Inches(1.5))
    value_w = int(Inches(1.9))
    bar_x = int(MARGIN) + name_w + int(Inches(0.10))
    bar_avail = (int(SLIDE_W) - int(MARGIN) - bar_x - value_w
                 - int(Inches(0.15)))
    sa = _to_float(bldg_a.get("annual_saving"))
    sb = _to_float(bldg_b.get("annual_saving"))
    vmax = max(sa or 0, sb or 0)
    for i, (nm, val) in enumerate([(name_a, sa), (name_b, sb)]):
        ry = row_y + i * int(bar_row_h)
        add_textbox(slide, MARGIN, ry, name_w, Inches(0.30), nm,
                    font_size_pt=SIZE_BODY, font_color=C_DARK, bold=True,
                    anchor=MSO_ANCHOR.MIDDLE)
        bw = int(bar_avail * (val / vmax)) if val and vmax > 0 else 0
        if bw > 0:
            add_rect(slide, bar_x, ry + int(Inches(0.05)), bw,
                     int(Inches(0.20)), C_NAVY)
        add_number_unit(slide, bar_x + bw + int(Inches(0.12)),
                        ry - int(Inches(0.02)), value_w, int(Inches(0.30)),
                        fmt_yen(val, "") if val else "—", "円/年",
                        number_size_pt=14, unit_size_pt=9)

    # ---- Totals: 合計効果 KPI cards (28pt) ----
    kpi_y = int(ys[3])
    add_section_header(slide, MARGIN, kpi_y, content_w, "合計効果（2棟合計）")
    card_y = kpi_y + int(Inches(0.36))
    card_h = Inches(1.00)
    total_cap = _safe_sum(bldg_a.get("capacity_kw"), bldg_b.get("capacity_kw"))
    total_gen = _safe_sum(bldg_a.get("annual_gen_kwh"),
                          bldg_b.get("annual_gen_kwh"))
    total_saving = _safe_sum(bldg_a.get("annual_saving"),
                             bldg_b.get("annual_saving"))
    total_co2 = _safe_sum(bldg_a.get("co2_reduction_t"),
                          bldg_b.get("co2_reduction_t"))
    totals = [
        (fmt_num(total_cap, 1), "kW", "合計設置容量"),
        (fmt_num(total_gen, 0), "kWh/年", "合計年間発電量"),
        (fmt_yen(total_saving, "") if total_saving else "—",
         "円/年", "合計年間削減額"),
        (fmt_num(total_co2, 1), "t-CO₂/年", "合計CO₂削減量"),
    ]
    for i, (number, unit, label) in enumerate(totals):
        add_kpi_card(slide, grid_x(i * 3), card_y, grid_w(3), card_h,
                     number, unit, label)

    add_footer(slide)


def _to_float(v):
    """Convert to float, returning None when impossible."""
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _safe_sum(a, b):
    """Safely sum two values that may be None."""
    try:
        va = float(a) if a is not None else 0
        vb = float(b) if b is not None else 0
        return va + vb
    except (TypeError, ValueError):
        return None
