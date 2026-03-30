"""
new_building_compare.py - 建物別 経済効果比較スライド

Two-column layout comparing Building A vs Building B.
KPIs per building: 容量, 年間発電量, 年間削減額, CO2削減量.
Uses data.get("buildings") if available, otherwise placeholder.
Summary row: 合計効果.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, C_BORDER,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, fmt_yen, fmt_num,
)

TITLE = "建物別 経済効果比較"

# KPI definitions: (key, label, unit, format_fn)
KPIS = [
    ("capacity_kw",      "設置容量",     "kW",       lambda v: fmt_num(v, 1)),
    ("annual_gen_kwh",   "年間発電量",   "kWh",      lambda v: fmt_num(v, 0)),
    ("annual_saving",    "年間削減額",   "円",       lambda v: fmt_yen(v, "")),
    ("co2_reduction_t",  "CO₂削減量",    "t-CO₂/年", lambda v: fmt_num(v, 1)),
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
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    company = data.get("company_name", "") or ""
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                f"{company}　様　|　建物別シミュレーション結果",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_SUB)
    y += Inches(0.35)

    # Get building data
    buildings = data.get("buildings")
    if not buildings or not isinstance(buildings, list) or len(buildings) < 2:
        buildings = DEFAULT_BUILDINGS

    bldg_a = buildings[0]
    bldg_b = buildings[1]

    # Two-column layout
    col_gap = Inches(0.25)
    col_w = (SLIDE_W - MARGIN * 2 - col_gap) / 2
    card_h = Inches(3.8)

    for idx, bldg in enumerate([bldg_a, bldg_b]):
        cx = MARGIN + idx * (col_w + col_gap)

        # Building card background
        add_rounded_rect(slide, cx, y, col_w, card_h, C_LIGHT_GRAY, radius_pt=6.0)

        # Building name header
        add_rounded_rect(slide, cx, y, col_w, Inches(0.5), C_ORANGE, radius_pt=6.0)
        # Square off bottom corners by overlaying a rect
        add_rect(slide, cx, y + Inches(0.3), col_w, Inches(0.2), C_ORANGE)
        add_textbox(slide, cx, y + Inches(0.08),
                    col_w, Inches(0.34),
                    bldg.get("name", f"建物{chr(65 + idx)}"),
                    font_name=FONT_BLACK, font_size_pt=18,
                    font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

        # KPI rows
        kpi_y = y + Inches(0.6)
        kpi_h = Inches(0.72)

        for k, (key, label, unit, fmt_fn) in enumerate(KPIS):
            ky = kpi_y + k * (kpi_h + Inches(0.08))
            row_bg = C_LIGHT_ORANGE if k % 2 == 0 else C_WHITE

            add_rounded_rect(slide, cx + Inches(0.1), ky,
                             col_w - Inches(0.2), kpi_h, row_bg, radius_pt=4.0)

            # Label
            add_textbox(slide, cx + Inches(0.18), ky + Inches(0.06),
                        col_w - Inches(0.36), Inches(0.2),
                        label,
                        font_name=FONT_BODY, font_size_pt=10,
                        font_color=C_SUB, bold=True,
                        align=PP_ALIGN.CENTER)

            # Value
            raw_val = bldg.get(key)
            display_val = fmt_fn(raw_val) if raw_val is not None else "—"
            add_textbox(slide, cx + Inches(0.18), ky + Inches(0.25),
                        col_w - Inches(0.36), Inches(0.3),
                        display_val,
                        font_name=FONT_BLACK, font_size_pt=20,
                        font_color=C_ORANGE, bold=True,
                        align=PP_ALIGN.CENTER)

            # Unit
            add_textbox(slide, cx + Inches(0.18), ky + Inches(0.53),
                        col_w - Inches(0.36), Inches(0.16),
                        unit,
                        font_name=FONT_BODY, font_size_pt=9,
                        font_color=C_SUB,
                        align=PP_ALIGN.CENTER)

    # Summary row: 合計効果
    summary_y = y + card_h + Inches(0.18)
    summary_h = Inches(0.65)
    add_rounded_rect(slide, MARGIN, summary_y,
                     SLIDE_W - MARGIN * 2, summary_h, C_ORANGE, radius_pt=6.0)

    add_textbox(slide, MARGIN + Inches(0.15), summary_y + Inches(0.06),
                Inches(1.5), Inches(0.2),
                "合計効果",
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_WHITE, bold=True)

    # Calculate totals
    total_cap = _safe_sum(bldg_a.get("capacity_kw"), bldg_b.get("capacity_kw"))
    total_gen = _safe_sum(bldg_a.get("annual_gen_kwh"), bldg_b.get("annual_gen_kwh"))
    total_saving = _safe_sum(bldg_a.get("annual_saving"), bldg_b.get("annual_saving"))
    total_co2 = _safe_sum(bldg_a.get("co2_reduction_t"), bldg_b.get("co2_reduction_t"))

    summaries = [
        (f"{fmt_num(total_cap, 1)} kW", "容量"),
        (f"{fmt_num(total_gen, 0)} kWh", "発電量"),
        (f"{fmt_yen(total_saving, '')} 円/年", "削減額"),
        (f"{fmt_num(total_co2, 1)} t", "CO₂"),
    ]

    item_w = (SLIDE_W - MARGIN * 2 - Inches(2.0)) / len(summaries)
    for s, (val_text, lbl) in enumerate(summaries):
        sx = MARGIN + Inches(2.0) + s * item_w
        add_textbox(slide, sx, summary_y + Inches(0.04),
                    item_w, Inches(0.3),
                    val_text,
                    font_name=FONT_BLACK, font_size_pt=14,
                    font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, sx, summary_y + Inches(0.36),
                    item_w, Inches(0.2),
                    lbl,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_WHITE,
                    align=PP_ALIGN.CENTER)

    add_footer(slide)


def _safe_sum(a, b):
    """Safely sum two values that may be None."""
    try:
        va = float(a) if a is not None else 0
        vb = float(b) if b is not None else 0
        return va + vb
    except (TypeError, ValueError):
        return None
