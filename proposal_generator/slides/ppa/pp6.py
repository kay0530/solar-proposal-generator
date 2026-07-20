"""
pp6.py - 発電シミュレーション (Design v2: Institutional Trust Grid)

Layout: A4 landscape
- White header (eyebrow + navy title + navy rule with orange tick)
- 4 KPI cards (cols 3+3+3+3): 年間発電量 / 自家消費量 / 自家消費率 / CO2削減
- Monthly generation bar chart spanning all 12 grid columns,
  sized to dominate the lower half of the content area
- Surplus electricity caption line at the bottom (when present)
"""
from __future__ import annotations

from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_ORANGE, C_SUB,
    FONT_BODY, GAP_BLOCK, GAP_CARD, MARGIN, SIZE_BODY, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_section_header,
    add_textbox, fmt_num, grid_w, grid_x, style_chart_base,
)

TITLE = "発電シミュレーション"
EYEBROW = "03｜効果シミュレーション"

# Typical monthly generation distribution (% of annual, approximate for Japan)
MONTHLY_PCT = [6.5, 7.0, 8.5, 9.5, 10.0, 9.5, 10.5, 10.0, 8.5, 8.0, 6.5, 5.5]
MONTH_NAMES = ["1月", "2月", "3月", "4月", "5月", "6月",
               "7月", "8月", "9月", "10月", "11月", "12月"]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP6 (generation simulation) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    annual_gen = data.get("annual_gen_kwh")
    self_kwh = data.get("self_consumption_kwh")
    self_pct = data.get("self_consumption_pct")
    surplus_kwh = data.get("surplus_kwh")
    _co2_raw = data.get("co2_annual_t")
    # Guard: co2 may be a descriptive string instead of a number
    try:
        co2_t = float(_co2_raw) if _co2_raw is not None else None
    except (ValueError, TypeError):
        co2_t = None

    # ---- KPI cards (4 across, cols 3+3+3+3) ----
    kpis = [
        (fmt_num(annual_gen, 0) if annual_gen else "—", "kWh/年", "年間発電量"),
        (fmt_num(self_kwh, 0) if self_kwh else "—", "kWh/年", "自家消費量"),
        (_fmt_pct(self_pct), "%", "自家消費率"),
        (fmt_num(co2_t, 1) if co2_t else "—", "t-CO₂/年", "年間CO₂削減量"),
    ]

    y = CONTENT_TOP
    kpi_h = Inches(1.00)
    for i, (number, unit, label) in enumerate(kpis):
        add_kpi_card(slide, grid_x(i * 3), y, grid_w(3), kpi_h,
                     number, unit, label)
    y = int(y) + int(kpi_h) + int(GAP_BLOCK)

    # ---- Monthly generation bar chart (12 cols, dominates lower half) ----
    add_section_header(slide, MARGIN, y, grid_w(12), "月別発電量（推定）")
    y += int(Inches(0.40))

    # Build monthly values
    monthly_kwh: list[float] = []
    raw_monthly = data.get("monthly_gen_kwh")
    if raw_monthly and len(raw_monthly) == 12:
        monthly_kwh = [float(v) for v in raw_monthly]
    elif annual_gen:
        for pct in MONTHLY_PCT:
            monthly_kwh.append(float(annual_gen) * pct / 100)

    surplus_h = int(Inches(0.26)) if surplus_kwh else 0
    chart_h = int(CONTENT_BOTTOM) - y - (surplus_h + int(GAP_CARD)
                                         if surplus_kwh else 0)

    if monthly_kwh:
        chart_data = CategoryChartData()
        chart_data.categories = MONTH_NAMES
        chart_data.add_series("月間発電量 (kWh)", monthly_kwh)

        chart_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            int(grid_x(0)), y, int(grid_w(12)), chart_h,
            chart_data,
        )
        chart = chart_frame.chart
        chart.has_legend = False

        # Bars: solid orange, gap_width 60
        series = chart.series[0]
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = C_ORANGE
        chart.plots[0].gap_width = 60

        style_chart_base(chart)
    else:
        add_textbox(slide, grid_x(0), y, grid_w(12), chart_h,
                    "発電量データ未入力",
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_SUB, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)

    # ---- Surplus electricity (caption line at content bottom) ----
    if surplus_kwh:
        surplus_text = f"余剰電力｜年間余剰電力量：{fmt_num(surplus_kwh, 0)} kWh"
        surplus_price = data.get("surplus_price")
        if surplus_price:
            surplus_text += f"　（売電単価：{fmt_num(surplus_price, 1)} 円/kWh）"
        add_textbox(slide, grid_x(0), int(CONTENT_BOTTOM) - surplus_h,
                    grid_w(12), surplus_h,
                    surplus_text,
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_DARK, anchor=MSO_ANCHOR.BOTTOM)

    add_footer(slide)


def _fmt_pct(val) -> str:
    """Format a percentage value (may be 0-1 float or 0-100 value)."""
    if val is None:
        return "—"
    try:
        v = float(val)
        # If value is <= 1 (with a small tolerance so an exact 1.0 ratio and
        # tiny float rounding errors are treated as a ratio, not an already-%
        # value), assume it's a 0-1 ratio and scale to percent.
        if v <= 1.0 + 1e-9:
            v *= 100
        return f"{v:.1f}"
    except (TypeError, ValueError):
        return "—"
