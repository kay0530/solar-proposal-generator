"""
pp9.py - デマンドカット試算スライド

Layout (A4 landscape):
  - Orange header bar with "デマンドカット試算"
  - 3 KPI cards: before peak / after peak / demand cut
  - Basic fee savings calculation box
  - 2-panel line chart (before/after PV) showing 2-week demand profile
  - Falls back to manual demand_reduction_kw when no iPals data
"""

from __future__ import annotations

from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from proposal_generator.demand_calc import calc_demand_cut
from proposal_generator.utils import (
    CONTENT_H, CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_NAVY,
    C_ORANGE, C_RED, C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, HEADER_H,
    MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_rect, add_rounded_rect,
    add_section_header, add_textbox, fmt_num, fmt_yen,
)

TITLE = "デマンドカット試算"

# Fallback unit price when no contract master data
DEMAND_UNIT_PRICE_FALLBACK = 1879.72


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP9 (demand cut simulation) onto an already-added blank slide."""
    add_header_bar(slide, TITLE, logo_path)

    hourly_rows = data.get("hourly_rows")
    basic_rate = float(data.get("basic_rate_kw", 0) or 0)
    pf_pct = int(data.get("power_factor_pct", 85) or 85)

    if basic_rate <= 0:
        basic_rate = DEMAND_UNIT_PRICE_FALLBACK

    has_ipals = hourly_rows and len(hourly_rows) > 0

    if has_ipals:
        result = calc_demand_cut(hourly_rows, basic_rate, pf_pct)
        peak_before = result["peak_before_kw"]
        peak_after = result["peak_after_kw"]
        demand_cut = result["demand_cut_kw"]
        monthly_saving = result["monthly_basic_saving"]
        annual_saving = result["annual_basic_saving"]
        pf_factor = result["pf_factor"]
        chart_before = result["peak_week_before"]
        chart_after = result["peak_week_after"]
    else:
        # Fallback to manual input
        reduction_kw = float(data.get("demand_reduction_kw", 0) or 0)
        capacity_kw = float(data.get("system_capacity_kw", 0) or 0)
        peak_before = reduction_kw * 3 if reduction_kw else capacity_kw * 0.8
        demand_cut = reduction_kw
        peak_after = peak_before - demand_cut
        pf_factor = (185 - pf_pct) / 100
        monthly_saving = demand_cut * basic_rate * pf_factor
        annual_saving = monthly_saving * 12
        chart_before = []
        chart_after = []

    y = CONTENT_TOP + Inches(0.05)

    # ---- Header + 4-column row: 3 KPI cards + savings box ----
    add_section_header(slide, MARGIN, y, Inches(5.0), "デマンドカット効果")
    y += Inches(0.4)

    # Split row: 3 KPI cards take ~55% width, savings box takes ~45%
    gap = Inches(0.15)
    total_w = SLIDE_W - MARGIN * 2
    kpi_area_w = total_w * 0.45
    savings_area_w = total_w * 0.55 - gap
    card_w = (kpi_area_w - gap * 2) / 3
    card_h = Inches(1.1)

    add_kpi_card(slide, MARGIN, y, card_w, card_h,
                 f"{fmt_num(peak_before, 0)}", "kW",
                 "①導入前ピークデマンド",
                 bg_color=C_LIGHT_GRAY, number_size_pt=24)

    add_kpi_card(slide, MARGIN + card_w + gap, y, card_w, card_h,
                 f"{fmt_num(peak_after, 0)}", "kW",
                 "②導入後ピークデマンド",
                 bg_color=C_LIGHT_GRAY, number_size_pt=24)

    add_kpi_card(slide, MARGIN + (card_w + gap) * 2, y, card_w, card_h,
                 f"▲{fmt_num(demand_cut, 0)}", "kW",
                 "デマンド削減量",
                 bg_color=C_LIGHT_ORANGE, number_size_pt=24)

    # Savings box (right side of KPI row)
    savings_x = MARGIN + kpi_area_w + gap
    savings_h = card_h
    add_rounded_rect(slide, savings_x, y, savings_area_w, savings_h, C_LIGHT_ORANGE)
    add_rect(slide, savings_x, y, Inches(0.06), savings_h, C_ORANGE)

    # Label + amount (left half of savings box)
    amount_w = savings_area_w * 0.45
    add_textbox(slide, savings_x + Inches(0.15), y + Inches(0.06),
                amount_w, Inches(0.22),
                "基本料金削減効果",
                font_name=FONT_BODY, font_size_pt=10, font_color=C_DARK, bold=True)
    add_textbox(slide, savings_x + Inches(0.15), y + Inches(0.30),
                amount_w, Inches(0.55),
                fmt_yen(annual_saving) + "/年",
                font_name=FONT_BLACK, font_size_pt=24, font_color=C_ORANGE, bold=True)

    # Calc detail (right half of savings box)
    calc_text = (
        f"基本料金単価: {fmt_num(basic_rate, 1)} 円/kW × 力率補正: {pf_factor:.2f}\n"
        f"月額削減: ▲{fmt_num(demand_cut, 0)} kW × {fmt_num(basic_rate, 1)} 円 × {pf_factor:.2f}\n"
        f"        = {fmt_yen(monthly_saving)}/月\n"
        f"年間削減: {fmt_yen(monthly_saving)} × 12 = {fmt_yen(annual_saving)}/年"
    )
    calc_x = savings_x + amount_w + Inches(0.1)
    calc_w = savings_area_w - amount_w - Inches(0.25)
    add_textbox(slide, calc_x, y + Inches(0.08),
                calc_w, Inches(1.0),
                calc_text,
                font_name=FONT_BODY, font_size_pt=7, font_color=C_SUB)

    y += card_h + Inches(0.15)

    # Keep savings_w var defined for fallback text message below
    savings_w = total_w

    # ---- Line charts (2 panels: before / after) ----
    if chart_before and chart_after:
        chart_w = SLIDE_W - MARGIN * 2
        chart_h = (SLIDE_H - y - Inches(0.5)) / 2  # split remaining space

        _add_demand_chart(slide, MARGIN, y, chart_w, chart_h,
                          "PV導入前 デマンド推移", chart_before, peak_before)
        y += chart_h + Inches(0.08)

        _add_demand_chart(slide, MARGIN, y, chart_w, chart_h,
                          "PV導入後 デマンド推移", chart_after, peak_after)
    elif not has_ipals:
        add_textbox(slide, MARGIN, y, savings_w, Inches(0.5),
                    "※ iPals CSVをアップロードすると、2週間のデマンド推移グラフが表示されます。",
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB)

    add_footer(slide)


def _add_demand_chart(slide, x, y, w, h, title: str,
                      chart_data_list: list[dict], peak_kw: float) -> None:
    """Add a line chart showing demand profile with a peak reference line."""
    # Extract data arrays
    labels = [d["label"] for d in chart_data_list]
    values = [d["value"] for d in chart_data_list]
    self_c_values = [d.get("self_c", 0) for d in chart_data_list]

    # Chart title
    add_textbox(slide, x, y, w, Inches(0.22),
                f"◆ {title}",
                font_name=FONT_BODY, font_size_pt=9, font_color=C_DARK, bold=True)
    y += Inches(0.22)
    h -= Inches(0.22)

    # Sample down to ~56 points (every 6 hours)
    step = max(1, len(values) // 56) if values else 1
    sampled_labels = labels[::step]
    sampled_values = values[::step]
    sampled_self_c = self_c_values[::step]

    cd = CategoryChartData()
    # Show date on every label - powerpoint will auto-thin ticks
    display_labels = [lbl.split(" ")[0] if " " in lbl else lbl for lbl in sampled_labels]
    cd.categories = display_labels

    cd.add_series("使用電力量 (kW)", sampled_values)
    cd.add_series("自家消費量 (kW)", sampled_self_c)
    cd.add_series("ピークライン", [peak_kw] * len(sampled_values))

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE, int(x), int(y), int(w), int(h), cd
    )
    chart = chart_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    # Style: demand=navy, self_c=orange, peak=red dashed
    from pptx.dml.color import RGBColor
    plot = chart.plots[0]
    series_demand = plot.series[0]
    series_demand.format.line.color.rgb = RGBColor(0x00, 0x20, 0x60)  # navy
    series_demand.format.line.width = Pt(1.5)
    series_demand.smooth = False

    series_self_c = plot.series[1]
    series_self_c.format.line.color.rgb = RGBColor(0xE8, 0x49, 0x0F)  # orange
    series_self_c.format.line.width = Pt(1.0)
    series_self_c.smooth = False

    series_peak = plot.series[2]
    series_peak.format.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # red
    series_peak.format.line.width = Pt(1.0)
    series_peak.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    series_peak.smooth = False

    # Value axis
    value_axis = chart.value_axis
    value_axis.has_title = False
    value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)

    # Category axis: reduce font size
    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(7)
    # cat_axis.tick_label_position left at default for better rendering
