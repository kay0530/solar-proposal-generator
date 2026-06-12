"""
ep5.py - デマンドカット試算（EPC） — design system v2

Same logic as v2 PP9 but titled for the EPC model.
Reuses demand_calc engine for peak detection and chart data.

Layout (A4 landscape):
  - White header (eyebrow + navy title + navy rule with orange tick)
  - Consolidated row: 3 KPI cards (before / after / cut) + savings band
    (C_PANEL + orange left bar + 28pt number + 8pt calc detail)
  - 2-panel line chart (before/after PV) showing 2-week demand profile:
      before demand = gray dashed, after demand = navy solid,
      self-consumption = orange filled area, peak ref = navy thin dashed
  - Falls back to manual demand_reduction_kw when no iPals data
"""

from __future__ import annotations

from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from proposal_generator.demand_calc import calc_demand_cut
from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_NAVY, C_ORANGE, C_PANEL, C_SUB,
    FONT_BODY, GAP_CARD, GAP_IN_CARD, MARGIN, SIZE_BODY, SIZE_CAPTION,
    SIZE_SMALL, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_multiline_textbox,
    add_number_unit, add_rect, add_section_header, add_textbox,
    fmt_num, fmt_yen, rotate_category_labels, style_chart_base,
    style_series_before, vstack,
)

TITLE = "デマンドカット試算（EPC）"
EYEBROW = "03｜効果シミュレーション"

# Fallback unit price when no contract master data
DEMAND_UNIT_PRICE_FALLBACK = 1879.72


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render EP5 (demand cut simulation for EPC) onto an already-added blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

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

    y = CONTENT_TOP

    # ---- Section header + consolidated row: 3 KPI cards + savings band ----
    add_section_header(slide, MARGIN, y, Inches(5.0), "デマンドカット効果")
    y = int(y) + int(Inches(0.42))

    total_w = SLIDE_W - MARGIN * 2
    kpi_area_w = int(total_w) * 45 // 100
    card_w = (kpi_area_w - int(GAP_IN_CARD) * 2) // 3
    card_h = Inches(1.05)

    add_kpi_card(slide, MARGIN, y, card_w, card_h,
                 fmt_num(peak_before, 0), "kW", "導入前ピーク")
    add_kpi_card(slide, int(MARGIN) + card_w + int(GAP_IN_CARD), y,
                 card_w, card_h,
                 fmt_num(peak_after, 0), "kW", "導入後ピーク")
    add_kpi_card(slide, int(MARGIN) + (card_w + int(GAP_IN_CARD)) * 2, y,
                 card_w, card_h,
                 f"▲{fmt_num(demand_cut, 0)}", "kW", "デマンド削減量")

    # Savings band: C_PANEL + orange left bar + 28pt number + 8pt detail
    savings_x = int(MARGIN) + kpi_area_w + int(GAP_CARD)
    savings_w = int(SLIDE_W) - int(MARGIN) - savings_x
    add_rect(slide, savings_x, y, savings_w, card_h, C_PANEL)
    add_rect(slide, savings_x, y, Inches(0.05), card_h, C_ORANGE)

    amount_x = savings_x + int(Inches(0.18))
    amount_w = int(savings_w * 0.40)
    add_textbox(slide, amount_x, y + int(Inches(0.12)),
                amount_w, Inches(0.20),
                "基本料金削減効果",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, bold=True)
    add_number_unit(slide, amount_x, y + int(Inches(0.30)),
                    amount_w, int(card_h) - int(Inches(0.44)),
                    fmt_yen(annual_saving), "/年")

    calc_x = amount_x + amount_w + int(Inches(0.10))
    calc_w = savings_x + savings_w - calc_x - int(Inches(0.15))
    calc_lines = [
        (f"基本料金単価: {fmt_num(basic_rate, 1)} 円/kW × 力率補正: {pf_factor:.2f}",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        (f"月額削減: ▲{fmt_num(demand_cut, 0)} kW × {fmt_num(basic_rate, 1)} 円"
         f" × {pf_factor:.2f} = {fmt_yen(monthly_saving)}/月",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
        (f"年間削減: {fmt_yen(monthly_saving)} × 12 = {fmt_yen(annual_saving)}/年",
         FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT),
    ]
    add_multiline_textbox(slide, calc_x, y + int(Inches(0.20)),
                          calc_w, int(card_h) - int(Inches(0.32)),
                          calc_lines, line_spacing=1.35)

    y = y + int(card_h) + int(GAP_CARD)

    # ---- Line charts (2 panels: before / after) ----
    if chart_before and chart_after:
        chart_w = total_w
        panel_h = (int(CONTENT_BOTTOM) - y - int(GAP_IN_CARD)) // 2

        _add_demand_chart(slide, MARGIN, y, chart_w, panel_h,
                          "PV導入前 デマンド推移", chart_before, peak_before,
                          mode="before")
        _add_demand_chart(slide, MARGIN, y + panel_h + int(GAP_IN_CARD),
                          chart_w, panel_h,
                          "PV導入後 デマンド推移", chart_after, peak_after,
                          mode="after")
    elif not has_ipals:
        note_h = Inches(0.60)
        ys = vstack(y, CONTENT_BOTTOM, [note_h])
        add_rect(slide, MARGIN, ys[0], total_w, note_h, C_PANEL)
        add_textbox(slide, MARGIN, ys[0], total_w, note_h,
                    "※ iPals CSVをアップロードすると、2週間のデマンド推移グラフが表示されます。",
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_SUB, align=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)

    add_footer(slide)


def _add_demand_chart(slide, x, y, w, h, title: str,
                      chart_data_list: list[dict], peak_kw: float,
                      mode: str = "before") -> None:
    """Add a line chart showing demand profile with a peak reference line.

    mode="before": demand line rendered gray dashed (pre-PV baseline).
    mode="after":  demand line rendered navy solid (post-PV result) to
                   avoid double-orange with the orange self-consumption area.
    """
    # Extract data arrays
    labels = [d["label"] for d in chart_data_list]
    values = [d["value"] for d in chart_data_list]
    self_c_values = [d.get("self_c", 0) for d in chart_data_list]

    # Chart title
    add_textbox(slide, x, y, w, Inches(0.22),
                title,
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_NAVY, bold=True)
    y = int(y) + int(Inches(0.24))
    h = int(h) - int(Inches(0.24))

    # Use full hourly resolution for Streamlit-level detail
    step = 1 if len(values) <= 360 else max(1, len(values) // 336)
    sampled_labels = labels[::step]
    sampled_values = values[::step]
    sampled_self_c = self_c_values[::step]

    cd = CategoryChartData()
    # Show date only on first sample of each day to avoid duplicates
    display_labels = []
    _last_date = None
    for lbl in sampled_labels:
        _date = lbl.split(" ")[0] if " " in lbl else lbl
        if _date != _last_date:
            display_labels.append(_date)
            _last_date = _date
        else:
            display_labels.append("")
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

    plot = chart.plots[0]
    series_demand = plot.series[0]
    if mode == "before":
        # Pre-PV baseline: gray dashed (grayscale-safe "before" encoding)
        style_series_before(series_demand)
    else:
        # Post-PV result: navy solid (avoids double-orange with area)
        series_demand.format.line.color.rgb = C_NAVY
        series_demand.format.line.width = Pt(2.25)
        series_demand.smooth = False

    series_self_c = plot.series[1]
    series_self_c.format.line.color.rgb = C_ORANGE
    series_self_c.format.line.width = Pt(1.0)
    series_self_c.smooth = False

    # Peak reference: navy thin dashed (no pure red in v2)
    series_peak = plot.series[2]
    series_peak.format.line.color.rgb = C_NAVY
    series_peak.format.line.width = Pt(1.0)
    series_peak.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    series_peak.smooth = False

    # Convert self_c series (index 1) to an areaChart so it shows as filled
    # orange area below the line, matching the Streamlit UI appearance.
    try:
        from lxml import etree as _etree
        _chartSpace = chart._chartSpace
        _ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        _a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        _plotArea = _chartSpace.find(f".//{{{_ns}}}plotArea")
        _lineChart = _plotArea.find(f"{{{_ns}}}lineChart")
        _sers = _lineChart.findall(f"{{{_ns}}}ser")
        _self_c_ser = _sers[1]
        _lineChart.remove(_self_c_ser)

        # Build areaChart element
        _area_xml = (
            f'<c:areaChart xmlns:c="{_ns}" xmlns:a="{_a_ns}">'
            f'<c:grouping val="standard"/>'
            f'<c:varyColors val="0"/>'
            f'</c:areaChart>'
        )
        _areaChart = _etree.fromstring(_area_xml)

        # Set orange solidFill on the series
        _spPr_xml = (
            f'<c:spPr xmlns:c="{_ns}" xmlns:a="{_a_ns}">'
            f'<a:solidFill><a:srgbClr val="E8490F"><a:alpha val="55000"/></a:srgbClr></a:solidFill>'
            f'<a:ln><a:solidFill><a:srgbClr val="E8490F"/></a:solidFill></a:ln>'
            f'</c:spPr>'
        )
        _spPr_new = _etree.fromstring(_spPr_xml)
        # Insert spPr after c:tx / c:order. If existing spPr present, replace.
        _existing_spPr = _self_c_ser.find(f"{{{_ns}}}spPr")
        if _existing_spPr is not None:
            _self_c_ser.remove(_existing_spPr)
        _tx = _self_c_ser.find(f"{{{_ns}}}tx")
        if _tx is not None:
            _tx.addnext(_spPr_new)
        else:
            _self_c_ser.insert(0, _spPr_new)

        _areaChart.append(_self_c_ser)

        # Copy axis references from lineChart to areaChart
        for _axId in _lineChart.findall(f"{{{_ns}}}axId"):
            _areaChart.append(_etree.fromstring(_etree.tostring(_axId)))

        # Insert areaChart BEFORE lineChart so it renders behind
        _lineChart.addprevious(_areaChart)
    except Exception:
        # If XML manipulation fails, fall back to styled line (no fill)
        series_self_c.format.line.width = Pt(1.5)

    # v2 chart chrome: frameless, dashed warm gridlines, 9pt axes/legend
    style_chart_base(chart)
    rotate_category_labels(chart, -45)
