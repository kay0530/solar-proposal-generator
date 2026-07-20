"""
ep4.py - 経済効果試算（EPC） — design system v2

Mirrors the v2 PP8 layout, adapted for EPC economics:
- Inline 28pt metric band: 累計削減額 / 投資回収年 / 初期費用（補助金適用後）
  (payback is a valid EPC concept — customer owns the asset)
- Section header + trial-condition caption (契約電力, 従量単価, 基本料金,
  保守費用, 償却資産税)
- 20-year simulation table split into two stacked halves (1-10年 / 11-20年)
  at 9pt with 累積削減額 as the audited total row. 8pt note at bottom.
"""
from __future__ import annotations

from pathlib import Path

from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, MARGIN, SLIDE_W,
    C_HAIR, C_SUB,
    SIZE_CAPTION, SIZE_SMALL,
    TABLE_ROW_H,
    add_footer, add_header_bar, add_line, add_number_unit,
    add_section_header, add_table, add_textbox,
    fmt_yen, grid_w, grid_x, vstack,
)

TITLE = "経済効果試算"
EYEBROW = "03｜効果シミュレーション"
DEGRADATION = 0.005  # 0.5% annual degradation
SURCHARGE_DEFAULT = 3.60  # 賦課金+燃料費等調整 (円/kWh)


def _yen_parts(v: float) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for add_number_unit."""
    if v >= 1_0000_0000:
        return f"{v / 1_0000_0000:.2f}", "億円"
    if v >= 10_000:
        return f"{v / 10_000:,.0f}", "万円"
    return f"{v:,.0f}", "円"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render EP4 (economic effect simulation for EPC) - 20-year table format."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    # ---- Extract data ----
    elec_company = data.get("elec_company", "") or ""
    elec_contract = data.get("elec_contract", "") or ""
    contract_kw = float(data.get("contract_kw", 0) or 0)
    self_kwh = float(data.get("self_consumption_kwh", 0) or 0)
    demand_kw = float(data.get("demand_reduction_kw", 0) or 0)
    years = int(data.get("contract_years", 20) or 20)
    tax_display = data.get("tax_display", "税抜") or "税抜"

    # Investment costs
    selling_price = float(data.get("selling_price", 0) or 0)
    subsidy_amount = float(data.get("subsidy_amount", 0) or 0)
    system_kw = float(data.get("system_capacity_kw", 0) or 0)
    initial_cost = selling_price - subsidy_amount

    # Maintenance cost: prefer explicit value, fallback to system_kw * 1200
    annual_om_cost = float(data.get("annual_om_cost", 0) or 0)
    if annual_om_cost <= 0 and system_kw > 0:
        annual_om_cost = system_kw * 1200

    # Depreciation tax (from data if available)
    depreciation_tax = float(data.get("depreciation_tax_y1", 0) or 0)

    # Electricity rate calculation
    annual_cost = data.get("annual_cost")
    annual_kwh = float(data.get("annual_kwh", 0) or 0)

    if annual_cost and annual_kwh > 0:
        avg_unit_price = float(annual_cost) / annual_kwh
    else:
        avg_unit_price = 0

    # Basic charge for demand reduction — comes from app.py (contract master
    # / manual input). Do NOT try to derive it as annual_cost -
    # avg_unit_price*annual_kwh: avg_unit_price == annual_cost/annual_kwh
    # makes that an identity whose float noise could bypass the fallback
    # and zero out the demand saving.
    basic_rate_kw = float(data.get("basic_rate_kw", 0) or 0)
    if basic_rate_kw <= 0:
        basic_rate_kw = 1500.0  # last resort typical high-voltage basic rate

    y1_demand_saving = demand_kw * basic_rate_kw * 12 if demand_kw > 0 else 0

    # ---- Per-year simulation arrays ----
    sim_years = min(years, 20)
    per_year = []   # (supply, usage_saving, demand_saving, total)
    cum_by_year = []
    cumulative = 0.0
    for yr in range(1, sim_years + 1):
        supply = self_kwh * (1 - DEGRADATION) ** (yr - 1) if self_kwh > 0 else 0
        usage_saving = supply * avg_unit_price if avg_unit_price > 0 else 0
        demand_saving = y1_demand_saving
        total_s = usage_saving + demand_saving
        cumulative += total_s
        per_year.append((supply, usage_saving, demand_saving, total_s))
        cum_by_year.append(cumulative)

    # ---- Inline metric band items (28pt, EPC: payback is valid) ----
    cum_num, cum_unit = _yen_parts(cumulative) if cumulative else ("—", "")

    recovery = data.get("investment_recovery_yr")
    try:
        rec_f = float(recovery) if recovery is not None else None
    except (TypeError, ValueError):
        rec_f = None

    if initial_cost > 0:
        init_num, init_unit = _yen_parts(initial_cost)
    else:
        init_num, init_unit = "—", ""

    band = [
        (f"累計削減額（{sim_years}年間）", cum_num, cum_unit),
        ("投資回収年", f"{rec_f:.1f}" if rec_f else "—", "年" if rec_f else ""),
        ("初期費用（補助金適用後）", init_num, init_unit),
    ]

    # ---- Table builder (7 rows each half) ----
    def _build_half_table(yr_range: list[int]) -> list[list[str]]:
        header = [""] + [f"{yr}年目" for yr in yr_range]
        row_a = ["自家消費電力量(kWh)"]
        row_b = ["平均従量単価(円/kWh)"]
        row_d1 = ["従量料金削減額(円)"]
        row_d2 = ["基本料金削減額(円)"]
        row_d3 = ["年間削減合計(円)"]
        row_cum = ["累積削減額(円)"]
        for yr in yr_range:
            supply, usage_saving, demand_saving, total_s = per_year[yr - 1]
            row_a.append(f"{supply:,.0f}")
            row_b.append(f"{avg_unit_price:.2f}" if avg_unit_price > 0 else "—")
            row_d1.append(f"{usage_saving:,.0f}" if avg_unit_price > 0 else "—")
            row_d2.append(f"{demand_saving:,.0f}" if demand_saving > 0 else "—")
            row_d3.append(f"{total_s:,.0f}"
                          if (avg_unit_price > 0 or demand_saving > 0) else "—")
            cum_v = cum_by_year[yr - 1]
            row_cum.append(f"{cum_v:,.0f}" if cum_v > 0 else "—")
        return [header, row_a, row_b, row_d1, row_d2, row_d3, row_cum]

    first_years = list(range(1, min(10, sim_years) + 1))
    second_years = list(range(11, sim_years + 1)) if sim_years > 10 else []

    # ---- Vertical layout via vstack (exact table heights) ----
    band_h = Inches(0.72)
    sect_h = Inches(0.52)
    t1_h = TABLE_ROW_H * 7
    t2_h = TABLE_ROW_H * 7 if second_years else None
    note_h = Inches(0.26)
    blocks = [band_h, sect_h, t1_h] + ([t2_h] if t2_h else []) + [note_h]
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, blocks)

    # ---- Metric band (no cards, hairline separators) ----
    band_y = ys[0]
    span = 12 // len(band)
    for i, (label, number, unit) in enumerate(band):
        bx = grid_x(i * span)
        bw = grid_w(span) - Inches(0.20)
        add_textbox(slide, bx, band_y, bw, Inches(0.20),
                    label,
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB, bold=True)
        add_number_unit(slide, bx, band_y + Inches(0.22), bw, Inches(0.46),
                        number, unit)
        if i > 0:
            sep_x = bx - Inches(0.12)
            add_line(slide, sep_x, band_y + Inches(0.04),
                     sep_x, band_y + band_h - Inches(0.06),
                     C_HAIR, width_pt=0.5)

    # ---- Section header + trial conditions caption ----
    sect_y = ys[1]
    add_section_header(slide, MARGIN, sect_y, SLIDE_W - MARGIN * 2,
                       f"{sim_years}年間の削減効果")
    cond_head = (f"{elec_company} {elec_contract} {contract_kw:.0f}kW"
                 if elec_company else "未設定")
    unit_txt = f"{avg_unit_price:.2f}円/kWh" if avg_unit_price > 0 else "未設定"
    om_txt = fmt_yen(annual_om_cost) if annual_om_cost > 0 else "—"
    dep_txt = fmt_yen(depreciation_tax) if depreciation_tax > 0 else "—"
    cond = (
        f"試算条件：契約電力 {cond_head} ／ 従量単価 {unit_txt} ／ "
        f"基本料金 {basic_rate_kw:,.0f}円/kW・月（削減対象 {demand_kw:.0f}kW）／ "
        f"保守費用 {om_txt}/年 ／ 償却資産税（初年度）{dep_txt}"
    )
    add_textbox(slide, MARGIN, sect_y + Inches(0.30),
                SLIDE_W - MARGIN * 2, Inches(0.18),
                cond, font_size_pt=SIZE_SMALL, font_color=C_SUB)

    # ---- 20-year tables (1-10 / 11-20), exact TABLE_ROW_H advance ----
    table_w = SLIDE_W - MARGIN * 2
    label_col_w = Inches(1.65)

    def _render_half(y, yr_range):
        rows = _build_half_table(yr_range)
        n_data = len(yr_range)
        data_col_w = (table_w - label_col_w) / n_data
        col_widths = [label_col_w] + [data_col_w] * n_data
        add_table(slide, MARGIN, y, table_w, rows, col_widths,
                  font_size_pt=SIZE_CAPTION, total_row=len(rows) - 1)

    _render_half(ys[2], first_years)
    if second_years:
        _render_half(ys[3], second_years)

    # ---- Note (8pt) ----
    note = (f"※ 金額は全て{tax_display}表記。発電量は年0.5%の経年劣化を考慮して試算。"
            "保守費用・償却資産税は年間削減額に含みません。")
    add_textbox(slide, MARGIN, ys[-1], SLIDE_W - MARGIN * 2, note_h,
                note, font_size_pt=SIZE_SMALL, font_color=C_SUB)

    add_footer(slide)
