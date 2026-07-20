"""
pp8.py - 経済効果試算 (Economic effect simulation) — design system v2

Top: inline metric band (累計削減額 / 投資回収年 / IRR if present) as
28pt number+unit pairs with hairline separators (no cards).
Then: section header + trial-condition caption, and the 20-year saving
table split into two stacked tables (1-10年 / 11-20年) at 9pt with the
per-year total rendered as an audited total row. 8pt note at bottom.
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
    grid_w, grid_x, vstack,
)

TITLE = "経済効果試算"
EYEBROW = "04｜ご契約条件"
DEGRADATION = 0.005  # 0.5% annual degradation
SURCHARGE_DEFAULT = 3.60  # 賦課金+燃料費等調整 (円/kWh)


def _yen_parts(v: float) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for add_number_unit."""
    if v >= 1_0000_0000:
        return f"{v / 1_0000_0000:.2f}", "億円"
    if v >= 10_000:
        return f"{v / 10_000:,.0f}", "万円"
    return f"{v:,.0f}", "円"


def _fmt_amt(v: float) -> str:
    """Format a yen amount; negatives via ▲ (auto-colored by add_table)."""
    if v < 0:
        return f"▲{abs(v):,.0f}"
    return f"{v:,.0f}"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP8 (economic effect simulation) - 20-year table format."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    # ---- Extract data ----
    elec_company = data.get("elec_company", "") or ""
    elec_contract = data.get("elec_contract", "") or ""
    contract_kw = float(data.get("contract_kw", 0) or 0)
    self_kwh = float(data.get("self_consumption_kwh", 0) or 0)
    ppa_price = float(data.get("ppa_unit_price", 0) or 0)
    demand_kw = float(data.get("demand_reduction_kw", 0) or 0)
    years = int(data.get("contract_years", 20) or 20)
    tax_display = data.get("tax_display", "税抜") or "税抜"

    annual_cost = data.get("annual_cost")
    annual_kwh = float(data.get("annual_kwh", 0) or 0)

    if annual_cost and annual_kwh > 0:
        avg_unit_price = float(annual_cost) / annual_kwh
    else:
        avg_unit_price = 0

    surcharge = SURCHARGE_DEFAULT
    elec_rate = max(avg_unit_price - surcharge, 0) if avg_unit_price > 0 else 0
    total_unit = elec_rate + surcharge if avg_unit_price > 0 else 0

    # basic_rate_kw comes from app.py (contract master / manual input).
    # Do NOT try to derive it as annual_cost - avg_unit_price*annual_kwh:
    # avg_unit_price == annual_cost/annual_kwh makes that an identity whose
    # float noise could bypass the fallback and zero out the demand saving.
    basic_rate_kw = float(data.get("basic_rate_kw", 0) or 0)
    if basic_rate_kw <= 0:
        basic_rate_kw = 1500.0  # last resort typical high-voltage basic rate

    y1_demand_saving = 0.0
    if self_kwh > 0 and total_unit > 0:
        y1_demand_saving = demand_kw * basic_rate_kw * 12

    # ---- Per-year simulation arrays ----
    sim_years = min(years, 20)
    per_year = []  # (supply, usage_saving, demand_saving, total)
    cumulative = 0.0
    for yr in range(1, sim_years + 1):
        supply = self_kwh * (1 - DEGRADATION) ** (yr - 1) if self_kwh > 0 else 0
        usage_saving = supply * (total_unit - ppa_price) if total_unit > 0 else 0
        demand_saving = y1_demand_saving
        total_s = usage_saving + demand_saving
        per_year.append((supply, usage_saving, demand_saving, total_s))
        if total_unit > 0 or demand_saving > 0:
            cumulative += total_s

    # ---- Inline metric band items ----
    if cumulative:
        cum_num, cum_unit = _yen_parts(cumulative)
    else:
        cum_num, cum_unit = "—", ""
    band = [(f"累計削減額（{sim_years}年間）", cum_num, cum_unit)]

    # 投資回収年 is an EPC concept — PPA customers make no upfront
    # investment, so showing a payback period is misleading. Show the
    # annual average saving instead on PPA decks.
    is_ppa = str(data.get("proposal_type", "ppa")).lower() != "epc"
    if is_ppa:
        if cumulative and sim_years:
            avg_num, avg_unit = _yen_parts(cumulative / sim_years)
            band.append(("年間平均削減額", avg_num, avg_unit))
    else:
        recovery = data.get("investment_recovery_yr")
        try:
            rec_f = float(recovery) if recovery is not None else None
        except (TypeError, ValueError):
            rec_f = None
        band.append(("投資回収年", f"{rec_f:.1f}" if rec_f else "—",
                     "年" if rec_f else ""))

    irr_v = data.get("irr")
    if isinstance(irr_v, (int, float)) and irr_v:
        band.append(("IRR（参考）", f"{irr_v * 100:.1f}", "%"))

    # ---- Table builder (7 rows each half) ----
    def _build_half_table(yr_range: list[int]) -> list[list[str]]:
        header = [""] + [f"{yr}年目" for yr in yr_range]
        row_a = ["供給電力量(kWh)"]
        row_b = ["従量単価(円/kWh)"]
        row_c = ["PPA単価(円/kWh)"]
        row_d1 = ["従量料金削減(円)"]
        row_d2 = ["基本料金削減(円)"]
        row_d3 = ["合計削減額(円)"]
        for yr in yr_range:
            supply, usage_saving, demand_saving, total_s = per_year[yr - 1]
            row_a.append(f"{supply:,.0f}")
            row_b.append(f"{total_unit:.2f}" if total_unit > 0 else "—")
            row_c.append(f"{ppa_price:.2f}" if ppa_price > 0 else "—")
            row_d1.append(_fmt_amt(usage_saving) if total_unit > 0 else "—")
            row_d2.append(_fmt_amt(demand_saving) if demand_saving > 0 else "—")
            row_d3.append(_fmt_amt(total_s)
                          if (total_unit > 0 or demand_saving > 0) else "—")
        return [header, row_a, row_b, row_c, row_d1, row_d2, row_d3]

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
                 if elec_company else "契約電力未設定")
    cond = (
        f"試算条件：契約電力 {cond_head} ／ "
        f"従量単価 {total_unit:.2f}円/kWh（電力量料金{elec_rate:.2f}＋賦課金等{surcharge:.2f}）／ "
        f"PPA単価 {ppa_price:.2f}円/kWh ／ "
        f"基本料金 {basic_rate_kw:,.0f}円/kW・月（削減対象 {demand_kw:.0f}kW）"
    )
    add_textbox(slide, MARGIN, sect_y + Inches(0.30),
                SLIDE_W - MARGIN * 2, Inches(0.18),
                cond, font_size_pt=SIZE_SMALL, font_color=C_SUB)

    # ---- 20-year tables (1-10 / 11-20), exact TABLE_ROW_H advance ----
    table_w = SLIDE_W - MARGIN * 2
    label_col_w = Inches(1.55)

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
            "初期費用・保守点検費用・償却資産税は0円（PPA事業者負担）です。")
    if any(p[1] < 0 or p[3] < 0 for p in per_year):
        note += "　※ ▲はマイナスを表します。"
    add_textbox(slide, MARGIN, ys[-1], SLIDE_W - MARGIN * 2, note_h,
                note, font_size_pt=SIZE_SMALL, font_color=C_SUB)

    add_footer(slide)
