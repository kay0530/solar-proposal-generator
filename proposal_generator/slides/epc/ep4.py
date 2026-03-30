"""
ep4.py - 経済効果試算 (EPC)

Reproduces the PP8 20-year simulation table format, adapted for EPC economics:
- Trial conditions box (電力会社, 契約種別, 契約電力, 従量単価)
- Initial-year KPI cards (従量料金削減, 基本料金削減, 初期費用, 保守費用, 償却資産税)
- 20-year simulation table split into two halves (1-10, 11-20)
  Rows: (A)自家消費電力量, (B)平均従量単価, 従量料金削減額, 基本料金削減額, 年間削減合計, 累積削減額
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, C_NAVY, C_LIGHT_CYAN, C_RED,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, add_table, add_kpi_card, fmt_yen, fmt_num,
)

TITLE = "経済効果試算"
DEGRADATION = 0.005  # 0.5% annual degradation
SURCHARGE_DEFAULT = 3.60  # 賦課金+燃料費等調整 (円/kWh)


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render EP4 (economic effect simulation for EPC) - 20-year table format."""
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP

    # ---- Extract data ----
    elec_company = data.get("elec_company", "")
    elec_contract = data.get("elec_contract", "")
    contract_kw = float(data.get("contract_kw", 0) or 0)
    self_kwh = float(data.get("self_consumption_kwh", 0) or 0)
    demand_kw = float(data.get("demand_reduction_kw", 0) or 0)
    years = int(data.get("contract_years", 20) or 20)
    tax_display = data.get("tax_display", "税抜")

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

    # Calculate average unit price from annual cost
    if annual_cost and annual_kwh > 0:
        avg_unit_price = float(annual_cost) / annual_kwh
    else:
        avg_unit_price = 0

    # Basic charge for demand reduction - prefer explicit value from electricity master
    basic_rate_kw = float(data.get("basic_rate_kw", 0) or 0)
    if basic_rate_kw <= 0 and annual_cost and contract_kw > 0 and annual_kwh > 0:
        # Fallback: estimate from total cost
        usage_cost = avg_unit_price * annual_kwh
        basic_annual = float(annual_cost) - usage_cost
        if basic_annual > 0:
            basic_rate_kw = basic_annual / contract_kw / 12
    if basic_rate_kw <= 0:
        basic_rate_kw = 1500.0  # last resort typical high-voltage basic rate

    # ---- Trial conditions box ----
    cond_h = Inches(0.55)
    add_rounded_rect(slide, MARGIN, y, SLIDE_W - MARGIN * 2, cond_h, C_LIGHT_GRAY)
    add_textbox(slide, MARGIN + Inches(0.1), y + Inches(0.03),
                Inches(1.2), Inches(0.22),
                "★試算条件", font_name=FONT_BODY, font_size_pt=9,
                font_color=C_ORANGE, bold=True)

    cond_text = (
        f"契約電力: {elec_company} {elec_contract} {contract_kw:.0f}kW"
        if elec_company else "契約電力: 未設定"
    )
    add_textbox(slide, MARGIN + Inches(0.1), y + Inches(0.22),
                SLIDE_W - MARGIN * 2 - Inches(0.2), Inches(0.18),
                cond_text, font_name=FONT_BODY, font_size_pt=8, font_color=C_DARK)

    cond2 = f"従量単価: {avg_unit_price:.2f}円/kWh" if avg_unit_price > 0 else "従量単価: 未設定"
    add_textbox(slide, MARGIN + Inches(0.1), y + Inches(0.35),
                SLIDE_W - MARGIN * 2 - Inches(0.2), Inches(0.18),
                cond2, font_name=FONT_BODY, font_size_pt=8, font_color=C_DARK)

    y += cond_h + Inches(0.08)

    # ---- Initial year KPIs ----
    y1_usage_saving = self_kwh * avg_unit_price if (self_kwh > 0 and avg_unit_price > 0) else 0
    y1_demand_saving = demand_kw * basic_rate_kw * 12 if demand_kw > 0 else 0

    kpi_data = [
        (fmt_yen(y1_usage_saving) if y1_usage_saving else "—", "従量料金削減"),
        (fmt_yen(y1_demand_saving) if y1_demand_saving else "—", "基本料金削減"),
        (fmt_yen(initial_cost) if initial_cost > 0 else "—", "初期費用"),
        (fmt_yen(annual_om_cost) if annual_om_cost > 0 else "—", "保守費用"),
        (fmt_yen(depreciation_tax) if depreciation_tax > 0 else "—", "償却資産税"),
    ]

    kpi_w = (SLIDE_W - MARGIN * 2 - Inches(0.1) * 4) / 5
    kpi_h = Inches(0.65)
    for i, (val, label) in enumerate(kpi_data):
        kx = MARGIN + i * (kpi_w + Inches(0.1))
        add_rounded_rect(slide, kx, y, kpi_w, kpi_h, C_LIGHT_ORANGE)
        add_textbox(slide, kx, y + Inches(0.05), kpi_w, Inches(0.30),
                    val, font_name=FONT_BLACK, font_size_pt=14,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)
        add_textbox(slide, kx, y + Inches(0.38), kpi_w, Inches(0.22),
                    label, font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

    y += kpi_h + Inches(0.12)

    # ---- 20-year simulation table ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       f"◆{years}年間の削減効果", font_size_pt=10)
    y += Inches(0.25)

    # Build simulation data
    sim_years = min(years, 20)
    half = (sim_years + 1) // 2  # split point

    def _build_half_table(start_yr: int, end_yr: int) -> list[list[str]]:
        """Build table rows for a range of years."""
        yr_range = list(range(start_yr, end_yr + 1))

        header = [""] + [f"{yr}年目" for yr in yr_range]
        row_a = ["(A)自家消費電力量(kWh)"]
        row_b = ["(B)平均従量単価(円/kWh)"]
        row_d1 = ["従量料金削減額(円)"]
        row_d2 = ["基本料金削減額(円)"]
        row_d3 = ["年間削減合計(円)"]
        row_cum = ["累積削減額(円)"]

        # Calculate cumulative up to start_yr - 1
        cumulative = 0
        for yr in range(1, start_yr):
            supply = self_kwh * (1 - DEGRADATION) ** (yr - 1) if self_kwh > 0 else 0
            usage_s = supply * avg_unit_price if avg_unit_price > 0 else 0
            demand_s = y1_demand_saving
            cumulative += usage_s + demand_s

        for yr in yr_range:
            # (A) self-consumption with degradation
            supply = self_kwh * (1 - DEGRADATION) ** (yr - 1) if self_kwh > 0 else 0
            row_a.append(f"{supply:,.0f}")

            # (B) average unit price (constant)
            row_b.append(f"{avg_unit_price:.2f}" if avg_unit_price > 0 else "—")

            # Savings
            usage_saving = supply * avg_unit_price if avg_unit_price > 0 else 0
            demand_saving = y1_demand_saving  # constant
            total_s = usage_saving + demand_saving
            cumulative += total_s

            row_d1.append(f"{usage_saving:,.0f}" if avg_unit_price > 0 else "—")
            row_d2.append(f"{demand_saving:,.0f}" if demand_saving > 0 else "—")
            row_d3.append(f"{total_s:,.0f}" if (avg_unit_price > 0 or demand_saving > 0) else "—")
            row_cum.append(f"{cumulative:,.0f}" if cumulative > 0 else "—")

        return [header, row_a, row_b, row_d1, row_d2, row_d3, row_cum]

    # First half: years 1 to half
    table1 = _build_half_table(1, half)
    n_cols1 = len(table1[0])
    table_w = SLIDE_W - MARGIN * 2
    label_col_w = Inches(2.2)
    data_col_w = (table_w - label_col_w) / (n_cols1 - 1) if n_cols1 > 1 else Inches(0.8)
    col_widths1 = [label_col_w] + [data_col_w] * (n_cols1 - 1)

    add_table(slide, MARGIN, y, table_w, table1, col_widths1, font_size_pt=7)
    table1_h = Inches(0.22) * len(table1)
    y += table1_h + Inches(0.1)

    # Second half: years half+1 to sim_years
    if half < sim_years:
        table2 = _build_half_table(half + 1, sim_years)
        n_cols2 = len(table2[0])
        data_col_w2 = (table_w - label_col_w) / (n_cols2 - 1) if n_cols2 > 1 else Inches(0.8)
        col_widths2 = [label_col_w] + [data_col_w2] * (n_cols2 - 1)

        add_table(slide, MARGIN, y, table_w, table2, col_widths2, font_size_pt=7)

    # ---- Note ----
    note_y = SLIDE_H - Inches(0.55)
    add_textbox(slide, MARGIN, note_y, SLIDE_W - MARGIN * 2, Inches(0.20),
                f"金額は全て{tax_display}表記　発電量は年▲0.5%低減で試算",
                font_name=FONT_BODY, font_size_pt=7, font_color=C_SUB)

    add_footer(slide)
