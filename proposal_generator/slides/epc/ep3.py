"""
ep3.py - システム概要（EPC）

Shows system specifications including purchase price and investment details.
Uses customer-specific data from the data dict.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_rect, add_rounded_rect,
    add_section_header, add_table, add_textbox,
    fmt_num, fmt_yen,
)

TITLE = "システム概要（EPC）"


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
    add_header_bar(slide, TITLE, logo_path)

    capacity      = data.get("system_capacity_kw")
    panel_kw      = data.get("panel_total_kw")
    selling_price = data.get("selling_price")
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    kw_unit_cost  = data.get("kw_unit_cost", 0) or 0
    panels        = data.get("panels", [])
    pcs_list      = data.get("pcs_list", [])
    batteries     = data.get("batteries") or []
    battery_kwh   = data.get("battery_total_kwh", 0) or 0

    # Net investment = selling_price - subsidy (if subsidy exists)
    net_price = None
    if selling_price is not None:
        net_price = selling_price - subsidy_amount if subsidy_amount else selling_price

    y = CONTENT_TOP + Inches(0.08)

    # ---- KPI cards row ----
    # Show 4 cards when kw_unit_cost is available, otherwise 3
    has_kw_cost = kw_unit_cost and kw_unit_cost > 0
    card_cols = 4 if has_kw_cost else 3
    card_gap = Inches(0.14)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * (card_cols - 1)) / card_cols
    card_h = Inches(1.0)

    kpis = [
        (fmt_num(capacity, 1), "kW", "システム容量"),
        (fmt_yen(selling_price), "", "販売価格（税別）"),
    ]
    # Show net investment if subsidy reduces the price
    if subsidy_amount > 0:
        kpis.append((fmt_yen(net_price), "", "実質投資額"))
    else:
        kpis.append((fmt_yen(selling_price), "", "投資額"))
    if has_kw_cost:
        kpis.append((fmt_num(kw_unit_cost, 0), "円/kW", "kW単価"))

    for i, (number, unit, label) in enumerate(kpis):
        cx = MARGIN + i * (card_w + card_gap)
        add_kpi_card(slide, cx, y, card_w, card_h,
                     number, unit, label,
                     number_size_pt=24)

    y += card_h + Inches(0.2)

    # ---- Panel specs section ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "太陽光パネル仕様")
    y += Inches(0.35)

    col_widths_4 = [Inches(4.2), Inches(1.8), Inches(1.2), Inches(1.5)]
    if panels:
        rows = [["メーカー / 型式", "出力 (W)", "枚数", "合計 (kW)"]]
        for p in panels:
            model = p.get("model", "—")
            watt  = p.get("watt_per_unit", 0)
            count = p.get("count", 0)
            total = watt * count / 1000
            rows.append([model, str(watt), str(count), f"{total:.1f}"])
        add_table(slide, MARGIN, y, SLIDE_W - MARGIN * 2, rows, col_widths_4, font_size_pt=9)
        y += Inches(0.28) * len(rows) + Inches(0.15)
    else:
        add_textbox(slide, MARGIN, y,
                    SLIDE_W - MARGIN * 2, Inches(0.3),
                    f"パネル合計出力：{fmt_num(panel_kw, 1)} kW",
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB)
        y += Inches(0.35)

    # ---- PCS specs section ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "パワーコンディショナ（PCS）仕様")
    y += Inches(0.35)

    col_widths_3 = [Inches(4.7), Inches(2.0), Inches(2.0)]
    if pcs_list:
        rows = [["メーカー / 型式", "容量 (kW)", "台数"]]
        for pcs in pcs_list:
            model    = pcs.get("model", "—")
            cap      = pcs.get("capacity_kw", "—")
            count    = pcs.get("count", 1)
            rows.append([model, str(cap), str(count)])
        add_table(slide, MARGIN, y, SLIDE_W - MARGIN * 2, rows, col_widths_3, font_size_pt=9)
        y += Inches(0.28) * len(rows) + Inches(0.15)
    else:
        add_textbox(slide, MARGIN, y,
                    SLIDE_W - MARGIN * 2, Inches(0.3),
                    "PCS仕様：詳細は別途ご案内",
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB)
        y += Inches(0.35)

    # ---- Battery specs section (if data exists) ----
    if batteries or battery_kwh:
        add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "蓄電池仕様")
        y += Inches(0.35)

        if batteries:
            rows = [["メーカー / 型式", "容量 (kWh)", "台数", "合計 (kWh)"]]
            for bat in batteries:
                model    = bat.get("model", "—")
                kwh      = bat.get("kwh_per_unit", 0)
                count    = bat.get("count", 1)
                total    = bat.get("total_kwh", kwh * count)
                rows.append([model, str(kwh), str(count), f"{total:.1f}"])
            add_table(slide, MARGIN, y, SLIDE_W - MARGIN * 2, rows, col_widths_4, font_size_pt=9)
            y += Inches(0.28) * len(rows) + Inches(0.15)
        else:
            add_textbox(slide, MARGIN, y,
                        SLIDE_W - MARGIN * 2, Inches(0.3),
                        f"蓄電池合計容量：{fmt_num(battery_kwh, 1)} kWh",
                        font_name=FONT_BODY, font_size_pt=10, font_color=C_SUB)
            y += Inches(0.35)

    # ---- Note ----
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.3),
                "※ 上記価格には設計・施工・申請費用を含みます。表示価格は税別です。",
                font_name=FONT_BODY, font_size_pt=8, font_color=C_SUB)

    add_footer(slide)
