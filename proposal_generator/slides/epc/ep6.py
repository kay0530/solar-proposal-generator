"""
ep6.py - 補助金活用（EPC）

Shows available subsidies for EPC solar system purchase.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_rect, add_rounded_rect,
    add_section_header, add_textbox,
    fmt_yen,
)

TITLE = "補助金活用（EPC）"

SUBSIDY_PROGRAMS = [
    {
        "name": "需要家主導型太陽光発電導入促進事業",
        "body": "環境省",
        "rate": "設備費の1/3〜1/2",
        "note": "自家消費率50%以上が条件。蓄電池併設で補助率UP。",
    },
    {
        "name": "中小企業経営強化税制",
        "body": "経済産業省",
        "rate": "即時償却 or 税額控除10%",
        "note": "中小企業が対象。設備取得価額の全額を初年度に費用計上可能。",
    },
    {
        "name": "ストレージパリティ達成に向けた太陽光発電設備導入支援事業",
        "body": "環境省",
        "rate": "定額補助（4〜5万円/kW）",
        "note": "蓄電池の同時導入が必須。自家消費型に限る。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP6 (subsidy utilization for EPC) onto an already-added blank slide.

    data keys used:
        selling_price, subsidy_name, subsidy_amount, system_capacity_kw,
        investment_recovery_yr
    """
    add_header_bar(slide, TITLE, logo_path)

    selling_price  = data.get("selling_price")
    subsidy_name   = data.get("subsidy_name", "") or ""
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    capacity       = data.get("system_capacity_kw")
    recovery_yr    = data.get("investment_recovery_yr")

    y = CONTENT_TOP + Inches(0.08)

    # ---- KPI cards: investment vs subsidy ----
    has_recovery = recovery_yr is not None and recovery_yr > 0
    card_cols = 4 if has_recovery else 3
    card_gap = Inches(0.14)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * (card_cols - 1)) / card_cols
    card_h = Inches(1.0)

    net_cost = None
    if selling_price is not None:
        net_cost = selling_price - subsidy_amount

    kpis = [
        (fmt_yen(selling_price), "", "設備投資額"),
        (fmt_yen(subsidy_amount), "", "補助金額"),
        (fmt_yen(net_cost), "", "実質負担額"),
    ]
    if has_recovery:
        kpis.append((f"{recovery_yr:.1f}", "年", "投資回収年数"))

    for i, (number, unit, label) in enumerate(kpis):
        cx = MARGIN + i * (card_w + card_gap)
        add_kpi_card(slide, cx, y, card_w, card_h,
                     number, unit, label,
                     number_size_pt=24)

    y += card_h + Inches(0.22)

    # ---- Applied subsidy highlight (conditional) ----
    has_highlight = bool(subsidy_name)
    if has_highlight:
        add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "適用予定補助金")
        y += Inches(0.32)
        highlight_h = Inches(0.48)
        add_rounded_rect(slide, MARGIN, y, SLIDE_W - MARGIN * 2, highlight_h, C_LIGHT_ORANGE)
        add_rect(slide, MARGIN, y, Inches(0.08), highlight_h, C_ORANGE)
        add_textbox(slide,
                    MARGIN + Inches(0.18), y + Inches(0.10),
                    SLIDE_W - MARGIN * 2 - Inches(0.25), highlight_h - Inches(0.16),
                    f"{subsidy_name}　→　補助金額：{fmt_yen(subsidy_amount)}",
                    font_name=FONT_BODY, font_size_pt=12,
                    font_color=C_DARK, bold=True)
        y += highlight_h + Inches(0.18)

    # ---- Available subsidy programs ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "主な補助金制度一覧")
    y += Inches(0.32)

    # Limit cards to avoid overflow: 2 if highlight is shown, 3 otherwise
    max_programs = 2 if has_highlight else 3
    programs_to_show = SUBSIDY_PROGRAMS[:max_programs]

    card_h_item = Inches(0.78)
    for prog in programs_to_show:
        add_rounded_rect(slide, MARGIN, y, SLIDE_W - MARGIN * 2, card_h_item, C_LIGHT_GRAY)
        add_rect(slide, MARGIN, y, Inches(0.06), card_h_item, C_ORANGE)

        # Program name
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.04),
                    SLIDE_W - MARGIN * 2 - Inches(0.2), Inches(0.22),
                    f"{prog['name']}（{prog['body']}）",
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK, bold=True)
        # Rate
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.28),
                    Inches(3.5), Inches(0.20),
                    f"補助率：{prog['rate']}",
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_ORANGE, bold=True)
        # Note
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.50),
                    SLIDE_W - MARGIN * 2 - Inches(0.2), Inches(0.24),
                    prog["note"],
                    font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_SUB)
        y += card_h_item + Inches(0.08)

    # ---- Note ----
    add_textbox(slide, MARGIN, y + Inches(0.02),
                SLIDE_W - MARGIN * 2, Inches(0.22),
                "※ 補助金の採択は申請内容・予算状況により異なります。詳細はお問い合わせください。",
                font_name=FONT_BODY, font_size_pt=8,
                font_color=C_SUB)

    add_footer(slide)
