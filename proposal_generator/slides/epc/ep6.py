"""
ep6.py - 補助金活用（EPC） — design system v2

Shows available subsidies for EPC solar system purchase.

Layout (A4 landscape):
  - KPI card row: 設備投資額 / 補助金額 / 実質負担額 (+投資回収年数 if present)
  - Applied subsidy highlight band (C_PANEL + orange left bar + 28pt amount)
  - All subsidy programs as full-width accent cards (no truncation; the
    vstack layout absorbs the conditional highlight block)
  - 8pt disclaimer note
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_ORANGE, C_PANEL, C_SUB,
    FONT_BLACK, FONT_BODY, GAP_IN_CARD, MARGIN, SIZE_BODY, SIZE_CAPTION,
    SIZE_SMALL, SLIDE_W,
    add_card_with_accent, add_footer, add_header_bar, add_kpi_card,
    add_multiline_textbox, add_number_unit, add_rect, add_section_header,
    add_textbox, vstack,
)

TITLE = "補助金活用（EPC）"
EYEBROW = "04｜補助金活用"

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


def _yen_parts(v) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for add_number_unit / KPI."""
    if v is None:
        return "—", ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v), ""
    if abs(f) >= 1_0000_0000:
        return f"{f / 1_0000_0000:.2f}", "億円"
    if abs(f) >= 10_000:
        return f"{f / 10_000:,.0f}", "万円"
    return f"{f:,.0f}", "円"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP6 (subsidy utilization for EPC) onto an already-added blank slide.

    data keys used:
        selling_price, subsidy_name, subsidy_amount, system_capacity_kw,
        investment_recovery_yr
    """
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    selling_price  = data.get("selling_price")
    subsidy_name   = data.get("subsidy_name", "") or ""
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    capacity       = data.get("system_capacity_kw")  # noqa: F841 (kept binding)
    recovery_yr    = data.get("investment_recovery_yr")

    try:
        rec_f = float(recovery_yr) if recovery_yr is not None else None
    except (TypeError, ValueError):
        rec_f = None
    has_recovery = rec_f is not None and rec_f > 0

    net_cost = None
    if selling_price is not None:
        try:
            net_cost = float(selling_price) - float(subsidy_amount)
        except (TypeError, ValueError):
            net_cost = None

    total_w = SLIDE_W - MARGIN * 2

    # ---- Block heights -> vstack (kills bottom dead space) ----
    kpi_h = Inches(1.05)
    hl_head_h = Inches(0.36)
    hl_band_h = Inches(0.62)
    sect_h = Inches(0.40)
    prog_card_h = Inches(0.88)
    prog_gap = Inches(0.12)
    n_progs = len(SUBSIDY_PROGRAMS)
    list_h = (int(sect_h) + int(prog_card_h) * n_progs
              + int(prog_gap) * (n_progs - 1))
    note_h = Inches(0.22)

    has_highlight = bool(subsidy_name)
    blocks = [kpi_h]
    if has_highlight:
        blocks.append(int(hl_head_h) + int(hl_band_h))
    blocks += [list_h, note_h]
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, blocks)

    # ---- KPI cards: investment vs subsidy (28pt via add_kpi_card) ----
    card_cols = 4 if has_recovery else 3
    card_gap = GAP_IN_CARD
    card_w = (int(total_w) - int(card_gap) * (card_cols - 1)) // card_cols

    kpis = [
        (*_yen_parts(selling_price), "設備投資額"),
        (*_yen_parts(subsidy_amount), "補助金額"),
        (*_yen_parts(net_cost), "実質負担額"),
    ]
    if has_recovery:
        kpis.append((f"{rec_f:.1f}", "年", "投資回収年数"))

    for i, (number, unit, label) in enumerate(kpis):
        cx = int(MARGIN) + i * (card_w + int(card_gap))
        add_kpi_card(slide, cx, ys[0], card_w, kpi_h, number, unit, label)

    blk = 1

    # ---- Applied subsidy highlight (conditional) ----
    if has_highlight:
        hy = ys[blk]
        blk += 1
        add_section_header(slide, MARGIN, hy, total_w, "適用予定補助金")
        band_y = int(hy) + int(hl_head_h)
        add_rect(slide, MARGIN, band_y, total_w, hl_band_h, C_PANEL)
        add_rect(slide, MARGIN, band_y, Inches(0.05), hl_band_h, C_ORANGE)

        name_w = int(total_w) * 55 // 100
        add_textbox(slide, int(MARGIN) + int(Inches(0.18)), band_y,
                    name_w, hl_band_h,
                    subsidy_name,
                    font_name=FONT_BLACK, font_size_pt=SIZE_BODY,
                    font_color=C_DARK, bold=True, anchor=MSO_ANCHOR.MIDDLE)

        amt_num, amt_unit = _yen_parts(subsidy_amount)
        amt_x = int(MARGIN) + name_w + int(Inches(0.20))
        amt_w = int(total_w) - name_w - int(Inches(0.38))
        add_textbox(slide, amt_x, band_y + int(Inches(0.06)),
                    amt_w, Inches(0.16),
                    "補助金額",
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, bold=True, align=PP_ALIGN.RIGHT)
        add_number_unit(slide, amt_x, band_y + int(Inches(0.20)),
                        amt_w, int(hl_band_h) - int(Inches(0.24)),
                        amt_num, amt_unit, align=PP_ALIGN.RIGHT)

    # ---- Available subsidy programs (all shown, accent-left cards) ----
    ly = ys[blk]
    add_section_header(slide, MARGIN, ly, total_w, "主な補助金制度一覧")
    py = int(ly) + int(sect_h)
    for prog in SUBSIDY_PROGRAMS:
        cx, cy, cw, ch = add_card_with_accent(slide, MARGIN, py, total_w,
                                              prog_card_h,
                                              accent_position="left")
        lines = [
            (f"{prog['name']}（{prog['body']}）",
             FONT_BODY, SIZE_BODY, C_DARK, True, PP_ALIGN.LEFT),
            (f"補助率：{prog['rate']}",
             FONT_BODY, SIZE_CAPTION, C_NAVY, True, PP_ALIGN.LEFT),
            (prog["note"],
             FONT_BODY, SIZE_CAPTION, C_SUB, False, PP_ALIGN.LEFT),
        ]
        add_multiline_textbox(slide, cx, cy, cw, ch, lines, line_spacing=1.35)
        py += int(prog_card_h) + int(prog_gap)

    # ---- Note (8pt) ----
    add_textbox(slide, MARGIN, ys[-1], total_w, note_h,
                "※ 補助金の採択は申請内容・予算状況により異なります。詳細はお問い合わせください。",
                font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                font_color=C_SUB)

    add_footer(slide)
