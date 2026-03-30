"""
new_summary.py - まとめ・サマリースライド

Shows key KPIs from the proposal at a glance:
- Annual cost saving
- 20-year total saving
- CO2 reduction
- IRR / payback period
- Action items / next steps
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE, C_LIGHT_GRAY,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, fmt_yen, fmt_num,
)

TITLE = "まとめ"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)
    company = data.get("company_name", "") or ""

    # Company name intro
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                f"{company}　様への提案サマリー",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_SUB)
    y += Inches(0.32)

    # ---- KPI cards (2x3 grid) ----
    kpis = _build_kpis(data)
    card_cols = 3
    card_w = (SLIDE_W - MARGIN * 2 - Inches(0.18) * (card_cols - 1)) / card_cols
    card_h = Inches(1.25)

    for i, (number, unit, label) in enumerate(kpis):
        col = i % card_cols
        row = i // card_cols
        cx = MARGIN + col * (card_w + Inches(0.18))
        cy = y + row * (card_h + Inches(0.1))

        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE)
        add_rect(slide, cx, cy, card_w, Inches(0.06), C_ORANGE)
        add_textbox(slide, cx, cy + Inches(0.08), card_w, Inches(0.48),
                    number,
                    font_name=FONT_BLACK, font_size_pt=26,
                    font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, cx, cy + Inches(0.54), card_w, Inches(0.2),
                    unit,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, cx + Inches(0.08), cy + card_h - Inches(0.3),
                    card_w - Inches(0.16), Inches(0.26),
                    label,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)

    y += (len(kpis) // card_cols + (1 if len(kpis) % card_cols else 0)) * (card_h + Inches(0.1)) + Inches(0.1)

    # ---- Next steps (two-column layout for landscape) ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2, "次のステップ", font_size_pt=12)
    y += Inches(0.32)

    steps = [
        "①  現地調査の実施（屋根荷重・電気設備確認）",
        "②  補助金申請書類の準備・申請",
        "③  PPA契約書の確認・締結",
        "④  設備設計・施工（着工〜運転開始まで約3〜4ヶ月）",
    ]
    step_col_w = (SLIDE_W - MARGIN * 2 - Inches(0.2)) / 2
    for idx, step in enumerate(steps):
        col = idx % 2
        row = idx // 2
        sx = MARGIN + Inches(0.1) + col * (step_col_w + Inches(0.2))
        sy = y + row * Inches(0.28)
        add_textbox(slide, sx, sy,
                    step_col_w - Inches(0.1), Inches(0.26),
                    step,
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_DARK)

    add_footer(slide)


def _build_kpis(data: dict) -> list[tuple[str, str, str]]:
    is_epc = data.get("proposal_type") == "epc"
    years = data.get("contract_years", 20)
    saving = data.get("annual_cost_saving")
    co2 = data.get("co2_annual_t")

    def _safe_f(val, fmt=".1f", suffix=""):
        if val is None:
            return "—"
        try:
            return f"{float(val):{fmt}}{suffix}"
        except (ValueError, TypeError):
            return "—"

    kpis = [
        (fmt_yen(saving, "") if saving else "—", "円/年", "年間電気代削減額"),
        (fmt_yen(saving * years if saving and years else None, "") if saving else "—",
         "円（契約期間合計）" if not is_epc else "円（20年間）",
         f"{years or 20}年間削減総額" if not is_epc else "20年間削減総額"),
        (_safe_f(co2, ".1f"), "t-CO₂/年", "年間CO₂削減量"),
    ]

    if is_epc:
        # EPC: show selling price, payback, IRR
        sp = data.get("selling_price")
        recovery = data.get("investment_recovery_yr")
        irr = data.get("irr")
        kpis.append((fmt_yen(sp, "") if sp else "—", "円", "販売価格"))
        kpis.append((_safe_f(recovery, ".1f"), "年", "投資回収期間"))
        kpis.append((_safe_f(
            irr * 100 if isinstance(irr, (int, float)) and irr else irr, ".1f", "%"),
            "IRR", "投資利回り"))
    else:
        # PPA: show PPA unit price, lease payment, DSCR
        unit_price = data.get("ppa_unit_price")
        lease = data.get("annual_lease_payment")
        recovery = data.get("investment_recovery_yr")
        kpis.append((_safe_f(unit_price, ".1f"), "円/kWh", "PPA単価"))
        kpis.append((_safe_f(recovery, ".1f"), "年", "投資回収期間"))
        kpis.append((fmt_yen(lease, "") if lease else "—", "円/年", "年間リース料"))

    return kpis
