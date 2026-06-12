"""
new_summary.py - まとめ・サマリースライド (design v2: Institutional Trust Grid)

Three-tier closing, eyebrow '06｜まとめ':
  1. Three KPI cards (annual saving / CO2 / zero-upfront or payback) at 28pt
  2. 48pt metric hero: cumulative saving over the contract period (cols 3-8)
  3. Next steps with circle number markers, quiet spec strip,
     and the deck's only dark band: navy CTA with white outline pill
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_SUB, C_WHITE,
    FONT_BLACK, GAP_CARD, MARGIN, SIZE_BODY, SIZE_CAPTION, SLIDE_W,
    add_footer, add_header_bar, add_kpi_card, add_metric_hero,
    add_number_marker, add_pill_label, add_rect, add_section_header,
    add_textbox, fmt_yen, grid_w, grid_x, vstack,
)

TITLE = "まとめ"
EYEBROW = "06｜まとめ"

NEXT_STEPS = [
    "現地調査の実施（屋根荷重・電気設備確認）",
    "補助金申請書類の準備・申請",
    "PPA契約書の確認・締結",
    "設備設計・施工（着工〜運転開始まで約3〜4ヶ月）",
]


def _safe_f(val, fmt: str = ".1f", suffix: str = "") -> str:
    if val is None:
        return "—"
    try:
        return f"{float(val):{fmt}}{suffix}"
    except (ValueError, TypeError):
        return "—"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    company = data.get("company_name") or ""
    is_epc = data.get("proposal_type") == "epc"
    years = data.get("contract_years") or 20
    saving = data.get("annual_cost_saving")
    co2 = data.get("co2_annual_t")
    recovery = data.get("investment_recovery_yr")

    # ---- Block heights for vertical justify ----
    lead_h = Inches(0.28)
    kpi_h = Inches(1.00)
    hero_h = Inches(1.35)
    step_row_h = Inches(0.40)
    steps_h = int(Inches(0.34)) + int(step_row_h) * 2
    spec_h = Inches(0.24)
    cta_h = Inches(0.70)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, kpi_h, hero_h, steps_h, spec_h, cta_h],
                min_gap=GAP_CARD)

    # ---- Lead: company intro ----
    intro = f"{company}　様への提案サマリー" if company else "ご提案サマリー"
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h, intro,
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- Tier 1: three KPI cards (28pt, guarded) ----
    kpi_y = ys[1]
    add_kpi_card(slide, grid_x(0), kpi_y, grid_w(4), kpi_h,
                 fmt_yen(saving, "") if saving else "—", "円/年",
                 "年間電気代削減額")
    add_kpi_card(slide, grid_x(4), kpi_y, grid_w(4), kpi_h,
                 _safe_f(co2, ".1f"), "t-CO₂/年", "年間CO₂削減量")
    if is_epc:
        # Payback only makes sense when the customer invests (EPC)
        add_kpi_card(slide, grid_x(8), kpi_y, grid_w(4), kpi_h,
                     _safe_f(recovery, ".1f"), "年", "投資回収期間")
    else:
        add_kpi_card(slide, grid_x(8), kpi_y, grid_w(4), kpi_h,
                     "0", "円", "初期費用（お客様ご負担）")

    # ---- Tier 2: cumulative saving hero (48pt, cols 3-8) ----
    total = saving * years if saving and years else None
    hero_label = ("20年間削減総額" if is_epc
                  else f"{years}年間削減総額（契約期間合計）")
    add_metric_hero(slide, grid_x(3), ys[2], grid_w(6), hero_h,
                    fmt_yen(total, "") if total else "—", "円", hero_label)

    # ---- Tier 3: next steps with circle markers ----
    steps_y = ys[3]
    add_section_header(slide, MARGIN, steps_y, content_w, "次のステップ")
    list_y = int(steps_y) + int(Inches(0.34))
    marker_d = Inches(0.30)
    col_w = (int(content_w) - int(GAP_CARD)) // 2
    for idx, step in enumerate(NEXT_STEPS):
        col = idx % 2
        row = idx // 2
        sx = int(MARGIN) + col * (col_w + int(GAP_CARD))
        sy = list_y + row * int(step_row_h)
        add_number_marker(slide,
                          sx + int(marker_d) // 2,
                          sy + int(step_row_h) // 2,
                          str(idx + 1), diameter=marker_d)
        add_textbox(slide, sx + int(marker_d) + int(Inches(0.10)), sy,
                    col_w - int(marker_d) - int(Inches(0.10)), step_row_h,
                    step, font_size_pt=11, font_color=C_DARK,
                    anchor=MSO_ANCHOR.MIDDLE)

    # ---- Quiet spec strip (keeps remaining bindings) ----
    if is_epc:
        sp = data.get("selling_price")
        irr = data.get("irr")
        irr_txt = _safe_f(
            irr * 100 if isinstance(irr, (int, float)) and irr else irr,
            ".1f", "%")
        spec = (f"販売価格：{fmt_yen(sp) if sp else '—'}　｜　"
                f"IRR（投資利回り）：{irr_txt}　｜　算定期間：{years}年")
    else:
        unit_price = data.get("ppa_unit_price")
        lease = data.get("annual_lease_payment")
        spec = (f"PPA単価：{_safe_f(unit_price, '.1f')}円/kWh　｜　"
                f"年間リース料：{fmt_yen(lease) if lease else '—'}　｜　"
                f"契約期間：{years}年")
    add_textbox(slide, MARGIN, ys[4], content_w, spec_h, spec,
                font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                align=PP_ALIGN.CENTER)

    # ---- CTA band (the deck's only dark band) ----
    cta_y = ys[5]
    add_rect(slide, MARGIN, cta_y, content_w, cta_h, C_NAVY)
    add_textbox(slide, int(MARGIN) + int(Inches(0.35)), cta_y,
                int(content_w) - int(Inches(2.50)), cta_h,
                "まずは現地調査から——日程のご相談を承ります",
                font_name=FONT_BLACK, font_size_pt=13, font_color=C_WHITE,
                bold=True, anchor=MSO_ANCHOR.MIDDLE)
    pill_w = Inches(1.45)
    pill_h = Inches(0.34)
    add_pill_label(slide,
                   int(MARGIN) + int(content_w) - int(pill_w)
                   - int(Inches(0.35)),
                   int(cta_y) + (int(cta_h) - int(pill_h)) // 2,
                   pill_w, pill_h, "現地調査無料",
                   font_color=C_WHITE)

    add_footer(slide)
