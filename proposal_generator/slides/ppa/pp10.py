"""
pp10.py - 補助金活用のご案内 (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - v2 white header, eyebrow '04｜ご契約条件'
  - Applied subsidy block: section header + program name + inline metric
    band (28pt number pairs, hairline separators)
      PPA: 補助金額 / 初期費用0円 / 補助金適用後PPA単価 + provider-burden
           calc caption (NO payback figures on PPA decks)
      EPC: 補助金額 / 実質負担額
  - 3 program cards (accent-top white cards)
  - 8pt note
"""

from __future__ import annotations

from pathlib import Path

from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_HAIR, C_NAVY, C_SUB,
    FONT_BLACK, FONT_BODY, LINE_SPACING_BODY, MARGIN,
    SIZE_CAPTION, SIZE_H2, SIZE_LEAD, SIZE_SMALL, SLIDE_W,
    add_card_with_accent, add_footer, add_header_bar, add_line,
    add_number_unit, add_section_header, add_textbox, fmt_yen,
    grid_w, grid_x, vstack,
)

TITLE = "補助金活用のご案内"
EYEBROW = "04｜ご契約条件"

SUBSIDY_CARDS = [
    {
        "name": "環境省補助金",
        "program": "ストレージパリティの達成に向けた太陽光発電設備等の価格低減促進事業",
        "detail": "太陽光＋蓄電池の導入を支援。補助率は設備費の1/3〜1/2。",
    },
    {
        "name": "経産省補助金",
        "program": "需要家主導による太陽光発電導入促進補助金",
        "detail": "一定規模以上の自家消費型太陽光発電を支援。補助単価は5〜7万円/kW。",
    },
    {
        "name": "自治体補助金",
        "program": "各自治体独自の再エネ導入支援制度",
        "detail": "都道府県・市区町村ごとに異なる補助制度あり。国の補助金と併用可能な場合も。",
    },
]

NOTE = "※ 補助金の採択状況・申請時期により変動する場合がございます。詳細はお問い合わせください。"


def _yen_parts(v: float) -> tuple[str, str]:
    """Split a yen amount into (number, unit) for add_number_unit."""
    if v >= 1_0000_0000:
        return f"{v / 1_0000_0000:.2f}", "億円"
    if v >= 10_000:
        return f"{v / 10_000:,.0f}", "万円"
    return f"{v:,.0f}", "円"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP10 (subsidy utilization) onto an already-added blank slide.

    data keys used:
        subsidy_name, subsidy_amount, system_capacity_kw,
        selling_price (total system price before subsidy),
        proposal_type ('PPA' or 'EPC'),
        min_ppa_price (PPA unit price after subsidy),
        ppa_principal (PPA provider's equipment cost)
    """
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    subsidy_name = data.get("subsidy_name", "") or ""
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    selling_price = data.get("selling_price", 0) or 0
    proposal_type = (data.get("proposal_type", "") or "").upper()
    min_ppa_price = data.get("min_ppa_price", 0) or 0
    ppa_principal = data.get("ppa_principal", 0) or 0
    is_ppa = proposal_type == "PPA"

    # EPC: net customer burden = selling_price - subsidy
    net_amount = max(0, float(selling_price) - float(subsidy_amount))

    # ---- Metric band items + caption (PPA / EPC branching) ----
    band = []
    cap_lines = []
    if subsidy_name:
        band.append(("補助金額",
                     *(_yen_parts(float(subsidy_amount))
                       if subsidy_amount else ("—", ""))))
        if is_ppa:
            # Subsidy reduces the PPA unit price, not the customer cost.
            band.append(("初期費用（お客様ご負担）", "0", "円"))
            if min_ppa_price:
                band.append(("補助金適用後PPA単価",
                             f"{min_ppa_price}", "円/kWh"))
            if ppa_principal:
                burden = max(0, float(ppa_principal) - float(subsidy_amount))
                cap_lines.append(
                    f"設備費 {fmt_yen(ppa_principal)} − 補助金 "
                    f"{fmt_yen(subsidy_amount)} ＝ PPA事業者負担 "
                    f"{fmt_yen(burden)}")
            cap_lines.append(
                "補助金により、PPA単価が低減されます。"
                "PPA契約のため、お客様の初期費用は0円です。")
        else:
            if selling_price:
                band.append(("実質負担額（税別）", *_yen_parts(net_amount)))

    # ---- Block heights for vertical justify ----
    cap_h = Inches(0.20) * len(cap_lines)
    blockA_h = (int(Inches(0.36)) + int(Inches(0.34))
                + int(Inches(0.70)) + int(cap_h)) if subsidy_name else None
    cards_h = Inches(1.95)
    blockB_h = int(Inches(0.40)) + int(cards_h)
    note_h = Inches(0.22)

    blocks = ([blockA_h] if blockA_h else []) + [blockB_h, note_h]
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, blocks)
    yi = 0

    # ---- Block A: applied subsidy (name + metric band + caption) ----
    if subsidy_name:
        ay = ys[yi]
        yi += 1
        add_section_header(slide, MARGIN, ay, content_w, "適用補助金")
        name_y = int(ay) + int(Inches(0.36))
        add_textbox(slide, MARGIN, name_y, content_w, Inches(0.30),
                    subsidy_name,
                    font_name=FONT_BLACK, font_size_pt=SIZE_H2,
                    font_color=C_NAVY, bold=True)

        band_y = name_y + int(Inches(0.34))
        band_h = Inches(0.70)
        span = max(12 // max(len(band), 1), 3)
        for i, (label, number, unit) in enumerate(band):
            bx = grid_x(i * span)
            bw = grid_w(span) - Inches(0.20)
            add_textbox(slide, bx, band_y, bw, Inches(0.20),
                        label,
                        font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                        bold=True)
            add_number_unit(slide, bx, band_y + int(Inches(0.22)),
                            bw, Inches(0.44),
                            number, unit)
            if i > 0:
                sep_x = bx - Inches(0.12)
                add_line(slide, sep_x, band_y + int(Inches(0.04)),
                         sep_x, band_y + int(band_h) - int(Inches(0.06)),
                         C_HAIR, width_pt=0.5)

        cap_y = band_y + int(band_h)
        for j, line in enumerate(cap_lines):
            add_textbox(slide, MARGIN, cap_y + j * int(Inches(0.20)),
                        content_w, Inches(0.18),
                        line,
                        font_size_pt=SIZE_CAPTION, font_color=C_SUB)

    # ---- Block B: subsidy program cards ----
    by = ys[yi]
    yi += 1
    add_section_header(slide, MARGIN, by, content_w, "主な補助金制度")
    cards_y = int(by) + int(Inches(0.40))

    for i, card in enumerate(SUBSIDY_CARDS):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, cards_y, w, cards_h)

        add_textbox(slide, cx, int(cy) + int(Inches(0.02)),
                    cw, Inches(0.26),
                    card["name"],
                    font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                    font_color=C_DARK, bold=True)
        add_textbox(slide, cx, int(cy) + int(Inches(0.34)),
                    cw, Inches(0.55),
                    card["program"],
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_NAVY, bold=True,
                    line_spacing=LINE_SPACING_BODY)
        add_textbox(slide, cx, int(cy) + int(Inches(0.95)),
                    cw, int(ch) - int(Inches(0.95)),
                    card["detail"],
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, line_spacing=LINE_SPACING_BODY)

    # ---- Note ----
    add_textbox(slide, MARGIN, ys[yi], content_w, note_h,
                NOTE,
                font_size_pt=SIZE_SMALL, font_color=C_SUB)

    add_footer(slide)
