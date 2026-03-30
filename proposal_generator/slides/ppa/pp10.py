"""
pp10.py - 補助金活用のご案内スライド

Layout (A4 landscape):
  - Orange header bar with "補助金活用のご案内"
  - Applicable subsidy info from data
  - 3 info cards about common subsidy programs
  - Net investment after subsidy
"""

from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_H, CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE,
    C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect,
    add_section_header, add_textbox, fmt_yen,
)

TITLE = "補助金活用のご案内"

SUBSIDY_CARDS = [
    {
        "name": "環境省補助金",
        "program": "ストレージパリティの達成に\n向けた太陽光発電設備等の\n価格低減促進事業",
        "detail": "太陽光＋蓄電池の導入を支援。\n補助率: 設備費の1/3〜1/2",
    },
    {
        "name": "経産省補助金",
        "program": "需要家主導による太陽光発電\n導入促進補助金",
        "detail": "一定規模以上の自家消費型\n太陽光発電を支援。\n補助単価: 5〜7万円/kW",
    },
    {
        "name": "自治体補助金",
        "program": "各自治体独自の再エネ導入\n支援制度",
        "detail": "都道府県・市区町村ごとに\n異なる補助制度あり。\n国の補助金と併用可能な場合も",
    },
]


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
    add_header_bar(slide, TITLE, logo_path)

    subsidy_name = data.get("subsidy_name", "") or ""
    subsidy_amount = data.get("subsidy_amount", 0) or 0
    selling_price = data.get("selling_price", 0) or 0
    proposal_type = (data.get("proposal_type", "") or "").upper()
    min_ppa_price = data.get("min_ppa_price", 0) or 0
    ppa_principal = data.get("ppa_principal", 0) or 0
    is_ppa = proposal_type == "PPA"

    # EPC: net amount = selling_price - subsidy
    net_amount = max(0, float(selling_price) - float(subsidy_amount))

    y = CONTENT_TOP + Inches(0.05)

    # ---- Applicable subsidy highlight ----
    if subsidy_name:
        add_section_header(slide, MARGIN, y, Inches(6.0), "適用補助金")
        y += Inches(0.4)

        highlight_w = SLIDE_W - MARGIN * 2
        # PPA needs more vertical space for extra info lines
        highlight_h = Inches(1.6) if is_ppa else Inches(0.9)
        add_rounded_rect(slide, MARGIN, y, highlight_w, highlight_h, C_LIGHT_ORANGE)
        add_rect(slide, MARGIN, y, Inches(0.08), highlight_h, C_ORANGE)

        add_textbox(slide,
                    MARGIN + Inches(0.25), y + Inches(0.1),
                    Inches(6.0), Inches(0.3),
                    subsidy_name,
                    font_name=FONT_BLACK, font_size_pt=16, font_color=C_ORANGE, bold=True)

        add_textbox(slide,
                    MARGIN + Inches(0.25), y + Inches(0.48),
                    Inches(4.0), Inches(0.3),
                    f"補助金額: {fmt_yen(subsidy_amount)}",
                    font_name=FONT_BODY, font_size_pt=13, font_color=C_DARK, bold=True)

        if is_ppa:
            # -- PPA mode: subsidy reduces PPA unit price, not customer cost --
            add_textbox(slide,
                        MARGIN + Inches(5.5), y + Inches(0.48),
                        Inches(5.0), Inches(0.3),
                        "PPA単価への補助金効果",
                        font_name=FONT_BLACK, font_size_pt=13, font_color=C_ORANGE, bold=True)

            detail_y = y + Inches(0.85)
            if ppa_principal:
                ppa_provider_burden = max(0, float(ppa_principal) - float(subsidy_amount))
                add_textbox(slide,
                            MARGIN + Inches(0.25), detail_y,
                            Inches(8.0), Inches(0.3),
                            f"設備費 {fmt_yen(ppa_principal)} − 補助金 {fmt_yen(subsidy_amount)}"
                            f" = PPA事業者負担 {fmt_yen(ppa_provider_burden)}",
                            font_name=FONT_BODY, font_size_pt=12, font_color=C_DARK)
                detail_y += Inches(0.3)

            if min_ppa_price:
                add_textbox(slide,
                            MARGIN + Inches(0.25), detail_y,
                            Inches(8.0), Inches(0.3),
                            f"補助金により、PPA単価が低減されます（PPA単価: {min_ppa_price}円/kWh）",
                            font_name=FONT_BODY, font_size_pt=12, font_color=C_DARK)
                detail_y += Inches(0.3)
            else:
                add_textbox(slide,
                            MARGIN + Inches(0.25), detail_y,
                            Inches(8.0), Inches(0.3),
                            "補助金により、PPA単価が低減されます",
                            font_name=FONT_BODY, font_size_pt=12, font_color=C_DARK)
                detail_y += Inches(0.3)

            add_textbox(slide,
                        MARGIN + Inches(0.25), detail_y,
                        Inches(8.0), Inches(0.3),
                        "※ PPA契約のため、お客様の初期費用は0円です",
                        font_name=FONT_BODY, font_size_pt=11, font_color=C_ORANGE, bold=True)
        else:
            # -- EPC mode: show net customer burden --
            if selling_price:
                add_textbox(slide,
                            MARGIN + Inches(5.5), y + Inches(0.48),
                            Inches(5.0), Inches(0.3),
                            f"実質負担額: {fmt_yen(net_amount)}（税別）",
                            font_name=FONT_BODY, font_size_pt=13, font_color=C_DARK, bold=True)

        y += highlight_h + Inches(0.25)

    # ---- Subsidy programs section ----
    add_section_header(slide, MARGIN, y, Inches(6.0), "主な補助金制度")
    y += Inches(0.45)

    card_cols = 3
    gap = Inches(0.18)
    card_w = (SLIDE_W - MARGIN * 2 - gap * (card_cols - 1)) / card_cols
    card_h = Inches(2.8)

    for i, card in enumerate(SUBSIDY_CARDS):
        cx = MARGIN + i * (card_w + gap)

        # Card background
        add_rounded_rect(slide, cx, y, card_w, card_h, C_LIGHT_GRAY)

        # Orange top accent
        add_rect(slide, cx, y, card_w, Inches(0.07), C_ORANGE)

        # Card title
        add_textbox(slide,
                    cx + Inches(0.12), y + Inches(0.18),
                    card_w - Inches(0.24), Inches(0.3),
                    card["name"],
                    font_name=FONT_BLACK, font_size_pt=14, font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Program name
        add_textbox(slide,
                    cx + Inches(0.12), y + Inches(0.55),
                    card_w - Inches(0.24), Inches(1.0),
                    card["program"],
                    font_name=FONT_BODY, font_size_pt=10, font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)

        # Detail
        add_textbox(slide,
                    cx + Inches(0.12), y + Inches(1.6),
                    card_w - Inches(0.24), Inches(1.0),
                    card["detail"],
                    font_name=FONT_BODY, font_size_pt=9, font_color=C_SUB)

    y += card_h + Inches(0.2)

    # ---- Note ----
    add_textbox(slide,
                MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.3),
                "※ 補助金の採択状況・申請時期により変動する場合がございます。詳細はお問い合わせください。",
                font_name=FONT_BODY, font_size_pt=9, font_color=C_SUB)

    add_footer(slide)
