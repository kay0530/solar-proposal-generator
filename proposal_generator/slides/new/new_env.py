"""
new_env.py - 環境への貢献スライド (design v2: Institutional Trust Grid)

CO2 reduction equivalents as three 28pt KPI cards with flat leaf icons
(CO2削減量 / 杉の木換算 / ガソリン換算), then SDGs alignment cards
(Goal 7, 13) and an environmental-certification panel — all inside the
v2 palette (no SDG brand colors, no emoji glyph icons).
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CARD_PAD, CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_HAIR, C_NAVY, C_PANEL,
    C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, MARGIN,
    SIZE_CAPTION, SIZE_BODY, SIZE_SMALL, SLIDE_W,
    add_footer, add_header_bar, add_icon, add_kpi_card,
    add_multiline_textbox, add_rect, add_rounded_rect, add_section_header,
    add_textbox, fmt_num, grid_w, grid_x, vstack,
)

TITLE = "環境への貢献"
EYEBROW = "補足｜環境価値"

# Conversion factors (approximate)
SUGI_PER_TON_CO2 = 71.4     # cedar trees per ton CO2 absorbed/year
GASOLINE_PER_TON_CO2 = 430  # liters of gasoline per ton CO2


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    company = data.get("company_name", "") or ""

    co2_annual = data.get("co2_annual_t")
    try:
        co2_val = float(co2_annual) if co2_annual is not None else 85.0
    except (TypeError, ValueError):
        co2_val = 85.0

    sugi_count = co2_val * SUGI_PER_TON_CO2
    gasoline_liters = co2_val * GASOLINE_PER_TON_CO2

    # ---- Vertical layout ----
    lead_h = Inches(0.26)
    kpi_h = Inches(1.30)
    detail_h = Inches(0.50)
    kpi_block_h = kpi_h + Inches(0.08) + detail_h
    sdg_card_h = Inches(1.60)
    sdg_block_h = Inches(0.40) + sdg_card_h
    note_h = Inches(0.20)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, kpi_block_h, sdg_block_h, note_h])

    # ---- Lead ----
    lead = (f"{company}　様の太陽光導入による環境貢献効果"
            if company else "太陽光発電導入による環境貢献効果")
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h, lead,
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- CO2 equivalent KPI cards (28pt) + leaf icons + detail captions ----
    kpi_y = ys[1]
    cards = [
        (fmt_num(co2_val, 1), "t-CO₂/年", "年間CO₂削減量",
         "再生可能エネルギー発電による温室効果ガス排出削減効果"),
        (f"{sugi_count:,.0f}", "本相当", "杉の木換算",
         "杉の木が1年間に吸収するCO₂量に換算した本数"),
        (f"{gasoline_liters:,.0f}", "L相当", "ガソリン換算",
         "ガソリン燃焼で排出されるCO₂量に換算したリットル数"),
    ]
    for i, (number, unit, label, detail) in enumerate(cards):
        x = grid_x(i * 4)
        w = grid_w(4)
        add_kpi_card(slide, x, kpi_y, w, kpi_h, number, unit, label)
        add_icon(slide, "leaf", x + w - Inches(0.55), kpi_y + Inches(0.14),
                 size=Inches(0.36))
        add_textbox(slide, x + Inches(0.02), kpi_y + kpi_h + Inches(0.08),
                    w - Inches(0.04), detail_h, detail,
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                    word_wrap=True, line_spacing=1.3)

    # ---- SDGs alignment + environmental certifications ----
    sdg_y = ys[2]
    add_section_header(slide, MARGIN, sdg_y, content_w, "SDGsへの貢献")
    card_y = sdg_y + Inches(0.40)

    sdg_items = [
        ("SDG 7", "エネルギーをみんなに そしてクリーンに",
         "再生可能エネルギーの導入拡大に貢献"),
        ("SDG 13", "気候変動に具体的な対策を",
         "CO₂排出削減で気候変動対策に貢献"),
    ]
    for i, (tag, sdg_title, sdg_desc) in enumerate(sdg_items):
        x = grid_x(i * 4)
        w = grid_w(4)
        add_rounded_rect(slide, x, card_y, w, sdg_card_h, C_WHITE,
                         border_color=C_HAIR, border_pt=0.75)
        add_multiline_textbox(
            slide, x + CARD_PAD, card_y + CARD_PAD,
            w - CARD_PAD * 2, sdg_card_h - CARD_PAD * 2,
            [
                (tag, FONT_BLACK, SIZE_CAPTION, C_NAVY, True, PP_ALIGN.LEFT),
                (sdg_title, FONT_BODY, 10.5, C_DARK, True, PP_ALIGN.LEFT),
                (sdg_desc, FONT_BODY, SIZE_CAPTION, C_SUB, False,
                 PP_ALIGN.LEFT),
            ],
            line_spacing=1.5)

    # Certification panel (C_PANEL + navy left bar)
    cert_x = grid_x(8)
    cert_w = grid_w(4)
    add_rect(slide, cert_x, card_y, cert_w, sdg_card_h, C_PANEL)
    add_rect(slide, cert_x, card_y, Inches(0.045), sdg_card_h, C_NAVY)
    add_multiline_textbox(
        slide, cert_x + Inches(0.045) + CARD_PAD, card_y + CARD_PAD,
        cert_w - Inches(0.045) - CARD_PAD * 2, sdg_card_h - CARD_PAD * 2,
        [
            ("環境認証・制度への活用", FONT_BLACK, 10.5, C_DARK, True,
             PP_ALIGN.LEFT),
            ("RE100 / SBT / CDP", FONT_BODY, SIZE_CAPTION, C_SUB, False,
             PP_ALIGN.LEFT),
            ("グリーン電力証書・非化石証書", FONT_BODY, SIZE_CAPTION, C_SUB,
             False, PP_ALIGN.LEFT),
        ],
        line_spacing=1.5)

    # ---- Conversion-factor note (8pt) ----
    add_textbox(slide, MARGIN, ys[3], content_w, note_h,
                "※ 換算係数：杉の木 約71.4本/t-CO₂（年間吸収量）、"
                "ガソリン 約430L/t-CO₂",
                font_size_pt=SIZE_SMALL, font_color=C_SUB)

    add_footer(slide)
