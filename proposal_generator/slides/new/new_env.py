"""
new_env.py - 環境への貢献スライド

CO2 reduction equivalent visualization.
Cards: CO2削減量(t), 杉の木換算(本), ガソリン換算(L).
SDGs alignment badges (Goal 7, 13).
Environmental certifications mention.
"""
from __future__ import annotations
from pathlib import Path
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, fmt_num,
)

TITLE = "環境への貢献"

# SDGs color palette
C_SDG7 = RGBColor(0xFC, 0xC3, 0x0B)   # SDG7 Affordable and Clean Energy (yellow)
C_SDG13 = RGBColor(0x3F, 0x7E, 0x44)  # SDG13 Climate Action (green)
C_GREEN = RGBColor(0x2E, 0x7D, 0x32)  # General green accent

# Conversion factors (approximate)
SUGI_PER_TON_CO2 = 71.4    # cedar trees per ton CO2 absorbed/year
GASOLINE_PER_TON_CO2 = 430  # liters of gasoline per ton CO2


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Subtitle
    company = data.get("company_name", "") or ""
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                f"{company}　様の太陽光導入による環境貢献効果",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_SUB)
    y += Inches(0.35)

    # ---- CO2 equivalent cards (3 cards) ----
    co2_annual = data.get("co2_annual_t")
    try:
        co2_val = float(co2_annual) if co2_annual is not None else 85.0
    except (TypeError, ValueError):
        co2_val = 85.0

    sugi_count = co2_val * SUGI_PER_TON_CO2
    gasoline_liters = co2_val * GASOLINE_PER_TON_CO2

    cards = [
        {
            "icon": "🌿",
            "number": fmt_num(co2_val, 1),
            "unit": "t-CO₂/年",
            "label": "年間CO₂削減量",
            "detail": "再生可能エネルギー発電による\n温室効果ガス排出削減効果",
        },
        {
            "icon": "🌲",
            "number": f"{sugi_count:,.0f}",
            "unit": "本相当",
            "label": "杉の木換算",
            "detail": "杉の木が1年間に吸収する\nCO₂量に換算した本数",
        },
        {
            "icon": "⛽",
            "number": f"{gasoline_liters:,.0f}",
            "unit": "L相当",
            "label": "ガソリン換算",
            "detail": "ガソリン燃焼で排出される\nCO₂量に換算したリットル数",
        },
    ]

    card_cols = 3
    card_gap = Inches(0.2)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap * (card_cols - 1)) / card_cols
    card_h = Inches(2.5)

    for i, card in enumerate(cards):
        cx = MARGIN + i * (card_w + card_gap)

        # Card background
        add_rounded_rect(slide, cx, y, card_w, card_h, C_LIGHT_ORANGE, radius_pt=6.0)

        # Green top accent bar
        add_rect(slide, cx, y, card_w, Inches(0.06), C_GREEN)

        # Icon
        add_textbox(slide, cx, y + Inches(0.12),
                    card_w, Inches(0.4),
                    card["icon"],
                    font_name=FONT_BODY, font_size_pt=28,
                    font_color=C_DARK,
                    align=PP_ALIGN.CENTER)

        # Number (large)
        add_textbox(slide, cx, y + Inches(0.55),
                    card_w, Inches(0.5),
                    card["number"],
                    font_name=FONT_BLACK, font_size_pt=30,
                    font_color=C_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Unit
        add_textbox(slide, cx, y + Inches(1.05),
                    card_w, Inches(0.22),
                    card["unit"],
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB,
                    align=PP_ALIGN.CENTER)

        # Label
        add_textbox(slide, cx + Inches(0.08), y + Inches(1.3),
                    card_w - Inches(0.16), Inches(0.26),
                    card["label"],
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_DARK, bold=True,
                    align=PP_ALIGN.CENTER)

        # Divider
        add_rect(slide, cx + Inches(0.1), y + Inches(1.58),
                 card_w - Inches(0.2), Inches(0.02), C_GREEN)

        # Detail
        add_textbox(slide, cx + Inches(0.1), y + Inches(1.65),
                    card_w - Inches(0.2), Inches(0.7),
                    card["detail"],
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB,
                    word_wrap=True)

    y += card_h + Inches(0.2)

    # ---- SDGs alignment section ----
    add_section_header(slide, MARGIN, y, SLIDE_W - MARGIN * 2,
                       "SDGs への貢献", font_size_pt=12)
    y += Inches(0.3)

    # SDG badges side by side
    badge_w = Inches(2.8)
    badge_h = Inches(0.9)
    badge_gap = Inches(0.3)
    badges_total_w = badge_w * 2 + badge_gap
    badge_start_x = MARGIN

    # SDG 7
    add_rounded_rect(slide, badge_start_x, y, badge_w, badge_h, C_SDG7, radius_pt=6.0)
    add_textbox(slide, badge_start_x + Inches(0.1), y + Inches(0.08),
                badge_w - Inches(0.2), Inches(0.3),
                "SDG 7: エネルギーをみんなに そしてクリーンに",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_WHITE, bold=True)
    add_textbox(slide, badge_start_x + Inches(0.1), y + Inches(0.42),
                badge_w - Inches(0.2), Inches(0.4),
                "再生可能エネルギーの導入拡大に貢献",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_WHITE)

    # SDG 13
    add_rounded_rect(slide, badge_start_x + badge_w + badge_gap, y,
                     badge_w, badge_h, C_SDG13, radius_pt=6.0)
    add_textbox(slide, badge_start_x + badge_w + badge_gap + Inches(0.1), y + Inches(0.08),
                badge_w - Inches(0.2), Inches(0.3),
                "SDG 13: 気候変動に具体的な対策を",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_WHITE, bold=True)
    add_textbox(slide, badge_start_x + badge_w + badge_gap + Inches(0.1), y + Inches(0.42),
                badge_w - Inches(0.2), Inches(0.4),
                "CO₂排出削減で気候変動対策に貢献",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_WHITE)

    # Environmental certifications note
    cert_x = badge_start_x + badges_total_w + Inches(0.3)
    cert_w = SLIDE_W - cert_x - MARGIN
    add_rounded_rect(slide, cert_x, y, cert_w, badge_h, C_LIGHT_GRAY, radius_pt=6.0)
    add_rect(slide, cert_x, y, Inches(0.05), badge_h, C_GREEN)
    add_textbox(slide, cert_x + Inches(0.15), y + Inches(0.08),
                cert_w - Inches(0.25), Inches(0.22),
                "環境認証・制度への活用",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK, bold=True)
    add_textbox(slide, cert_x + Inches(0.15), y + Inches(0.32),
                cert_w - Inches(0.25), Inches(0.5),
                "RE100 / SBT / CDP\nグリーン電力証書\n非化石証書",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_SUB,
                word_wrap=True)

    add_footer(slide)
