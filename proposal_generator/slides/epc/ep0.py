"""
ep0.py - EPC表紙スライド (Premium cover design)

Layout: Split design matching pp0.py premium style
- Left 38%: Orange panel with logo, tagline, service badge
- Right 62%: Customer name, office, address, spec cards
- Bottom accent bar with copyright
"""
from __future__ import annotations

import re
from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from pptx.dml.color import RGBColor

from proposal_generator.utils import (
    C_DARK, C_LIGHT_GRAY, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_rect, add_rounded_rect, add_textbox,
    fmt_yen,
)

# Cover-specific colors
C_ORANGE_DARK = RGBColor(0xC0, 0x3A, 0x0A)   # darker orange for cover panel


def _fmt_date(val) -> str:
    """Format date string: '2026-03-28' or datetime -> '2026年3月28日'."""
    s = str(val).split(" ")[0]
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    return s


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP0 (EPC cover slide) with premium split-panel design.

    data keys used:
        company_name, office_name, address, proposal_date,
        system_capacity_kw, selling_price,
        subsidy_name, subsidy_amount,
        investment_recovery_yr
    """
    company = data.get("company_name", "") or ""
    office = data.get("office_name", "") or ""
    address = data.get("address", "") or ""
    prop_date = _fmt_date(data.get("proposal_date", "") or "")
    capacity = data.get("system_capacity_kw")
    selling_price = data.get("selling_price")
    subsidy_name = data.get("subsidy_name", "") or ""
    subsidy_amount = data.get("subsidy_amount")
    recovery_yr = data.get("investment_recovery_yr")

    # ---- Layout constants ----
    split_x = SLIDE_W * 0.38  # left panel width

    # ============================================================
    # LEFT PANEL (orange)
    # ============================================================
    add_rect(slide, 0, 0, split_x, SLIDE_H, C_ORANGE)
    # Darker accent strip at left edge
    add_rect(slide, 0, 0, Inches(0.06), SLIDE_H, C_ORANGE_DARK)

    # Logo (white version works best on orange)
    if logo_path and Path(logo_path).exists():
        from proposal_generator.utils import add_image_contain
        add_image_contain(slide,
                          Inches(0.5), Inches(0.5),
                          Inches(2.5), Inches(0.7), logo_path)

    # Main tagline
    add_textbox(slide, Inches(0.5), Inches(2.2),
                split_x - Inches(0.8), Inches(0.40),
                "自家消費型太陽光発電設備のご提案",
                font_name=FONT_BODY, font_size_pt=13,
                font_color=C_WHITE, bold=True)

    add_textbox(slide, Inches(0.5), Inches(2.7),
                split_x - Inches(0.8), Inches(1.2),
                "設備購入型の\n電気代 / CO₂\n削減プラン",
                font_name=FONT_BLACK, font_size_pt=28,
                font_color=C_WHITE, bold=True)

    # Service type badge (darker orange on orange)
    _badge_y = Inches(4.3)
    _badge_h = Inches(0.45)
    add_rounded_rect(slide, Inches(0.5), _badge_y,
                     Inches(3.2), _badge_h, C_ORANGE_DARK)
    add_textbox(slide, Inches(0.5), _badge_y + Inches(0.08),
                Inches(3.2), Inches(0.30),
                "EPC（設備販売）ご提案書",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    # Company info (bottom of left panel)
    add_textbox(slide, Inches(0.5), SLIDE_H - Inches(1.2),
                split_x - Inches(0.8), Inches(0.25),
                "株式会社オルテナジー",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_WHITE)
    add_textbox(slide, Inches(0.5), SLIDE_H - Inches(0.9),
                split_x - Inches(0.8), Inches(0.20),
                prop_date,
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_WHITE)

    # ============================================================
    # RIGHT PANEL (white)
    # ============================================================

    # Customer name (hero text)
    add_textbox(slide, split_x + Inches(0.5), Inches(1.0),
                SLIDE_W - split_x - Inches(0.8), Inches(1.2),
                f"{company}",
                font_name=FONT_BLACK, font_size_pt=30,
                font_color=C_DARK, bold=True)

    # "御中" below company name
    add_textbox(slide, split_x + Inches(0.5), Inches(2.0),
                SLIDE_W - split_x - Inches(0.8), Inches(0.45),
                "御中",
                font_name=FONT_BODY, font_size_pt=18,
                font_color=C_SUB)

    # Thin accent line
    add_rect(slide, split_x + Inches(0.5), Inches(2.6),
             Inches(2.0), Inches(0.04), C_ORANGE)

    # Office / Address
    info_x = split_x + Inches(0.5)
    info_w = SLIDE_W - split_x - Inches(0.8)

    if office:
        add_textbox(slide, info_x, Inches(2.85),
                    info_w, Inches(0.30),
                    office,
                    font_name=FONT_BODY, font_size_pt=14,
                    font_color=C_DARK)
    if address:
        add_textbox(slide, info_x, Inches(3.25),
                    info_w, Inches(0.25),
                    f"設置先住所：{address}",
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB)

    # ---- System spec cards ----
    y_spec = Inches(4.0)
    specs = []
    if capacity:
        specs.append(("設備容量", f"{capacity:.1f} kW"))
    if selling_price:
        specs.append(("販売価格", fmt_yen(selling_price)))
    if subsidy_name and subsidy_amount:
        specs.append(("補助金", fmt_yen(subsidy_amount)))
    if recovery_yr:
        specs.append(("想定回収年数", f"{recovery_yr}年"))

    if specs:
        card_gap = Inches(0.12)
        card_w = (info_w - card_gap * (len(specs) - 1)) / len(specs)
        card_h = Inches(1.2)

        for i, (label, value) in enumerate(specs):
            cx = info_x + i * (card_w + card_gap)
            add_rounded_rect(slide, cx, y_spec, card_w, card_h, C_LIGHT_GRAY)
            add_rect(slide, cx, y_spec, card_w, Inches(0.05), C_ORANGE)
            add_textbox(slide, cx, y_spec + Inches(0.15), card_w, Inches(0.25),
                        label,
                        font_name=FONT_BODY, font_size_pt=9,
                        font_color=C_SUB, align=PP_ALIGN.CENTER)
            add_textbox(slide, cx, y_spec + Inches(0.45), card_w, Inches(0.55),
                        value,
                        font_name=FONT_BLACK, font_size_pt=20,
                        font_color=C_ORANGE, bold=True,
                        align=PP_ALIGN.CENTER)

    # ---- Bottom accent bar ----
    add_rect(slide, 0, SLIDE_H - Inches(0.10), SLIDE_W, Inches(0.10), C_ORANGE)

    # Copyright
    add_textbox(slide, split_x, SLIDE_H - Inches(0.45),
                SLIDE_W - split_x, Inches(0.20),
                "Copyright 2026 Altenergy, Inc. All rights reserved",
                font_name=FONT_BODY, font_size_pt=7,
                font_color=C_SUB, align=PP_ALIGN.CENTER)
