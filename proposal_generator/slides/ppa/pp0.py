"""
pp0.py - PPA表紙スライド (Refined cover design)

Layout: Asymmetric split with stronger typographic hierarchy
- Left 38%: Orange panel with logo, eyebrow tag, hero copy, service badge
- Right 62%: Customer name as hero, refined spec cards with accent stripes
- Subtle geometric accents (corner triangle, vertical hairline)
"""
from __future__ import annotations

import re
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from pptx.dml.color import RGBColor

from proposal_generator.utils import (
    C_BORDER, C_DARK, C_LIGHT_GRAY, C_NAVY, C_ORANGE, C_ORANGE_DARK, C_SUB,
    C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    SIZE_HERO, SIZE_H1, SIZE_H2, SIZE_H3, SIZE_BODY, SIZE_CAPTION, SIZE_SMALL,
    add_image_contain, add_rect, add_rounded_rect, add_textbox, add_divider,
)


def _fmt_date(val) -> str:
    s = str(val).split(" ")[0]
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return f"{m.group(1)}年{int(m.group(2))}月{int(m.group(3))}日"
    return s


def generate(slide, data: dict, logo_path: Path = None) -> None:
    company = data.get("company_name", "") or ""
    office = data.get("office_name", "") or ""
    prop_date = _fmt_date(data.get("proposal_date", "") or "")
    capacity = data.get("system_capacity_kw")
    years = int(data.get("contract_years", 20) or 20)
    unit_price = data.get("ppa_unit_price")
    address = data.get("address", "") or ""

    split_x = SLIDE_W * 0.38

    # ============================================================
    # LEFT PANEL (orange, full bleed)
    # ============================================================
    add_rect(slide, 0, 0, split_x, SLIDE_H, C_ORANGE)
    # Darker left accent strip (slightly wider for visual weight)
    add_rect(slide, 0, 0, Inches(0.08), SLIDE_H, C_ORANGE_DARK)

    # Logo (white version)
    _white_logo = data.get("_logo_white_path")
    _use_logo = _white_logo if _white_logo and Path(_white_logo).exists() else logo_path
    if _use_logo and Path(_use_logo).exists():
        add_image_contain(slide,
                          Inches(0.55), Inches(0.55),
                          Inches(2.3), Inches(0.55), _use_logo)

    # Eyebrow tag (small caps style)
    add_textbox(slide, Inches(0.55), Inches(2.0),
                split_x - Inches(0.9), Inches(0.22),
                "PROPOSAL  |  ONSITE PPA",
                font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                font_color=C_WHITE, bold=True)

    # Thin separator line under eyebrow
    add_rect(slide, Inches(0.55), Inches(2.30), Inches(0.5), Inches(0.02), C_WHITE)

    # Hero category text
    add_textbox(slide, Inches(0.55), Inches(2.55),
                split_x - Inches(0.9), Inches(0.36),
                "自家消費型",
                font_name=FONT_BODY, font_size_pt=14,
                font_color=C_WHITE, bold=False)

    # Main hero (3 lines, large)
    add_textbox(slide, Inches(0.55), Inches(2.95),
                split_x - Inches(0.9), Inches(2.0),
                "太陽光発電\nシステム",
                font_name=FONT_BLACK, font_size_pt=34,
                font_color=C_WHITE, bold=True)

    # Sub-tagline
    add_textbox(slide, Inches(0.55), Inches(4.95),
                split_x - Inches(0.9), Inches(0.85),
                "導入費用ゼロで\n電気代とCO₂を削減する\n次世代エネルギープラン。",
                font_name=FONT_BODY, font_size_pt=11,
                font_color=C_WHITE, bold=False)

    # Service badge (outlined style instead of solid for elegance)
    _badge_y = Inches(6.4)
    _badge_h = Inches(0.42)
    _badge_w = Inches(3.4)
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        int(Inches(0.55)), int(_badge_y),
        int(_badge_w), int(_badge_h),
    )
    badge.fill.background()  # transparent
    badge.line.color.rgb = C_WHITE
    badge.line.width = Pt(1.2)
    add_textbox(slide, Inches(0.55), _badge_y + Inches(0.08),
                _badge_w, Inches(0.30),
                "オンサイトPPAサービスのご提案",
                font_name=FONT_BLACK, font_size_pt=11,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    # Bottom: company name + date (compact, two-line)
    add_textbox(slide, Inches(0.55), SLIDE_H - Inches(0.95),
                split_x - Inches(0.9), Inches(0.22),
                "株式会社オルテナジー",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_WHITE, bold=True)
    add_textbox(slide, Inches(0.55), SLIDE_H - Inches(0.70),
                split_x - Inches(0.9), Inches(0.20),
                prop_date,
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_WHITE)

    # ============================================================
    # RIGHT PANEL (white)
    # ============================================================

    info_x = split_x + Inches(0.7)
    info_w = SLIDE_W - split_x - Inches(1.1)

    # Top eyebrow on right side
    add_textbox(slide, info_x, Inches(0.9),
                info_w, Inches(0.22),
                "FOR",
                font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                font_color=C_ORANGE, bold=True)

    # Customer name (hero)
    add_textbox(slide, info_x, Inches(1.20),
                info_w, Inches(1.30),
                f"{company}",
                font_name=FONT_BLACK, font_size_pt=32,
                font_color=C_DARK, bold=True)

    # 御中
    add_textbox(slide, info_x, Inches(2.55),
                info_w, Inches(0.32),
                "御中",
                font_name=FONT_BODY, font_size_pt=16,
                font_color=C_SUB)

    # Decorative thin orange accent line
    add_rect(slide, info_x, Inches(3.05),
             Inches(2.4), Inches(0.04), C_ORANGE)

    # Office name + address
    if office:
        add_textbox(slide, info_x, Inches(3.25),
                    info_w, Inches(0.32),
                    office,
                    font_name=FONT_BODY, font_size_pt=14,
                    font_color=C_DARK, bold=True)
    if address:
        add_textbox(slide, info_x, Inches(3.65),
                    info_w, Inches(0.24),
                    f"設置先住所：{address}",
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_SUB)

    # ---- System spec cards (refined with accent bar via add_card_with_accent style) ----
    y_spec = Inches(4.45)
    specs = []
    if capacity:
        specs.append(("設備容量", f"{capacity:.1f}", "kW"))
    if unit_price:
        specs.append(("PPA単価", f"¥{unit_price:.2f}", "/kWh"))
    specs.append(("契約期間", f"{years}", "年"))

    card_gap = Inches(0.18)
    card_w = (info_w - card_gap * (len(specs) - 1)) / len(specs) if specs else Inches(2.0)
    card_h = Inches(1.55)

    for i, (label, value, unit) in enumerate(specs):
        cx = info_x + i * (card_w + card_gap)
        # Card background
        add_rounded_rect(slide, cx, y_spec, card_w, card_h, C_WHITE,
                         radius_pt=8.0,
                         border_color=C_BORDER, border_pt=0.5)
        # Top accent bar
        add_rect(slide, cx + Inches(0.15), y_spec,
                 card_w - Inches(0.30), Inches(0.05), C_ORANGE)
        # Label (top)
        add_textbox(slide, cx, y_spec + Inches(0.20),
                    card_w, Inches(0.25),
                    label,
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_SUB, bold=True, align=PP_ALIGN.CENTER)
        # Value (large)
        add_textbox(slide, cx, y_spec + Inches(0.55),
                    card_w, Inches(0.55),
                    value,
                    font_name=FONT_BLACK, font_size_pt=24,
                    font_color=C_ORANGE, bold=True, align=PP_ALIGN.CENTER)
        # Unit (small, below value)
        add_textbox(slide, cx, y_spec + Inches(1.10),
                    card_w, Inches(0.25),
                    unit,
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

    # ---- Bottom band with copyright ----
    bottom_band_y = SLIDE_H - Inches(0.45)
    add_rect(slide, split_x, bottom_band_y, SLIDE_W - split_x, Inches(0.02), C_BORDER)
    add_textbox(slide, split_x, SLIDE_H - Inches(0.30),
                SLIDE_W - split_x, Inches(0.20),
                "Copyright 2026 altenergy, Inc.   |   https://altenergy.co.jp/",
                font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                font_color=C_SUB, align=PP_ALIGN.CENTER)

    # Bottom accent bar (extends across slide)
    add_rect(slide, 0, SLIDE_H - Inches(0.08), SLIDE_W, Inches(0.08), C_ORANGE)
