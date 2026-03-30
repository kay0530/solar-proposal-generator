"""
pp6a.py - 事業スキーム (Business scheme diagram)

PDF P7: Shows the PPA contract flow diagram.
  顧客 ←→ オルテナジー ←→ リース会社
  with: PPA契約, 電気料金, 電力供給, リース契約, 返済, 資金提供
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "事業スキーム"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    company = data.get("company_name", "お客様")
    lease_company = data.get("lease_company", "リース会社")

    y = CONTENT_TOP + Inches(0.3)

    # ---- Header labels ----
    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.3),
                "オンサイトPPA契約　　　　　　　　　　　　　　　リース契約",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_SUB, align=PP_ALIGN.CENTER)
    y += Inches(0.35)

    # ---- Three entity boxes ----
    box_w = Inches(2.8)
    box_h = Inches(1.4)
    center_x = (SLIDE_W - box_w) / 2

    # Left box: Customer
    lx = MARGIN + Inches(0.5)
    add_rounded_rect(slide, lx, y, box_w, box_h, C_LIGHT_ORANGE)
    add_rect(slide, lx, y, box_w, Inches(0.06), C_ORANGE)
    add_textbox(slide, lx, y + Inches(0.20), box_w, Inches(0.35),
                f"{company}様",
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, lx, y + Inches(0.65), box_w, Inches(0.55),
                "電気料金お支払い\n↓\n電力をご使用",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_SUB, align=PP_ALIGN.CENTER)

    # Center box: Altenergy
    add_rounded_rect(slide, center_x, y, box_w, box_h, C_ORANGE)
    add_textbox(slide, center_x, y + Inches(0.15), box_w, Inches(0.35),
                "オルテナジーグループ",
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, center_x, y + Inches(0.55), box_w, Inches(0.70),
                "太陽光発電システム設置工事\n発電事業者\nメンテナンス実施\nEPC事業者",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_WHITE, align=PP_ALIGN.CENTER)

    # Right box: Lease company
    rx = SLIDE_W - MARGIN - Inches(0.5) - box_w
    add_rounded_rect(slide, rx, y, box_w, box_h, C_LIGHT_GRAY)
    add_rect(slide, rx, y, box_w, Inches(0.06), C_ORANGE)
    lease_name = lease_company if lease_company else "リース会社"
    add_textbox(slide, rx, y + Inches(0.20), box_w, Inches(0.35),
                lease_name,
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, rx, y + Inches(0.65), box_w, Inches(0.55),
                "資金提供\n↓\n返済受取",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_SUB, align=PP_ALIGN.CENTER)

    # ---- Arrow labels between boxes ----
    arrow_y = y + box_h / 2

    # Left ↔ Center arrows
    _draw_arrow_label(slide, lx + box_w, arrow_y - Inches(0.2),
                      center_x - lx - box_w, "電気料金", above=True)
    _draw_arrow_label(slide, lx + box_w, arrow_y + Inches(0.05),
                      center_x - lx - box_w, "電力供給", above=False)

    # Center ↔ Right arrows
    _draw_arrow_label(slide, center_x + box_w, arrow_y - Inches(0.2),
                      rx - center_x - box_w, "返済", above=True)
    _draw_arrow_label(slide, center_x + box_w, arrow_y + Inches(0.05),
                      rx - center_x - box_w, "資金提供", above=False)

    y += box_h + Inches(0.5)

    # ---- Solar system illustration box ----
    sys_box_w = SLIDE_W - MARGIN * 2
    sys_box_h = Inches(2.2)
    add_rounded_rect(slide, MARGIN, y, sys_box_w, sys_box_h, C_LIGHT_GRAY)

    add_textbox(slide, MARGIN + Inches(0.2), y + Inches(0.1),
                sys_box_w - Inches(0.4), Inches(0.30),
                "太陽光発電システム",
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_ORANGE, bold=True)

    # System details
    capacity = data.get("system_capacity_kw", 0)
    panel_info = ""
    panels = data.get("panels", [])
    if panels:
        p = panels[0]
        panel_info = f"パネル: {p.get('model', '')} {p.get('watt_per_unit', 0):.0f}W × {p.get('count', 0)}枚"

    pcs_info = ""
    pcs_list = data.get("pcs_list", [])
    if pcs_list:
        q = pcs_list[0]
        pcs_info = f"PCS: {q.get('model', '')} {q.get('kw_per_unit', 0):.1f}kW × {q.get('count', 0)}台"

    details = [
        f"設備容量: {capacity:.2f} kW" if capacity else "",
        panel_info,
        pcs_info,
        f"契約期間: {data.get('contract_years', 20)}年",
    ]
    detail_text = "\n".join(d for d in details if d)

    add_textbox(slide, MARGIN + Inches(0.2), y + Inches(0.50),
                sys_box_w / 2, sys_box_h - Inches(0.60),
                detail_text,
                font_name=FONT_BODY, font_size_pt=10, font_color=C_DARK)

    # Right side: key points
    key_points = (
        "● 設備の所有権はオルテナジーに帰属\n"
        "● 設備の設置・メンテナンス費用は全て当社負担\n"
        "● お客様は使用した電力量に応じた電気料金のみお支払い\n"
        "● 契約期間終了後、設備の無償譲渡または撤去"
    )
    add_textbox(slide, MARGIN + sys_box_w / 2, y + Inches(0.50),
                sys_box_w / 2 - Inches(0.2), sys_box_h - Inches(0.60),
                key_points,
                font_name=FONT_BODY, font_size_pt=10, font_color=C_DARK)

    add_footer(slide)


def _draw_arrow_label(slide, x, y, w, text, above=True):
    """Draw a text label between two boxes (arrow placeholder)."""
    add_textbox(slide, x, y, w, Inches(0.22),
                f"← {text} →" if not above else f"→ {text} →",
                font_name=FONT_BODY, font_size_pt=8,
                font_color=C_ORANGE, bold=True, align=PP_ALIGN.CENTER)
