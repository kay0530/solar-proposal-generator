"""
pp6a.py - 事業スキーム (Design v2: Institutional Trust Grid)

Three-entity scheme diagram on the 12-column grid:
  顧客 (cols 0-3, orange-tick card) ←→ オルテナジー (cols 4-7, navy-header
  card) ←→ リース会社 (cols 8-11, hairline card)
Flows are real chevron shapes with white chip labels (電気料金 / 電力供給 /
資金提供 / 返済) — no text-glyph arrows, no wrapped labels.
Bottom band: system spec list + key contract points incl. off-balance note.
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CARD_PAD, CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_HAIR, C_NAVY, C_ORANGE,
    C_PANEL, C_SUB, C_WHITE, FONT_BLACK, FONT_BODY, GAP_BLOCK, MARGIN,
    SIZE_BODY, SIZE_CAPTION, SIZE_SMALL, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    grid_w, grid_x, vstack,
)

TITLE = "事業スキーム"
EYEBROW = "03｜効果シミュレーション"


def _chip(slide, cx, cy, text: str, color=C_ORANGE):
    """White chip label centered at (cx, cy) — sits above/below a flow arrow."""
    w = Inches(1.10)
    h = Inches(0.26)
    x = int(cx) - int(w) // 2
    y = int(cy) - int(h) // 2
    add_rounded_rect(slide, x, y, w, h, C_WHITE,
                     border_color=color, border_pt=0.75)
    add_textbox(slide, x, y + Inches(0.025), w, Inches(0.20),
                text,
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=color, bold=True, align=PP_ALIGN.CENTER)


def _flow_arrow(slide, x, y, w, direction: str, color=C_NAVY):
    """Thin chevron arrow shape. direction: 'right' or 'left'."""
    h = Inches(0.16)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.CHEVRON, int(x), int(y), int(w), int(h))
    if direction == "left":
        shape.rotation = 180.0
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    company = data.get("company_name") or "お客様"
    lease_company = data.get("lease_company") or "リース会社"

    # ---- Vertical plan: contract labels / entity row / spec band ----
    label_h = Inches(0.26)
    box_h = Inches(1.95)
    spec_h = Inches(2.30)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [label_h, box_h, spec_h],
                min_gap=GAP_BLOCK)
    label_y, box_y, spec_y = ys[0], ys[1], ys[2]

    # ---- Entity columns on the grid (cards inset so arrows have room) ----
    inset = Inches(0.28)
    l_x, l_w = grid_x(0), grid_w(4) - inset
    c_x, c_w = grid_x(4) + inset // 2, grid_w(4) - inset
    r_x, r_w = grid_x(8) + inset, grid_w(4) - inset

    # ---- Contract-zone labels above the gaps ----
    gap1_cx = (int(l_x) + int(l_w) + int(c_x)) // 2
    gap2_cx = (int(c_x) + int(c_w) + int(r_x)) // 2
    add_textbox(slide, gap1_cx - Inches(1.2), label_y, Inches(2.4), label_h,
                "オンサイトPPA契約",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, gap2_cx - Inches(1.2), label_y, Inches(2.4), label_h,
                "リース契約",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, bold=True, align=PP_ALIGN.CENTER)

    # ---- Left card: customer (orange top tick) ----
    add_rounded_rect(slide, l_x, box_y, l_w, box_h, C_WHITE,
                     border_color=C_HAIR, border_pt=0.75)
    add_rect(slide, l_x + Inches(0.12), box_y,
             l_w - Inches(0.24), Inches(0.045), C_ORANGE)
    add_textbox(slide, l_x + CARD_PAD, box_y + Inches(0.18),
                l_w - CARD_PAD * 2, Inches(0.62),
                f"{company}様",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER,
                line_spacing=1.2)
    add_textbox(slide, l_x + CARD_PAD, box_y + box_h - Inches(0.84),
                l_w - CARD_PAD * 2, Inches(0.70),
                "太陽光発電の電力を使用し、使用した電力量に応じた電気料金のみお支払い",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, align=PP_ALIGN.CENTER, line_spacing=1.3)

    # ---- Center card: altenergy (navy header band) ----
    add_rounded_rect(slide, c_x, box_y, c_w, box_h, C_WHITE,
                     border_color=C_HAIR, border_pt=0.75)
    hdr_h = Inches(0.46)
    add_rect(slide, c_x, box_y, c_w, hdr_h, C_NAVY)
    add_textbox(slide, c_x, box_y + Inches(0.09),
                c_w, Inches(0.30),
                "オルテナジーグループ",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, c_x + CARD_PAD, box_y + hdr_h + Inches(0.16),
                c_w - CARD_PAD * 2, box_h - hdr_h - Inches(0.30),
                "発電事業者 ／ EPC事業者\n太陽光発電システムの設置工事・保守メンテナンスを一貫実施",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_DARK, align=PP_ALIGN.CENTER, line_spacing=1.35)

    # ---- Right card: lease company (hairline tick) ----
    add_rounded_rect(slide, r_x, box_y, r_w, box_h, C_WHITE,
                     border_color=C_HAIR, border_pt=0.75)
    add_rect(slide, r_x + Inches(0.12), box_y,
             r_w - Inches(0.24), Inches(0.045), C_HAIR)
    add_textbox(slide, r_x + CARD_PAD, box_y + Inches(0.18),
                r_w - CARD_PAD * 2, Inches(0.62),
                str(lease_company),
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER,
                line_spacing=1.2)
    add_textbox(slide, r_x + CARD_PAD, box_y + box_h - Inches(0.84),
                r_w - CARD_PAD * 2, Inches(0.70),
                "設備資金を提供し、リース料の返済を受領",
                font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                font_color=C_SUB, align=PP_ALIGN.CENTER, line_spacing=1.3)

    # ---- Flow arrows + chips (upper = rightward, lower = leftward) ----
    arr_up_y = box_y + int(box_h * 0.34)
    arr_dn_y = box_y + int(box_h * 0.64)

    # Customer <-> Altenergy
    g1_x = int(l_x) + int(l_w) + int(Inches(0.06))
    g1_w = int(c_x) - g1_x - int(Inches(0.06))
    _flow_arrow(slide, g1_x, arr_up_y, g1_w, "right", C_ORANGE)   # 電気料金 →
    _flow_arrow(slide, g1_x, arr_dn_y, g1_w, "left", C_NAVY)      # ← 電力供給
    _chip(slide, g1_x + g1_w // 2, arr_up_y - int(Inches(0.20)), "電気料金", C_ORANGE)
    _chip(slide, g1_x + g1_w // 2, arr_dn_y + int(Inches(0.34)), "電力供給", C_NAVY)

    # Altenergy <-> Lease
    g2_x = int(c_x) + int(c_w) + int(Inches(0.06))
    g2_w = int(r_x) - g2_x - int(Inches(0.06))
    _flow_arrow(slide, g2_x, arr_up_y, g2_w, "right", C_ORANGE)   # 返済 →
    _flow_arrow(slide, g2_x, arr_dn_y, g2_w, "left", C_NAVY)      # ← 資金提供
    _chip(slide, g2_x + g2_w // 2, arr_up_y - int(Inches(0.20)), "返済", C_ORANGE)
    _chip(slide, g2_x + g2_w // 2, arr_dn_y + int(Inches(0.34)), "資金提供", C_NAVY)

    # ---- Bottom band: system spec + key points ----
    content_w = SLIDE_W - MARGIN * 2
    add_rect(slide, MARGIN, spec_y, content_w, spec_h, C_PANEL)
    add_rect(slide, MARGIN, spec_y, Inches(0.06), spec_h, C_ORANGE)

    add_textbox(slide, MARGIN + Inches(0.30), spec_y + Inches(0.16),
                Inches(4.0), Inches(0.28),
                "太陽光発電システム",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_NAVY, bold=True)

    # Left: spec list
    capacity = data.get("system_capacity_kw", 0) or 0
    details = []
    if capacity:
        details.append(f"設備容量: {capacity:,.2f} kW")
    panels = data.get("panels") or []
    if panels:
        p = panels[0]
        details.append(
            f"パネル: {p.get('model', '')} "
            f"{(p.get('watt_per_unit', 0) or 0):.0f}W × {p.get('count', 0)}枚")
    pcs_list = data.get("pcs_list") or []
    if pcs_list:
        q = pcs_list[0]
        details.append(
            f"PCS: {q.get('model', '')} "
            f"{(q.get('kw_per_unit', 0) or 0):.1f}kW × {q.get('count', 0)}台")
    details.append(f"契約期間: {int(data.get('contract_years', 20) or 20)}年")

    add_textbox(slide, MARGIN + Inches(0.30), spec_y + Inches(0.56),
                int(content_w * 0.44), spec_h - Inches(0.85),
                "\n".join(details),
                font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                font_color=C_DARK, line_spacing=1.4)

    # Right: key contract points (incl. off-balance note from the live deck)
    points = [
        "設備の所有権はオルテナジーグループに帰属（リース資産のため、"
        "会計上お客様の計上はオフバランス）",
        "設置・保守メンテナンス費用は全て当社負担",
        "お客様は使用した電力量に応じた電気料金のみお支払い",
        "契約期間終了後は設備の無償譲渡または撤去を選択可能",
    ]
    px = MARGIN + int(content_w * 0.48)
    pw = int(content_w * 0.50)
    py = spec_y + Inches(0.22)
    for pt_text in points:
        add_rect(slide, px, py + Inches(0.07), Inches(0.08), Inches(0.08),
                 C_ORANGE)
        add_textbox(slide, px + Inches(0.18), py, pw - Inches(0.18),
                    Inches(0.46),
                    pt_text,
                    font_name=FONT_BODY, font_size_pt=SIZE_CAPTION,
                    font_color=C_DARK, line_spacing=1.25)
        py += Inches(0.50)

    # Lease company provenance note (from live deck P7)
    if "シーエナジー" in str(lease_company):
        add_textbox(slide, MARGIN + Inches(0.30),
                    spec_y + spec_h - Inches(0.26),
                    content_w - Inches(0.60), Inches(0.18),
                    "※ ㈱シーエナジーは中部電力グループ企業（中部電力100％出資会社）です。",
                    font_name=FONT_BODY, font_size_pt=SIZE_SMALL,
                    font_color=C_SUB)

    add_footer(slide)
