"""
pp4a.py - なぜオルテナジーが選ばれるのか (Why Altenergy is chosen)

Design v2 "Institutional Trust Grid":
  Lead standfirst (12.5pt) + 3 tall strength cards
  (grid cols 0-3 / 4-7 / 8-11): circle number marker top-left
  + 14pt navy title + 10.5pt body paragraphs.
- 競争力のある単価
- ワンストップサービス
- 豊富な実績
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_W,
    LINE_SPACING_BODY, LINE_SPACING_LEAD,
    SIZE_BODY, SIZE_H2, SIZE_LEAD, SIZE_SMALL,
    add_card_with_accent, add_divider, add_footer, add_header_bar,
    add_multiline_textbox, add_number_marker, add_textbox,
    grid_x, grid_w, vstack,
)

TITLE = "なぜオルテナジーが選ばれるのか"
EYEBROW = "02｜オルテナジーの強み"

LEAD = (
    "PPAにおいて実績を伸ばしている企業は多くありません。"
    "オルテナジーが選ばれ続けるのは、3つの強みがあるからです。"
)

STRENGTHS = [
    ("1", "競争力のある単価", [
        "単価はリース金利・パネル原価などの外的要因と、"
        "施工・管理コストなどの内的要因で決まります。",
        "オルテナジーは自社施工体制と効率的なオペレーションで"
        "両方のコストを最適化し、競争力のある単価を実現しています。",
    ]),
    ("2", "ワンストップサービス", [
        "設計・施工・メンテナンス・発電事業運営までを自社グループで"
        "一貫対応。お客様の窓口は一つです。",
        "EPC事業者としての実績と発電事業者としてのノウハウを併せ持ち、"
        "設計変更や障害対応も迅速に行えます。",
    ]),
    ("3", "豊富な実績", [
        "累計設置容量100MW以上・導入企業数200社以上。"
        "工場・倉庫・商業施設など多様な建物への導入実績があります。",
        "全国対応のネットワークで、北海道から沖縄まで"
        "日本全国のお客様にサービスを提供しています。",
    ]),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    lead_h = Inches(0.55)
    card_h = Inches(4.30)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [lead_h, card_h])

    # --- Lead standfirst ---
    add_textbox(slide, MARGIN, ys[0], SLIDE_W - MARGIN * 2, lead_h,
                LEAD,
                font_name=FONT_BODY, font_size_pt=SIZE_LEAD,
                font_color=C_DARK, line_spacing=LINE_SPACING_LEAD)

    # --- 3 strength cards ---
    marker_d = Inches(0.34)
    for i, (num, card_title, paras) in enumerate(STRENGTHS):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, ys[1], w, card_h)

        add_number_marker(slide, cx + marker_d // 2,
                          cy + Inches(0.10) + marker_d // 2, num)
        add_textbox(slide, cx + marker_d + Inches(0.12), cy + Inches(0.10),
                    cw - marker_d - Inches(0.12), Inches(0.34),
                    card_title,
                    font_name=FONT_BLACK, font_size_pt=SIZE_H2,
                    font_color=C_NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

        add_divider(slide, cx, cy + Inches(0.56), cw)

        lines = []
        for j, para in enumerate(paras):
            if j > 0:
                # blank spacer paragraph between body paragraphs
                lines.append(("", FONT_BODY, SIZE_SMALL, C_DARK, False,
                              PP_ALIGN.LEFT))
            lines.append((para, FONT_BODY, SIZE_BODY, C_DARK, False,
                          PP_ALIGN.LEFT))
        add_multiline_textbox(slide, cx, cy + Inches(0.72), cw,
                              ch - Inches(0.78),
                              lines, line_spacing=LINE_SPACING_BODY)

    add_footer(slide)
