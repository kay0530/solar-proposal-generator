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
        "1,000棟を超える産業用太陽光の施工管理実績に加え、遠隔監視"
        "システムや蓄電池制御システムを自社開発。",
        "営業から設計・調達・施工・監視・メンテナンスまでを一気通貫で"
        "行う体制により製造原価を圧縮し、競争力のある単価を実現しています。",
    ]),
    ("2", "柔軟な契約形態", [
        "全く同じ契約内容のお客様は一社もありません。前半10年と後半10年で"
        "単価を変える設計や、削減できたデマンド料金の一部還元など、",
        "お客様のご要望・状況に合わせた契約条件を個別に設計します。"
        "この柔軟性こそが選ばれ続ける理由です。",
    ]),
    ("3", "監視・制御システム開発力", [
        "発電量の監視・分析により、現地に行かずに問題を発見・予測する"
        "独自アルゴリズムを自社開発。1回の訪問で全対応できる体制を整え、",
        "メンテナンスコストを最小化します。システムの開発者であり"
        "ユーザーでもあることが、私たちの強みです。",
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
