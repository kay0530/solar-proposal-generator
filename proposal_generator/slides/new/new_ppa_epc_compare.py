"""
new_ppa_epc_compare.py - PPA vs EPC 比較スライド

Side-by-side comparison table:
- PPA: 初期費用ゼロ, 電力購入契約, メンテ込み, 契約期間中は解約不可
- EPC: 初期投資必要, 設備所有, 自社メンテ, 資産計上・償却可能
Orange for PPA column, blue/gray for EPC column.
Recommendation based on data.get("proposal_type").
"""
from __future__ import annotations
from pathlib import Path
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, C_BORDER,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header,
)

TITLE = "PPA vs EPC 比較"

# Blue/gray accent for EPC column
C_BLUE = RGBColor(0x4A, 0x6F, 0xA5)
C_LIGHT_BLUE = RGBColor(0xEB, 0xF0, 0xF7)

COMPARISON_ITEMS = [
    ("初期費用",     "ゼロ（PPA事業者が負担）",          "設備購入費用が必要"),
    ("設備所有権",   "PPA事業者が所有",                   "自社所有（資産計上）"),
    ("電力料金",     "PPA単価で固定（長期安定）",         "自家発電のため実質無料"),
    ("メンテナンス", "PPA事業者が全て対応（込み）",       "自社で手配（別途コスト）"),
    ("契約期間",     "15〜20年（期間中は原則解約不可）",  "制約なし（自社設備）"),
    ("税務メリット", "経費処理が可能",                    "減価償却・税額控除が可能"),
    ("リスク",       "発電リスクはPPA事業者側",           "故障・性能低下リスクは自社"),
    ("適している企業", "初期投資を抑えたい企業",           "自己資金・融資で投資可能な企業"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Subtitle
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                "導入方式の違いを分かりやすく比較",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_SUB)
    y += Inches(0.35)

    # Column layout: label | PPA | EPC
    label_w = Inches(2.2)
    col_w = (SLIDE_W - MARGIN * 2 - label_w - Inches(0.15)) / 2
    gap = Inches(0.08)

    ppa_x = MARGIN + label_w + gap
    epc_x = ppa_x + col_w + gap

    row_h = Inches(0.5)

    # Column headers
    # PPA header
    add_rounded_rect(slide, ppa_x, y, col_w, Inches(0.42), C_ORANGE, radius_pt=4.0)
    add_textbox(slide, ppa_x, y + Inches(0.06),
                col_w, Inches(0.3),
                "PPA（電力購入契約）",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    # EPC header
    add_rounded_rect(slide, epc_x, y, col_w, Inches(0.42), C_BLUE, radius_pt=4.0)
    add_textbox(slide, epc_x, y + Inches(0.06),
                col_w, Inches(0.3),
                "EPC（自社購入）",
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    y += Inches(0.5)

    # Comparison rows
    for r, (label, ppa_text, epc_text) in enumerate(COMPARISON_ITEMS):
        ry = y + r * row_h
        row_bg = C_LIGHT_GRAY if r % 2 == 0 else C_WHITE

        # Label cell
        add_rounded_rect(slide, MARGIN, ry + Inches(0.02),
                         label_w, row_h - Inches(0.04), row_bg,
                         radius_pt=3.0)
        add_rect(slide, MARGIN, ry + Inches(0.02),
                 Inches(0.05), row_h - Inches(0.04), C_ORANGE)
        add_textbox(slide, MARGIN + Inches(0.12), ry + Inches(0.06),
                    label_w - Inches(0.16), row_h - Inches(0.12),
                    label,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_DARK, bold=True)

        # PPA cell
        add_rounded_rect(slide, ppa_x, ry + Inches(0.02),
                         col_w, row_h - Inches(0.04), C_LIGHT_ORANGE,
                         radius_pt=3.0)
        add_textbox(slide, ppa_x + Inches(0.08), ry + Inches(0.06),
                    col_w - Inches(0.16), row_h - Inches(0.12),
                    ppa_text,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_DARK,
                    word_wrap=True)

        # EPC cell
        add_rounded_rect(slide, epc_x, ry + Inches(0.02),
                         col_w, row_h - Inches(0.04), C_LIGHT_BLUE,
                         radius_pt=3.0)
        add_textbox(slide, epc_x + Inches(0.08), ry + Inches(0.06),
                    col_w - Inches(0.16), row_h - Inches(0.12),
                    epc_text,
                    font_name=FONT_BODY, font_size_pt=9,
                    font_color=C_DARK,
                    word_wrap=True)

    # Recommendation banner
    proposal_type = data.get("proposal_type", "PPA")
    if proposal_type and "epc" in str(proposal_type).lower():
        rec_text = "御社にはEPC（自社購入）モデルをおすすめします"
        rec_bg = C_BLUE
    else:
        rec_text = "御社にはPPA（電力購入契約）モデルをおすすめします"
        rec_bg = C_ORANGE

    rec_y = y + len(COMPARISON_ITEMS) * row_h + Inches(0.12)
    rec_h = Inches(0.45)
    add_rounded_rect(slide, MARGIN, rec_y,
                     SLIDE_W - MARGIN * 2, rec_h, rec_bg, radius_pt=6.0)
    add_textbox(slide, MARGIN, rec_y + Inches(0.08),
                SLIDE_W - MARGIN * 2, Inches(0.28),
                f"▶  {rec_text}",
                font_name=FONT_BLACK, font_size_pt=14,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)

    add_footer(slide)
