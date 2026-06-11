"""
pp2.py - PPAモデルとは（static スライド）

Design v2 "Institutional Trust Grid":
  Band 1: 4-step flow — white cards + true circle number markers
          + navy chevron shapes between steps, descriptions below cards
  Band 2: contract feature notes on a C_PANEL band
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_HAIR, C_NAVY, C_PANEL, C_SUB,
    C_WHITE, FONT_BLACK, FONT_BODY, MARGIN, SLIDE_W,
    LINE_SPACING_BODY, SIZE_BODY, SIZE_CAPTION, SIZE_LEAD,
    add_footer, add_header_bar, add_multiline_textbox, add_number_marker,
    add_rect, add_rounded_rect, add_section_header, add_textbox, vstack,
)

TITLE = "PPAモデルとは"
EYEBROW = "01｜導入の背景"

FLOW_STEPS = [
    ("1", "PPA事業者", "設備を設置・所有・管理"),
    ("2", "発電", "太陽光で発電"),
    ("3", "自家消費", "発電電力をお客様が利用"),
    ("4", "電気代削減", "使用量に応じて電力料金を支払い"),
]

FEATURES = [
    ("初期費用ゼロ", "太陽光パネル・PCS等の設置費用はPPA事業者が全額負担します。"),
    ("維持管理不要", "メンテナンス・保険・修繕費も全てPPA事業者が対応します。"),
    ("長期固定単価", "契約期間中のPPA単価が固定されるため、電気代上昇リスクを回避できます。"),
    ("契約満了後", "契約終了後は設備を無償譲渡または撤去します（契約内容により異なります）。"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    header_h = Inches(0.34)
    card_h = Inches(1.25)
    desc_h = Inches(0.50)
    panel_h = Inches(2.10)
    flow_block_h = header_h + card_h + Inches(0.08) + desc_h
    feat_block_h = header_h + panel_h
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [flow_block_h, feat_block_h])

    # --- Band 1: 4-step flow ---
    y = ys[0]
    add_section_header(slide, MARGIN, y, content_w, "電力供給のしくみ（4ステップ）")
    y += header_h

    chevron_gap = Inches(0.30)
    step_w = (content_w - chevron_gap * 3) // 4
    for i, (num, label, desc) in enumerate(FLOW_STEPS):
        sx = MARGIN + i * (step_w + chevron_gap)
        add_rounded_rect(slide, sx, y, step_w, card_h, C_WHITE,
                         border_color=C_HAIR, border_pt=0.75)
        add_number_marker(slide, sx + step_w // 2, y + Inches(0.36), num)
        add_textbox(slide, sx + Inches(0.08), y + Inches(0.64),
                    step_w - Inches(0.16), Inches(0.30),
                    label,
                    font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                    font_color=C_NAVY, bold=True, align=PP_ALIGN.CENTER)
        add_textbox(slide, sx + Inches(0.04), y + card_h + Inches(0.08),
                    step_w - Inches(0.08), desc_h,
                    desc,
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_DARK, align=PP_ALIGN.CENTER,
                    line_spacing=LINE_SPACING_BODY)

        if i < len(FLOW_STEPS) - 1:
            ch_w = Inches(0.18)
            ch_h = Inches(0.24)
            chev = slide.shapes.add_shape(
                MSO_SHAPE.CHEVRON,
                int(sx + step_w + (chevron_gap - ch_w) // 2),
                int(y + card_h // 2 - ch_h // 2),
                int(ch_w), int(ch_h))
            chev.fill.solid()
            chev.fill.fore_color.rgb = C_NAVY
            chev.line.fill.background()
            chev.shadow.inherit = False

    # --- Band 2: feature notes on warm panel ---
    y2 = ys[1]
    add_section_header(slide, MARGIN, y2, content_w,
                       "契約のポイント（PPA事業者がすべて対応）")
    y2 += header_h
    add_rect(slide, MARGIN, y2, content_w, panel_h, C_PANEL)

    pad = Inches(0.25)
    col_gap = Inches(0.40)
    cell_w = (content_w - pad * 2 - col_gap) // 2
    cell_h = Inches(0.66)
    row_gap = panel_h - pad * 2 - cell_h * 2
    for i, (feat_title, feat_desc) in enumerate(FEATURES):
        col = i % 2
        row = i // 2
        fx = MARGIN + pad + col * (cell_w + col_gap)
        fy = y2 + pad + row * (cell_h + max(Inches(0.12), row_gap))
        add_multiline_textbox(
            slide, fx, fy, cell_w, cell_h,
            [
                (feat_title, FONT_BLACK, SIZE_BODY, C_NAVY, True, PP_ALIGN.LEFT),
                (feat_desc, FONT_BODY, SIZE_CAPTION, C_SUB, False, PP_ALIGN.LEFT),
            ],
            line_spacing=LINE_SPACING_BODY)

    add_footer(slide)
