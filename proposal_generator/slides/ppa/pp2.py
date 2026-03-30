"""
pp2.py - PPAモデルとは（static スライド）
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "PPAモデルとは"

FLOW_STEPS = [
    ("①", "PPA事業者",  "設備を設置・\n所有・管理"),
    ("②", "発電",       "太陽光で発電"),
    ("③", "自家消費",   "発電電力を\nお客様が利用"),
    ("④", "電気代削減", "使用量に応じて\n電力料金を支払い"),
]

FEATURES = [
    ("初期費用ゼロ",  "太陽光パネル・PCS等の設置費用はPPA事業者が全額負担します。"),
    ("維持管理不要",  "メンテナンス・保険・修繕費も全てPPA事業者が対応します。"),
    ("長期固定単価",  "契約期間中のPPA単価が固定されるため、電気代上昇リスクを回避できます。"),
    ("契約満了後",    "契約終了後は設備を無償譲渡または撤去します（契約内容により異なります）。"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    # Flow diagram (horizontal arrow steps) - wider gaps for landscape
    gap = Inches(0.18)
    step_w = (SLIDE_W - MARGIN * 2 - gap * 3) / 4
    step_h = Inches(1.3)
    for i, (num, label, desc) in enumerate(FLOW_STEPS):
        cx = MARGIN + i * (step_w + gap)
        add_rounded_rect(slide, cx, y, step_w, step_h, C_ORANGE)
        add_textbox(slide,
                    cx, y + Inches(0.05),
                    step_w, Inches(0.38),
                    num,
                    font_name=FONT_BLACK, font_size_pt=20,
                    font_color=C_LIGHT_ORANGE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide,
                    cx, y + Inches(0.4),
                    step_w, Inches(0.38),
                    label,
                    font_name=FONT_BLACK, font_size_pt=14,
                    font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide,
                    cx + Inches(0.1), y + Inches(0.75),
                    step_w - Inches(0.2), Inches(0.5),
                    desc,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_WHITE,
                    align=PP_ALIGN.CENTER)

        # Arrow between steps
        if i < len(FLOW_STEPS) - 1:
            arrow_x = cx + step_w + Inches(0.02)
            add_textbox(slide,
                        arrow_x, y + step_h / 2 - Inches(0.15),
                        Inches(0.14), Inches(0.3),
                        "▶",
                        font_name=FONT_BODY, font_size_pt=14,
                        font_color=C_ORANGE,
                        align=PP_ALIGN.CENTER)

    y += step_h + Inches(0.2)

    # Feature rows (2x2 grid for landscape layout)
    feat_w = (SLIDE_W - MARGIN * 2 - Inches(0.3)) / 2
    feat_h = Inches(0.5)
    for i, (feat_title, feat_desc) in enumerate(FEATURES):
        col = i % 2
        row = i // 2
        fx = MARGIN + col * (feat_w + Inches(0.3))
        fy = y + row * (feat_h + Inches(0.08))
        add_rect(slide, fx, fy, Inches(0.06), feat_h, C_ORANGE)
        add_textbox(slide,
                    fx + Inches(0.14), fy,
                    Inches(1.5), Inches(0.26),
                    feat_title,
                    font_name=FONT_BODY, font_size_pt=11,
                    font_color=C_DARK, bold=True)
        add_textbox(slide,
                    fx + Inches(0.14), fy + Inches(0.24),
                    feat_w - Inches(0.2), Inches(0.26),
                    feat_desc,
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB)

    add_footer(slide)
