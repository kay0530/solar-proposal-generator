"""
pp11.py - 導入スケジュールスライド (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - v2 white header, eyebrow '05｜導入の流れ'
  - Lead line: total period from contract to operation (12.5pt)
  - Horizontal chevron timeline: white fill + 1pt navy outline,
    grid-snapped equal widths; STEP number 16pt orange inside,
    duration 9pt below each chevron
  - Milestone table (audited-figures style) fills residual height
"""

from __future__ import annotations

from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, GAP_BLOCK, MARGIN, SIZE_CAPTION, SIZE_LEAD,
    SIZE_SMALL, SLIDE_W, TABLE_ROW_H,
    add_footer, add_header_bar, add_multiline_textbox, add_section_header,
    add_table, add_textbox, grid_w, grid_x, vstack,
)

TITLE = "導入スケジュール"
EYEBROW = "05｜導入の流れ"

# Phase definitions: (step_no, name, duration, milestone lines)
PHASES = [
    ("1", "ご契約・設計", "1〜2ヶ月",
     ["現地調査・電力需要分析・システム設計", "PPA契約締結・補助金申請手続き"]),
    ("2", "機器調達", "2〜3ヶ月",
     ["太陽光パネル・PCS等の機器発注", "架台・配線部材の手配"]),
    ("3", "施工・設置", "1〜2ヶ月",
     ["屋根防水処理・架台設置", "パネル設置・電気工事・系統連系"]),
    ("4", "運転開始", "—",
     ["試運転・検査完了後に発電開始", "遠隔監視システム稼働"]),
]

NOTE = "※ 上記は標準的な工期の目安です。補助金申請の審査期間・系統連系協議等により前後する場合があります。"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP11 (installation schedule). data keys used: none (static)."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    # ---- Block heights for vertical justify ----
    lead_h = Inches(0.30)
    chev_h = Inches(1.00)
    dur_h = Inches(0.22)
    timeline_h = int(chev_h) + int(Inches(0.06)) + int(dur_h)
    n_table_rows = 1 + sum(len(p[3]) for p in PHASES)  # header + milestones
    table_h = int(TABLE_ROW_H) * n_table_rows
    note_h = Inches(0.22)
    sect_h = int(Inches(0.34)) + table_h + int(Inches(0.06)) + int(note_h)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, timeline_h, sect_h], min_gap=GAP_BLOCK)

    # ---- Lead line ----
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h,
                "ご契約から運転開始まで：約4〜7ヶ月（目安）",
                font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                font_color=C_DARK, bold=True, line_spacing=1.4)

    # ---- Chevron timeline (grid-snapped, equal widths) ----
    span = 12 // len(PHASES)  # 4 phases -> 3 grid columns each
    chev_y = ys[1]
    for i, (step_no, name, duration, _lines) in enumerate(PHASES):
        cx = grid_x(i * span)
        cw = grid_w(span)
        chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON,
                                      int(cx), int(chev_y),
                                      int(cw), int(chev_h))
        chev.fill.solid()
        chev.fill.fore_color.rgb = C_WHITE
        chev.line.color.rgb = C_NAVY
        chev.line.width = Pt(1.0)
        chev.shadow.inherit = False

        add_multiline_textbox(
            slide,
            int(cx) + int(Inches(0.15)), int(chev_y) + int(Inches(0.18)),
            int(cw) - int(Inches(0.30)), Inches(0.64),
            [
                (f"STEP {step_no}", FONT_BLACK, 16, C_ORANGE, True,
                 PP_ALIGN.CENTER),
                (name, FONT_BODY, 11, C_DARK, True, PP_ALIGN.CENTER),
            ],
            line_spacing=1.2)

        add_textbox(slide, cx, int(chev_y) + int(chev_h) + int(Inches(0.06)),
                    cw, dur_h, duration,
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                    align=PP_ALIGN.CENTER)

    # ---- Milestone table ----
    sect_y = ys[2]
    add_section_header(slide, MARGIN, sect_y, content_w, "工程別マイルストーン")
    tbl_y = int(sect_y) + int(Inches(0.34))

    rows = [["工程", "期間", "主なマイルストーン"]]
    for step_no, name, duration, lines in PHASES:
        for j, line in enumerate(lines):
            if j == 0:
                rows.append([f"STEP {step_no}　{name}", duration, line])
            else:
                rows.append(["", "", line])

    c1 = Inches(2.6)
    c2 = Inches(1.5)
    c3 = int(content_w) - int(c1) - int(c2)
    add_table(slide, MARGIN, tbl_y, content_w, rows, [c1, c2, c3])

    # ---- Note (hugs the table bottom) ----
    add_textbox(slide, MARGIN, tbl_y + table_h + int(Inches(0.06)),
                content_w, note_h, NOTE,
                font_size_pt=SIZE_SMALL, font_color=C_SUB)

    add_footer(slide)
