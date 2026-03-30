"""
new_competitor.py - 他社比較表スライド

Comparison table: アルトエナジー vs 他社A vs 他社B
Rows: PPA単価, サービス内容, メンテナンス, 契約柔軟性, 実績
Highlights アルトエナジー column with orange accent.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, C_BORDER,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_table, _set_cell_bg,
)

TITLE = "他社比較"

# Default comparison data (used when data.get("competitors") is not available)
_DEFAULT_ROWS_PPA = [
    ["比較項目",       "アルトエナジー",                  "他社A",                    "他社B"],
    ["PPA単価",        "業界最安水準",                    "標準的",                   "やや高め"],
    ["サービス内容",   "設計〜施工〜運用まで一貫対応",     "設計・施工のみ",           "施工・運用"],
    ["メンテナンス",   "24時間遠隔監視＋定期点検込み",     "オプション（別途費用）",    "年1回点検のみ"],
    ["契約柔軟性",     "契約期間・条件をカスタマイズ可能", "固定プランのみ",           "一部カスタマイズ可"],
    ["実績",           "全国500件以上の導入実績",          "関東中心100件程度",        "50件程度"],
]
_DEFAULT_ROWS_EPC = [
    ["比較項目",       "アルトエナジー",                  "他社A",                    "他社B"],
    ["kW単価",         "業界最安水準",                    "標準的",                   "やや高め"],
    ["サービス内容",   "設計〜施工〜運用まで一貫対応",     "設計・施工のみ",           "施工・運用"],
    ["メンテナンス",   "24時間遠隔監視＋定期点検込み",     "オプション（別途費用）",    "年1回点検のみ"],
    ["保証内容",       "パネル25年・PCS10年・施工10年",    "パネル25年・PCS5年",       "パネル25年のみ"],
    ["実績",           "全国500件以上の導入実績",          "関東中心100件程度",        "50件程度"],
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.05)

    # Subtitle
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                "主要項目での他社比較",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_SUB)
    y += Inches(0.35)

    # Build rows from data or defaults
    competitors = data.get("competitors")
    if competitors and isinstance(competitors, list) and len(competitors) > 0:
        rows = competitors  # expect list of lists
    else:
        rows = _DEFAULT_ROWS_EPC if data.get("proposal_type") == "epc" else _DEFAULT_ROWS_PPA

    n_cols = len(rows[0])
    col_w = (SLIDE_W - MARGIN * 2) / n_cols
    col_widths = [int(col_w)] * n_cols

    # Draw table manually for better styling control
    row_h = Inches(0.55)
    table_x = MARGIN
    table_w = SLIDE_W - MARGIN * 2

    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cx = table_x + col_w * c
            cy = y + row_h * r

            # Determine background color
            if r == 0:
                # Header row
                bg = C_ORANGE
                font_color = C_WHITE
                bold = True
                font_size = 11
            elif c == 1:
                # アルトエナジー column - orange accent
                bg = C_LIGHT_ORANGE
                font_color = C_DARK
                bold = False
                font_size = 10
            elif r % 2 == 0:
                bg = C_WHITE
                font_color = C_DARK
                bold = False
                font_size = 10
            else:
                bg = C_LIGHT_GRAY
                font_color = C_DARK
                bold = False
                font_size = 10

            # First column (labels) is bold
            if c == 0 and r > 0:
                bold = True

            add_rounded_rect(slide, cx + Inches(0.02), cy + Inches(0.02),
                             col_w - Inches(0.04), row_h - Inches(0.04),
                             bg, radius_pt=3.0,
                             border_color=C_BORDER, border_pt=0.5)

            add_textbox(slide, cx + Inches(0.08), cy + Inches(0.08),
                        col_w - Inches(0.16), row_h - Inches(0.16),
                        str(cell_text),
                        font_name=FONT_BODY if not (r == 0) else FONT_BLACK,
                        font_size_pt=font_size,
                        font_color=font_color,
                        bold=bold,
                        align=PP_ALIGN.CENTER,
                        word_wrap=True)

    # Orange accent bar on アルトエナジー column
    accent_x = table_x + col_w
    accent_y = y + row_h  # skip header
    accent_h = row_h * (len(rows) - 1)
    add_rect(slide, accent_x + Inches(0.02), accent_y, Inches(0.05), accent_h, C_ORANGE)

    # Bottom note
    note_y = y + row_h * len(rows) + Inches(0.15)
    add_textbox(slide, MARGIN, note_y,
                SLIDE_W - MARGIN * 2, Inches(0.24),
                "※ 他社情報は一般的な市場調査に基づく参考値です",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_SUB,
                align=PP_ALIGN.RIGHT)

    add_footer(slide)
