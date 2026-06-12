"""
new_ppa_epc_compare.py - PPA vs EPC 比較スライド (design v2)

Audited comparison table (比較項目 | PPA | EPC):
  - row labels left-aligned, descriptions centered, 9pt
  - highlight_col tints the recommended side; 'おすすめ' pill above it
  - 適している企業 row promoted to two cards under the table
  - Navy recommendation band driven by data["proposal_type"]
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_NAVY_LIGHT, C_ORANGE,
    C_SUB, C_WHITE, FONT_BLACK, GAP_CARD, MARGIN, SIZE_BODY, SIZE_CAPTION,
    SIZE_H2, SIZE_LEAD, SIZE_TABLE, SLIDE_W, TABLE_ROW_H,
    add_card_with_accent, add_footer, add_header_bar, add_pill_label,
    add_rect, add_section_header, add_table, add_textbox, vstack,
)

TITLE = "PPA vs EPC 比較"
EYEBROW = "02｜比較検討"

COMPARISON_ITEMS = [
    ("初期費用",     "ゼロ（PPA事業者が負担）",           "設備購入費用が必要"),
    ("設備所有権",   "PPA事業者が所有",                   "自社所有（資産計上）"),
    ("電力料金",     "PPA単価で固定（長期安定）",         "自家発電のため実質無料"),
    ("メンテナンス", "PPA事業者が全て対応（込み）",       "自社で手配（別途コスト）"),
    ("契約期間",     "15〜20年（期間中は原則解約不可）",  "制約なし（自社設備）"),
    ("税務メリット", "経費処理が可能",                    "減価償却・税額控除が可能"),
    ("リスク",       "発電リスクはPPA事業者側",           "故障・性能低下リスクは自社"),
]

# 適している企業 (former table row, shown as cards)
SUITABLE_PPA = "初期投資を抑えたい企業"
SUITABLE_EPC = "自己資金・融資で投資可能な企業"

HEADER_PPA = "PPA（電力購入契約）"
HEADER_EPC = "EPC（自社購入）"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)
    content_w = SLIDE_W - MARGIN * 2

    proposal_type = data.get("proposal_type", "PPA")
    is_epc = bool(proposal_type) and "epc" in str(proposal_type).lower()

    rows = [["比較項目", HEADER_PPA, HEADER_EPC]]
    rows += [[label, ppa, epc] for label, ppa, epc in COMPARISON_ITEMS]
    n_rows = len(rows)

    # ---- Block heights for vertical justify ----
    lead_h = Inches(0.26)
    pill_h = Inches(0.26)
    table_block_h = (int(pill_h) + int(Inches(0.08))
                     + int(TABLE_ROW_H) * n_rows)
    suit_block_h = int(Inches(0.36)) + int(Inches(0.95))
    band_h = Inches(0.52)

    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [lead_h, table_block_h, suit_block_h, band_h],
                min_gap=GAP_CARD)

    # ---- Lead ----
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h,
                "導入方式の違いを分かりやすく比較",
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- Comparison table with highlight on the recommended column ----
    label_w = int(Inches(1.7))
    col_w = (int(content_w) - label_w) // 2
    col_widths = [label_w, col_w, col_w]
    highlight_col = 2 if is_epc else 1

    # 'おすすめ' pill centered above the recommended column
    pill_w = int(Inches(1.05))
    rec_col_x = int(MARGIN) + label_w + col_w * (highlight_col - 1)
    add_pill_label(slide, rec_col_x + (col_w - pill_w) // 2, ys[1],
                   pill_w, pill_h, "おすすめ")

    table_y = int(ys[1]) + int(pill_h) + int(Inches(0.08))
    add_table(slide, MARGIN, table_y, content_w, rows, col_widths,
              font_size_pt=SIZE_TABLE, highlight_col=highlight_col)

    # ---- 適している企業 cards ----
    suit_y = int(ys[2])
    add_section_header(slide, MARGIN, suit_y, content_w, "適している企業")
    card_y = suit_y + int(Inches(0.36))
    card_h = Inches(0.95)
    card_w = (int(content_w) - int(GAP_CARD)) // 2
    cards = [
        (HEADER_PPA, SUITABLE_PPA, not is_epc),
        (HEADER_EPC, SUITABLE_EPC, is_epc),
    ]
    for i, (tag, text, recommended) in enumerate(cards):
        x = int(MARGIN) + i * (card_w + int(GAP_CARD))
        cx, cy, cw, ch = add_card_with_accent(
            slide, x, card_y, card_w, card_h,
            accent_color=C_ORANGE if recommended else C_NAVY_LIGHT,
            accent_position="top")
        add_textbox(slide, cx, int(cy) + int(Inches(0.04)), cw, Inches(0.20),
                    tag, font_size_pt=SIZE_CAPTION,
                    font_color=C_ORANGE if recommended else C_SUB, bold=True)
        add_textbox(slide, cx, int(cy) + int(Inches(0.30)), cw, Inches(0.34),
                    text, font_size_pt=SIZE_LEAD, font_color=C_DARK,
                    bold=True)

    # ---- Recommendation band (navy CTA + orange tick) ----
    band_y = ys[3]
    add_rect(slide, MARGIN, band_y, content_w, band_h, C_NAVY)
    add_rect(slide, MARGIN, band_y, Inches(0.06), band_h, C_ORANGE)
    model = HEADER_EPC if is_epc else HEADER_PPA
    add_textbox(slide, int(MARGIN) + int(Inches(0.20)), band_y,
                int(content_w) - int(Inches(0.40)), band_h,
                f"御社には {model} モデルをおすすめします",
                font_name=FONT_BLACK, font_size_pt=SIZE_H2,
                font_color=C_WHITE, bold=True, align=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE)

    add_footer(slide)
