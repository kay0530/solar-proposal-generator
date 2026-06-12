"""
new_urgency.py - 緊急性訴求スライド (design v2: Institutional Trust Grid)

Title: "今が導入のベストタイミング"
3-card band (電気料金動向 / 排出権 / 補助金期限) with key stats as orange
number runs, plus a C_PANEL conclusion band carrying the equipment-cost
outlook and the call to action.
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_ORANGE, C_PANEL, C_SUB,
    FONT_BLACK, FONT_BODY, LINE_SPACING_BODY, MARGIN,
    SIZE_BODY, SIZE_CAPTION, SLIDE_W,
    add_card_with_accent, add_divider, add_footer, add_header_bar,
    add_multiline_textbox, add_number_unit, add_rect, add_textbox,
    grid_w, grid_x, vstack,
)

TITLE = "今が導入のベストタイミング"
EYEBROW = "補足｜いま導入すべき理由"

URGENCY_CARDS = [
    {
        "label": "電気料金動向",
        "title": "電力料金の高騰リスク",
        "stat": ("3.49", "円/kWh｜再エネ賦課金（2024年度）"),
        "bullets": [
            "燃料費調整額の上昇トレンド（2022年以降、高止まり継続）",
            "容量拠出金の新規導入（2024年度〜）による更なる上乗せ",
        ],
        "emphasis": "今の電気代がベースラインではない — 上がり続ける",
    },
    {
        "label": "排出権",
        "title": "カーボンプライシングの本格化",
        "stat": ("2026", "年｜GX-ETS（排出量取引制度）本格稼働"),
        "bullets": [
            "炭素賦課金の段階的導入（2028年〜）",
            "RE100/SBTi対応企業の急増 → 排出枠の早期確保が有利",
        ],
        "emphasis": "対策が遅れるほどコストが増大する",
    },
    {
        "label": "補助金期限",
        "title": "補助金の縮小傾向",
        "stat": ("1/3", "へ縮小見込み｜需要家主導型 補助率（R7）"),
        "bullets": [
            "補助率は年々低下（R4: 2/3 → R6: 1/2）",
            "自治体補助金も予算枠が縮小傾向",
            "申請件数の増加により競争率が上昇中",
        ],
        "emphasis": "今年度が最も有利な条件で導入できるタイミング",
    },
]

EQUIP_TITLE = "設備コストの動向 — 待つリスク"
EQUIP_BODY = ("パネル価格は低下傾向ですが、PCS・蓄電池の供給逼迫、施工人材不足による"
              "工事費の上昇、半導体・原材料の価格変動リスクにより、"
              "総コストは必ずしも下がりません。")
CTA_TEXT = ("→ まずは無料シミュレーション・現地調査をご依頼ください。"
            "補助金申請のサポートも対応いたします。")


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    # ---- Vertical layout: lead + 3-card band + conclusion band ----
    lead_h = Inches(0.26)
    card_h = Inches(3.55)
    band_h = Inches(1.55)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [lead_h, card_h, band_h])

    add_textbox(slide, MARGIN, ys[0], content_w, lead_h,
                "導入を先送りにするリスクと、今動くメリット",
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- 3 urgency cards (top accent, orange stat runs) ----
    for i, card in enumerate(URGENCY_CARDS):
        x = grid_x(i * 4)
        w = grid_w(4)
        cx, cy, cw, ch = add_card_with_accent(slide, x, ys[1], w, card_h,
                                              accent_position="top")

        # Category label (tracked caption)
        add_textbox(slide, cx, cy, cw, Inches(0.18), card["label"],
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                    bold=True, tracking_pt=1.2)
        # Card title
        add_textbox(slide, cx, cy + Inches(0.22), cw, Inches(0.26),
                    card["title"],
                    font_name=FONT_BLACK, font_size_pt=12.5,
                    font_color=C_NAVY, bold=True)
        # Key stat: orange number + context unit on one baseline
        num, unit = card["stat"]
        add_number_unit(slide, cx, cy + Inches(0.50), cw, Inches(0.42),
                        num, unit, number_size_pt=20, unit_size_pt=9,
                        unit_color=C_SUB)
        add_divider(slide, cx, cy + Inches(1.04), cw)

        # Supporting bullets
        bullet_lines = [(f"・{b}", FONT_BODY, 9.5, C_DARK, False,
                         PP_ALIGN.LEFT) for b in card["bullets"]]
        add_multiline_textbox(slide, cx, cy + Inches(1.16),
                              cw, ch - Inches(1.16) - Inches(0.60),
                              bullet_lines, line_spacing=LINE_SPACING_BODY)

        # Emphasis line pinned to card bottom
        add_textbox(slide, cx, cy + ch - Inches(0.55), cw, Inches(0.55),
                    f"→ {card['emphasis']}",
                    font_size_pt=10, font_color=C_ORANGE, bold=True,
                    word_wrap=True, line_spacing=1.2)

    # ---- Conclusion band: equipment-cost outlook + CTA (C_PANEL) ----
    band_y = ys[2]
    add_rect(slide, MARGIN, band_y, content_w, band_h, C_PANEL)
    add_rect(slide, MARGIN, band_y, Inches(0.05), band_h, C_ORANGE)
    bx = MARGIN + Inches(0.20)
    bw = content_w - Inches(0.40)
    add_textbox(slide, bx, band_y + Inches(0.13), bw, Inches(0.24),
                EQUIP_TITLE,
                font_name=FONT_BLACK, font_size_pt=12.5,
                font_color=C_NAVY, bold=True)
    add_textbox(slide, bx, band_y + Inches(0.44), bw, Inches(0.58),
                EQUIP_BODY, font_size_pt=SIZE_BODY, font_color=C_DARK,
                word_wrap=True, line_spacing=LINE_SPACING_BODY)
    add_textbox(slide, bx, band_y + band_h - Inches(0.40), bw, Inches(0.30),
                CTA_TEXT, font_size_pt=11, font_color=C_ORANGE, bold=True)

    add_footer(slide)
