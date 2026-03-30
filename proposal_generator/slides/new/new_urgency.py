"""
new_urgency.py - 緊急性訴求スライド

Title: "今が導入のベストタイミング"
4 urgency points with detailed market data, plus call-to-action banner.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
)

TITLE = "今が導入のベストタイミング"

URGENCY_POINTS = [
    {
        "label": "電力料金",
        "title": "電力料金の高騰リスク",
        "bullets": [
            "燃料費調整額の上昇トレンド（2022年以降、高止まり継続）",
            "再エネ賦課金の段階的引き上げ（2024年度: 3.49円/kWh）",
            "容量拠出金の新規導入（2024年度〜、更なる上乗せ要因）",
        ],
        "emphasis": "今の電気代がベースラインではない — 上がり続ける",
    },
    {
        "label": "排出権",
        "title": "カーボンプライシングの本格化",
        "bullets": [
            "GX-ETS（排出量取引制度）2026年本格稼働",
            "炭素賦課金の段階的導入（2028年〜）",
            "RE100/SBTi対応企業の急増 → 排出枠の早期確保が有利",
        ],
        "emphasis": "対策が遅れるほどコストが増大する",
    },
    {
        "label": "補助金",
        "title": "補助金の縮小傾向",
        "bullets": [
            "需要家主導型太陽光: 補助率が年々低下（R4: 2/3 → R6: 1/2 → R7: 1/3見込み）",
            "自治体補助金も予算枠が縮小傾向",
            "申請件数増加により競争率が上昇中",
        ],
        "emphasis": "今年度が最も有利な条件で導入できるタイミング",
    },
    {
        "label": "設備価格",
        "title": "設備コストの動向",
        "bullets": [
            "パネル価格は低下傾向だが、PCS・蓄電池は供給逼迫",
            "施工人材の不足 → 工事費が上昇傾向",
            "半導体・原材料の価格変動リスク",
        ],
        "emphasis": "総コストは必ずしも下がらない — 待つリスクがある",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    # Subtitle
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.28),
                "導入を先送りにするリスクと、今動くメリット",
                font_name=FONT_BODY, font_size_pt=12,
                font_color=C_SUB)
    y += Inches(0.35)

    # 4 urgency cards (2x2 grid)
    card_cols = 2
    card_gap = Inches(0.2)
    card_w = (SLIDE_W - MARGIN * 2 - card_gap) / card_cols
    card_h = Inches(2.45)
    row_gap = Inches(0.12)

    for i, point in enumerate(URGENCY_POINTS):
        col = i % card_cols
        row = i // card_cols
        cx = MARGIN + col * (card_w + card_gap)
        cy = y + row * (card_h + row_gap)

        # Card background
        add_rounded_rect(slide, cx, cy, card_w, card_h, C_LIGHT_ORANGE, radius_pt=6.0)

        # Orange top accent bar
        add_rect(slide, cx, cy, card_w, Inches(0.06), C_ORANGE)

        # Label badge
        badge_w = Inches(1.0)
        badge_h = Inches(0.30)
        add_rounded_rect(slide, cx + Inches(0.12), cy + Inches(0.15),
                         badge_w, badge_h, C_ORANGE, radius_pt=4.0)
        add_textbox(slide, cx + Inches(0.12), cy + Inches(0.16),
                    badge_w, badge_h - Inches(0.02),
                    point["label"],
                    font_name=FONT_BLACK, font_size_pt=11,
                    font_color=C_WHITE, bold=True,
                    align=PP_ALIGN.CENTER)

        # Title
        add_textbox(slide, cx + Inches(1.2), cy + Inches(0.15),
                    card_w - Inches(1.35), Inches(0.30),
                    point["title"],
                    font_name=FONT_BODY, font_size_pt=13,
                    font_color=C_DARK, bold=True)

        # Divider
        add_rect(slide, cx + Inches(0.12), cy + Inches(0.52),
                 card_w - Inches(0.24), Inches(0.015), C_ORANGE)

        # Bullet points
        bullet_y = cy + Inches(0.60)
        for bullet in point["bullets"]:
            add_textbox(slide, cx + Inches(0.18), bullet_y,
                        card_w - Inches(0.36), Inches(0.30),
                        f"・{bullet}",
                        font_name=FONT_BODY, font_size_pt=9,
                        font_color=C_SUB, word_wrap=True)
            bullet_y += Inches(0.32)

        # Emphasis line (orange bold)
        add_textbox(slide, cx + Inches(0.12), cy + card_h - Inches(0.48),
                    card_w - Inches(0.24), Inches(0.35),
                    f"→ {point['emphasis']}",
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_ORANGE, bold=True, word_wrap=True)

    # Call-to-action banner at bottom
    cta_y = y + 2 * (card_h + row_gap) + Inches(0.08)
    cta_h = Inches(0.50)
    add_rounded_rect(slide, MARGIN, cta_y,
                     SLIDE_W - MARGIN * 2, cta_h, C_ORANGE, radius_pt=6.0)
    add_textbox(slide, MARGIN, cta_y + Inches(0.06),
                SLIDE_W - MARGIN * 2, Inches(0.20),
                "今すぐ無料シミュレーションをご依頼ください",
                font_name=FONT_BLACK, font_size_pt=15,
                font_color=C_WHITE, bold=True,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, MARGIN, cta_y + Inches(0.28),
                SLIDE_W - MARGIN * 2, Inches(0.18),
                "お見積り・現地調査は無料で承ります。補助金申請のサポートも対応いたします。",
                font_name=FONT_BODY, font_size_pt=9,
                font_color=C_WHITE,
                align=PP_ALIGN.CENTER)

    add_footer(slide)
