"""
new_ff.py - FF振り返り（前回ヒアリング結果）スライド (design v2: Institutional Trust Grid)

FF = Fact Findings. Shows what was learned in the previous customer visit
as a 2x2 grid of white cards with orange left accents:
- Current situation & challenges
- Person-in-charge needs
- Management appeal points
- Constraints / concerns
"""
from __future__ import annotations

from pathlib import Path

from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_SUB,
    FONT_BLACK, LINE_SPACING_BODY, MARGIN, SIZE_BODY, SLIDE_W,
    add_card_with_accent, add_divider, add_footer, add_header_bar,
    add_textbox, grid_w, grid_x, vstack,
)

TITLE = "ヒアリング結果（FF振り返り）"
EYEBROW = "補足｜ヒアリング振り返り"

# (data key, card title, empty-state hint)
SECTIONS = [
    ("ff_current_situation", "現状・課題",
     "電気代・設備状況・運用課題など"),
    ("ff_customer_needs", "担当者ニーズ",
     "担当者が上司・経営者に訴えたいこと"),
    ("ff_mgmt_needs", "経営者へのアピール",
     "経営者が関心を持つポイント（ROI・リスク・環境）"),
    ("ff_constraints", "制約・懸念事項",
     "屋根強度・予算・タイムライン・補助金期限など"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    company = data.get("company_name", "") or ""

    # ---- Vertical layout: lead line + two card rows ----
    lead_h = Inches(0.26)
    card_h = Inches(2.70)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [lead_h, card_h, card_h])

    # ---- Lead: customer name + hearing date ----
    lead_who = f"{company}　様" if company else "お客様"
    # Excel-calc path may return proposal_date as datetime; keep date part only
    prop_date = str(data.get("proposal_date", "") or "—").split(" ")[0]
    lead = f"{lead_who}　｜　ヒアリング実施日：{prop_date}"
    add_textbox(slide, MARGIN, ys[0], content_w, lead_h, lead,
                font_size_pt=SIZE_BODY, font_color=C_SUB)

    # ---- 2x2 white cards with orange left accent ----
    for i, (key, section_title, hint) in enumerate(SECTIONS):
        col = i % 2
        row = i // 2
        x = grid_x(col * 6)
        w = grid_w(6)
        y = ys[1 + row]
        cx, cy, cw, ch = add_card_with_accent(slide, x, y, w, card_h,
                                              accent_position="left")

        # Card title (navy bold) + hairline divider
        add_textbox(slide, cx, cy, cw, Inches(0.26), section_title,
                    font_name=FONT_BLACK, font_size_pt=12.5,
                    font_color=C_NAVY, bold=True)
        add_divider(slide, cx, cy + Inches(0.36), cw)

        # Body: hearing content, or quiet hint when not yet filled
        content = str(data.get(key, "") or "").strip()
        body = content if content else f"（{hint}）"
        add_textbox(slide, cx, cy + Inches(0.48), cw, ch - Inches(0.48),
                    body, font_size_pt=SIZE_BODY,
                    font_color=C_DARK if content else C_SUB,
                    word_wrap=True, line_spacing=LINE_SPACING_BODY)

    add_footer(slide)
