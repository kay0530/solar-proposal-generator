"""
pp13.py - よくあるご質問 (design v2: Institutional Trust Grid)

Layout (A4 landscape):
  - v2 white header, eyebrow '07｜よくあるご質問'
  - 5 Q&A rows: outline pill 'Qn' marker + bold question + body answer,
    hairline dividers between rows, vstack-justified
"""

from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_SUB,
    FONT_BLACK, FONT_BODY, LINE_SPACING_BODY, MARGIN,
    SIZE_BODY, SLIDE_W,
    add_divider, add_footer, add_header_bar, add_pill_label,
    add_textbox, vstack,
)

TITLE = "よくあるご質問"
EYEBROW = "07｜よくあるご質問"

FAQ_ITEMS = [
    {
        "q": "初期費用はかかりますか？",
        "a": "PPAモデルでは初期費用ゼロでご導入いただけます。設備の設置費用・メンテナンス費用は全てPPA事業者が負担します。",
    },
    {
        "q": "メンテナンスは必要ですか？",
        "a": "設備の保守・点検・修理は全て当社が対応いたします。24時間遠隔監視により、異常の早期発見・迅速対応を行います。",
    },
    {
        "q": "契約期間中に移転したらどうなりますか？",
        "a": "移転先への設備の移設対応が可能です。移転先の条件を確認の上、最適なプランをご提案いたします。",
    },
    {
        "q": "停電時はどうなりますか？",
        "a": "蓄電池を併設することで、停電時にも非常用電源としてご利用いただけます。BCP対策としても有効です。",
    },
    {
        "q": "屋根が劣化しませんか？",
        "a": "設置前に防水処理を実施し、定期点検で屋根の状態を確認します。むしろパネルが直射日光を遮り、屋根の劣化を抑える効果もあります。",
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render PP13 (FAQ) onto an already-added blank slide.

    data keys used: (none required - static content)
    """
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2

    item_h = Inches(0.82)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM,
                [item_h] * len(FAQ_ITEMS))

    pill_w = Inches(0.55)
    pill_h = Inches(0.28)
    text_x = int(MARGIN) + int(Inches(0.75))
    text_w = int(content_w) - int(Inches(0.75))

    for i, item in enumerate(FAQ_ITEMS):
        iy = ys[i]

        # Q marker (outline pill) + question on the marker's centerline
        add_pill_label(slide, MARGIN, iy, pill_w, pill_h, f"Q{i + 1}")
        add_textbox(slide, text_x, iy, text_w, pill_h,
                    item["q"],
                    font_name=FONT_BLACK, font_size_pt=11,
                    font_color=C_DARK, bold=True,
                    anchor=MSO_ANCHOR.MIDDLE)

        # Answer body (word-wrapped, no manual breaks)
        add_textbox(slide, text_x, iy + int(Inches(0.36)),
                    text_w, int(item_h) - int(Inches(0.36)),
                    item["a"],
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_SUB, line_spacing=LINE_SPACING_BODY)

        # Hairline divider between rows (midway in the gap)
        if i < len(FAQ_ITEMS) - 1:
            div_y = (int(iy) + int(item_h) + int(ys[i + 1])) // 2
            add_divider(slide, MARGIN, div_y, content_w)

    add_footer(slide)
