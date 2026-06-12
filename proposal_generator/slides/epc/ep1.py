"""
ep1.py - なぜ今「EPC（設備購入）」なのか？（static slide）

Design v2 "Institutional Trust Grid":
  Band 1: standfirst lead paragraph (12.5pt) + brand illustration
  Band 2: 4 merit cards (2x2 white cards w/ accent bar + flat line icons)
  Band 3: conclusion band (C_PANEL + orange left bar)
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, C_DARK, C_NAVY, C_ORANGE, C_PANEL,
    FONT_BLACK, FONT_BODY, GAP_CARD, GAP_IN_CARD, MARGIN, SLIDE_W,
    LINE_SPACING_BODY, LINE_SPACING_LEAD, SIZE_BODY, SIZE_LEAD,
    add_card_with_accent, add_footer, add_header_bar, add_icon,
    add_image_contain, add_rect, add_textbox, asset_path,
    grid_x, grid_w, vstack,
)

TITLE = "なぜ今「EPC（設備購入）」なのか？"
EYEBROW = "01｜導入の背景"

STANDFIRST = (
    "燃料費や再エネ賦課金の影響で電気料金の上昇基調が続く中、"
    "自家消費型太陽光発電の導入は企業の競争力強化に直結します。"
    "EPC（設備購入）モデルは、お客様自身が設備を所有することで"
    "長期の電気代削減効果を最大限に享受できる、PPAにはないメリットを持つ選択肢です。"
    "減価償却による節税、契約期間の制約なし——税制優遇や補助金の活用により、"
    "実質的な投資負担も大幅に軽減できます。"
)

# (icon, title, body) — bodies fold in the live-deck merits
# (遮熱効果 / BCP対策 / 工場立地法対応) on the asset card.
CARDS = [
    ("factory", "資産所有",
     "設備はお客様の資産となり、減価償却で節税効果を享受。"
     "屋根の遮熱効果による空調負荷低減、BCP対策（非常時の電源確保）、"
     "工場立地法の環境施設対応にも寄与します。"),
    ("yen", "長期コスト削減",
     "PPA単価の支払いが不要となり、発電した電力はすべて電気代削減に直結。"
     "長期にわたる削減効果を最大化できます。"),
    ("doc", "税制優遇",
     "中小企業経営強化税制等の活用により、即時償却・税額控除が可能。"
     "導入初年度の税負担を大きく抑えられます。"),
    ("check", "補助金活用",
     "国・自治体の補助金を活用することで、初期投資を大幅に圧縮できます。"),
]

CONCLUSION = (
    "設備を「所有」することで、PPAにはないメリットを最大化——"
    "自社資産として太陽光発電を活用する時代です"
)


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    content_w = SLIDE_W - MARGIN * 2
    illust = asset_path("illust_decarb.png")
    band1_h = Inches(1.30) if illust else Inches(1.00)
    card_h = Inches(1.50)
    band2_h = card_h * 2 + GAP_CARD
    band3_h = Inches(0.80)
    ys = vstack(CONTENT_TOP, CONTENT_BOTTOM, [band1_h, band2_h, band3_h])

    # --- Band 1: standfirst lead paragraph (+ optional illustration) ---
    sf_w = grid_w(9) if illust else content_w
    add_textbox(slide, MARGIN, ys[0], sf_w, band1_h,
                STANDFIRST,
                font_name=FONT_BODY, font_size_pt=SIZE_LEAD,
                font_color=C_DARK, line_spacing=LINE_SPACING_LEAD,
                anchor=MSO_ANCHOR.TOP)
    if illust:
        try:
            add_image_contain(slide, grid_x(9) + Inches(0.2),
                              ys[0] - Inches(0.05),
                              grid_w(3) - Inches(0.2), band1_h + Inches(0.10),
                              illust)
        except Exception:
            pass

    # --- Band 2: 4 merit cards (2x2) ---
    for i, (icon_name, card_title, body) in enumerate(CARDS):
        col = i % 2
        row = i // 2
        x = grid_x(col * 6)
        w = grid_w(6)
        y0 = ys[1] + row * (card_h + GAP_CARD)
        cx, cy, cw, ch = add_card_with_accent(slide, x, y0, w, card_h)

        # flat line icon, top-right of the card
        add_icon(slide, icon_name, cx + cw - Inches(0.44), cy + Inches(0.02),
                 Inches(0.40))

        add_textbox(slide, cx, cy + Inches(0.04), cw - Inches(0.55),
                    Inches(0.26),
                    card_title,
                    font_name=FONT_BLACK, font_size_pt=SIZE_LEAD,
                    font_color=C_DARK, bold=True)

        body_y = cy + Inches(0.30) + GAP_IN_CARD
        add_textbox(slide, cx, body_y, cw, ch - Inches(0.46),
                    body,
                    font_name=FONT_BODY, font_size_pt=SIZE_BODY,
                    font_color=C_DARK, line_spacing=LINE_SPACING_BODY)

    # --- Band 3: conclusion band ---
    add_rect(slide, MARGIN, ys[2], content_w, band3_h, C_PANEL)
    add_rect(slide, MARGIN, ys[2], Inches(0.06), band3_h, C_ORANGE)
    add_textbox(slide, MARGIN + Inches(0.30), ys[2],
                content_w - Inches(0.60), band3_h,
                CONCLUSION,
                font_name=FONT_BLACK, font_size_pt=12,
                font_color=C_NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)

    add_footer(slide)
