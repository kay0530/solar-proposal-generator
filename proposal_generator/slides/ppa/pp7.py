"""
pp7.py - ご利用料金 (PPA pricing) — design system v2 "one-number slide"

Left (cols 0-6): 48pt metric hero (PPA単価, 20-year fixed) + 2 lines of
supporting copy + 初期ご負担金額 0円 sub-metric (28pt number/unit pair).
Right (cols 7-11): mini comparison table (現行単価 / PPA単価 / 削減率,
PPA row highlighted) + 4 merit lines with small orange square markers.
Bottom: 8pt assumptions note.
"""
from __future__ import annotations

from pathlib import Path

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

from proposal_generator.utils import (
    CONTENT_BOTTOM, CONTENT_TOP, MARGIN, SLIDE_W,
    C_DARK, C_ORANGE, C_SUB,
    FONT_BODY,
    SIZE_BODY, SIZE_CAPTION, SIZE_SMALL,
    GAP_BLOCK, TABLE_ROW_H,
    add_footer, add_header_bar, add_metric_hero, add_multiline_textbox,
    add_number_unit, add_rect, add_section_header, add_table, add_textbox,
    grid_w, grid_x, vstack,
)

TITLE = "ご利用料金"
EYEBROW = "04｜ご契約条件"

MERITS = [
    ("再エネ賦課金の上昇対策",
     "基本的には上昇傾向の再エネ賦課金。"
     "上がれば上がるほど自家消費によるメリットは大きくなります。"),
    ("燃料費等調整単価の上昇対策",
     "燃料費が高騰すればプラスに振れる調整額。"
     "再エネ賦課金同様、支払う必要がなくなります。"),
    ("環境価値（CO2排出量抑制）",
     "RE100達成には環境クレジットの購入が必須である企業も。"
     "市場調達する量を削減可能。"),
    ("炭素税対策",
     "2050年までの脱炭素化に向けて段階的に炭素税率を"
     "引き上げていくという計画があります。"),
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP7 (PPA pricing) onto a blank slide."""
    add_header_bar(slide, TITLE, logo_path, eyebrow=EYEBROW)

    tax_display = data.get("tax_display", "税抜") or "税抜"
    ppa_price = float(data.get("ppa_unit_price", 0) or 0)
    years = int(data.get("contract_years", 20) or 20)

    # Current average unit price (for the comparison table)
    cur_price = None
    try:
        annual_cost = data.get("annual_cost")
        annual_kwh = float(data.get("annual_kwh", 0) or 0)
        if annual_cost and annual_kwh > 0:
            cur_price = float(annual_cost) / annual_kwh
    except (TypeError, ValueError):
        cur_price = None

    price_str = f"{ppa_price:.2f}" if ppa_price else "—"

    # ---- Bottom assumptions note (anchored, 8pt) ----
    notes = [
        f"※ 金額は全て{tax_display}表記です。通常の電力供給契約で発生する基本料金はかかりません。",
        "※ 再生可能エネルギー促進賦課金および燃料費調整額のお支払いが不要となります。",
        "※ 環境価値は貴社に帰属するものとし、上記ご提案単価は環境価値を含めた金額です。",
        "※ ご提案単価の有効期限はご提案日より1ヶ月間となります。",
    ]
    if cur_price is not None:
        notes.append("※ 現行平均単価はご提供いただいた電気料金実績より算出した参考値です。")
    note_h = Inches(0.16) * len(notes) + Inches(0.06)
    note_y = CONTENT_BOTTOM - note_h
    add_multiline_textbox(
        slide, MARGIN, note_y, SLIDE_W - MARGIN * 2, note_h,
        [(t, FONT_BODY, SIZE_SMALL, C_SUB, False, PP_ALIGN.LEFT) for t in notes],
        line_spacing=1.3)

    area_top = CONTENT_TOP + Inches(0.05)
    area_bottom = note_y - GAP_BLOCK

    # ---- Left cols 0-6: price hero + copy + zero-upfront sub-metric ----
    lx = grid_x(0)
    lw = grid_w(7)
    hero_h = Inches(1.60)
    copy_h = Inches(0.65)
    sub_h = Inches(0.85)
    ys = vstack(area_top, area_bottom, [hero_h, copy_h, sub_h])

    add_metric_hero(slide, lx, ys[0], lw, hero_h,
                    price_str, "円/kWh", f"PPA単価（{years}年固定）")

    add_multiline_textbox(
        slide, lx, ys[1], lw, copy_h,
        [("PPA（Power Purchase Agreement）は、太陽光発電システムによる電力供給契約です。",
          FONT_BODY, SIZE_BODY, C_DARK, False, PP_ALIGN.LEFT),
         ("ご使用いただいた電力量のみをお支払いいただくため、設備投資のご負担はありません。",
          FONT_BODY, SIZE_BODY, C_DARK, False, PP_ALIGN.LEFT)],
        line_spacing=1.35)

    add_textbox(slide, lx, ys[2], lw, Inches(0.20),
                "初期ご負担金額",
                font_size_pt=SIZE_CAPTION, font_color=C_SUB, bold=True)
    add_number_unit(slide, lx, ys[2] + Inches(0.22), lw, Inches(0.50),
                    "0", "円")

    # ---- Right cols 7-11: comparison table + merit lines ----
    rx = grid_x(7)
    rw = grid_w(5)

    reduction_str = "—"
    if cur_price and ppa_price and cur_price > 0:
        reduction_str = f"{(cur_price - ppa_price) / cur_price * 100:.1f}%"
    cmp_rows = [
        ["区分", "単価（円/kWh）"],
        ["現行平均単価", f"{cur_price:.2f}" if cur_price is not None else "—"],
        ["PPA単価", price_str],
        ["削減率", reduction_str],
    ]
    cmp_h = Inches(0.34) + TABLE_ROW_H * len(cmp_rows)

    merit_item_h = Inches(0.68)
    merit_h = Inches(0.34) + merit_item_h * len(MERITS)

    rys = vstack(area_top, area_bottom, [cmp_h, merit_h])

    add_section_header(slide, rx, rys[0], rw, "現行単価との比較")
    add_table(slide, rx, rys[0] + Inches(0.34), rw, cmp_rows,
              [rw - Inches(1.9), Inches(1.9)],
              total_row=2)  # highlight the PPA row

    add_section_header(slide, rx, rys[1], rw, "PPAのメリット")
    my = rys[1] + Inches(0.34)
    sq = Inches(0.07)
    indent = Inches(0.16)
    for title, desc in MERITS:
        add_rect(slide, rx, my + Inches(0.06), sq, sq, C_ORANGE)
        add_textbox(slide, rx + indent, my, rw - indent, Inches(0.20),
                    title,
                    font_size_pt=SIZE_BODY, font_color=C_DARK, bold=True)
        add_textbox(slide, rx + indent, my + Inches(0.21),
                    rw - indent, Inches(0.40),
                    desc,
                    font_size_pt=SIZE_CAPTION, font_color=C_SUB,
                    line_spacing=1.3)
        my += merit_item_h

    add_footer(slide)
