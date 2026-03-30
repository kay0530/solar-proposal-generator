"""
pp8a.py - ご契約条件に関して（サマリー）

PDF P13: Contract conditions summary with two-column layout.
Left: 設備/期間/終了後/保険/保守
Right: 中途解約/所有権移転/建物改修
Bottom: 違約金テーブル
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_LIGHT_GRAY, FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, add_table, fmt_yen, fmt_num,
)

TITLE = "ご契約条件に関して（サマリー）"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP

    add_textbox(slide, MARGIN, y, SLIDE_W - MARGIN * 2, Inches(0.25),
                "オンサイトPPAサービス契約に関して",
                font_name=FONT_BLACK, font_size_pt=12,
                font_color=C_ORANGE, bold=True)
    y += Inches(0.30)

    # ---- Two-column layout ----
    col_w = (SLIDE_W - MARGIN * 2 - Inches(0.3)) / 2
    left_x = MARGIN
    right_x = MARGIN + col_w + Inches(0.3)

    years = data.get("contract_years", 20)

    # Left column items
    left_items = [
        ("【設備について】",
         "発電設備の施工及びメンテナンスは当社が実施させていただきます。"
         "施工及びメンテナンスの実施のための場所の提供、"
         "その他、契約者にご協力いただくことがあります。"),
        ("【契約の期間に関して】",
         f"契約期間は{years}年になります。"
         "契約期間満了後は新たに更新契約を締結することにより契約の更新が可能です。"),
        ("【契約期間終了後の手続き】",
         "契約期間終了後、契約が更新されない場合には太陽光パネル及びPCS等の"
         "主要機器に関しては、弊社の費用負担で撤去いたします。"),
        ("【保険に関して】",
         "発電設備の火災保険及び賠償責任保険は当社が加入いたします。"
         "当社帰責事由により貴社及び第三者に損害を与えた場合には、"
         "当社が加入する保険により担保される範囲で補償をいたします。"),
        ("【発電設備の保守】",
         "発電設備の定期点検及び故障が発生した場合の補修作業は、"
         "当社の費用負担により実施するものとします。"),
    ]

    # Right column items
    right_items = [
        ("【期間中の貴社事由による解約】",
         "契約期間中に貴社の都合により解約をされた場合には、"
         "予め設定した違約金が発生いたします。"),
        ("【建物等の所有権の移転】",
         "建物等の所有権が移転する場合には、新たなる所有者が"
         "オンサイトPPAサービス契約を承継する場合などは"
         "ペナルティが発生しないことといたします。"),
        ("【建物等の改修に関して】",
         "建物等の改装のために当該設備を一時的に撤去が必要である場合には、"
         "期間及び改修方法を協議の上、貴社の負担で一時的な撤去を行います。"),
    ]

    ly = y
    for title, body in left_items:
        add_textbox(slide, left_x, ly, col_w, Inches(0.20),
                    title,
                    font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_ORANGE, bold=True)
        ly += Inches(0.18)
        add_textbox(slide, left_x, ly, col_w, Inches(0.45),
                    body,
                    font_name=FONT_BODY, font_size_pt=7, font_color=C_DARK)
        ly += Inches(0.48)

    ry = y
    for title, body in right_items:
        add_textbox(slide, right_x, ry, col_w, Inches(0.20),
                    title,
                    font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_ORANGE, bold=True)
        ry += Inches(0.18)
        add_textbox(slide, right_x, ry, col_w, Inches(0.50),
                    body,
                    font_name=FONT_BODY, font_size_pt=7, font_color=C_DARK)
        ry += Inches(0.53)

    # ---- Penalty table ----
    table_y = max(ly, ry) + Inches(0.1)

    capacity = data.get("system_capacity_kw", 0)
    selling_price = data.get("selling_price", 0)
    proposal_type = data.get("proposal_type", "ppa")
    depreciation_years = years

    add_section_header(slide, MARGIN, table_y, SLIDE_W - MARGIN * 2,
                       "中途解約による違約金に係る設備価額（税抜）", font_size_pt=9)
    table_y += Inches(0.22)

    # Info line
    depr_limit = int(selling_price / depreciation_years) if depreciation_years > 0 else 0
    equip_label = "設備価額（PPA事業者負担）" if proposal_type == "ppa" else "設備価額"
    info = (
        f"設備容量: {capacity:.2f}kW　"
        f"{equip_label}: {fmt_yen(selling_price)}　"
        f"償却年数: {depreciation_years}年　"
        f"償却限度額: {fmt_yen(depr_limit)}"
    )
    add_textbox(slide, MARGIN, table_y, SLIDE_W - MARGIN * 2, Inches(0.18),
                info, font_name=FONT_BODY, font_size_pt=7, font_color=C_DARK)
    table_y += Inches(0.22)

    # Penalty table: milestones
    milestones = [1, 3, 5, 7, 10, 15, 20]
    header = ["経過年数"] + [f"{yr}年" for yr in milestones if yr <= years]
    values = ["違約金"]
    for yr in milestones:
        if yr > years:
            break
        remaining = max(selling_price - depr_limit * yr, depr_limit)
        values.append(fmt_yen(remaining))

    if len(header) > 1:
        table_w = SLIDE_W - MARGIN * 2
        col_w_each = table_w / len(header)
        col_widths = [col_w_each] * len(header)
        add_table(slide, MARGIN, table_y, table_w, [header, values],
                  col_widths, font_size_pt=7)
        table_y += Inches(0.45)

    # PPA note below penalty table
    if proposal_type == "ppa":
        add_textbox(slide, MARGIN, table_y, SLIDE_W - MARGIN * 2, Inches(0.18),
                    "※ 上記は設備の残存簿価に基づく概算違約金です",
                    font_name=FONT_BODY, font_size_pt=7, font_color=C_SUB)

    add_footer(slide)
