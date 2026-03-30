"""
ep7.py - 導入スケジュール（EPC）

Shows the installation timeline for EPC solar system.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    FONT_BLACK, FONT_BODY, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header,
)

TITLE = "導入スケジュール（EPC）"

SCHEDULE_PHASES = [
    {
        "phase": "STEP 1",
        "title": "ご契約・設計",
        "duration": "1〜2ヶ月",
        "tasks": [
            "売買契約の締結",
            "現地調査・詳細設計",
            "電力会社への接続申請",
            "補助金申請（該当する場合）",
        ],
    },
    {
        "phase": "STEP 2",
        "title": "機器調達・施工準備",
        "duration": "1〜2ヶ月",
        "tasks": [
            "太陽光パネル・PCS等の発注",
            "施工計画の策定",
            "必要な許認可の取得",
        ],
    },
    {
        "phase": "STEP 3",
        "title": "施工・設置",
        "duration": "2〜4週間",
        "tasks": [
            "架台・パネルの設置工事",
            "電気工事（PCS・配線）",
            "検査・試運転",
        ],
    },
    {
        "phase": "STEP 4",
        "title": "運転開始",
        "duration": "—",
        "tasks": [
            "系統連系・運転開始",
            "発電モニタリング開始",
            "お客様への引渡し完了",
        ],
    },
]


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """
    Render EP7 (installation schedule for EPC) onto an already-added blank slide.

    data keys used: (none - static content)
    """
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP + Inches(0.08)

    # ---- Lead text ----
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.5),
                "ご契約から運転開始まで、約3〜5ヶ月が目安です。\n"
                "補助金申請がある場合は、採択スケジュールに合わせて調整いたします。",
                font_name=FONT_BODY, font_size_pt=10,
                font_color=C_DARK)
    y += Inches(0.65)

    # ---- Schedule phases ----
    phase_w = SLIDE_W - MARGIN * 2
    for phase_data in SCHEDULE_PHASES:
        phase_h = Inches(0.35) + Inches(0.25) * len(phase_data["tasks"]) + Inches(0.15)

        # Phase card background
        add_rounded_rect(slide, MARGIN, y, phase_w, phase_h, C_LIGHT_GRAY)

        # Orange left accent
        add_rect(slide, MARGIN, y, Inches(0.07), phase_h, C_ORANGE)

        # Phase label + title + duration
        add_textbox(slide,
                    MARGIN + Inches(0.16), y + Inches(0.06),
                    Inches(1.2), Inches(0.25),
                    phase_data["phase"],
                    font_name=FONT_BLACK, font_size_pt=11,
                    font_color=C_ORANGE, bold=True)
        add_textbox(slide,
                    MARGIN + Inches(1.4), y + Inches(0.06),
                    Inches(3.5), Inches(0.25),
                    phase_data["title"],
                    font_name=FONT_BODY, font_size_pt=12,
                    font_color=C_DARK, bold=True)
        add_textbox(slide,
                    SLIDE_W - MARGIN - Inches(2.0), y + Inches(0.06),
                    Inches(2.0), Inches(0.25),
                    phase_data["duration"],
                    font_name=FONT_BODY, font_size_pt=10,
                    font_color=C_SUB,
                    align=PP_ALIGN.RIGHT)

        # Task list
        task_y = y + Inches(0.38)
        for task in phase_data["tasks"]:
            add_textbox(slide,
                        MARGIN + Inches(0.35), task_y,
                        phase_w - Inches(0.5), Inches(0.22),
                        f"・{task}",
                        font_name=FONT_BODY, font_size_pt=9,
                        font_color=C_SUB)
            task_y += Inches(0.25)

        y += phase_h + Inches(0.12)

    # ---- Note ----
    add_textbox(slide, MARGIN, y,
                SLIDE_W - MARGIN * 2, Inches(0.25),
                "※ スケジュールは目安です。現地条件や申請状況により変動する場合があります。",
                font_name=FONT_BODY, font_size_pt=8,
                font_color=C_SUB)

    add_footer(slide)
