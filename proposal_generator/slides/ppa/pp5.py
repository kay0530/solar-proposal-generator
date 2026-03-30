"""
pp5.py - 設備レイアウト・積載荷重 (Equipment Layout & Load Calculation)

Left side: uploaded layout image (roof layout screenshot)
Right side: load calculation summary table (from 積載荷重計算表 Excel)
Gracefully degrades when either image or load data is missing.
"""
from __future__ import annotations
from pathlib import Path
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from proposal_generator.utils import (
    CONTENT_TOP, C_DARK, C_LIGHT_GRAY, C_LIGHT_ORANGE, C_ORANGE, C_SUB, C_WHITE,
    C_BORDER,
    FONT_BLACK, FONT_BODY, HEADER_H, MARGIN, SLIDE_H, SLIDE_W,
    add_footer, add_header_bar, add_rect, add_rounded_rect, add_textbox,
    add_section_header, add_image_contain, _set_cell_bg, fmt_num,
)

TITLE = "設備レイアウト・積載荷重"


def generate(slide, data: dict, logo_path: Path = None) -> None:
    """Render PP5 (equipment layout & load calculation)."""
    add_header_bar(slide, TITLE, logo_path)

    y = CONTENT_TOP
    has_image = bool(data.get("layout_image_path"))
    has_load = bool(data.get("load_calc"))

    if has_image and has_load:
        # Two-column layout: image left, table right
        _render_two_column(slide, data, y)
    elif has_image:
        # Image only: centered, with basic system info
        _render_image_only(slide, data, y)
    elif has_load:
        # Table only: centered load calc summary
        _render_load_only(slide, data, y)
    else:
        # Fallback: show basic system info from customer_data
        _render_fallback(slide, data, y)

    add_footer(slide)


def _render_two_column(slide, data: dict, y) -> None:
    """Two-column layout: layout image (left) + load calc table (right)."""
    col_gap = Inches(0.3)
    left_w = (SLIDE_W - MARGIN * 2 - col_gap) * 0.55
    right_w = (SLIDE_W - MARGIN * 2 - col_gap) * 0.45
    left_x = MARGIN
    right_x = MARGIN + left_w + col_gap

    # -- Left: Layout image --
    add_section_header(slide, left_x, y, left_w, "◆ 設備レイアウト図", font_size_pt=11)
    img_y = y + Inches(0.35)
    img_h = SLIDE_H - img_y - Inches(0.6)

    img_path = Path(data["layout_image_path"])
    if img_path.exists():
        add_image_contain(slide, left_x, img_y, left_w, img_h, img_path)
    else:
        # Placeholder box
        add_rounded_rect(slide, left_x, img_y, left_w, img_h, C_LIGHT_GRAY)
        add_textbox(slide, left_x, img_y + img_h / 2 - Inches(0.15),
                    left_w, Inches(0.3),
                    "レイアウト画像なし",
                    font_name=FONT_BODY, font_size_pt=12,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)

    # -- Right: Load calculation summary --
    add_section_header(slide, right_x, y, right_w, "◆ 積載荷重計算", font_size_pt=11)
    table_y = y + Inches(0.35)
    _render_load_table(slide, data, right_x, table_y, right_w)


def _render_image_only(slide, data: dict, y) -> None:
    """Full-width layout image with basic system info below."""
    full_w = SLIDE_W - MARGIN * 2

    # System info bar
    _render_system_info_bar(slide, data, MARGIN, y, full_w)
    y += Inches(0.7)

    add_section_header(slide, MARGIN, y, full_w, "◆ 設備レイアウト図", font_size_pt=12)
    img_y = y + Inches(0.35)
    img_h = SLIDE_H - img_y - Inches(0.6)

    img_path = Path(data["layout_image_path"])
    if img_path.exists():
        add_image_contain(slide, MARGIN, img_y, full_w, img_h, img_path)


def _render_load_only(slide, data: dict, y) -> None:
    """Centered load calc summary (no image)."""
    full_w = SLIDE_W - MARGIN * 2

    # System info bar
    _render_system_info_bar(slide, data, MARGIN, y, full_w)
    y += Inches(0.7)

    add_section_header(slide, MARGIN, y, full_w, "◆ 積載荷重計算結果", font_size_pt=12)
    table_y = y + Inches(0.35)

    # Center the table with max width
    table_w = min(full_w, Inches(7.0))
    table_x = MARGIN + (full_w - table_w) / 2
    _render_load_table(slide, data, table_x, table_y, table_w)


def _render_fallback(slide, data: dict, y) -> None:
    """Fallback: show basic system info when no image or load data."""
    full_w = SLIDE_W - MARGIN * 2

    _render_system_info_bar(slide, data, MARGIN, y, full_w)
    y += Inches(0.9)

    add_textbox(slide, MARGIN, y, full_w, Inches(0.5),
                "レイアウト画像または積載荷重計算表をアップロードしてください。",
                font_name=FONT_BODY, font_size_pt=14,
                font_color=C_SUB, align=PP_ALIGN.CENTER)


def _render_system_info_bar(slide, data: dict, x, y, w) -> None:
    """Render a compact system info bar with key specs."""
    panel_kw = data.get("panel_total_kw", data.get("system_capacity_kw", 0)) or 0
    panel_count = data.get("panel_total_count", data.get("panel_count", 0)) or 0
    pcs_kw = data.get("pcs_total_kw", data.get("pcs_output_kw", 0)) or 0
    battery_kwh = data.get("battery_total_kwh", data.get("battery_kwh", 0)) or 0

    bar_h = Inches(0.55)
    add_rounded_rect(slide, x, y, w, bar_h, C_LIGHT_ORANGE)

    items = []
    if panel_kw:
        items.append(f"PV出力: {panel_kw:,.2f} kW")
    if panel_count:
        items.append(f"パネル枚数: {panel_count:,}枚")
    if pcs_kw:
        items.append(f"PCS出力: {pcs_kw:,.1f} kW")
    if battery_kwh:
        items.append(f"蓄電池: {battery_kwh:,.1f} kWh")

    info_text = "　｜　".join(items) if items else "設備情報未入力"

    add_textbox(slide, x + Inches(0.15), y + Inches(0.1),
                w - Inches(0.3), Inches(0.35),
                info_text,
                font_name=FONT_BLACK, font_size_pt=13,
                font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)


def _render_load_table(slide, data: dict, x, y, w) -> None:
    """Render load calculation summary as a styled table."""
    lc = data.get("load_calc", {})
    if not lc:
        return

    # KPI cards at top
    kpi_items = [
        (f"{lc.get('load_per_roof_area', 0):.1f}", "kg/m\u00b2", "対屋根面積"),
        (f"{lc.get('total_weight_kg', 0):,.0f}", "kg", "総重量"),
        (f"{lc.get('panel_count', 0):,}", "枚", "パネル枚数"),
    ]

    kpi_w = (w - Inches(0.2)) / 3
    kpi_h = Inches(0.75)
    for i, (val, unit, label) in enumerate(kpi_items):
        kx = x + i * (kpi_w + Inches(0.1))
        add_rounded_rect(slide, kx, y, kpi_w, kpi_h, C_LIGHT_ORANGE)
        add_textbox(slide, kx, y + Inches(0.05), kpi_w, Inches(0.30),
                    val, font_name=FONT_BLACK, font_size_pt=18,
                    font_color=C_ORANGE, bold=True, align=PP_ALIGN.CENTER)
        add_textbox(slide, kx, y + Inches(0.35), kpi_w, Inches(0.18),
                    unit, font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_SUB, align=PP_ALIGN.CENTER)
        add_textbox(slide, kx, y + Inches(0.53), kpi_w, Inches(0.18),
                    label, font_name=FONT_BODY, font_size_pt=8,
                    font_color=C_DARK, bold=True, align=PP_ALIGN.CENTER)

    y += kpi_h + Inches(0.2)

    # Detail table
    rows = [
        ["項目", "値"],
        ["PV型番", lc.get("panel_model", "—")],
        ["パネル枚数", f"{lc.get('panel_count', 0):,} 枚"],
        ["パネル単体重量", f"{lc.get('panel_unit_weight_kg', 0):.1f} kg"],
        ["パネル重量 (計)", f"{lc.get('panel_weight_kg', 0):,.1f} kg"],
        ["架台重量", f"{lc.get('frame_weight_kg', 0):,.1f} kg"],
        ["配線重量", f"{lc.get('wiring_weight_kg', 0):,.1f} kg"],
        ["総重量", f"{lc.get('total_weight_kg', 0):,.1f} kg"],
        ["パネル面積", f"{lc.get('panel_area_m2', 0):,.1f} m\u00b2"],
        ["屋根面積", f"{lc.get('roof_area_m2', 0):,.1f} m\u00b2"],
        ["積載荷重 (対パネル面積)", f"{lc.get('load_per_panel_area', 0):.2f} kg/m\u00b2"],
        ["積載荷重 (対屋根面積)", f"{lc.get('load_per_roof_area', 0):.2f} kg/m\u00b2"],
    ]

    # Filter out rows where value is "0 kg" or similar empty
    n_rows = len(rows)
    n_cols = 2
    label_w = w * 0.55
    value_w = w * 0.45
    row_h = Inches(0.26)

    tbl = slide.shapes.add_table(n_rows, n_cols, int(x), int(y), int(w), int(row_h * n_rows)).table
    tbl.columns[0].width = int(label_w)
    tbl.columns[1].width = int(value_w)

    for r, row_data in enumerate(rows):
        for c, cell_text in enumerate(row_data):
            cell = tbl.cell(r, c)
            cell.text = str(cell_text)
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
                for run in para.runs:
                    run.font.name = FONT_BODY
                    run.font.size = Pt(9)
                    run.font.bold = (r == 0)
                    run.font.color.rgb = C_WHITE if r == 0 else C_DARK

            if r == 0:
                _set_cell_bg(cell, C_ORANGE)
            elif r % 2 == 0:
                _set_cell_bg(cell, C_WHITE)
            else:
                _set_cell_bg(cell, C_LIGHT_GRAY)
