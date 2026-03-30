"""
utils.py
Shared python-pptx utility functions for all slide generators.
Design spec: sales_sheet_design_guide.md (A4 landscape, DIC160 orange #E8490F)
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Design constants
# ---------------------------------------------------------------------------

# A4 landscape in inches
SLIDE_W = Inches(11.69)   # A4 landscape width
SLIDE_H = Inches(8.27)    # A4 landscape height

# Margins
MARGIN = Inches(0.35)

# Color palette
C_ORANGE  = RGBColor(0xE8, 0x49, 0x0F)   # DIC160 #E8490F
C_DARK    = RGBColor(0x33, 0x33, 0x33)   # #333333 main text
C_SUB     = RGBColor(0x66, 0x66, 0x66)   # #666666 sub text
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)   # #FFFFFF
C_LIGHT_ORANGE = RGBColor(0xFE, 0xF0, 0xEB)  # #FEF0EB card bg
C_LIGHT_GRAY   = RGBColor(0xF5, 0xF5, 0xF5)  # #F5F5F5 card bg
C_BORDER  = RGBColor(0xE0, 0xE0, 0xE0)   # #E0E0E0 table border
C_CARD_BG = RGBColor(0x1A, 0x1A, 0x1A)   # #1A1A1A dark card
# PDF reference colors (navy/teal scheme from Excel output)
C_NAVY    = RGBColor(0x00, 0x20, 0x60)   # #002060 primary dark
C_TEAL    = RGBColor(0x46, 0xAA, 0xC5)   # #46AAC5 accent teal
C_LIGHT_TEAL = RGBColor(0x76, 0xC5, 0xD8) # #76C5D8 price box bg
C_LIGHT_CYAN = RGBColor(0xCC, 0xFF, 0xFF) # #CCFFFF table header
C_RED     = RGBColor(0xFF, 0x00, 0x00)   # #FF0000 negative values

# Fonts (Meiryo unified)
FONT_BLACK  = "Meiryo"              # titles, accent numbers — use with bold=True
FONT_BODY   = "Meiryo"              # body text
_BOLD_FONT  = True                  # flag: FONT_BLACK callers should set bold=True

# Header bar height
HEADER_H = Inches(0.9)

# Footer height
FOOTER_H = Inches(0.28)

# Content area
CONTENT_TOP = HEADER_H + Inches(0.12)
CONTENT_H   = SLIDE_H - HEADER_H - FOOTER_H - Inches(0.24)


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------

def load_template(template_path: Path) -> Presentation:
    """Open the base PPTX template (contains logo + orange header bar layout)."""
    return Presentation(str(template_path))


def add_blank_slide(prs: Presentation, layout_index: int = 6) -> object:
    """Add a blank slide using the specified layout index."""
    layout = prs.slide_layouts[layout_index]
    return prs.slides.add_slide(layout)


# ---------------------------------------------------------------------------
# Shape primitives
# ---------------------------------------------------------------------------

def add_rect(slide, x, y, w, h, fill_color: RGBColor, border_color: Optional[RGBColor] = None, border_pt: float = 0.0):
    """Add a filled rectangle with optional border."""
    from pptx.util import Emu
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        x, y, w, h
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape


def add_rounded_rect(slide, x, y, w, h, fill_color: RGBColor, radius_pt: float = 6.0, border_color: Optional[RGBColor] = None, border_pt: float = 0.0):
    """Add a rounded rectangle."""
    from pptx.util import Pt as _Pt
    shape = slide.shapes.add_shape(
        5,  # MSO_SHAPE_TYPE.ROUNDED_RECTANGLE
        x, y, w, h
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    # Set corner radius via XML
    sp_pr = shape.element.find(qn("p:spPr"))
    if sp_pr is not None:
        prstgeom = sp_pr.find(qn("a:prstGeom"))
        if prstgeom is not None:
            av_lst = prstgeom.find(qn("a:avLst"))
            if av_lst is not None:
                for gd in av_lst.findall(qn("a:gd")):
                    if gd.get("name") == "adj":
                        # radius as fraction of min(w,h); clamp to reasonable value
                        from pptx.util import Emu
                        min_dim = min(w, h)
                        frac = min(radius_pt * 12700 / min_dim, 50000)
                        gd.set("fmla", f"val {int(frac)}")
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = _Pt(border_pt)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, x, y, w, h, text: str,
                font_name: str = FONT_BODY,
                font_size_pt: float = 11,
                font_color: RGBColor = C_DARK,
                bold: bool = False,
                align: PP_ALIGN = PP_ALIGN.LEFT,
                word_wrap: bool = True) -> object:
    """Add a simple single-run textbox."""
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    tf.auto_size = None
    # Remove default margins
    from pptx.util import Pt
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = font_color
    run.font.bold = bold
    return txBox


def add_multiline_textbox(slide, x, y, w, h, lines: list[tuple],
                          word_wrap: bool = True) -> object:
    """
    Add a textbox with multiple paragraphs/runs.
    lines: list of (text, font_name, font_size_pt, color, bold, align)
    """
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    from pptx.util import Pt
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)

    for i, (text, font_name, size, color, bold, align) in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold
    return txBox


def add_image_contain(slide, x, y, w, h, image_path: Path) -> object:
    """Add an image maintaining aspect ratio (contain mode)."""
    from PIL import Image as PILImage
    img = PILImage.open(str(image_path))
    img_w, img_h = img.size
    aspect = img_w / img_h

    box_aspect = w / h
    if aspect > box_aspect:
        # wider than box: fit width
        render_w = w
        render_h = w / aspect
        render_x = x
        render_y = y + (h - render_h) / 2
    else:
        # taller than box: fit height
        render_h = h
        render_w = h * aspect
        render_x = x + (w - render_w) / 2
        render_y = y

    return slide.shapes.add_picture(str(image_path), render_x, render_y, render_w, render_h)


def add_line(slide, x1, y1, x2, y2, color: RGBColor, width_pt: float = 1.0):
    """Add a straight line."""
    from pptx.util import Pt
    connector = slide.shapes.add_connector(1, x1, y1, x2, y2)
    connector.line.color.rgb = color
    connector.line.width = Pt(width_pt)
    return connector


# ---------------------------------------------------------------------------
# Composite components
# ---------------------------------------------------------------------------

def _add_gradient_rect(slide, x, y, w, h, color_top: RGBColor, color_bottom: RGBColor, angle: int = 90):
    """Add a rectangle with linear gradient fill (top→bottom by default)."""
    from lxml import etree

    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE
        int(x), int(y), int(w), int(h),
    )
    shape.line.fill.background()  # no border

    # Access spPr and insert gradFill right after prstGeom (before a:ln)
    sp_pr = shape._element.find(qn("p:spPr"))
    if sp_pr is None:
        sp_pr = shape._element.find(qn("a:spPr"))
    if sp_pr is None:
        sp_pr = etree.SubElement(shape._element, qn("p:spPr"))

    # Remove existing fills
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for child in list(sp_pr.findall(qn(tag))):
            sp_pr.remove(child)

    # Build gradFill element
    gf = etree.Element(qn("a:gradFill"))
    gsl = etree.SubElement(gf, qn("a:gsLst"))
    gs1 = etree.SubElement(gsl, qn("a:gs"), attrib={"pos": "0"})
    etree.SubElement(gs1, qn("a:srgbClr"), attrib={"val": str(color_top)})
    gs2 = etree.SubElement(gsl, qn("a:gs"), attrib={"pos": "100000"})
    etree.SubElement(gs2, qn("a:srgbClr"), attrib={"val": str(color_bottom)})
    etree.SubElement(gf, qn("a:lin"), attrib={"ang": str(angle * 60000), "scaled": "1"})

    # Insert gradFill after prstGeom (correct OOXML element order)
    prst_geom = sp_pr.find(qn("a:prstGeom"))
    if prst_geom is not None:
        idx = list(sp_pr).index(prst_geom) + 1
        sp_pr.insert(idx, gf)
    else:
        sp_pr.insert(0, gf)

    # Remove theme style override (p:style) which can override our fill
    for style_el in list(shape._element.findall(qn("p:style"))):
        shape._element.remove(style_el)


def add_header_bar(slide, title: str, logo_path: Optional[Path] = None):
    """
    Orange gradient header bar (dark at top → light at bottom) with white title.
    Single shape with gradient fill.
    """
    # Gradient: #E8490F (top) → #F0935A (bottom) - stays warm enough for white text
    _add_gradient_rect(slide, 0, 0, SLIDE_W, HEADER_H,
                       C_ORANGE, RGBColor(0xF0, 0x93, 0x5A), angle=90)

    # Logo (left side)
    if logo_path and logo_path.exists():
        logo_h = Inches(0.55)
        logo_w = Inches(1.6)
        add_image_contain(slide,
                          Inches(0.3), (HEADER_H - logo_h) / 2,
                          logo_w, logo_h,
                          logo_path)
        text_x = Inches(2.1)
    else:
        text_x = Inches(0.35)

    # Title text
    add_textbox(slide,
                text_x, Inches(0.15),
                SLIDE_W - text_x - Inches(0.3), Inches(0.8),
                title,
                font_name=FONT_BLACK,
                font_size_pt=28,
                font_color=C_WHITE,
                bold=True,
                align=PP_ALIGN.LEFT)


def add_footer(slide, text: str = "株式会社オルテナジー  |  https://altenergy.co.jp/"):
    """Add footer bar at the bottom of the slide."""
    y = SLIDE_H - FOOTER_H
    add_rect(slide, 0, y, SLIDE_W, FOOTER_H, C_DARK)
    add_textbox(slide,
                Inches(0.35), y + Inches(0.03),
                SLIDE_W - Inches(0.7), FOOTER_H - Inches(0.06),
                text,
                font_name=FONT_BODY,
                font_size_pt=8,
                font_color=C_WHITE,
                align=PP_ALIGN.CENTER)


def add_section_header(slide, x, y, w, text: str, font_size_pt: float = 14):
    """Section header: orange left border + bold text."""
    border_w = Inches(0.05)
    add_rect(slide, x, y, border_w, Inches(0.28), C_ORANGE)
    add_textbox(slide,
                x + border_w + Inches(0.08), y,
                w - border_w - Inches(0.08), Inches(0.3),
                text,
                font_name=FONT_BODY,
                font_size_pt=font_size_pt,
                font_color=C_DARK,
                bold=True)


def add_kpi_card(slide, x, y, w, h,
                 number: str, unit: str, label: str,
                 bg_color: RGBColor = C_LIGHT_ORANGE,
                 number_size_pt: float = 36):
    """KPI card with large number, unit, and label."""
    add_rounded_rect(slide, x, y, w, h, bg_color)
    # Number (large)
    add_textbox(slide,
                x + Inches(0.1), y + Inches(0.12),
                w - Inches(0.2), Inches(0.5),
                number,
                font_name=FONT_BLACK,
                font_size_pt=number_size_pt,
                font_color=C_ORANGE,
                bold=True,
                align=PP_ALIGN.CENTER)
    # Unit (small, right of number area)
    add_textbox(slide,
                x + Inches(0.1), y + Inches(0.5),
                w - Inches(0.2), Inches(0.22),
                unit,
                font_name=FONT_BODY,
                font_size_pt=9,
                font_color=C_SUB,
                align=PP_ALIGN.CENTER)
    # Label
    add_textbox(slide,
                x + Inches(0.08), y + h - Inches(0.3),
                w - Inches(0.16), Inches(0.28),
                label,
                font_name=FONT_BODY,
                font_size_pt=9,
                font_color=C_DARK,
                bold=True,
                align=PP_ALIGN.CENTER)


def add_table(slide, x, y, w, rows_data: list[list],
              col_widths: list,
              header_bg: RGBColor = C_ORANGE,
              row_bg_even: RGBColor = C_WHITE,
              row_bg_odd: RGBColor = C_LIGHT_GRAY,
              font_size_pt: float = 10):
    """
    Add a simple styled table.
    rows_data[0] = header row, rows_data[1:] = data rows.
    col_widths: list of Inches values (should sum to w).
    """
    from pptx.util import Pt
    n_rows = len(rows_data)
    n_cols = len(rows_data[0])
    row_h = Inches(0.28)
    tbl = slide.shapes.add_table(n_rows, n_cols, x, y, w, row_h * n_rows).table

    # Column widths
    for c, cw in enumerate(col_widths):
        tbl.columns[c].width = int(cw)

    for r, row in enumerate(rows_data):
        for c, cell_text in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = str(cell_text) if cell_text is not None else ""

            # Font
            for para in cell.text_frame.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.name = FONT_BODY
                    run.font.size = Pt(font_size_pt)
                    run.font.bold = (r == 0)
                    run.font.color.rgb = C_WHITE if r == 0 else C_DARK

            # Background
            if r == 0:
                _set_cell_bg(cell, header_bg)
            elif r % 2 == 0:
                _set_cell_bg(cell, row_bg_even)
            else:
                _set_cell_bg(cell, row_bg_odd)

    return tbl


def _set_cell_bg(cell, color: RGBColor):
    """Set table cell background color."""
    from lxml import etree
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    solidFill = etree.SubElement(tcPr, qn("a:solidFill"))
    srgbClr = etree.SubElement(solidFill, qn("a:srgbClr"))
    srgbClr.set("val", str(color))


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def fmt_yen(value, unit: str = "円") -> str:
    """Format a number as Japanese yen string."""
    if value is None:
        return "—"
    try:
        v = float(value)
        if abs(v) >= 1_0000_0000:
            return f"{v / 1_0000_0000:.1f}億{unit}"
        if abs(v) >= 10_000:
            return f"{v / 10_000:.0f}万{unit}"
        return f"{v:,.0f}{unit}"
    except (TypeError, ValueError):
        return str(value)


def fmt_num(value, decimals: int = 1, suffix: str = "") -> str:
    """Format a number with given decimal places."""
    if value is None:
        return "—"
    try:
        return f"{float(value):,.{decimals}f}{suffix}"
    except (TypeError, ValueError):
        return str(value)
