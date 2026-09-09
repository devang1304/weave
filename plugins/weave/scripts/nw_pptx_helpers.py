#!/usr/bin/env python3
"""
nw_pptx_helpers.py - shared runtime helpers for Netwoven decks built on
NW_Presentation_Base_2026.pptx ("Netwoven Default Theme", 39 layouts).

Used by new_deck_pptx.py (create / finalize / demo) and, through the lazy
import in validate_deliverable.py, by the validator (validate_deck).

Runtime requirements: Python 3.9+, python-pptx, lxml. Nothing else.

Placeholder idx values in LAYOUTS were verified against the base with
python-pptx (placeholder_format.idx). python-pptx reports title AND ctrTitle
placeholders as idx 0. The base's layouts use non-standard idx values
(10..38), so always address placeholders through ph(slide, idx) and the
LAYOUTS map, never through positional indexing.
"""
import datetime as _dt
import os
import re
import zipfile

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.shapes.group import GroupShape
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Layout map (name -> placeholder idx roles), verified on the base
# ---------------------------------------------------------------------------

LAYOUTS = {
    # cover: ctrTitle 48pt cyan, subtitle 18pt, two portrait picture slots
    "Title Slide for Verticals": {"title": 0, "subtitle": 1, "pics": [15, 16]},
    "Title and Content": {"title": 0, "body": 1},
    "Section Header": {"title": 0, "body": 1, "pic": 32},
    "Two Content": {"title": 0, "left": 1, "right": 2},
    "Three Column": {"title": 0, "columns": [13, 14, 15]},
    "Comparison": {"title": 0, "heads": [1, 3], "bodies": [2, 4]},
    "Title Only": {"title": 0},
    "Title Only (Centered)": {"title": 0},
    # (number, label) pairs, left to right; numerals are 96pt cyan
    "Three Stat": {"stats": [(10, 11), (15, 14), (17, 16)]},
    # idx 16 = 72pt statement, idx 14 = 21pt eyebrow (verified in layout XML)
    "Big Statement with Illustration": {"statement": 16, "eyebrow": 14, "pic": 13},
    "Three Image Horizontal": {"title": 0, "captions": [13, 14, 15], "pics": [16, 17, 18]},
    "Six Icon": {"title": 0, "cells": [13, 14, 15, 16, 17, 18],
                 "icons": [19, 20, 21, 25, 26, 27]},
    "Grid Layout 1-2": {"title": 0, "eyebrows": [14, 15, 17], "cells": [12, 13, 16]},
    "Grid Layout 2-1": {"title": 0, "eyebrows": [14, 15, 17], "cells": [12, 13, 16]},
    "Grid Layout 2-2": {"title": 0, "eyebrows": [14, 15, 18, 19], "cells": [12, 13, 16, 17]},
    "Grid Layout 1-2-2": {"title": 0, "eyebrows": [14, 15, 17, 19, 21],
                          "cells": [12, 13, 16, 18, 20]},
    "Grid Layout 2-2-2": {"title": 0, "eyebrows": [14, 15, 18, 19, 22, 23],
                          "cells": [12, 13, 16, 17, 20, 21]},
    "Blank": {},
    "Closing Slide": {"pic": 25},
}

# Layouts that must never carry a generated slide.
# "Title Slide" has five picture slots and a subtitle but NO title placeholder.
BROKEN_LAYOUTS = {"Title Slide"}

# ---------------------------------------------------------------------------
# Brand tokens (see shared/references/brand-tokens.md)
# ---------------------------------------------------------------------------

PALETTE_CATEGORICAL = ["00B0F0", "32355F", "00549F", "B30738", "0080B0", "93E2FF"]
PALETTE_SEQUENTIAL = ["93E2FF", "00B0F0", "0080B0", "32355F"]
PALETTE_DIVERGING = ["B30738", "F2F2F2", "00549F"]
CHROME_TEXT = "595959"
GRIDLINE = "D0CECE"
TABLE_HEADER = "00B0F0"
TABLE_BAND = "F2F2F2"
FONT = "Segoe UI"
WHITE = "FFFFFF"
BLACK = "000000"

# Every srgbClr a generated slide may carry (template-native values included).
APPROVED_COLOURS = set(PALETTE_CATEGORICAL) | set(PALETTE_SEQUENTIAL) | set(PALETTE_DIVERGING) | {
    CHROME_TEXT, GRIDLINE, TABLE_HEADER, TABLE_BAND, WHITE, BLACK, "404040",
}
APPROVED_FONTS = {"Segoe UI", "Segoe UI Semibold", "Segoe UI Light", "Segoe UI Black", "Arial"}

# Content area of a "Title Only" slide (below the 0.4in/1.11in title box).
CONTENT_LEFT = Inches(0.92)
CONTENT_TOP = Inches(1.6)
CONTENT_WIDTH = Inches(11.5)
CONTENT_HEIGHT = Inches(5.0)
TITLE_ONLY_CONTENT = {"left": CONTENT_LEFT, "top": CONTENT_TOP,
                      "width": CONTENT_WIDTH, "height": CONTENT_HEIGHT}
SOURCE_LINE_TOP = Inches(6.65)   # sits between content and the master footer

# ---------------------------------------------------------------------------
# Demo / instructional text that must never survive into a deck
# ---------------------------------------------------------------------------

DROP_MARKERS = [
    "How to Use This Template",
    "Delete this slide",
    "Remove this section",
    "To Set a background image",
    "SECTION DIVIDER",
]
# Placeholder prompt strings that show a placeholder was left unfilled.
PROMPT_MARKERS = [
    "Click to edit",
    "Image Holder",
    "Click icon to add picture",
    "Click the icon or drag and drop",
    "Icon name",
    "Picture Description",
]

# Closer tags, stored in slide.name (p:cSld/@name). Order = required tail order.
TAG_CONFIDENTIALITY = "NW Confidentiality"
TAG_THANKYOU = "NW Thank You"
TAG_CLOSING = "NW Closing"
CLOSER_TAGS = [TAG_CONFIDENTIALITY, TAG_THANKYOU, TAG_CLOSING]

CONFIDENTIALITY_TITLE = "Statement of Confidentiality"
# Wording taken from NW_Document_Base_2026.docx (the three paragraphs that follow
# "Statement of Confidentiality"); the bound "Company Name Here" SDT text is the
# {client} slot, the copyright year is {year}.
CONFIDENTIALITY_PARAGRAPHS = [
    "This document contains information that is proprietary and confidential to "
    "Netwoven, Inc. and {client} which shall not be disclosed, transmitted, or "
    "duplicated, used in whole or in part for any purpose other than its intended "
    "purpose. Any use or disclosure in whole or in part of this information without "
    "the express written permission of Netwoven, Inc. and {client} are prohibited.",
    "Any other company and product names mentioned are used for identification "
    "purposes only, and may be trademarks of their respective owners.",
    "©2001 - {year} Netwoven, Inc. All rights reserved. Any use or distribution "
    "of these materials without express authorization of Netwoven, Inc. and {client} "
    "are strictly prohibited.",
]
CONFIDENTIALITY_FONT_PT = 16
CONFIDENTIALITY_SPACE_AFTER_PT = 12

# Personal names that leaked from the source templates; never accepted as creator.
CREATOR_DENYLIST = {"Devang Dhanuka"}

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"


# ---------------------------------------------------------------------------
# Reading helpers
# ---------------------------------------------------------------------------

def _iter_shapes(shapes):
    """Yield every shape, descending into groups."""
    for sh in shapes:
        yield sh
        if isinstance(sh, GroupShape):
            for sub in _iter_shapes(sh.shapes):
                yield sub


def shape_text(shape):
    """All text carried by one shape: text frame, table cells, nothing for pictures."""
    parts = []
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        parts.append(shape.text_frame.text)
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(p for p in parts if p)


def slide_text(slide):
    """Every visible text on a slide (recurses groups and tables)."""
    return "\n".join(t for t in (shape_text(sh) for sh in _iter_shapes(slide.shapes)) if t)


def layout_name(slide):
    try:
        return slide.slide_layout.name or ""
    except Exception:  # noqa: BLE001 - detached layout
        return ""


def _title_shape(slide):
    for sh in slide.placeholders:
        if sh.placeholder_format.type in (1, 3):  # TITLE, CENTER_TITLE
            return sh
    return None


def slide_title_text(slide):
    """Title text with sensible fallbacks for layouts that have no title placeholder."""
    t = _title_shape(slide)
    if t is not None and t.has_text_frame and t.text_frame.text.strip():
        return t.text_frame.text.strip().replace("\n", " ")
    name = slide.name or ""
    if name in CLOSER_TAGS:
        return {TAG_CONFIDENTIALITY: CONFIDENTIALITY_TITLE, TAG_THANKYOU: "Thank You",
                TAG_CLOSING: "Closing"}[name]
    lname = layout_name(slide)
    spec = LAYOUTS.get(lname, {})
    if "statement" in spec:
        p = ph(slide, spec["statement"])
        if p is not None and p.has_text_frame and p.text_frame.text.strip():
            return p.text_frame.text.strip().replace("\n", " ")
    if "stats" in spec:
        labels = []
        for _num, lab in spec["stats"]:
            p = ph(slide, lab)
            if p is not None and p.has_text_frame and p.text_frame.text.strip():
                labels.append(p.text_frame.text.strip())
        if labels:
            return " / ".join(labels)
    if lname == "Closing Slide":
        return "Closing"
    txt = slide_text(slide)
    if "THANK YOU" in txt.upper():
        return "Thank You"
    first = next((ln.strip() for ln in txt.splitlines() if ln.strip()), "")
    return first[:80]


def ph(slide, idx):
    """Placeholder with this idx, or None."""
    for sh in slide.placeholders:
        if sh.placeholder_format.idx == idx:
            return sh
    return None


def get_layout(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise KeyError("layout %r not found in this deck" % name)


def layout_has_title(layout):
    return any(p.placeholder_format.type in (1, 3) for p in layout.placeholders)


# ---------------------------------------------------------------------------
# Low-level writing helpers
# ---------------------------------------------------------------------------

def remove_shape(shape):
    if shape is None:
        return
    el = shape._element
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def is_empty_placeholder(shape):
    """A p:sp placeholder with no text (covers unfilled picture placeholders too)."""
    if not shape.is_placeholder:
        return False
    if not shape._element.tag.endswith("}sp"):
        return False  # filled pictures (p:pic) and tables/charts (graphicFrame)
    if not shape.has_text_frame:
        return True
    return not shape.text_frame.text.strip()


def remove_empty_placeholders(slide):
    """Delete unfilled placeholders so no 'Click to edit' prompt renders. Returns count."""
    removed = 0
    for sh in list(slide.shapes):
        if is_empty_placeholder(sh):
            remove_shape(sh)
            removed += 1
    return removed


def set_paragraph_no_bullet(paragraph):
    """Force <a:buNone/> on a paragraph (kills the layout's inherited bullet)."""
    pPr = paragraph._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buBlip"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    bu = etree.SubElement(pPr, qn("a:buNone"))
    # schema order: lnSpc, spcBef, spcAft, buClr*, buSz*, buFont*, bu*, tabLst, defRPr
    for tag in ("a:tabLst", "a:defRPr", "a:extLst"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
            pPr.append(el)
    return bu


def set_paragraph_space_after(paragraph, pt):
    paragraph.space_after = Pt(pt)


def set_alt_text(shape, text):
    """Accessibility description via cNvPr/@descr (the shape's own cNvPr is first)."""
    cNvPr = shape._element.find(".//" + qn("p:cNvPr"))
    if cNvPr is not None:
        cNvPr.set("descr", text)
    return cNvPr


def _style_run(run, size_pt=None, bold=None, color=None, name=FONT, italic=None):
    f = run.font
    f.name = name
    if size_pt is not None:
        f.size = Pt(size_pt)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color:
        f.color.rgb = RGBColor.from_string(color)


def fill_text_frame(shape, lines, levels=None, no_bullets=False, size_pt=None, bold=None,
                    color=None, space_after_pt=None):
    """Replace a placeholder's text with `lines` (str or list). Levels are 0-based."""
    if shape is None:
        # Every caller resolves `shape` via ph(slide, idx), which returns
        # None when that idx isn't on the layout (see ph()'s docstring).
        # Guarding centrally here, the way set_title() already guards its
        # own placeholder lookup, turns a layout/LAYOUTS-map mismatch into
        # a clear error instead of an AttributeError on None.text_frame.
        raise ValueError("fill_text_frame: shape is None (placeholder idx not found on this layout)")
    if isinstance(lines, str):
        lines = [lines]
    lines = [str(x) for x in lines]
    levels = list(levels or [0] * len(lines))
    tf = shape.text_frame
    # A layout's placeholder can be sized for much shorter text than a real
    # spec ever supplies (e.g. Three Stat's labels): without this, PowerPoint
    # trusts the template's cached geometry and overflows the shape instead
    # of wrapping or shrinking, which LibreOffice's renderer papers over --
    # the gap that let a genuinely broken slide pass a LibreOffice-only check.
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    # keep the first paragraph element (it carries any layout-specific pPr), clear the rest
    paras = list(tf.paragraphs)
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    first = tf.paragraphs[0]
    for r in list(first.runs):
        r._r.getparent().remove(r._r)
    for i, text in enumerate(lines):
        p = first if i == 0 else tf.add_paragraph()
        p.level = max(0, min(8, int(levels[i] if i < len(levels) else 0)))
        run = p.add_run()
        run.text = text
        if size_pt or bold is not None or color:
            _style_run(run, size_pt=size_pt, bold=bold, color=color)
        if no_bullets:
            set_paragraph_no_bullet(p)
        if space_after_pt is not None:
            p.space_after = Pt(space_after_pt)
    return tf


def set_title(slide, text):
    t = _title_shape(slide)
    if t is None:
        raise ValueError("slide layout %r has no title placeholder" % layout_name(slide))
    fill_text_frame(t, text)
    return t


def _prep(prs, name):
    if name in BROKEN_LAYOUTS:
        raise ValueError("layout %r is unusable (no title placeholder)" % name)
    return prs.slides.add_slide(get_layout(prs, name))


def _fill_or_remove(slide, idx, lines, **kw):
    shape = ph(slide, idx)
    if shape is None:
        return None
    if lines:
        fill_text_frame(shape, lines, **kw)
    else:
        remove_shape(shape)
    return shape


def _insert_picture_or_remove(slide, idx, image_path, alt=None):
    shape = ph(slide, idx)
    if shape is None:
        return None
    if image_path:
        pic = shape.insert_picture(image_path)
        if alt:
            set_alt_text(pic, alt)
        return pic
    remove_shape(shape)
    return None


# ---------------------------------------------------------------------------
# Slide builders (every helper returns the new slide)
# ---------------------------------------------------------------------------

def add_cover(prs, title, subtitle=None, images=None):
    """Cover on 'Title Slide for Verticals'. images: up to two file paths (portrait slots)."""
    spec = LAYOUTS["Title Slide for Verticals"]
    slide = _prep(prs, "Title Slide for Verticals")
    fill_text_frame(ph(slide, spec["title"]), title)
    _fill_or_remove(slide, spec["subtitle"], subtitle)
    images = list(images or [])
    for i, idx in enumerate(spec["pics"]):
        _insert_picture_or_remove(slide, idx, images[i] if i < len(images) else None,
                                  alt="Cover image %d" % (i + 1))
    return slide


def add_title_only(prs, title, layout_name="Title Only"):
    """`layout_name` lets a caller request the sibling "Title Only (Centered)"
    layout instead of the default "Title Only" -- both share the same title
    placeholder idx (see LAYOUTS), only their layout XML differs."""
    slide = _prep(prs, layout_name)
    set_title(slide, title)
    return slide


def add_title_content(prs, title, bullets, levels=None):
    spec = LAYOUTS["Title and Content"]
    slide = _prep(prs, "Title and Content")
    set_title(slide, title)
    body = ph(slide, spec["body"])
    fill_text_frame(body, bullets, levels=levels)
    # A short bullet list (the common case on a lean, one-idea-per-slide deck)
    # otherwise sits stranded at the top of a placeholder sized for a much
    # longer one, leaving most of the slide looking unfinished rather than
    # deliberately spacious.
    body.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return slide


def add_section_header(prs, title, subtitle=None):
    spec = LAYOUTS["Section Header"]
    slide = _prep(prs, "Section Header")
    set_title(slide, title)
    _fill_or_remove(slide, spec["body"], subtitle)
    _insert_picture_or_remove(slide, spec["pic"], None)
    return slide


def add_two_content(prs, title, left, right, left_levels=None, right_levels=None):
    spec = LAYOUTS["Two Content"]
    slide = _prep(prs, "Two Content")
    set_title(slide, title)
    left_ph = ph(slide, spec["left"])
    right_ph = ph(slide, spec["right"])
    fill_text_frame(left_ph, left, levels=left_levels)
    fill_text_frame(right_ph, right, levels=right_levels)
    left_ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    right_ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return slide


def add_three_column(prs, title, columns, headings=None):
    """columns: three lists of bullet strings; headings: optional three bold lead lines."""
    spec = LAYOUTS["Three Column"]
    slide = _prep(prs, "Three Column")
    set_title(slide, title)
    for i, idx in enumerate(spec["columns"]):
        col = list(columns[i]) if i < len(columns) else []
        shape = ph(slide, idx)
        if headings and i < len(headings) and headings[i]:
            fill_text_frame(shape, [headings[i]] + col)
            p0 = shape.text_frame.paragraphs[0]
            set_paragraph_no_bullet(p0)
            for r in p0.runs:
                r.font.bold = True
        elif col:
            fill_text_frame(shape, col)
        else:
            remove_shape(shape)
    return slide


def add_comparison(prs, title, left_heading, left, right_heading, right):
    spec = LAYOUTS["Comparison"]
    slide = _prep(prs, "Comparison")
    set_title(slide, title)
    fill_text_frame(ph(slide, spec["heads"][0]), left_heading)
    fill_text_frame(ph(slide, spec["heads"][1]), right_heading)
    fill_text_frame(ph(slide, spec["bodies"][0]), left)
    fill_text_frame(ph(slide, spec["bodies"][1]), right)
    return slide


def add_three_stat(prs, stats):
    """stats: three (number, label) tuples, e.g. ("42%", "of tickets closed same day")."""
    spec = LAYOUTS["Three Stat"]
    slide = _prep(prs, "Three Stat")
    for i, (num_idx, lab_idx) in enumerate(spec["stats"]):
        if i < len(stats) and stats[i]:
            number, label = stats[i]
            fill_text_frame(ph(slide, num_idx), str(number))
            fill_text_frame(ph(slide, lab_idx), str(label), size_pt=24)
        else:
            _fill_or_remove(slide, num_idx, None)
            _fill_or_remove(slide, lab_idx, None)
    return slide


def add_big_statement(prs, statement, eyebrow=None, image=None):
    spec = LAYOUTS["Big Statement with Illustration"]
    slide = _prep(prs, "Big Statement with Illustration")
    fill_text_frame(ph(slide, spec["statement"]), statement)
    _fill_or_remove(slide, spec["eyebrow"], eyebrow)
    _insert_picture_or_remove(slide, spec["pic"], image, alt="Illustration")
    return slide


def add_confidentiality(prs, client, year=None):
    """Statement of Confidentiality on 'Title and Content', tagged TAG_CONFIDENTIALITY."""
    if not client:
        raise ValueError("a client name is required for the confidentiality slide")
    year = year or _dt.date.today().year
    spec = LAYOUTS["Title and Content"]
    slide = _prep(prs, "Title and Content")
    set_title(slide, CONFIDENTIALITY_TITLE)
    paras = [p.format(client=client, year=year) for p in CONFIDENTIALITY_PARAGRAPHS]
    fill_text_frame(ph(slide, spec["body"]), paras, no_bullets=True,
                    size_pt=CONFIDENTIALITY_FONT_PT,
                    space_after_pt=CONFIDENTIALITY_SPACE_AFTER_PT)
    slide.name = TAG_CONFIDENTIALITY
    return slide


# ---------------------------------------------------------------------------
# Free-floating content: tables, charts, text boxes
# ---------------------------------------------------------------------------

def _cell_fill(cell, hex_colour):
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor.from_string(hex_colour)


def _cell_text(cell, text, size_pt, bold, colour, align=None):
    cell.text = ""
    tf = cell.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "" if text is None else str(text)
    _style_run(run, size_pt=size_pt, bold=bold, color=colour)
    if align is not None:
        p.alignment = align
    cell.margin_left = cell.margin_right = Inches(0.08)
    cell.margin_top = cell.margin_bottom = Inches(0.04)
    # A row sized for one line of text otherwise leaves it stranded at the
    # top of whatever height add_table (or a caller's col_widths_in-style
    # sizing) actually gave the row -- most visible on a table with few rows
    # in a generously sized content area.
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE


def add_table(slide, header, rows, left, top, width, height, col_widths=None, font_pt=None,
              align_numbers=True):
    """Table with explicit Netwoven fills: header bold white Segoe UI, body rows
    banded FFFFFF / F2F2F2 (explicit, not style-driven). Returns the graphicFrame.
    font_pt=None scales to the row height PowerPoint will actually give each row
    (height split evenly across rows) -- a flat size looks sparse when a small
    table (few rows) is placed in a large content area, and cramped when a
    large one is; pass an explicit font_pt to opt out of the scaling."""
    n_rows = len(rows) + 1
    n_cols = len(header)
    if font_pt is None:
        row_height_pt = (height / n_rows) / 12700  # EMU -> pt
        font_pt = max(13, min(22, int(row_height_pt * 0.3)))
    gf = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = gf.table
    table.first_row = True
    table.horz_banding = False
    table.vert_banding = False
    if col_widths:
        total = float(sum(col_widths))
        for i, w in enumerate(col_widths):
            table.columns[i].width = Emu(int(width * (w / total)))
    for c, text in enumerate(header):
        cell = table.cell(0, c)
        _cell_fill(cell, TABLE_HEADER)
        _cell_text(cell, text, font_pt + 2, True, WHITE, PP_ALIGN.CENTER)
    num_re = re.compile(r"^[\s$€£(+-]*[\d.,]+\s*[%)]?[A-Za-z]{0,2}$")
    for r, row in enumerate(rows, start=1):
        band = WHITE if r % 2 == 1 else TABLE_BAND
        for c in range(n_cols):
            val = row[c] if c < len(row) else ""
            cell = table.cell(r, c)
            _cell_fill(cell, band)
            align = None
            if align_numbers and c > 0 and isinstance(val, (int, float)):
                align = PP_ALIGN.RIGHT
            elif align_numbers and c > 0 and isinstance(val, str) and num_re.match(val.strip() or "x"):
                align = PP_ALIGN.RIGHT
            _cell_text(cell, val, font_pt, False, CHROME_TEXT, align)
    set_alt_text(gf, "Table: " + ", ".join(str(h) for h in header))
    return gf


_CHART_KINDS = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "stacked": XL_CHART_TYPE.COLUMN_STACKED,
}


def _rgb(hex_colour):
    return RGBColor.from_string(hex_colour)


def add_chart(slide, kind, categories, series, left, top, width, height, data_labels=False,
              number_format=None, palette=None, legend=True):
    """Native chart with Netwoven chrome. kind in {column, bar, line, pie, stacked};
    series: dict name -> list of values (ordered). Returns the graphicFrame."""
    if kind not in _CHART_KINDS:
        raise ValueError("kind must be one of %s" % sorted(_CHART_KINDS))
    palette = list(palette or PALETTE_CATEGORICAL)
    cd = CategoryChartData()
    cd.categories = list(categories)
    for name, values in series.items():
        cd.add_series(name, list(values))
    gf = slide.shapes.add_chart(_CHART_KINDS[kind], left, top, width, height, cd)
    chart = gf.chart
    chart.has_title = False
    chart.font.name = FONT
    chart.font.size = Pt(12)
    chart.font.color.rgb = _rgb(CHROME_TEXT)
    plot = chart.plots[0]
    if kind == "pie":
        ser = plot.series[0]
        for j in range(len(categories)):
            pt = ser.points[j]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _rgb(palette[j % len(palette)])
            pt.format.line.color.rgb = _rgb(WHITE)
        # a series-level fill so the "explicit fill" check also holds at c:ser level
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = _rgb(palette[0])
        chart.has_legend = bool(legend)
        if legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(12)
            chart.legend.font.name = FONT
            chart.legend.font.color.rgb = _rgb(CHROME_TEXT)
    else:
        for i, ser in enumerate(plot.series):
            colour = _rgb(palette[i % len(palette)])
            if kind == "line":
                ser.format.line.color.rgb = colour
                ser.format.line.width = Pt(2.25)
                ser.smooth = False
                try:
                    ser.marker.format.fill.solid()
                    ser.marker.format.fill.fore_color.rgb = colour
                    ser.marker.format.line.color.rgb = colour
                except Exception:  # noqa: BLE001 - marker API differences
                    pass
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = colour
        if kind in ("column", "bar", "stacked"):
            plot.gap_width = 80
            if kind == "stacked":
                plot.overlap = 100
        va, ca = chart.value_axis, chart.category_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = _rgb(GRIDLINE)
        va.has_minor_gridlines = False
        va.format.line.color.rgb = _rgb(GRIDLINE)
        ca.format.line.color.rgb = _rgb(GRIDLINE)
        ca.has_major_gridlines = False
        for ax in (va, ca):
            ax.tick_labels.font.size = Pt(12)
            ax.tick_labels.font.name = FONT
            ax.tick_labels.font.color.rgb = _rgb(CHROME_TEXT)
        if number_format:
            va.tick_labels.number_format = number_format
            va.tick_labels.number_format_is_linked = False
        multi = len(series) > 1
        chart.has_legend = bool(legend) and multi
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(12)
            chart.legend.font.name = FONT
            chart.legend.font.color.rgb = _rgb(CHROME_TEXT)
    if data_labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(12)
        dl.font.name = FONT
        dl.font.color.rgb = _rgb(CHROME_TEXT)
        if number_format:
            dl.number_format = number_format
            dl.number_format_is_linked = False
        if kind == "pie":
            dl.show_percentage = True
            dl.show_value = False
            dl.show_category_name = False
    set_alt_text(gf, "%s chart: %s" % (kind.capitalize(), ", ".join(series.keys())))
    return gf


def add_textbox(slide, text, left, top, width, height, size_pt=10, color=CHROME_TEXT, bold=False,
                italic=False, align=None, name=FONT):
    """Plain text box (default styling suits source lines / footnotes: 10pt 595959)."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = str(line)
        _style_run(run, size_pt=size_pt, bold=bold, color=color, name=name, italic=italic)
        if align is not None:
            p.alignment = align
    return tb


def add_source_line(slide, text):
    """Source / note line under the content area of a Title Only slide."""
    return add_textbox(slide, text, CONTENT_LEFT, SOURCE_LINE_TOP, CONTENT_WIDTH, Inches(0.3),
                       size_pt=10, color=CHROME_TEXT, italic=False)


# ---------------------------------------------------------------------------
# Deck-level surgery
# ---------------------------------------------------------------------------

def slide_id_elements(prs):
    return list(prs.slides._sldIdLst)


def delete_slide(prs, slide):
    """Drop a slide (and its relationship) so its part is not written on save."""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get("{%s}id" % R_NS)
        if prs.part.related_part(rId) is slide.part:
            prs.part.drop_rel(rId)
            sldIdLst.remove(sldId)
            return True
    return False


def renumber_slide_parts(prs):
    """Rename slide parts to /ppt/slides/slide1..N in presentation order.

    python-pptx names a new slide part slide{len(slides)+1}.xml. After deleting
    slides that number can collide with a surviving part (the base's closers are
    slide5/slide6), and the zip writer then silently overwrites the closer with
    the new slide. Call this after every deletion and before adding slides."""
    from pptx.opc.packuri import PackURI
    slides = list(prs.slides)
    # two passes so a temporary name never equals another part's final name
    for i, s in enumerate(slides):
        s.part.partname = PackURI("/ppt/slides/slide_tmp%d.xml" % (i + 1))
    for i, s in enumerate(slides):
        s.part.partname = PackURI("/ppt/slides/slide%d.xml" % (i + 1))
    return len(slides)


def move_slide(prs, slide, new_index):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get("{%s}id" % R_NS)
        if prs.part.related_part(rId) is slide.part:
            sldIdLst.remove(sldId)
            sldIdLst.insert(new_index, sldId)
            return True
    return False


def closer_kind(slide):
    """Return the CLOSER_TAGS entry this slide is (tag, then text/layout/title fallbacks)."""
    name = slide.name or ""
    if name in CLOSER_TAGS:
        return name
    title = slide_title_text(slide)
    if title == CONFIDENTIALITY_TITLE:
        return TAG_CONFIDENTIALITY
    if layout_name(slide) == "Closing Slide":
        return TAG_CLOSING
    txt = slide_text(slide).upper()
    if "THANK YOU" in txt and "TIME TODAY" in txt:
        return TAG_THANKYOU
    if layout_name(slide) == "Blank" and "THANK YOU" in txt:
        return TAG_THANKYOU
    return None


def tag_closers(prs):
    """Write CLOSER_TAGS into slide.name for every detected closer; returns {tag: slide}."""
    found = {}
    for slide in prs.slides:
        kind = closer_kind(slide)
        if kind and kind not in found:
            slide.name = kind
            found[kind] = slide
    return found


def move_closers_last(prs):
    """Closers to the tail in CLOSER_TAGS order. Returns the ordered list of tags moved."""
    slides = list(prs.slides)
    by_tag = {}
    for s in slides:
        k = closer_kind(s)
        if k and k not in by_tag:
            by_tag[k] = s
    moved = []
    n = len(slides)
    for tag in CLOSER_TAGS:
        s = by_tag.get(tag)
        if s is not None:
            s.name = tag
            move_slide(prs, s, n - 1)
            moved.append(tag)
    return moved


def remove_dangling_section_ids(prs):
    """Drop p14:sldId entries in a p14:sectionLst that point to deleted slides. Returns count."""
    pres = prs.part._element
    valid = {s.get("id") for s in prs.slides._sldIdLst}
    removed = 0
    for sldId in pres.iter("{%s}sldId" % P14_NS):
        if sldId.get("id") not in valid:
            sldId.getparent().remove(sldId)
            removed += 1
    return removed


def patch_app_xml(path):
    """Zip-level fix of docProps/app.xml: drop HeadingPairs/TitlesOfParts (stale demo slide
    titles), set <Slides> to the real count. Rewrites the zip in place; returns slide count."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        pres = z.read("ppt/presentation.xml")
        n_slides = len(re.findall(rb"<p:sldId\b", pres))
        app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else None
    if app is None:
        return n_slides
    app = re.sub(r"<HeadingPairs>.*?</HeadingPairs>", "", app, flags=re.S)
    app = re.sub(r"<TitlesOfParts>.*?</TitlesOfParts>", "", app, flags=re.S)
    app = re.sub(r"<HeadingPairs/>|<TitlesOfParts/>", "", app)
    if re.search(r"<Slides>\d*</Slides>", app):
        app = re.sub(r"<Slides>\d*</Slides>", "<Slides>%d</Slides>" % n_slides, app)
    else:
        app = app.replace("</Properties>", "<Slides>%d</Slides></Properties>" % n_slides)
    app = re.sub(r"<Notes>\d*</Notes>", "<Notes>0</Notes>", app)
    app = re.sub(r"<HiddenSlides>\d*</HiddenSlides>", "<HiddenSlides>0</HiddenSlides>", app)
    _rewrite_zip_member(path, "docProps/app.xml", app.encode("utf-8"))
    return n_slides


def _rewrite_zip_member(path, member, data):
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            payload = data if item.filename == member else zin.read(item.filename)
            zout.writestr(item, payload)
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# Validation (called by validate_deliverable.py: validate_deck(prs, path, args))
# ---------------------------------------------------------------------------

_WORD_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9'%$.,-]*")


def _words(text):
    return len(_WORD_RE.findall(text or ""))


def _read_zip_parts(path):
    parts = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if (n.startswith("ppt/slides/slide") and n.endswith(".xml")) or \
               n.startswith("ppt/charts/chart") and n.endswith(".xml") or \
               n in ("docProps/app.xml", "docProps/core.xml", "ppt/presentation.xml") or \
               n.startswith("ppt/slides/_rels/"):
                parts[n] = z.read(n).decode("utf-8", "replace")
    return parts


def _slide_part_order(prs):
    """[(slide, 'ppt/slides/slideN.xml')] in presentation order."""
    out = []
    for s in prs.slides:
        out.append((s, s.part.partname.lstrip("/")))
    return out


def validate_deck(prs, path, args):
    """Deck checks from the plan (Track 1 item 6). Returns a list of
    {"check", "passed", "evidence", "hard"} dicts. Text/colour/font checks read the raw
    slide XML from the saved file at `path`."""
    internal = bool(getattr(args, "internal", False))
    client = getattr(args, "client", None)
    denylist = set(CREATOR_DENYLIST) | set(getattr(args, "creator_denylist", None) or [])
    results = []

    def add(name, passed, evidence, hard=True):
        results.append({"check": name, "passed": bool(passed), "evidence": evidence, "hard": hard})

    parts = _read_zip_parts(path)
    order = _slide_part_order(prs)
    slides = [s for s, _ in order]
    n = len(slides)

    # --- hard: cover
    if n == 0:
        add("deck has slides", False, "no slides")
        return results
    first = slides[0]
    add("first slide is the cover layout", layout_name(first) == "Title Slide for Verticals",
        "slide 1 layout = %r" % layout_name(first))
    t = _title_shape(first)
    cover_title = t.text_frame.text.strip() if (t is not None and t.has_text_frame) else ""
    add("cover has a title", bool(cover_title), "cover title = %r" % cover_title)

    # --- hard: broken layouts
    bad = [i + 1 for i, s in enumerate(slides) if layout_name(s) in BROKEN_LAYOUTS]
    add("no slide on a broken layout", not bad,
        "slides on %s: %s" % (sorted(BROKEN_LAYOUTS), bad or "none"))

    # --- hard: demo / prompt markers in slide XML and app.xml
    marker_hits = []
    text_of = {}
    for i, (s, partname) in enumerate(order, start=1):
        xml = parts.get(partname, "")
        texts = " ".join(re.findall(r"<a:t>([^<]*)</a:t>", xml))
        text_of[i] = texts
        for m in DROP_MARKERS + PROMPT_MARKERS:
            if m.lower() in texts.lower():
                marker_hits.append("slide %d: %r" % (i, m))
    app_xml = parts.get("docProps/app.xml", "")
    for m in DROP_MARKERS + PROMPT_MARKERS + ["PowerPoint Presentation"]:
        if m.lower() in app_xml.lower():
            marker_hits.append("app.xml: %r" % m)
    add("no demo or prompt text in slides / app.xml", not marker_hits,
        "; ".join(marker_hits) if marker_hits else "clean")
    add("app.xml has no stale TitlesOfParts", "TitlesOfParts" not in app_xml and "HeadingPairs" not in app_xml,
        "TitlesOfParts present" if "TitlesOfParts" in app_xml else "removed")

    # --- hard: closers present and last, in order
    kinds = [closer_kind(s) for s in slides]
    expected = [tag for tag in CLOSER_TAGS if not (internal and tag == TAG_CONFIDENTIALITY)]
    present = [k for k in kinds if k]
    missing = [tag for tag in expected if tag not in present]
    add("closers present", not missing,
        "expected %s; missing %s" % (expected, missing or "none"))
    tail = kinds[-len(present):] if present else []
    tail_ok = bool(present) and tail == present and present == [t for t in CLOSER_TAGS if t in present]
    body_has_closer = any(k for k in kinds[: n - len(present)]) if present else False
    add("closers are last and in order", tail_ok and not body_has_closer,
        "tail = %s" % [k or layout_name(s) for k, s in zip(kinds, slides)][-4:])
    if internal:
        add("no confidentiality slide in an internal deck", TAG_CONFIDENTIALITY not in present,
            "confidentiality %s" % ("present" if TAG_CONFIDENTIALITY in present else "absent"))
    conf_idx = [i for i, k in enumerate(kinds) if k == TAG_CONFIDENTIALITY]
    if conf_idx:
        ctext = slide_text(slides[conf_idx[0]])
        add("confidentiality slide has the statement title",
            CONFIDENTIALITY_TITLE in ctext, "title %s" % ("found" if CONFIDENTIALITY_TITLE in ctext else "missing"))
        add("confidentiality slide carries the three paragraphs",
            "Netwoven, Inc." in ctext and "trademarks" in ctext and "©2001" in ctext,
            "%d words" % _words(ctext))
        if client:
            add("confidentiality slide names the client", client in ctext,
                "client %r %s" % (client, "found" if client in ctext else "missing"))

    # --- hard: no empty placeholders
    empties = []
    for i, s in enumerate(slides, start=1):
        for sh in s.shapes:
            if is_empty_placeholder(sh):
                empties.append("slide %d idx %s" % (i, sh.placeholder_format.idx))
    add("no empty placeholders", not empties, "; ".join(empties) if empties else "none")

    # --- hard: dangling section ids
    pres_xml = parts.get("ppt/presentation.xml", "")
    valid_ids = set(re.findall(r'<p:sldId id="(\d+)"', pres_xml))
    sect_ids = re.findall(r'<p14:sldId id="(\d+)"', pres_xml)
    dangling = [i for i in sect_ids if i not in valid_ids]
    add("no dangling section slide ids", not dangling,
        "dangling %s" % dangling if dangling else ("no sectionLst" if not sect_ids else "all valid"))

    # --- hard: core metadata
    core_title = (prs.core_properties.title or "").strip()
    add("core title set", bool(core_title), "dc:title = %r" % core_title)
    creator = (prs.core_properties.author or "").strip()
    lastmod = (prs.core_properties.last_modified_by or "").strip()
    leaked = [v for v in (creator, lastmod) if v in denylist]
    add("creator scrubbed", not leaked, "creator=%r lastModifiedBy=%r" % (creator, lastmod))
    add("creator is Netwoven", creator in ("", "Netwoven"), "creator=%r" % creator, hard=False)

    # --- warnings per slide
    closer_set = set()
    for i, k in enumerate(kinds, start=1):
        if k:
            closer_set.add(i)
    untitled, long_titles, dense, bad_fonts, bad_colours, many_visuals = [], [], [], [], [], []
    for i, (s, partname) in enumerate(order, start=1):
        if i in closer_set:
            continue
        try:
            lay = s.slide_layout
        except Exception:  # noqa: BLE001 - detached layout, same as layout_name()
            lay = None
        has_title_ph = layout_has_title(lay) if lay is not None else False
        tshape = _title_shape(s)
        ttext = tshape.text_frame.text.strip() if (tshape is not None and tshape.has_text_frame) else ""
        if has_title_ph and not ttext:
            untitled.append(i)
        if _words(ttext) > 12:
            long_titles.append("slide %d (%d words)" % (i, _words(ttext)))
        body_paras, body_words = 0, 0
        title_el = tshape._element if tshape is not None else None
        for sh in _iter_shapes(s.shapes):
            # python-pptx hands out fresh proxies on each iteration: compare elements, not proxies
            if sh._element is title_el or not getattr(sh, "has_text_frame", False) or not sh.has_text_frame:
                continue
            if getattr(sh, "has_table", False) and sh.has_table:
                continue
            for p in sh.text_frame.paragraphs:
                if p.text.strip():
                    body_paras += 1
                    body_words += _words(p.text)
        if i != 1 and (body_paras > 8 or body_words > 90):
            dense.append("slide %d (%d paras, %d words)" % (i, body_paras, body_words))
        xml = parts.get(partname, "")
        fonts = set(re.findall(r'<a:latin typeface="([^"]+)"', xml))
        off = sorted(f for f in fonts if not f.startswith("+") and f not in APPROVED_FONTS)
        if off:
            bad_fonts.append("slide %d: %s" % (i, off))
        cols = set(c.upper() for c in re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', xml))
        offc = sorted(c for c in cols if c not in APPROVED_COLOURS)
        if offc:
            bad_colours.append("slide %d: %s" % (i, offc))
        gfs = len(re.findall(r"<p:graphicFrame\b", xml)) + len(re.findall(r"<p:pic\b", xml))
        if gfs > 2:
            many_visuals.append("slide %d (%d)" % (i, gfs))
    add("every content slide has a title", not untitled, "untitled: %s" % (untitled or "none"), hard=False)
    add("titles are 12 words or fewer", not long_titles, "; ".join(long_titles) or "ok", hard=False)
    add("body density (<= 8 paragraphs, <= 90 words)", not dense, "; ".join(dense) or "ok", hard=False)
    add("fonts within the Segoe UI family / Arial", not bad_fonts, "; ".join(bad_fonts) or "ok", hard=False)
    add("colours within the approved palette", not bad_colours, "; ".join(bad_colours) or "ok", hard=False)
    add("at most 2 visuals per slide", not many_visuals, "; ".join(many_visuals) or "ok", hard=False)

    # --- warning: chart series without explicit fill
    unfilled = []
    for name, xml in parts.items():
        if not name.startswith("ppt/charts/"):
            continue
        try:
            root = etree.fromstring(xml.encode("utf-8"))
        except etree.XMLSyntaxError:
            continue
        is_pie = root.find(".//{%s}pieChart" % C_NS) is not None or root.find(".//{%s}doughnutChart" % C_NS) is not None
        for k, ser in enumerate(root.iter("{%s}ser" % C_NS)):
            spPr = ser.find("{%s}spPr" % C_NS)
            ok = spPr is not None and spPr.find(".//{%s}solidFill" % A_NS) is not None
            if is_pie:
                pts = ser.findall("{%s}dPt" % C_NS)
                ok = ok or (pts and all(p.find(".//{%s}solidFill" % A_NS) is not None for p in pts))
            if not ok:
                unfilled.append("%s series %d" % (name.split("/")[-1], k + 1))
    add("chart series carry explicit palette fills", not unfilled, "; ".join(unfilled) or "ok", hard=False)

    return results


def summarize(results):
    hard_fail = [r for r in results if r["hard"] and not r["passed"]]
    warn = [r for r in results if not r["hard"] and not r["passed"]]
    return hard_fail, warn
