#!/usr/bin/env python3
"""
Shared Word-side helpers for the Netwoven 2026 deliverable builders.

Used by new_deliverable_docx.py (general document), new_sow_docx.py
(Milestone / T&M statements of work) and validate_deliverable.py, and
importable from a skill session that wants to fill a generated shell with
python-docx while keeping the template's own mechanics intact.

Design rules (they explain most of the shape of this file):

* Everything is lxml on the raw XML, never `row.cells`: the SOW product table
  wraps its checkbox cells in w:sdt elements that python-docx's row.cells does
  not see, and fee tables carry live fields and bookmarks that must survive.
* The SOW placeholders (<<Client>>, <<Project Name>>, MM/DD/YYYY) live inside
  bound content controls (w:sdt with w:dataBinding).  Word re-resolves them
  from docProps/app.xml, docProps/core.xml and customXml/item1.xml on open,
  but LibreOffice and Word-before-rebind show the *cached* text, so
  `set_bound_fields` rewrites the cached w:t text in every bound SDT of every
  story part (document, headers, footers) as well as the bound values.
* Bookmark pairs must never be half-deleted (a dangling bookmarkStart/End
  corrupts TOC anchors and REF targets), hence `preserve_half_bookmarks`.
* Runtime dependencies: python-docx, lxml only.  Python 3.9+ syntax.

Nothing in here renders or needs LibreOffice.
"""
from __future__ import annotations

import copy
import datetime as _dt
import re
import shutil
import sys
import zipfile
from copy import deepcopy

from lxml import etree

# --------------------------------------------------------------------------
# namespaces
# --------------------------------------------------------------------------
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
W = "{%s}" % NS["w"]
W14 = "{%s}" % NS["w14"]
R = "{%s}" % NS["r"]
WP = "{%s}" % NS["wp"]
A = "{%s}" % NS["a"]
PIC = "{%s}" % NS["pic"]
MC = "{%s}" % NS["mc"]
V = "{%s}" % NS["v"]
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
RT_PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
RT_OLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
RT_DIAGRAM_PREFIXES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagram",
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
)

LOGO_BOX_NAME = "Rectangle: Rounded Corners 5"
LOGO_BOX_CX, LOGO_BOX_CY = 1390650, 600075  # EMU, the template's rounded box

CHECKED_GLYPH = "☒"    # ballot box with X
UNCHECKED_GLYPH = "☐"  # ballot box


def log(msg):
    print(msg, file=sys.stderr)


# --------------------------------------------------------------------------
# tiny XML utilities
# --------------------------------------------------------------------------
def E(tag, attrib=None, text=None, ns=W):
    """Build a w:* element (or another namespace via ns=)."""
    el = etree.Element(ns + tag)
    for k, v in (attrib or {}).items():
        if k.startswith("{"):
            el.set(k, v)
        else:
            el.set(ns + k, v)
    if text is not None:
        el.text = text
    return el


def para_text(p):
    """Concatenated w:t text of any element (paragraph, cell, sdt...)."""
    return "".join(t.text or "" for t in p.iter(W + "t"))


def norm(s):
    """Whitespace-collapsed, nbsp-normalised, stripped text for matching."""
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


def is_p(el):
    return el.tag == W + "p"


def is_tbl(el):
    return el.tag == W + "tbl"


def p_style(el):
    if el.tag != W + "p":
        return None
    st = el.find(W + "pPr/" + W + "pStyle")
    return st.get(W + "val") if st is not None else None


def is_p_with_style(el, style):
    return p_style(el) == style


def p_numpr(el):
    """(ilvl, numId) of a paragraph or None."""
    n = el.find(W + "pPr/" + W + "numPr") if is_p(el) else None
    if n is None:
        return None
    il = n.find(W + "ilvl")
    ni = n.find(W + "numId")
    return (il.get(W + "val") if il is not None else "0",
            ni.get(W + "val") if ni is not None else None)


def is_empty_p(el):
    """A paragraph with no text, no drawing, no field and no object."""
    if not is_p(el):
        return False
    if para_text(el).strip():
        return False
    for tag in ("drawing", "object", "fldChar", "instrText", "br", "pict"):
        if el.find(".//" + W + tag) is not None:
            return False
    if el.find(".//" + MC + "AlternateContent") is not None:
        return False
    return True


def find_index(children, pred, start=0):
    for i in range(start, len(children)):
        if pred(children[i]):
            return i
    return None


def find_heading(children, style, text, start=0):
    """Index of the paragraph with paragraph style `style` whose text is
    `text` (whitespace-normalised, case-insensitive)."""
    want = norm(text).lower()
    return find_index(children,
                      lambda el: is_p_with_style(el, style)
                      and norm(para_text(el)).lower() == want, start)


def body_children(doc):
    return list(doc.element.body)


# --------------------------------------------------------------------------
# bookmark-safe removal
# --------------------------------------------------------------------------
def preserve_half_bookmarks(body, doomed):
    """Before deleting `doomed` elements, move bookmarkStart/End whose pair
    lives outside the doomed range to just before the first doomed element.
    Fully-inside pairs die with the range."""
    doomed_set = set(id(el) for el in doomed)

    def inside(el):
        anc = el
        while anc is not None:
            if id(anc) in doomed_set:
                return True
            anc = anc.getparent()
        return False

    starts, ends = {}, {}
    for el in body.iter(W + "bookmarkStart"):
        starts[el.get(W + "id")] = el
    for el in body.iter(W + "bookmarkEnd"):
        ends[el.get(W + "id")] = el
    rescued = []
    for bid, s in starts.items():
        e = ends.get(bid)
        s_in, e_in = inside(s), (e is not None and inside(e))
        if s_in != e_in:
            rescued.append(s if s_in else e)
    if not doomed:
        return
    # Insert a COPY of each rescued marker just before the doomed range,
    # and always leave the ORIGINAL untouched. Moving the original instead
    # would corrupt the rescue whenever it IS one of the doomed elements
    # (the bookmark that starts the doomed range with its pair outside it,
    # a real case, not just a hypothetical): the caller's own removal loop
    # (remove_range/remove_elements) iterates the fixed `doomed` list of
    # element references afterward and deletes every one of them
    # unconditionally, so a moved-in-place original would simply be
    # deleted right back out, silently losing the bookmark again. A clone
    # sidesteps this entirely -- the original still dies with its range
    # exactly as `doomed` expects, and the surviving copy is a distinct
    # object the removal loop never touches.
    before = doomed[0].getprevious()
    for el in rescued:
        clone = deepcopy(el)
        if before is not None:
            before.addnext(clone)
        else:
            # doomed[0] was body's first child: insert at the front.
            body.insert(0, clone)
        before = clone  # keep multiple rescued markers in their original order


def remove_range(body, children, start_idx, end_idx):
    """Remove body children [start_idx, end_idx) with bookmark rescue."""
    doomed = children[start_idx:end_idx]
    preserve_half_bookmarks(body, doomed)
    for el in doomed:
        el.getparent().remove(el)
    return len(doomed)


def remove_elements(body, elements):
    """Remove arbitrary elements (paragraphs, tables, bookmark markers) with
    bookmark rescue.  Returns the count removed."""
    elements = [el for el in elements if el is not None and el.getparent() is not None]
    if not elements:
        return 0
    preserve_half_bookmarks(body, elements)
    for el in elements:
        par = el.getparent()
        if par is not None:
            par.remove(el)
    return len(elements)


def remove_paragraphs_where(body, pred):
    """Remove every body-level paragraph for which pred(p) is true."""
    doomed = [el for el in body if is_p(el) and pred(el)]
    return remove_elements(body, doomed)


def section_end(children, start_idx, stop_pred):
    """Index of the first child after start_idx satisfying stop_pred, or the
    index of the final sectPr (never past it)."""
    end = find_index(children, stop_pred, start_idx + 1)
    if end is None:
        end = len(children)
        while end > 0 and children[end - 1].tag == W + "sectPr":
            end -= 1
    return end


def collapse_empty_paragraphs(body, around, keep=1):
    """Collapse runs of consecutive empty paragraphs adjacent to `around`
    (an element still in the body, or None to use the whole body) down to
    `keep` paragraphs."""
    children = list(body)
    removed = 0
    if around is not None and around in children:
        i = children.index(around)
        lo = i
        while lo - 1 >= 0 and (is_empty_p(children[lo - 1])
                               or children[lo - 1].tag in (W + "bookmarkStart", W + "bookmarkEnd")):
            lo -= 1
        hi = i
        while hi + 1 < len(children) and (is_empty_p(children[hi + 1])
                                          or children[hi + 1].tag in (W + "bookmarkStart", W + "bookmarkEnd")):
            hi += 1
        empties = [c for c in children[lo:hi + 1] if is_empty_p(c)]
        doomed = empties[keep:]
        removed += remove_elements(body, doomed)
    return removed


def collapse_empty_runs(body, max_run=2, keep=1):
    """Everywhere in the body, shrink any run of >= max_run consecutive empty
    paragraphs to `keep`.  Bookmark markers between them are ignored (kept).
    Returns the number of paragraphs removed."""
    children = list(body)
    removed = 0
    run = []
    def flush():
        nonlocal removed
        if len(run) >= max_run:
            removed += remove_elements(body, run[keep:])
    for el in children:
        if is_empty_p(el):
            run.append(el)
        elif el.tag in (W + "bookmarkStart", W + "bookmarkEnd"):
            continue
        else:
            flush()
            run = []
    flush()
    return removed


def remove_orphan_list_labels(body, candidates):
    """After deleting list items, drop label-only parent items that lost all
    their children.  `candidates` are paragraphs that were the previous
    sibling of a deleted numbered item.  A candidate is removed when it is a
    numbered paragraph, has no sentence punctuation, and the next body
    element is not a deeper-level numbered paragraph.  Iterates to a fixed
    point.  Returns the count removed."""
    removed = 0
    changed = True
    cands = [c for c in candidates if c is not None]
    while changed:
        changed = False
        for p in list(cands):
            if p.getparent() is None or not is_p(p):
                cands.remove(p)
                continue
            np_ = p_numpr(p)
            text = norm(para_text(p))
            if np_ is None or not text or re.search(r"[.;!?]$", text) or len(text) > 60:
                cands.remove(p)
                continue
            nxt = p.getnext()
            while nxt is not None and nxt.tag in (W + "bookmarkStart", W + "bookmarkEnd"):
                nxt = nxt.getnext()
            nxt_np = p_numpr(nxt) if nxt is not None and is_p(nxt) else None
            if nxt_np is not None and int(nxt_np[0]) > int(np_[0]):
                continue  # still has a child
            prev = p.getprevious()
            while prev is not None and prev.tag in (W + "bookmarkStart", W + "bookmarkEnd"):
                prev = prev.getprevious()
            remove_elements(body, [p])
            cands.remove(p)
            removed += 1
            changed = True
            if prev is not None and is_p(prev) and prev not in cands:
                cands.append(prev)
    return removed


# --------------------------------------------------------------------------
# runs, text, fields
# --------------------------------------------------------------------------
def first_run_rpr(p):
    """rPr to reuse for new text in paragraph p: first run's rPr, else the
    paragraph mark rPr, else None."""
    r = p.find(W + "r")
    if r is not None and r.find(W + "rPr") is not None:
        return deepcopy(r.find(W + "rPr"))
    ppr = p.find(W + "pPr")
    if ppr is not None and ppr.find(W + "rPr") is not None:
        return deepcopy(ppr.find(W + "rPr"))
    return None


def make_run(text, rpr=None):
    r = E("r")
    if rpr is not None:
        r.append(deepcopy(rpr))
    t = E("t", text=text)
    if text != text.strip() or "  " in text:
        t.set(XML_SPACE, "preserve")
    r.append(t)
    return r


def clear_paragraph_content(p):
    """Remove everything from a paragraph except pPr (runs, hyperlinks, sdts,
    field runs, proofErr...).  Bookmarks inside are kept."""
    for child in list(p):
        if child.tag in (W + "pPr", W + "bookmarkStart", W + "bookmarkEnd"):
            continue
        p.remove(child)


def set_paragraph_text(p, text, rpr=None):
    """Replace the paragraph's content with a single run of `text`, reusing
    the first run's formatting unless rpr is given."""
    if rpr is None:
        rpr = first_run_rpr(p)
    clear_paragraph_content(p)
    p.append(make_run(text, rpr))
    return p


def cell_paragraphs(tc):
    return [c for c in tc if c.tag == W + "p"]


def set_cell_text(tc, text, keep_extra_paragraphs=False):
    """Set a table cell's text (first paragraph, single run).  Extra
    paragraphs are removed unless keep_extra_paragraphs."""
    ps = cell_paragraphs(tc)
    if not ps:
        p = E("p")
        tc.append(p)
        ps = [p]
    set_paragraph_text(ps[0], text)
    if not keep_extra_paragraphs:
        for extra in ps[1:]:
            tc.remove(extra)
    return ps[0]


def set_cell_lines(tc, lines):
    """One paragraph per line, all cloned from the cell's first paragraph."""
    ps = cell_paragraphs(tc)
    proto = deepcopy(ps[0]) if ps else E("p")
    for p in ps:
        tc.remove(p)
    lines = list(lines) or [""]
    for line in lines:
        p = deepcopy(proto)
        set_paragraph_text(p, line)
        tc.append(p)


def set_fee_cell(tc, text):
    """Alias kept for the SOW API: set a fee-table cell's plain text."""
    return set_cell_text(tc, text)


def row_cells(tr):
    """All w:tc of a row in order, INCLUDING cells wrapped in w:sdt (the SOW
    product table's checkbox cells).  Never use python-docx row.cells here."""
    return [tc for tc in tr.iter(W + "tc") if _nearest_row(tc) is tr]


def _nearest_row(el):
    anc = el.getparent()
    while anc is not None and anc.tag != W + "tr":
        anc = anc.getparent()
    return anc


def table_rows(tbl):
    return [tr for tr in tbl if tr.tag == W + "tr"]


def iter_fields(p):
    """Top-level complex fields in a paragraph (or any element containing
    runs).  Yields dicts: instr (str), result_runs (list of w:r between
    separate and end), sep (run holding separate or None), begin, end."""
    fields = []
    depth = 0
    cur = None
    for r in p.iter(W + "r"):
        fc = r.find(W + "fldChar")
        if fc is not None:
            kind = fc.get(W + "fldCharType")
            if kind == "begin":
                depth += 1
                if depth == 1:
                    cur = {"begin": r, "instr": "", "sep": None,
                           "result_runs": [], "end": None}
            elif kind == "separate":
                if depth == 1 and cur is not None:
                    cur["sep"] = r
            elif kind == "end":
                if depth == 1 and cur is not None:
                    cur["end"] = r
                    fields.append(cur)
                    cur = None
                depth = max(0, depth - 1)
            continue
        if cur is None:
            continue
        if depth == 1:
            it = r.find(W + "instrText")
            if it is not None and cur["sep"] is None:
                cur["instr"] += it.text or ""
            elif cur["sep"] is not None:
                cur["result_runs"].append(r)
    for f in fields:
        f["instr"] = f["instr"].strip()
    return fields


def set_field_result(container, text, instr_contains=None):
    """Write the cached result of the (first matching) complex field inside
    `container` (a w:tc, w:p or any element).  The field code is untouched
    so Word still recomputes it.  Returns True when a field was updated."""
    for p in ([container] if is_p(container) else container.iter(W + "p")):
        for f in iter_fields(p):
            if instr_contains and instr_contains not in f["instr"]:
                continue
            runs = f["result_runs"]
            if runs:
                first = runs[0]
                t = first.find(W + "t")
                if t is None:
                    t = E("t")
                    first.append(t)
                t.text = text
                for extra in runs[1:]:
                    for tt in extra.findall(W + "t"):
                        extra.remove(tt)
            else:
                rpr = f["sep"].find(W + "rPr") if f["sep"] is not None else None
                nr = make_run(text, rpr)
                f["sep"].addnext(nr)
            return True
    return False


def refresh_ref_caches(root, values):
    """Rewrite the cached result of every ` REF <name> ` field whose bookmark
    name is a key of `values` (name -> display text)."""
    n = 0
    for p in root.iter(W + "p"):
        for f in iter_fields(p):
            m = re.match(r"REF\s+(\S+)", f["instr"])
            if not m or m.group(1) not in values:
                continue
            if set_field_result(p, values[m.group(1)], instr_contains="REF " + m.group(1)):
                n += 1
    return n


def parse_amount(text):
    """'$1,250.00' / '-0.00' / '' -> float (0.0 when blank)."""
    s = norm(text).replace("$", "").replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else 0.0


def fmt_amount(x):
    """Word's `\\# "0.00"` picture: plain two decimals, no thousands sep."""
    return "%.2f" % x


# --------------------------------------------------------------------------
# tables (SOW fill API)
# --------------------------------------------------------------------------
def tbl_style(tbl):
    st = tbl.find(W + "tblPr/" + W + "tblStyle")
    return st.get(W + "val") if st is not None else None


def find_table(doc_or_body, style=None, header_startswith=None):
    """First w:tbl whose style id equals `style` and/or whose header cells
    start with the given texts.  `header_startswith` is a sequence of strings
    compared (normalised, case-insensitive) with the first cells of row 0."""
    body = doc_or_body.element.body if hasattr(doc_or_body, "element") else doc_or_body
    for tbl in body.iter(W + "tbl"):
        if style is not None and tbl_style(tbl) != style:
            continue
        if header_startswith:
            rows = table_rows(tbl)
            if not rows:
                continue
            cells = [norm(para_text(tc)).lower() for tc in row_cells(rows[0])]
            want = [norm(h).lower() for h in header_startswith]
            if len(cells) < len(want):
                continue
            if not all(c.startswith(w) for c, w in zip(cells, want)):
                continue
        return tbl
    return None


def table_label_paragraph(tbl):
    """The non-empty paragraph immediately above a table (its label), or None."""
    prev = tbl.getprevious()
    while prev is not None and is_empty_p(prev):
        prev = prev.getprevious()
    if prev is not None and is_p(prev) and para_text(prev).strip():
        return prev
    return None


def clear_sample_rows(tbl):
    """Keep the header row plus ONE empty data row (cloned from the first data
    row, all text removed).  Returns a deepcopy of the original first data row
    to use as a prototype for add_* calls."""
    rows = table_rows(tbl)
    if len(rows) < 2:
        return None
    proto = deepcopy(rows[1])
    for tr in rows[2:]:
        tbl.remove(tr)
    empty = rows[1]
    for tc in row_cells(empty):
        set_cell_text(tc, "")
    return proto


def _row_is_blank(tr):
    return not norm(para_text(tr))


def _append_data_row(tbl, new_tr):
    """Insert new_tr after the last data row, replacing a blank placeholder
    row if that is the only data row."""
    rows = table_rows(tbl)
    data = rows[1:]
    if len(data) == 1 and _row_is_blank(data[0]):
        data[0].addprevious(new_tr)
        tbl.remove(data[0])
    else:
        rows[-1].addnext(new_tr)
    return new_tr


def add_deliverable_row(tbl, values, prototype=None):
    """Append a row to the PSOGrid deliverables table.  `values` is a list of
    cell texts (MS: Scope Section, Milestone, Deliverable, Description,
    Format; TM: Scope Section, Deliverable, Description, Format).  Missing
    values are blank; extras are ignored."""
    rows = table_rows(tbl)
    proto = deepcopy(prototype) if prototype is not None else deepcopy(rows[-1])
    cells = row_cells(proto)
    if len(values) > len(cells):
        # A value list longer than the table's real column count almost
        # always means the wrong variant's column shape was used (e.g. the
        # 5-column Milestone shape against a 4-column T&M table) -- values
        # would otherwise fill left to right with the extras silently
        # dropped, quietly losing the last column instead of just "ignoring
        # extras" the way a genuinely spare value would.
        log("warning: add_deliverable_row got %d values for a %d-column table "
            "(wrong SOW variant column shape?); the extra value(s) are dropped: %r"
            % (len(values), len(cells), values[len(cells):]))
    for i, tc in enumerate(cells):
        set_cell_text(tc, str(values[i]) if i < len(values) and values[i] is not None else "")
    return _append_data_row(tbl, proto)


def add_timeline_row(tbl, phase, lines, duration, prototype=None):
    """Append a Phase / Description / Duration row to the NetwovenTable1
    timeline table.  `lines` is a list of description lines (one paragraph
    each, as in the template)."""
    rows = table_rows(tbl)
    proto = deepcopy(prototype) if prototype is not None else deepcopy(rows[-1])
    cells = row_cells(proto)
    if len(cells) < 3:
        raise ValueError("timeline table row has fewer than 3 cells")
    set_cell_text(cells[0], phase)
    set_cell_lines(cells[1], [lines] if isinstance(lines, str) else list(lines))
    set_cell_text(cells[2], duration)
    return _append_data_row(tbl, proto)


def add_table_row(tbl, values, prototype=None):
    """Generic: clone a data row and set plain cell texts."""
    return add_deliverable_row(tbl, values, prototype)


# ---- Milestone fee table (TableGrid, 4 columns) ---------------------------
def fee_rows_ms(tbl):
    """Classify the Milestone SOW fee table rows.  Returns a dict with
    header, milestones (list), spacer1, subtotal, discount, license, spacer2,
    total (each a w:tr or None)."""
    rows = table_rows(tbl)
    out = {"header": rows[0] if rows else None, "milestones": [],
           "spacer1": None, "subtotal": None, "discount": None,
           "license": None, "spacer2": None, "total": None}
    state = "ms"
    for tr in rows[1:]:
        cells = row_cells(tr)
        first = norm(para_text(cells[0])).lower() if cells else ""
        instr = " ".join(f["instr"] for f in iter_fields(tr))
        if first.startswith("sub-total") or first.startswith("sub total") or first.startswith("subtotal"):
            out["subtotal"] = tr
            state = "post"
        elif first.startswith("client discount"):
            out["discount"] = tr
        elif first.startswith("license cost"):
            out["license"] = tr
        elif first.startswith("total") and "SUM(ABOVE)" in instr:
            out["total"] = tr
        elif _row_is_blank(tr):
            if state == "ms":
                out["spacer1"] = tr
                state = "post"
            else:
                out["spacer2"] = tr
        elif state == "ms":
            out["milestones"].append(tr)
    return out


def reset_fee_rows_ms(tbl):
    """Remove all milestone rows (keeping =SUM(ABOVE) rows, bookmark TE, the
    Client Discount and License Cost rows).  Returns a deepcopy of the first
    milestone row to use as the prototype for add_fee_row_ms."""
    info = fee_rows_ms(tbl)
    proto = deepcopy(info["milestones"][0]) if info["milestones"] else None
    for tr in info["milestones"]:
        tbl.remove(tr)
    return proto


def add_fee_row_ms(tbl, name, description, amount, completion, prototype=None):
    """Insert a milestone row (Milestone, Description, Payment, Expected
    Completion) before the spacer row that precedes Sub-Total."""
    info = fee_rows_ms(tbl)
    if prototype is None:
        if info["milestones"]:
            prototype = deepcopy(info["milestones"][0])
        else:
            raise ValueError("no milestone row prototype available")
    tr = deepcopy(prototype)
    cells = row_cells(tr)
    values = [name, description,
              fmt_amount(parse_amount(str(amount))) if amount not in (None, "") else "",
              completion or ""]
    for i, tc in enumerate(cells[:4]):
        set_cell_text(tc, str(values[i]))
    anchor = info["spacer1"] if info["spacer1"] is not None else info["subtotal"]
    if anchor is None:
        table_rows(tbl)[-1].addnext(tr)
    else:
        anchor.addprevious(tr)
    return tr


def recompute_ms_totals(tbl, body=None):
    """Write cached Sub-Total and Total (=SUM(ABOVE) fields keep their codes),
    honouring the Client Discount / License Cost rows.  Refreshes any
    `REF TE` caches in `body`.  Returns the total as float."""
    info = fee_rows_ms(tbl)
    subtotal = sum(parse_amount(para_text(row_cells(tr)[2])) for tr in info["milestones"])
    if info["subtotal"] is not None:
        set_field_result(row_cells(info["subtotal"])[2], fmt_amount(subtotal))
    disc = parse_amount(para_text(row_cells(info["discount"])[2])) if info["discount"] is not None else 0.0
    lic = parse_amount(para_text(row_cells(info["license"])[2])) if info["license"] is not None else 0.0
    total = subtotal + disc + lic
    if info["total"] is not None:
        set_field_result(row_cells(info["total"])[2], fmt_amount(total))
    if body is not None:
        refresh_ref_caches(body, {"TE": fmt_amount(total)})
    return total


# ---- T&M rate table (TableGrid, 7 columns) --------------------------------
def _has_instr(tr, needle):
    return any(needle in f["instr"] for f in iter_fields(tr))


def rate_sections_tm(tbl):
    """Classify the T&M rate card.  Returns dict with onshore/offshore each
    {'rows': [product rows], 'estimate': tr}, plus onshore_ref, offshore_ref,
    subtotal, discount, license, total rows."""
    rows = table_rows(tbl)
    out = {"onshore": {"rows": [], "estimate": None},
           "offshore": {"rows": [], "estimate": None},
           "onshore_ref": None, "offshore_ref": None,
           "subtotal": None, "discount": None, "license": None, "total": None}
    section = None
    for tr in rows[1:]:
        cells = row_cells(tr)
        c0 = norm(para_text(cells[0])).lower() if cells else ""
        c1 = norm(para_text(cells[1])).lower() if len(cells) > 1 else ""
        if c0 == "onshore":
            section = "onshore"
        elif c0 == "offshore":
            section = "offshore"
        if _has_instr(tr, "PRODUCT(LEFT)") and section:
            out[section]["rows"].append(tr)
        elif c1.startswith("onshore estimate"):
            out["onshore"]["estimate"] = tr
        elif c1.startswith("offshore estimate"):
            out["offshore"]["estimate"] = tr
            section = None
        elif c1.startswith("onshore total"):
            out["onshore_ref"] = tr
        elif c1.startswith("offshore total"):
            out["offshore_ref"] = tr
        elif c1.startswith("sub total") or c1.startswith("sub-total") or c1.startswith("subtotal"):
            out["subtotal"] = tr
        elif c1.startswith("client discount"):
            out["discount"] = tr
        elif c1.startswith("license cost"):
            out["license"] = tr
        elif c1.startswith("total project estimate"):
            out["total"] = tr
    return out


def _role_key(s):
    return norm(s).rstrip("*").strip().lower()


def set_rate_row_tm(tbl, role, location, resources, weekly_hours, weeks, rate):
    """Fill (or add) a role row in the T&M rate card.  `location` is
    'onshore' or 'offshore'.  Numbers may be int/float/str.  The
    =PRODUCT(LEFT) field keeps its code; its cached result is written.
    Returns the row."""
    loc = (location or "").strip().lower()
    if loc not in ("onshore", "offshore"):
        raise ValueError("location must be onshore|offshore, got %r" % location)
    info = rate_sections_tm(tbl)
    sec = info[loc]
    key = _role_key(role)
    if not key:
        raise ValueError("role must not be blank")
    target = None
    for tr in sec["rows"]:
        k = _role_key(para_text(row_cells(tr)[1]))
        if k == key:
            target = tr
            break
    # No prefix fallback: _role_key already normalises case/whitespace/the
    # trailing "*", so an exact match after normalising is sufficient to
    # reuse a template row (e.g. "Consultant" matching "Consultant*"). A
    # prefix fallback risked two distinct roles where one name is a prefix
    # of the other (e.g. "Consultant" / "Consultant Lead") colliding into
    # one mislabeled row with the wrong numbers.
    if target is None:
        if not sec["rows"]:
            raise ValueError("no %s role rows to clone" % loc)
        target = deepcopy(sec["rows"][-1])
        cells = row_cells(target)
        set_cell_text(cells[0], "\xa0")
        set_cell_text(cells[1], role if role.endswith("*") else role + "*")
        anchor = sec["estimate"] if sec["estimate"] is not None else sec["rows"][-1]
        if sec["estimate"] is not None:
            anchor.addprevious(target)
        else:
            anchor.addnext(target)
    cells = row_cells(target)
    nums = [resources, weekly_hours, weeks, rate]
    if any(n in (None, "") for n in nums):
        # A field genuinely missing from the spec (None/"") is a gap, not a
        # zero -- writing "0" here would fabricate a number and a $0.00
        # total the source never supplied. Leave the whole row blank (the
        # four numbers form one product; a partial product is not a real
        # value either) so the gap stays visibly unfilled instead of
        # silently reading as "this role costs nothing", matching how
        # add_fee_row_ms leaves a missing Milestone amount blank rather
        # than writing a fabricated 0.00.
        for tc in cells[2:6]:
            set_cell_text(tc, "")
        set_field_result(cells[6], "")
        return target
    vals = [parse_amount(str(n)) for n in nums]
    texts = [("%d" % v) if float(v).is_integer() else ("%g" % v) for v in vals[:3]]
    texts.append(("%d" % vals[3]) if float(vals[3]).is_integer() else fmt_amount(vals[3]))
    for tc, t in zip(cells[2:6], texts):
        set_cell_text(tc, t)
    product = vals[0] * vals[1] * vals[2] * vals[3]
    set_field_result(cells[6], fmt_amount(product))
    return target


def drop_unused_rate_rows_tm(tbl):
    """Remove role rows whose resource count (or hours/weeks) is zero/blank.
    Section 'Onshore'/'Offshore' labels are moved to the first surviving row
    of the section.  Returns the count removed."""
    info = rate_sections_tm(tbl)
    removed = 0
    for loc in ("onshore", "offshore"):
        rows = info[loc]["rows"]
        keep = []
        for tr in rows:
            cells = row_cells(tr)
            used = all(parse_amount(para_text(tc)) > 0 for tc in cells[2:5])
            if used:
                keep.append(tr)
        label = None
        for tr in rows:
            c0 = norm(para_text(row_cells(tr)[0]))
            if c0.lower() == loc:
                label = c0
        for tr in rows:
            if tr not in keep:
                tbl.remove(tr)
                removed += 1
        if keep and label:
            # make sure the section label survives on the first kept row
            first_cells = row_cells(keep[0])
            if norm(para_text(first_cells[0])).lower() != loc:
                set_cell_text(first_cells[0], label)
            for tr in keep[1:]:
                c = row_cells(tr)[0]
                if norm(para_text(c)).lower() == loc:
                    set_cell_text(c, "\xa0")
    return removed


def recompute_tm_totals(tbl, body=None):
    """Write cached results for Onshore/Offshore Estimate (=SUM(ABOVE)),
    the two REF rows, Sub Total and Total Project Estimate; then refresh
    REF Onshore_Total / Offshore_Total / Total_Estimate caches in `body`
    (including the Executive Summary line).  Returns the total as float."""
    info = rate_sections_tm(tbl)
    sums = {}
    for loc in ("onshore", "offshore"):
        s = 0.0
        for tr in info[loc]["rows"]:
            cells = row_cells(tr)
            vals = [parse_amount(para_text(tc)) for tc in cells[2:6]]
            prod = vals[0] * vals[1] * vals[2] * vals[3]
            set_field_result(cells[6], fmt_amount(prod))
            s += prod
        sums[loc] = s
        if info[loc]["estimate"] is not None:
            set_field_result(row_cells(info[loc]["estimate"])[-1], fmt_amount(s))
    if info["onshore_ref"] is not None:
        set_field_result(row_cells(info["onshore_ref"])[-1], fmt_amount(sums["onshore"]))
    if info["offshore_ref"] is not None:
        set_field_result(row_cells(info["offshore_ref"])[-1], fmt_amount(sums["offshore"]))
    subtotal = sums["onshore"] + sums["offshore"]
    if info["subtotal"] is not None:
        set_field_result(row_cells(info["subtotal"])[-1], fmt_amount(subtotal))
    disc = parse_amount(para_text(row_cells(info["discount"])[-1])) if info["discount"] is not None else 0.0
    lic = parse_amount(para_text(row_cells(info["license"])[-1])) if info["license"] is not None else 0.0
    total = subtotal + disc + lic
    if info["total"] is not None:
        set_field_result(row_cells(info["total"])[-1], fmt_amount(total))
    if body is not None:
        refresh_ref_caches(body, {"Total_Estimate": fmt_amount(total),
                                  "Onshore_Total": fmt_amount(sums["onshore"]),
                                  "Offshore_Total": fmt_amount(sums["offshore"])})
    return total


# ---- product checkbox table ----------------------------------------------
def products_table(doc_or_body):
    tbl = find_table(doc_or_body, style="GridTable6Colorful-Accent5")
    if tbl is None:
        tbl = find_table(doc_or_body, header_startswith=["Microsoft Products"])
    return tbl


def _checkbox_sdt(tr):
    for sdt in tr.iter(W + "sdt"):
        if sdt.find(W + "sdtPr/" + W14 + "checkbox") is not None:
            return sdt
    return None


def product_rows(tbl):
    """[(row, name, checked_bool)] for every checkbox row."""
    out = []
    for tr in table_rows(tbl)[1:]:
        sdt = _checkbox_sdt(tr)
        if sdt is None:
            continue
        cells = row_cells(tr)
        name = norm(para_text(cells[0])) if cells else ""
        chk = sdt.find(W + "sdtPr/" + W14 + "checkbox/" + W14 + "checked")
        checked = chk is not None and chk.get(W14 + "val") in ("1", "true")
        out.append((tr, name, checked))
    return out


def _set_checkbox(tr, checked):
    sdt = _checkbox_sdt(tr)
    if sdt is None:
        return False
    cb = sdt.find(W + "sdtPr/" + W14 + "checkbox")
    chk = cb.find(W14 + "checked")
    if chk is None:
        chk = etree.SubElement(cb, W14 + "checked")
        cb.insert(0, chk)
    chk.set(W14 + "val", "1" if checked else "0")
    glyph = CHECKED_GLYPH if checked else UNCHECKED_GLYPH
    content = sdt.find(W + "sdtContent")
    ts = list(content.iter(W + "t"))
    if ts:
        ts[0].text = glyph
        for t in ts[1:]:
            t.getparent().remove(t)
    else:
        p = content.find(".//" + W + "p")
        if p is not None:
            p.append(make_run(glyph, first_run_rpr(p)))
    return True


def tick_product(doc_or_tbl, name, checked=True):
    """Tick (or untick) a product by name: sets w14:checked and the glyph
    (U+2612 / U+2610).  Match is case-insensitive on the normalised name,
    falling back to prefix match.  Returns True if a row was changed."""
    tbl = doc_or_tbl if isinstance(doc_or_tbl, etree._Element) and doc_or_tbl.tag == W + "tbl" \
        else products_table(doc_or_tbl)
    if tbl is None:
        return False
    key = norm(name).lower()
    if not key:
        return False
    rows = product_rows(tbl)
    # Exact (case-insensitive) match only: a prefix fallback risked ticking
    # an unrelated product whose name happens to start with (or be a
    # prefix of) the requested one, e.g. "SharePoint" matching
    # "SharePoint Migration Tool".
    for tr, rname, _ in rows:
        if rname.lower() == key:
            return _set_checkbox(tr, checked)
    return False


def drop_unticked_products(doc_or_tbl):
    """Remove every unticked checkbox row.  Returns count removed."""
    tbl = doc_or_tbl if isinstance(doc_or_tbl, etree._Element) and doc_or_tbl.tag == W + "tbl" \
        else products_table(doc_or_tbl)
    if tbl is None:
        return 0
    n = 0
    for tr, _, checked in product_rows(tbl):
        if not checked:
            tbl.remove(tr)
            n += 1
    return n


# ---- Biz Apps table ------------------------------------------------------
def biz_apps_table(doc_or_body):
    return find_table(doc_or_body, header_startswith=["Serial No", "Product Name"])


def set_biz_apps(doc_or_body, names):
    """Replace the Biz Apps rows with `names` (serial numbers regenerated).
    With an empty list the label paragraph + table are removed entirely.
    Returns the number of rows written (0 when removed / absent)."""
    body = doc_or_body.element.body if hasattr(doc_or_body, "element") else doc_or_body
    tbl = biz_apps_table(body)
    if tbl is None:
        return 0
    names = [n for n in (names or []) if str(n).strip()]
    if not names:
        label = table_label_paragraph(tbl)
        nxt = tbl.getnext()
        doomed = [tbl]
        if label is not None and norm(para_text(label)).lower() == "biz apps":
            doomed.append(label)
        if nxt is not None and is_empty_p(nxt) and label is not None:
            prev = label.getprevious()
            if prev is not None and is_empty_p(prev):
                doomed.append(nxt)
        remove_elements(body, doomed)
        return 0
    rows = table_rows(tbl)
    proto = deepcopy(rows[1])
    for tr in rows[1:]:
        tbl.remove(tr)
    for i, name in enumerate(names, start=1):
        tr = deepcopy(proto)
        cells = row_cells(tr)
        set_cell_text(cells[0], str(i))
        set_cell_text(cells[1], str(name))
        tbl.append(tr)
    return len(names)


# --------------------------------------------------------------------------
# client logo box (SOW cover)
# --------------------------------------------------------------------------
def find_logo_box(body):
    """The mc:AlternateContent holding the 'Rectangle: Rounded Corners 5'
    anchor (Choice) and its VML fallback, or None."""
    for ac in body.iter(MC + "AlternateContent"):
        for dp in ac.iter(WP + "docPr"):
            if dp.get("name") == LOGO_BOX_NAME:
                return ac
        for rr in ac.iter(V + "roundrect"):
            if rr.get("id") == LOGO_BOX_NAME:
                return ac
    return None


def remove_client_logo_box(doc):
    """Delete the 'Please Insert Client Logo' rounded box (both the
    Choice drawing and the Fallback VML).  Returns True if removed."""
    body = doc.element.body
    ac = find_logo_box(body)
    if ac is None:
        return False
    run = ac.getparent()
    ac.getparent().remove(ac)
    if run is not None and run.tag == W + "r" and not [c for c in run if c.tag != W + "rPr"]:
        run.getparent().remove(run)
    return True


def set_client_logo(doc, image_path):
    """Replace the logo box with the client's logo picture, scaled to fit the
    box (1390650 x 600075 EMU) preserving aspect ratio, right-aligned and
    vertically centred where the box was.  Returns True on success."""
    from docx.image.image import Image
    body = doc.element.body
    ac = find_logo_box(body)
    if ac is None:
        return False
    anchor = None
    for cand in ac.iter(WP + "anchor"):
        anchor = cand
        break
    if anchor is None:
        return False
    img = Image.from_file(image_path)
    rId, _ = doc.part.get_or_add_image(image_path)
    px_w, px_h = img.px_width, img.px_height
    scale = min(LOGO_BOX_CX / float(px_w), LOGO_BOX_CY / float(px_h))
    cx, cy = int(px_w * scale), int(px_h * scale)

    new_anchor = deepcopy(anchor)
    ext = new_anchor.find(WP + "extent")
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))
    eff = new_anchor.find(WP + "effectExtent")
    if eff is not None:
        for k in ("l", "t", "r", "b"):
            eff.set(k, "0")
    posv = new_anchor.find(WP + "positionV/" + WP + "posOffset")
    if posv is not None:
        try:
            posv.text = str(int(posv.text or "0") + (LOGO_BOX_CY - cy) // 2)
        except ValueError:
            pass
    docpr = new_anchor.find(WP + "docPr")
    docpr.set("name", "Client Logo")
    docpr.set("descr", "Client logo")
    # Build the picture graphic from literal XML so the a:/pic:/r: prefixes are
    # declared exactly the way Word writes them (lxml would otherwise invent
    # ns0/ns1 prefixes where no ancestor declares the drawingml namespaces).
    cnv = new_anchor.find(WP + "cNvGraphicFramePr")
    if cnv is not None:
        for c in list(cnv):
            cnv.remove(c)
        cnv.append(etree.fromstring(
            '<a:graphicFrameLocks xmlns:a="%s" noChangeAspect="1"/>' % NS["a"]))
    graphic = new_anchor.find(A + "graphic")
    new_graphic = etree.fromstring(
        '<a:graphic xmlns:a="{a}"><a:graphicData uri="{picns}">'
        '<pic:pic xmlns:pic="{picns}"><pic:nvPicPr><pic:cNvPr id="0" name="Client Logo"/>'
        '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="{r}" r:embed="{rid}"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic>'.format(a=NS["a"], picns=NS["pic"], r=NS["r"], rid=rId, cx=cx, cy=cy))
    graphic.addprevious(new_graphic)
    new_anchor.remove(graphic)
    # wp14 size-relative hints belong to shapes, not pictures
    for c in list(new_anchor):
        if c.tag.startswith("{http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing}"):
            new_anchor.remove(c)
    drawing = E("drawing")
    drawing.append(new_anchor)
    ac.addprevious(drawing)
    ac.getparent().remove(ac)
    return True


# --------------------------------------------------------------------------
# OLE timeline, hyperlinks, relationships
# --------------------------------------------------------------------------
def drop_rel(part, rId):
    """Remove a relationship from a part unconditionally (python-docx's
    drop_rel refuses when the rId is still referenced; callers here have
    already removed the referencing XML)."""
    rels = part.rels
    if rId in rels:
        del rels[rId]
        return True
    return False


def referenced_rids(root):
    """Every r:* attribute value in an XML tree (r:id, r:embed, r:dm...)."""
    out = set()
    for el in root.iter():
        for k, v in el.attrib.items():
            if k.startswith(R):
                out.add(v)
    return out


def drop_unreferenced_rels(part, reltypes=None):
    """Drop relationships of the given types (default: hyperlink, image,
    package, oleObject, diagram*) that nothing in the part's XML references
    any more.  Returns the list of dropped rIds."""
    used = referenced_rids(part.element)
    dropped = []
    for rId, rel in list(part.rels.items()):
        rt = rel.reltype
        ok_type = (reltypes is None and (rt in (RT_HYPERLINK, RT_IMAGE, RT_PACKAGE, RT_OLE)
                                         or rt.startswith(RT_DIAGRAM_PREFIXES))) \
            or (reltypes is not None and rt in reltypes)
        if ok_type and rId not in used:
            del part.rels[rId]
            dropped.append(rId)
    return dropped


def remove_ole_timeline(doc):
    """Remove the embedded Excel Gantt: the w:object paragraph, the _MON_*
    bookmarks, the EMF preview + package relationships, and collapse the
    surrounding empty paragraphs to one.  Returns list of dropped rIds."""
    body = doc.element.body
    dropped = []
    objs = [p for p in body.iter(W + "p") if p.find(".//" + W + "object") is not None]
    for p in objs:
        rids = referenced_rids(p)
        nxt = p.getnext()
        prev = p.getprevious()
        mons = [b for b in body.iter(W + "bookmarkStart") if (b.get(W + "name") or "").startswith("_MON_")]
        mon_ids = set(b.get(W + "id") for b in mons)
        mon_ends = [b for b in body.iter(W + "bookmarkEnd") if b.get(W + "id") in mon_ids]
        remove_elements(body, mons + mon_ends)
        remove_elements(body, [p])
        for rid in rids:
            if drop_rel(doc.part, rid):
                dropped.append(rid)
        # collapse empties around the hole
        marker = nxt if (nxt is not None and nxt.getparent() is not None) else prev
        if marker is not None and marker.getparent() is not None:
            if is_empty_p(marker):
                collapse_empty_paragraphs(body, marker, keep=1)
            else:
                # marker is content: collapse empties just before it
                before = marker.getprevious()
                if before is not None and is_empty_p(before):
                    collapse_empty_paragraphs(body, before, keep=1)
    return dropped


def remove_hyperlink_paragraph(body, rId):
    """Delete every body-level paragraph containing an external hyperlink
    with relationship id rId.  Returns count removed."""
    doomed = []
    for p in body:
        if not is_p(p):
            continue
        for h in p.iter(W + "hyperlink"):
            if h.get(R + "id") == rId:
                doomed.append(p)
                break
    return remove_elements(body, doomed)


def strip_external_hyperlink_rel(part, rId):
    return drop_rel(part, rId)


def external_hyperlink_rids(part, url_contains):
    """rIds of hyperlink relationships whose target contains url_contains."""
    out = []
    for rId, rel in part.rels.items():
        if rel.reltype == RT_HYPERLINK and rel.is_external and url_contains in rel.target_ref:
            out.append(rId)
    return out


# --------------------------------------------------------------------------
# highlighted author instructions
# --------------------------------------------------------------------------
def text_runs(p):
    return [r for r in p.iter(W + "r") if r.find(W + "t") is not None]


def is_highlighted(r):
    return r.find(W + "rPr/" + W + "highlight") is not None


def strip_highlight(r):
    hl = r.find(W + "rPr/" + W + "highlight")
    if hl is not None:
        hl.getparent().remove(hl)


def strip_highlighted_instructions(body, rules=None):
    """Handle every body-level paragraph that contains yellow-highlighted
    author instructions.

    * All text runs highlighted -> the paragraph is deleted.
    * Otherwise the first matching rule applies.  A rule is a dict:
        {"contains": "substring of paragraph text",
         "action": "delete" | "empty_slot" | "replace_highlighted",
         "style": "BodyText"   # empty_slot: paragraph style for the slot
         "text": "10%"}        # replace_highlighted: replacement text
    * No rule -> highlight removed, text kept, a warning is logged.
    Returns a list of (action, text_excerpt) tuples."""
    rules = rules or []
    actions = []
    for p in [el for el in list(body) if is_p(el)]:
        runs = text_runs(p)
        hl = [r for r in runs if is_highlighted(r)]
        if not hl:
            continue
        text = norm(para_text(p))
        if len(hl) == len(runs):
            remove_elements(body, [p])
            actions.append(("delete", text[:60]))
            continue
        rule = None
        for rl in rules:
            if rl.get("contains") and rl["contains"].lower() in text.lower():
                rule = rl
                break
        if rule is None:
            for r in hl:
                strip_highlight(r)
            actions.append(("unhighlight-kept", text[:60]))
            log("warning: partially highlighted paragraph kept without a rule: %r" % text[:80])
            continue
        act = rule.get("action", "delete")
        if act == "delete":
            remove_elements(body, [p])
            actions.append(("delete", text[:60]))
        elif act == "empty_slot":
            clear_paragraph_content(p)
            ppr = p.find(W + "pPr")
            if ppr is None:
                ppr = E("pPr")
                p.insert(0, ppr)
            for k in ("numPr", "ind", "jc"):
                el = ppr.find(W + k)
                if el is not None:
                    ppr.remove(el)
            st = ppr.find(W + "pStyle")
            style = rule.get("style", "BodyText")
            if st is None:
                st = E("pStyle")
                ppr.insert(0, st)
            st.set(W + "val", style)
            nxt = p.getnext()
            if nxt is not None and is_empty_p(nxt):
                remove_elements(body, [nxt])
            actions.append(("empty_slot", text[:60]))
        elif act == "replace_highlighted":
            rpr = hl[0].find(W + "rPr")
            rpr = deepcopy(rpr) if rpr is not None else None
            if rpr is not None:
                h = rpr.find(W + "highlight")
                if h is not None:
                    rpr.remove(h)
            new = make_run(rule.get("text", ""), rpr)
            hl[0].addprevious(new)
            for r in hl:
                r.getparent().remove(r)
            actions.append(("replace_highlighted", text[:60]))
        else:
            raise ValueError("unknown highlight rule action %r" % act)
    return actions


# --------------------------------------------------------------------------
# TOC
# --------------------------------------------------------------------------
def toc_sdt(body):
    for sdt in body.iter(W + "sdt"):
        gallery = sdt.find(W + "sdtPr/" + W + "docPartObj/" + W + "docPartGallery")
        if gallery is not None and gallery.get(W + "val") == "Table of Contents":
            return sdt
    return None


def toc_heading_text(body):
    sdt = toc_sdt(body)
    if sdt is None:
        return None
    content = sdt.find(W + "sdtContent")
    for p in content.iter(W + "p"):
        return norm(para_text(p))
    return None


def reset_toc(body, heading_text):
    """Replace the TOC's cached entries (they list deleted template
    headings) with the template's own heading paragraph + a dirty
    `TOC \\o "1-3" \\h \\z \\u` field that Word rebuilds on open.  The heading
    pPr is deep-copied from the first paragraph inside the TOC sdt; numPr
    numId=0 is added only if the copied pPr lacks it (without it the heading
    would join the heading outline numbering)."""
    sdt = toc_sdt(body)
    if sdt is None:
        return False
    content = sdt.find(W + "sdtContent")
    if content is None:
        return False
    first_p = None
    for p in content.iter(W + "p"):
        first_p = p
        break
    ppr = deepcopy(first_p.find(W + "pPr")) if first_p is not None and first_p.find(W + "pPr") is not None else None
    run_rpr = first_run_rpr(first_p) if first_p is not None else None
    if ppr is None:
        ppr = E("pPr")
        ppr.append(E("pStyle", {"val": "TOCHeading"}))
    if ppr.find(W + "numPr") is None:
        npr = E("numPr")
        npr.append(E("ilvl", {"val": "0"}))
        npr.append(E("numId", {"val": "0"}))
        st = ppr.find(W + "pStyle")
        if st is not None:
            st.addnext(npr)
        else:
            ppr.insert(0, npr)
    for child in list(content):
        content.remove(child)
    h = E("p")
    h.append(ppr)
    h.append(make_run(heading_text, run_rpr))
    content.append(h)
    p = E("p")
    ppr2 = E("pPr")
    ppr2.append(E("pStyle", {"val": "TOC1"}))
    p.append(ppr2)
    r1 = E("r")
    r1.append(E("fldChar", {"fldCharType": "begin", "dirty": "true"}))
    p.append(r1)
    r2 = E("r")
    instr = E("instrText", text=' TOC \\o "1-3" \\h \\z \\u ')
    instr.set(XML_SPACE, "preserve")
    r2.append(instr)
    p.append(r2)
    r3 = E("r")
    r3.append(E("fldChar", {"fldCharType": "separate"}))
    p.append(r3)
    p.append(make_run("The table of contents fills in automatically when this document opens in Word."))
    r5 = E("r")
    r5.append(E("fldChar", {"fldCharType": "end"}))
    p.append(r5)
    content.append(p)
    return True


def reset_figure_list(body, kind, heading_text):
    """Like reset_toc, but for the plain-body "List of Figures" / "List of
    Tables" sections (TableofFigures-styled paragraphs, not inside an sdt).
    Their cached entries name the template's own demo figures/tables; once
    that demo content is removed the cache is stale and a non-Word viewer
    (LibreOffice, a PDF export) would show it as-is. Replace the whole
    TableofFigures paragraph group under `heading_text` with one dirty
    `TOC \\h \\z \\c "kind"` field, matching the template's own field
    switches, so it reads as a single placeholder line until Word rebuilds
    it on open. `kind` is "Figure" or "Table". Returns True if a group was
    found and reset, False if `heading_text` is absent (nothing to do)."""
    children = list(body)
    start = find_index(children, lambda el: is_p(el) and norm(para_text(el)) == heading_text)
    if start is None:
        return False
    group_start = start + 1
    end = group_start
    while end < len(children):
        el = children[end]
        st = el.find(f"{W}pPr/{W}pStyle")
        if el.tag != W + "p" or st is None or st.get(W + "val") != "TableofFigures":
            break
        end += 1
    if end == group_start:
        return False
    first_p = children[group_start]
    run_rpr = first_run_rpr(first_p)
    heading_para = children[start]  # never removed: safe insertion anchor
    for el in children[group_start:end]:
        body.remove(el)
    p = E("p")
    ppr = E("pPr")
    ppr.append(E("pStyle", {"val": "TableofFigures"}))
    p.append(ppr)
    r1 = E("r")
    r1.append(E("fldChar", {"fldCharType": "begin", "dirty": "true"}))
    p.append(r1)
    r2 = E("r")
    instr = E("instrText", text=' TOC \\h \\z \\c "%s" ' % kind)
    instr.set(XML_SPACE, "preserve")
    r2.append(instr)
    p.append(r2)
    r3 = E("r")
    r3.append(E("fldChar", {"fldCharType": "separate"}))
    p.append(r3)
    p.append(make_run("The list of %ss fills in automatically when this document opens in Word." % kind.lower(),
                       run_rpr))
    r5 = E("r")
    r5.append(E("fldChar", {"fldCharType": "end"}))
    p.append(r5)
    heading_para.addnext(p)
    return True


# --------------------------------------------------------------------------
# numbering
# --------------------------------------------------------------------------
def _story_roots(doc):
    """lxml roots of document + headers/footers/footnotes/endnotes."""
    roots = [doc.element]
    for part in doc.part.package.iter_parts():
        pn = str(part.partname)
        if re.search(r"/word/(header|footer|footnotes|endnotes)\d*\.xml$", pn) and hasattr(part, "element"):
            roots.append(part.element)
    return roots


def _num_kind(numbering, num):
    """'bullet' / 'decimal' / other numFmt of level 0 of a w:num."""
    a = num.find(W + "abstractNumId")
    if a is None:
        return None
    for ab in numbering.findall(W + "abstractNum"):
        if ab.get(W + "abstractNumId") == a.get(W + "val"):
            lvl0 = ab.find(W + "lvl")
            if lvl0 is None:
                return None
            fmt = lvl0.find(W + "numFmt")
            return fmt.get(W + "val") if fmt is not None else None
    return None


def _numids_used_by_styles(doc):
    """The raw set of numIds any style references for its own numPr (e.g.
    the NW Heading styles' own numId=2). Shared by list_num_id() and
    prune_numbering() so "which numIds a style owns" has one definition;
    each caller still applies its own further logic on top (list_num_id
    treats any style-referenced numId as unsafe to hand out full stop;
    prune_numbering additionally lets one through if a body paragraph is
    also already using it, since that makes it a real, present-tense body
    list rather than purely a style's private counter)."""
    return {ni.get(W + "val") for ni in doc.styles.element.iter(W + "numId")}


def list_num_id(doc, kind="bullet"):
    """numId of a list definition whose level 0 is `kind` ('bullet' or
    'decimal'), for skills that add lists with python-docx: put
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="ID"/></w:numPr> on a
    'List Paragraph' paragraph.  Never returns a numId a STYLE already owns
    for its own auto-numbering (e.g. the NW Heading section-number list) --
    attaching a body paragraph to that numId would splice it into the
    style's own sequence instead of numbering independently. Falls back to
    a style-owned numId only if truly nothing else of that kind exists (so
    callers can still detect the collision and build their own list, rather
    than getting None with no numId at all). None if the document has no
    list of that kind whatsoever."""
    try:
        numbering = doc.part.numbering_part.element
    except (AttributeError, KeyError):
        return None
    style_only = _numids_used_by_styles(doc)
    fallback = None
    for num in numbering.findall(W + "num"):
        if _num_kind(numbering, num) != kind:
            continue
        nid = num.get(W + "numId")
        if nid in style_only:
            if fallback is None:
                fallback = nid
            continue
        return nid
    return fallback


def prune_numbering(doc, keep_kinds=("bullet", "decimal")):
    """Remove w:num entries no paragraph or style references and then
    w:abstractNum entries no surviving w:num uses (keeping any with a
    styleLink/numStyleLink relation).  One unreferenced list of each kind in
    `keep_kinds` (by level-0 numFmt) is kept so skills can still add bullet /
    numbered lists via list_num_id().  Returns (nums_removed, abstract_removed).

    A numId used only via a STYLE's own numPr (e.g. numId=2, which the NW
    Heading styles own for section-number auto-numbering) is never offered as
    that "kept alive" generic list: attaching a body paragraph to a
    style-owned numId splices it into that style's own numbering sequence
    instead of giving it independent 1/2/3 numbers. Such a numId still
    counts as "used" for deletion purposes (it must never be pruned), it
    just cannot satisfy keep_kinds on style-ownership alone."""
    try:
        numbering = doc.part.numbering_part.element
    except (AttributeError, KeyError):
        return (0, 0)
    used = set()
    used_by_para = set()
    for root in _story_roots(doc):
        for ni in root.iter(W + "numId"):
            if ni.getparent().tag == W + "numPr":
                val = ni.get(W + "val")
                used.add(val)
                used_by_para.add(val)
    styled = _numids_used_by_styles(doc)
    used |= styled
    style_only = styled - used_by_para
    used.discard("0")
    used_by_para.discard("0")
    have_kinds = set(_num_kind(numbering, n) for n in numbering.findall(W + "num")
                     if n.get(W + "numId") in used_by_para)
    for kind in keep_kinds or ():
        if kind in have_kinds:
            continue
        for num in numbering.findall(W + "num"):
            nid = num.get(W + "numId")
            if nid in style_only:
                continue  # never adopt a style-owned list as the generic pick
            if _num_kind(numbering, num) == kind:
                used.add(nid)
                break
    nums_removed = 0
    for num in list(numbering.findall(W + "num")):
        if num.get(W + "numId") not in used:
            numbering.remove(num)
            nums_removed += 1
    live_abs = set()
    for num in numbering.findall(W + "num"):
        a = num.find(W + "abstractNumId")
        if a is not None:
            live_abs.add(a.get(W + "val"))
    linked = set()
    for ab in numbering.findall(W + "abstractNum"):
        if ab.find(W + "styleLink") is not None or ab.find(W + "numStyleLink") is not None:
            linked.add(ab.get(W + "abstractNumId"))
    abs_removed = 0
    for ab in list(numbering.findall(W + "abstractNum")):
        aid = ab.get(W + "abstractNumId")
        if aid not in live_abs and aid not in linked:
            numbering.remove(ab)
            abs_removed += 1
    return (nums_removed, abs_removed)


# --------------------------------------------------------------------------
# dates
# --------------------------------------------------------------------------
DATE_FORMATS = ("%Y-%m-%d", "%m/%d/%Y", "%B %d, %Y", "%d %B %Y", "%B %Y", "%b %Y",
                "%b %d, %Y", "%d/%m/%Y", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M:%S")


def parse_date(s):
    """Parse a user date string; returns a datetime.date or None."""
    if not s:
        return None
    s = s.strip()
    if s.lower() == "today":
        return _dt.date.today()
    for fmt in DATE_FORMATS:
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def to_iso_datetime(s):
    """'2026-09-01' / 'September 2026' / '9/1/2026' -> '2026-09-01T00:00:00Z'
    (the bound coverPageProps node is storeMappedDataAs=dateTime).  Unparsable
    strings are returned unchanged with a note."""
    d = parse_date(s)
    if d is None:
        if s:
            log("note: could not parse date %r; storing as-is" % s)
        return s
    return d.strftime("%Y-%m-%dT00:00:00Z")


def date_display(s):
    """Display form used by the date SDTs: M/d/yyyy (no leading zeros)."""
    d = parse_date(s) if not isinstance(s, _dt.date) else s
    if d is None:
        return s
    return "%d/%d/%d" % (d.month, d.day, d.year)


# --------------------------------------------------------------------------
# zip-level patches (parts python-docx does not expose or re-serialise)
# --------------------------------------------------------------------------
def rewrite_zip(path, fn):
    """Rewrite every entry of the zip at `path` through fn(name, bytes) ->
    bytes (return the input unchanged to keep an entry)."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            new = fn(item.filename, data)
            zout.writestr(item, data if new is None else new)
    shutil.move(tmp, path)


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _sdt_binding_kind(sdt):
    db = sdt.find(W + "sdtPr/" + W + "dataBinding")
    if db is None:
        return None
    xp = db.get(W + "xpath") or ""
    if xp.endswith(":title[1]") or xp.endswith("/title[1]"):
        return "title"
    if xp.endswith(":Company[1]") or xp.endswith("/Company[1]"):
        return "company"
    if xp.endswith(":PublishDate[1]") or xp.endswith("/PublishDate[1]"):
        return "date"
    return None


def _set_sdt_cached_text(sdt, value, iso=None):
    pr = sdt.find(W + "sdtPr")
    plc = pr.find(W + "showingPlcHdr")
    if plc is not None:
        pr.remove(plc)
    if iso and pr.find(W + "date") is not None:
        pr.find(W + "date").set(W + "fullDate", iso)
    content = sdt.find(W + "sdtContent")
    if content is None:
        return
    runs = [r for r in content.iter(W + "r")
            if r.find(W + "t") is not None or r.find(W + "tab") is not None]
    sdt_rpr = pr.find(W + "rPr")
    if runs:
        first = runs[0]
        rpr = first.find(W + "rPr")
        if rpr is not None and rpr.find(W + "rStyle") is not None \
                and rpr.find(W + "rStyle").get(W + "val") == "PlaceholderText":
            if sdt_rpr is not None:
                first.replace(rpr, deepcopy(sdt_rpr))
            else:
                rpr.remove(rpr.find(W + "rStyle"))
        t = first.find(W + "t")
        if t is None:
            t = E("t")
            first.append(t)
        t.text = value
        if value != value.strip():
            t.set(XML_SPACE, "preserve")
        for extra in first.findall(W + "t")[1:]:
            first.remove(extra)
        for r in runs[1:]:
            r.getparent().remove(r)
    else:
        p = None
        for cand in content.iter(W + "p"):
            p = cand
        target = p if p is not None else content
        target.append(make_run(value, sdt_rpr))


def set_bound_fields(path, title=None, company=None, date=None):
    """Zip-level pass after doc.save(): for every bound SDT in document,
    header* and footer* parts write the cached text (title / company / date
    display), drop showingPlcHdr, set w:date/@w:fullDate; patch core.xml
    dc:title, app.xml Company + TitlesOfParts, customXml item PublishDate
    (self-closing or populated).  `date` may be any parseable date string;
    display is M/d/yyyy, stored value ISO dateTime.  Returns counts."""
    iso = to_iso_datetime(date) if date else None
    disp = date_display(date) if date else None
    counts = {"title": 0, "company": 0, "date": 0}

    def patch(name, data):
        if re.match(r"word/(document|header\d*|footer\d*)\.xml$", name):
            root = etree.fromstring(data)
            changed = False
            for sdt in root.iter(W + "sdt"):
                kind = _sdt_binding_kind(sdt)
                if kind == "title" and title is not None:
                    _set_sdt_cached_text(sdt, title)
                elif kind == "company" and company is not None:
                    _set_sdt_cached_text(sdt, company)
                elif kind == "date" and disp is not None:
                    _set_sdt_cached_text(sdt, disp, iso)
                else:
                    continue
                counts[kind] += 1
                changed = True
            return _serialize(root) if changed else data
        if name == "docProps/core.xml" and title is not None:
            text = data.decode("utf-8")
            if re.search(r"<dc:title\b[^>]*/>", text):
                text = re.sub(r"<dc:title\b[^>]*/>", lambda m: "<dc:title>%s</dc:title>" % xml_escape(title), text)
            elif "<dc:title" in text:
                text = re.sub(r"<dc:title>.*?</dc:title>", lambda m: "<dc:title>%s</dc:title>" % xml_escape(title), text, flags=re.S)
            else:
                text = text.replace("</cp:coreProperties>", "<dc:title>%s</dc:title></cp:coreProperties>" % xml_escape(title))
            return text.encode("utf-8")
        if name == "docProps/app.xml":
            text = data.decode("utf-8")
            if company is not None:
                if re.search(r"<Company\b[^>]*/>", text):
                    text = re.sub(r"<Company\b[^>]*/>", lambda m: "<Company>%s</Company>" % xml_escape(company), text)
                elif "<Company>" in text:
                    text = re.sub(r"<Company>.*?</Company>", lambda m: "<Company>%s</Company>" % xml_escape(company), text, flags=re.S)
                else:
                    text = text.replace("</Properties>", "<Company>%s</Company></Properties>" % xml_escape(company))
            if title is not None:
                text = re.sub(r"(<TitlesOfParts>\s*<vt:vector[^>]*>\s*<vt:lpstr>).*?(</vt:lpstr>)",
                              lambda m: m.group(1) + xml_escape(title) + m.group(2), text, count=1, flags=re.S)
            return text.encode("utf-8")
        if re.match(r"customXml/item\d+\.xml$", name) and iso:
            text = data.decode("utf-8")
            if "PublishDate" in text and "coverPageProps" in text:
                text = re.sub(r"<((?:\w+:)?PublishDate)\s*/>", r"<\g<1>>%s</\g<1>>" % iso, text)
                text = re.sub(r"(<(?:\w+:)?PublishDate(?:\s[^>/]*)?>).*?(</(?:\w+:)?PublishDate>)",
                              r"\g<1>%s\g<2>" % iso, text, flags=re.S)
                return text.encode("utf-8")
        return data

    rewrite_zip(path, patch)
    return counts


def patch_settings_updatefields(path):
    """Ensure word/settings.xml carries <w:updateFields w:val="true"/>
    (ECMA order: before hdrShapeDefaults/footnotePr/endnotePr/compat)."""
    def patch(name, data):
        if name == "word/settings.xml" and b"updateFields" not in data:
            text = data.decode("utf-8")
            for anchor in ("<w:hdrShapeDefaults", "<w:footnotePr", "<w:endnotePr", "<w:compat"):
                idx = text.find(anchor)
                if idx != -1:
                    return (text[:idx] + '<w:updateFields w:val="true"/>' + text[idx:]).encode("utf-8")
            return text.replace("</w:settings>", '<w:updateFields w:val="true"/></w:settings>').encode("utf-8")
        return data
    rewrite_zip(path, patch)


def patch_footer_dateformat(path, fmt="M/d/yyyy"):
    """Give every footer w:date SDT without a dateFormat the template's
    cover format so renderers do not show a raw serial."""
    def patch(name, data):
        if re.match(r"word/(footer|header)\d*\.xml$", name) and (b"<w:date>" in data or b"<w:date " in data):
            root = etree.fromstring(data)
            changed = False
            for d in root.iter(W + "date"):
                if d.find(W + "dateFormat") is None:
                    df = E("dateFormat", {"val": fmt})
                    d.insert(0, df)
                    changed = True
            return _serialize(root) if changed else data
        return data
    rewrite_zip(path, patch)


def scrub_metadata(path, creator="Netwoven"):
    """core.xml: creator/lastModifiedBy -> Netwoven, drop lastPrinted,
    created/modified -> now, revision 1, subject/keywords/description
    emptied.  app.xml: TotalTime 0."""
    now = _dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    def patch(name, data):
        if name == "docProps/core.xml":
            text = data.decode("utf-8")
            text = re.sub(r"<cp:lastPrinted>.*?</cp:lastPrinted>", "", text, flags=re.S)
            text = re.sub(r"<cp:lastPrinted\b[^>]*/>", "", text)
            for tag in ("dc:creator", "cp:lastModifiedBy"):
                if re.search(r"<%s\b[^>]*/>" % tag, text):
                    text = re.sub(r"<%s\b[^>]*/>" % tag, lambda m: "<%s>%s</%s>" % (tag, xml_escape(creator), tag), text)
                elif "<%s>" % tag in text:
                    text = re.sub(r"<%s>.*?</%s>" % (tag, tag), lambda m: "<%s>%s</%s>" % (tag, xml_escape(creator), tag), text, flags=re.S)
                else:
                    text = text.replace("</cp:coreProperties>", "<%s>%s</%s></cp:coreProperties>" % (tag, xml_escape(creator), tag))
            for tag in ("dcterms:created", "dcterms:modified"):
                text = re.sub(r"(<%s\b[^>]*>).*?(</%s>)" % (tag, tag), r"\g<1>%s\g<2>" % now, text, flags=re.S)
            text = re.sub(r"<cp:revision>.*?</cp:revision>", "<cp:revision>1</cp:revision>", text, flags=re.S)
            for tag in ("dc:subject", "cp:keywords", "dc:description"):
                text = re.sub(r"<%s>.*?</%s>" % (tag, tag), "<%s/>" % tag, text, flags=re.S)
            return text.encode("utf-8")
        if name == "docProps/app.xml":
            text = data.decode("utf-8")
            text = re.sub(r"<TotalTime>.*?</TotalTime>", "<TotalTime>0</TotalTime>", text, flags=re.S)
            return text.encode("utf-8")
        return data
    rewrite_zip(path, patch)


def zip_names(path):
    with zipfile.ZipFile(path) as z:
        return z.namelist()


def zip_read(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name)


# --------------------------------------------------------------------------
# python-docx content helpers for skills (figures, tables)
# --------------------------------------------------------------------------
def _next_seq_number(doc, seq_name):
    n = 0
    for p in doc.element.body.iter(W + "p"):
        for f in iter_fields(p):
            if re.match(r"SEQ\s+%s\b" % re.escape(seq_name), f["instr"]):
                n += 1
    return n + 1


def add_caption(doc, seq_name, text, style="Caption"):
    """Append a Caption paragraph 'Figure N text' / 'Table N text' with a
    live SEQ field (cached number computed from existing SEQ fields)."""
    number = _next_seq_number(doc, seq_name)
    para = doc.add_paragraph(style=style)
    p = para._p
    p.append(make_run(seq_name + " "))
    r1 = E("r")
    r1.append(E("fldChar", {"fldCharType": "begin", "dirty": "true"}))
    p.append(r1)
    r2 = E("r")
    it = E("instrText", text=" SEQ %s \\* ARABIC " % seq_name)
    it.set(XML_SPACE, "preserve")
    r2.append(it)
    p.append(r2)
    r3 = E("r")
    r3.append(E("fldChar", {"fldCharType": "separate"}))
    p.append(r3)
    rr = make_run(str(number))
    rpr = E("rPr")
    rpr.append(E("noProof"))
    rr.insert(0, rpr)
    p.append(rr)
    r5 = E("r")
    r5.append(E("fldChar", {"fldCharType": "end"}))
    p.append(r5)
    p.append(make_run(" " + text))
    return para


def add_captioned_picture(doc, image_path, caption, max_width_in=6.5, style="Caption"):
    """Centred picture no wider than max_width_in (aspect preserved) followed
    by a 'Figure N caption' paragraph in the Caption style with a SEQ field.
    Returns (picture_paragraph, caption_paragraph)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.image.image import Image
    from docx.shared import Emu, Inches
    img = Image.from_file(image_path)
    width_emu = min(Inches(max_width_in), Emu(img.width))
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(image_path, width=width_emu)
    cap = add_caption(doc, "Figure", caption, style=style)
    return para, cap


def add_styled_table(doc, header, rows, style="Netwoven Table 1", caption=None,
                     col_widths_in=None):
    """Append a table in a Netwoven table style with a bold header row (the
    style handles the look) and optional 'Table N caption' below it, as the
    template does.  Returns the python-docx Table."""
    from docx.shared import Inches
    ncols = len(header)
    table = doc.add_table(rows=1, cols=ncols)
    try:
        table.style = doc.styles[style]
    except KeyError:
        log("warning: table style %r not in this document; leaving default" % style)
    for i, h in enumerate(header):
        table.rows[0].cells[i].text = str(h)
    for row in rows:
        cells = table.add_row().cells
        for i in range(ncols):
            cells[i].text = str(row[i]) if i < len(row) and row[i] is not None else ""
    if col_widths_in:
        for row in table.rows:
            for i, w in enumerate(col_widths_in[:ncols]):
                row.cells[i].width = Inches(w)
    if caption:
        add_caption(doc, "Table", caption)
    return table


__all__ = [n for n in dir() if not n.startswith("_")]
