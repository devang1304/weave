#!/usr/bin/env python3
"""
Validate a built Netwoven deliverable (.docx document, .docx SOW, .pptx deck)
before handing it over.

  validate_deliverable.py FILE [--type doc|sow|deck] [--internal]
      [--client NAME] [--render] [--png-dir DIR] [--json]

Type is auto-detected when omitted: .pptx -> deck; .docx whose styles carry
"PSO Grid" -> sow; else doc.  Text checks run over the raw XML of every story
part (document, headers, footers) plus docProps/app.xml, core.xml,
custom.xml, so text inside content controls, text boxes and split runs is
seen.  Hard checks fail the run (exit 1); warnings only report.

--render converts through LibreOffice (soffice discovered on PATH or in
/Applications/LibreOffice.app) and reports the page count via pdfinfo;
--png-dir also writes page PNGs via pdftoppm.  A missing soffice is a
warning, never a failure.  Decks are delegated to
nw_pptx_helpers.validate_deck (imported lazily, so this file also ships in
skills that have no pptx helper).

--json prints {"file","type","pages","passed","results":[{"check","passed",
"evidence","hard"}]}.  Requires python-docx / lxml (python-pptx for decks).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import nw_docx_helpers as H  # noqa: E402 - after sys.path setup

# Shared with every builder script, so a fix to namespace values or to
# para_text/norm/p_style's matching semantics reaches the validator too
# instead of silently diverging in a second, hand-copied implementation.
W = H.W
W14 = H.W14
R = H.R
para_text = H.para_text
norm = H.norm
p_style = H.p_style

# (label, regex, flags) -- case-sensitive word matches for proper nouns that
# would otherwise hit ordinary words (Intel vs intelligence).
FORBIDDEN = [
    ("Intel", r"\bIntel\b", 0),
    ("Ross Stores", r"Ross Stores", re.I),
    ("Cypress", r"\bCypress\b", 0),
    ("sharepoint.com", r"sharepoint\.com", re.I),
    ("Sendnya", r"Sendnya", re.I),
    ("Segue UI", r"Segue UI", re.I),
    ("How to use this template", r"How to use this template", re.I),
    ("Sample Normal Text", r"Sample Normal Text", re.I),
    ("Company Name Here", r"Company Name Here", re.I),
    ("Document Title Here", r"Document Title Here", re.I),
    ("MM/DD/YYYY", r"MM/DD/YYYY", 0),
    ("THIS SECTION HAS BEEN CREATED AS A GUIDELINE", r"THIS SECTION HAS BEEN CREATED AS A GUIDELINE", re.I),
    ("delete this line", r"delete this line", re.I),
    ("right click to refresh", r"right click to refresh", re.I),
    ("Update Field", r"Update Field\b", 0),
    ("Please Insert Client Logo", r"Please Insert Client Logo", re.I),
    ("Click on the placeholder logo", r"Click on the placeholder logo", re.I),
    ("delete everything within Section 2", r"delete everything within Section 2", re.I),
]
DEFAULT_TITLES = {"", "document title here", "<<project name>>", "title"}
DEFAULT_COMPANIES = {"", "company name here", "<<client>>"}
SOW_REQUIRED_H1 = ["Glossary", "Executive Summary", "Microsoft Products Utilized in Project",
                   "Scope of Request and Deliverables", "Assumptions", "Project Operations",
                   "Budget and Timeline", "Signatures"]
SOW_TABLE_STYLES = {"TableGrid", "PSOGrid", "GridTable6Colorful-Accent5",
                    "NetwovenTable1", "NetwovenTable2", "NetwovenTable3", None}
DOC_TABLE_STYLES = {"NetwovenTable1", "NetwovenTable2", "NetwovenTable3"}
DASHES = "–—"
WEAVE_SPEC_NS = "http://netwoven.com/weave/spec/2.0"  # root namespace of the embedded spec part
# Exact set of docProps/custom.xml properties the v2 spec embedder writes
# (spec_embed.py's PROP_RECIPE / PROP_SHA / PROP_VERSION, hand-copied here as
# WEAVE_SPEC_NS already is above).  MSIP_Label_* legitimately needs prefix
# matching -- the suffix is a variable GUID -- but Weave_* does not, so a
# stray or truncated Weave_* property must fail this check, not silently pass.
SANCTIONED_WEAVE_PROPS = {"Weave_Recipe", "Weave_SpecSha256", "Weave_Version"}


def add(results, name, passed, evidence, hard=True):
    results.append({"check": name, "passed": bool(passed), "evidence": evidence, "hard": hard})


def own_text(p):
    """Text of a paragraph excluding paragraphs nested in its text boxes."""
    out = []
    for t in p.iter(W + "t"):
        anc = t.getparent()
        while anc is not None and anc.tag != W + "p":
            anc = anc.getparent()
        if anc is p:
            out.append(t.text or "")
    return "".join(out)


def detect_type(path):
    if path.suffix.lower() == ".pptx":
        return "deck"
    with zipfile.ZipFile(path) as z:
        styles = z.read("word/styles.xml").decode("utf-8", "replace") if "word/styles.xml" in z.namelist() else ""
    if 'w:styleId="PSOGrid"' in styles or 'w:val="PSO Grid"' in styles:
        return "sow"
    return "doc"


# --------------------------------------------------------------------------
# docx
# --------------------------------------------------------------------------
class Pkg:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = self.z.namelist()
        self.story = {}
        for n in self.names:
            if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n):
                self.story[n] = etree.fromstring(self.z.read(n))
        self.doc = self.story.get("word/document.xml")
        self.body = self.doc.find(W + "body") if self.doc is not None else None

    def read(self, name):
        return self.z.read(name).decode("utf-8", "replace") if name in self.names else ""

    def xml(self, name):
        return etree.fromstring(self.z.read(name)) if name in self.names else None

    def paragraphs(self, parts=None):
        """Every w:p (including nested in sdt/txbx/table) of the story parts."""
        for n, root in self.story.items():
            if parts and not any(n.startswith(p) for p in parts):
                continue
            for p in root.iter(W + "p"):
                yield n, p

    def texts(self):
        return [(n, norm(para_text(p))) for n, p in self.paragraphs()]

    def body_children(self):
        """Body-level elements, looking through body-level sdt wrappers."""
        for el in self.body:
            if el.tag == W + "sdt":
                content = el.find(W + "sdtContent")
                if content is not None:
                    for c in content:
                        yield c
            else:
                yield el

    def body_paragraphs(self):
        """Top-level body paragraphs + paragraphs inside body tables (text
        boxes excluded), for prose checks."""
        out = []
        for el in self.body_children():
            if el.tag == W + "p":
                out.append(el)
            elif el.tag == W + "tbl":
                out.extend(p for p in el.iter(W + "p") if not _in_textbox(p))
        return out

    def body_tables(self):
        return [el for el in self.body_children() if el.tag == W + "tbl"]


def _in_textbox(p):
    anc = p.getparent()
    while anc is not None:
        if anc.tag == W + "txbxContent":
            return True
        anc = anc.getparent()
    return False


def check_docx_common(pkg, results, kind, internal, client):
    texts = pkg.texts()
    joined = "\n".join(t for _, t in texts)
    hits = sorted(set(n for n, t in texts if "<<" in t))
    add(results, "no <<placeholder>> text left", not hits,
        "in %s" % ", ".join(hits) if hits else "none")
    hl = [n for n, root in pkg.story.items() if root.find(".//" + W + "highlight") is not None]
    add(results, "no yellow highlight (author instructions)", not hl, "in %s" % ", ".join(hl) if hl else "none")
    found = []
    for label, rx, flags in FORBIDDEN:
        if re.search(rx, joined, flags):
            found.append(label)
    add(results, "no forbidden template/legacy strings", not found, "found: %s" % (", ".join(found) or "none"))

    core = pkg.read("docProps/core.xml")
    app = pkg.read("docProps/app.xml")
    title = norm(html_unescape(re.search(r"<dc:title>(.*?)</dc:title>", core, re.S).group(1))) if re.search(r"<dc:title>(.*?)</dc:title>", core, re.S) else ""
    add(results, "dc:title set (not template default)", title.lower() not in DEFAULT_TITLES, "title=%r" % title)
    m = re.search(r"<Company>(.*?)</Company>", app, re.S)
    company = norm(html_unescape(m.group(1))) if m else ""
    if client:
        ok = company == norm(client)
        ev = "Company=%r, expected %r" % (company, client)
    elif company.lower() == "netwoven":
        ok = bool(internal and kind == "doc")
        ev = "Company='Netwoven' (%s)" % ("allowed with --internal on a document" if ok else "only allowed with --internal on a document")
    else:
        ok = company.lower() not in DEFAULT_COMPANIES
        ev = "Company=%r" % company
    add(results, "app.xml Company is the client", ok, ev)
    m = re.search(r"<TitlesOfParts>.*?<vt:lpstr>(.*?)</vt:lpstr>", app, re.S)
    tparts = norm(html_unescape(m.group(1))) if m else ""
    add(results, "app.xml TitlesOfParts matches dc:title", (not tparts) or tparts == title,
        "TitlesOfParts=%r" % tparts, hard=False)

    # bound content controls cache the same values
    mism = []
    n_company = 0
    for n, root in pkg.story.items():
        for sdt in root.iter(W + "sdt"):
            db = sdt.find(W + "sdtPr/" + W + "dataBinding")
            if db is None:
                continue
            xp = db.get(W + "xpath") or ""
            cached = norm(para_text(sdt.find(W + "sdtContent"))) if sdt.find(W + "sdtContent") is not None else ""
            if xp.endswith("Company[1]"):
                n_company += 1
                if cached != company:
                    mism.append("%s Company=%r" % (n, cached))
            elif xp.endswith("title[1]"):
                if cached != title:
                    mism.append("%s Title=%r" % (n, cached))
            elif xp.endswith("PublishDate[1]"):
                if not re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", cached):
                    mism.append("%s PublishDate=%r" % (n, cached))
            if sdt.find(W + "sdtPr/" + W + "showingPlcHdr") is not None:
                mism.append("%s showingPlcHdr on %s" % (n, xp.rsplit(":", 1)[-1]))
    add(results, "bound content controls cache title/company/date", not mism,
        "%d Company SDTs; mismatches: %s" % (n_company, "; ".join(mism[:4]) or "none"))

    iso_ok = False
    pd = ""
    for n in pkg.names:
        if re.match(r"customXml/item\d+\.xml$", n):
            x = pkg.read(n)
            m = re.search(r"<(?:\w+:)?PublishDate>(.*?)</(?:\w+:)?PublishDate>", x, re.S)
            if m:
                pd = m.group(1).strip()
                iso_ok = re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z?", pd) is not None
    add(results, "coverPageProps PublishDate is an ISO dateTime", iso_ok, "PublishDate=%r" % pd)

    creator = re.search(r"<dc:creator>(.*?)</dc:creator>", core, re.S)
    lmb = re.search(r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>", core, re.S)
    creator = norm(creator.group(1)) if creator else ""
    lmb = norm(lmb.group(1)) if lmb else ""
    add(results, "creator/lastModifiedBy scrubbed to Netwoven",
        creator in ("Netwoven", "") and lmb in ("Netwoven", ""), "creator=%r lastModifiedBy=%r" % (creator, lmb))
    add(results, "no lastPrinted timestamp", "lastPrinted" not in core, "present" if "lastPrinted" in core else "absent", hard=False)

    if "docProps/custom.xml" in pkg.names:
        cx = pkg.xml("docProps/custom.xml")
        names = [p.get("name") for p in cx.iter("{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property")]
        # Weave_* (the v2 spec embedder's Recipe / SpecSha256 / Version) sit
        # beside the Purview MSIP_Label_* properties by design.
        bad = [n for n in names
               if not (n or "").startswith("MSIP_Label_") and (n or "") not in SANCTIONED_WEAVE_PROPS]
        add(results, "custom.xml holds only MSIP_Label_* / Weave_* properties", not bad, "other properties: %s" % (", ".join(bad) or "none"))
    add(results, "sensitivity label part present (LabelInfo.xml)", "docMetadata/LabelInfo.xml" in pkg.names,
        "present" if "docMetadata/LabelInfo.xml" in pkg.names else "absent (template shipped one)", hard=False)
    others = []
    for n in pkg.names:
        if re.match(r"customXml/item\d+\.xml$", n):
            x = pkg.read(n)
            # the embedded weave spec part (SPEC.md "Embedding") is the one
            # other legitimate custom XML part
            if "coverPageProps" not in x and WEAVE_SPEC_NS not in x:
                others.append(n)
    add(results, "no customXml parts beyond coverPageProps / weave spec", not others, ", ".join(others) or "none")
    settings = pkg.read("word/settings.xml")
    add(results, "no documentProtection", "documentProtection" not in settings, "present" if "documentProtection" in settings else "absent")
    dirty_toc = 'w:dirty="true"' in pkg.read("word/document.xml") and "TOC" in pkg.read("word/document.xml")
    add(results, "updateFields on or dirty TOC", "updateFields" in settings or dirty_toc,
        "updateFields=%s dirtyTOC=%s" % ("updateFields" in settings, dirty_toc))
    for bad in ("word/embeddings/", "word/diagrams/"):
        if any(n.startswith(bad) for n in pkg.names):
            add(results, "no embedded OLE / SmartArt parts", False, bad + " present")
            break
    else:
        add(results, "no embedded OLE / SmartArt parts", True, "none")

    # warnings: house style (SOW template boilerplate that must stay verbatim
    # is skipped: the Glossary definitions carry en dashes, the estimate line
    # a colon)
    body_ps = pkg.body_paragraphs()
    skip = set()
    if kind == "sow":
        in_gloss = False
        for p in body_ps:
            st = p_style(p)
            if st == "Heading1":
                in_gloss = norm(para_text(p)).lower() == "glossary"
            if in_gloss or norm(own_text(p)).startswith("Total Project Estimate"):
                skip.add(id(p))
    prose = [norm(own_text(p)) for p in body_ps if norm(own_text(p)) and id(p) not in skip]
    dash_hits = [t[:70] for t in prose if any(d in t for d in DASHES)]
    add(results, "house style: no em/en dashes in body", not dash_hits,
        "%d paragraph(s), e.g. %s" % (len(dash_hits), dash_hits[:2]), hard=False)
    midcolon = [t[:70] for t in prose if re.search(r":\s+\S", t) and not t.rstrip().endswith(":")
                and not re.match(r"^(Note|Figure|Table)\b", t)]
    add(results, "house style: no mid-sentence colons", not midcolon,
        "%d paragraph(s), e.g. %s" % (len(midcolon), midcolon[:2]), hard=False)
    longh = [norm(para_text(p))[:60] for p in body_ps
             if (p_style(p) or "").startswith(("Heading", "NWHeading")) and len(norm(para_text(p)).split()) > 12]
    add(results, "headings <= 12 words", not longh, "%s" % (longh[:3] or "ok"), hard=False)
    return body_ps


def html_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def table_styles(pkg):
    out = []
    for tbl in pkg.body_tables():
        st = tbl.find(W + "tblPr/" + W + "tblStyle")
        out.append(st.get(W + "val") if st is not None else None)
    return out


def check_doc(pkg, results, internal, body_ps):
    styles = {}
    for p in body_ps:
        s = p_style(p) or "Normal"
        styles[s] = styles.get(s, 0) + 1
    n_h1 = styles.get("NWHeading1", 0)
    add(results, "uses NW Heading 1", n_h1 > 0,
        "heading style counts: %s" % {k: v for k, v in styles.items() if "Heading" in k} if n_h1
        else "no NW Heading 1 paragraph: this is still an unfilled shell, not a deliverable")
    add(results, "multi-level structure (NW Heading 2 present)", styles.get("NWHeading2", 0) > 0,
        "at least one subsection expected in a full deliverable", hard=False)
    direct = [k for k in styles if re.fullmatch(r"Heading[1-9]", k)]
    add(results, "no built-in Heading 1-9 (use NW Heading n)", not direct, "found: %s" % (direct or "none"))
    conf = any("Statement of Confidentiality" in norm(para_text(p)) for p in body_ps)
    if internal:
        add(results, "internal: client confidentiality page removed", not conf, "present" if conf else "absent as expected")
    else:
        add(results, "client: confidentiality page present", conf, "present" if conf else "missing")
    unchanged = [t for t in ("Author Name", "Job Position") if any(norm(para_text(p)) == t for p in body_ps)]
    add(results, "revision page author/role filled", not unchanged,
        "unchanged: %s" % (", ".join(unchanged) or "none"), hard=False)
    ts = table_styles(pkg)
    bad = [s for s in ts if s not in DOC_TABLE_STYLES]
    add(results, "tables use Netwoven table styles", not bad, "tables: %s" % (ts or "none"), hard=False)
    toc_head = H.toc_heading_text(pkg.body)
    add(results, "TOC heading is 'Table of Contents'", toc_head in (None, "Table of Contents"),
        "heading=%r (absent TOC is allowed for a minimal shell)" % toc_head, hard=False)


def check_sow(pkg, results, body_ps):
    h1 = [norm(para_text(p)) for p in body_ps if p_style(p) == "Heading1"]
    add(results, "SOW: at least 7 Heading 1 sections", len(h1) >= 7, "Heading1: %s" % h1)
    missing = [h for h in SOW_REQUIRED_H1 if not any(x.lower() == h.lower() for x in h1)]
    add(results, "SOW: required Heading 1 set present", not missing, "missing: %s" % (missing or "none"))
    bad = sorted(set(p_style(p) for p in body_ps if (p_style(p) or "").startswith("NWHeading") or p_style(p) == "Heading4"))
    add(results, "SOW: no NW Heading / Heading 4 usage", not bad, "found: %s" % (bad or "none"))
    add(results, "SOW: Appendix removed", not any(x.lower() == "appendix" for x in h1), "Heading1 'Appendix' %s" % ("present" if any(x.lower() == "appendix" for x in h1) else "absent"))
    obj = any(root.find(".//" + W + "object") is not None for root in pkg.story.values())
    add(results, "SOW: no embedded OLE object (Excel Gantt)", not obj, "w:object present" if obj else "none")
    instr = " ".join((it.text or "") for it in pkg.doc.iter(W + "instrText"))
    bookmarks = set(b.get(W + "name") for b in pkg.doc.iter(W + "bookmarkStart"))
    has_sum = "SUM(ABOVE)" in instr
    is_tm = "PRODUCT(LEFT)" in instr or "Total_Estimate" in bookmarks
    need = {"Onshore_Total", "Offshore_Total", "Total_Estimate"} if is_tm else {"TE"}
    add(results, "SOW: fee formula fields present", has_sum and (not is_tm or "PRODUCT(LEFT)" in instr),
        "SUM(ABOVE)=%s PRODUCT(LEFT)=%s" % (has_sum, "PRODUCT(LEFT)" in instr))
    add(results, "SOW: fee bookmarks present", need <= bookmarks, "missing: %s" % (sorted(need - bookmarks) or "none"))
    refs = set(re.findall(r"REF\s+(\S+)", instr))
    dangling = sorted(r for r in refs if r not in bookmarks)
    add(results, "SOW: REF field targets exist", not dangling, "dangling: %s" % (dangling or "none"))
    toc_head = H.toc_heading_text(pkg.body)
    add(results, "SOW: TOC heading is 'Contents'", toc_head == "Contents", "heading=%r" % toc_head)
    footers = "\n".join(norm(para_text(p)) for n, p in pkg.paragraphs(parts=("word/footer",)))
    add(results, "SOW: 'Netwoven Confidential' footer", "Netwoven Confidential" in footers, footers[:80] or "no footer text")
    nsect = len(list(pkg.doc.iter(W + "sectPr")))
    add(results, "SOW: exactly one section", nsect == 1, "%d sectPr" % nsect)
    # Scoped to the products table itself (via the canonical helper), not a
    # whole-document scan for any w14:checked -- an unrelated checkbox
    # elsewhere in the document must never make this look like a pass.
    products_tbl = H.products_table(pkg.body)
    ticked = [r for r in H.product_rows(products_tbl) if r[2]] if products_tbl is not None else []
    add(results, "SOW: at least one product ticked", bool(ticked), "%d ticked" % len(ticked), hard=False)
    amounts = re.findall(r"<w:t[^>]*>(-?\d[\d,]*\.\d\d)</w:t>", pkg.read("word/document.xml"))
    nonzero = [a for a in amounts if float(a.replace(",", "")) != 0.0]
    add(results, "SOW: fee table filled (non-zero amounts)", bool(nonzero), "%d amounts, %d non-zero" % (len(amounts), len(nonzero)), hard=False)
    grid_rows = 0
    for tbl in pkg.body.iter(W + "tbl"):
        st = tbl.find(W + "tblPr/" + W + "tblStyle")
        if st is not None and st.get(W + "val") == "PSOGrid":
            rows = tbl.findall(W + "tr")
            grid_rows = sum(1 for tr in rows[1:] if norm(para_text(tr)))
    add(results, "SOW: deliverables table has rows", grid_rows > 0, "%d filled PSO Grid rows" % grid_rows, hard=False)
    ts = table_styles(pkg)
    bad = [s for s in ts if s not in SOW_TABLE_STYLES]
    add(results, "SOW: table styles from the template set", not bad, "tables: %s" % ts, hard=False)
    exec_ok = False
    children = list(pkg.body)
    for i, el in enumerate(children):
        if el.tag == W + "p" and p_style(el) == "Heading1" and norm(para_text(el)).lower() == "executive summary":
            for nxt in children[i + 1:]:
                if nxt.tag == W + "p" and p_style(nxt) == "Heading1":
                    break
                if nxt.tag == W + "p" and len(norm(para_text(nxt))) > 40 and "Statement of Work" not in para_text(nxt):
                    exec_ok = True
                    break
    add(results, "SOW: Executive Summary written", exec_ok, "first paragraph under Executive Summary is %s" % ("filled" if exec_ok else "empty (slot)"), hard=False)


def check_docx(path, kind, internal, client):
    results = []
    pkg = Pkg(path)
    body_ps = check_docx_common(pkg, results, kind, internal, client)
    if kind == "sow":
        check_sow(pkg, results, body_ps)
    else:
        check_doc(pkg, results, internal, body_ps)
    return results


# --------------------------------------------------------------------------
# rendering (optional)
# --------------------------------------------------------------------------
def find_soffice():
    for cand in ("soffice", "libreoffice"):
        p = shutil.which(cand)
        if p:
            return p
    for cand in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
                 "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
                 "/usr/lib/libreoffice/program/soffice", "/usr/bin/soffice"):
        if os.path.exists(cand):
            return cand
    return None


def render_pages(path, png_dir=None, results=None):
    """Convert to PDF with LibreOffice; return page count (None when not
    available).  Adds warning-level results instead of failing."""
    soffice = find_soffice()
    if soffice is None:
        if results is not None:
            add(results, "render: LibreOffice available", False,
                "soffice not found (brew install --cask libreoffice); render skipped", hard=False)
        return None
    with tempfile.TemporaryDirectory() as td:
        profile = os.path.join(td, "profile")
        cmd = [soffice, "--headless", "-env:UserInstallation=file://%s" % profile,
               "--convert-to", "pdf", "--outdir", td, str(path)]
        try:
            subprocess.run(cmd, capture_output=True, timeout=240)
        except (subprocess.TimeoutExpired, OSError) as exc:
            if results is not None:
                add(results, "render: conversion", False, "soffice failed: %s" % exc, hard=False)
            return None
        pdfs = list(Path(td).glob("*.pdf"))
        if not pdfs:
            if results is not None:
                add(results, "render: conversion", False, "soffice produced no PDF", hard=False)
            return None
        pages = None
        if shutil.which("pdfinfo"):
            out = subprocess.run(["pdfinfo", str(pdfs[0])], capture_output=True, text=True, encoding="utf-8", errors="replace")
            m = re.search(r"Pages:\s+(\d+)", out.stdout)
            pages = int(m.group(1)) if m else None
        else:
            if results is not None:
                add(results, "render: pdfinfo available", False, "pdfinfo not found; page count unknown", hard=False)
        if png_dir:
            os.makedirs(png_dir, exist_ok=True)
            if shutil.which("pdftoppm"):
                subprocess.run(["pdftoppm", "-r", "80", "-png", str(pdfs[0]),
                                os.path.join(png_dir, Path(path).stem)], capture_output=True)
                shutil.copy(str(pdfs[0]), os.path.join(png_dir, Path(path).stem + ".pdf"))
            elif results is not None:
                add(results, "render: pdftoppm available", False, "pdftoppm not found; no PNGs", hard=False)
        if results is not None:
            add(results, "render: converted with LibreOffice", True, "%s pages" % pages, hard=False)
        return pages


# --------------------------------------------------------------------------
def _fail(args, error, hint=""):
    """Exit before a results list exists (bad path / bad --type combo).

    Prints the {"ok": false, "error", "hint"} shape every other script in
    this pipeline uses on failure when --json was requested, so weave.py's
    run_child()/_last_json_object() can parse this script's own specific
    error straight into its `error` field.  Without --json, keeps printing
    the plain "error: ..." message to stderr this always printed."""
    if getattr(args, "json", False):
        print(json.dumps({"ok": False, "error": error, "hint": hint}))
    else:
        print("error: %s" % error, file=sys.stderr)
    sys.exit(1)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--type", choices=["doc", "sow", "deck"], default=None)
    ap.add_argument("--internal", action="store_true")
    ap.add_argument("--client", default=None, help="expected client name (Company)")
    ap.add_argument("--render", action="store_true")
    ap.add_argument("--png-dir", default=None)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    path = Path(args.file)
    if not path.exists():
        _fail(args, "%s not found" % path, "Check the path.")
    kind = args.type or detect_type(path)
    results = []
    if kind == "deck":
        if path.suffix.lower() != ".pptx":
            _fail(args, "--type deck needs a .pptx file", "Pass a .pptx file, or drop --type to auto-detect.")
        try:
            from nw_pptx_helpers import validate_deck  # lazy: absent in Word-only skills
        except ImportError as exc:
            add(results, "deck validator available", False,
                "nw_pptx_helpers.validate_deck not importable beside validate_deliverable.py (%s)" % exc)
        else:
            import pptx
            prs = pptx.Presentation(str(path))
            results.extend(validate_deck(prs, str(path), args))
    else:
        if path.suffix.lower() != ".docx":
            _fail(args, "unsupported file type %s" % path.suffix, "Supported: .docx (doc/sow) or .pptx (deck).")
        results = check_docx(path, kind, args.internal, args.client)

    pages = None
    if args.render or args.png_dir:
        pages = render_pages(path, args.png_dir, results)
    hard_fails = [r for r in results if not r["passed"] and r["hard"]]
    report = {"file": str(path), "type": kind, "pages": pages, "passed": not hard_fails, "results": results}
    if args.json:
        print(json.dumps(report, indent=2))
    else:
        print("%s (%s)" % (path.name, kind))
        for r in results:
            mark = "PASS" if r["passed"] else ("FAIL" if r["hard"] else "warn")
            print("[%s] %s -- %s" % (mark, r["check"], r["evidence"]))
        if pages is not None:
            print("rendered pages: %d" % pages)
        print("RESULT:", "PASS" if not hard_fails else "FAIL (%d hard failure%s)" % (len(hard_fails), "s" if len(hard_fails) != 1 else ""))
    sys.exit(0 if not hard_fails else 1)


if __name__ == "__main__":
    main()
