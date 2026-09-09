#!/usr/bin/env python3
"""
Create a client-ready Netwoven Statement of Work shell from a 2026 SOW base.

  new_sow_docx.py --type milestone|tm BASE OUT --title "Project Name"
      --client "Client Name" --date 2026-09-01
      [--client-logo logo.png] [--keep-products]
      [--keep-appendix-timeline-table] [--json-spec spec.json]
      [--expenses-cap 10]

Bases: NW_SOW_Milestone_Base_2026.docx (--type milestone) and
NW_SOW_TM_Base_2026.docx (--type tm).  The script asserts the base matches
the type.  SOWs have no internal variant: --internal is rejected.

What is stripped, unconditionally and located by content (never by index):
  * the whole Appendix (Heading1 "Appendix" to the final sectPr) including
    the SmartArt diagram parts; for the Milestone base its
    Phase/Description/Duration timeline table is first cloned into the
    Timeline section (the T&M base already has one there);
  * the embedded Excel OLE Gantt (w:object, _MON_ bookmarks, EMF preview,
    embedded workbook) and its instruction line;
  * every yellow-highlighted author instruction (whole paragraph when fully
    highlighted; partial paragraphs follow rules: the Executive Summary
    fill-in sentence becomes an empty Body Text slot, "Post Go Live support
    for X weeks" and the Milestone plain-text total line are deleted, the
    travel "X%" cap becomes "[__]%" or --expenses-cap);
  * the unhighlighted instruction lines: the rate-card SharePoint link
    paragraph (+ its relationship), "Please Click the Executive Summary
    link", every "<<Insert ...>>" prompt;
  * sample scope / deliverable content (Milestone outline, T&M Scope 1/2,
    PSO Grid sample row -> header + one empty row, Biz Apps rows unless the
    spec lists biz_apps (then the block goes), products unticked unless the
    spec ticks them; ticked -> unticked rows dropped unless --keep-products).

Then: TOC reset to the template's "Contents" heading + dirty TOC field, every
bound Company / Title / Publish Date content control (46/35 Company SDTs, the
cover text boxes, header, footer) gets its cached text, the client logo is
placed in the cover box or the box is removed, numbering pruned, metadata
scrubbed, updateFields on, app.xml TitlesOfParts set.

--json-spec (all keys optional):
{
  "milestones": [{"name": "...", "description": "...", "amount": 25000,
                  "completion": "Week 4"}],                  # milestone only
  "rates": [{"role": "Senior Consultant", "location": "onshore|offshore",
             "resources": 1, "weekly_hours": 40, "weeks": 12, "rate": 185}],  # tm only
  "products": ["Microsoft Teams", "SharePoint Online"],
  "biz_apps": ["D365 Customer Engagement"],
  # deliverables rows must match the variant's real column count or values
  # silently shift left and the last column is dropped -- add_deliverable_row
  # writes values[i] into cells[i] positionally with no shape check.
  "deliverables": [["Scope Section", "Milestone", "Deliverable", "Description", "Format"]],  # milestone: 5 cols
  # "deliverables": [["Scope Section", "Deliverable", "Description", "Format"]],              # tm: 4 cols (no Milestone column)
  "timeline": [["Phase", ["line 1", "line 2"], "Week 1 - 4"]],
  "expenses_cap": 10,
  "assumptions_remove": ["Govern 365 Promotional Offer",
                         "Establish M365 Claiming Partner of Record (CPOR)"]
}
Fee/rate tables keep their =SUM(ABOVE) / =PRODUCT(LEFT) fields and bookmarks
(TE, Onshore_Total, Offshore_Total, Total_Estimate); cached totals and REF
caches are written so the numbers show before Word recomputes.

Post-save assertions: no word/embeddings/, word/diagrams/, image8.emf in the
zip; no yellow highlight; no "<<"; no "Please Insert Client Logo"; no
Appendix heading; fee fields + bookmarks intact; every Company SDT caches
the client name.  Requires python-docx, lxml, nw_docx_helpers.py.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from copy import deepcopy

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx  # noqa: E402
from lxml import etree  # noqa: E402

import nw_docx_helpers as H  # noqa: E402
from nw_docx_helpers import W  # noqa: E402

RATE_CARD_URL_FRAGMENT = "sharepoint.com"


def highlight_rules(expenses_cap):
    cap_text = ("%s%%" % expenses_cap) if expenses_cap not in (None, "") else "[__]%"
    return [
        {"contains": "would like to migrate/develop", "action": "empty_slot", "style": "BodyText"},
        {"contains": "Post Go Live support for", "action": "delete"},
        {"contains": "Total Project Estimate (right click", "action": "delete"},
        {"contains": "Total expenses for the project will be limited", "action": "replace_highlighted",
         "text": cap_text},
    ]


def detect_base_type(body):
    if H.find_table(body, header_startswith=["Milestone", "Milestone Description"]) is not None:
        return "milestone"
    if H.find_table(body, header_startswith=["Location"]) is not None:
        return "tm"
    return None


def heading_index(children, level, text, start=0):
    return H.find_heading(children, "Heading%d" % level, text, start)


def is_heading(el, levels=(1, 2)):
    return H.p_style(el) in tuple("Heading%d" % lv for lv in levels)


def range_after_heading(children, h_idx, stop_levels=(1, 2)):
    """(start, end) of the content after heading index h_idx up to the next
    heading of the given levels or the final sectPr."""
    end = H.section_end(children, h_idx, lambda el: is_heading(el, stop_levels))
    return h_idx + 1, end


def clone_ms_timeline_table(doc, log):
    """Milestone base: copy the Appendix Phase/Description/Duration table into
    the Timeline section (after the last text paragraph before the OLE)."""
    body = doc.element.body
    children = list(body)
    app = heading_index(children, 1, "Appendix")
    if app is None:
        return None
    tbl = None
    for el in children[app:]:
        if H.is_tbl(el) and H.tbl_style(el) == "NetwovenTable1":
            hdr = [H.norm(H.para_text(tc)).lower() for tc in H.row_cells(H.table_rows(el)[0])]
            if hdr[:1] == ["phase"]:
                tbl = el
                break
    if tbl is None:
        log.append("warning: no Phase/Description/Duration table found in the Appendix")
        return None
    clone = deepcopy(tbl)
    tl = heading_index(children, 2, "Timeline")
    if tl is None:
        log.append("warning: no Timeline heading; timeline table not cloned")
        return None
    ole = None
    for el in children[tl + 1:]:
        if is_heading(el):
            break
        if H.is_p(el) and el.find(".//" + W + "object") is not None:
            ole = el
            break
    anchor = None
    if ole is not None:
        prev = ole.getprevious()
        while prev is not None and (H.is_empty_p(prev) or prev.tag in (W + "bookmarkStart", W + "bookmarkEnd")):
            prev = prev.getprevious()
        anchor = prev
    if anchor is None:
        # fall back: after the last non-empty paragraph directly under the heading
        anchor = children[tl]
        for el in children[tl + 1:]:
            if is_heading(el) or H.is_tbl(el):
                break
            if H.is_p(el) and H.para_text(el).strip():
                anchor = el
    anchor.addnext(clone)
    log.append("cloned Appendix timeline table into the Timeline section")
    return clone


def remove_appendix(doc, log):
    body = doc.element.body
    children = list(body)
    app = heading_index(children, 1, "Appendix")
    if app is None:
        log.append("warning: no Appendix heading found")
        return 0
    end = len(children)
    while children[end - 1].tag == W + "sectPr":
        end -= 1
    n = H.remove_range(body, children, app, end)
    dropped = H.drop_unreferenced_rels(doc.part)
    log.append("removed Appendix (%d body elements, dropped rels %s)" % (n, ", ".join(dropped) or "none"))
    return n


def remove_instruction_lines(doc, log):
    body = doc.element.body
    part = doc.part
    n = 0
    for rid in H.external_hyperlink_rids(part, RATE_CARD_URL_FRAGMENT):
        n += H.remove_hyperlink_paragraph(body, rid)
        H.strip_external_hyperlink_rel(part, rid)
        log.append("removed rate-card link paragraph and relationship %s" % rid)
    n += H.remove_paragraphs_where(
        body, lambda p: H.norm(H.para_text(p)).lower().startswith("please click the executive summary link"))
    inserts = [p for p in body if H.is_p(p) and H.norm(H.para_text(p)).lower().startswith("<<insert")]
    parents = []
    for p in inserts:
        prev = p.getprevious()
        while prev is not None and prev.tag in (W + "bookmarkStart", W + "bookmarkEnd"):
            prev = prev.getprevious()
        # the item above a deleted prompt may be a label that only introduced
        # the prompt ("Use Cases & Prototypes" / "Miscellaneous"); the
        # template's levels are inconsistent, so any numbered neighbour is a
        # candidate and remove_orphan_list_labels applies the label test
        if prev is not None and H.is_p(prev) and H.p_numpr(prev) is not None and H.p_numpr(p) is not None:
            parents.append(prev)
    k = H.remove_elements(body, inserts)
    n += k
    log.append("removed %d '<<Insert ...>>' prompt paragraphs" % k)
    o = H.remove_orphan_list_labels(body, parents)
    if o:
        log.append("removed %d list labels left without items" % o)
    # any remaining paragraph that is only a <<placeholder>> outside an SDT
    k = H.remove_paragraphs_where(
        body, lambda p: re.fullmatch(r"<<[^<>]*>>", H.norm(H.para_text(p))) is not None
        and p.find(".//" + W + "sdt") is None)
    if k:
        log.append("removed %d bare <<placeholder>> paragraphs" % k)
    return n


def capture_prototypes(children, start, end):
    """First BodyText/ListParagraph numbered paragraph per ilvl in a range."""
    protos = {}
    for el in children[start:end]:
        np_ = H.p_numpr(el)
        if np_ is None:
            continue
        lvl = int(np_[0])
        if lvl not in protos and H.para_text(el).strip():
            protos[lvl] = deepcopy(el)
    return protos


def empty_slot(style="BodyText"):
    p = H.E("p")
    ppr = H.E("pPr")
    ppr.append(H.E("pStyle", {"val": style}))
    p.append(ppr)
    return p


def strip_sample_scope(doc, kind, spec, log):
    """Scope of Work sample outline -> generated milestone outline (spec) or
    one empty Body Text slot.  Deliverable lists after the PSO Grid -> gone."""
    body = doc.element.body
    children = list(body)
    scope_h = heading_index(children, 2, "Scope of Work")
    protos = {}
    if scope_h is not None:
        start, end = range_after_heading(children, scope_h)
        protos = capture_prototypes(children, start, end)
        n = H.remove_range(body, children, start, end)
        log.append("removed %d sample Scope of Work paragraphs" % n)
        children = list(body)
        anchor = children[scope_h]
        milestones = (spec or {}).get("milestones") or []
        if kind == "milestone" and milestones and 0 in protos:
            new = []
            for i, m in enumerate(milestones, start=1):
                p0 = deepcopy(protos[0])
                H.set_paragraph_text(p0, "Milestone %d" % i)
                new.append(p0)
                name = str(m.get("name") or "").strip()
                if name and 1 in protos:
                    p1 = deepcopy(protos[1])
                    H.set_paragraph_text(p1, name)
                    new.append(p1)
                desc = str(m.get("description") or "").strip()
                if desc and (2 in protos or 1 in protos):
                    p2 = deepcopy(protos.get(2, protos.get(1)))
                    H.set_paragraph_text(p2, desc)
                    new.append(p2)
            for el in new:
                anchor.addnext(el)
                anchor = el
            log.append("wrote %d milestones into Scope of Work" % len(milestones))
        else:
            anchor.addnext(empty_slot("BodyText"))
            log.append("left one empty Body Text slot under Scope of Work")
    else:
        log.append("warning: no 'Scope of Work' heading")

    # deliverable lists after the PSO Grid (Milestone base)
    children = list(body)
    grid = H.find_table(body, style="PSOGrid")
    if grid is not None:
        gi = children.index(grid)
        end = H.section_end(children, gi, lambda el: is_heading(el))
        doomed = [el for el in children[gi + 1:end] if not H.is_empty_p(el)]
        if doomed:
            H.remove_elements(body, doomed)
            log.append("removed %d sample deliverable list paragraphs after the PSO Grid" % len(doomed))
        H.collapse_empty_paragraphs(body, grid.getnext(), keep=1) if grid.getnext() is not None else None
    return protos


def fill_deliverables(doc, spec, log):
    body = doc.element.body
    grid = H.find_table(body, style="PSOGrid")
    if grid is None:
        log.append("warning: no PSO Grid table")
        return
    proto = H.clear_sample_rows(grid)
    rows = (spec or {}).get("deliverables") or []
    for values in rows:
        H.add_deliverable_row(grid, list(values), proto)
    log.append("PSO Grid: header + %s" % ("%d deliverable rows" % len(rows) if rows else "one empty row"))


def fill_timeline(doc, spec, keep_sample, log):
    body = doc.element.body
    tbl = H.find_table(body, style="NetwovenTable1", header_startswith=["Phase", "Description", "Duration"])
    if tbl is None:
        log.append("warning: no Phase/Description/Duration timeline table")
        return
    rows = (spec or {}).get("timeline") or []
    if rows:
        proto = H.clear_sample_rows(tbl)
        for row in rows:
            phase, lines, duration = (list(row) + ["", [], ""])[:3]
            H.add_timeline_row(tbl, str(phase), lines, str(duration), proto)
        log.append("timeline table: %d phase rows" % len(rows))
    elif keep_sample:
        log.append("timeline table: kept the template's illustrative rows (--keep-appendix-timeline-table)")
    else:
        H.clear_sample_rows(tbl)
        log.append("timeline table: header + one empty row")


def fill_products(doc, spec, keep_products, log):
    body = doc.element.body
    names = (spec or {}).get("products") or []
    # Resolve the products table once: tick_product() re-scans the whole
    # body for it on every call otherwise, an O(products x body size) cost
    # for what should be a handful of lookups.
    tbl = H.products_table(body)
    ticked, missing = [], []
    for name in names:
        if H.tick_product(tbl, str(name), True):
            ticked.append(name)
        else:
            missing.append(name)
    if missing:
        log.append("warning: products not found in the table: %s" % ", ".join(map(str, missing)))
    if ticked and not keep_products:
        n = H.drop_unticked_products(body)
        log.append("products: ticked %d, dropped %d unticked rows" % (len(ticked), n))
    elif ticked:
        log.append("products: ticked %d, kept unticked rows (--keep-products)" % len(ticked))
    else:
        log.append("products: nothing ticked (all rows kept unticked)")
    biz = (spec or {}).get("biz_apps") or []
    n = H.set_biz_apps(body, biz)
    log.append("Biz Apps: %s" % ("%d rows" % n if n else "block removed (none in spec)"))


def remove_assumptions(doc, titles, log):
    body = doc.element.body
    for title in titles or []:
        children = list(body)
        idx = heading_index(children, 2, title)
        if idx is None:
            log.append("warning: assumption section %r not found" % title)
            continue
        end = H.section_end(children, idx, lambda el: is_heading(el))
        n = H.remove_range(body, children, idx, end)
        log.append("removed assumption section %r (%d elements)" % (title, n))
    dropped = H.drop_unreferenced_rels(doc.part, reltypes=(H.RT_HYPERLINK,))
    if dropped:
        log.append("dropped unreferenced hyperlink rels %s" % ", ".join(dropped))


def fill_fees(doc, kind, spec, log):
    body = doc.element.body
    if kind == "milestone":
        tbl = H.find_table(body, header_startswith=["Milestone", "Milestone Description"])
        if tbl is None:
            log.append("warning: no milestone fee table")
            return
        ms = (spec or {}).get("milestones") or []
        if ms:
            proto = H.reset_fee_rows_ms(tbl)
            for i, m in enumerate(ms, start=1):
                name = str(m.get("name") or "").strip()
                label = "Milestone %d" % i
                H.add_fee_row_ms(tbl, label, name or str(m.get("description") or ""),
                                 m.get("amount", ""), str(m.get("completion") or ""), proto)
        total = H.recompute_ms_totals(tbl, body)
        log.append("fee table: %s, cached total %s" % ("%d milestone rows" % len(ms) if ms else "template rows kept", H.fmt_amount(total)))
    else:
        tbl = H.find_table(body, header_startswith=["Location"])
        if tbl is None:
            log.append("warning: no rate table")
            return
        rates = (spec or {}).get("rates") or []
        incomplete = []
        for r in rates:
            # No default: a field missing from the spec must reach
            # set_rate_row_tm as None (a gap to leave blank), never as a
            # fabricated 0 (a real, invented number in a signable document).
            H.set_rate_row_tm(tbl, str(r.get("role") or ""), str(r.get("location") or "onshore"),
                              r.get("resources"), r.get("weekly_hours"), r.get("weeks"), r.get("rate"))
            if any(r.get(k) in (None, "") for k in ("resources", "weekly_hours", "weeks", "rate")):
                incomplete.append(r.get("role"))
        if incomplete:
            log.append("warning: rate row(s) left blank, missing resources/weekly_hours/weeks/rate: %s"
                       % ", ".join(map(str, incomplete)))
        if rates:
            n = H.drop_unused_rate_rows_tm(tbl)
            log.append("rate table: %d roles set, %d unused role rows dropped" % (len(rates), n))
        total = H.recompute_tm_totals(tbl, body)
        log.append("rate table: cached total %s (REF caches refreshed)" % H.fmt_amount(total))


def post_save_assertions(path, kind, client):
    names = H.zip_names(path)
    problems = []
    for bad in ("word/embeddings/", "word/diagrams/"):
        if any(n.startswith(bad) for n in names):
            problems.append("%s still in package" % bad)
    if any(n.endswith("image8.emf") for n in names):
        problems.append("image8.emf still in package")
    if "docMetadata/LabelInfo.xml" not in names:
        problems.append("docMetadata/LabelInfo.xml missing (sensitivity label must stay)")
    company_total = company_ok = 0
    instr_all = ""
    bookmarks = set()
    for n in names:
        if not re.match(r"word/(document|header\d*|footer\d*)\.xml$", n):
            continue
        root = etree.fromstring(H.zip_read(path, n))
        if root.find(".//" + W + "highlight") is not None:
            problems.append("%s: yellow highlight survives" % n)
        for p in root.iter(W + "p"):
            t = H.para_text(p)
            if "<<" in t:
                problems.append("%s: '<<' survives in %r" % (n, H.norm(t)[:60]))
                break
        for p in root.iter(W + "p"):
            if "Please Insert Client Logo" in H.norm(H.para_text(p)):
                problems.append("%s: 'Please Insert Client Logo' survives" % n)
                break
        for p in root.iter(W + "p"):
            if H.p_style(p) == "Heading1" and H.norm(H.para_text(p)).lower() == "appendix":
                problems.append("Appendix heading survives")
        if root.find(".//" + W + "object") is not None:
            problems.append("%s: w:object survives" % n)
        for sdt in root.iter(W + "sdt"):
            if H._sdt_binding_kind(sdt) == "company":
                company_total += 1
                if H.norm(H.para_text(sdt.find(W + "sdtContent"))) == H.norm(client):
                    company_ok += 1
        for it in root.iter(W + "instrText"):
            instr_all += (it.text or "") + " "
        for b in root.iter(W + "bookmarkStart"):
            bookmarks.add(b.get(W + "name"))
    if company_ok != company_total or company_total == 0:
        problems.append("Company SDT caches: %d of %d equal %r" % (company_ok, company_total, client))
    if "SUM(ABOVE)" not in instr_all:
        problems.append("=SUM(ABOVE) fee field missing")
    need = {"TE"} if kind == "milestone" else {"Onshore_Total", "Offshore_Total", "Total_Estimate"}
    if kind == "tm" and "PRODUCT(LEFT)" not in instr_all:
        problems.append("=PRODUCT(LEFT) rate field missing")
    missing = need - bookmarks
    if missing:
        problems.append("fee bookmarks missing: %s" % ", ".join(sorted(missing)))
    return problems, {"company_sdts": company_total}


class SowError(Exception):
    """Raised by build_sow(); main() turns it into sys.exit with the message."""


def build_sow(kind, base, out, title, client, date, spec=None, client_logo=None, keep_products=False,
              keep_appendix_timeline_table=False, expenses_cap=None, before_save=None):
    """Library entry point behind the CLI.  `spec` is the v1 --json-spec dict
    (all keys optional).  `before_save(doc, log)` runs after every strip/fill
    step and before the save + zip passes, so a caller can write extra prose
    into the shell's slots with the template mechanics already settled.
    Returns {"log": [...], "counts": {...}, "problems": [...], "info": {...}}.
    Raises SowError for the CLI's argument/base errors."""
    if kind not in ("milestone", "tm"):
        raise SowError("error: --type must be milestone or tm, got %r" % kind)
    if H.parse_date(date) is None:
        raise SowError("error: could not parse --date %r (use 2026-09-01 or 9/1/2026)" % date)
    spec = dict(spec or {})
    if expenses_cap is None:
        expenses_cap = spec.get("expenses_cap")
    if client_logo and not os.path.exists(client_logo):
        raise SowError("error: --client-logo %r not found" % client_logo)

    doc = docx.Document(base)
    body = doc.element.body
    detected = detect_base_type(body)
    if detected != kind:
        raise SowError("error: base %s looks like a %s SOW, but --type %s was given" % (
            os.path.basename(base), detected or "non-SOW", kind))
    if H.toc_heading_text(body) is None:
        raise SowError("error: base has no Table of Contents content control; wrong base?")

    log = []
    # 1) Appendix (Milestone: clone its timeline table first), diagram rels
    if kind == "milestone":
        clone_ms_timeline_table(doc, log)
    remove_appendix(doc, log)
    # 2) highlighted author instructions (rules for partial paragraphs)
    actions = H.strip_highlighted_instructions(body, highlight_rules(expenses_cap))
    log.append("highlighted instructions: %s" % ", ".join(
        "%d %s" % (sum(1 for a in actions if a[0] == k), k) for k in sorted(set(a[0] for a in actions))) or "none")
    # 3) OLE Gantt
    dropped = H.remove_ole_timeline(doc)
    log.append("removed OLE timeline (rels %s)" % (", ".join(dropped) or "none"))
    # 4) unhighlighted instruction lines and <<Insert>> prompts
    remove_instruction_lines(doc, log)
    # 5) sample content
    strip_sample_scope(doc, kind, spec, log)
    fill_deliverables(doc, spec, log)
    fill_timeline(doc, spec, keep_appendix_timeline_table, log)
    fill_products(doc, spec, keep_products, log)
    remove_assumptions(doc, spec.get("assumptions_remove"), log)
    fill_fees(doc, kind, spec, log)
    # 6) TOC, blank-paragraph runs, logo, numbering, metadata
    H.reset_toc(body, "Contents")
    k = H.collapse_empty_runs(body, max_run=3, keep=1)
    if k:
        log.append("collapsed runs of 3+ blank paragraphs (%d removed)" % k)
    if client_logo:
        if H.set_client_logo(doc, client_logo):
            log.append("client logo placed in the cover box")
        else:
            log.append("warning: logo box not found; logo not placed")
    else:
        log.append("client logo box %s" % ("removed" if H.remove_client_logo_box(doc) else "not found"))
    if before_save is not None:
        before_save(doc, log)
    nums, abstracts = H.prune_numbering(doc)
    log.append("numbering pruned: %d num, %d abstractNum" % (nums, abstracts))
    doc.core_properties.title = title
    # tmp + move, not doc.save(out) directly -- see new_deliverable_docx.py's
    # finish_document() for why a signed-price SOW is exactly the file this
    # protects.
    tmp = out + ".tmp"
    doc.save(tmp)
    shutil.move(tmp, out)

    counts = H.set_bound_fields(out, title=title, company=client, date=date)
    H.patch_settings_updatefields(out)
    H.patch_footer_dateformat(out)
    H.scrub_metadata(out)
    log.append("bound SDTs rewritten: title=%d company=%d date=%d" % (counts["title"], counts["company"], counts["date"]))

    problems, info = post_save_assertions(out, kind, client)
    return {"log": log, "counts": counts, "problems": problems, "info": info}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--type", required=True, choices=["milestone", "tm"])
    ap.add_argument("base")
    ap.add_argument("out")
    ap.add_argument("--title", required=True, help="project name (cover, header, dc:title)")
    ap.add_argument("--client", required=True, help="client legal name (all Company content controls)")
    ap.add_argument("--date", required=True, help="effective/publish date (2026-09-01 or 9/1/2026)")
    ap.add_argument("--client-logo", default=None, help="PNG/JPEG placed in the cover logo box; omitted -> box removed")
    ap.add_argument("--keep-products", action="store_true", help="keep unticked product rows")
    ap.add_argument("--keep-appendix-timeline-table", action="store_true",
                    help="keep the template's illustrative timeline rows instead of clearing them")
    ap.add_argument("--json-spec", default=None)
    ap.add_argument("--expenses-cap", default=None, help="travel expense cap percent (Milestone base)")
    ap.add_argument("--internal", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args(argv)

    if args.internal:
        sys.exit("error: SOWs have no internal variant; --internal is not supported by new_sow_docx.py")
    spec = {}
    if args.json_spec:
        with open(args.json_spec, "r", encoding="utf-8") as fh:
            spec = json.load(fh)
    try:
        res = build_sow(args.type, args.base, args.out, args.title, args.client, args.date, spec=spec,
                        client_logo=args.client_logo, keep_products=args.keep_products,
                        keep_appendix_timeline_table=args.keep_appendix_timeline_table,
                        expenses_cap=args.expenses_cap)
    except SowError as exc:
        sys.exit(str(exc))

    print("OK: wrote %s (type=%s, title=%r, client=%r, date=%s)" % (
        args.out, args.type, args.title, args.client, H.date_display(args.date)))
    for line in res["log"]:
        print("  - " + line)
    print("  - Company SDTs in document/header/footer: %d" % res["info"]["company_sdts"])
    if res["problems"]:
        print("POST-SAVE ASSERTIONS FAILED:")
        for pr in res["problems"]:
            print("  ! " + pr)
        sys.exit(1)
    print("  post-save assertions: all passed")


if __name__ == "__main__":
    main()
