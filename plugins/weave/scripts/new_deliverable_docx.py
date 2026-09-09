#!/usr/bin/env python3
"""
Create a ready-to-fill Netwoven Word deliverable shell from the 2026 base
(NW_Document_Base_2026.docx).

Why this script exists: the base ships with demo content ("Heading 1" style
stubs + a long "How to use this template" section) that must never reach a
client, and its section/header/footer mechanics are easy to break by naive
stripping (the final sectPr carries NO header/footer references of its own;
it inherits them from the front-matter section).  This script removes exactly
the right content, preserves the section mechanics, rewrites the cached text
of every bound content control (Title / Company / Publish Date) and scrubs
metadata, so every invocation starts from a correct shell.

Usage:
  new_deliverable_docx.py BASE OUT --title "Document Title" --company "Client"
      [--date 2026-09-01] [--internal] [--keep-figure-lists] [--minimal]
      [--author "Jane Doe"] [--role "Senior Consultant"]

Modes:
  default     : keep cover, confidentiality page, revision history, TOC;
                drop List of Figures/Tables (stale after demo removal; keep
                with --keep-figure-lists if you will add captioned
                figures/tables); remove demo + instructional content.
  --internal  : additionally remove the client confidentiality page; company
                defaults to "Netwoven" if --company not given.
  --minimal   : bare shell: strip ALL body content, union header/footer
                references into the surviving final sectPr (required, or the
                output silently loses its header/footer).

Always: bound Title/Company/Publish Date SDT caches + dc:title + app.xml
Company/TitlesOfParts + coverPageProps PublishDate, settings updateFields,
footer date format, metadata scrub, numbering prune.  --date defaults to
today.  Requires python-docx, lxml (nw_docx_helpers.py beside this file).
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
from copy import deepcopy

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx  # noqa: E402

import nw_docx_helpers as H  # noqa: E402
from nw_docx_helpers import W  # noqa: E402


def fill_author(doc, author, role):
    """Fill the 'Author Name' / 'Job Position' paragraphs of the revision
    history page and the Author cell of every revision-table data row."""
    body = doc.element.body
    filled = []
    for p in body:
        if not H.is_p(p):
            continue
        t = H.norm(H.para_text(p))
        if author and t == "Author Name":
            H.set_paragraph_text(p, author)
            filled.append("Author Name")
        elif role and t == "Job Position":
            H.set_paragraph_text(p, role)
            filled.append("Job Position")
    if author:
        tbl = H.find_table(body, header_startswith=["Date", "Version", "Revision Description", "Author"])
        if tbl is not None:
            for tr in H.table_rows(tbl)[1:]:
                cells = H.row_cells(tr)
                if len(cells) >= 4:
                    H.set_cell_text(cells[3], author)
            filled.append("revision table Author cells")
    return filled


class ShellError(Exception):
    """Raised by the library entry points; main() turns it into sys.exit."""


def resolve_company_date(company, internal, date):
    """Shared argument normalisation: company defaults to Netwoven for the
    internal variant, date defaults to today.  Raises ShellError with the
    CLI's exact messages."""
    company = company or ("Netwoven" if internal else None)
    if company is None:
        raise ShellError("error: --company is required for client deliverables")
    date = date or H.date_display(H.parse_date("today"))
    if H.parse_date(date) is None:
        raise ShellError("error: could not parse --date %r (use 2026-09-01, 9/1/2026 or 'September 2026')" % date)
    return company, date


def prepare_document_shell(doc, internal=False, minimal=False, keep_figure_lists=False,
                           author=None, role=None):
    """Strip the opened base down to a shell IN MEMORY (no save).  Returns the
    list of 'removed' log lines.  Callers may append body content afterwards
    and must then call finish_document()."""
    body = doc.element.body
    final_sectPr = body.find(W + "sectPr")
    if final_sectPr is None:
        raise ShellError("error: base file has no final sectPr; wrong base?")
    if H.find_index(list(body), lambda el: H.is_p_with_style(el, "NWHeading1")) is None and not minimal:
        raise ShellError("error: base has no NWHeading1 paragraph; this is not the 2026 document base")

    removed = []
    if minimal:
        # Union header/footer refs from ALL earlier sections into the final
        # sectPr FIRST: it has none of its own and inherits from sections this
        # mode deletes.
        have = {(el.tag, el.get(W + "type")) for el in final_sectPr
                if el.tag in (W + "headerReference", W + "footerReference")}
        for p in body.findall(W + "p"):
            sect = p.find(W + "pPr/" + W + "sectPr")
            if sect is None:
                continue
            for ref in sect:
                if ref.tag in (W + "headerReference", W + "footerReference"):
                    key = (ref.tag, ref.get(W + "type"))
                    if key not in have:
                        final_sectPr.insert(0, deepcopy(ref))
                        have.add(key)
            # a titlePg on the front-matter section suppressed the cover
            # header; the shell has no cover so we do not carry it over.
        for child in list(body):
            if child is not final_sectPr:
                body.remove(child)
        removed.append("all body content (minimal shell)")
    else:
        children = list(body)
        # 1) Demo + instructional content: from the FIRST NWHeading1 paragraph
        #    ("Heading 1" stub) to the end of body (excl. final sectPr).
        demo_start = H.find_index(children, lambda el: H.is_p_with_style(el, "NWHeading1"))
        end = len(children)
        while children[end - 1] is final_sectPr:
            end -= 1
        n = H.remove_range(body, children, demo_start, end)
        removed.append("%d demo/instruction body elements" % n)

        # 2) List of Figures / List of Tables (their cached entries point at
        #    demo figures that no longer exist).  Drop by default; on
        #    request, keep the section but replace the cached entries with
        #    a dirty field (same trap as the TOC: a non-Word viewer would
        #    otherwise show the deleted demo figures verbatim).
        if not keep_figure_lists:
            children = list(body)
            lof = H.find_index(children, lambda el: H.is_p(el) and H.norm(H.para_text(el)) == "List of Figures")
            if lof is not None:
                sect_end = H.find_index(children,
                                        lambda el: H.is_p(el) and el.find(W + "pPr/" + W + "sectPr") is not None,
                                        start=lof)
                if sect_end is not None:
                    n = H.remove_range(body, children, lof, sect_end)
                    removed.append("List of Figures/Tables (%d elements)" % n)
        else:
            if H.reset_figure_list(body, "Figure", "List of Figures"):
                removed.append("List of Figures cache reset (dirty field)")
            if H.reset_figure_list(body, "Table", "List of Tables"):
                removed.append("List of Tables cache reset (dirty field)")

        # 2b) TOC cache lists the deleted demo headings: replace with the
        #     template's own heading + a dirty TOC field.
        H.reset_toc(body, "Table of Contents")

        # 3) Internal variant: drop the client confidentiality page.
        if internal:
            children = list(body)
            conf = H.find_index(children, lambda el: H.is_p(el) and "Statement of Confidentiality" in H.para_text(el))
            if conf is not None:
                rev = H.find_index(children, lambda el: H.is_p(el) and "Document Revision History" in H.para_text(el), start=conf)
                if rev is not None:
                    n = H.remove_range(body, children, conf, rev)
                    removed.append("confidentiality page (%d elements)" % n)

        # 4) Author / role on the revision-history page.
        filled = fill_author(doc, author, role)
        if filled:
            removed.append("filled " + ", ".join(filled))
    return removed


def finish_document(doc, out, title, company, date):
    """Prune numbering, set dc:title, save, then run the zip-level passes in
    the documented order (set_bound_fields first among them, all after the
    save; nothing may re-save the file afterwards).  Returns
    {"counts": {...}, "nums": n, "abstracts": n}."""
    nums, abstracts = H.prune_numbering(doc)
    doc.core_properties.title = title
    # tmp + move, not doc.save(out) directly: a process death mid-save (OOM,
    # crash, power loss) must never leave a truncated file at the path a
    # caller expects a finished deliverable -- the same guarantee every
    # zip-level patch below already gets from H.set_bound_fields et al.
    tmp = out + ".tmp"
    doc.save(tmp)
    shutil.move(tmp, out)

    # ---- zip-level patches (parts python-docx does not expose) ----
    counts = H.set_bound_fields(out, title=title, company=company, date=date)
    H.patch_settings_updatefields(out)
    H.patch_footer_dateformat(out)
    H.scrub_metadata(out)
    return {"counts": counts, "nums": nums, "abstracts": abstracts}


def build_document(base, out, title, company=None, date=None, internal=False, minimal=False,
                   keep_figure_lists=False, author=None, role=None, fill=None):
    """Library entry point: shell + optional `fill(doc)` callback for body
    content + finish.  Returns a dict with company, date, removed, counts."""
    company, date = resolve_company_date(company, internal, date)
    doc = docx.Document(base)
    removed = prepare_document_shell(doc, internal=internal, minimal=minimal,
                                     keep_figure_lists=keep_figure_lists, author=author, role=role)
    if fill is not None:
        fill(doc)
    info = finish_document(doc, out, title, company, date)
    info.update({"company": company, "date": date, "removed": removed})
    return info


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("base")
    ap.add_argument("out")
    ap.add_argument("--title", required=True)
    ap.add_argument("--company", default=None)
    ap.add_argument("--date", default=None, help="publish date (2026-09-01, 9/1/2026, 'September 2026'); default today")
    ap.add_argument("--internal", action="store_true")
    ap.add_argument("--minimal", action="store_true")
    ap.add_argument("--keep-figure-lists", action="store_true")
    ap.add_argument("--author", default=None, help="fills 'Author Name' and the revision table Author cells")
    ap.add_argument("--role", default=None, help="fills 'Job Position'")
    args = ap.parse_args(argv)

    try:
        info = build_document(args.base, args.out, args.title, company=args.company, date=args.date,
                              internal=args.internal, minimal=args.minimal,
                              keep_figure_lists=args.keep_figure_lists, author=args.author, role=args.role)
    except ShellError as exc:
        sys.exit(str(exc))

    counts = info["counts"]
    mode = "minimal" if args.minimal else "internal" if args.internal else "client"
    print("OK: wrote %s (mode=%s, title=%r, company=%r, date=%s)" % (
        args.out, mode, args.title, info["company"], H.date_display(info["date"])))
    print("  removed: " + "; ".join(info["removed"]))
    print("  bound SDTs rewritten: title=%d company=%d date=%d; numbering pruned: %d num, %d abstractNum" % (
        counts["title"], counts["company"], counts["date"], info["nums"], info["abstracts"]))


if __name__ == "__main__":
    main()
