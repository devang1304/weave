#!/usr/bin/env python3
"""
spec_from_file.py -- reconstruct a Weave content spec from a built .docx/.pptx
(recipes/SPEC.md "Embedding"), reconciling it against any hand edits.

  spec_from_file.py FILE [--out spec.json] [--json]

If FILE carries no embedded spec (spec_embed.read_spec returns None), the
result is a best-effort spec built from structure alone: kind guessed from
the extension (+ validate_deliverable.detect_type for .docx doc vs sow),
blocks/slides built from heading styles / slide layouts with fresh ids, no
"sow" key, weave.generator = "spec_from_file (foreign)". `foreign: true`
says so in the result.

If FILE carries an embedded spec, every block/slide is matched to its live
element by bookmark / slide-name id (spec_embed.id_from_bookmark /
id_from_slide_name). Per match:
  - live text != stored text  -> the FILE wins: the block/slide is updated
    in place and marked "edited": true (a key on the block/slide dict, not a
    side structure, so the returned spec rebuilds directly as given).
  - a live id with no matching spec block/slide -> appended as a new block
    with a fresh id, listed under drift.added.
  - a spec block/slide with no matching live id -> dropped (removed by
    hand), noted under drift.lost_if_rebuilt.
  - content the walker cannot faithfully re-express (merged table cells,
    run-level bold/italic collapsed to plain text by a hand edit, slide
    content that changed beyond its title) -> also drift.lost_if_rebuilt,
    a list of {"location", "why"}.
Orphan-bookmark scanning (drift.added) runs for doc/deck files only: a SOW's
blocks live nested under sow.sections.*, and there is no reliable way to
tell which section an orphan bookmark belonged to.

SOW fee/milestone/rate/deliverable/timeline tables are never rebuilt from
the live document (too fragile, and risks silently corrupting a live
formula field): sow.* stays exactly as stored; only drift.sow_tables_differ
(bool) says whether those tables' visible text still matches what the
stored spec would produce.

Result: {"ok": true, "spec": {...}, "foreign": bool,
         "drift": {"added": [...], "lost_if_rebuilt": [...],
                   "sow_tables_differ": bool|null}}
Always prints this object (indented without --json, single line with).
--out writes just the "spec" value (ready to feed back to
build_deliverable.py --spec) to that file.

Requires python-docx, python-pptx, lxml. Python 3.9+.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys
import traceback
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import docx  # noqa: E402
import pptx  # noqa: E402

import nw_docx_helpers as H  # noqa: E402
from nw_docx_helpers import W  # noqa: E402
import spec_embed  # noqa: E402
import validate_deliverable as V  # noqa: E402

# nw_docx_helpers (H/W) is a genuine dependency for BOTH file types --
# _reconcile_pptx below reuses H.norm()/H.para_text() as shared text-
# normalization utilities even for deck files -- so it stays a plain,
# unconditional import (every skill's manifest needs nw_docx_helpers.py
# regardless of kind). nw_pptx_helpers (PH) is genuinely deck-only and is
# loaded lazily so a doc/sow-only skill's Copilot Cowork per-skill copy
# (which never carries nw_pptx_helpers.py) doesn't ModuleNotFoundError.
PH = None


def _ensure_ext_imports(is_pptx):
    global PH
    if is_pptx and PH is None:
        try:
            import nw_pptx_helpers as _PH
        except ImportError as exc:
            raise RuntimeError(
                "this copy of Weave is missing nw_pptx_helpers.py, needed to reconcile a "
                "PowerPoint file (%s) -- likely a per-skill packaging limitation" % exc)
        PH = _PH

SOW_SECTION_NAMES = ("executive_summary", "scope", "out_of_scope", "assumptions_add", "roles_notes")
# Layouts whose entire visible content is (title +) body (+ source): the
# only shape the deck reconciler's text fallback can fully reconstruct.
SIMPLE_BODY_LAYOUTS = {"Title and Content", "Section Header", "Title Only", "Title Only (Centered)"}


# ---------------------------------------------------------------------------
# foreign files (no embedded spec): best-effort reconstruction
# ---------------------------------------------------------------------------
def _guess_kind(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".pptx", ".potx"):
        return "deck"
    if ext not in (".docx", ".dotx"):
        return "doc"
    try:
        return V.detect_type(Path(path))
    except Exception:  # noqa: BLE001 - best effort only
        return "doc"


def _foreign_doc_blocks(doc):
    blocks = []
    counter = [0]
    bullets = []

    def next_id():
        counter[0] += 1
        return "b%d" % counter[0]

    def flush_bullets():
        if bullets:
            blocks.append({"id": next_id(), "type": "bullets", "items": list(bullets)})
            bullets[:] = []

    for el in doc.element.body:
        if el.tag == W + "p":
            if H.is_empty_p(el):
                continue
            text = H.norm(H.para_text(el))
            if not text:
                continue
            style = H.p_style(el) or "Normal"
            m = re.match(r"(?:NW)?Heading(\d)$", style)
            if m:
                flush_bullets()
                blocks.append({"id": next_id(), "type": "heading", "level": int(m.group(1)), "text": text})
            elif style == "ListParagraph":
                bullets.append(text)
            else:
                flush_bullets()
                blocks.append({"id": next_id(), "type": "para", "text": text})
        elif el.tag == W + "tbl":
            flush_bullets()
            rows = H.table_rows(el)
            if rows:
                header = [H.norm(H.para_text(tc)) for tc in H.row_cells(rows[0])]
                data = [[H.norm(H.para_text(tc)) for tc in H.row_cells(tr)] for tr in rows[1:]]
                blocks.append({"id": next_id(), "type": "table", "header": header, "rows": data})
    flush_bullets()
    return blocks


def _foreign_deck_slides(prs):
    slides = []
    for i, slide in enumerate(prs.slides, start=1):
        layout = PH.layout_name(slide)
        if layout not in PH.LAYOUTS:
            layout = "Blank"
        slides.append({"id": "s%d" % i, "layout": layout, "title": PH.slide_title_text(slide)})
    return slides


def _foreign_spec(path, kind, is_pptx):
    now = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    spec = {
        "weave": {"spec_version": "2.0", "kind": kind, "recipe": None, "base": None, "revision": 1,
                 "generator": "spec_from_file (foreign)", "built_at": now},
        "meta": {"title": "", "client": None, "client_slug": None, "date": None, "internal": False,
                "author": "", "role": "", "label": None, "keep_figure_lists": False},
        "blocks": [], "slides": [], "sow": None,
    }
    if is_pptx:
        prs = pptx.Presentation(path)
        spec["meta"]["title"] = (prs.core_properties.title or "").strip()
        spec["slides"] = _foreign_deck_slides(prs)
    else:
        doc = docx.Document(path)
        spec["meta"]["title"] = (doc.core_properties.title or "").strip()
        spec["blocks"] = _foreign_doc_blocks(doc)
    return spec


# ---------------------------------------------------------------------------
# docx block reconciliation
# ---------------------------------------------------------------------------
def _bookmark_ranges(body):
    """{safe_id: (first_content_el, last_content_el)} for every _weave_*
    bookmark pair (see spec_embed.mark_block: the markers themselves are
    siblings adjacent to, not wrapping, the content)."""
    children = list(body)
    starts = {}
    out = {}
    for i, el in enumerate(children):
        if el.tag == W + "bookmarkStart":
            bid = spec_embed.id_from_bookmark(el.get(W + "name"))
            if bid is not None:
                starts[el.get(W + "id")] = (i, bid)
        elif el.tag == W + "bookmarkEnd":
            wid = el.get(W + "id")
            if wid in starts:
                start_i, bid = starts.pop(wid)
                content = children[start_i + 1:i]
                if content:
                    out[bid] = (content[0], content[-1])
    return out


def _live_text_of_range(first_el, last_el):
    parts = []
    cur = first_el
    while True:
        parts.append(H.para_text(cur))
        if cur is last_el:
            break
        cur = cur.getnext()
    return H.norm(" ".join(parts))


def _expected_text_of_block(block):
    t = block.get("type")
    if t == "heading":
        return H.norm(block.get("text") or "")
    if t == "para":
        if block.get("runs"):
            return H.norm("".join(r.get("text") or "" for r in block["runs"]))
        return H.norm(block.get("text") or "")
    if t == "bullets":
        return H.norm(" ".join(block.get("items") or []))
    if t == "callout":
        return H.norm(block.get("text") or "")
    return None  # table, figure, page_break: compared separately or not at all


def _table_has_merges(tbl_el):
    for tc in tbl_el.iter(W + "tc"):
        tcpr = tc.find(W + "tcPr")
        if tcpr is not None and (tcpr.find(W + "gridSpan") is not None or tcpr.find(W + "vMerge") is not None):
            return True
    return False


def _extract_table(tbl_el):
    rows = H.table_rows(tbl_el)
    if not rows:
        return [], []
    header = [H.norm(H.para_text(tc)) for tc in H.row_cells(rows[0])]
    data = [[H.norm(H.para_text(tc)) for tc in H.row_cells(tr)] for tr in rows[1:]]
    return header, data


_CAPTION_PREFIX_RE = re.compile(r"^(?:Table|Figure)\s+\d+\s*", re.I)


def _live_caption_text(caption_el):
    """add_caption() always prepends the SEQ field's cached 'Table N '/
    'Figure N ' to the paragraph; strip it so the comparison is against the
    same user-supplied caption text the spec stores (never the number)."""
    return _CAPTION_PREFIX_RE.sub("", H.norm(H.para_text(caption_el)), count=1).strip()


def _reconcile_block(block, first_el, last_el, lost):
    bid = block.get("id", "?")
    bt = block.get("type")

    if bt == "table" and H.is_tbl(first_el):
        if _table_has_merges(first_el):
            lost.append({"location": "block %s" % bid, "why": "table has merged cells; a rebuild would lose them"})
        else:
            header, rows = _extract_table(first_el)
            expected = H.norm(" ".join([str(x) for x in (block.get("header") or [])] +
                                       [str(c) for row in (block.get("rows") or []) for c in row]))
            live = H.norm(" ".join(header + [c for row in rows for c in row]))
            if live != expected:
                block["header"], block["rows"], block["edited"] = header, rows, True
        caption_el = last_el if last_el is not first_el else None
        if caption_el is not None:
            live_cap = _live_caption_text(caption_el)
            if live_cap != H.norm(block.get("caption") or ""):
                block["caption"], block["edited"] = live_cap, True
        return

    if bt == "figure":
        caption_el = last_el if last_el is not first_el else None
        if caption_el is not None:
            live_cap = _live_caption_text(caption_el)
            if live_cap != H.norm(block.get("caption") or ""):
                block["caption"], block["edited"] = live_cap, True
        return

    if bt == "page_break":
        return

    expected = _expected_text_of_block(block)
    if expected is None:
        return
    live = _live_text_of_range(first_el, last_el)
    if live != expected:
        block["edited"] = True
        if bt == "bullets":
            items = []
            cur = first_el
            while True:
                items.append(H.norm(H.para_text(cur)))
                if cur is last_el:
                    break
                cur = cur.getnext()
            block["items"] = items
        else:
            if bt == "para" and block.get("runs"):
                lost.append({"location": "block %s" % bid,
                            "why": "run-level bold/italic formatting was collapsed to plain text after a hand edit"})
                block.pop("runs", None)
            block["text"] = live


def _prune_missing(blocks, found_safe_ids, lost):
    keep = []
    for b in blocks:
        if spec_embed.safe_id(b.get("id")) in found_safe_ids:
            keep.append(b)
        else:
            lost.append({"location": "block %s" % b.get("id"), "why": "no longer found in the file (removed by hand)"})
    return keep


def _sow_tables_differ(body, sow_spec):
    """Best-effort: never rebuilds sow.* from the live document (SPEC.md);
    only reports whether the fee/deliverable/timeline tables' visible text
    still contains what the stored spec would produce. A pure containment
    check (never a row count): the fee/rate tables carry fixed template
    rows alongside the data rows (Sub-Total, Client Discount, License Cost,
    Total...), so "how many rows" is not a fair comparison against "how
    many milestones/rates the spec lists"."""
    sow_spec = sow_spec or {}
    variant = sow_spec.get("variant")
    checks = []
    if variant == "milestone":
        tbl = H.find_table(body, header_startswith=["Milestone", "Milestone Description"])
        checks.append((tbl, sow_spec.get("milestones") or [],
                       lambda m: str(m.get("name") or m.get("description") or "")))
    elif variant == "tm":
        tbl = H.find_table(body, header_startswith=["Location"])
        checks.append((tbl, sow_spec.get("rates") or [], lambda r: str(r.get("role") or "")))
    grid = H.find_table(body, style="PSOGrid")
    checks.append((grid, sow_spec.get("deliverables") or [], lambda row: str(row[0]) if row else ""))
    timeline = H.find_table(body, style="NetwovenTable1", header_startswith=["Phase", "Description", "Duration"])
    checks.append((timeline, sow_spec.get("timeline") or [], lambda row: str(row[0]) if row else ""))

    for tbl, items, key_fn in checks:
        if tbl is None or not items:
            continue
        live_text = H.norm(H.para_text(tbl)).lower()
        for item in items:
            key = H.norm(key_fn(item)).lower()
            if key and key not in live_text:
                return True
    return False


def _reconcile_docx(path, spec, drift):
    doc = docx.Document(path)
    body = doc.element.body
    lost = drift["lost_if_rebuilt"]
    kind = (spec.get("weave") or {}).get("kind")

    all_blocks = list(spec.get("blocks") or [])
    if isinstance(spec.get("sow"), dict):
        sections = spec["sow"].get("sections") or {}
        for name in SOW_SECTION_NAMES:
            all_blocks.extend(sections.get(name) or [])

    by_safe_id = {}
    for b in all_blocks:
        by_safe_id.setdefault(spec_embed.safe_id(b.get("id")), []).append(b)

    ranges = _bookmark_ranges(body)
    found_safe_ids = set()
    for safe_id, (first_el, last_el) in ranges.items():
        candidates = by_safe_id.get(safe_id)
        if not candidates:
            if kind != "sow":  # a sow's blocks live nested under sections.*; no reliable "which section"
                text = _live_text_of_range(first_el, last_el)
                spec.setdefault("blocks", []).append({"id": safe_id, "type": "para", "text": text})
                drift["added"].append(safe_id)
            continue
        found_safe_ids.add(safe_id)
        for block in candidates:
            _reconcile_block(block, first_el, last_el, lost)

    if spec.get("blocks"):
        spec["blocks"] = _prune_missing(spec["blocks"], found_safe_ids, lost)
    if isinstance(spec.get("sow"), dict):
        sections = spec["sow"].get("sections") or {}
        for name in SOW_SECTION_NAMES:
            if sections.get(name):
                sections[name] = _prune_missing(sections[name], found_safe_ids, lost)
        drift["sow_tables_differ"] = _sow_tables_differ(body, spec["sow"])
    return spec


def _reconcile_pptx(path, spec, drift):
    prs = pptx.Presentation(path)
    lost = drift["lost_if_rebuilt"]
    by_id = {s.get("id"): s for s in (spec.get("slides") or [])}
    found = set()

    for slide in prs.slides:
        wid = spec_embed.id_from_slide_name(slide.name)
        if wid is None:
            continue  # cover / confidentiality / closer: not spec-tracked (SPEC.md)
        stored_slide = by_id.get(wid)
        if stored_slide is None:
            new_slide = {"id": wid, "layout": PH.layout_name(slide), "title": PH.slide_title_text(slide)}
            spec.setdefault("slides", []).append(new_slide)
            drift["added"].append(wid)
            continue
        found.add(wid)
        layout = stored_slide.get("layout") or ""
        # Only layouts with a real title placeholder carry a comparable
        # "title" -- slide_title_text() synthesizes a stand-in for the rest
        # (e.g. Three Stat's stat labels joined with "/"), which is not a
        # stored spec field and must never be compared against or written.
        if "title" in PH.LAYOUTS.get(layout, {}):
            live_title = PH.slide_title_text(slide)
            if H.norm(live_title) != H.norm(stored_slide.get("title") or ""):
                stored_slide["title"] = live_title
                stored_slide["edited"] = True
        # The broad "did anything change" fallback only works for the plain
        # title+body(+source) layouts: it is the only shape of content this
        # walker can fully re-render as text to compare against. Any other
        # layout (Two Content, Comparison, Three Column, Three Stat, Big
        # Statement...) uses fields this check does not model (columns,
        # stats, statement...) and a visual's own text (chart/table) is
        # never reconstructed either -- both would be permanent false
        # positives, so those slides rely on the title comparison alone.
        visual = stored_slide.get("visual") or {}
        if layout in SIMPLE_BODY_LAYOUTS and visual.get("type") not in ("chart", "table") and not stored_slide.get("edited"):
            expected_all = H.norm(" ".join(
                [stored_slide.get("title") or ""] + list(stored_slide.get("body") or []) +
                ([stored_slide["source"]] if stored_slide.get("source") else [])))
            if H.norm(PH.slide_text(slide)) != expected_all:
                stored_slide["edited"] = True
                lost.append({"location": "slide %s" % wid,
                            "why": "content changed beyond the title; body text was not re-extracted"})

    if spec.get("slides"):
        keep = []
        for s in spec["slides"]:
            if s.get("id") in found or s.get("id") in drift["added"]:
                keep.append(s)
            else:
                lost.append({"location": "slide %s" % s.get("id"), "why": "no longer found in the file (removed by hand)"})
        spec["slides"] = keep
    return spec


# ---------------------------------------------------------------------------
def reconcile(path):
    ext = os.path.splitext(path)[1].lower()
    is_pptx = ext in (".pptx", ".potx")
    _ensure_ext_imports(is_pptx)
    stored = spec_embed.read_spec(path)
    drift = {"added": [], "lost_if_rebuilt": [], "sow_tables_differ": None}

    if stored is None:
        kind = _guess_kind(path)
        spec = _foreign_spec(path, kind, is_pptx)
        return {"ok": True, "spec": spec, "foreign": True, "drift": drift}

    spec = json.loads(json.dumps(stored))  # deep copy; freely mutated below
    if is_pptx:
        spec = _reconcile_pptx(path, spec, drift)
    else:
        spec = _reconcile_docx(path, spec, drift)
    _remember_hand_edits(stored, spec, drift)
    return {"ok": True, "spec": spec, "foreign": False, "drift": drift}


def _remember_hand_edits(stored, spec, drift):
    """Hand the differences to memory.py (the local preference memory) so a
    repeated hand edit can become a preference. Best effort: memory is a side
    channel, absent on some packaging surfaces, and never changes the result."""
    try:
        import memory  # noqa: WPS433 - sibling stdlib script beside this file

        memory.record(memory.observe(stored, spec, drift))
        memory.promote()
    except Exception:  # noqa: BLE001 - see docstring
        pass


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("--out", default=None)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    if not os.path.exists(args.file):
        print(json.dumps({"ok": False, "error": "file not found: %s" % args.file, "hint": "Check the path."}))
        return 1

    try:
        result = reconcile(os.path.abspath(args.file))
    except Exception as exc:  # noqa: BLE001 - top-level guard: always emit the JSON contract
        traceback.print_exc()
        print(json.dumps({"ok": False, "error": "unexpected error: %s" % exc,
                          "hint": "Run again without --json to see the traceback on stderr."}))
        return 1

    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(result["spec"], fh, indent=2, ensure_ascii=False)
            fh.write("\n")

    print(json.dumps(result, indent=None if args.json else 2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
