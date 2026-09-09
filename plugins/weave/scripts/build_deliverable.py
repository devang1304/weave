#!/usr/bin/env python3
"""
build_deliverable.py -- build a Netwoven deliverable (doc / deck / sow) from
a Weave content spec (recipes/SPEC.md), embed the spec for later editing, and
validate the result.

  build_deliverable.py --spec FILE [--out FILE] [--replace]
      [--client-slug S] [--recipe R] [--json]

Dispatches on spec["weave"]["kind"]:
  doc  : walks spec["blocks"] in order, writing each with the docx helpers,
         inside the fill(doc) callback new_deliverable_docx.build_document()
         runs before the shell is saved.
  deck : builds the shell (cover / confidentiality / closers) with
         new_deck_pptx.build_shell(), then walks spec["slides"] adding one
         slide per entry via nw_pptx_helpers, then new_deck_pptx.finalize_deck().
  sow  : new_sow_docx.build_sow() with a before_save(doc, log) callback that
         writes spec["sow"]["sections"] prose into the template's known slots.

Every block/slide id becomes a hidden bookmark / slide name via spec_embed,
so a later hand edit can be reconciled by spec_from_file.py. The full spec
(with weave.revision/built_at/generator resolved) is embedded with
spec_embed.embed_spec() AFTER the file's own save, since embed_spec works
directly on the zip and nothing may re-save the package afterward.

When spec.weave.recipe names a file under shared/recipes/ with a .json half,
this loads and enforces it (recipes/SPEC.md "Recipe files"): title_pattern
fills a blank meta.title; blocks[].framework / visual.framework resolve from
the recipe's frameworks table; sections[].required missing from the spec is
a BuildError; validator.require_headings/owners_and_dates_in and a missing
required input become validation warnings, never hard failures. A --recipe
naming a file with no .json half yet (Phase 1's generic-doc/generic-deck/sow)
is accepted but not enforced, same as before.

Hard content rules (SPEC.md "Blocks") are checked for every block BEFORE any
writing starts, so a bad spec never produces a half-written file: heading
levels never skip; the first block of a client (non-internal) document is a
level-1 heading; no literal "Table N"/"Figure N" in a caption (the SEQ field
numbers it); no em/en dash anywhere unless meta.allow_dashes is true. Inside
sow.sections.*, heading blocks are only allowed in assumptions_add, and only
at level 3 (recipes/sow.md); every other section is para/bullets/table/etc.

Prints exactly one JSON object: on success (always, --json or not, though it
is only PRINTED without --json as a short human summary line)
  {"ok": true, "out", "kind", "revision", "recipe", "validation":
   {"hard_fails", "warnings", "results"}, "flags": [...]}
A build that produces content problems (validation.hard_fails > 0) still
exits 0: the file was written and the caller can inspect/fix it. Exit 1 is
reserved for a build that could not be produced at all (bad spec, missing
base, an exception) and ALWAYS prints {"ok": false, "error", "hint"} as one
JSON object, matching weave.py's "the whole stdout parses as one JSON
object, or the last line does" contract -- regardless of --json.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import glob
import json
import os
import re
import sys
import traceback
from types import SimpleNamespace

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx.enum.text import WD_BREAK  # noqa: E402
import pptx  # noqa: E402
from pptx.util import Inches  # noqa: E402

import nw_docx_helpers as H  # noqa: E402
from nw_docx_helpers import W  # noqa: E402
import spec_embed  # noqa: E402
import validate_deliverable as V  # noqa: E402
import weave as WV  # noqa: E402

# nw_docx_helpers (H/W) stays a plain, unconditional top-level import: every
# skill's manifest.json entry ships it regardless of kind (spec_from_file.py
# needs it unconditionally too, and _default_out_path_and_revision below
# calls H.parse_date() for every kind, not just doc/sow -- an earlier,
# kind-gated version of this import left H unset for "deck", which crashed
# any deck build that omitted --out with a real meta.date, e.g. every
# companion-pair skill's second (deck) build call).
#
# new_deliverable_docx/new_sow_docx/nw_pptx_helpers/new_deck_pptx are
# genuinely kind-specific and stay lazy, imported by _ensure_kind_imports()
# below: a Copilot Cowork per-skill copy (manifest.json's m365 target) only
# carries the builder scripts that ONE skill's own kind(s) need, so importing
# all four unconditionally would ModuleNotFoundError the instant a doc-only
# skill's copy of this file loads (it never got new_deck_pptx.py). Each is
# wrapped so a genuinely-missing module raises a clear BuildError naming the
# gap, instead of a bare ModuleNotFoundError deep inside this function.
ND = NS = PH = NDeck = None


def _ensure_kind_imports(kind):
    """Import only the builder module(s) `kind` actually needs. Idempotent --
    safe to call once per kind per process, including once per kind across a
    mixed doc+deck companion-pair run."""
    global ND, NS, PH, NDeck
    if kind == "doc" and ND is None:
        try:
            import new_deliverable_docx as _ND
        except ImportError as exc:
            raise BuildError(
                "this copy of Weave is missing new_deliverable_docx.py, needed to build a "
                "'doc' kind file (%s) -- likely a per-skill packaging limitation" % exc)
        ND = _ND
    if kind == "sow" and NS is None:
        try:
            import new_sow_docx as _NS
        except ImportError as exc:
            raise BuildError(
                "this copy of Weave is missing new_sow_docx.py, needed to build a 'sow' "
                "kind file (%s) -- likely a per-skill packaging limitation" % exc)
        NS = _NS
    if kind == "deck" and PH is None:
        try:
            import nw_pptx_helpers as _PH
        except ImportError as exc:
            raise BuildError(
                "this copy of Weave is missing nw_pptx_helpers.py, needed to build a 'deck' "
                "kind file (%s) -- likely a per-skill packaging limitation" % exc)
        PH = _PH
    if kind == "deck" and NDeck is None:
        try:
            import new_deck_pptx as _NDeck
        except ImportError as exc:
            raise BuildError(
                "this copy of Weave is missing new_deck_pptx.py, needed to build a 'deck' "
                "kind file (%s) -- likely a per-skill packaging limitation" % exc)
        NDeck = _NDeck


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASES_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "assets", "bases"))
RECIPES_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "recipes"))
EXT_BY_KIND = {"doc": ".docx", "sow": ".docx", "deck": ".pptx"}
DASH_CHARS = "–—"  # en dash, em dash
CAPTION_BAD_RE = re.compile(r"\b(table|figure)\s+\d+\b", re.I)
BODY_BLOCK_HEIGHT = Inches(1.5)
_SILENT = lambda *a, **k: None  # noqa: E731 - suppress the builders' own print(...) logging


class BuildError(Exception):
    """Raised for anything that means no file can be produced; main() turns
    it into the {"ok": false, "error", "hint"} envelope."""


# ---------------------------------------------------------------------------
# spec loading / naming / revision
# ---------------------------------------------------------------------------
def load_spec(path):
    if not os.path.exists(path):
        raise BuildError("spec file not found: %s" % path)
    try:
        with open(path, "r", encoding="utf-8") as fh:
            spec = json.load(fh)
    except (OSError, ValueError) as exc:
        raise BuildError("could not read spec %s: %s" % (path, exc))
    if not isinstance(spec, dict):
        raise BuildError("spec %s must be a JSON object" % path)
    return spec


def resolve_base_path(base_name):
    if not base_name:
        raise BuildError("spec.weave.base is required")
    path = os.path.join(BASES_DIR, base_name)
    if not os.path.exists(path):
        raise BuildError("base %r not found under %s" % (base_name, BASES_DIR))
    return path


def _existing_revisions(prefix, ext):
    """weave.revision embedded in every existing '<prefix>_v*<ext>' file in the
    current directory -- read from each file itself, never trusted from the
    file's OWN name, since a stale filename next to a since-advanced embedded
    revision is exactly the bug this exists to catch. A same-prefix file that
    cannot be read as a Weave file is skipped, not treated as revision 0."""
    revisions = []
    pattern = glob.escape(prefix) + "_v*" + glob.escape(ext)
    for path in glob.glob(pattern):
        try:
            old = spec_embed.read_spec(path)
        except Exception:  # noqa: BLE001 - a same-prefix file that isn't a readable Weave file
            continue
        if old and isinstance(old.get("weave"), dict):
            try:
                revisions.append(int(old["weave"].get("revision") or 0))
            except (TypeError, ValueError):
                pass
    return revisions


def _default_out_path_and_revision(spec, kind):
    """(path, revision) for the auto-generated filename (see default_out_path).
    The revision baked into the filename is the TRUE next revision -- one
    past the highest weave.revision embedded in any existing same-prefix
    file -- not the incoming spec's own (often stale / never-updated by the
    caller) weave.revision. Without this, repeated --replace rebuilds with no
    explicit --out kept re-resolving to the very same filename (silently
    overwriting the prior file) while only the EMBEDDED revision climbed."""
    meta = spec.get("meta") or {}
    weave = spec.get("weave") or {}
    client_slug = meta.get("client_slug") or WV.slugify(meta.get("client") or "")
    client_slug = client_slug or "deliverable"
    type_slug = WV.slugify(str(weave.get("recipe") or kind)) or kind
    d = H.parse_date(meta.get("date")) if meta.get("date") else None
    if d is None:
        d = _dt.date.today()
    ext = EXT_BY_KIND[kind]
    prefix = "%s_%s_%s" % (client_slug, type_slug, d.strftime("%Y-%m-%d"))
    try:
        incoming = int(weave.get("revision") or 1)
    except (TypeError, ValueError):
        incoming = 1
    existing = _existing_revisions(prefix, ext)
    revision = (max(existing) + 1) if existing else incoming
    return "%s_v%s%s" % (prefix, revision, ext), revision


def default_out_path(spec, kind):
    """<ClientShort>_<Type>_<YYYY-MM-DD>_v<revision>.<ext> (SPEC.md "Script
    contract"), revision resolved as described in _default_out_path_and_revision."""
    path, _revision = _default_out_path_and_revision(spec, kind)
    return path


def resolve_revision(out_path, spec):
    """New revision for this build: bumped from whatever is already embedded
    in `out_path` when it exists, else the incoming spec's own revision
    (default 1) unmodified."""
    incoming = (spec.get("weave") or {}).get("revision")
    try:
        incoming = int(incoming) if incoming is not None else 1
    except (TypeError, ValueError):
        incoming = 1
    if os.path.exists(out_path):
        old = spec_embed.read_spec(out_path)
        if old and isinstance(old.get("weave"), dict):
            try:
                return int(old["weave"].get("revision") or 0) + 1
            except (TypeError, ValueError):
                pass
    return incoming


# ---------------------------------------------------------------------------
# hard content rules (shared by doc blocks and sow section blocks)
# ---------------------------------------------------------------------------
def _texts_of_block(block):
    """Yield (field label, text) for every text-bearing field of a block."""
    t = block.get("type")
    if t == "heading":
        yield "text", block.get("text") or ""
    elif t == "para":
        if block.get("runs"):
            for i, r in enumerate(block["runs"]):
                yield "runs[%d].text" % i, r.get("text") or ""
        else:
            yield "text", block.get("text") or ""
    elif t == "bullets":
        for i, item in enumerate(block.get("items") or []):
            yield "items[%d]" % i, item
    elif t == "table":
        for i, hcell in enumerate(block.get("header") or []):
            yield "header[%d]" % i, hcell
        for ri, row in enumerate(block.get("rows") or []):
            for ci, cell in enumerate(row):
                yield "rows[%d][%d]" % (ri, ci), "" if cell is None else str(cell)
        if block.get("caption"):
            yield "caption", block["caption"]
    elif t == "figure":
        if block.get("caption"):
            yield "caption", block["caption"]
        if block.get("alt"):
            yield "alt", block["alt"]
    elif t == "callout":
        yield "text", block.get("text") or ""


def check_block_text_rules(block, allow_dashes):
    bid = block.get("id", "?")
    for field, text in _texts_of_block(block):
        text = text or ""
        if not allow_dashes:
            hit = next((c for c in text if c in DASH_CHARS), None)
            if hit:
                raise BuildError(
                    "block %r: %s contains an em/en dash (%r); set meta.allow_dashes to allow it"
                    % (bid, field, text[:70]))
        if field == "caption" and CAPTION_BAD_RE.search(text):
            raise BuildError(
                "block %r: caption must not include a literal 'Table N'/'Figure N' (%r); "
                "the SEQ field numbers it automatically" % (bid, text[:70]))


def _check_unique_ids(blocks, where, seen_ids=None, seen_safe=None):
    """Raise BuildError if any two blocks share a raw id, or if two DIFFERENT
    raw ids collide once run through spec_embed.safe_id() -- the sanitized,
    32-char-truncated name spec_embed.mark_block() bookmarks under. Either
    collision is invisible until the second block's write silently deletes
    the first block's same-named bookmark (mark_block() removes any existing
    bookmark with that name before writing the new one), destroying its
    content with no error.

    `seen_ids`/`seen_safe` map an already-seen raw id / safe id back to
    (raw_id, where) of its first sighting. Passing them in (instead of
    leaving them to default to fresh, per-call dicts) lets callers share one
    pool across several lists checked one after another -- e.g. every
    sow.sections.* list in one SOW build -- so a collision spanning two
    different `where`s is also caught, and named in the error."""
    if seen_ids is None:
        seen_ids = {}
    if seen_safe is None:
        seen_safe = {}
    for b in blocks:
        bid = b.get("id")
        if not bid:
            raise BuildError("%s: every block needs a non-empty id" % where)
        if bid in seen_ids:
            prev_where = seen_ids[bid]
            location = "" if prev_where == where else (" (first seen in %s)" % prev_where)
            raise BuildError("%s: block id %r is used more than once%s" % (where, bid, location))
        safe = spec_embed.safe_id(bid)
        if safe in seen_safe:
            prev_id, prev_where = seen_safe[safe]
            location = "" if prev_where == where else (" in %s" % prev_where)
            raise BuildError(
                "%s: block id %r collides with id %r%s once both are sanitized to the same "
                "bookmark name %r (spec_embed.safe_id keeps only letters/digits/underscore and "
                "truncates to 32 chars) -- rename one of the two ids so they stay distinct after "
                "sanitization" % (where, bid, prev_id, location, spec_embed.bookmark_name(bid)))
        seen_ids[bid] = where
        seen_safe[safe] = (bid, where)
    return seen_ids, seen_safe


def validate_doc_blocks(blocks, internal, allow_dashes):
    _check_unique_ids(blocks, "blocks")
    if blocks and not internal:
        first = blocks[0]
        if not (first.get("type") == "heading" and first.get("level") == 1):
            raise BuildError(
                "block %r: the first block of a client document must be a level-1 heading"
                % first.get("id", "?"))
    prev_level = 0
    for b in blocks:
        if b.get("type") != "heading":
            continue
        level = b.get("level")
        if level not in (1, 2, 3, 4):
            raise BuildError("block %r: heading level must be 1-4 (got %r)" % (b.get("id", "?"), level))
        if level > prev_level + 1:
            raise BuildError(
                "block %r: heading level %d skips a level (previous heading was level %d)"
                % (b.get("id", "?"), level, prev_level))
        prev_level = level
    for b in blocks:
        check_block_text_rules(b, allow_dashes)


def validate_section_blocks(blocks, section_name, allow_dashes, seen_ids=None, seen_safe=None):
    """sow.sections.*: heading blocks are only allowed inside assumptions_add
    (recipes/sow.md), and only at level 3 (SPEC.md). `seen_ids`/`seen_safe`
    are the shared id pool passed in by build_sow_kind's loop over every
    sections.* list in this SOW build, so a raw-id or safe-id collision
    spanning two different sections (not just within one) is caught too."""
    _check_unique_ids(blocks, "sow.sections.%s" % section_name, seen_ids, seen_safe)
    for b in blocks:
        if b.get("type") == "heading":
            if section_name != "assumptions_add":
                raise BuildError(
                    "sow.sections.%s block %r: heading blocks are only allowed in assumptions_add"
                    % (section_name, b.get("id", "?")))
            if b.get("level") != 3:
                raise BuildError(
                    "sow.sections.%s block %r: heading level must be 3 (got %r)"
                    % (section_name, b.get("id", "?"), b.get("level")))
        check_block_text_rules(b, allow_dashes)


# ---------------------------------------------------------------------------
# recipe application (recipes/SPEC.md "Recipe files")
# ---------------------------------------------------------------------------
class _SafeDict(dict):
    def __missing__(self, key):
        return ""


def load_recipe(spec, kind):
    """Load and validate shared/recipes/<weave.recipe>.json when the spec
    names a recipe. Returns None when weave.recipe is empty, or when the
    named recipe has no .json half yet (Phase 1 shipped generic-doc/
    generic-deck/sow with only the .md half: --recipe is accepted for the
    output filename but nothing is enforced until a matching .json exists).
    A mismatched recipe/kind/base is a BuildError: loading a recipe against
    the wrong kind or base spec is a caller bug, not a content problem the
    flag list can surface."""
    name = (spec.get("weave") or {}).get("recipe")
    if not name:
        return None
    path = os.path.join(RECIPES_DIR, "%s.json" % name)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            recipe = json.load(fh)
    except (OSError, ValueError) as exc:
        raise BuildError("could not read recipe %s: %s" % (path, exc))
    expected = {"recipe": name, "kind": kind, "base": (spec.get("weave") or {}).get("base")}
    for field, want in expected.items():
        got = recipe.get(field)
        if got != want:
            raise BuildError(
                "recipe %r has %s %r but the spec has %s %r" % (name, field, got, field, want))
    return recipe


def _apply_title_pattern(spec, recipe):
    """meta.title defaults from recipe['title_pattern'] when the spec left
    it empty, substituting {client}/{date}/{period}/... from meta's own
    string fields. A spec-supplied title always wins; a template referencing
    a field meta doesn't have resolves to '' rather than raising -- the gap
    surfaces separately as a missing-input warning (_check_recipe_inputs),
    not as a build crash."""
    if recipe is None:
        return
    meta = spec.setdefault("meta", {})
    pattern = recipe.get("title_pattern")
    if meta.get("title") or not pattern:
        return
    values = _SafeDict((k, v) for k, v in meta.items() if isinstance(v, str))
    meta["title"] = pattern.format_map(values)


def _resolve_framework_on(container, recipe, where):
    """Mutate `container` (a table-shaped block or visual dict) in place:
    when it names a `framework`, fill in any of that framework's own keys
    the container doesn't already set itself, leaving the `framework` key in
    place so the embedded spec still records where the shape came from.
    List-valued framework fields (header) are copied, not aliased, so two
    blocks sharing one framework never share a mutable list."""
    name = container.get("framework")
    if not name:
        return
    if recipe is None:
        raise BuildError("%s: framework %r requested but no recipe is loaded (pass --recipe)" % (where, name))
    fw = (recipe.get("frameworks") or {}).get(name)
    if fw is None:
        raise BuildError(
            "%s: framework %r is not defined by recipe %r" % (where, name, recipe.get("recipe")))
    for key, value in fw.items():
        if key == "type" or key in container:
            continue
        container[key] = list(value) if isinstance(value, list) else value


def _resolve_block_frameworks(blocks, recipe):
    for b in blocks:
        if b.get("type") == "table":
            _resolve_framework_on(b, recipe, "block %r" % b.get("id", "?"))


def _resolve_slide_frameworks(slides_spec, recipe):
    for sd in slides_spec:
        visual = sd.get("visual")
        if visual and visual.get("type") == "table":
            _resolve_framework_on(visual, recipe, "slide %r" % sd.get("id", "?"))


_NAME_TOKEN_RE = re.compile(r"\b[A-Z][a-z]+\s+[A-Z][a-z]+\b")
_DATE_TOKEN_RE = re.compile(
    r"\b\d{4}-\d{2}-\d{2}\b"
    r"|\b\d{1,2}/\d{1,2}/\d{2,4}\b"
    r"|\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?\b",
    re.I)


def _check_recipe_inputs(recipe, meta):
    """SPEC.md 'inputs': a required input with no meta.* counterpart is a
    validation warning naming the recipe's own `ask` prompt -- never a hard
    failure, since a spec built without asking should still surface the gap
    in the flag list rather than fail the build outright."""
    results = []
    for inp in recipe.get("inputs") or []:
        key = inp.get("key")
        if inp.get("required") and key and not meta.get(key):
            results.append({
                "check": "recipe: input %r provided" % key, "passed": False,
                "evidence": inp.get("ask") or ("meta.%s is required by this recipe" % key),
                "hard": False,
            })
    return results


def _block_search_text(block):
    return " ".join(text for _field, text in _texts_of_block(block) if text)


def _section_found_in_blocks(blocks, section, by_id):
    sid = section.get("id")
    if sid and sid in by_id:
        return by_id[sid]
    heading = section.get("heading")
    if heading:
        level = section.get("level")
        for b in blocks:
            if (b.get("type") == "heading"
                    and (b.get("text") or "").strip().lower() == heading.strip().lower()
                    and (level is None or b.get("level") == level)):
                return b
    return None


def _section_range_text(blocks, section_block):
    """Concatenated text of a doc section: `section_block` (usually its
    heading) plus every following block up to -- not including -- the next
    heading at the same or a shallower level. Doc sections are a flat
    blocks[] list with no explicit grouping (SPEC.md "Blocks"), so a
    heading's "section" is implicit: everything after it until a heading
    that could stand beside or above it."""
    try:
        start = blocks.index(section_block)
    except ValueError:
        return _block_search_text(section_block)
    level = section_block.get("level") if section_block.get("type") == "heading" else None
    texts = [_block_search_text(section_block)]
    for b in blocks[start + 1:]:
        if level is not None and b.get("type") == "heading" and (b.get("level") or 0) <= level:
            break
        texts.append(_block_search_text(b))
    return " ".join(t for t in texts if t)


def _check_doc_recipe(recipe, blocks, meta):
    results = _check_recipe_inputs(recipe, meta)
    by_id = {b["id"]: b for b in blocks if b.get("id")}
    section_by_id = {}
    for section in recipe.get("sections") or []:
        found = _section_found_in_blocks(blocks, section, by_id)
        if section.get("required") and found is None:
            raise BuildError(
                "recipe %r: required section %r is missing from the spec"
                % (recipe.get("recipe"), section.get("id") or section.get("heading")))
        if found is not None and section.get("id"):
            section_by_id[section["id"]] = found

    validator = recipe.get("validator") or {}
    heading_texts = {(b.get("text") or "").strip().lower() for b in blocks if b.get("type") == "heading"}
    for heading in validator.get("require_headings") or []:
        present = heading.strip().lower() in heading_texts
        results.append({
            "check": "recipe: heading %r present" % heading, "passed": present,
            "evidence": "" if present else "no heading block matches %r" % heading, "hard": False,
        })

    owners_id = validator.get("owners_and_dates_in")
    if owners_id:
        section_block = section_by_id.get(owners_id) or by_id.get(owners_id)
        text = _section_range_text(blocks, section_block) if section_block else ""
        has_name, has_date = bool(_NAME_TOKEN_RE.search(text)), bool(_DATE_TOKEN_RE.search(text))
        passed = bool(section_block) and has_name and has_date
        if not section_block:
            evidence = "section %r not found" % owners_id
        elif not passed:
            evidence = "no owner-like and date-like text found in section %r" % owners_id
        else:
            evidence = ""
        results.append({
            "check": "recipe: owners and dates in %r" % owners_id, "passed": passed,
            "evidence": evidence, "hard": False,
        })
    return results


def _slide_search_text(sd):
    parts = [sd.get("title") or ""]
    parts.extend(sd.get("body") or [])
    visual = sd.get("visual") or {}
    if visual.get("type") == "table":
        for row in visual.get("rows") or []:
            parts.extend("" if c is None else str(c) for c in row)
    return " ".join(p for p in parts if p)


def _check_deck_recipe(recipe, slides_spec, meta):
    results = _check_recipe_inputs(recipe, meta)
    by_id = {sd["id"]: sd for sd in slides_spec if sd.get("id")}
    for section in recipe.get("sections") or []:
        sid = section.get("id")
        if section.get("required") and (not sid or sid not in by_id):
            raise BuildError(
                "recipe %r: required section %r is missing from the spec" % (recipe.get("recipe"), sid))

    validator = recipe.get("validator") or {}
    owners_id = validator.get("owners_and_dates_in")
    if owners_id:
        sd = by_id.get(owners_id)
        text = _slide_search_text(sd) if sd else ""
        has_name, has_date = bool(_NAME_TOKEN_RE.search(text)), bool(_DATE_TOKEN_RE.search(text))
        passed = bool(sd) and has_name and has_date
        if not sd:
            evidence = "section %r not found" % owners_id
        elif not passed:
            evidence = "no owner-like and date-like text found in slide %r" % owners_id
        else:
            evidence = ""
        results.append({
            "check": "recipe: owners and dates in %r" % owners_id, "passed": passed,
            "evidence": evidence, "hard": False,
        })
    return results


# ---------------------------------------------------------------------------
# shared block-writing engine (doc blocks + sow section blocks)
# ---------------------------------------------------------------------------
def _set_numpr(p_el, num_id, ilvl):
    ppr = p_el.find(W + "pPr")
    if ppr is None:
        ppr = H.E("pPr")
        p_el.insert(0, ppr)
    npr = H.E("numPr")
    npr.append(H.E("ilvl", {"val": str(int(ilvl))}))
    npr.append(H.E("numId", {"val": str(num_id)}))
    pstyle = ppr.find(W + "pStyle")
    if pstyle is not None:
        pstyle.addnext(npr)
    else:
        ppr.insert(0, npr)


def _capture_new_elements(body, build_fn):
    """Run build_fn() (which appends w:p/w:tbl children via python-docx) and
    return the elements it produced, in order. python-docx's Document.add_*
    always inserts new body content just before the final sectPr -- NOT at
    the old len(body), which is one past where the new content actually
    lands, since sectPr itself shifts forward -- so a before/after length
    slice silently grabs sectPr instead of the new elements. Anchoring on
    "whatever directly preceded the final sectPr before the call" sidesteps
    that: it is never itself new content, and it never moves."""
    final_sectpr = body.find(W + "sectPr")
    prev_anchor = final_sectpr.getprevious() if final_sectpr is not None else (list(body)[-1] if len(body) else None)
    build_fn()
    start = prev_anchor.getnext() if prev_anchor is not None else next(iter(body), None)
    new_els = []
    cur = start
    while cur is not None and cur is not final_sectpr:
        new_els.append(cur)
        cur = cur.getnext()
    return new_els


def _write_block(doc, block, heading_style):
    """Append one block's paragraph(s)/table at the CURRENT end of `doc`.
    Returns the list of new body-level elements it produced, in order, for
    mark_block()/relocation. `heading_style` is a "%d"-format template,
    e.g. "NW Heading %d" (doc) or "Heading %d" (sow)."""
    t = block["type"]
    bid = block.get("id", "?")

    def build():
        if t == "heading":
            doc.add_paragraph(block["text"], style=heading_style % block["level"])
        elif t == "para":
            style = block.get("style") or "Normal"
            p = doc.add_paragraph(style=style)
            runs = block.get("runs")
            if runs:
                for r in runs:
                    run = p.add_run(r.get("text") or "")
                    if r.get("bold"):
                        run.bold = True
                    if r.get("italic"):
                        run.italic = True
            else:
                p.add_run(block.get("text") or "")
        elif t == "bullets":
            items = block.get("items") or []
            levels = block.get("levels") or [0] * len(items)
            num_id = H.list_num_id(doc, "decimal" if block.get("numbered") else "bullet")
            for i, item in enumerate(items):
                p = doc.add_paragraph(item, style="List Paragraph")
                if num_id is not None:
                    _set_numpr(p._p, num_id, levels[i] if i < len(levels) else 0)
        elif t == "table":
            style = block.get("style") or "Netwoven Table 1"
            H.add_styled_table(doc, block.get("header") or [], block.get("rows") or [],
                               style=style, col_widths_in=block.get("col_widths_in"))
            if block.get("caption"):
                H.add_caption(doc, "Table", block["caption"])
        elif t == "figure":
            if not block.get("image"):
                raise BuildError("block %r: figure needs an image path" % bid)
            para, _cap = H.add_captioned_picture(doc, block["image"], block.get("caption") or "",
                                                 max_width_in=block.get("width_in", 6.5))
            if block.get("alt"):
                docPr = para._p.find(".//" + H.WP + "docPr")
                if docPr is not None:
                    docPr.set("descr", block["alt"])
        elif t == "callout":
            tbl = doc.add_table(rows=1, cols=1)
            try:
                tbl.style = doc.styles["Netwoven Table 3"]
            except KeyError:
                H.log("warning: 'Netwoven Table 3' style not found; callout left unstyled")
            tbl.rows[0].cells[0].text = block.get("text") or ""
        elif t == "page_break":
            p = doc.add_paragraph()
            p.add_run().add_break(WD_BREAK.PAGE)
        else:
            raise BuildError("block %r: unknown block type %r" % (bid, t))

    try:
        new_els = _capture_new_elements(doc.element.body, build)
    except KeyError as exc:
        raise BuildError("block %r: %s" % (bid, exc))
    if not new_els:
        raise BuildError("block %r: produced no content" % bid)
    return new_els


# ---------------------------------------------------------------------------
# doc kind
# ---------------------------------------------------------------------------
def build_doc(spec, out, recipe):
    meta = spec.get("meta") or {}
    weave = spec.get("weave") or {}
    blocks = spec.get("blocks") or []
    internal = bool(meta.get("internal"))
    allow_dashes = bool(meta.get("allow_dashes"))
    title = meta.get("title") or ""
    if not title:
        raise BuildError("meta.title is required")

    _resolve_block_frameworks(blocks, recipe)
    validate_doc_blocks(blocks, internal, allow_dashes)
    recipe_results = _check_doc_recipe(recipe, blocks, meta) if recipe else []
    base = resolve_base_path(weave.get("base"))

    def fill(doc):
        body = doc.element.body
        for block in blocks:
            new_els = _write_block(doc, block, "NW Heading %d")
            spec_embed.mark_block(body, new_els[0], new_els[-1], block["id"])

    try:
        ND.build_document(
            base, out, title,
            company=meta.get("client"), date=meta.get("date"), internal=internal,
            keep_figure_lists=bool(meta.get("keep_figure_lists")),
            author=meta.get("author") or os.environ.get("WEAVE_AUTHOR"),
            role=meta.get("role") or os.environ.get("WEAVE_ROLE"),
            fill=fill,
        )
    except ND.ShellError as exc:
        raise BuildError(str(exc))

    spec_embed.embed_spec(out, spec)
    results = V.check_docx(out, "doc", internal, meta.get("client"))
    results.extend(recipe_results)
    flags = []
    if not (meta.get("author") or os.environ.get("WEAVE_AUTHOR")):
        flags.append("no author given; the revision-history page's Author Name / Job Position were left blank")
    return _summarize(results), flags


# ---------------------------------------------------------------------------
# deck kind
# ---------------------------------------------------------------------------
def _resolve_palette(name):
    return {"categorical": PH.PALETTE_CATEGORICAL, "sequential": PH.PALETTE_SEQUENTIAL,
            "diverging": PH.PALETTE_DIVERGING}.get(name, PH.PALETTE_CATEGORICAL)


def _add_visual(slide, visual, left, top, width, height, sid):
    vtype = visual.get("type")
    if vtype == "chart":
        PH.add_chart(slide, visual.get("kind", "column"), visual.get("categories") or [],
                    visual.get("series") or {}, left, top, width, height,
                    number_format=visual.get("number_format"), palette=_resolve_palette(visual.get("palette")))
    elif vtype == "table":
        PH.add_table(slide, visual.get("header") or [], visual.get("rows") or [], left, top, width, height,
                    col_widths=visual.get("col_widths_in"))
    elif vtype == "image":
        if not visual.get("path"):
            raise BuildError("slide %r: visual.image needs a path" % sid)
        # Insert at native size first so the true aspect ratio is known, then
        # scale to fit inside (width, height) without stretching past either
        # bound (passing only height= here let a wide image overflow the
        # content area with no width cap at all), and center it in the box.
        pic = slide.shapes.add_picture(visual["path"], left, top)
        scale = min(width / pic.width, height / pic.height, 1)
        pic.width = int(pic.width * scale)
        pic.height = int(pic.height * scale)
        pic.left = left + (width - pic.width) // 2
        pic.top = top + (height - pic.height) // 2
        if visual.get("alt"):
            PH.set_alt_text(pic, visual["alt"])
    elif vtype == "frame":
        raise BuildError("slide %r: frame visuals (%s) are a Phase 2 feature" % (sid, visual.get("kind")))
    else:
        raise BuildError("slide %r: unknown visual.type %r" % (sid, vtype))


REQUIRED_SLIDE_FIELDS = {
    "Two Content": ("left", "right"),
    "Comparison": ("left_head", "right_head", "left", "right"),
    "Three Column": ("columns",),
    "Three Stat": ("stats",),
    "Big Statement with Illustration": ("statement",),
}

# Layouts with no placeholder content of their own in the CONTENT_* rectangle
# a `visual` is placed into -- the only ones a slide's `visual` can safely
# share the slide with. Every other layout (Two Content, Comparison, Three
# Column, Three Stat, Title and Content, Big Statement, ...) fills that same
# rectangle with its own placeholders, so a visual there would land almost
# entirely on top of existing content.
VISUAL_COMPATIBLE_LAYOUTS = ("Title Only", "Title Only (Centered)", "Blank")


def _add_content_slide(prs, sd):
    layout = sd.get("layout")
    sid = sd.get("id", "?")
    if layout not in PH.LAYOUTS:
        raise BuildError("slide %r: unknown layout %r" % (sid, layout))
    missing = [f for f in REQUIRED_SLIDE_FIELDS.get(layout, ()) if not sd.get(f)]
    if missing:
        raise BuildError("slide %r: layout %r needs %s" % (sid, layout, ", ".join(missing)))
    if sd.get("visual") and layout not in VISUAL_COMPATIBLE_LAYOUTS:
        raise BuildError(
            "slide %r: layout %r cannot be paired with a visual -- it would be placed on top of "
            "that layout's own placeholders. Use 'Title Only' for a slide that needs a visual, or "
            "drop the visual and let %r's own content (left/right/columns/stats) carry the slide"
            % (sid, layout, layout))

    title = sd.get("title") or ""
    body = sd.get("body")
    body_levels = sd.get("body_levels")

    if layout == "Title and Content":
        slide = PH.add_title_content(prs, title, body or [], levels=body_levels)
    elif layout == "Section Header":
        slide = PH.add_section_header(prs, title, sd.get("subtitle"))
    elif layout in ("Title Only", "Title Only (Centered)"):
        slide = PH.add_title_only(prs, title, layout_name=layout)
    elif layout == "Two Content":
        slide = PH.add_two_content(prs, title, sd["left"], sd["right"],
                                   left_levels=sd.get("left_levels"), right_levels=sd.get("right_levels"))
    elif layout == "Comparison":
        slide = PH.add_comparison(prs, title, sd["left_head"], sd["left"], sd["right_head"], sd["right"])
    elif layout == "Three Column":
        slide = PH.add_three_column(prs, title, sd["columns"], headings=sd.get("headings"))
    elif layout == "Three Stat":
        stats = [(s.get("value"), s.get("label")) for s in sd["stats"]]
        slide = PH.add_three_stat(prs, stats)
    elif layout == "Big Statement with Illustration":
        slide = PH.add_big_statement(prs, sd["statement"], eyebrow=sd.get("eyebrow"), image=sd.get("image"))
    elif layout == "Title Slide for Verticals":
        slide = PH.add_cover(prs, title, sd.get("subtitle"))
    elif layout == "Blank":
        slide = prs.slides.add_slide(PH.get_layout(prs, "Blank"))
        if title:
            PH.add_textbox(slide, title, PH.CONTENT_LEFT, PH.CONTENT_TOP, PH.CONTENT_WIDTH, Inches(0.6),
                           size_pt=24, bold=True, color=PH.CHROME_TEXT)
    else:
        raise BuildError("slide %r: layout %r is not yet supported by the builder" % (sid, layout))

    # Body text + visual placement for layouts with no dedicated body
    # placeholder of their own (SPEC.md pairs "Title Only" with both a
    # bullet body and a visual): a fixed-height text box up top, the visual
    # taking the remaining content area beneath it.
    top, height = PH.CONTENT_TOP, PH.CONTENT_HEIGHT
    if layout in VISUAL_COMPATIBLE_LAYOUTS and body:
        tb = slide.shapes.add_textbox(PH.CONTENT_LEFT, top, PH.CONTENT_WIDTH, BODY_BLOCK_HEIGHT)
        PH.fill_text_frame(tb, body, levels=body_levels)
        top = top + BODY_BLOCK_HEIGHT + Inches(0.15)
        height = PH.CONTENT_HEIGHT - BODY_BLOCK_HEIGHT - Inches(0.15)
    visual = sd.get("visual")
    if visual:
        _add_visual(slide, visual, PH.CONTENT_LEFT, top, PH.CONTENT_WIDTH, height, sid)
    if sd.get("source"):
        PH.add_source_line(slide, sd["source"])
    if sd.get("notes"):
        slide.notes_slide.notes_text_frame.text = sd["notes"]
    return slide


def build_deck(spec, out, recipe):
    meta = spec.get("meta") or {}
    weave = spec.get("weave") or {}
    slides_spec = spec.get("slides") or []
    internal = bool(meta.get("internal"))
    client = meta.get("client")
    title = meta.get("title") or ""
    if not title:
        raise BuildError("meta.title is required")

    _resolve_slide_frameworks(slides_spec, recipe)
    recipe_results = _check_deck_recipe(recipe, slides_spec, meta) if recipe else []

    base = resolve_base_path(weave.get("base"))
    prs = pptx.Presentation(base)
    subtitle = meta.get("subtitle")
    if not subtitle:
        parts = [p for p in ((client if not internal else None), meta.get("date")) if p]
        subtitle = " / ".join(parts) or None
    cover_images = meta.get("cover_images")
    if cover_images is None:
        cover_images = NDeck.pick_cover_images(2)
    try:
        NDeck.build_shell(prs, title, subtitle=subtitle, client=client, internal=internal,
                          cover_images=cover_images, log=_SILENT)
    except SystemExit as exc:
        raise BuildError(str(exc))

    seen_ids = set()
    for sd in slides_spec:
        sid = sd.get("id")
        if not sid:
            raise BuildError("every slide needs a non-empty id")
        if sid in seen_ids:
            raise BuildError("slide id %r is used more than once" % sid)
        seen_ids.add(sid)
        try:
            slide = _add_content_slide(prs, sd)
        except (ValueError, KeyError) as exc:
            raise BuildError("slide %r: %s" % (sid, exc))
        spec_embed.mark_slide(slide, sid)

    prs.save(out)
    final_prs = NDeck.finalize_deck(out, title=title, prune_empty=True, log=_SILENT)
    spec_embed.embed_spec(out, spec)

    args_ns = SimpleNamespace(internal=internal, client=client)
    results = PH.validate_deck(final_prs, out, args_ns)
    results.extend(recipe_results)
    flags = []
    for sd in slides_spec:
        visual = sd.get("visual") or {}
        if visual.get("type") in ("chart", "table") and not sd.get("source"):
            flags.append("slide %r has a %s but no source line" % (sd.get("id", "?"), visual.get("type")))
    return _summarize(results), flags


# ---------------------------------------------------------------------------
# sow kind
# ---------------------------------------------------------------------------
def _fill_signature_line(p_el, left_text, right_text):
    """The Signatures 'Name' blank line is one paragraph with two
    tab-separated runs of underscores (client side, Netwoven side); only
    those two runs' text is touched, so the tab layout survives."""
    blanks = [r.find(W + "t") for r in p_el.findall(W + "r")
             if r.find(W + "t") is not None and re.fullmatch(r"_+\s*", r.find(W + "t").text or "")]
    if len(blanks) < 2:
        return False
    if left_text:
        blanks[0].text = left_text
    if right_text:
        blanks[-1].text = right_text
    return True


def _fill_signatures(doc, body, sig, log):
    children = list(body)
    h_idx = NS.heading_index(children, 1, "Signatures")
    if h_idx is None:
        log.append("warning: sections.signatures: 'Signatures' heading not found; left blank")
        return
    block_p = None
    for el in children[h_idx + 1:]:
        if H.is_p(el) and H.p_style(el) == "Heading1":
            break
        if H.is_p(el) and H.p_style(el) == "Block":
            block_p = el
            break
    if block_p is None:
        log.append("warning: sections.signatures: could not find the signature blank line; left blank")
        return
    left = sig.get("client_name") or ""
    if left and sig.get("client_title"):
        left = "%s, %s" % (left, sig["client_title"])
    right = sig.get("netwoven_name") or ""
    if right and sig.get("netwoven_title"):
        right = "%s, %s" % (right, sig["netwoven_title"])
    if _fill_signature_line(block_p, left, right):
        log.append("signatures: filled the Name line (client=%r, netwoven=%r)" % (left, right))
    else:
        log.append("warning: sections.signatures: signature blank line did not match the expected pattern; left blank")


def _append_glossary(doc, body, pairs, log):
    children = list(body)
    h_idx = NS.heading_index(children, 1, "Glossary")
    if h_idx is None:
        log.append("warning: sections.glossary_add: 'Glossary' heading not found; terms not written")
        return
    _start, end = NS.range_after_heading(children, h_idx, stop_levels=(1,))
    anchor = children[end - 1]

    def build():
        for pair in pairs:
            term, definition = (list(pair) + ["", ""])[:2]
            if definition is None:
                definition = ""  # an explicit null is a gap (SPEC.md); never write the literal "None"
            p = doc.add_paragraph(style="List Paragraph")
            r1 = p.add_run(str(term))
            r1.bold = True
            if definition:
                p.add_run(" – " + str(definition))  # en dash, matching the template's own glossary format

    new_els = _capture_new_elements(body, build)
    cur = anchor
    for el in new_els:
        cur.addnext(el)
        cur = el
    log.append("added %d glossary term(s)" % len(pairs))


def _write_blocks_at(doc, anchor, blocks, log):
    """Write `blocks` in order right after `anchor`. Each block's content is
    bookmarked immediately, and `cur` MUST then jump to the bookmarkEnd (not
    stay on the block's last content element): mark_block() places
    bookmarkEnd via `last_el.addnext(end)`, and a later `cur.addnext(...)`
    against that same last_el would insert the NEXT block's content between
    the two -- inside this block's own bookmark pair -- instead of after it."""
    body = doc.element.body
    cur = anchor
    for block in blocks:
        new_els = _write_block(doc, block, "Heading %d")
        for el in new_els:
            cur.addnext(el)
            cur = el
        _start, end = spec_embed.mark_block(body, new_els[0], new_els[-1], block["id"])
        cur = end
    return cur


def _fill_known_slot(doc, body, level, heading_text, blocks, name, log):
    """Insert right after the heading, consuming the single empty paragraph
    strip_highlighted_instructions() always leaves there (e.g. the Executive
    Summary fill-in sentence)."""
    children = list(body)
    h_idx = NS.heading_index(children, level, heading_text)
    if h_idx is None:
        log.append("warning: sections.%s: %r heading not found; content not written" % (name, heading_text))
        return
    slot = children[h_idx + 1] if h_idx + 1 < len(children) and H.is_empty_p(children[h_idx + 1]) else None
    _write_blocks_at(doc, children[h_idx], blocks, log)
    if slot is not None:
        body.remove(slot)


def _append_section_end(doc, body, level, heading_text, blocks, name, log, stop_levels):
    """Append at the end of the section (after any spec-generated content
    such as the milestone outline), consuming one trailing empty paragraph
    (the lone Body Text slot new_sow_docx leaves for an unfilled section)."""
    children = list(body)
    h_idx = NS.heading_index(children, level, heading_text)
    if h_idx is None:
        log.append("warning: sections.%s: %r heading not found; content not written" % (name, heading_text))
        return
    _start, end = NS.range_after_heading(children, h_idx, stop_levels=stop_levels)
    slot, anchor_idx = None, end - 1
    if end - 1 > h_idx and H.is_empty_p(children[end - 1]):
        slot, anchor_idx = children[end - 1], end - 2
    _write_blocks_at(doc, children[anchor_idx], blocks, log)
    if slot is not None:
        body.remove(slot)


def _write_sow_sections(doc, sections, log):
    body = doc.element.body
    if sections.get("executive_summary"):
        _fill_known_slot(doc, body, 1, "Executive Summary", sections["executive_summary"], "executive_summary", log)
    if sections.get("scope"):
        _append_section_end(doc, body, 2, "Scope of Work", sections["scope"], "scope", log, stop_levels=(1, 2))
    if sections.get("out_of_scope"):
        _append_section_end(doc, body, 2, "Out of Scope", sections["out_of_scope"], "out_of_scope", log,
                            stop_levels=(1, 2))
    if sections.get("assumptions_add"):
        _append_section_end(doc, body, 1, "Assumptions", sections["assumptions_add"], "assumptions_add", log,
                            stop_levels=(1,))
    if sections.get("roles_notes"):
        _append_section_end(doc, body, 1, "Project Operations", sections["roles_notes"], "roles_notes", log,
                            stop_levels=(1,))
    if sections.get("glossary_add"):
        _append_glossary(doc, body, sections["glossary_add"], log)
    sig = sections.get("signatures") or {}
    if any(sig.get(k) for k in ("client_name", "client_title", "netwoven_name", "netwoven_title")):
        _fill_signatures(doc, body, sig, log)


def build_sow_kind(spec, out):
    meta = spec.get("meta") or {}
    weave = spec.get("weave") or {}
    sow_spec = dict(spec.get("sow") or {})
    variant = sow_spec.get("variant")
    if variant not in ("milestone", "tm"):
        raise BuildError("sow.variant must be 'milestone' or 'tm' (got %r)" % variant)
    sections = sow_spec.pop("sections", None) or {}
    title = meta.get("title") or ""
    client = meta.get("client")
    date = meta.get("date")
    if not title:
        raise BuildError("meta.title is required")
    if not client:
        raise BuildError("meta.client is required for a SOW")
    if not date:
        raise BuildError("meta.date is required for a SOW")

    allow_dashes = bool(meta.get("allow_dashes"))
    seen_ids, seen_safe = {}, {}
    for name in ("executive_summary", "scope", "out_of_scope", "assumptions_add", "roles_notes"):
        validate_section_blocks(sections.get(name) or [], name, allow_dashes, seen_ids, seen_safe)

    base = resolve_base_path(weave.get("base"))

    def before_save(doc, log):
        _write_sow_sections(doc, sections, log)

    try:
        res = NS.build_sow(
            variant, base, out, title, client, date, spec=sow_spec,
            client_logo=sow_spec.get("client_logo"),
            keep_products=bool(sow_spec.get("keep_products")),
            keep_appendix_timeline_table=bool(sow_spec.get("keep_appendix_timeline_table")),
            expenses_cap=sow_spec.get("expenses_cap"),
            before_save=before_save,
        )
    except NS.SowError as exc:
        raise BuildError(str(exc))

    spec_embed.embed_spec(out, spec)
    results = V.check_docx(out, "sow", False, client)
    for problem in res["problems"]:
        results.append({"check": "post-save assertion", "passed": False, "evidence": problem, "hard": True})
    flags = [ln[len("warning: "):] for ln in res["log"] if ln.startswith("warning: ")]
    return _summarize(results), flags


# ---------------------------------------------------------------------------
def _summarize(results):
    hard_fails = sum(1 for r in results if r["hard"] and not r["passed"])
    warnings = sum(1 for r in results if not r["hard"] and not r["passed"])
    return {"hard_fails": hard_fails, "warnings": warnings, "results": results}


def _build(args):
    spec = load_spec(args.spec)
    weave = spec.setdefault("weave", {})
    spec.setdefault("meta", {})
    spec.setdefault("blocks", [])
    spec.setdefault("slides", [])
    spec.setdefault("sow", None)
    kind = weave.get("kind")
    if kind not in ("doc", "deck", "sow"):
        raise BuildError("spec.weave.kind must be 'doc', 'deck' or 'sow' (got %r)" % kind)
    _ensure_kind_imports(kind)

    if args.client_slug:
        spec["meta"]["client_slug"] = args.client_slug
    if args.recipe:
        weave["recipe"] = args.recipe

    recipe = load_recipe(spec, kind)
    _apply_title_pattern(spec, recipe)

    if args.out:
        out = os.path.abspath(args.out)
        revision = None  # resolved below via resolve_revision, against this exact explicit path
    else:
        out_name, revision = _default_out_path_and_revision(spec, kind)
        out = os.path.abspath(out_name)

    if os.path.exists(out) and not args.replace:
        raise BuildError("%s already exists; pass --replace to build a new revision over it" % out)

    # For an explicit --out, resolve_revision's own existence check against
    # that literal path is authoritative. For the auto-generated name, the
    # revision was already resolved above FROM THE SAME SCAN that picked the
    # filename's own v<revision> suffix, and must not be re-derived here: by
    # construction `out` never already exists in that branch (each build gets
    # a fresh, not-yet-used filename), so resolve_revision(out, spec) would
    # find nothing there and silently fall back to the spec's stale incoming
    # revision -- reintroducing the filename/embedded-revision mismatch this
    # fixes.
    weave["revision"] = resolve_revision(out, spec) if revision is None else revision
    weave.setdefault("spec_version", "2.0")
    weave.setdefault("generator", WV.plugin_version())
    weave["built_at"] = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    if kind == "doc":
        validation, flags = build_doc(spec, out, recipe)
    elif kind == "deck":
        validation, flags = build_deck(spec, out, recipe)
    else:
        validation, flags = build_sow_kind(spec, out)

    return {
        "ok": True,
        "out": out,
        "kind": kind,
        "revision": weave["revision"],
        "recipe": weave.get("recipe"),
        "validation": validation,
        "flags": flags,
    }


def _fail(error, hint):
    """Always JSON, regardless of --json (matches weave.py's own emit_error)."""
    print(json.dumps({"ok": False, "error": error, "hint": hint}))
    return 1


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--replace", action="store_true")
    ap.add_argument("--client-slug", default=None)
    ap.add_argument("--recipe", default=None)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    try:
        result = _build(args)
    except BuildError as exc:
        return _fail(str(exc), "Check the spec against recipes/SPEC.md.")
    except SystemExit as exc:
        return _fail(str(exc.code), "Check the spec against recipes/SPEC.md.")
    except Exception as exc:  # noqa: BLE001 - top-level guard: always emit the JSON contract
        traceback.print_exc()
        return _fail("unexpected error: %s" % exc, "Run again without --json to see the traceback on stderr.")

    if args.json:
        print(json.dumps(result))
    else:
        v = result["validation"]
        print("OK: wrote %s (kind=%s, revision=%d, hard_fails=%d, warnings=%d)" % (
            result["out"], result["kind"], result["revision"], v["hard_fails"], v["warnings"]))
        for f in result["flags"]:
            print("  ! " + f)
    return 0


if __name__ == "__main__":
    sys.exit(main())
