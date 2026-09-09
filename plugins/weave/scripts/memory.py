#!/usr/bin/env python3
"""
memory.py -- Weave's local preference memory. Standard library only.

Weave learns from what users change by hand. Whenever a skill reads a
Weave-built file back (`weave.py spec-from-file`), spec_from_file.py compares
the stored plan with the live document and hands the differences to
observe(): a heading the user deleted, a slide they removed, a word they
replaced. Each difference becomes a pattern with a stable key. When the same
key has been seen `promote_after` times it is promoted to a preference, and
the build skills read the preferences before writing a spec.

  memory.py show     [--json]          promoted preferences, one per line
  memory.py observe  --before FILE --after FILE --drift FILE [--json]
  memory.py promote  [--json]
  memory.py forget   KEY [--json]      drop a preference and never re-promote it
  memory.py status   [--json]

Files under <data>/memory/ (mode 0700): patterns.jsonl (one observation per
line), preferences.json (promoted keys, their text and counts, and forgotten
keys), memory.json (settings that override DEFAULTS).

Everything here stays on the device: memory is never sent with telemetry,
and an observation stores at most `max_term_words` words of the user's own
edit, or one heading or slide title. WEAVE_MEMORY=0 turns observation off.
"""
from __future__ import annotations

import argparse
import contextlib
import datetime as _dt
import difflib
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional

DEFAULTS = {
    # The same hand edit on three separate deliverables is a habit, not a
    # one-off; one or two are too easily noise from a single engagement.
    "promote_after": 3,
    # A "term" swap is a few words ("utilize" -> "use"). Longer spans are
    # rewrites: recorded for the count, never promoted as a preference.
    "max_term_words": 3,
}
PATTERNS_FILE = "patterns.jsonl"
PREFERENCES_FILE = "preferences.json"
SETTINGS_FILE = "memory.json"
# _locked()'s tunables: how long a healthy caller waits, how old an
# unreleased lock must be before it's presumed abandoned by a crashed
# holder, and the poll interval while waiting.
LOCK_TIMEOUT_S = 5.0
LOCK_STALE_AFTER_S = 60.0
LOCK_POLL_INTERVAL_S = 0.05
SOW_SECTIONS = ("executive_summary", "scope", "out_of_scope", "assumptions_add", "roles_notes")
# Observation kinds that are precise enough to become a preference sentence.
PROMOTABLE = ("heading_removed", "slide_removed", "term", "phrase_dropped")


def data_dir(create: bool = True) -> Path:
    raw = os.environ.get("CLAUDE_PLUGIN_DATA")
    d = Path(raw).expanduser() if raw else Path.home() / ".weave"
    if create:
        try:
            d.mkdir(parents=True, exist_ok=True)
            os.chmod(d, 0o700)
        except OSError:
            pass
    return d


def memory_dir(create: bool = True) -> Path:
    d = data_dir(create) / "memory"
    if create:
        try:
            d.mkdir(parents=True, exist_ok=True)
            os.chmod(d, 0o700)
        except OSError:
            pass
    return d


def enabled() -> bool:
    return os.environ.get("WEAVE_MEMORY") != "0"


def settings(mdir: Optional[Path] = None) -> dict:
    mdir = mdir or memory_dir()
    merged = dict(DEFAULTS)
    try:
        raw = json.loads((mdir / SETTINGS_FILE).read_text(encoding="utf-8"))
        merged.update({k: int(raw[k]) for k in DEFAULTS if k in raw})
    except (OSError, ValueError, TypeError):
        pass
    return merged


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip().lower()


# ---------------------------------------------------------------------------
# observing hand edits
# ---------------------------------------------------------------------------
def _text_of(block: dict) -> Optional[str]:
    kind = block.get("type")
    if kind in ("heading", "para", "callout"):
        if block.get("runs"):
            return " ".join(str(r.get("text", "")) for r in block["runs"] if isinstance(r, dict))
        return block.get("text")
    if kind == "bullets":
        return " ".join(str(i) for i in (block.get("items") or []))
    if kind == "figure":
        return block.get("caption")
    return None  # tables and page breaks carry no prose to compare


def _all_blocks(spec: dict) -> List[dict]:
    blocks = list(spec.get("blocks") or [])
    sow = spec.get("sow")
    if isinstance(sow, dict):
        sections = sow.get("sections") or {}
        for name in SOW_SECTIONS:
            blocks.extend(sections.get(name) or [])
    return [b for b in blocks if isinstance(b, dict)]


def term_change(old: str, new: str, max_term_words: int) -> Optional[dict]:
    """The single small change between two texts, if that is all that changed:
    {"kind": "term", "old", "new"} for a replacement, {"kind": "phrase_dropped",
    "old"} for a deletion. Anything larger, or more than one change, is None."""
    a, b = old.split(), new.split()
    ops = [op for op in difflib.SequenceMatcher(a=a, b=b, autojunk=False).get_opcodes() if op[0] != "equal"]
    if len(ops) != 1:
        return None
    tag, i1, i2, j1, j2 = ops[0]
    if tag == "replace" and (i2 - i1) <= max_term_words and (j2 - j1) <= max_term_words:
        return {"kind": "term", "old": " ".join(a[i1:i2]), "new": " ".join(b[j1:j2])}
    if tag == "delete" and (i2 - i1) <= max_term_words:
        return {"kind": "phrase_dropped", "old": " ".join(a[i1:i2])}
    return None


def _observation(kind: str, spec: dict, **detail) -> dict:
    weave = spec.get("weave") or {}
    parts = [kind] + [_norm(str(detail[k])) for k in ("text", "old", "new") if k in detail]
    return {"ts": _now(), "kind": kind, "deliverable": weave.get("kind"), "recipe": weave.get("recipe"),
            "key": "|".join(parts), "detail": detail, "promotable": kind in PROMOTABLE}


def observe(stored: dict, reconciled: dict, drift: dict, max_term_words: Optional[int] = None) -> List[dict]:
    """Observations from one spec-from-file reconcile: `stored` is the plan
    embedded in the file, `reconciled` the plan after hand edits were
    folded in, `drift` spec_from_file.py's drift record."""
    max_words = settings()["max_term_words"] if max_term_words is None else max_term_words
    out: List[dict] = []
    if (stored.get("weave") or {}).get("kind") == "deck":
        after = {s.get("id"): s for s in (reconciled.get("slides") or []) if isinstance(s, dict)}
        for slide in stored.get("slides") or []:
            live = after.get(slide.get("id"))
            if live is None:
                if slide.get("title"):
                    out.append(_observation("slide_removed", stored, text=slide["title"]))
            elif live.get("edited") and slide.get("title") and live.get("title") and slide["title"] != live["title"]:
                change = term_change(slide["title"], live["title"], max_words)
                out.append(_observation(change["kind"], stored, **{k: v for k, v in change.items() if k != "kind"})
                           if change else _observation("edited", stored, text="slide"))
        return out
    after = {b.get("id"): b for b in _all_blocks(reconciled)}
    for block in _all_blocks(stored):
        live = after.get(block.get("id"))
        if live is None:
            if block.get("type") == "heading" and block.get("text"):
                out.append(_observation("heading_removed", stored, text=block["text"]))
            continue
        if not live.get("edited"):
            continue
        old, new = _text_of(block), _text_of(live)
        if not old or not new or old == new:
            continue
        change = term_change(old, new, max_words)
        out.append(_observation(change["kind"], stored, **{k: v for k, v in change.items() if k != "kind"})
                   if change else _observation("edited", stored, text=block.get("type") or "block"))
    return out


# ---------------------------------------------------------------------------
# storing, promoting, showing
# ---------------------------------------------------------------------------
def record(observations: Iterable[dict], mdir: Optional[Path] = None) -> int:
    if not enabled():
        return 0
    mdir = mdir or memory_dir()
    n = 0
    with open(mdir / PATTERNS_FILE, "a", encoding="utf-8") as fh:
        for obs in observations:
            fh.write(json.dumps(obs, sort_keys=True) + "\n")
            n += 1
    return n


def _read_patterns(mdir: Path) -> List[dict]:
    path = mdir / PATTERNS_FILE
    if not path.exists():
        return []
    out = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            try:
                out.append(json.loads(line))
            except ValueError:
                continue
    return out


def _load_preferences(mdir: Path) -> dict:
    try:
        raw = json.loads((mdir / PREFERENCES_FILE).read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            raw.setdefault("preferences", {})
            raw.setdefault("forgotten", [])
            return raw
    except (OSError, ValueError):
        pass
    return {"preferences": {}, "forgotten": []}


def _save_preferences(mdir: Path, prefs: dict) -> None:
    path = mdir / PREFERENCES_FILE
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(prefs, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    tmp.replace(path)


@contextlib.contextmanager
def _locked(lock_path: Path, timeout_s: float = LOCK_TIMEOUT_S, stale_after_s: float = LOCK_STALE_AFTER_S,
            poll_s: float = LOCK_POLL_INTERVAL_S):
    """Cross-platform advisory lock via exclusive file creation (no fcntl,
    so it also works on Windows), so a concurrent promote() and forget()
    can't race on preferences.json. Waits up to `timeout_s` for a healthy
    holder to finish; a lock older than the separate `stale_after_s` is
    presumed abandoned by a crashed holder and reclaimed on any retry,
    independent of how impatient this particular caller is.

    Reclaiming is stat-then-unlink, which is a TOCTOU by itself: two waiters
    could both see the same stale lock and both reclaim it, or a waiter
    could unlink a lock a healthy holder only just recreated. Re-checking
    the mtime immediately before unlinking narrows that window to the gap
    between the two stat() calls -- not zero, but the only race left
    requires another process's full unlink-then-recreate to land inside it,
    vanishingly unlikely for a local file and proportionate for an advisory
    lock like this one (not a distributed/financial one)."""
    deadline = time.monotonic() + timeout_s
    fd = None
    while fd is None:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                mtime = lock_path.stat().st_mtime
                if time.time() - mtime > stale_after_s and lock_path.stat().st_mtime == mtime:
                    lock_path.unlink()
                    continue
            except OSError:
                pass
            if time.monotonic() >= deadline:
                raise TimeoutError("could not acquire lock: %s" % lock_path)
            time.sleep(poll_s)
    try:
        yield
    finally:
        os.close(fd)
        try:
            lock_path.unlink()
        except OSError:
            pass


def render(kind: str, detail: dict, count: int) -> str:
    """The one sentence a build skill reads. Wording states the habit, never
    an instruction the recipe could conflict with."""
    what = {"deck": "decks", "doc": "documents", "sow": "statements of work"}
    if kind == "heading_removed":
        return "Leave out the section '%s' (removed by hand %d times)." % (detail.get("text"), count)
    if kind == "slide_removed":
        return "Leave out slides like '%s' (removed by hand %d times)." % (detail.get("text"), count)
    if kind == "term":
        return "Write '%s' instead of '%s' (changed by hand %d times)." % (detail.get("new"), detail.get("old"), count)
    if kind == "phrase_dropped":
        return "Drop the phrase '%s' (deleted by hand %d times)." % (detail.get("old"), count)
    return "Edited %s %d times." % (what.get(detail.get("text"), detail.get("text")), count)


def promote(mdir: Optional[Path] = None) -> dict:
    """Promote every promotable key seen at least `promote_after` times that
    the user has not forgotten. Idempotent: counts are recomputed from
    patterns.jsonl each time. A no-op when memory is off: existing
    preferences are returned unchanged, and no directory or file is created,
    matching record()'s own WEAVE_MEMORY=0 gate."""
    if not enabled():
        return _load_preferences(mdir or memory_dir(create=False))
    mdir = mdir or memory_dir()
    with _locked(mdir / "preferences.lock"):
        threshold = settings(mdir)["promote_after"]
        prefs = _load_preferences(mdir)
        groups: Dict[str, List[dict]] = {}
        for obs in _read_patterns(mdir):
            if obs.get("promotable"):
                groups.setdefault(obs["key"], []).append(obs)
        for key, seen in groups.items():
            if len(seen) < threshold or key in prefs["forgotten"]:
                continue
            first = seen[0]
            prefs["preferences"][key] = {
                "text": render(first["kind"], first.get("detail") or {}, len(seen)),
                "count": len(seen),
                "deliverable": first.get("deliverable"),
                "first_seen": min(o["ts"] for o in seen),
                "last_seen": max(o["ts"] for o in seen),
            }
        _save_preferences(mdir, prefs)
        return prefs


def show(mdir: Optional[Path] = None) -> List[dict]:
    prefs = _load_preferences(mdir or memory_dir())["preferences"]
    return [dict(key=k, **v) for k, v in sorted(prefs.items(), key=lambda kv: (-kv[1]["count"], kv[0]))]


def forget(key: str, mdir: Optional[Path] = None) -> bool:
    mdir = mdir or memory_dir()
    with _locked(mdir / "preferences.lock"):
        prefs = _load_preferences(mdir)
        existed = key in prefs["preferences"]
        prefs["preferences"].pop(key, None)
        if key not in prefs["forgotten"]:
            prefs["forgotten"].append(key)
        _save_preferences(mdir, prefs)
        return existed


# ---------------------------------------------------------------------------
def _emit(obj: dict, as_json: bool, human: str) -> int:
    print(json.dumps(obj) if as_json else human)
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, help_text in (("show", "promoted preferences"), ("promote", "promote patterns seen often enough"),
                            ("status", "counts and settings")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--json", action="store_true")
    f = sub.add_parser("forget", help="drop a preference by key and never re-promote it")
    f.add_argument("key")
    f.add_argument("--json", action="store_true")
    o = sub.add_parser("observe", help="record observations from a stored spec, a reconciled spec and a drift record")
    o.add_argument("--before", required=True)
    o.add_argument("--after", required=True)
    o.add_argument("--drift", required=True)
    o.add_argument("--json", action="store_true")
    args = ap.parse_args(argv)

    try:
        if args.cmd == "show":
            prefs = show()
            return _emit({"ok": True, "preferences": prefs}, args.json,
                         "\n".join("- %s" % p["text"] for p in prefs) or "No saved preferences yet.")
        if args.cmd == "promote":
            prefs = promote()
            return _emit({"ok": True, "promoted": len(prefs["preferences"])}, args.json,
                         "%d preference(s) promoted." % len(prefs["preferences"]))
        if args.cmd == "forget":
            existed = forget(args.key)
            return _emit({"ok": True, "forgot": existed}, args.json,
                         "Forgot %s." % args.key if existed else "%s was not a saved preference; it will not be promoted." % args.key)
        if args.cmd == "status":
            mdir = memory_dir()
            obj = {"ok": True, "enabled": enabled(), "patterns": len(_read_patterns(mdir)),
                   "preferences": len(_load_preferences(mdir)["preferences"]), "settings": settings(mdir),
                   "dir": str(mdir)}
            return _emit(obj, args.json, "Memory %s: %d pattern(s), %d preference(s) at %s." % (
                "on" if obj["enabled"] else "off (WEAVE_MEMORY=0)", obj["patterns"], obj["preferences"], mdir))
        with open(args.before, encoding="utf-8") as fh:
            before = json.load(fh)
        with open(args.after, encoding="utf-8") as fh:
            after = json.load(fh)
        with open(args.drift, encoding="utf-8") as fh:
            drift = json.load(fh)
        observations = observe(before, after, drift)
        n = record(observations)
        promote()
        return _emit({"ok": True, "recorded": n, "observations": observations}, args.json,
                     "Recorded %d observation(s)." % n)
    except Exception as exc:  # noqa: BLE001 - always the JSON contract, never a traceback
        print(json.dumps({"ok": False, "error": "unexpected error: %s" % exc, "hint": "Check the file paths."}))
        return 1


if __name__ == "__main__":
    sys.exit(main())
