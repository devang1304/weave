#!/usr/bin/env python3
"""
weave.py -- the one command Weave skills invoke (see recipes/SPEC.md).

  weave.py [--json] doctor          [--quiet] [--install]
  weave.py [--json] build           --spec FILE --out FILE [...]
  weave.py [--json] spec-from-file  FILE [--out spec.json]
  weave.py [--json] review          FILE [--quick] [--rubric ...] [--apply] [--report out.md]
  weave.py [--json] validate        FILE [--type doc|sow|deck] [--internal] [--client NAME] [--render]
  weave.py [--json] read            PATH [--out md] [--tables-json [FILE]] [--max-chars N]
  weave.py [--json] client          list | show SLUG | set SLUG key=value ... | delete SLUG --yes
  weave.py [--json] label           FILE --show | --expect NAME | --apply NAME
  weave.py [--json] telemetry       event ... | hook | flush | status | login-start | login-poll | logout | configure ...
  weave.py [--json] memory          show | promote | forget KEY | status
  weave.py --version

Every subcommand prints exactly one JSON object when --json is given.
Failures print {"ok": false, "error": "...", "hint": "..."} and exit 1.

Only the dispatcher logic lives here (stdlib).  Subcommands that need
python-docx / python-pptx are run after an import check; when the check fails
this script re-executes itself with the interpreter recorded in
${CLAUDE_PLUGIN_DATA}/env.json (written by `doctor --install`).

After build, review, validate, read and spec-from-file the dispatcher records
one usage event (telemetry.py, GOVERNANCE.md Sec 5): kind, recipe, outcome,
duration, never content. WEAVE_TELEMETRY=0 turns that off.
"""
from __future__ import annotations

import json
import os
import re
import stat
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

SCRIPT_DIR = Path(__file__).resolve().parent

# subcommand -> script beside this file
DISPATCH = {
    "build": "build_deliverable.py",
    "spec-from-file": "spec_from_file.py",
    "review": "review_deliverable.py",
    "validate": "validate_deliverable.py",
    "read": "read_source.py",
    "label": "purview_label.py",
    "telemetry": "telemetry.py",
    "memory": "memory.py",
}
NO_DEPS = {"doctor", "client", "telemetry", "memory"}  # stdlib-only targets: never need python-docx/pptx/lxml
SUBCOMMANDS_HINT = "Use one of: doctor, build, spec-from-file, review, validate, read, client, label, telemetry, memory"
PLAIN_TEXT_SUFFIXES = {".txt", ".md", ".vtt", ".markdown"}
# read_source.py's own optional deps (pypdf / openpyxl) cover .pdf / .xlsx --
# never python-docx / python-pptx / lxml.  Only .docx / .pptx (and their
# legacy .doc / .ppt cousins) genuinely need the trio gated by ensure_deps()
# below, so a missing openpyxl/pypdf can surface read_source.py's own
# specific ImportError message instead of being pre-empted by this blanket
# gate's generic one.
READ_NO_OFFICE_DEPS_SUFFIXES = PLAIN_TEXT_SUFFIXES | {".pdf", ".xlsx"}

CLIENT_KEYS = ("legal_name", "short_name", "logo_path", "standard_assumptions",
               "ai_permitted", "default_label", "products_in_scope", "notes",
               "author_name", "author_role")
CLIENT_LIST_KEYS = {"standard_assumptions", "products_in_scope"}
CLIENT_BOOL_KEYS = {"ai_permitted"}
SETUP_HINT = "Run /weave:setup to install what Weave needs."


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
def data_dir() -> Path:
    raw = os.environ.get("CLAUDE_PLUGIN_DATA")
    d = Path(raw).expanduser() if raw else Path.home() / ".weave"
    try:
        d.mkdir(parents=True, exist_ok=True)
        os.chmod(d, 0o700)
    except OSError:
        pass
    return d


def plugin_version() -> str:
    p = SCRIPT_DIR
    for _ in range(5):  # the script itself plus up to 4 parents
        cand = p / ".claude-plugin" / "plugin.json"
        if cand.exists():
            try:
                v = json.loads(cand.read_text(encoding="utf-8")).get("version")
                if v:
                    return "weave %s" % v
            except (OSError, ValueError):
                break
        if p.parent == p:
            break
        p = p.parent
    return "weave (dev)"


def emit_error(error: str, hint: str, code: int = 1) -> int:
    print(json.dumps({"ok": False, "error": error, "hint": hint}))
    return code


def slugify(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", name.strip().lower()).strip("-")
    return re.sub(r"-{2,}", "-", s)


def _last_json_object(text: str) -> Optional[dict]:
    """Parse the whole stdout as one JSON object, or fall back to the last line."""
    text = text.strip()
    if not text:
        return None
    try:
        obj = json.loads(text)
        return obj if isinstance(obj, dict) else {"ok": True, "result": obj}
    except ValueError:
        pass
    for line in reversed(text.splitlines()):
        line = line.strip()
        if line.startswith("{"):
            try:
                obj = json.loads(line)
                return obj if isinstance(obj, dict) else None
            except ValueError:
                continue
    return None


# ---------------------------------------------------------------------------
# dependency check / re-exec
# ---------------------------------------------------------------------------
def deps_ok() -> bool:
    try:
        import docx  # noqa: F401
        import lxml  # noqa: F401
        import pptx  # noqa: F401
    except ImportError:
        return False
    return True


def ensure_deps() -> Optional[int]:
    """Return None when imports work (possibly after re-exec), else an exit code."""
    if deps_ok():
        return None
    if os.environ.get("WEAVE_REEXEC") != "1":
        env_path = data_dir() / "env.json"
        try:
            env = json.loads(env_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            env = {}
        python = env.get("python")
        # abspath, not realpath: a venv's python is normally a symlink to the
        # system interpreter, so realpath() would call them "the same
        # interpreter" and skip the exec -- but venv activation depends on
        # the path used to launch python, not the binary it resolves to, so
        # only the exact recorded path actually picks up the venv's packages.
        if python and os.path.exists(python) and os.path.abspath(python) != os.path.abspath(sys.executable):
            os.environ["WEAVE_REEXEC"] = "1"
            try:
                sys.stdout.flush()
                os.execv(python, [python, os.path.abspath(__file__)] + sys.argv[1:])
            except OSError:
                pass  # fall through to the error below
    return emit_error("Weave's Python packages (python-docx, python-pptx, lxml) are not installed.", SETUP_HINT)


# ---------------------------------------------------------------------------
# subprocess dispatch
# ---------------------------------------------------------------------------
def run_child(script: str, args: List[str], as_json: bool, result_out: Optional[dict] = None) -> int:
    """Run a sibling script. When `result_out` is given and the child printed
    a JSON object, it is stored under result_out["result"] for the caller."""
    target = SCRIPT_DIR / script
    if not target.exists():
        return emit_error("%s is not available (%s not found)." % (script, target),
                          "This part of Weave is not installed yet")
    cmd = [sys.executable, str(target)] + args
    if as_json and "--json" not in args:
        cmd.append("--json")
    if not as_json:
        try:
            return subprocess.call(cmd)
        except OSError as exc:
            return emit_error("Could not start %s: %s" % (script, exc), SETUP_HINT)
    try:
        cp = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    except OSError as exc:
        return emit_error("Could not start %s: %s" % (script, exc), SETUP_HINT)
    obj = _last_json_object(cp.stdout)
    if obj is None:
        tail = (cp.stderr or cp.stdout or "").strip().splitlines()[-6:]
        # Propagate the child's real exit code so weave.py's own process exit
        # status matches what actually failed; a nonfailing code here would
        # be a contradiction (obj is None only because the run failed to
        # produce parseable JSON), so fall back to 1 defensively.
        code = cp.returncode if cp.returncode else 1
        return emit_error("%s exited with code %d without a JSON result." % (script, cp.returncode),
                          " | ".join(tail) or "Run the command again without --json to see the full output.",
                          code=code)
    if result_out is not None:
        result_out["result"] = obj
    print(json.dumps(obj, indent=2))
    return cp.returncode


def dispatch(sub: str, args: List[str], as_json: bool) -> int:
    """run_child plus the passive usage event for the subcommands that do
    real work. Telemetry can never change the exit code or the output."""
    started = time.monotonic()
    captured: dict = {}
    rc = run_child(DISPATCH[sub], args, as_json, captured)
    try:
        sys.path.insert(0, str(SCRIPT_DIR))
        import telemetry  # noqa: WPS433 - sibling stdlib script; absent on some packaging surfaces

        if sub in telemetry.PASSIVE_EVENT_SUBCOMMANDS:
            telemetry.record(telemetry.passive_fields(sub, args, captured.get("result"), rc, time.monotonic() - started))
    except Exception:  # noqa: BLE001 - telemetry is a side channel, never a failure mode
        pass
    return rc


# ---------------------------------------------------------------------------
# client profiles
# ---------------------------------------------------------------------------
def clients_dir() -> Path:
    d = data_dir() / "clients"
    d.mkdir(parents=True, exist_ok=True)
    try:
        os.chmod(d, 0o700)
    except OSError:
        pass
    return d


def _profile_path(slug: str) -> Path:
    return clients_dir() / ("%s.json" % slug)


def _load_profile(slug: str) -> Optional[dict]:
    p = _profile_path(slug)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _save_profile(slug: str, profile: dict) -> Path:
    p = _profile_path(slug)
    fd = os.open(str(p), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        fh.write(json.dumps(profile, indent=2, ensure_ascii=False) + "\n")
    os.chmod(p, stat.S_IRUSR | stat.S_IWUSR)
    return p


def _parse_value(key: str, raw: str):
    if key in CLIENT_LIST_KEYS:
        raw = raw.strip()
        if raw == "":
            return []
        if raw.startswith("["):
            try:
                val = json.loads(raw)
            except ValueError as exc:
                raise ValueError("%s must be a JSON list or a ';'-separated string (%s)" % (key, exc))
            if not isinstance(val, list):
                raise ValueError("%s must be a list" % key)
            return [str(x) for x in val]
        return [part.strip() for part in raw.split(";") if part.strip()]
    if key in CLIENT_BOOL_KEYS:
        low = raw.strip().lower()
        if low in ("true", "yes", "1", "on"):
            return True
        if low in ("false", "no", "0", "off"):
            return False
        raise ValueError("%s must be true or false" % key)
    return raw


def _new_profile(slug: str, display: str) -> dict:
    return {
        "slug": slug,
        "legal_name": display,
        "short_name": display,
        "logo_path": None,
        "standard_assumptions": [],
        "ai_permitted": True,
        "default_label": None,
        "products_in_scope": [],
        "notes": "",
        "author_name": None,
        "author_role": None,
    }


def cmd_client(args: List[str], as_json: bool) -> int:
    if not args:
        return emit_error("client needs an action: list | show SLUG | set SLUG key=value ... | delete SLUG --yes",
                          "Example: weave.py client set \"Contoso Ltd\" short_name=Contoso")
    action, rest = args[0], args[1:]

    if action == "list":
        rows = []
        for p in sorted(clients_dir().glob("*.json")):
            prof = _load_profile(p.stem) or {}
            rows.append({"slug": p.stem, "legal_name": prof.get("legal_name"), "short_name": prof.get("short_name")})
        if as_json:
            print(json.dumps({"ok": True, "clients": rows, "dir": str(clients_dir())}, indent=2))
        elif rows:
            for r in rows:
                print("%s\t%s\t%s" % (r["slug"], r.get("legal_name") or "", r.get("short_name") or ""))
        else:
            print("No client profiles yet (%s)." % clients_dir())
        return 0

    if not rest:
        return emit_error("client %s needs a client name or slug." % action, "Example: weave.py client %s contoso" % action)
    display = rest[0]
    slug = slugify(display)
    if not slug:
        return emit_error("'%s' does not make a usable client slug." % display, "Use letters or digits in the name.")

    if action == "show":
        prof = _load_profile(slug)
        if prof is None:
            return emit_error("No client profile for '%s'." % slug,
                              "Create one with: weave.py client set %s legal_name=\"...\"" % slug)
        if as_json:
            print(json.dumps({"ok": True, "slug": slug, "profile": prof}, indent=2))
        else:
            for k in ("slug",) + CLIENT_KEYS:
                print("%s: %s" % (k, json.dumps(prof.get(k), ensure_ascii=False)))
        return 0

    if action == "set":
        prof = _load_profile(slug)
        created = prof is None
        if prof is None:
            prof = _new_profile(slug, display)
        for kv in rest[1:]:
            if "=" not in kv:
                return emit_error("'%s' is not key=value." % kv, "Allowed keys: %s" % ", ".join(CLIENT_KEYS))
            key, raw = kv.split("=", 1)
            key = key.strip()
            if key not in CLIENT_KEYS:
                return emit_error("Unknown client key '%s'." % key, "Allowed keys: %s" % ", ".join(CLIENT_KEYS))
            try:
                prof[key] = _parse_value(key, raw)
            except ValueError as exc:
                return emit_error(str(exc), "Lists: a;b;c or [\"a\",\"b\"]. Booleans: true/false.")
        prof["slug"] = slug
        prof["updated_at"] = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        path = _save_profile(slug, prof)
        if as_json:
            print(json.dumps({"ok": True, "action": "created" if created else "updated", "slug": slug,
                              "path": str(path), "profile": prof}, indent=2))
        else:
            print("%s client profile %s (%s)" % ("Created" if created else "Updated", slug, path))
        return 0

    if action == "delete":
        if "--yes" not in rest[1:]:
            return emit_error("Deleting a client profile needs --yes.",
                              "weave.py client delete %s --yes" % slug)
        p = _profile_path(slug)
        if not p.exists():
            return emit_error("No client profile for '%s'." % slug, "weave.py client list shows what exists.")
        p.unlink()
        if as_json:
            print(json.dumps({"ok": True, "action": "deleted", "slug": slug}))
        else:
            print("Deleted client profile %s" % slug)
        return 0

    return emit_error("Unknown client action '%s'." % action, "Use list | show SLUG | set SLUG key=value ... | delete SLUG --yes")


# ---------------------------------------------------------------------------
def usage() -> str:
    return (__doc__ or "").strip().split("\n\n")[1]


def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    if argv and argv[0] in ("--version", "-V"):
        print(plugin_version())
        return 0
    if not argv or argv[0] in ("-h", "--help"):
        print(usage())
        return 0 if argv else 1

    as_json = "--json" in argv
    # Global --json may sit before the subcommand; strip it there and rely on
    # `as_json` (children get --json appended when they need it).
    while argv and argv[0] == "--json":
        argv.pop(0)
    if not argv:
        return emit_error("No subcommand given.", SUBCOMMANDS_HINT)
    sub, rest = argv[0], argv[1:]

    if sub == "doctor":
        return run_child("doctor.py", rest, as_json)
    if sub == "client":
        return cmd_client([a for a in rest if a != "--json"], as_json)
    if sub not in DISPATCH:
        return emit_error("Unknown subcommand '%s'." % sub, SUBCOMMANDS_HINT)

    needs_deps = sub not in NO_DEPS
    if needs_deps and sub == "read":
        paths = [a for a in rest if not a.startswith("-")]
        if paths and Path(paths[0]).suffix.lower() in READ_NO_OFFICE_DEPS_SUFFIXES:
            needs_deps = False
    if needs_deps:
        rc = ensure_deps()
        if rc is not None:
            return rc
    return dispatch(sub, rest, as_json)


if __name__ == "__main__":
    sys.exit(main())
