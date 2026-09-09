#!/usr/bin/env python3
"""
telemetry.py -- Weave's usage-event log (GOVERNANCE.md Sec 5). Standard
library only.

  telemetry.py event     --type TYPE [--skill S] [--recipe R] [--kind doc|deck|sow]
                         [--outcome pass|fail] [--hard-fails N] [--warnings N]
                         [--duration-s F] [--edit-iterations N] [--source-types S[,S...]]
                         [--rating 1-5] [--text T] [--client-slug S] [--json]
  telemetry.py hook      reads one Claude Code hook payload on stdin and records a
                         skill_invoked event when it names a Weave skill; always
                         exits 0 and prints nothing
  telemetry.py flush        [--quiet] [--json]   send queued events to the configured sink
  telemetry.py status       [--json]
  telemetry.py login-start  [--json]             start device-code sign-in; returns immediately
  telemetry.py login-poll   [--json]             one poll attempt for the sign-in login-start began
  telemetry.py logout       [--json]
  telemetry.py configure --endpoint URL --tenant-id ID --client-id ID [--json]

Events are one JSON object per line in <data>/telemetry/events.jsonl,
append-only. user_hash and client_hash are salted SHA-256 digests whose salt
rotates monthly, so an event never carries a name, prompt, path, or title and
never correlates across months. A feedback event carries no client_hash.

Three things write events:
  * weave.py records one event per dispatched subcommand (build, review,
    validate, read, spec-from-file) with kind, recipe, outcome, and duration,
    through passive_fields() below;
  * the plugin's hooks run `telemetry.py hook` when a Weave skill is invoked,
    whether Claude chose it (PostToolUse for the Skill tool) or the user
    typed it (UserPromptExpansion);
  * /weave:feedback records `--type feedback`.

WEAVE_TELEMETRY=0 disables passive events (every type except feedback); a
feedback report still files, since filing it is the point of the command.

Sending: `flush` transmits queued events to the sink telemetry_sink.py
describes (a SharePoint list reached through Microsoft Graph) once an
administrator has configured it and the user has signed in with `login`.
Until then `flush` only reports what is queued. The plugin's SessionStart
hook runs `flush --quiet`, so events leave the device at the start of the
next session, as GOVERNANCE.md promises.

Data dir: $CLAUDE_PLUGIN_DATA, else ~/.weave (created when missing, mode 0700).
Exit codes: 0 recorded (or reported), 1 bad arguments. Never a traceback.
"""
from __future__ import annotations

import argparse
import contextlib
import datetime as _dt
import hashlib
import json
import os
import re
import secrets
import sys
import time
import uuid
from pathlib import Path
from typing import List, Optional

SCRIPT_DIR = Path(__file__).resolve().parent

# 64 bits of a monthly-salted SHA-256: unlinkable across months and short
# enough to be a plain text column in the usage list.
HASH_HEX_CHARS = 16
# The tool Claude Code calls when it invokes a skill itself; the hook payload
# then carries the skill as "<plugin>:<skill>" (verified 2026-09-05 against
# Claude Code 2.1.228 with a probe plugin).
SKILL_TOOL = "Skill"
# weave.py subcommands that produce a passive event; the others are either
# telemetry itself, stdlib bookkeeping (client, memory), or the doctor.
PASSIVE_EVENT_SUBCOMMANDS = ("build", "review", "validate", "read", "spec-from-file")
FEEDBACK_TYPE = "feedback"
RATING_MIN, RATING_MAX = 1, 5
# How long cmd_flush waits for a concurrent flush (another session's
# SessionStart hook) to release the queue lock before reporting "busy"
# instead of a failure; a module constant, not an inline default, so tests
# can shrink it rather than waiting out a real multi-second timeout.
FLUSH_LOCK_TIMEOUT_S = 5.0
# How old an unreleased lock file must be before it's presumed abandoned by
# a crashed holder (never by how long a healthy caller is willing to wait --
# see _locked()'s docstring). The queue lock's own critical sections never
# include network calls (see cmd_flush), so a healthy hold time is always a
# small fraction of this.
LOCK_STALE_AFTER_S = 60.0
LOCK_POLL_INTERVAL_S = 0.05
# Module-level indirection (not a bare time.sleep call) purely so tests can
# replace it and skip real waiting; cmd_login_poll's pacing sleep is real
# production behavior, not a test seam otherwise.
_sleep = time.sleep


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


def telemetry_dir(create: bool = True) -> Path:
    d = data_dir(create) / "telemetry"
    if create:
        try:
            d.mkdir(parents=True, exist_ok=True)
            os.chmod(d, 0o700)
        except OSError:
            pass
    return d


def events_path() -> Path:
    return telemetry_dir() / "events.jsonl"


def detect_surface() -> str:
    env = os.environ.get("WEAVE_SURFACE")
    if env:
        return env
    if os.environ.get("CLAUDE_PLUGIN_ROOT"):
        return "code"
    here = str(SCRIPT_DIR)
    if "/sessions/" in here or here.startswith("/mnt/"):
        return "cowork"
    return "unknown"


def _plugin_field(name: str) -> Optional[str]:
    """Read one field of the nearest .claude-plugin/plugin.json above this script."""
    p = SCRIPT_DIR
    for _ in range(5):  # the script itself plus up to 4 parents
        cand = p / ".claude-plugin" / "plugin.json"
        if cand.exists():
            try:
                return json.loads(cand.read_text(encoding="utf-8")).get(name)
            except (OSError, ValueError):
                return None
        if p.parent == p:
            break
        p = p.parent
    return None


def plugin_version() -> str:
    return _plugin_field("version") or "dev"


def plugin_prefix() -> str:
    """How Claude Code names this plugin's skills in hook payloads: '<plugin>:'."""
    return (_plugin_field("name") or "weave") + ":"


def _fail(error: str, hint: str) -> int:
    """Always JSON, regardless of --json (matches weave.py's own emit_error)."""
    print(json.dumps({"ok": False, "error": error, "hint": hint}))
    return 1


def _load_json(path: Path) -> Optional[dict]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _write_json_atomic(path: Path, obj: dict) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(obj), encoding="utf-8")
    tmp.replace(path)


def _current_salt() -> str:
    """One salt per calendar month (UTC), so a hash from one month can never
    be correlated with the same user/client's hash from another month."""
    month = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m")
    path = telemetry_dir() / "salt.json"
    rec = _load_json(path) or {}
    if rec.get("month") == month and rec.get("salt"):
        return rec["salt"]
    salt = secrets.token_hex(32)
    _write_json_atomic(path, {"month": month, "salt": salt})
    return salt


def _install_id() -> str:
    """A random, anonymous per-install identifier -- not tied to any real
    account -- that _hash() salts to make user_hash. Generated once and
    reused after that; rotating it on every call would be indistinguishable
    from a fresh install every time, which defeats a hash that is supposed
    to stay stable within one salt month."""
    path = telemetry_dir() / "install_id.json"
    rec = _load_json(path) or {}
    if rec.get("id"):
        return rec["id"]
    new_id = uuid.uuid4().hex
    _write_json_atomic(path, {"id": new_id})
    return new_id


def _hash(value: str) -> str:
    return hashlib.sha256((_current_salt() + value).encode("utf-8")).hexdigest()[:HASH_HEX_CHARS]


def _slug(value: str) -> str:
    """Same shape as weave.py's slugify, so a client passed as a display name
    to `review --client` and as a slug to `build --client-slug` hash alike."""
    s = re.sub(r"[^a-z0-9]+", "-", value.strip().lower()).strip("-")
    return re.sub(r"-{2,}", "-", s)


def _timestamp() -> str:
    return _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def enabled(event_type: Optional[str]) -> bool:
    """Off when either the plugin-settings toggle or WEAVE_TELEMETRY says so
    (either turns it off, per GOVERNANCE.md Sec 5); feedback always records,
    since filing it is the point of the command."""
    if event_type == FEEDBACK_TYPE:
        return True
    if os.environ.get("WEAVE_TELEMETRY") == "0":
        return False
    toggle = os.environ.get("CLAUDE_PLUGIN_OPTION_TELEMETRY_ENABLED")
    return toggle is None or toggle.strip().lower() not in ("false", "0", "no", "off", "")


# ---------------------------------------------------------------------------
# recording
# ---------------------------------------------------------------------------
def record(fields: dict) -> Optional[dict]:
    """Append one event built from `fields` (a `type` plus GOVERNANCE.md Sec 5
    fields; None values are dropped). `client_slug` is hashed, never stored,
    and ignored for feedback events. Returns the event, or None when
    telemetry is off for this type."""
    event_type = fields.get("type") or "event"
    if not enabled(event_type):
        return None
    event = {
        "event_id": uuid.uuid4().hex,
        "timestamp": _timestamp(),
        "plugin_version": plugin_version(),
        "surface": detect_surface(),
        "user_hash": _hash(_install_id()),
    }
    client = fields.get("client_slug")
    for key, value in fields.items():
        if key == "client_slug" or value is None:
            continue
        event[key] = value
    if client and event_type != FEEDBACK_TYPE:
        event["client_hash"] = _hash(_slug(str(client)))
    path = events_path()
    # Shares queue.lock with cmd_flush's drop step: without it, an append
    # landing between a flush's read and its atomic replace of events.jsonl
    # is silently lost. Cheap here because the lock's critical sections are
    # local file I/O only (cmd_flush never holds it across a network call).
    with _locked(telemetry_dir() / "queue.lock", timeout_s=FLUSH_LOCK_TIMEOUT_S):
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(event, sort_keys=True) + "\n")
    return event


def _arg_value(args: List[str], flag: str) -> Optional[str]:
    for i, a in enumerate(args):
        if a == flag and i + 1 < len(args):
            return args[i + 1]
        if a.startswith(flag + "="):
            return a.split("=", 1)[1]
    return None


def passive_fields(sub: str, args: List[str], result: Optional[dict], rc: int, duration_s: float) -> dict:
    """The event weave.py records after dispatching `sub`. `result` is the
    child's JSON object when there was one (None in human-output mode), so
    outcome falls back to the exit code alone."""
    r = result if isinstance(result, dict) else {}
    validation = r.get("validation") if isinstance(r.get("validation"), dict) else {}
    results = r.get("results") if isinstance(r.get("results"), list) else None
    hard_fails = validation.get("hard_fails")
    if hard_fails is None and results is not None:  # validate_deliverable.py's shape
        hard_fails = sum(1 for x in results if isinstance(x, dict) and x.get("hard") and not x.get("passed"))
    ok = rc == 0 and r.get("ok", True) is not False
    # spec-from-file's result nests kind/recipe under spec.weave (the
    # embedded plan), not at the top level like build/review/validate/read.
    weave = r.get("spec", {}).get("weave") if sub == "spec-from-file" and isinstance(r.get("spec"), dict) else None
    fields = {
        "type": sub,
        "recipe": weave.get("recipe") if weave is not None else r.get("recipe"),
        "outcome": "pass" if ok and not hard_fails else "fail",
        "hard_fails": hard_fails,
        "warnings": validation.get("warnings"),
        "score": r.get("score"),
        "duration_s": round(duration_s, 3),
        "client_slug": _arg_value(args, "--client-slug") or _arg_value(args, "--client"),
    }
    if sub == "read":
        fields["source_types"] = [r["type"]] if r.get("type") else None
    elif weave is not None:
        fields["kind"] = weave.get("kind")
    else:
        fields["kind"] = r.get("kind") or r.get("type")
    return fields


def _model_chosen_skill(payload: dict) -> Optional[str]:
    tool_input = payload.get("tool_input") if isinstance(payload.get("tool_input"), dict) else {}
    return tool_input.get("skill")


def _user_typed_skill(payload: dict) -> Optional[str]:
    return payload.get("command_name")


# One row per hook payload shape Claude Code is known to send for a skill
# invocation (hook_event_name, an extra match on the payload, the invocation
# label, how to pull the skill name out) -- verified empirically 2026-09-05
# against Claude Code 2.1.228 with a probe plugin (not documented). Add a row
# here, not a new branch in skill_from_hook, when Claude Code adds a third way.
HOOK_SHAPES = (
    ("PostToolUse", lambda p: p.get("tool_name") == SKILL_TOOL, "model", _model_chosen_skill),
    ("UserPromptExpansion", lambda p: True, "user", _user_typed_skill),
)


def skill_from_hook(payload: dict, prefix: str) -> Optional[tuple]:
    """(skill, invocation) when a hook payload names one of this plugin's
    skills, else None."""
    event = payload.get("hook_event_name")
    for hook_event, matches, invocation, extract in HOOK_SHAPES:
        if event != hook_event or not matches(payload):
            continue
        name = extract(payload)
        if isinstance(name, str) and name.startswith(prefix):
            return name[len(prefix):], invocation
        return None
    return None


def cmd_event(args) -> int:
    if not enabled(args.type):
        if args.json:
            print(json.dumps({"ok": True, "recorded": False, "reason": "WEAVE_TELEMETRY=0"}))
        else:
            print("Telemetry is off (WEAVE_TELEMETRY=0); event not recorded.")
        return 0
    if args.rating is not None and not (RATING_MIN <= args.rating <= RATING_MAX):
        return _fail("--rating must be %d-%d" % (RATING_MIN, RATING_MAX), "Omit --rating if there is none.")
    event = record({
        "type": args.type,
        "skill": args.skill,
        "recipe": args.recipe,
        "kind": args.kind,
        "outcome": args.outcome,
        "hard_fails": args.hard_fails,
        "warnings": args.warnings,
        "duration_s": args.duration_s,
        "edit_iterations": args.edit_iterations,
        "source_types": args.source_types.split(",") if args.source_types else None,
        "rating": args.rating,
        "text": args.text,
        "client_slug": args.client_slug,
    })
    if args.json:
        print(json.dumps({"ok": True, "recorded": True, "event_id": event["event_id"], "file": str(events_path())}))
    else:
        print("Recorded %s event %s -> %s" % (args.type, event["event_id"], events_path()))
    return 0


def cmd_hook(args) -> int:
    """Hook stdout is shown to Claude or the user, so stay silent; a hook
    that fails must never disturb the session, so always exit 0."""
    try:
        payload = json.loads(sys.stdin.read() or "{}")
        found = skill_from_hook(payload, plugin_prefix()) if isinstance(payload, dict) else None
        if found is not None:
            skill, invocation = found
            ms = payload.get("duration_ms")
            record({"type": "skill_invoked", "skill": skill, "invocation": invocation,
                    "duration_s": round(ms / 1000.0, 3) if isinstance(ms, (int, float)) else None})
    except Exception:  # noqa: BLE001 - a telemetry hook must never break a session
        pass
    return 0


# ---------------------------------------------------------------------------
# queue, status, sending
# ---------------------------------------------------------------------------
def _read_events(path: Path) -> List[str]:
    if not path.exists():
        return []
    with open(path, encoding="utf-8") as fh:
        return [line for line in fh if line.strip()]


@contextlib.contextmanager
def _locked(lock_path: Path, timeout_s: float = FLUSH_LOCK_TIMEOUT_S,
            stale_after_s: float = LOCK_STALE_AFTER_S, poll_s: float = LOCK_POLL_INTERVAL_S):
    """Cross-platform advisory lock via exclusive file creation (no fcntl,
    so it also works on Windows). Waits up to `timeout_s` for a healthy
    holder to finish, then raises TimeoutError -- a separate, much larger
    `stale_after_s` is what decides a lock is abandoned (its holder crashed
    without cleaning up): a lock only that old is reclaimed on sight, on any
    retry, independent of how impatient this particular caller is. Using
    `timeout_s` for both would make every wait-then-give-up look "stale" by
    definition, reclaiming locks healthy processes still hold.

    Reclaiming is stat-then-unlink, which is a TOCTOU by itself: two waiters
    could both see the same stale lock and both reclaim it, or a waiter could
    unlink a lock a healthy holder only just recreated. Re-checking the mtime
    immediately before unlinking narrows that window to the gap between the
    two stat() calls -- not zero, but the only race left requires another
    process's full unlink-then-recreate to land inside it, vanishingly
    unlikely for a local file and proportionate for an advisory lock like
    this one (not a distributed/financial one)."""
    deadline = time.monotonic() + timeout_s
    fd = None
    while fd is None:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                mtime = lock_path.stat().st_mtime  # wall-clock, unlike monotonic()
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


def _drop_sent_and_bad(path: Path, sent_ids: set) -> None:
    """Remove every line that is malformed (can never be sent, whenever it
    arrived) or whose event_id is in `sent_ids` (just transmitted). Matches
    by identity against a FRESH read under queue.lock, never by position or
    count -- so a line appended by a concurrent record() since the batch was
    read (it lands at the end, with an event_id not in `sent_ids`) survives
    regardless of where it ends up, and a batch that only partially sent
    (throttled/network/auth error partway through) never loses an event
    that was never actually transmitted."""
    remaining = []
    for line in _read_events(path):
        try:
            ev = json.loads(line)
        except ValueError:
            continue  # malformed: can never be sent, drop it on sight
        if ev.get("event_id") in sent_ids:
            continue  # already sent this flush (or a concurrent one -- idempotent either way)
        remaining.append(line)
    tmp = path.with_suffix(".jsonl.tmp")
    tmp.write_text("".join(remaining), encoding="utf-8")
    tmp.replace(path)


def cmd_flush(args) -> int:
    quiet = getattr(args, "quiet", False)
    path = events_path()
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402 - sibling script

    cfg = sink.load_config(telemetry_dir())
    # Safe unlocked: tmp.replace() is atomic, so this is always a consistent
    # snapshot (old or new content, never torn) -- just possibly a little
    # stale relative to a concurrent flush, which the identity-based drop
    # below tolerates.
    queued = _read_events(path)
    report = {"ok": True, "queued": len(queued), "sent": 0, "file": str(path)}
    if cfg is None:
        report["note"] = "local-only: no usage-list sink is configured, so nothing is transmitted"
        return _emit_flush(report, args.json, quiet, notable=False)
    state = _load_json(telemetry_dir() / "state.json") or {}
    now = _dt.datetime.now(_dt.timezone.utc).timestamp()
    if state.get("retry_after_until", 0) > now:
        report["note"] = "the usage list asked us to wait; will retry after %s" % (
            _dt.datetime.fromtimestamp(state["retry_after_until"], _dt.timezone.utc).isoformat())
        return _emit_flush(report, args.json, quiet, notable=False)
    try:
        token = sink.get_token(cfg, telemetry_dir(), version=plugin_version())
    except sink.SinkTransportError:
        report["note"] = "could not reach the usage list (network unavailable); will retry next session"
        return _emit_flush(report, args.json, quiet, notable=len(queued) > 0)
    if token is None:
        report["note"] = "not signed in to the usage list; run /weave:setup to sign in"
        return _emit_flush(report, args.json, quiet, notable=len(queued) > 0)
    if not queued:
        return _emit_flush(report, args.json, quiet, notable=False)
    batch = queued[: cfg["flush_max_events"]]
    events = []
    for line in batch:
        try:
            events.append(json.loads(line))
        except ValueError:  # a partial write or crash mid-append; _drop_sent_and_bad drops it regardless of send outcome
            pass
    # No lock held here: get_token/send are network calls, and the file lock
    # exists to protect fast local I/O (see _locked's docstring), not to
    # serialize this device's network latency across every flush attempt.
    result = sink.send(cfg, token, events, plugin_version())
    sent_ids = {events[i]["event_id"] for i in range(result["sent"])}
    try:
        with _locked(telemetry_dir() / "queue.lock", timeout_s=FLUSH_LOCK_TIMEOUT_S):
            _drop_sent_and_bad(path, sent_ids)
            if result.get("retry_after_s"):
                _write_json_atomic(telemetry_dir() / "state.json", {"retry_after_until": now + result["retry_after_s"]})
    except TimeoutError:
        report["note"] = "sent, but could not update the local queue this time; unsent items will resend next flush"
    report.update({"sent": result["sent"], "queued": len(_read_events(path)), "stopped": result.get("stopped")})
    return _emit_flush(report, args.json, quiet, notable=bool(result.get("stopped")) or bool(report.get("note")))


def _emit_flush(report: dict, as_json: bool, quiet: bool, notable: bool) -> int:
    if as_json:
        print(json.dumps(report))
    elif not quiet:
        print("%d event(s) sent, %d queued at %s. %s" % (
            report["sent"], report["queued"], report["file"], report.get("note") or report.get("stopped") or ""))
    elif notable:
        print("Weave telemetry: %s" % (report.get("note") or report.get("stopped")))
    return 0


def cmd_status(args) -> int:
    on = enabled(None)  # any non-feedback type: is passive telemetry on?
    path = events_path()
    n = len(_read_events(path))
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402

    cfg = sink.load_config(telemetry_dir())
    signed_in = bool(cfg) and sink.load_tokens(telemetry_dir()) is not None
    note = "feedback events (--type feedback) always record; WEAVE_TELEMETRY=0 only silences passive usage events"
    if args.json:
        print(json.dumps({"ok": True, "enabled": on, "queued": n, "file": str(path),
                          "sink_configured": bool(cfg), "signed_in": signed_in, "note": note}))
    else:
        print("Telemetry %s. %d event(s) queued at %s." % ("on" if on else "off (WEAVE_TELEMETRY=0)", n, path))
        print("Usage list: %s%s." % ("configured" if cfg else "not configured",
                                     ", signed in" if signed_in else (", not signed in" if cfg else "")))
        print(note)
    return 0


def _login_pending_path() -> Path:
    return telemetry_dir() / "login_pending.json"


def _pending_sign_in(now: float) -> Optional[dict]:
    """The pending device-code state, or None if there isn't one or it has
    already expired (Microsoft's own expires_in window, tracked locally as
    an absolute _expires_at so login-poll never has to guess)."""
    start = _load_json(_login_pending_path())
    if start is None or start.get("_expires_at", 0) <= now:
        return None
    return start


def cmd_login_start(args) -> int:
    """Start a device-code sign-in and return immediately with Microsoft's
    instruction text; login-poll checks on it. Split from a single blocking
    call so one sign-in is never one multi-minute tool invocation."""
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402

    cfg = sink.load_config(telemetry_dir())
    if cfg is None:
        return _fail("No usage-list sink is configured.",
                     "An administrator sets telemetry_endpoint, telemetry_tenant_id and telemetry_client_id "
                     "in the plugin settings, or runs: telemetry.py configure --endpoint ... --tenant-id ... --client-id ...")
    now = time.time()
    existing = _pending_sign_in(now)
    if existing is not None:
        # A still-valid sign-in from an earlier login-start is in progress;
        # requesting a second device code here would silently orphan it (its
        # browser tab, if the user completes it, would authorize a code
        # nothing polls for anymore). Re-show the same one instead.
        start = existing
    else:
        try:
            start = sink.device_code_start(cfg, version=plugin_version())
        except sink.SinkError as exc:
            return _fail("Could not start sign-in: %s" % exc, "Run /weave:setup again to retry.")
        start["_expires_at"] = now + float(start["expires_in"])
        _write_json_atomic(_login_pending_path(), start)
    if args.json:
        print(json.dumps({"ok": True, "message": start["message"], "interval_s": start["interval"],
                          "expires_in_s": max(0, start["_expires_at"] - now)}))
    else:
        # The message is Microsoft's own instruction text; show it verbatim.
        print(start["message"])
    return 0


def cmd_login_poll(args) -> int:
    """One poll attempt against the sign-in login-start began, waiting the
    server-paced interval first (bounded to a few seconds -- never the whole
    device-code window) so the skill's own re-invocation timing can't poll
    faster than the server allows. Returns signed_in: false while the user
    is still completing it in the browser; the skill calls this again after
    the interval it was told."""
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402

    cfg = sink.load_config(telemetry_dir())
    if cfg is None:
        return _fail("No usage-list sink is configured.", "Run telemetry.py login-start first.")
    start = _pending_sign_in(time.time())
    if start is None:
        _login_pending_path().unlink(missing_ok=True)  # clears a merely-expired file too
        return _fail("No sign-in is in progress (it may have expired).", "Run telemetry.py login-start again.")
    _sleep(float(start["interval"]))
    try:
        state, tokens = sink.device_code_poll_once(cfg, start, version=plugin_version())
    except sink.SinkTransportError:
        # A one-off network blip, not a real sign-in failure: the device
        # code and pending state are both still valid, so just ask the
        # caller to try again rather than discarding either.
        if args.json:
            print(json.dumps({"ok": True, "signed_in": False, "retry_in_s": start["interval"]}))
        else:
            print("Could not reach the sign-in server; will try again.")
        return 0
    except sink.SinkError as exc:
        _login_pending_path().unlink(missing_ok=True)
        return _fail("Sign-in did not complete: %s" % exc, "Run telemetry.py login-start again when ready.")
    if state == "slow_down":
        start["interval"] = float(start["interval"]) + sink.SLOW_DOWN_INCREMENT_S
        _write_json_atomic(_login_pending_path(), start)
        state = "pending"
    if state == "pending":
        if args.json:
            print(json.dumps({"ok": True, "signed_in": False, "retry_in_s": start["interval"]}))
        else:
            print("Still waiting for sign-in.")
        return 0
    sink.save_tokens(telemetry_dir(), tokens)
    _login_pending_path().unlink(missing_ok=True)
    if args.json:
        print(json.dumps({"ok": True, "signed_in": True}))
    else:
        print("Signed in. Queued usage events will be sent at the start of your next session.")
    return 0


def cmd_logout(args) -> int:
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402

    removed = sink.clear_tokens(telemetry_dir())
    print(json.dumps({"ok": True, "signed_out": removed}) if args.json else
          ("Signed out of the usage list." if removed else "Not signed in."))
    return 0


def cmd_configure(args) -> int:
    sys.path.insert(0, str(SCRIPT_DIR))
    import telemetry_sink as sink  # noqa: E402

    try:
        cfg = sink.save_config(telemetry_dir(), {"endpoint": args.endpoint, "tenant_id": args.tenant_id,
                                                  "client_id": args.client_id})
    except sink.SinkError as exc:
        return _fail(str(exc), "The endpoint is the Graph URL of the list's items collection over https.")
    print(json.dumps({"ok": True, "sink": cfg}) if args.json else "Usage-list sink configured: %s" % cfg["endpoint"])
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    e = sub.add_parser("event", help="record one usage or feedback event")
    e.add_argument("--type", required=True, help="e.g. feedback, build, review")
    e.add_argument("--skill", default=None)
    e.add_argument("--recipe", default=None)
    e.add_argument("--kind", default=None, choices=["doc", "deck", "sow"])
    e.add_argument("--outcome", default=None, choices=["pass", "fail"])
    e.add_argument("--hard-fails", type=int, default=None)
    e.add_argument("--warnings", type=int, default=None)
    e.add_argument("--duration-s", type=float, default=None)
    e.add_argument("--edit-iterations", type=int, default=None)
    e.add_argument("--source-types", default=None, help="comma-separated, e.g. docx,pdf")
    e.add_argument("--rating", type=int, default=None, help="%d-%d" % (RATING_MIN, RATING_MAX))
    e.add_argument("--text", default=None, help="general description only -- never prompt/file/client text")
    e.add_argument("--client-slug", default=None, help="hashed, never stored raw; ignored for --type feedback")
    e.add_argument("--json", action="store_true")
    e.set_defaults(func=cmd_event)

    h = sub.add_parser("hook", help="record a skill_invoked event from a Claude Code hook payload on stdin")
    h.set_defaults(func=cmd_hook)

    f = sub.add_parser("flush", help="send queued events to the configured sink")
    f.add_argument("--quiet", action="store_true", help="print only when the user needs to act (for hooks)")
    f.add_argument("--json", action="store_true")
    f.set_defaults(func=cmd_flush)

    s = sub.add_parser("status", help="whether telemetry is enabled, and how much is queued")
    s.add_argument("--json", action="store_true")
    s.set_defaults(func=cmd_status)

    ls = sub.add_parser("login-start", help="start device-code sign-in to the usage list; returns immediately")
    ls.add_argument("--json", action="store_true")
    ls.set_defaults(func=cmd_login_start)

    lp = sub.add_parser("login-poll", help="one poll attempt for the sign-in login-start began")
    lp.add_argument("--json", action="store_true")
    lp.set_defaults(func=cmd_login_poll)

    lo = sub.add_parser("logout", help="forget the usage-list sign-in")
    lo.add_argument("--json", action="store_true")
    lo.set_defaults(func=cmd_logout)

    c = sub.add_parser("configure", help="record where usage events are sent")
    c.add_argument("--endpoint", required=True)
    c.add_argument("--tenant-id", required=True)
    c.add_argument("--client-id", required=True)
    c.add_argument("--json", action="store_true")
    c.set_defaults(func=cmd_configure)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001 - top-level guard: always emit the JSON contract
        if args.cmd == "hook":
            return 0
        return _fail("unexpected error: %s" % exc, "Run again without --json to see the traceback on stderr.")


if __name__ == "__main__":
    sys.exit(main())
