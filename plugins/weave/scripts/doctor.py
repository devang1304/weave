#!/usr/bin/env python3
"""
Weave environment check and one-shot installer.  Standard library only, so it
runs before any dependency exists (it is what the SessionStart hook calls).

  doctor.py            plain-language report
  doctor.py --quiet    one line, always exit 0 (safe for hooks)
  doctor.py --json     machine-readable report
  doctor.py --install  make the environment ready: reuse the current
                       interpreter when it already has the required packages,
                       otherwise create <data>/venv and pip install the pinned
                       requirements.txt beside this script; writes <data>/env.json

Data dir: $CLAUDE_PLUGIN_DATA, else ~/.weave (created when missing, mode 0700).

Exit codes: 0 ready (or --quiet), 1 something required is missing, 2 install
could not complete (the message says what to run by hand).  Never a traceback.
"""
from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

MIN_PYTHON = (3, 9)
# (distribution name, import name)
REQUIRED = [("python-docx", "docx"), ("python-pptx", "pptx"), ("lxml", "lxml")]
OPTIONAL = [("openpyxl", "openpyxl"), ("pypdf", "pypdf")]
SOFFICE_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/usr/bin/soffice",
)
SCRIPT_DIR = Path(__file__).resolve().parent
REQUIREMENTS = SCRIPT_DIR / "requirements.txt"

# Small stdlib program run inside another interpreter to learn what it has.
_PROBE = r"""
import json, sys, importlib
out = {"python": sys.executable, "version": list(sys.version_info[:3]), "modules": {}}
for dist, mod in %s:
    try:
        m = importlib.import_module(mod)
        try:
            from importlib import metadata
            ver = metadata.version(dist)
        except Exception:
            ver = getattr(m, "__version__", None)
        out["modules"][dist] = {"present": True, "version": ver}
    except Exception:
        out["modules"][dist] = {"present": False, "version": None}
print(json.dumps(out))
"""


# ---------------------------------------------------------------------------
# discovery
# ---------------------------------------------------------------------------
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


def find_soffice() -> Optional[str]:
    for cand in ("soffice", "libreoffice"):
        p = shutil.which(cand)
        if p:
            return p
    for cand in SOFFICE_CANDIDATES:
        if os.path.exists(cand):
            return cand
    return None


def probe_here(pairs: List[Tuple[str, str]]) -> Dict[str, Dict[str, object]]:
    """Import each module in this interpreter; report presence and version."""
    import importlib

    found: Dict[str, Dict[str, object]] = {}
    for dist, mod in pairs:
        try:
            m = importlib.import_module(mod)
        except Exception:  # noqa: BLE001 - any import failure means "absent"
            found[dist] = {"present": False, "version": None}
            continue
        ver = None
        try:
            from importlib import metadata

            ver = metadata.version(dist)
        except Exception:  # noqa: BLE001
            ver = getattr(m, "__version__", None)
        found[dist] = {"present": True, "version": ver}
    return found


def probe_interpreter(python: str, pairs: List[Tuple[str, str]]) -> Optional[Dict[str, object]]:
    """Run the probe inside another interpreter.  None when it cannot run."""
    if not python or not os.path.exists(python):
        return None
    code = _PROBE % json.dumps([list(p) for p in pairs])
    try:
        cp = subprocess.run([python, "-c", code], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60)
    except (OSError, subprocess.TimeoutExpired):
        return None
    if cp.returncode != 0:
        return None
    try:
        return json.loads(cp.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError):
        return None


def writable(d: Path) -> bool:
    try:
        probe = d / ".weave-write-test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def read_env_json(d: Path) -> Optional[dict]:
    p = d / "env.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


# ---------------------------------------------------------------------------
# report
# ---------------------------------------------------------------------------
def build_report() -> dict:
    d = data_dir()
    py_ok = sys.version_info[:2] >= MIN_PYTHON
    required = probe_here(REQUIRED)
    optional = probe_here(OPTIONAL)
    interpreter = sys.executable
    env = read_env_json(d)
    env_json_path = str(d / "env.json") if env else None

    # When this interpreter lacks something, the recorded venv may have it.
    missing_here = [k for k, v in required.items() if not v["present"]]
    if missing_here and env and env.get("python") and env["python"] != sys.executable:
        probed = probe_interpreter(env["python"], REQUIRED + OPTIONAL)
        if probed:
            mods = probed.get("modules", {})
            alt_required = {k: mods.get(k, {"present": False, "version": None}) for k, _ in REQUIRED}
            if all(v["present"] for v in alt_required.values()):
                required = alt_required
                optional = {k: mods.get(k, {"present": False, "version": None}) for k, _ in OPTIONAL}
                interpreter = probed.get("python", env["python"])
                py_ok = tuple(probed.get("version", [0, 0]))[:2] >= MIN_PYTHON

    missing = [k for k, v in required.items() if not v["present"]]
    if not py_ok:
        missing.insert(0, "python>=%d.%d" % MIN_PYTHON)
    return {
        "ok": py_ok and not [k for k, v in required.items() if not v["present"]],
        "python": platform.python_version(),
        "python_ok": py_ok,
        "interpreter": interpreter,
        "surface": detect_surface(),
        "data_dir": str(d),
        "data_dir_writable": writable(d),
        "required": required,
        "optional": optional,
        "soffice": find_soffice(),
        "env_json": env_json_path,
        "missing": missing,
    }


def print_human(rep: dict) -> None:
    mark = lambda ok: "ok     " if ok else "missing"  # noqa: E731
    print("Weave environment check")
    print("  Python %s at %s  [%s]" % (rep["python"], rep["interpreter"],
                                       "ok" if rep["python_ok"] else "too old, need %d.%d+" % MIN_PYTHON))
    print("  Required packages:")
    for name, info in rep["required"].items():
        ver = (" %s" % info["version"]) if info.get("version") else ""
        print("    %s  %s%s" % (mark(info["present"]), name, ver))
    print("  Optional packages (only for reading .xlsx / .pdf sources):")
    for name, info in rep["optional"].items():
        ver = (" %s" % info["version"]) if info.get("version") else ""
        print("    %s  %s%s" % (mark(info["present"]), name, ver))
    print("  LibreOffice (optional, used to render page previews): %s" % (rep["soffice"] or "not found"))
    print("  Data folder: %s (%s)" % (rep["data_dir"], "writable" if rep["data_dir_writable"] else "NOT writable"))
    print("  Surface: %s" % rep["surface"])
    if rep["env_json"]:
        print("  Recorded environment: %s" % rep["env_json"])
    if rep["ok"]:
        print("Weave is ready.")
    else:
        print("Weave is missing: %s." % ", ".join(rep["missing"]))
        print("Run /weave:setup (or: python3 %s --install) to install what Weave needs." % Path(__file__).name)


# ---------------------------------------------------------------------------
# install
# ---------------------------------------------------------------------------
def _sha256(path: Path) -> Optional[str]:
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError:
        return None


def _write_env_json(d: Path, python: str, installed: List[str]) -> Path:
    payload = {
        "python": python,
        "installed": installed,
        "created_at": _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "requirements_sha256": _sha256(REQUIREMENTS),
    }
    p = d / "env.json"
    p.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    try:
        os.chmod(p, 0o600)
    except OSError:
        pass
    return p


def _installed_list(mods: Dict[str, Dict[str, object]]) -> List[str]:
    return ["%s==%s" % (k, v["version"]) if v.get("version") else k
            for k, v in mods.items() if v.get("present")]


def _looks_like_network_failure(text: str) -> bool:
    text = text.lower()
    needles = ("could not fetch", "connection", "temporary failure", "proxyerror",
               "newconnectionerror", "max retries", "network is unreachable",
               "no matching distribution", "read timed out", "ssl", "name or service not known")
    return any(n in text for n in needles)


def _venv_python(venv: Path) -> Path:
    if os.name == "nt":
        return venv / "Scripts" / "python.exe"
    return venv / "bin" / "python"


def install(as_json: bool) -> int:
    d = data_dir()
    if not writable(d):
        return _fail(as_json, "The Weave data folder %s is not writable." % d,
                     "Set CLAUDE_PLUGIN_DATA to a folder you can write to, then run /weave:setup again.")

    # 1. Current interpreter already good enough?  Record it and stop.
    here = probe_here(REQUIRED + OPTIONAL)
    if sys.version_info[:2] >= MIN_PYTHON and all(here[k]["present"] for k, _ in REQUIRED):
        p = _write_env_json(d, sys.executable, _installed_list(here))
        return _done(as_json, sys.executable, p, venv=None,
                     note="This Python already has everything Weave needs; no download required.")

    if not REQUIREMENTS.exists():
        return _fail(as_json, "requirements.txt is missing beside %s." % Path(__file__).name,
                     "Reinstall the Weave plugin; the file should ship with it.")

    # 2. Create (or reuse) the venv.
    venv = d / "venv"
    vpy = _venv_python(venv)
    if not vpy.exists():
        cp = subprocess.run([sys.executable, "-m", "venv", str(venv)], capture_output=True, text=True, encoding="utf-8", errors="replace")
        if cp.returncode != 0 or not vpy.exists():
            shutil.rmtree(venv, ignore_errors=True)
            cp = subprocess.run([sys.executable, "-m", "venv", "--without-pip", str(venv)],
                                capture_output=True, text=True, encoding="utf-8", errors="replace")
            if cp.returncode != 0 or not vpy.exists():
                return _fail(as_json, "Could not create a Python environment at %s." % venv,
                             "Run by hand: %s -m venv %s   then run /weave:setup again.\n%s"
                             % (sys.executable, venv, (cp.stderr or "").strip()[-600:]))
            ep = subprocess.run([str(vpy), "-m", "ensurepip", "--upgrade"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            if ep.returncode != 0:
                return _fail(as_json, "The new environment has no pip.",
                             "Run by hand: %s -m ensurepip --upgrade   then run /weave:setup again." % vpy)

    # 3. pip install pinned requirements.
    cmd = [str(vpy), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(REQUIREMENTS)]
    try:
        cp = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
    except subprocess.TimeoutExpired:
        return _fail(as_json, "Installing Weave's packages took too long and was stopped.",
                     "Check your internet connection, then run by hand: %s" % " ".join(cmd))
    if cp.returncode != 0:
        blob = (cp.stdout or "") + (cp.stderr or "")
        if _looks_like_network_failure(blob):
            return _fail(as_json, "Weave could not download its packages (no network, or a proxy is blocking pip).",
                         "When you are online, run this one command and then run /weave:setup again:\n  %s"
                         % " ".join(cmd))
        return _fail(as_json, "pip could not install Weave's packages.",
                     "Run by hand and read the message:\n  %s\n%s" % (" ".join(cmd), blob.strip()[-800:]))

    probed = probe_interpreter(str(vpy), REQUIRED + OPTIONAL)
    mods = (probed or {}).get("modules", {})
    if not all(mods.get(k, {}).get("present") for k, _ in REQUIRED):
        return _fail(as_json, "pip finished but the packages still do not import inside %s." % vpy,
                     "Delete %s and run /weave:setup again." % venv)
    p = _write_env_json(d, str(vpy), _installed_list(mods))
    return _done(as_json, str(vpy), p, venv=str(venv), note="Installed into a private environment; nothing else on your machine was changed.")


def _fail(as_json: bool, error: str, hint: str) -> int:
    if as_json:
        print(json.dumps({"ok": False, "error": error, "hint": hint}))
    else:
        print("Weave setup did not finish.")
        print("  " + error)
        print("  " + hint.replace("\n", "\n  "))
    return 2


def _done(as_json: bool, python: str, env_json: Path, venv: Optional[str], note: str) -> int:
    if as_json:
        rep = build_report()
        rep.update({"installed": True, "env_json": str(env_json), "venv": venv, "note": note})
        print(json.dumps(rep, indent=2))
    else:
        print("Weave is ready.")
        print("  Using Python: %s" % python)
        print("  Recorded in:  %s" % env_json)
        print("  " + note)
    return 0


# ---------------------------------------------------------------------------
def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Check (or set up) what Weave needs to run.")
    ap.add_argument("--quiet", action="store_true", help="one line, always exit 0 (for hooks)")
    ap.add_argument("--install", action="store_true", help="create the venv and install pinned requirements")
    ap.add_argument("--json", action="store_true", help="machine-readable report")
    args = ap.parse_args(argv)

    try:
        if args.install:
            return install(args.json)
        rep = build_report()
        if args.quiet:
            if rep["ok"]:
                print("Weave: ready")
            else:
                print("Weave: missing %s. Run /weave:setup" % ", ".join(rep["missing"]))
            return 0
        if args.json:
            print(json.dumps(rep, indent=2))
        else:
            print_human(rep)
        return 0 if rep["ok"] else 1
    except Exception as exc:  # noqa: BLE001 - never show a traceback to a non-technical user
        if args.quiet:
            print("Weave: check could not run (%s). Run /weave:setup" % exc)
            return 0
        if args.json:
            print(json.dumps({"ok": False, "error": str(exc), "hint": "Run /weave:setup to install what Weave needs."}))
        else:
            print("Weave could not check its environment: %s" % exc)
            print("Run /weave:setup to install what Weave needs.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
