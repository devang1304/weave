#!/usr/bin/env python3
"""
telemetry_sink.py -- send Weave usage events to a SharePoint list through
Microsoft Graph, signed in as the user with the OAuth 2.0 device code flow.
Standard library only. Imported by telemetry.py; skills never run it directly.

Configuration (all three required; none is shipped inside the plugin):

  endpoint   Graph URL of the list's items collection, for example
             https://graph.microsoft.com/v1.0/sites/<site-id>/lists/<list-id>/items
  tenant_id  the Microsoft Entra tenant that owns the M365 group's site
  client_id  Application (client) ID of the "Netwoven Weave telemetry" app
             registration: a public client with "Allow public client flows"
             on, delegated permission Sites.Selected (admin-consented), and
             `write` granted on the one site with POST /sites/{id}/permissions

Sources, in order: the environment variables CLAUDE_PLUGIN_OPTION_TELEMETRY_ENDPOINT,
_TENANT_ID and _CLIENT_ID (Claude Code exposes plugin userConfig values to
shell commands under that prefix), then <data>/telemetry/sink.json written by
`telemetry.py configure`.

Tokens live in <data>/telemetry/token.json (mode 0600) and are refreshed
before they expire. Every request carries the SharePoint traffic-decoration
User-Agent "NONISV|Netwoven|Weave/<version>" so the service can attribute and
prioritise the traffic. A 429 or 503 stops the batch and hands the
Retry-After value back to the caller, which is the only correct response to
SharePoint throttling.

Docs: device code https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code
      list items  https://learn.microsoft.com/en-us/graph/api/listitem-create
      Selected    https://learn.microsoft.com/en-us/graph/permissions-selected-overview
      throttling  https://learn.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
"""
from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

ENV_PREFIX = "CLAUDE_PLUGIN_OPTION_TELEMETRY_"
REQUIRED = ("endpoint", "tenant_id", "client_id")
DEFAULTS = {
    "authority": "https://login.microsoftonline.com",
    # Graph's own Selected scope, plus offline_access so a refresh token is
    # issued and the user signs in once per device, not once per session.
    "scope": "https://graph.microsoft.com/Sites.Selected offline_access",
    # One Graph call. Long enough for a slow corporate proxy; short enough
    # that a SessionStart flush can never hold a session open for minutes.
    "http_timeout_s": 30,
    # SharePoint throttles one user at 3,000 requests per 5 minutes; one flush
    # of 200 single-item posts stays under a tenth of that budget.
    "flush_max_events": 200,
    # Refresh this long before expiry so a flush never starts with a token
    # that lapses mid-batch and turns into a 401 half-way through.
    "token_refresh_skew_s": 60,
}
USER_AGENT_FORMAT = "NONISV|Netwoven|Weave/{version}"
DEVICE_CODE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"
# Device-code polling answers that mean "keep waiting" without error, per the
# Entra docs and RFC 8628 sec. 3.5. "slow_down" is handled separately from
# "authorization_pending" (see device_code_poll_once): it also requires the
# client to increase its polling interval, which folding both into one
# undifferentiated "pending" bucket would silently drop.
PENDING_ERROR = "authorization_pending"
SLOW_DOWN_ERROR = "slow_down"
# RFC 8628 sec. 3.5: on slow_down, increase the polling interval by at least
# this many seconds.
SLOW_DOWN_INCREMENT_S = 5
TOKEN_FILE = "token.json"
CONFIG_FILE = "sink.json"

Http = Callable[[str, str, Dict[str, str], Optional[bytes], float], Tuple[int, Dict[str, str], str]]


class SinkError(Exception):
    """A configuration or sign-in problem the user can act on."""


class SinkTransportError(SinkError):
    """Could not reach the host at all (offline, DNS, proxy, timeout) -- as
    opposed to a host that responded with a real OAuth/HTTP error. Distinct
    from the base SinkError so a caller polling repeatedly (device-code
    sign-in) can treat a one-off network blip as 'still waiting', not as the
    sign-in itself having failed."""


def _http(method: str, url: str, headers: Dict[str, str], body: Optional[bytes], timeout: float):
    """(status, response headers, response text). HTTP errors are returned,
    not raised, so callers read Retry-After off a 429; only a failure to
    reach the host raises."""
    req = urllib.request.Request(url, data=body, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, dict(resp.headers.items()), resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        return exc.code, dict(exc.headers.items()), exc.read().decode("utf-8", "replace")
    except (urllib.error.URLError, OSError) as exc:
        raise SinkTransportError("could not reach %s (%s)" % (urllib.parse.urlsplit(url).netloc, exc)) from exc


# ---------------------------------------------------------------------------
# configuration
# ---------------------------------------------------------------------------
def _validate(cfg: dict) -> dict:
    missing = [k for k in REQUIRED if not cfg.get(k)]
    if missing:
        raise SinkError("sink configuration is missing: %s" % ", ".join(missing))
    if not str(cfg["endpoint"]).startswith("https://"):
        raise SinkError("the sink endpoint must be an https URL")
    merged = dict(DEFAULTS)
    merged.update({k: v for k, v in cfg.items() if v is not None})
    return merged


def load_config(tdir: Path, environ: Optional[dict] = None) -> Optional[dict]:
    """The merged sink configuration, or None when nothing is configured.
    Environment (plugin userConfig) wins over sink.json."""
    env = os.environ if environ is None else environ
    from_env = {k: env.get(ENV_PREFIX + k.upper()) for k in REQUIRED}
    if all(from_env.values()):
        return _validate(from_env)
    path = tdir / CONFIG_FILE
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(raw, dict) or not all(raw.get(k) for k in REQUIRED):
        return None
    return _validate(raw)


def save_config(tdir: Path, cfg: dict) -> dict:
    merged = _validate(cfg)
    stored = {k: cfg[k] for k in REQUIRED}
    path = tdir / CONFIG_FILE
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(stored, indent=2) + "\n", encoding="utf-8")
    _chmod_private(tmp)
    tmp.replace(path)
    return merged


def _chmod_private(path: Path) -> None:
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# tokens
# ---------------------------------------------------------------------------
def _token_endpoint(cfg: dict) -> str:
    return "%s/%s/oauth2/v2.0/token" % (cfg["authority"], cfg["tenant_id"])


def _post_form(url: str, data: dict, timeout: float, http: Http, version: str) -> Tuple[int, dict]:
    body = urllib.parse.urlencode(data).encode("utf-8")
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json",
               "User-Agent": USER_AGENT_FORMAT.format(version=version)}
    status, _, text = http("POST", url, headers, body, timeout)
    try:
        parsed = json.loads(text) if text else {}
    except ValueError:
        parsed = {}
    return status, parsed if isinstance(parsed, dict) else {}


def _tokens_from(response: dict, now: float, previous_refresh: Optional[str] = None) -> dict:
    if "access_token" not in response or "expires_in" not in response:
        raise SinkError(response.get("error_description") or response.get("error") or "token response had no access_token")
    return {"access_token": response["access_token"],
            "refresh_token": response.get("refresh_token") or previous_refresh,
            "expires_at": now + float(response["expires_in"])}


def _resolve_http(http: Optional[Http]) -> Http:
    """Look the transport up at call time, so telemetry.py (and the tests)
    can replace module-level `_http`; a default argument would freeze it."""
    return _http if http is None else http


def device_code_start(cfg: dict, http: Optional[Http] = None, version: str = "dev") -> dict:
    """POST /devicecode; returns Microsoft's response (device_code, user_code,
    verification_uri, expires_in, interval, message)."""
    http = _resolve_http(http)
    url = "%s/%s/oauth2/v2.0/devicecode" % (cfg["authority"], cfg["tenant_id"])
    status, resp = _post_form(url, {"client_id": cfg["client_id"], "scope": cfg["scope"]}, cfg["http_timeout_s"], http, version)
    if status != 200 or "device_code" not in resp:
        raise SinkError(resp.get("error_description") or "device code request failed with HTTP %d" % status)
    return resp


def device_code_poll_once(cfg: dict, start: dict, http: Optional[Http] = None, now=time.time,
                          version: str = "dev") -> Tuple[str, Optional[dict]]:
    """One /token poll attempt: ("pending", None) to keep waiting at the same
    interval, ("slow_down", None) to keep waiting but the caller must
    increase its interval by at least SLOW_DOWN_INCREMENT_S (RFC 8628 sec.
    3.5), or ("done", tokens) once the user finished signing in. Raises
    SinkError on a real failure (denied, expired, misconfigured) or
    SinkTransportError if the host couldn't be reached at all -- the caller
    should treat the latter as transient, not as the sign-in having failed.
    The caller -- telemetry.py's login-poll -- makes one HTTP call per
    invocation rather than blocking a single tool call for the whole
    device-code expiry window."""
    http = _resolve_http(http)
    data = {"grant_type": DEVICE_CODE_GRANT, "client_id": cfg["client_id"], "device_code": start["device_code"]}
    status, resp = _post_form(_token_endpoint(cfg), data, cfg["http_timeout_s"], http, version)
    if status == 200:
        return "done", _tokens_from(resp, now())
    error = resp.get("error")
    if error == SLOW_DOWN_ERROR:
        return "slow_down", None
    if error == PENDING_ERROR:
        return "pending", None
    raise SinkError(resp.get("error_description") or resp.get("error") or "HTTP %d" % status)


def load_tokens(tdir: Path) -> Optional[dict]:
    try:
        tok = json.loads((tdir / TOKEN_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return tok if isinstance(tok, dict) and tok.get("access_token") else None


def save_tokens(tdir: Path, tokens: dict) -> None:
    path = tdir / TOKEN_FILE
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(tokens), encoding="utf-8")
    _chmod_private(tmp)
    tmp.replace(path)


def clear_tokens(tdir: Path) -> bool:
    path = tdir / TOKEN_FILE
    if path.exists():
        path.unlink()
        return True
    return False


def get_token(cfg: dict, tdir: Path, http: Optional[Http] = None, now=time.time, version: str = "dev") -> Optional[str]:
    """A usable access token, refreshed when within the skew of expiry;
    None when the user has never signed in or the refresh was explicitly
    refused. Raises SinkTransportError (not caught here) when the refresh
    couldn't even reach the server, so the caller can tell "not signed in"
    apart from "network unavailable, try again later" -- collapsing both
    into None previously made cmd_flush send a signed-out user to redo
    OAuth consent over what may have been a dropped connection."""
    http = _resolve_http(http)
    tok = load_tokens(tdir)
    if tok is None:
        return None
    if now() + cfg["token_refresh_skew_s"] < float(tok.get("expires_at", 0)):
        return tok["access_token"]
    if not tok.get("refresh_token"):
        return None
    data = {"grant_type": "refresh_token", "client_id": cfg["client_id"], "refresh_token": tok["refresh_token"],
            "scope": cfg["scope"]}
    status, resp = _post_form(_token_endpoint(cfg), data, cfg["http_timeout_s"], http, version)  # SinkTransportError propagates
    if status != 200:
        return None
    try:
        refreshed = _tokens_from(resp, now(), previous_refresh=tok["refresh_token"])
    except SinkError:
        return None
    save_tokens(tdir, refreshed)
    return refreshed["access_token"]


# ---------------------------------------------------------------------------
# sending
# ---------------------------------------------------------------------------
def to_fields(event: dict) -> dict:
    """One event as a SharePoint list item's `fields`: Title is the event id;
    every other event key is a column of the same name; lists join with
    commas. Column names in the list must match the event keys (the setup
    runbook in docs/TELEMETRY-FEEDBACK-MEMORY.md creates them)."""
    fields = {"Title": event.get("event_id")}
    for key, value in event.items():
        if key == "event_id":
            continue
        fields[key] = ",".join(str(v) for v in value) if isinstance(value, list) else value
    return fields


def _retry_after(headers: Dict[str, str]) -> Optional[int]:
    for name, value in headers.items():
        if name.lower() == "retry-after":
            try:
                return int(value)
            except ValueError:
                return None
    return None


def send(cfg: dict, token: str, events: List[dict], version: str = "dev", http: Optional[Http] = None) -> dict:
    """POST events one at a time, in order, stopping at the first failure so
    the caller can drop exactly the sent prefix from the queue.
    Returns {"sent": n, "stopped": reason or None, "retry_after_s": s or None}."""
    http = _resolve_http(http)
    headers = {"Authorization": "Bearer %s" % token, "Content-Type": "application/json", "Accept": "application/json",
               "User-Agent": USER_AGENT_FORMAT.format(version=version)}
    sent = 0
    for event in events:
        body = json.dumps({"fields": to_fields(event)}).encode("utf-8")
        try:
            status, resp_headers, text = http("POST", cfg["endpoint"], headers, body, cfg["http_timeout_s"])
        except SinkError as exc:
            return {"sent": sent, "stopped": str(exc), "retry_after_s": None}
        if status in (200, 201):
            sent += 1
            continue
        if status in (429, 503):
            return {"sent": sent, "stopped": "throttled (HTTP %d)" % status, "retry_after_s": _retry_after(resp_headers)}
        if status in (401, 403):
            return {"sent": sent, "stopped": "not authorized (HTTP %d); sign in again with /weave:setup" % status,
                    "retry_after_s": None}
        return {"sent": sent, "stopped": "HTTP %d from the usage list: %s" % (status, text[:200]), "retry_after_s": None}
    return {"sent": sent, "stopped": None, "retry_after_s": None}
