---
name: feedback
description: "Reports a problem with Netwoven Weave or rates a result: asks what happened and what you expected, optionally a 1 to 5 rating, and files it for the plugin team without sending your prompt, file, or client name. Use after a build or review that missed the mark, or to suggest a deliverable type Weave should support."
disable-model-invocation: true
argument-hint: "[what happened]"
allowed-tools: Read, Glob, Write, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Weave feedback

One question, then file it. Never include prompt text, file contents, file
names, or client names in what is filed; describe the problem in general
terms and keep any specifics in the user's own copy.

## 1. Ask once

"What happened, and what did you expect instead? Optionally, rate the
result 1 to 5." If the user already said what happened in the invocation,
skip the question and only ask for the rating if they want to give one.

## 2. File it

Try the telemetry channel first (it exists from Phase 2):

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" telemetry event --type feedback --skill <skill> --rating <n> --text "<general description>" --json
```

If that reports `ok: true`: the report is stored on the device and sent to
Netwoven's usage list at the start of the next session once the user is
signed in (`/weave:setup` does that when the list is configured). If
`weave.py telemetry status --json` shows `sink_configured` false, tell the
user the report stays on this device until an administrator configures the
usage list -- it is already filed, no further action needed.

If the command itself fails (telemetry is not installed on this packaging
surface), nothing was recorded anywhere. Run
`python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --json` (a separate script,
available even when telemetry.py is not) and read its `data_dir` field --
that folder always exists and is writable once setup has run. Write the
feedback yourself with the Write tool instead, as
`<data_dir>/feedback/feedback-<YYYYMMDD-HHMM>.json` (create the folder if
needed) containing `{"skill": ..., "rating": ..., "text": ..., "plugin_version":
..., "surface": ...}` (the same fields a recorded event carries; get
`plugin_version` from `.claude-plugin/plugin.json` next to `${CLAUDE_PLUGIN_ROOT}`
and `surface` from whether this is Claude Code or Claude Cowork), then tell
the user it is saved at that path and to paste it into the Weave Teams
channel until the automatic channel is live.

## 3. Reply

Two lines: what was filed (without repeating the text back) and where it
went. Thank the user without ceremony. If the feedback asks for a new
deliverable type, say it will be reviewed against the recipe list and that
new types usually become recipes rather than new commands.

Rules in `${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` apply to the filed
content.
