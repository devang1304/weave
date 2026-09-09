---
name: setup
description: "Sets up Netwoven Weave on this device: checks and installs what it needs, records your name and role for revision pages, adds or updates client profiles (legal name, logo, standard assumptions, AI-permitted flag), and shows the usage-data disclosure and how to turn it off. Run once after installing, or whenever Weave says something is missing."
disable-model-invocation: true
argument-hint: "[client <name>]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Weave setup

Plain language throughout. The user may never have installed anything on a
command line; you are doing it for them.

## 1. Check, then install if needed

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --json
```

If `ok` is true: "Weave is ready." Otherwise run the installer and report
the result in one sentence:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --install --json
```

If the install cannot reach the internet, the result carries a `hint` with
the exact command an administrator can run; show it verbatim and stop.
Never show tracebacks.

## 2. Your name and role

Ask once, in one question: "What name and role should appear on the
revision page of documents you build?" Store them:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" client set me author_name="Jane Doe" author_role="Senior Consultant" --json
```

Mention that the same values can also be set in the plugin's settings; the
profile set here wins whenever both exist, since it is the more specific,
easily-updated value.

## 3. Client profiles (optional, or when invoked as `/weave:setup client <name>`)

A profile saves re-typing on every deliverable and encodes engagement
terms. Ask for: legal name, short name, logo file path (optional),
standard assumptions to include in SOWs (optional), whether the client's
engagement terms permit AI-assisted drafting (default yes), default
sensitivity label (optional).

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" client set contoso legal_name="Contoso Ltd" short_name="Contoso" ai_permitted=true --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" client list --json
```

Profiles live only on this device, in a folder only the user can read.
They are never sent anywhere.

## 4. Usage data disclosure (show once, verbatim)

> Weave records anonymised usage events: which kind of deliverable was
> built, which template version, whether the checks passed, and how long it
> took. It never records what you typed, file names, or client names. This
> helps the plugin team see what works. Turn it off any time in the plugin
> settings ("Share anonymised usage events") or by setting
> `WEAVE_TELEMETRY=0`.

Point to `${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` for the full rules.

## 5. Usage list sign-in (only when the plugin team has configured it)

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" telemetry status --json
```

If `sink_configured` is true and `signed_in` is false, start the sign-in:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" telemetry login-start --json
```

Show the `message` field exactly as printed (a web address and a short
code) and tell the user to sign in there. Then poll:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" telemetry login-poll --json
```

Each call already waits the server-paced interval before checking, so if
`signed_in` is false just call it again (no separate wait step, and never
call it back-to-back in a tight loop). Once `signed_in` is true, one
sentence: usage events now leave this device at the start of each session.
If `sink_configured` is false, skip this step: events stay on the device. If the plugin team gave you an endpoint, tenant id and
client id to enter by hand (surfaces without plugin settings), record them
first:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" telemetry configure --endpoint URL --tenant-id ID --client-id ID --json
```

## 6. Finish

Three sentences to try, tailored to the user's role if known:

- "Turn these workshop notes into a client-ready assessment for Contoso."
- "Make a 10-slide QBR deck from this status email."
- "Review this deck before I send it."
