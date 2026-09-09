---
name: status
description: "Builds a client or internal status report (RAG table, risks, next steps) on the Netwoven document template, optionally paired with a steering-committee deck built from the same facts. Fills workstream status, this-period/next-period notes, and owners-and-dates for next steps from meeting notes, a tracker export, or prior status reports; never invents a RAG colour. Use when the user asks for a status report, steerco deck, weekly or monthly update, project status, or RAG/RAID summary. Not for a one-off proposal, a kickoff deck, or a full assessment readout."
disable-model-invocation: true
argument-hint: "[source files] [reporting period]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Status report / steerco deck

Recipes: `${CLAUDE_PLUGIN_ROOT}/recipes/status-report.md` (doc, `status-report`) and
`${CLAUDE_PLUGIN_ROOT}/recipes/steerco.md` (deck, `steerco`) -- read whichever
recipe applies before writing the spec. They are a companion pair: build the
report first when both are wanted, and reuse its client/date/period answers
for the deck rather than asking twice.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

## 2. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. `period` is the only input the
recipes require and can't infer from most source material -- ask for it in
one question if missing, offering the source's own date range as the
default phrasing. Never invent a workstream's RAG status: a gap becomes
`[TO BE PROVIDED]` in the table and a line in the flag list.

## 3. Write the spec(s) and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write each spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md`, with `weave.recipe`
set to `status-report` and/or `steerco`. Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe status-report --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec deck-spec.json --recipe steerco --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

A required section missing from the spec is a hard build failure: check the
recipe's `sections` list and rebuild. `validation.warnings` (a missing
heading, no owner/date found in Next period) go in the flag list, not the
error path.

## 4. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`. Keep both specs' shared fields (client, date, period) in sync
by hand if only one side is edited.

## 5. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path(s)>`.
