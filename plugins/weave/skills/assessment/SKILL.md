---
name: assessment
description: "Builds a client or internal assessment report (current-state findings, recommendations, roadmap) on the Netwoven document template, optionally paired with a readout deck built from the same facts. Fills area-by-area current/target state, priority, and sequenced next steps from tenant inventories, discovery notes, or prior assessments; never invents a maturity or readiness rating. Use when the user asks to assess a current state, produce a maturity or readiness assessment, or build an assessment report and/or readout deck. Not for a proposal or SOW, a full HLD, or an ongoing status report."
disable-model-invocation: true
argument-hint: "[source files] [what's being assessed]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Assessment report / readout deck

Recipes: `${CLAUDE_PLUGIN_ROOT}/recipes/assessment-report.md` (doc,
`assessment-report`) and `${CLAUDE_PLUGIN_ROOT}/recipes/assessment-readout.md`
(deck, `assessment-readout`) -- read whichever recipe applies before writing
the spec. They are a companion pair: build the report first when both are
wanted, and reuse its client/date/scope answers for the deck rather than
asking twice.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

## 2. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. `scope` is the only input the
recipes require and can't infer from most source material -- ask for it in
one question if missing, offering the source's own phrasing as the default.
Never invent a maturity or readiness rating: a gap becomes `[TO BE PROVIDED]`
in the findings table and a line in the flag list.

## 3. Write the spec(s) and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write each spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md`, with `weave.recipe`
set to `assessment-report` and/or `assessment-readout`. Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe assessment-report --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec deck-spec.json --recipe assessment-readout --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

A required section missing from the spec is a hard build failure: check the
recipe's `sections` list and rebuild. `validation.warnings` (a missing
heading, a rating with no supporting finding) go in the flag list, not the
error path.

## 4. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`. Keep both specs' shared fields (client, date, scope) in sync by
hand if only one side is edited.

## 5. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path(s)>`.
