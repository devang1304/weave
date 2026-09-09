---
name: runbook
description: "Builds a migration runbook for a broad end-to-end system, workload, or tenant migration, or a cutover runbook for a focused go-live cutover window, on the Netwoven document template -- picking the right shape from the source or request rather than assuming. Fills the pre-migration checklist, numbered migration steps with owners and dates, rollback triggers, and validation checks for a migration runbook, or the go/no-go checklist, time-boxed cutover sequence with owners, abort criteria, and post-cutover validation for a cutover runbook, from notes, transcripts, or a prior draft; never invents a step, owner, date, or abort threshold the source doesn't support. Use when the user asks to plan a migration, write a migration runbook, plan a go-live or cutover, or write a cutover runbook or cutover plan. Not for a project charter, a kickoff deck, a full assessment readout, or a design document."
disable-model-invocation: true
argument-hint: "[source files] [system or workload]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Migration or cutover runbook

Input: notes, a migration plan discussion, a transcript, or a prior draft.
Output: either a migration runbook (the whole migration, prerequisites
through validation) or a cutover runbook (just the go-live window), on the
Netwoven document template, plus a short flag list in chat. These are two
independent recipes, not a companion pair -- build exactly one per run.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 2. Pick the recipe

| Source shows | Recipe | Read |
|---|---|---|
| A broad end-to-end migration plan (prerequisites, the full move, rollback, validation) | `migration-runbook` | `${CLAUDE_PLUGIN_ROOT}/recipes/migration-runbook.md` |
| A focused go-live cutover-window procedure (a time-boxed sequence around one cutover event) | `cutover-runbook` | `${CLAUDE_PLUGIN_ROOT}/recipes/cutover-runbook.md` |

If the source genuinely supports either reading, ask one question: "Is
this the whole migration plan, or just the cutover window?" Do not guess
past that -- a cutover runbook drafted for a whole migration, or a
migration runbook drafted for a single go-live window, is the wrong shape
either way.

## 3. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. Ask the one `required` input the
chosen recipe lists ("what is being migrated" for `migration-runbook`,
"what is cutting over, and to what target" for `cutover-runbook`) if the
source doesn't make it obvious. Never invent a step, owner, date, or abort
threshold the source doesn't support -- a gap is a flag, not a guess.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` and the chosen
recipe's `.md`, with `weave.recipe` set to `migration-runbook` or
`cutover-runbook` (never both). Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe migration-runbook --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

(Use `--recipe cutover-runbook` instead when that's the one chosen.) A
required section missing from the spec is a hard build failure: check the
recipe's `sections` list and rebuild. `validation.warnings` -- including
the `owners_and_dates_in` check flagging a step with no clear owner or
date -- go in the flag list, not the error path.

## 5. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`.

## 6. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path>`.
