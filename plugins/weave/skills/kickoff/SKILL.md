---
name: kickoff
description: "Builds a kickoff deck and/or project charter for a new engagement on the Netwoven templates, built from the same underlying facts. Fills objectives, team roles and names, milestone timeline, and first next-steps with owners from a SOW, proposal, or kickoff-call notes; never invents a stakeholder name, date, or milestone the source doesn't give. Use when the user asks to kick off a project, introduce the team to a client sponsor, or start an engagement. Not for an ongoing status update (that's `status`), a proposal, or a SOW."
disable-model-invocation: true
argument-hint: "[source files] [project name]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Kickoff deck / project charter

Recipes: `${CLAUDE_PLUGIN_ROOT}/recipes/project-charter.md` (doc, `project-charter`)
and `${CLAUDE_PLUGIN_ROOT}/recipes/kickoff-deck.md` (deck, `kickoff-deck`) -- read
whichever recipe applies before writing the spec. They are a companion pair:
build the charter first when both are wanted, and reuse its client/date/
project_name answers for the deck rather than asking twice.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

## 2. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. `project_name` is the only input
the recipes require and can't infer from most source material -- ask for it
in one question if missing. Never invent a stakeholder name, milestone
date, or success metric: a gap becomes `[TO BE PROVIDED]` in the block and a
line in the flag list.

## 3. Write the spec(s) and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write each spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md`, with `weave.recipe`
set to `project-charter` and/or `kickoff-deck`. Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe project-charter --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec deck-spec.json --recipe kickoff-deck --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

A required section missing from the spec is a hard build failure: check the
recipe's `sections` list and rebuild. `validation.warnings` (a missing
heading, no owner found in Next Steps) go in the flag list, not the error
path.

## 4. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`. Keep both specs' shared fields (client, date, project_name) in
sync by hand if only one side is edited.

## 5. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path(s)>`.
