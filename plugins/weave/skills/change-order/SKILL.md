---
name: change-order
description: "Drafts a signable Netwoven change order amending an already-signed SOW on the document template: summary of change, scope impact, schedule impact, and cost impact, each stated against the original engagement rather than restating it. Fills these from a change request, email thread, redline, or scoping note; names the original SOW being amended; never invents a cost or schedule delta; keeps the document short by referencing the original SOW instead of duplicating its scope or fee tables. Use when the user asks for a change order, amendment, scope change, contract modification, or a cost or schedule adjustment on an engagement that already has a signed SOW. Not for a brand-new SOW (that's the sow skill), a proposal without a prior signed engagement, or an internal document."
disable-model-invocation: true
argument-hint: "[source files] [original SOW title]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven change order

Input: a change request, email thread, redline, or scoping note describing
a modification to an engagement that already has a signed SOW. Output: a
short amendment on the Netwoven document template naming the original SOW
and stating what changed in scope, schedule, and cost, plus a flag list in
chat. This is not a new SOW: it references the original document rather
than restating its scope or fee tables.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 2. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. `original_sow` is the one
required input the recipe can't reliably infer -- if the source doesn't
name the original SOW clearly, ask for its title (and date if known) in one
question. Never invent a cost or schedule delta: a gap becomes
`[TO BE PROVIDED]` in the relevant section and a line in the flag list.

## 3. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` and
`${CLAUDE_PLUGIN_ROOT}/recipes/change-order.md`, with `weave.recipe` set to
`change-order`.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe change-order --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

A required section missing from the spec is a hard build failure: check
the recipe's `sections` list and rebuild. `validation.warnings` go in the
flag list, not the error path.

## 4. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`.

## 5. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path>`.
