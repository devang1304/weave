---
name: design
description: "Builds an Architecture Decision Record (ADR) for a single architecture decision, or a High-Level Design (HLD) document for a whole system or solution, on the Netwoven document template -- picking the right shape from the source or request rather than assuming. Fills Context/Decision/Consequences for an ADR or Overview/Requirements/Architecture/Components for an HLD from notes, transcripts, or a prior draft; never invents a decision the source hasn't made, a component it doesn't describe, or a tradeoff-free consequence. Use when the user asks to record, document, or write up an architecture decision, ADR, design decision, high-level design, HLD, system design, or solution architecture. Not for a full runbook, a kickoff charter, or a priced proposal."
disable-model-invocation: true
argument-hint: "[source files] [decision or system name]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Architecture design (ADR or HLD)

Input: notes, a design discussion, a transcript, or a prior draft. Output:
either an ADR (one decision) or an HLD (one system or solution), on the
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
| One specific choice already being made, with reasoning for and against | `adr` | `${CLAUDE_PLUGIN_ROOT}/recipes/adr.md` |
| A whole system or solution to design (multiple components, requirements) | `hld` | `${CLAUDE_PLUGIN_ROOT}/recipes/hld.md` |

If the source genuinely supports either reading, ask one question: "Is this
recording one decision, or designing the whole solution?" Do not guess past
that -- an ADR drafted for an undecided question, or an HLD drafted for a
single choice, is the wrong shape either way.

## 3. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. Ask the one `required` input the
chosen recipe lists (`decision_title` for `adr`, `system_name` for `hld`) if
the source doesn't make it obvious. Never invent a decision the source
hasn't made (`adr`) or a component it doesn't describe (`hld`); a gap is a
flag, not a guess.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` and the chosen recipe's
`.md`, with `weave.recipe` set to `adr` or `hld` (never both). Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe adr --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

(Use `--recipe hld` instead when that's the one chosen.) A required section
missing from the spec is a hard build failure: check the recipe's
`sections` list and rebuild. `validation.warnings` go in the flag list, not
the error path.

## 5. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`.

## 6. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path>`.
