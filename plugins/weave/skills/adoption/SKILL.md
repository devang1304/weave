---
name: adoption
description: "Builds an adoption strategy document, a communications plan document, or a training deck for a Microsoft rollout on the Netwoven templates -- picking the one the request actually needs rather than building all three. Fills Current State/Adoption Goals/Approach/Success Metrics for a strategy, Audience/Key Messages/Channels and Cadence for a comms plan, or Objectives/Agenda/Key Topics for a training deck from notes, transcripts, or a prior draft; never invents an adoption percentage, a communication date, or a feature the source doesn't describe. Use when the user asks to drive adoption, plan rollout communications, or build end-user training for a Microsoft rollout. Not for a kickoff charter, a status report, or a proposal."
disable-model-invocation: true
argument-hint: "[source files] [what's being adopted]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Adoption (strategy, comms, or training)

Input: notes, a rollout plan, a communications draft, or training material.
Output: exactly one of an adoption strategy document, a communications plan
document, or a training deck, on the Netwoven templates, plus a short flag
list in chat. These are three independent recipes, not a bundle built
together -- a rollout plan, its comms, and its training session are usually
written at different times by different people, so build the one the
request actually needs.

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 2. Pick the recipe

| Source or request is about | Recipe | Read |
|---|---|---|
| The overall rollout plan -- current state, goals, tactics, metrics | `adoption-strategy` | `${CLAUDE_PLUGIN_ROOT}/recipes/adoption-strategy.md` |
| Who hears what, through which channel, when | `adoption-comms` | `${CLAUDE_PLUGIN_ROOT}/recipes/adoption-comms.md` |
| An end-user training session | `adoption-training` | recipes/adoption-training.md, when present (see note) |

If the request genuinely covers more than one of these, ask which one to
build first rather than guessing -- do not build two or three in the same
run on the assumption they were wanted together.

Note: `adoption-training` is a deck-kind recipe, distinct from the two
doc-kind recipes above -- on some packaging surfaces the build for it may
fail outright with a clear "missing nw_pptx_helpers.py" error rather than
run. See `${CLAUDE_PLUGIN_ROOT}/references/routing.md` for when that's expected.

## 3. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. Ask the one `required` input the
chosen recipe lists (`initiative` -- what is being adopted or communicated
about or trained on) if the source doesn't make it obvious. Never invent an
adoption baseline, a target percentage, a communication date or channel, or
a feature the source doesn't describe as real; a gap is a flag, not a
guess.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` and the chosen
recipe's `.md`, with `weave.recipe` set to `adoption-strategy`,
`adoption-comms`, or `adoption-training` (exactly one). Follow
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the prose register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe adoption-strategy --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

(Use `--recipe adoption-comms` or `--recipe adoption-training` instead when
that's the one chosen.) A required section missing from the spec is a hard
build failure: check the recipe's `sections` list and rebuild.
`validation.warnings` go in the flag list, not the error path.

## 5. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`.

## 6. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path>`.
