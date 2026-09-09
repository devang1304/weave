---
name: proposal
description: "Builds a pre-sales Netwoven proposal (executive summary, understanding of the client's needs, proposed approach, optional why-Netwoven, rough-order-of-magnitude investment summary, next steps) on the Netwoven document template. Restates the client's problem in their own language and names the recommended Microsoft offering from scoping notes, a discovery call, or an RFP; investment figures are always ranges, never a fabricated number, and no signed-commitment language leaks in. Use when the user asks for a proposal, pitch document, offering write-up, pre-sales narrative, or a preliminary investment estimate. Not for a signed SOW or fee-bearing scope, a full assessment or discovery readout, or an internal design doc."
disable-model-invocation: true
argument-hint: "[source files] [offering]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven proposal

Recipe: `${CLAUDE_PLUGIN_ROOT}/recipes/proposal.md` (doc, `proposal`) -- read it
before writing the spec. Pre-sales only: this document is never a signed
commitment (see `${CLAUDE_PLUGIN_ROOT}/references/routing.md` for that boundary).

## 1. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

## 2. Collect the inputs

Read the source(s) with `weave.py read`. Apply
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`. `offering` is the only input the
recipe requires and can't reliably be inferred from most source material --
ask for it in one question if missing (which Microsoft offering or
solution this proposal is for). Never invent an investment figure: a gap
becomes `[TO BE PROVIDED]` in the table and a line in the flag list.

## 3. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write the spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md`, with `weave.recipe`
set to `proposal`. Follow `${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for
the prose register; the Investment Summary states plainly that figures are
rough order of magnitude, never final pricing.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --recipe proposal --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT --quick --json
```

A required section missing from the spec is a hard build failure: check the
recipe's `sections` list and rebuild. `validation.warnings` (a missing
heading, a dropped optional section) go in the flag list, not the error
path.

## 4. Edit or rebuild

`weave.py spec-from-file FILE --out spec.json`, edit, rebuild with
`--replace`.

## 5. Deliver

**Check before sending** · **I assumed** · **I left out** · **Open
decisions**, then `Saved to: <absolute path>`.
