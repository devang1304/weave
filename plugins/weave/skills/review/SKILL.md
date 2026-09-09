---
name: review
description: "Reviews a finished Netwoven deliverable (Word document or PowerPoint deck) before it goes to a client and returns a scored checklist: template integrity and leakage, client name and date fields, sensitivity label, house writing standards, action titles and slide density, unsupported numbers, internal candour, and brand palette. Fixes mechanical issues on request and lists judgement calls for the author. Use when the user asks to review, check, QA, proofread, or sanity-check a document or deck, asks whether it is ready to send, or attaches a Netwoven file and asks what is wrong with it. Not for writing or rewriting content, legal review of contracts, or files not on Netwoven templates, where it offers a lighter check."
when_to_use: "Trigger phrases: review this, check this, QA, proofread, ready to send, client-ready check, what is wrong with this deck, record the reviewer. Requests to write or rewrite go to the document, deck, or sow skill; emails go to the voice-align plugin."
argument-hint: "[file] [client name]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven deliverable review

Input: a .docx or .pptx, built by Weave or not. Output: a verdict, a score by
area, the storyline as the titles alone tell it, and a fix list split into
"fixed for you" and "your call". Nothing in the file changes unless the
user asks. Read `${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` before opening
any file.

## 0. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 1. Run the checks

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review FILE --client "Client, Inc." --report review.md --json
```

Use `--rubric sow` for statements of work. `--quick` when the user only
wants a fast leakage and hygiene pass. If the file was not built on a
Netwoven template the result says so; offer the lighter check (writing,
structure, numbers) and continue only on a yes.

## 2. Add the judgement layer

The checks extract the storyline (slide titles or section headings in
order). Read it as a reader would and grade five things, citing the title
or heading each finding refers to:

1. **Governing thought**: does the first content slide or the executive
   summary state the answer, not the topic?
2. **Key line**: are the top-level sections or slides MECE, and do they
   together prove the governing thought?
3. **Insight titles**: which titles are topic labels ("Findings",
   "Timeline") rather than sentences with a verb?
4. **Audience fit**: internal candour, jargon, or system names a client
   should not see; hedges that were hardened; commitments not in the source.
5. **Unsupported claims**: numbers without a source line, "best practice"
   assertions, superlatives.

Standards to grade against: `${CLAUDE_PLUGIN_ROOT}/references/deck-storyline.md`
(action titles, horizontal and vertical logic, density),
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` (house punctuation,
client-facing hygiene), `${CLAUDE_PLUGIN_ROOT}/references/brand-tokens.md`
(palette), `${CLAUDE_PLUGIN_ROOT}/references/visual-standards.md` (chart choice,
source lines).

## 3. Fix only what is safe, and only when asked

Ask once: "Fix the mechanical items now?" (dashes, empty placeholders,
closers in the wrong place, missing alt text, creator metadata, empty
title). On yes:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review FILE --apply --json
```

A backup is written first. Never rewrite substance, never touch fee tables,
never change the sensitivity label. Content fixes (a topic title, an
unsupported number) are proposed as edits for the author.

When the user names the reviewer of record ("I reviewed it" or "Jane
reviewed this"), record it:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review FILE --apply --reviewer "Jane Doe" --json
```

## 4. Deliver

Reply shape, plain language, no scores below the verdict line unless asked:

- **Verdict**: ready to send after the checklist / fix the listed items /
  not ready. Score out of 100 in parentheses.
- **The story your titles tell**: the ordered titles, 1 line each, with a
  mark on the ones that are topics not insights.
- **Fixed for you** (only if `--apply` ran).
- **Your call**: findings by severity, each with the location and a
  one-line fix.
- **Before you send**: the items from
  `${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` that apply.

Never paste JSON or script output. Say "checks", not "validator".

## Hard rules

- Nothing changes in the file without an explicit yes; a backup always
  exists before `--apply`.
- Fee tables, legal boilerplate, and Purview labels are never edited by
  this skill.
- Findings cite a location. No finding without a fix or a decision.
- A file from another firm's template is reviewed for writing and
  structure only; do not restyle it onto the Netwoven template unless asked
  (that is a build, not a review).
