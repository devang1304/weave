---
name: document
description: "Builds a Netwoven-branded Word document on the 2026 document template from notes, transcripts, drafts, findings, or a prior deliverable: client-ready reports, assessments, readouts, recommendations, runbooks, specs, postmortems, and memos. Answer first, evidence grouped, unsupported claims flagged in chat, template leakage blocked by checks. Use when the user asks to write, rewrite, restructure, polish, formalize, or make client-ready a document, or mentions a Netwoven document or template. Also edits a document Weave built earlier. Not for slide decks, statements of work, proposals with fees, change orders, spreadsheets, emails, or chat messages."
when_to_use: "Trigger phrases: client-ready, assessment, readout, findings report, recommendation, runbook, SOP, spec, postmortem, decision memo, Word doc. Decks go to the deck skill; anything the client signs or prices goes to the sow skill; emails and Teams posts go to the voice-align plugin. Ask one question when client versus internal is unclear."
argument-hint: "[source files] [client] [title]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven document

Input: any rough material. Output: a Word document on the Netwoven 2026
document template, or structured Markdown for light internal notes, plus a
short flag list in chat. The substance is the user's; structure, prose,
visuals, and branding are this skill's job. Never invent facts, numbers,
commitments, or confidence the source does not contain.

Read `${CLAUDE_PLUGIN_ROOT}/references/routing.md` when the route is unclear and
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` before reading any source.

## 0. Check the environment (every run, silent when healthy)

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

If it reports anything missing, tell the user in one line to run
`/weave:setup` and stop. Never paste script output into chat.

## 1. Decide audience and weight, then say which you picked

| Audience | Weight | Output |
|---|---|---|
| Client or external reader | Branded client document | Confidentiality page kept, Company = client |
| Colleagues or leadership, formal artefact (spec, SOP, runbook, postmortem, assessment) | Branded internal document | No confidentiality page, Company = Netwoven |
| Colleagues, working note (memo, meeting notes, wiki page, under ~2 pages) | Structured Markdown | Skeleton in section 4b |

One line in the reply names the choice ("Built as a client document" /
"Built as an internal document" / "Kept as Markdown") and why. Unclear
client versus internal: ask one question with the two consequences.

**Editing a file Weave built earlier** (the user attaches a .docx and asks
for a change): skip to section 5.

## 2. Read everything

For .docx, .pptx, .xlsx, .pdf, .vtt sources:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" read SOURCE --out source.md --json
```

The Markdown carries provenance markers and the source's sensitivity label.
Apply `guardrails.md`: refuse credentials and sensitive personal data,
warn on internal candour and other clients' pricing, stop on restricted
labels or an `ai_permitted: false` client profile
(`weave.py client show <slug>`).

Collect the deliverable's identity while reading: **title**, **client
company name** (the template's Company field; never "Netwoven" on a client
document), **publish date or version**. Ask only for what is missing and
decision-critical, in one question.

## 3. Analyse before writing

Run `${CLAUDE_PLUGIN_ROOT}/references/writing-method.md`: reader, purpose,
substance inventory (load-bearing, context, filler, orphan, unsupported),
assumption check, then the recursive Minto rebuild. Produce the **pyramid
outline** (governing thought, key line, per-section points) and the flag
list before drafting. For more than about five pages of source, show the
outline and flags first and get a nod.

## 4a. Write the spec and build (branded document)

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write a spec file following `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` with
`weave.kind: "doc"` and `weave.recipe: "generic-doc"` (see
`${CLAUDE_PLUGIN_ROOT}/recipes/generic-doc.md` for the section pattern). Map the
pyramid: the governing thought plus SCQA becomes the executive summary as
the first level-1 heading; each key-line statement becomes a level-1 heading
that states the insight; recurse with levels 2 to 4; never skip a level.
Styles and mechanics come from `${CLAUDE_PLUGIN_ROOT}/references/word-template-guide.md`;
prose from `${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` (no em or en
dashes; colons only as list lead-ins; client-facing hygiene); tables and
figures from `${CLAUDE_PLUGIN_ROOT}/references/visual-standards.md`; colours only
from `${CLAUDE_PLUGIN_ROOT}/references/brand-tokens.md`. Set `meta.internal: true`
for internal documents. Fill `meta.author`/`meta.role` from the `me` profile
(`weave.py client show me`) when present.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --out "<Client>_<Type>_<YYYY-MM-DD>_v1.docx" --json
```

The result carries `validation.hard_fails`. Fix every hard failure in the
spec and rebuild; warnings need judgement. Then run the quick review and act
on anything high severity:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT.docx --quick --json
```

## 4b. Structured Markdown (working notes)

No template. Skeleton, adapted not padded:

```markdown
# <Title that states the governing thought>
**Status:** Draft | **Owner:** <name> | **Date:** <date> | **Audience:** <who>
> **Bottom line:** <the decision needed, the change, or the answer, in 1 to 2 sentences>
## Decisions needed / Actions
| # | Decision or action | Owner | By |
## <Key-line insight 1>
## <Key-line insight 2>
## Open questions and flags
```

Writing standards apply in full. Internal notes may keep system names,
ticket ids, and colleague names; the fidelity rule still holds.

## 5. Edit a document Weave built earlier

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" spec-from-file FILE.docx --out spec.json --json
```

If `drift.lost_if_rebuilt` is not empty, list those items to the user in
plain words and ask before continuing. Make the requested change in the
spec (text edits win over the stored plan; new sections get new ids), then
rebuild with `--replace` so the revision number increments, and re-check.
If the file has no Weave plan inside (built by hand or by another tool),
say so, offer to rebuild it on the template from its content, and proceed
only on a yes.

## 6. Deliver

Reply in this shape, at most eight lines, omitting empty headings:

**Check before sending** (from `${CLAUDE_PLUGIN_ROOT}/references/guardrails.md`,
shortened to what applies) · **I assumed** · **I left out** · **Open
decisions**. Then `Saved to: <absolute path>`.

Plain language only: "Word document", "checks passed", never script names,
flags, or JSON. The flag list lives in chat, never in the file.

## Hard rules

- Never let demo or instructional template content reach a file. The checks
  are a hard gate.
- Never invent or "improve" numbers, dates, prices, scope, or certainty.
  Hedges in the source are information. A gap is a flag, not a guess.
- The confidentiality page stays on client documents; its wording comes
  from the template. Company is the client's name on client documents.
- The templates' look is fixed configuration: no restyling, recolouring, or
  new styles. The guides' quirks exist so you reproduce them exactly.
- Internal candour is allowed in internal documents; recklessness is not.
  Flag anything that would be a problem in writing even internally.
