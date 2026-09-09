---
name: sow
description: "Drafts a signable Netwoven Statement of Work on the 2026 Milestone (fixed-fee) or Time and Materials template, choosing the variant from the fee model in the source and asking one question when unclear. Fills scope, deliverables, assumptions, roles, fees, timeline, and Microsoft products from scoping notes, an estimate, a proposal, or a discovery call; keeps legal and commercial boilerplate verbatim; never invents rates, hours, or amounts; checks that no placeholder or former client name survives. Use when the user asks for a SOW, statement of work, scope with fees, rate card, engagement letter, fixed-fee or T&M proposal, or the scope document the client will sign. Also edits a SOW Weave built earlier. Not for change orders on a signed engagement, proposals without fees, reviewing a counterparty's SOW, or emails."
when_to_use: "Trigger phrases: SOW, statement of work, scope with fees, fixed fee, milestones, T&M, rate card, weekly hours, engagement letter, signable. Change orders go to the change-order command; proposals without fees and decks go to the document or deck skill; emails go to the voice-align plugin. Ask before treating a small edit as new SOW work."
argument-hint: "[source files] [client legal name] [project]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven statement of work

Input: scoping notes, an estimate, a proposal, a call transcript, or an
earlier SOW draft. Output: a signable SOW on the Netwoven 2026 Milestone or
T&M template with the legal and commercial skeleton intact and every
project-specific section filled from the source, plus a flag list in chat.
The reader signs this document, so fidelity to the source's numbers, scope,
and commitments outranks everything else.

Read `${CLAUDE_PLUGIN_ROOT}/references/routing.md` when the route is unclear and
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` before reading any source.

## 0. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 1. Pick the variant

| Signals in the source | Variant |
|---|---|
| Fixed fee, milestone payments, deliverable-based acceptance, "not to exceed" tied to deliverables, payment on completion | **Milestone** |
| Hourly or daily rates, rate card, weekly hours, number of weeks, monthly invoicing, "estimate", "actuals", staffing plan by role | **T&M** |
| Mixed, absent, or contradictory | **Ask one question** offering both, one-line consequence each |

Never infer the fee model from the client's industry or from habit. A
change order or amendment on a signed engagement is not a new SOW: say so,
and point to `/weave:change-order`. A small edit to an existing SOW (a
typo, a date, one paragraph): section 5.

## 2. Read everything and collect the inputs

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" read SOURCE --out source.md --tables-json tables.json --json
```

Apply `guardrails.md`. Collect from the source; ask only for
decision-critical gaps, in one question:

- Client legal name; project name; effective date (a real date)
- Scope of work and out of scope
- Deliverables (Milestone: scope section, milestone, deliverable,
  description, format, in that order; T&M: scope section, deliverable,
  description, format, no milestone column)
- Project-specific assumptions beyond the template's standard set
- Netwoven roles (onshore or offshore) and the client roles that apply
- Fees (Milestone: amount and expected completion per milestone; T&M: role,
  location, resources, weekly hours, weeks, hourly rate)
- Timeline phases; Microsoft products in scope; whether CPOR and Govern 365
  sections apply; expenses cap if travel is expected; client logo file if any

Missing fees or rates are never invented. Leave them out of the spec so the
cell stays blank, write `[TO BE PROVIDED]` where prose needs a number, and
list each gap under Open decisions. Check `weave.py client show <slug>` for
a profile: legal name, logo path, standard assumptions, `ai_permitted`.

## 3. Analyse

Run `${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` with the purpose fixed:
the reader signs. The governing thought becomes the first paragraph of the
Executive Summary (what Netwoven will deliver, for whom, by when, under which
commercial model). Inventory every scope item, deliverable, assumption,
number, and date; anything unsupported goes to the flag list.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write a spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` with `weave.kind: "sow"`,
`weave.recipe: "sow"`, `sow.variant` set, the fee, deliverable, timeline,
product, and business-application tables in the `sow` keys (shapes in
`${CLAUDE_PLUGIN_ROOT}/recipes/sow.md`), and the prose in `sow.sections`:
executive summary (governing thought first; the effective-date sentence stays
bound), scope as bullets, out of scope stated explicitly, project-specific
assumptions to add, glossary terms the source defines, roles notes,
signature names and titles when given (never sign for anyone). Every style,
table, field, and placeholder decision follows
`${CLAUDE_PLUGIN_ROOT}/references/sow-template-guide.md`: Heading 1 to 3 only,
never NW Heading or Heading 4. Prose follows
`${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` in the SOW register.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --out "<Client>_SOW_<YYYY-MM-DD>_v1.docx" --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT.docx --rubric sow --quick --json
```

Leftover placeholders, highlights, the internal Appendix, the rate-card
link, or broken fee fields are hard failures: fix the spec and rebuild.

## 5. Edit a SOW Weave built earlier

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" spec-from-file FILE.docx --out spec.json --json
```

Show any `drift.lost_if_rebuilt` items and ask before continuing. Fee
tables are edited through the `sow` keys, never by retyping totals. Rebuild
with `--replace`, re-check.

## 6. Deliver

At most eight lines: **Check before sending** (legal name, effective date,
every amount and rate, the sensitivity label, named reviewer) · **I
assumed** · **I left out** · **Open decisions** (fee or rate gaps, optional
sections dropped, boilerplate the user asked to change), then
`Saved to: <absolute path>`. Plain language only.

## Hard rules

- Legal and commercial boilerplate stays verbatim unless the user edits it
  explicitly: Change Requests, Status Reporting, deemed acceptance and late
  fee (Milestone), invoicing and pause-or-stop terms (T&M), Staffing, the
  Signatures preamble, the Statement of Confidentiality.
- Never invent, round, or "improve" rates, discounts, hours, amounts, or
  dates. A gap is a flag, not a guess, and never a zero.
- The rate-card link, the internal Appendix, and the template's former
  client references never reach the output.
- No yellow highlight, `<<...>>` prompt, or `MM/DD/YYYY` may survive.
- The client logo box is filled from a supplied file or removed; never draw
  or approximate a logo.
- The templates' look is fixed configuration: no restyling, no new table
  styles, no header or footer edits.
