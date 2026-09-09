---
name: deck
description: "Builds a Netwoven-branded PowerPoint deck on the 2026 presentation template: QBRs, steering committee readouts, kickoff decks, findings presentations, roadmaps, executive briefings, and proposal decks without fees. Turns notes, transcripts, or a document into a consulting storyline: one message per slide, full-sentence action titles, charts chosen by message, brand palette, confidentiality slide, then checks the file. Use when the user asks for slides, a deck, a presentation, a readout, or a QBR, or when the material is presentation-shaped (an agenda, a talk track, a meeting on a date). Also edits a deck Weave built earlier. Not for Word documents, statements of work, spreadsheets, emails, or chat messages."
when_to_use: "Trigger phrases: slides, deck, presentation, QBR, steerco, readout, kickoff, roadmap deck, exec briefing, pitch. Word documents go to the document skill; SOWs and fee proposals to the sow skill; emails and Teams posts to the voice-align plugin. Ask one question when document versus deck is unclear."
argument-hint: "[source files] [client] [title]"
allowed-tools: Read, Glob, Bash(python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" *)
---

# Netwoven deck

Input: any rough material. Output: a PowerPoint deck on the Netwoven 2026
presentation template plus a short flag list in chat. One message per slide,
every content-slide title a full-sentence insight. Never invent facts,
numbers, commitments, or confidence the source does not contain.

Read `${CLAUDE_PLUGIN_ROOT}/references/routing.md` when the route is unclear and
`${CLAUDE_PLUGIN_ROOT}/references/guardrails.md` before reading any source.

## 0. Check the environment

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" doctor --quiet
```

Anything missing: one line telling the user to run `/weave:setup`, then stop.

## 1. Route and identity

A deck is the right output when the user asks for slides or the material is
presentation-shaped (agenda, talk track, QBR, steerco, pitch, a meeting on a
date). A 20-page findings report and a 12-slide readout are different
deliverables; when torn, ask one question. Internal audience: still a deck,
with `meta.internal: true` (no confidentiality slide).

Collect **deck title**, **subtitle** (usually "Client | Month Year"),
**client name**, and audience while reading. Ask only for what is missing
and decision-critical. **Editing a deck Weave built earlier**: go to section 5.

## 2. Read everything

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" read SOURCE --out source.md --tables-json tables.json --json
```

Apply `guardrails.md` (refuse, warn, stop rules; sensitivity label of the
source). Numeric tables from `--tables-json` feed charts later.

## 3. Analyse, then write the ghost deck

Run `${CLAUDE_PLUGIN_ROOT}/references/writing-method.md` for the pyramid, then
`${CLAUDE_PLUGIN_ROOT}/references/deck-storyline.md` for the **ghost deck**: the
ordered list of action titles, one full-sentence insight per slide, before
any body content. Check horizontal logic (titles alone tell the story, MECE
siblings) and vertical logic (each body proves its title). For more than
about five pages of source, show the ghost deck and flags first and get a nod.

## 4. Write the spec and build

Run `python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" memory show` first and apply any saved preferences that do not conflict with the source or the recipe; they are habits learned from this user's own hand edits.

Write a spec per `${CLAUDE_PLUGIN_ROOT}/recipes/SPEC.md` with `weave.kind: "deck"`
and `weave.recipe: "generic-deck"` (pattern in
`${CLAUDE_PLUGIN_ROOT}/recipes/generic-deck.md`). Slide-type to layout map, density
rules, and anti-patterns are in `deck-storyline.md`; placeholder details in
`${CLAUDE_PLUGIN_ROOT}/references/ppt-template-guide.md`; chart and framework
choice in `${CLAUDE_PLUGIN_ROOT}/references/visual-standards.md`; colours only from
`${CLAUDE_PLUGIN_ROOT}/references/brand-tokens.md` (the builder applies the palette;
never the theme's default cycle). Each evidence slide carries one visual
and a source line when it shows numbers. End on next steps with owners and
dates, never a "Questions?" slide. The builder adds the cover, the Statement
of Confidentiality slide (client decks), and keeps Thank You and Closing
last.

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" build --spec spec.json --out "<Client>_<Type>_<YYYY-MM-DD>_v1.pptx" --json
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" review OUT.pptx --quick --json
```

Fix every hard failure and every high-severity finding, then rebuild.

## 5. Edit a deck Weave built earlier

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/weave.py" spec-from-file FILE.pptx --out spec.json --json
```

List any `drift.lost_if_rebuilt` items in plain words and ask before
continuing. Change the spec (slide text edits win; new slides get new ids),
rebuild with `--replace`, re-check. A deck without a Weave plan inside: say
so and offer to rebuild it on the template from its content.

## 6. Deliver

At most eight lines: **Check before sending** · **I assumed** · **I left
out** · **Open decisions**, then `Saved to: <absolute path>`. Plain language
only ("PowerPoint deck", "checks passed"); never script names, flags, or
JSON. The flag list lives in chat, never in the file.

## Hard rules

- Action titles on every content slide; no topic titles; no "Questions?"
  slide; one message per slide; at most one visual per evidence slide.
- Never invent or "improve" numbers, dates, or certainty. Hedged figures
  stay hedged on the slide, with attribution.
- The Statement of Confidentiality slide precedes the closers on client
  decks; its wording comes from the template.
- Colours only from brand-tokens; no restyling, recolouring, or new layouts.
  The template's quirks are reproduced, not corrected.
- Demo and instructional template content never reaches a file.
