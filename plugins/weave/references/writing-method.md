# The writing method: what to say, then how to say it

The shared writing core for every skill in this plugin. Part 1 (the analysis
method) decides *what* is said and *in what order*, run before writing a
single sentence of output; Part 2 (writing standards) decides *how each
sentence reads* once the structure is settled. A beautifully-templated
document with a buried lead is a failure, and so is a perfectly pyramidal
document that quietly invents support for a claim the source never earned,
or that states its case in flabby prose.

## Contents

- Part 1: The analysis method (recursive Minto Pyramid, first-principles)
  - Phase A. Deconstruct
  - Phase B. Rebuild
  - Phase C. Verify
- Part 2: Writing standards
  - Voice and tone, house punctuation, mechanics, structure, accessibility
  - Client-facing specifics, decks, SOW register

---

# Part 1: The analysis method, a recursive Minto Pyramid on first-principles ground

Run this **before writing a single sentence of output**. The method has
three phases: deconstruct to fundamentals, rebuild as a pyramid, verify. The
output of phases A+B is a **pyramid outline**. Produce it as a visible
intermediate artifact (in your working notes or, for long documents, shown
to the user) before drafting prose.

## Phase A. Deconstruct (first principles)

Most drafts are pattern-matched assemblies: boilerplate openings, sections
that exist because similar documents have them, claims repeated from earlier
documents without re-examination. First-principles analysis strips a document
back to what is actually known, for whom, and why it matters, then rebuilds
from only that.

### A1. Establish the fundamentals

Answer these from the source and the user's request; ask the user only for
what is decision-critical and genuinely absent:

- **Reader:** who exactly reads this, and what do they already know/believe?
- **Purpose:** what single decision or action should this document produce?
  A document that "informs" is underspecified. Informs *toward what*?
- **Stakes:** what happens if the reader does nothing, or decides wrong?
- **Constraints:** length, deadline, confidentiality, mandated sections.

### A2. Inventory the substance

Extract every distinct claim, fact, number, recommendation, risk, commitment,
and open question from the source. Classify each:

| Class | Meaning | Fate |
|---|---|---|
| **Load-bearing** | Directly supports what the reader must decide/do | Keeps, becomes pyramid material |
| **Context** | Reader needs it to understand a load-bearing item | Keeps, subordinated under what it serves |
| **Filler** | Restatement, throat-clearing, generic boilerplate ("in today's fast-paced digital landscape"), hedges that carry no information | Cut without mercy |
| **Orphan** | True but serves no reader question this document answers | Cut; mention to the user if it looks like it belongs in a *different* document |
| **Unsupported** | Asserted but not evidenced in the source | Flag. Never silently keep as fact, never invent support, never silently drop |

### A3. Interrogate the assumptions

For each load-bearing item, ask: *is this established, or inherited?* Numbers
copied between drafts drift; "best practice" claims often trace to nothing;
recommendations sometimes survive after their original rationale died. Anything
that fails this test moves to **Unsupported** and gets flagged. The rewrite
must never make the source sound more certain than it is. Confidence levels
are substance, not style.

---

## Phase B. Rebuild (the Minto Pyramid)

### B1. Find the governing thought with SCQA

The governing thought is one sentence that answers the reader's question. To
find it, build the SCQA chain:

- **Situation.** What the reader already accepts as true.
- **Complication.** What changed or is wrong, creating tension.
- **Question.** The question that tension raises in the reader's mind.
- **Answer.** The governing thought. This leads the document.

The introduction of the final document tells S→C→A in a few sentences and
then gets out of the way. If you cannot state the governing thought in one
sentence, Phase A is not finished.

### B2. Build the key line (MECE)

The key line is the 2 to 4 statements that together make the governing thought
true. Requirements:

- **Insight statements, not topic labels.** "Migration risk concentrates in
  ungoverned site sprawl", not "Risks". A reader of *only* the headings must
  receive the argument.
- **MECE:** mutually exclusive (no point belongs in two groups; if one does,
  the grouping is wrong, not the point), collectively exhaustive (together
  they fully support the governing thought; no load-bearing item is homeless).
- **Same kind:** siblings are all reasons, all steps, or all criteria, never
  a mix.
- **Deliberate order:** deductive (premise → conclusion) only when the reader
  must follow a chain of reasoning; inductive (parallel evidence under one
  insight) as the default; within a group, order by time, structure, or
  degree, and be able to say which.

### B3. Recurse

**Every section is itself a pyramid.** Its heading is its governing thought;
its paragraphs are its key line; each paragraph's first sentence is that
paragraph's point and the rest of the paragraph supports it. Apply B1 to B2 at
each level until a unit is a single paragraph. Two tests at every level:

- **Vertical:** each level answers exactly the question the statement above
  raises in the reader's mind (Why? How? Which ones?). If a section raises
  "why?" and its children answer "how?", the pyramid is broken.
- **Horizontal:** siblings are MECE and of the same kind.

Depth guidance: match structure to substance. A 2-page memo is one pyramid
with 2 to 3 key-line sections and no subsections. Do not manufacture heading
levels to look thorough. Empty hierarchy is filler wearing a suit.

### B4. Two shapes the pyramid must take

- **Decks.** The pyramid becomes a ghost deck before any slide exists: the
  governing thought is the executive-summary title, each key-line point is
  a section, and every supporting point that earns a slide becomes an
  action title. Continue in `deck-storyline.md` (ghost deck, action titles,
  layout map, density rules) once Phase B is done.
- **Statements of Work.** The reader signs, so the pyramid is short (what
  Netwoven delivers, for how much, by when, under which assumptions) and it
  is poured into the template's fixed section order rather than a free
  outline: Glossary, Executive Summary, Microsoft Products Utilized in
  Project, Scope of Request and Deliverables, Assumptions, Project
  Operations, Budget and Timeline, Signatures. Do not reorder or rename
  those sections; put the governing thought in the Executive Summary and the
  key line into Scope of Work and Deliverables. See `sow-template-guide.md`
  §9 for what each section keeps, and §10 for the paragraphs that are never
  rewritten.

---

## Phase C. Verify

Run all three checks before delivering; fix and re-check on failure.

1. **The skim test.** Read only: governing thought, key-line headings, and
   the first sentence of each paragraph. That skim must deliver the complete
   argument and the ask. If it doesn't, points are buried.
2. **The fidelity check.** Every number, commitment, name, date, and
   confidence level from the source survives, unchanged, or appears in the
   flag list. Nothing is invented. No new facts, numbers, quotes, or
   certainty. Reordering and regrouping are allowed; meaning drift is not.
3. **The flag list.** Deliver alongside the document (in chat, never inside a
   client deliverable): unsupported claims found, assumptions you had to
   make, orphan content set aside, and decisions only the author can make.
   An honest flag list is part of the deliverable's quality, not an
   admission of failure.

---

# Part 2: Writing standards for Netwoven deliverables

Prose-level rules for the rewritten document. Part 1 above decides *what* is
said and *in what order*; this part decides *how each sentence reads*.
Grounded in the Microsoft Writing Style Guide (Netwoven is a
Microsoft-ecosystem consultancy; its clients read Microsoft docs all day, so
matching that register reads as native quality to them), plus Netwoven house
rules carried over from the voice-align plugin.

## Voice and tone

- **Active voice, named actors.** "Netwoven will migrate the sites", not
  "the sites will be migrated." Passive is acceptable only when the actor is
  genuinely unknown or irrelevant.
- **Warm-professional, never chummy, never stiff.** Confident declaratives.
  No exclamation marks in deliverables.
- **Bigger ideas, fewer words.** Target sentences under ~25 words. One idea
  per sentence. If a sentence needs a breath mid-way, it's two sentences.
- **Second person for instructions** ("you can review the mapping in
  Appendix A"), third person for findings and commitments.
- Hedges are substance: keep the *fact* of uncertainty in plain words
  ("we have not yet validated X"), cut decorative hedging ("it could perhaps
  be argued that").

## Netwoven house punctuation (adapted from voice-align for documents)

- **Em dashes and en dashes: never.** Rejoin the sentence, split it, or use a
  comma. Numeric ranges use "to" (3 to 5 weeks) or the word "through".
- **Colons: only as the terminator of a lead-in to a rendered list or
  table.** Never mid-sentence ("three risks: cost, time, scope" → "Three
  risks stand out. Cost..."). A lead-in sentence before a list may end with a
  colon; nothing else may.
- Semicolons sparingly, only joining two short, tightly parallel sentences.
- Oxford comma always.

## Mechanics (Microsoft style)

- **Headings: sentence case** ("Migration approach and timeline", not
  "Migration Approach And Timeline"). Small caps are a style effect, not a
  writing instruction: the document template's `NW Heading 1/2` and the
  SOW templates' harmonized `Heading1/2` render small caps regardless of
  what you type, and the deck's Grid-layout eyebrows render all caps. Write
  sentence case everywhere and let the style do its work; typed capitals
  would survive into the TOC, where small caps are not applied.
- **Never skip heading levels.** NW Heading 3 only under an NW Heading 2.
- **Numbers:** digits for 10 and above, and for all measurements, versions,
  percentages, and counts a reader might compare (3 sites, 12 workflows).
- **Acronyms:** spell out at first use with the acronym in parentheses, then
  acronym only. Never define what the reader's role guarantees they know
  (a SharePoint admin needs no expansion of "SharePoint").
- **One name per concept.** Pick "site collection" or "site" and hold it;
  synonym variety reads as different things in technical prose.
- Product names exactly as Microsoft spells them: SharePoint Online,
  Microsoft 365, Teams, Power Platform, Entra ID.

## Structure devices

- **Lists** only for 3+ genuinely parallel items; each item a fragment or a
  sentence, consistently, never a mix. If items have two dimensions, it's a
  table, not a list.
- **Tables** for any comparison, mapping, or dataset. In Word deliverables,
  always a named Netwoven table style (see `visual-standards.md` §10.3 for
  the 1/2/3 selection rule and `word-template-guide.md` §3.2 for the
  styles), never a manually-formatted grid. Charts, frameworks, stat tiles,
  and figure geometry follow `visual-standards.md`; colours only from
  `brand-tokens.md`.
- **Captions** on every figure and table ("Table 3 Site inventory by
  workload"), using the template's Caption style with a `SEQ` field so the
  numbering and the lists of figures and tables update (`visual-standards.md`
  §10.2).
- **No wall of text:** paragraphs of 2 to 5 sentences.
- **No decoration:** bold only for genuine warnings or defined terms at
  first definition. Never bold whole sentences for emphasis.

## Accessibility

- Meaning never carried by color or position alone.
- Alt text on any image that carries information.
- Descriptive link/reference text ("the assessment workbook in Appendix B",
  never "click here").

## Client-facing specifics

- No internal jargon, ticket numbers, internal system names, or colleague
  first-names-only. "Our engineering team", not "Raj's team".
- No internal candor artifacts: cost-cutting rationales, internal
  disagreements, "the client won't notice" phrasing. If the source contains
  material that must not reach a client, cut it and put it in the flag list.
- Commitments (dates, prices, scope) are quoted exactly from the source or
  flagged, never rounded, never "improved".

## Decks

Prose rules above apply to slide text, with the deck-specific limits in
`deck-storyline.md`: action titles are full sentences of 12 words or fewer,
body slides carry 75 words or fewer, bullets are one level deep, and the
deck ends with next steps rather than "Questions?". The Confidentiality
slide's three paragraphs are copied from the Word template verbatim and are
exempt from every rule here.

## SOW register

A Statement of Work is a contract the client signs, so its prose is drier
than a report and much of it is fixed. In a SOW:

- **Boilerplate is verbatim.** Glossary definitions, assumptions, change
  request, status reporting, payment, staffing, signature, and
  confidentiality paragraphs are copied exactly from the template
  (`sow-template-guide.md` §10). House punctuation rules do not apply inside
  them; do not "improve" a legal sentence.
- **Netwoven and the client are named in full** as the template names them
  ("Netwoven" or "Netwoven, Inc.", and the client's legal name that the
  Company field carries). No "we" or "you" in scope, assumptions, or
  payment text; "we" is acceptable only in the Executive Summary.
- **Every number is sourced.** Rates, hours, weeks, milestone payments,
  discounts, the expenses cap, and dates come from the source or are flagged;
  never estimate a fee to fill a table. Money uses the template's format
  inside tables (`0.00`) and `$` with thousands separators in prose.
- **Scope is written as commitments, not aspirations.** "Netwoven will
  configure…" and "Netwoven will deliver…" for in-scope items; "Out of
  Scope" is a real list, never left empty.
- **Deliverables are nouns with a format** ("Migration plan, document";
  "Configured tenant, environment") and each maps to a milestone (Milestone
  variant) or a scope section (T&M variant).
- **Assumptions are conditions, one per bullet**, written as facts about the
  client's environment or obligations, in the template's own list style.
- **Keep the template's section titles and order** (Part 1, §B4 above); use
  `Heading1/2/3` only, never `Heading4` or the `NW Heading` styles.
- **Dates** in prose as Month D, YYYY; the bound date fields take the
  builder's `M/d/yyyy` form. Never leave `MM/DD/YYYY`, `<<…>>`, or a yellow
  highlight in the output.
