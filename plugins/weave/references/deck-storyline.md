---
title: Deck Storyline Method
covers: consulting-deck structure for decks built on NW_Presentation_Base_2026.pptx
last_verified: 2026-09-02
---

# Deck storyline method

`writing-method.md` produces a pyramid outline. This file turns that outline
into a deck. Run it after Phase B of the analysis and before any slide is
built. The layout names, placeholder indexes, and geometry below are exact for
the 2026 presentation template (see `ppt-template-guide.md` §3 for the full
catalog). Colours come from `brand-tokens.md`; chart and framework recipes
from `visual-standards.md`.

---

## 1. The ghost deck comes first

A ghost deck is the storyline written as a numbered list of action titles,
one per slide, with a one-line note of the evidence each slide will carry.
No slide is created until the ghost deck reads as a complete argument on its
own.

Procedure:

1. Take the governing thought and the key line from the pyramid.
2. Write the executive summary title (the governing thought as a sentence).
3. Write one title per key-line point. Under each, write one title per
   supporting point that deserves its own slide.
4. Read only the titles, top to bottom. They must tell the whole story, in
   order, with no gaps and no repeats. This is the dot-dash test.
5. Show the ghost deck to the user for long or high-stakes decks before
   building slides. Fixing an argument costs minutes here and hours later.

A ghost deck entry looks like this:

```
07  Migration risk concentrates in 14 ungoverned site collections
    evidence: bar chart, sites by governance state, source inventory 2026-08
    layout: Title Only + chart
```

---

## 2. One message per slide

Each slide answers exactly one question raised by the slide above it in the
pyramid. If a slide needs two "and also" clauses in its title, it is two
slides. If a slide has no message, only a topic ("Timeline", "Risks"), it is
not finished.

The body of a slide exists to prove the title. Anything on the slide that does
not support the title is removed or moved to the appendix.

---

## 3. Action titles

An action title is a full sentence that states the conclusion of the slide.

| Rule | Detail |
|---|---|
| Full sentence | Subject, verb, object. "Three workloads carry 80 percent of the migration effort", not "Migration effort" |
| Two lines at most | About 90 characters at the master's 40pt title in an 11.5in wide placeholder. Longer titles wrap to three lines and collide with the body. To be confirmed by render |
| Twelve words or fewer | The validator warns above 12 words |
| Sentence case | Match `writing-method.md`; no title case, no trailing period |
| Numbers where you have them | "Cuts license spend by 18 percent", not "Reduces license spend" |
| No hedging filler | "may potentially" and "it is believed that" are cut |
| Consistent voice | All titles declarative; no questions except the one question the deck exists to answer, if any |

Test each title by removing the slide body. If the audience could still act on
the title alone, it is an action title.

---

## 4. Horizontal and vertical logic

Horizontal logic is read across the titles. The sequence of titles at one
level must be MECE siblings of the same kind (all reasons, all steps, all
options) in a deliberate order (time, structure, or importance). Reading the
titles of section 3 must give the whole of section 3.

Vertical logic is read down a single slide. The title states the "so what";
the body is the "because". Apply the so-what test to every body element: if
you cannot say how it supports the title, it does not belong. Apply the
"why should I believe you" test to every title: the body must answer it.

Breadcrumbs keep the audience oriented in long decks. On evidence slides
built on `Title Only` (L12), the builder may add a 10.5pt `#595959` line at
the top left with the section name in all caps and letter spacing 70 (the
same eyebrow style the Grid layouts use). Never repeat the breadcrumb in the
title.

---

## 5. Slide type to layout map

Layout names are exact strings from the template. `idx` values are the
placeholder indexes to pass to `ph(slide, idx)` from `nw_pptx_helpers.py`.
Cover geometry is in inches from the layout XML.

| Slide type | Layout (index in master) | Placeholders to fill | Notes |
|---|---|---|---|
| Cover | `Title Slide for Verticals` (L3) | `ctrTitle` idx 0 (48pt `#00B0F0`, bottom anchored, 6.51in x 1.23in, 5.44in wide), `subTitle` idx 1 (18pt `#595959`), pictures idx 15 and 16 | Never `Title Slide` (L1): it has no title placeholder. Subtitle carries client name and date. Two picture placeholders are removed unless `--cover-image` supplies photos |
| Executive summary | `Title and Content` (L2) for prose; `Three Column` (L10, body idx 13, 14, 15) when the key line has three points | Title idx 0, body idx 1 | See §6 |
| Situation, complication, answer | `Title and Content` (L2) or `Comparison` (L11: headers idx 1 and 3 at 24pt bold, bodies idx 2 and 4) | | Use Comparison for "today versus target" |
| Agenda | `Agenda` (L4) | body idx 21 (18pt `#44546A`), picture idx 38 | The word "Agenda" is baked into the layout as static text; do not add a title. Skip the agenda for decks under 12 body slides |
| Section divider | `Section Header` (L8) | title idx 0 (60pt), body idx 1 (24pt), optional picture idx 32 | The template's own "SECTION DIVIDER" demo slide is dropped by the builder |
| Evidence (chart, table, framework) | `Title Only` (L12) | title idx 0; visual drawn into the content area x 0.92in, y 1.6in, w 11.5in, h 5.0in (to be confirmed by render) | One visual per slide |
| Two-sided argument | `Two Content` (L9: bodies idx 1 and 2, each 5.67in wide) or `Comparison` (L11) | | Bullets left, chart right is acceptable on Two Content |
| Three parallel points | `Three Column` (L10) | idx 13, 14, 15 | Headings inside each column are bold first lines, not extra placeholders |
| Four to six parallel points | `Six Icon` (L28: labels idx 13 to 18, icons idx 19, 20, 21, 25, 26, 27; title idx 0 sits at left, 3.94in wide), `Grid Layout 2-2` (L32), `Grid Layout 2-2-2` (L33) | Grid eyebrows are 10.5pt all caps tracked | Six is the ceiling |
| Big numbers | `Three Stat` (L26: numerals idx 10, 15, 17 at 96pt `#00B0F0`; labels idx 11, 14, 16 at 21pt black) | No title placeholder; the numbers are the message | `Big Statement with Illustration` (L25: statement idx 16 at 72pt, eyebrow idx 14, picture idx 13) for one number or one sentence |
| Quote | `Testimonial On Right` (L19) | title idx 0, quote idx 21, name idx 35, position idx 36, logo idx 34, photo idx 32 | Not L20 (its bar is purple accent2). L19 and L20 suppress the master footer |
| Next steps | `Title and Content` (L2) with an Action / Owner / Date table, or `Title Only` (L12) plus `add_table` | | See §7 |
| Appendix divider | `Section Header` (L8) titled "Appendix" | | Everything after it is backup |
| Closers | Confidentiality on `Title and Content` (L2), then the template's Thank You (Blank, L21) and Closing Slide (L38) | | Generated and ordered by `new_deck_pptx.py`; see `ppt-template-guide.md` §7 |

Avoid `Company Profile` (L5, two paragraphs of baked boilerplate), `Photo on
Left` (L6, body is 18pt with no bullets and inherits nothing useful),
`Content with Caption` (L34, 32pt body), `Title and Vertical Text` (L36) and
`Vertical Title and Text` (L37).

---

## 6. Executive summary rules

1. One slide, first body slide after the cover (and after the agenda if
   present).
2. Title is the governing thought.
3. Body is the key line: two to four statements, each one sentence, in the
   same order as the sections that follow. Each statement is the action title
   of its section, lightly shortened.
4. No bullets under bullets. No numbers that do not reappear later.
5. The last line is the ask: the decision or action the audience must take.
6. For a three-point key line, use `Three Column` with the point as a bold
   first line and one supporting sentence beneath.
7. The executive summary and the next-steps slide must agree. Read them
   together before finalizing.

---

## 7. Closing and next steps

The deck ends with what happens next, not with "Questions?". Never title a
slide "Questions", "Thank you" (the template's own Thank You closer already
exists), or "Discussion".

The next-steps slide:

| Column | Content |
|---|---|
| Action | Verb first, one line |
| Owner | A role or named person the audience recognises |
| Date | A real date or week, never "TBD" (flag it instead) |

Three to six rows. The first row is the decision requested in the executive
summary. The table uses `add_table` from `nw_pptx_helpers.py` (header
`#00B0F0`, bands `#FFFFFF` / `#F2F2F2`, Segoe UI 12pt).

---

## 8. Density rules

| Measure | Limit | Why |
|---|---|---|
| Words on a body slide | 75 or fewer (validator warns above 90) | The audience reads or listens, not both |
| Bullets | 6 or fewer, 12 words each, one level | Nested bullets are a pyramid that belongs on two slides |
| Visuals per slide | 1 (validator warns above 2) | One message, one proof |
| Table size | 6 columns by 8 rows at most | Anything larger becomes a chart or an appendix table |
| Smallest text | 14pt on body slides; 12pt allowed only for chart chrome, source lines, and table cells | Segoe UI at 12pt is at the limit of projector legibility |
| Body slides per 30 minutes | 12 to 20 | Two minutes a slide including questions |
| Title length | 2 lines, 12 words | See §3 |
| Appendix | Unlimited, after the Appendix divider | Backup never competes with the storyline |

---

## 9. Anti-patterns

| Anti-pattern | What to do instead |
|---|---|
| Topic titles ("Background", "Findings") | Rewrite as the conclusion of the slide |
| Title on one slide, "continued" on the next | Split the message or cut the content |
| Chart with a chart title and a slide title | Delete the chart title; the slide title is the message |
| Rainbow charts using the theme accent cycle | Use `PALETTE_DECK` in order; highlight one series, grey the rest |
| Bullets that restate the title | Delete; add evidence or delete the slide |
| Agenda for a 6-slide deck | Skip the agenda |
| "Questions?" closer | Next steps slide, then the template closers |
| Logo or client name repeated on every slide | The master footer already carries the wordmark |
| Text boxes drawn over an empty placeholder | Fill the placeholder or remove it (`finalize --prune-empty`) |
| Colour as the only encoding ("red items are at risk") | Add a label or marker; see `brand-tokens.md` §5 |
| Stock photos on evidence slides | Photos only on cover, section headers, and testimonial slides |
| Screenshots without a callout | Crop to the region, add one callout box, state the point in the title |

---

## 10. Research bases

These are the sources the method draws on. They are named so an author can
go deeper; nothing here is a quotation, and no page numbers are claimed.

| Source | What this method takes from it |
|---|---|
| Barbara Minto, *The Pyramid Principle* | The governing thought, key line, SCQA introduction, MECE grouping, and the rule that ideas at any level summarise the ideas below them. This is the spine of `writing-method.md` and of §1 to §4 here |
| Gene Zelazny, *Say It With Charts* | Choose the comparison first (component, item, time series, frequency, correlation), then the chart form; the chart's message is its title. Applied in `visual-standards.md` §1 |
| Edward Tufte, *The Visual Display of Quantitative Information* and later essays | Maximise the share of ink that carries data; remove decorative chrome; avoid slide bullets as a substitute for analysis. Behind the chrome rules and the "one visual" limit |
| Cole Nussbaumer Knaflic, *Storytelling with Data* | Declutter, focus attention with one highlight colour against grey, and lead with the takeaway. Behind the highlight-series rule and action titles |
| Nancy Duarte, *slide:ology* and *Resonate* | Audience-first structure, contrast between what is and what could be, and the discipline of one idea per slide |
| Richard Mayer, multimedia learning research, and John Sweller, cognitive load theory | People process words and pictures in limited channels; redundant text read aloud and decorative elements add extraneous load. Behind the 75-word ceiling, the ban on nested bullets, and the rule that the visual must support the title rather than decorate it |
| Gestalt principles of perception (proximity, similarity, enclosure, continuity, closure) | Group related items by spacing rather than boxes, align to a grid, use enclosure sparingly. Behind the framework geometry in `visual-standards.md` |

Where these sources disagree (Tufte's scepticism of slides versus Duarte's
embrace of them), this method sides with the client reading a deck on a
laptop after the meeting. The deck must stand alone.
