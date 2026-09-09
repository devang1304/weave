# Recipe: generic-deck

The default recipe for any PowerPoint deck without a more specific recipe.
`deck-storyline.md` is the method; this file is the spec pattern.

## Slide pattern

| Position | Layout | Content |
|---|---|---|
| 1 | Title Slide for Verticals | Title, subtitle "Client \| Month Year" (added by the builder from `meta`) |
| 2 | Title and Content (client decks; builder adds it) | Statement of Confidentiality |
| 3 | Title and Content | Executive summary: governing thought as the title, key line as 3 to 5 bullets that mirror the section headers |
| 4 | Title and Content or Section Header | Agenda when more than ~8 body slides |
| body | Title Only + one visual, Two Content, Comparison, Three Column, Three Stat | One message per slide; the title is a full sentence with a verb; a source line under any numbers |
| n-1 | Title and Content | Next steps with owner and date per line; never "Questions?" |
| last | (builder) | Thank You, Closing |

Density: 12 to 20 body slides per 30 minutes; ≤ 75 words per slide; ≤ 6
bullets of ≤ 12 words; 14 pt floor; tables ≤ 6 × 8.

## Spec skeleton

```json
{"weave": {"spec_version": "2.0", "kind": "deck", "recipe": "generic-deck",
           "base": "NW_Presentation_Base_2026.pptx", "revision": 1},
 "meta": {"title": "...", "subtitle": "Contoso | September 2026", "client": "Contoso", "date": "2026-09-15", "internal": false},
 "slides": [
   {"id": "s1", "layout": "Title and Content", "title": "<Governing thought>", "body": ["...", "..."]},
   {"id": "s2", "layout": "Title Only", "title": "<Insight>",
    "visual": {"type": "chart", "kind": "column", "categories": ["A", "B"], "series": {"Sites": [312, 88]}},
    "source": "Source: tenant inventory, Aug 2026"},
   {"id": "s9", "layout": "Title and Content", "title": "Three decisions this week keep the plan on track",
    "body": ["Approve the wave plan (J. Doe, 20 Sep)", "..."]}
 ]}
```
