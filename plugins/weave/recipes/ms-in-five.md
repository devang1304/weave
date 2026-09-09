# Recipe: ms-in-five

Client-facing distillation of one Microsoft offering into five slides: what
it is, why it matters, how it works, what to do next. No companion recipe --
this deck stands alone.

## Citation discipline (the whole point of this recipe)

Every factual claim about the offering -- what it does, what it costs, how
it is licensed, whether a feature is GA, preview, or retired -- must trace to
a page fetched live via WebFetch from `learn.microsoft.com` or
`www.microsoft.com`. Never state a capability, price, licensing detail, or
release status from memory or training data alone, even for a product that
seems well known: Microsoft offerings rename, re-bundle, and re-price often
enough that last year's understanding is not this week's fact.

Each content slide's `source` field (the slide `source` line in SPEC.md)
names the exact URL and the date it was retrieved, e.g.:

```
"source": "Source: learn.microsoft.com/en-us/purview/purview-overview, retrieved 2026-09-04"
```

A claim that cannot be verified against one of the two allowed domains is
not a reason to guess: either drop it, or keep it and add it to the flag
list as unverified ("I could not confirm X via Microsoft's own docs; treat
as unconfirmed"). Never fabricate a citation to make a claim look sourced.

## Slide pattern

1. **Cover** (built automatically from `meta`).
2. **What It Is** (block/section id `what-it-is`, required): one plain-
   language sentence a non-technical exec would understand on first read.
   No jargon, no acronym left unexplained, no feature list -- just what the
   thing is.
3. **Why It Matters** (block/section id `why-it-matters`, required):
   translates the capability into a business outcome -- time saved, risk
   reduced, cost avoided, a decision made easier -- not a restatement of the
   marketing feature name. This is the slide a non-technical exec remembers.
4. **How It Works** (block/section id `how-it-works`, required): the
   simplest accurate mechanism, as 2 to 4 bullets. Not a feature dump and
   not an architecture diagram's worth of detail -- just enough that the
   audience trusts the "why it matters" claim is real.
5. **Next Steps** (block/section id `next-steps`, required): concrete and
   specific to this offering -- a pilot, a licensing conversation, a scoped
   demo, a fit assessment -- with an owner where one is known. Never a
   generic "Questions?" close and never a next step that could apply to any
   Microsoft product interchangeably.
6. Closers (Confidentiality, Thank You, Closing) are added automatically.

## Hard rules specific to this recipe

- Five slides total (cover plus the four sections above). Resist the pull
  to add a fifth content slide for "one more important detail" -- cut detail
  before cutting accuracy, and cut scope before cutting slide count.
- No slide states a number, price, licensing tier, or availability claim
  without a matching `source` citation on that same slide.
- Fetch before writing: read the fetched pages first, then draft. Do not
  draft from assumption and patch citations in afterward -- that is how a
  plausible-sounding but wrong claim survives to the deck.
- If the two allowed domains do not settle a claim (a fast-moving preview
  feature, a region-specific price), say so in the flag list rather than
  rounding it off to a confident-sounding sentence.
