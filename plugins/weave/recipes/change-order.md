# Recipe: change-order

An amendment to an already-signed Netwoven SOW, on the document template —
not a new SOW. It exists to record what changed against a document that
already exists and is already signed, so it references the original rather
than restating it.

## What each section says

- **Summary of Change**: names the original SOW being amended (its title,
  and its date if known) in the first sentence, then one or two sentences
  on what this change order does overall (adds scope, removes scope, shifts
  schedule, changes cost, or some combination). This is the governing
  thought; the reader should know what changed and against which document
  before reading further.
- **Scope Impact**: states exactly what is added, removed, or changed in
  scope, referencing the specific original scope item(s) being touched (by
  name or section reference from the original SOW, not a vague "the scope").
  Do not restate scope items that are unaffected.
- **Schedule Impact** (optional — drop the section if the source shows no
  schedule change): the delta to dates or duration, stated as a change from
  what the original SOW set, not as a fresh timeline.
- **Cost Impact**: the delta — increase or decrease, and the new total if
  known. A missing number is `[TO BE PROVIDED]`, never a guess and never a
  fabricated delta, even when the direction of the change is obvious from
  the source.
- **Approval**: blank signature lines mirroring how an amendment is
  typically countersigned (client signatory and Netwoven signatory, each
  with name/title/date lines), not the full legal signature-block
  boilerplate restated from a SOW.

## Asking for what's missing

Only `original_sow` is required and usually can't be inferred reliably from
the change request itself — ask for it in one question if the source
doesn't name the original SOW clearly (title, and date if the source or the
user gives one).

## Hard rules specific to this recipe

- Never invent a cost or schedule delta that isn't stated or isn't clearly
  computable from the source. A gap is `[TO BE PROVIDED]` and a line in the
  flag list, the same as everywhere else — the direction of a change (up or
  down) is not license to guess the amount.
- Never treat this as a full new SOW: no fee tables, no milestones list, no
  restatement of the full scope of work. This document is intentionally
  short — it points at the original SOW for everything unaffected rather
  than duplicating it.
- If the source suggests the "change" is actually a brand-new engagement
  (new client, unrelated scope, no reference to an existing signed SOW),
  say so and confirm before proceeding — that belongs in the `sow` skill,
  not here.
