# Recipe: steerco

Steering committee deck: the same period's facts as the `status-report`
recipe, answering four questions a steering committee actually has. Build
the status report first when both are wanted; reuse its answers here
(`companion.share`: client, date, period) rather than asking twice.

## Slide pattern

1. **Cover** (built automatically from `meta`).
2. **Summary slide** (block/section id `summary`, required): one slide,
   title states the overall answer to "are we on track" as a full sentence
   ("Migration is on track; two decisions need the committee this week"),
   body is 3 to 5 bullets covering the four questions a steerco answers:
   where things stand, whether the plan is on track, what needs attention,
   what the committee needs to decide. Use the `rag-table` framework as this
   slide's visual when a workstream-by-workstream view earns its own slide
   instead of folding into the bullets — do not add both a bullet summary
   and a table saying the same thing.
3. **Next slide** (block/section id `next`, required): next steps with
   owner and date per line, or a decision table (option, recommendation,
   owner). Never a "Questions?" slide.
4. Closers (Confidentiality, Thank You, Closing) are added automatically.

## Hard rules specific to this recipe

- Exactly the four-question shape above; do not add a slide that repeats
  the status report's full section list. A steerco deck is the two-slide
  distillation, not the report re-typeset.
- Every RAG or status word on the summary slide must trace to the same
  answer given for the paired status report, if one was built in the same
  turn. Never let the deck read more optimistic than the document.
