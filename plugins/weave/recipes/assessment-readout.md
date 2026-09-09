# Recipe: assessment-readout

Assessment readout deck: the same assessment's facts as the
`assessment-report` recipe, distilled into a 3-slide readout. Build the
report first when both are wanted; reuse its answers here
(`companion.share`: client, date, scope) rather than asking twice.

## Slide pattern

1. **Cover** (built automatically from `meta`).
2. **Summary slide** (block/section id `summary`, required): one slide,
   title states the overall readiness/maturity verdict for `scope` as a
   full sentence ("Governance is maturing but ownership gaps block scale"),
   body is 3 to 5 bullets covering the headline verdict and the single most
   important recommendation, matching the report's Executive Summary.
3. **Findings slide** (block/section id `findings`, required): the
   distillation of Current State Findings, not the full table re-typeset.
   Use the `finding-table` framework as this slide's visual only when a
   table view of the highest-priority areas earns its own slide; otherwise
   summarize in 3 to 5 bullets. Do not add both a bullet summary and a table
   saying the same thing.
4. **Next Steps slide** (block/section id `next-steps`, required): the
   near-term actions from the report's Roadmap and Next Steps, with owner
   and date per line where known. Never a "Questions?" slide.
5. Closers (Confidentiality, Thank You, Closing) are added automatically.

## Hard rules specific to this recipe

- Exactly the three-slide shape above (summary, findings, next-steps); do
  not add a slide that repeats the full report's section list. A readout
  deck is the distillation, not the report re-typeset.
- Every rating or verdict on the readout must trace to the same finding in
  the paired report, if one was built in the same turn. Never let the deck
  read more positive than the report.
