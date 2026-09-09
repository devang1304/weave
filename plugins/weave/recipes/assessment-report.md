# Recipe: assessment-report

Client or internal assessment report on the document template. Pairs with
the `assessment-readout` deck recipe when the same assessment's facts also
need a readout deck; build the report first, then reuse its answers for the
deck (`companion.share`: client, date, scope).

## What each section says

- **Executive Summary**: the governing thought is the overall readiness or
  maturity verdict for `scope`, and the single most important recommendation,
  first sentence. Do not bury it under a recap of what was assessed or how.
- **Current State Findings** (`finding-table`): one row per assessed area.
  Current State and Target State are short phrases, never paragraphs (e.g.
  "Ad hoc, no owner" / "Defined policy with named owner"), and Priority is a
  single word or short label (High / Medium / Low), not a sentence.
- **Recommendations**: ranked by priority, highest first. Each recommendation
  ties back to a specific row in Current State Findings -- do not introduce a
  recommendation with no matching finding.
- **Roadmap and Next Steps**: sequencing in near-term, mid-term, and
  long-term groups. Each item is an action, not a restatement of a finding.

## Asking for what's missing

Only `scope` is required and has no way to be inferred from most source
material -- ask for it in one question, offering the phrasing the user
already used if the source names what was assessed. Never ask for anything
else this recipe doesn't list as `required`.

## Hard rules specific to this recipe

- Never invent a maturity or readiness rating the source doesn't support. If
  the source doesn't say how an area stands, the row's Current State (and/or
  Priority) cell states `[TO BE PROVIDED]` and the gap goes in the flag list
  -- the same "never fabricate" rule as everywhere else, ratings included.
- Current State Findings, Recommendations, and Roadmap and Next Steps stay
  three distinct sections; do not merge a recommendation into the findings
  table or fold the roadmap into the recommendations list even when the
  source blends them together.
