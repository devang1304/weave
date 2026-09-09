# Recipe: status-report

Client or internal status report on the document template. Pairs with the
`steerco` deck recipe when the same period's facts also need a steering
committee readout; build the doc first, then reuse its answers for the deck
(`companion.share`: client, date, period).

## What each section says

- **Executive summary**: the governing thought is the overall RAG status and
  the one decision needed this period, first sentence. Do not bury it under
  a recap of what the project is.
- **Progress this period** (`rag-table`): one row per workstream. Status is
  a single word (On track / At risk / Blocked), never a paragraph. "This
  period" and "next period" columns are short phrases, not full sentences.
- **Risks and issues** (`raid-table`, optional — drop the section if the
  source has nothing new to add for this period rather than repeating last
  period's list unchanged): only items that are new, changed, or need the
  reader's attention now.
- **Next period and decisions needed**: a list, each line an action with an
  owner and a date. This is the section the `owners_and_dates_in` check
  reads; every line needs both, not just some.

## Asking for what's missing

Only `period` is required and has no way to be inferred from most source
material — ask for it in one question, offering the phrasing the user
already used if the source names a date range. Never ask for anything else
this recipe doesn't list as `required`; a missing optional section is a
gap to flag, not a question to ask.

## Hard rules specific to this recipe

- Never invent a RAG status. If the source doesn't say how a workstream is
  doing, the row's Status cell states `[TO BE PROVIDED]` and the gap goes in
  the flag list — the same "never fabricate" rule as everywhere else, RAG
  colour included.
- "Progress this period" and "Next period" must stay two distinct sections;
  do not merge them even when the source's own notes blend them together.
