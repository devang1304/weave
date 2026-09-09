# Recipe: cutover-runbook

A cutover runbook on the document template: the narrow, time-critical
go-live window procedure, not the whole migration. Use `migration-runbook`
instead when the source covers the broader end-to-end plan.

## What each section says

- **Pre-Cutover Checklist**: the final go/no-go gate items checked
  immediately before the window opens, each independently verifiable --
  not a restatement of the migration's earlier prerequisites.
- **Cutover Sequence**: time-boxed and numbered, either relative to go-live
  (T-2h, T-1h, T-0, T+30m) or as explicit clock times, each step naming an
  owner. This is the section the `owners_and_dates_in` check reads; every
  step needs both a time and an owner.
- **Rollback and Abort Criteria**: the specific conditions that trigger an
  abort (a failure threshold, an error rate, a missed checkpoint) and who
  has the authority to call it -- never just "abort if needed."
- **Post-Cutover Validation**: concrete checks -- specific tests, counts,
  confirmations -- that must pass before declaring the cutover done, not
  just "verify it worked."
- **Communication Plan** (optional -- drop the section if the source
  doesn't name who gets notified): who is notified at which milestones
  (start, complete, abort) and by what channel.

## Asking for what's missing

Only `system_name` is required and usually can't be inferred reliably
without asking -- request it in one question if the source doesn't name
what is cutting over and to what target.

## Hard rules specific to this recipe

- Never invent a time, an owner, or an abort threshold the source doesn't
  support. A gap becomes `[TO BE PROVIDED]` in that step and a line in the
  flag list, not a filled-in guess.
- Rollback and Abort Criteria must name both the trigger condition and who
  has authority to call it -- a criteria list with no named decision-maker
  is incomplete.
- Keep this recipe narrower than `migration-runbook`: it covers only the
  go-live window, not pre-migration data prep or the broader project
  timeline.
