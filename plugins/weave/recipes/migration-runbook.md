# Recipe: migration-runbook

A migration runbook on the document template: the broad end-to-end plan
for moving a system, workload, or tenant, from prerequisites through
validation. Use `cutover-runbook` instead when the source is really about
one go-live window, not the whole migration.

## What each section says

- **Overview and Scope**: what's migrating, from and to, and the target
  completion window, in plain language before any procedural detail. A
  reader who stops here knows what's moving and by when, not yet how.
- **Pre-Migration Checklist**: a concrete list of prerequisites -- access
  provisioned, backups taken, communication sent -- each phrased so it can
  be verified independently (checked off yes/no), not a vague readiness
  statement.
- **Migration Steps**: numbered, one action per step, each naming an owner
  and either a date or a relative time offset (for example "Day 2" or
  "T+4h"). This is the section the `owners_and_dates_in` check reads; every
  step needs both an owner and a date/offset, not just some.
- **Rollback Plan**: the actual trigger conditions that call for rolling
  back, stated as concrete thresholds or failure conditions, never just
  "roll back if needed."
- **Validation Steps**: concrete checks that prove the migration
  succeeded -- specific tests, counts, or confirmations -- never just
  "verify it worked."

## Asking for what's missing

Only `system_name` is required and usually can't be inferred reliably
without asking -- request it in one question if the source doesn't name the
system, workload, or tenant being migrated.

## Hard rules specific to this recipe

- Never invent a step, an owner, or a date/time offset the source doesn't
  support. A gap becomes `[TO BE PROVIDED]` in that step or checklist item
  and a line in the flag list, not a filled-in guess.
- Rollback Plan must state real trigger conditions (a failure threshold, an
  error rate, a named stakeholder call), not a placeholder sentence that
  says only to roll back if something goes wrong.
- Validation Steps must be independently checkable (a count, a test, a
  login, a report) -- do not accept "confirm it worked" as a step on its
  own.
