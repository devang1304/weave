# Recipe: project-charter

Project charter on the document template. Pairs with the `kickoff-deck`
recipe when the same engagement's facts also need a kickoff deck; build the
charter first, then reuse its answers for the deck (`companion.share`:
client, date, project_name).

## What each section says

- **Purpose and Objectives**: the governing sentence first -- why the
  project exists and what it will achieve. Do not bury it under background.
- **Scope**: what's in, and explicitly what's out. A scope section with no
  out-of-scope statement is incomplete.
- **Stakeholders and Roles**: names and roles on both the Netwoven and
  client sides. A missing name is a gap for the flag list, not an invented
  placeholder person -- never guess who fills a seat.
- **Timeline and Milestones**: a list of dated milestones, not a detailed
  task-by-task plan.
- **Success Criteria**: measurable, checkable outcomes -- "90% of mailboxes
  migrated by 15 October," not "the migration goes smoothly." Vague
  aspirations belong in Purpose and Objectives, not here.
- **Assumptions and Risks** (optional -- drop the section if the source has
  nothing to add): what could go wrong and what's assumed true going in.

## Asking for what's missing

Only `project_name` is required and has no way to be inferred from most
source material -- ask for it in one question. Never ask for anything else
this recipe doesn't list as `required`; a missing optional section is a gap
to flag, not a question to ask.

## Hard rules specific to this recipe

- Never invent a stakeholder name, a milestone date, or a success metric the
  source doesn't support. A gap becomes `[TO BE PROVIDED]` in the block and
  a line in the flag list.
- Success Criteria must stay measurable; if the source only offers a vague
  goal, flag it rather than sharpening it into a number the source never
  gave.
