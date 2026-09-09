# Routing: which Weave skill owns a request

This is the only place routing rules live. Skill descriptions name siblings
only in their `when_to_use` line; everything else points here.

## Contents
- The route table
- Tie-breaks
- The one-question rule
- Words to use and words to avoid in chat
- Where the file lands

## The route table

| The user wants | Owner | Notes |
|---|---|---|
| An email, a Teams or Slack message, a short reply | **voice-align** plugin | Never build a file nobody asked for |
| Anything the client signs or that prices work: SOW, statement of work, scope with fees, rate card, milestones, engagement letter | `sow` | Variant chosen from the fee model; ask one question if unclear |
| A change to a signed SOW: change order, amendment, extra hours, scope delta | `/weave:change-order` | Explicit only; a short amendment referencing the original SOW, never a new one |
| A pre-sales pitch, offering write-up, or rough investment estimate before any commercial terms are fixed | `/weave:proposal` | Explicit only; the moment fees, milestones, or rates enter it becomes `sow` |
| A PowerPoint deck: slides, presentation, QBR, roadmap deck, exec briefing, pitch deck | `deck` | Presentation-shaped material (agenda, talk track, meeting on a date) counts even without the word "deck" |
| A recurring status update, weekly/monthly report, or a steering-committee deck | `/weave:status` | Explicit only; report and deck share one set of facts |
| A Word document for a client or external reader: findings report, recommendation | `document` | Confidentiality page kept; Company = the client |
| A Word document or Markdown for colleagues: spec, SOP, postmortem, decision memo, meeting notes, wiki page | `document` (internal) | Picks docx vs Markdown by weight and says which |
| An assessment, readiness review, or maturity/gap analysis, with or without a readout deck | `/weave:assessment` | Explicit only |
| An architecture decision record or a high-level design / solution design document | `/weave:design` | Explicit only; one decision (ADR) vs. a whole system (HLD) |
| A migration plan or a go-live cutover procedure | `/weave:runbook` | Explicit only; whole migration vs. just the cutover window |
| A project kickoff deck or a project charter for a new engagement | `/weave:kickoff` | Explicit only; deck and charter share one set of facts |
| An adoption strategy, a rollout communications plan, or end-user training material | `/weave:adoption` | Explicit only; usually built one at a time, not together |
| A short, cited explainer on one specific Microsoft product or offering | `/weave:ms-in-five` | Explicit only; every fact must trace to a live learn.microsoft.com/www.microsoft.com fetch |
| Check, review, QA, proofread, "is this ready to send", "what is wrong with this deck" | `review` | Works on any .docx/.pptx; offers a lighter check on non-Netwoven files |
| Edit a file Weave built earlier ("change slide 3", "swap the client name", "add a risk") | the router or command that built it | Edit loop: read the embedded spec, change it, rebuild, re-check |
| Summarise, explain, compare, answer a question about a document | plain answer | No file |
| Review a counterparty's contract or SOW for risk | legal tools, not Weave | Say so |
| Set up Weave, install what it needs, record your name and role, add a client | `/weave:setup` | Explicit only |
| Report a problem or rate a result | `/weave:feedback` | Explicit only |

## Tie-breaks

- **Document or deck unclear**: a 20-page findings report and a 12-slide
  readout are different deliverables. Ask one question.
- **Client or internal unclear**: the failure modes point opposite ways
  (leaked candour vs a confidentiality page on a wiki note). Ask one question.
- **Deck for an internal audience**: still `deck`, with `internal: true`
  (no confidentiality slide).
- **SOW drafted internally first**: still `sow`. The reader signs it
  eventually.
- **A proposal that includes fees, milestones or rates**: `sow`. A proposal
  without commercials: `document` (or `deck` if slides were asked for).
- **Small edit to a SOW** (a typo, a date, one paragraph): the `sow` router
  in edit mode, not a new SOW. **Change order**: the command.
- **A findings/recommendations document that names Current State/Target
  State**: `/weave:assessment`, not plain `document` -- the recipe's
  finding-table and roadmap sections are the point.
- **A steering-committee or kickoff deck**: `/weave:status` or
  `/weave:kickoff`, not plain `deck` -- the paired doc (report or charter)
  is usually wanted from the same facts even if only the deck was asked for
  first; ask.
- **A migration or cutover procedure a client will hold**: `/weave:runbook`,
  not plain `document` -- it still gets the client confidentiality page and
  Company binding, but the recipe's checklist/sequence/rollback structure is
  the point.
- **`/weave:adoption --recipe adoption-training` fails with "this copy of
  Weave is missing nw_pptx_helpers.py"**: on a packaging surface with a
  per-skill file limit, `adoption`'s own companion files ship only its two
  doc-kind recipes (strategy, comms) -- the deck-kind training toolchain
  didn't fit under the cap. The error is expected there, not a bug; build
  training material through a deck-capable surface instead.

## The one-question rule

A skill asks at most one clarifying question before producing something.
The question offers two or three labelled options with a one-line
consequence each. Everything else becomes a stated assumption under
"I assumed" in the flag list. Never ask for information the source already
contains.

## Words to use and words to avoid in chat

Use: proposal, SOW, change order, rate card, fixed fee, T&M, readout,
findings report, assessment, recommendation, QBR, steerco, kickoff, status
report, roadmap, runbook, SOP, spec, postmortem, decision memo, client-ready,
Netwoven template, Word document, PowerPoint deck, checks.

Avoid: invoke, frontmatter, CLI, flag, argument, stdout, script output,
validator (say "checks"), docx/pptx (say "Word document", "PowerPoint deck"),
JSON, spec (say "the plan for the file" if you must mention it).

Reply shape after every build (max 8 lines under these headings, omit empty
ones): **Check before sending** · **I assumed** · **I left out** · **Open
decisions**. Then one line: `Saved to: <absolute path>`.

## Where the file lands

Default file name `<ClientShort>_<Type>_<YYYY-MM-DD>_v<n>.<ext>` in the
working folder. In Cowork the file also appears in the session's output
list; in Microsoft Copilot Cowork it appears in the OneDrive `Cowork` folder.
Always state the absolute path in the reply.
