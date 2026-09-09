# Recipe: adr

An Architecture Decision Record on the document template: one decision,
stated and justified, not a whole design. Normally short and normally
internal (`meta.internal: true`) unless the source clearly indicates this
record is being shared externally with a client.

## What each section says

- **Context**: the problem and the forces at play (constraints, competing
  priorities, what prompted this decision) stated without yet giving the
  answer. A reader who stops here should understand why a decision was
  needed, not what it was.
- **Decision**: the single chosen answer, stated plainly in one paragraph.
  This is the governing thought of the document; do not hedge it or bury it
  under the reasoning that led here.
- **Alternatives Considered** (optional -- drop the section if the source
  doesn't discuss any): what else was weighed and why it lost. Each
  alternative gets a real reason, not "was not chosen."
- **Consequences**: both the upside AND the real tradeoffs or costs
  accepted -- never only the positive case. If the decision creates new
  work, risk, or a dependency, say so here.

## Asking for what's missing

Only `decision_title` is required and usually can't be inferred reliably
without asking -- request it in one question if the source doesn't make the
decision's subject obvious from its first lines.

## Hard rules specific to this recipe

- Never invent a "Decision" the source doesn't actually state. A question
  the source discusses but has not yet resolved is not an ADR yet -- say so
  instead of drafting one, even if the surrounding material reads like it is
  heading toward an answer.
- Consequences must include at least one real tradeoff or cost, not only
  benefits restated from the Decision section. A decision with no downside
  at all is a sign the section was not thought through, not a sign the
  decision is free.
- `meta.internal` is usually `true` for this recipe; only set it `false`
  when the source clearly indicates external, client-facing distribution.
