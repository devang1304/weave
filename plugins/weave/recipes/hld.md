# Recipe: hld

A High-Level Design document on the document template: the shape of a whole
system or solution, not a single decision. Use `adr` instead when the
source is really about one architecture choice.

## What each section says

- **Overview**: what's being built and for whom, in plain language, before
  any technical detail. A reader who stops here should know the purpose and
  the audience, not yet the mechanism.
- **Requirements**: functional and non-functional requirements kept clearly
  separate, either as labelled subsections or a table -- performance,
  security, and compliance requirements never get folded silently into the
  functional list.
- **Architecture**: the shape of the solution. A `frame` visual (for
  example `capability_map`) or a diagram-worthy table is the right call per
  `{{WEAVE_ROOT}}/recipes/SPEC.md`'s visual types, but `frame` visuals are
  not yet implemented in `build_deliverable.py` as of Phase 2 -- until they
  ship, use a plain table or bullets describing the components and their
  connections as the safe fallback.
- **Components**: the major pieces and their responsibilities, one entry
  per component. Name what each piece does, not just what it's called.
- **Risks and Open Questions** (optional -- drop the section if the source
  presents nothing unresolved): what's still unresolved, stated as open
  questions rather than papered over as settled.

## Asking for what's missing

Only `system_name` is required and usually can't be inferred reliably
without asking -- request it in one question if the source doesn't name the
system or solution clearly.

## Hard rules specific to this recipe

- Never invent a component or integration the source doesn't describe or
  that isn't a clearly labelled Netwoven recommendation. A gap in the
  source is a gap in the document, flagged in chat, not filled in from
  what a system "probably" needs.
- Do not present Risks and Open Questions as though everything else in the
  document is fully settled when it isn't -- an unresolved requirement
  belongs in that section even if it would read more cleanly folded into
  Requirements.
- Keep Architecture and Components distinct: Architecture describes the
  shape and connections, Components describes the pieces and their
  responsibilities. Do not let one section absorb the other.
