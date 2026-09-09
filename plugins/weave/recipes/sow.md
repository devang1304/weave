# Recipe: sow

Spec pattern for the `sow` kind. The template guide
(`references/sow-template-guide.md`) is the ground truth for every table,
field, and placeholder; this file only shows the shapes the builder reads.

## Variant and tables (`sow` object)

```json
"sow": {
  "variant": "milestone",
  "milestones": [{"name": "Discovery", "description": "...", "amount": 42000, "completion": "Week 3"}],
  "rates": [{"role": "Consultant", "location": "onshore", "resources": 1, "weekly_hours": 40, "weeks": 12, "rate": 185}],
  "products": ["SharePoint Online", "Microsoft Purview"],
  "biz_apps": [],
  "deliverables": [["Discovery", "Discovery", "Findings from stakeholder interviews", "Word document"]],
  "timeline": [["Phase 1: Discovery", ["Stakeholder interviews", "Tenant inventory"], "3 weeks"]],
  "expenses_cap": 10,
  "assumptions_remove": ["Govern 365"],
  "client_logo": null,
  "keep_products": false,
  "sections": {
    "executive_summary": [{"id": "e1", "type": "para", "text": "..."}],
    "scope": [{"id": "sc1", "type": "bullets", "items": ["..."]}],
    "out_of_scope": [{"id": "o1", "type": "bullets", "items": ["..."]}],
    "assumptions_add": [{"id": "a1", "type": "bullets", "items": ["..."]}],
    "glossary_add": [["Term", "Definition"]],
    "roles_notes": [],
    "signatures": {"client_name": "", "client_title": "", "netwoven_name": "", "netwoven_title": ""}
  }
}
```

- Milestone uses `milestones` and 5-cell `deliverables` rows in the order
  `[scope section, milestone, deliverable, description, format]`; T&M uses
  `rates` and 4-cell rows in the order `[scope section, deliverable,
  description, format]` (no milestone column). The example above is the
  Milestone shape; leave the other list (`milestones` or `rates`) empty for
  the variant you are not using. Positions are filled left to right with no
  name-based mapping, so a row with the wrong cell count silently shifts
  every value after the gap into the wrong column.
- A missing `amount` or `rate` is a gap: omit the key so the cell stays
  blank. Never write 0.
- `meta.title` is the project name (cover, header, document title);
  `meta.client` the legal name (every Company field); `meta.date` a real date.
- `sections.*` prose uses `para` and `bullets` blocks; headings only at
  level 3 and only inside `assumptions_add`.
- Boilerplate sections (Change Requests, Status Reporting, Payment Schedule,
  Staffing, Signatures preamble, Confidentiality) are not in the spec; the
  builder keeps them verbatim from the template.
