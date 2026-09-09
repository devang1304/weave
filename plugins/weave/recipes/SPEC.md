# Weave content spec and script contract (v2.0)

This file is the contract between the skills, the builder scripts, and the
review/eval tooling. Every generated file carries its spec inside the package
(see "Embedding"), so a later edit is a spec edit followed by a rebuild.

## Contents
- Spec shape (all kinds)
- Blocks (documents and SOW prose)
- Slides (decks)
- SOW section
- Embedding inside .docx / .pptx
- Script contract (`weave.py` subcommands)
- Recipe files (Phase 2)

## Spec shape

```json
{
  "weave": {"spec_version": "2.0", "kind": "doc", "recipe": "generic-doc",
            "base": "NW_Document_Base_2026.docx", "revision": 1,
            "generator": "weave 2.0.0", "built_at": "2026-09-03T10:00:00Z"},
  "meta": {"title": "SharePoint Online Governance Assessment",
           "client": "Contoso Ltd", "client_slug": "contoso",
           "date": "September 2026", "internal": false,
           "author": "", "role": "", "label": null,
           "keep_figure_lists": false},
  "blocks": [],
  "slides": [],
  "sow": null
}
```

- `kind`: `doc` | `deck` | `sow`. A `doc` spec uses `blocks`; a `deck` spec
  uses `slides`; a `sow` spec uses `sow` plus `blocks` inside `sow.sections`.
- `base`: file name under `shared/assets/bases/`. The builder resolves it
  relative to its own location; skills never pass base paths.
- `meta.date`: free text for documents ("September 2026"); SOWs need a real
  date (`2026-09-15` or `9/15/2026`) because the template binds it.
- `meta.internal`: documents only. Removes the confidentiality page and sets
  Company to "Netwoven" (`new_deliverable_docx.py --internal`).
- `meta.label`: `null` = keep the base template's Purview label. Phase 3
  allows a catalog name.
- Every block and slide has a stable `id` (`b1`, `s3`, or any string). Ids
  survive rebuilds; new items get fresh ids.

## Blocks

| type | fields | maps to |
|---|---|---|
| `heading` | `level` (1-4), `text` | NW Heading 1-4 (doc) or Heading 1-3 (SOW; level 4 rejected) |
| `para` | `text` or `runs[]` (`{text, bold, italic}`), optional `style` | Normal (default), or a named template style |
| `bullets` | `items[]`, optional `levels[]` (0-based per item), `numbered` (bool) | List Paragraph with the template's list numbering (`list_num_id`) |
| `table` | `style` ("Netwoven Table 1"/"2"/"3", SOW: "PSO Grid"), `header[]`, `rows[][]`, optional `caption`, `col_widths_in[]`, optional `framework` (a name from the active recipe's `frameworks` map — resolves `style`/`header` from there instead of requiring them inline; a `framework` name with no active recipe or no matching entry is a `BuildError`) | `add_styled_table` + `add_caption` (SEQ field auto-numbers; never put "Table N" in the caption text) |
| `figure` | `image` (path), optional `width_in` (default 6.5), `caption`, `alt` | `add_captioned_picture` |
| `callout` | `text` | Netwoven Table 3 single-cell (doc only) |
| `page_break` | none | hard page break |

Rules the builder enforces (hard errors): heading levels never skip; first
block of a client doc is a level-1 heading; no `Table N`/`Figure N` prefix in
captions; text never contains em/en dashes unless `meta.allow_dashes` is set.

## Slides

```json
{"id": "s3", "layout": "Title Only",
 "title": "Migration readiness is blocked by three permission gaps",
 "body": ["..."], "body_levels": [0],
 "visual": {"type": "chart", "kind": "column", "categories": ["A","B"],
            "series": {"Sites": [312, 88]}, "number_format": "0",
            "palette": "categorical"},
 "source": "Source: tenant inventory, Aug 2026",
 "notes": "speaker notes", "section": "findings"}
```

- `layout`: a key of `LAYOUTS` in `nw_pptx_helpers.py` ("Title Slide for
  Verticals", "Title and Content", "Section Header", "Two Content", "Three
  Column", "Comparison", "Title Only", "Three Stat", "Big Statement with
  Illustration", "Blank"). Layout-specific fields: `subtitle` (cover),
  `left`/`right` (Two Content, Comparison, with `left_head`/`right_head`),
  `columns[]` (Three Column), `stats[]` of `{value, label}` (Three Stat),
  `statement`/`eyebrow`/`image` (Big Statement -- `image` is an optional
  local file path for the illustration slot; omitted or missing leaves that
  slot cleanly removed, not a broken placeholder).
- `meta.cover_images` (optional, deck only): a list of 0-2 local file paths
  for the cover's two portrait picture slots. Omitted or `null` picks up to 2
  at random from `shared/assets/cover-images/` (empty pool = no images,
  same as an explicit `[]`).
- `visual.type`: `chart` (kinds: column, bar, line, pie, stacked), `table`
  (`header`, `rows`, `col_widths_in`, optional `framework` — same resolution
  rule as a doc `table` block above; a deck table has no `style` concept,
  only `header`/`rows`, so a framework used here supplies `header` only),
  `image` (`path`, `alt`), `frame` (Phase 2: `licensing_stack`,
  `capability_map`, `maturity_heatmap`, `offering_ladder`, `phase_roadmap`),
  or absent.
- The builder always: puts the cover first, inserts the Statement of
  Confidentiality slide unless `meta.internal`, keeps Thank You and Closing
  last (`finalize_deck`), sets alt text on every visual, prunes empty
  placeholders.

## SOW section

```json
"sow": {"variant": "milestone",
        "milestones": [...], "rates": [...], "products": [...], "biz_apps": [...],
        "deliverables": [...], "timeline": [...], "expenses_cap": 10,
        "assumptions_remove": [...],
        "client_logo": null, "keep_products": false,
        "sections": {"executive_summary": [blocks], "scope": [blocks],
                     "out_of_scope": [blocks], "assumptions_add": [blocks],
                     "glossary_add": [["Term", "Definition"]],
                     "roles_notes": [blocks],
                     "signatures": {"client_name": "", "client_title": "",
                                    "netwoven_name": "", "netwoven_title": ""}}}
```

The v1 `--json-spec` keys keep their exact shapes (see
`shared/scripts/README-docx-scripts.md`). A missing rate or amount stays a
gap (blank cell, `[TO BE PROVIDED]` in prose, entry in the flag list), never
a fabricated zero. `sections.*` blocks may use `heading` only at level 3.

## Embedding inside .docx / .pptx

- Part `customXml/itemN.xml` with root `<weave:spec
  xmlns:weave="http://netwoven.com/weave/spec/2.0">` containing the JSON in
  CDATA; matching `customXml/itemPropsN.xml` (datastore item with a fixed
  GUID `{7B7C4D6E-9A1F-4C2B-8E3D-WEAVE0000002}` style id generated once per
  file) and relationships from `word/document.xml.rels` or
  `ppt/presentation.xml.rels`. Custom XML parts survive Office re-saves.
- `docProps/custom.xml` gets three short properties only: `Weave_Recipe`,
  `Weave_SpecSha256`, `Weave_Version`. The pptx base has no `custom.xml`;
  the embedder creates it. `MSIP_Label_*` properties and
  `docMetadata/LabelInfo.xml` are never touched.
- Block ids become hidden bookmarks `_weave_<id>` around each block's first
  paragraph in Word; slide ids are written to `p:cSld/@name` as `weave:<id>`.
- `spec_from_file` reads the part, then walks the document and reconciles:
  text edited by hand wins over the stored spec; anything the walker cannot
  express (manual formatting, shapes it does not know) is listed under
  `drift.lost_if_rebuilt` and the skill shows that list before rebuilding.

## Script contract (`weave.py`)

`weave.py` is the only command skills invoke. It lives at
`{{WEAVE_ROOT}}/scripts/weave.py`; the build substitutes the placeholder per
target. Every subcommand accepts `--json` and prints one JSON object on
success or `{"ok": false, "error": "...", "hint": "..."}` on failure with
exit 1. If imports fail, `weave.py` re-executes itself with the interpreter
recorded in `${CLAUDE_PLUGIN_DATA}/env.json` (written by `doctor --install`).

| Subcommand | Arguments | Result |
|---|---|---|
| `doctor` | `[--quiet] [--install] [--json]` | environment report; `--install` creates the venv and pins deps |
| `build` | `--spec FILE --out FILE [--replace] [--client-slug S] [--recipe R] [--json]` | builds the file, embeds the spec, runs validation, returns `{out, revision, validation:{hard_fails, warnings}, flags}` |
| `spec-from-file` | `FILE [--out spec.json] [--json]` | reconstructed spec plus `drift` |
| `review` | `FILE [--quick] [--rubric default|sow|deck] [--apply] [--report out.md] [--json]` | scored report (see plan §1.4) |
| `validate` | `FILE [--type doc|sow|deck] [--internal] [--client NAME] [--render] [--json]` | passthrough to `validate_deliverable.py` |
| `read` | `PATH [--out md] [--tables-json] [--json]` | normalized markdown with provenance markers |
| `client` | `list` / `show SLUG` / `set SLUG key=value ...` | client profile CRUD under `${CLAUDE_PLUGIN_DATA}/clients/` |
| `label` | `FILE --show` (Phase 2) / `--expect NAME` / `--apply NAME` (Phase 3) | Purview label read/check/apply |
| `telemetry` | `event --type T ...` / `flush` / `status` | appends to `<data>/telemetry/events.jsonl`; `flush` reports what's queued (no send endpoint exists yet); only `/weave:feedback` calls this today -- passive per-build/review events (GOVERNANCE.md Sec 5's full event schema) are not wired into the other skills |

Output naming when the skill does not specify `--out`:
`<ClientShort>_<Type>_<YYYY-MM-DD>_v<revision>.<ext>` in the current working
directory; the result JSON always carries the absolute path.

## Recipe files (Phase 2)

Two files per recipe, both under `shared/recipes/`: `<name>.json` (machine-
checked constraints, read by `build_deliverable.py --recipe <name>`) and
`<name>.md` (prose guidance for Claude: what each section should say, how to
ask for missing inputs; the only file a command skill's SKILL.md points at
directly). Phase 1 shipped three near-empty recipes with only the `.md` half
(`generic-doc`, `generic-deck`, `sow`) so every skill already has a recipe
name to cite; their `--recipe` value was accepted but not loaded. Phase 2
adds the `.json` half and makes `--recipe` load and enforce it.

### `<name>.json` shape

```json
{
  "recipe": "status-report", "kind": "doc", "base": "NW_Document_Base_2026.docx",
  "title_pattern": "{client} status report: {period}",
  "inputs": [{"key": "period", "required": true, "ask": "Which reporting period?"}],
  "sections": [
    {"id": "summary", "heading": "Executive summary", "level": 1, "required": true},
    {"id": "progress", "heading": "Progress this period", "level": 1, "required": true,
     "framework": "rag-table"}
  ],
  "frameworks": {
    "rag-table": {"type": "table", "style": "Netwoven Table 2",
                  "header": ["Workstream", "Status", "This period", "Next period"]}
  },
  "validator": {"require_headings": ["Executive summary", "Next period and decisions needed"],
                "owners_and_dates_in": "next"},
  "companion": {"recipe": "steerco", "share": ["client", "date", "period"]}
}
```

- `recipe`/`kind`/`base` must match the spec's own `weave.recipe`/`weave.kind`/
  `weave.base` exactly, or `build_deliverable.py` raises a `BuildError` naming
  the mismatch (a recipe loaded against the wrong kind of spec is a caller
  bug, not a content problem, so it is a hard build failure, not a
  `validation.hard_fails` entry).
- `title_pattern` (optional): when the spec's `meta.title` is empty, fill it
  by substituting `{client}`, `{date}`, `{period}` (and any other `meta.*` or
  top-level `inputs[].key` value) into this template. A spec-supplied title
  always wins; the pattern is a default, not an override.
- `inputs` (optional, informational): names the values a command's SKILL.md
  should collect and, when one is missing, the exact one-question prompt to
  ask. `build_deliverable.py` does not itself collect input — that is the
  skill's job before it writes the spec — but a `required: true` input whose
  `meta.*` counterpart is empty becomes a `validation` warning (not a hard
  failure) naming which input is missing, so a spec built without asking
  still surfaces the gap in the flag list instead of shipping silently thin.
- `sections` (optional): each entry names a section by `id`. For a `kind:
  doc` (or the `sow` prose blocks) recipe, `id` must match a `blocks[].id`
  (case-sensitive), or the entry may instead give `heading` text
  (case-insensitive match against the block's rendered heading text) at the
  given `level`. For a `kind: deck` recipe, matching is by `id` against
  `slides[].id` only — slides have no heading/level concept, so a deck
  recipe's section entries always set `heading` and `level` to `null` (see
  `steerco.json`). `required: true` sections missing from the spec become a
  `BuildError` (the file is not written at all — a required section is a
  caller bug in the same way a missing `meta.title` is, not content the
  author can fix after the fact by reading a flag list). `required: false`
  (or omitted) sections are informational only, used for
  `owners_and_dates_in` resolution and the companion `share` list.
- `frameworks` (optional): named table shapes a section can request via its
  matching `blocks[].framework` key (a new optional key on a `table` block,
  SPEC.md "Blocks" table) instead of spelling out `style`/`header` inline.
  `build_deliverable.py` resolves `framework` to the recipe's definition
  before falling back to whatever `style`/`header` the block already
  supplies; a `framework` name with no matching entry in the recipe (or no
  recipe loaded at all) is a `BuildError`.
- `validator` (optional): `require_headings` (list of heading texts that
  must all be present, checked the same way `sections[].required` is, kept
  as a separate list for readability when it duplicates `sections`);
  `owners_and_dates_in` (a section `id`): that section's block text must
  contain something matching a name-like token and something matching a
  date-like token (a deliberately loose heuristic — this is a `validation`
  warning, never a hard failure, since "does this read as an owner and a
  date" is a judgement call the flag list surfaces, not one code should
  silently guess at or block on).
- `companion` (optional, informational only): names a sibling recipe meant
  to be built from the same underlying facts in the same skill turn (a
  status report and its steerco deck), and which `meta`/`inputs` keys the
  two share. `build_deliverable.py` does not build the companion itself —
  each recipe is still one `--spec`/`--out` call — this field only lets a
  command's SKILL.md and `<name>.md` say "build both from one set of
  answers" without repeating the shared inputs' prompts twice.

### `<name>.md` shape

Unstructured prose: what each section should say (beyond the generic
recipe pattern), the exact wording for each `inputs[].ask` prompt, and any
recipe-specific hard rules. Keep it as short as the generic recipe files
(`generic-doc.md`, `generic-deck.md`, `sow.md`) — a command's SKILL.md is
already ~25 lines that hands off here for everything recipe-specific.
