# Changelog

All notable changes to Netwoven Weave. Format follows Keep a Changelog;
versions follow semantic versioning. The plugin `version` pins what users
receive, so every user-visible change bumps it.

## [2.0.1] - Unreleased

### Fixed
- The release pipeline published the Claude marketplace artifacts only after
  the Microsoft 365 packaging step, so `atk package` rejecting the 18 MB app
  package (its cap is 10 MiB for the whole package) blocked the entire
  release, marketplace pull request included. The Claude artifacts now
  publish first, and the two M365 steps may fail without failing the release.

### Changed
- Copilot Cowork delivery (ADR-001, revised 2026-09-09): the thin skills
  package plus the MCP connector is the Cowork build path. Cowork's
  self-contained skill folders would carry `NW_Document_Base_2026.docx` nine
  times (13.3 MB) against the 10 MiB cap, so no M365 package ships until the
  server exists. Claude surfaces are unaffected.

## [2.0.0] - 2026-09-09

Tagged but never published: the first release-pipeline run failed in the
Microsoft 365 packaging step, and the fix above ships as 2.0.1. Everything
below is what 2.0.1 delivers.

### Changed
- Plugin renamed from `netwoven-deliverables` to `weave` (display name
  Netwoven Weave); commands read `/weave:...`. Marketplaces map the old
  name with `renames`.
- Skills reorganised into four implicit routers (`document`, `deck`, `sow`,
  `review`), a hidden `house-rules` skill, and explicit commands (`setup`,
  `feedback`). `client-deliverable` and `internal-doc` merge into
  `document` plus `deck`.
- Routing lives once in `references/routing.md`; descriptions follow one
  template and stay under 1,024 characters.
- One source, two build targets: Claude (marketplace layout, resources once
  at plugin root) and Microsoft Copilot Cowork (per-skill companions within
  the importer's limits).

### Added
- Usage events for every skill call: `weave.py` records one event after each
  build, review, validate, read and read-back; the plugin's hooks record a
  `skill_invoked` event whether Claude chose the skill or the user typed it;
  `weave.py telemetry flush` sends the queue to Netwoven's usage list (a
  SharePoint list, through Microsoft Graph, signed in with a device code by
  `/weave:setup`) at the start of each session. Nothing is sent until the
  plugin team configures the list.
- Preference memory (`weave.py memory`): hand edits seen when a Weave file is
  read back become patterns, a pattern seen three times becomes a preference,
  and the build skills read the preferences before writing a spec. Local
  only; `WEAVE_MEMORY=0` turns it off.
- `tools/template_index.py` indexes every template asset reference, checks
  hashes against `shared/manifest.json`, fails on residue, and writes
  `docs/TEMPLATE-INDEX.md`; `tools/release_check.py` and CI run it.
- `tools/mirror_marketplace.py` assembles the marketplace tree the release
  pipeline publishes to GitHub for Claude's organization marketplace and
  refuses a publish without a version bump.
- Staged `azure-pipelines.yml` (Validate, Package, Publish) and the design
  documents `docs/adr/ADR-001-stateless-mcp-on-azure-functions.md`,
  `docs/PLATFORM-FEATURE-MAP.md`, `docs/DISTRIBUTION-PIPELINE.md`,
  `docs/TELEMETRY-FEEDBACK-MEMORY.md`, `docs/COST-ANALYSIS.md`.
- Content spec (`recipes/SPEC.md`) embedded in every generated file; edit
  an existing deliverable by changing its plan and rebuilding.
- `weave.py` dispatcher, `doctor.py` environment check and installer,
  `read_source.py` ingestion for .docx/.pptx/.xlsx/.pdf/.vtt, `review_deliverable.py`
  scored review with safe fixes, client profiles, reviewer-of-record metadata.
- Evaluation harness (`evals/`) with trigger sets per router, migrated
  output cases, deterministic asserts, seeded-defect review tests, and
  release gates.
- GOVERNANCE.md: acceptable use, data handling, labels, disclosure,
  telemetry notice.

### Fixed
- The three template guides pointed example commands at `assets/<base>`; the
  bases live under `assets/bases/`.
- `spec-from-file` usage events always recorded `kind`/`recipe` as null;
  `passive_fields()` now reads them from the embedded spec's nested shape.
- The plugin-settings "Share anonymised usage events" toggle had no code
  path; `enabled()` now reads it alongside `WEAVE_TELEMETRY=0`.
- `memory.py promote()` ignored `WEAVE_MEMORY=0` (only `record()` checked
  it), so a pattern already at threshold could still be silently promoted
  after the user turned memory off; `promote()` now self-gates the same way.
- Two concurrent `telemetry.py flush` runs (e.g. two sessions' SessionStart
  hooks) could race and drop an event that was never sent; concurrent
  `memory.py promote`/`forget` could race and silently revert a "forget".
  Both now hold a small cross-platform file lock around the read-modify-write.
- One malformed line in `events.jsonl` (a crash mid-append) used to abort
  the whole flush before any good lines were dropped, wedging the queue on
  that line forever; bad lines are now parsed best-effort and dropped.
- `telemetry.py login` blocked a single tool call for the device code's full
  ~15-minute expiry window; split into `login-start` (returns immediately)
  and `login-poll` (one poll attempt), matching how `/weave:setup` narrates it.
- `telemetry_sink.py`'s device-code and token-refresh requests tagged their
  User-Agent `Weave/dev` instead of the real plugin version (only `send()`
  passed it); `save_config()` wrote `sink.json` directly instead of through
  the atomic tmp-file-then-replace `save_tokens()` already used.
- `plugin.json`'s `userConfig` was missing `telemetry_tenant_id` and
  `telemetry_client_id`, so the documented plugin-settings sink configuration
  never actually worked for Claude Code users.
- The feedback skill's local-file fallback (for when telemetry isn't
  installed) was dropped in the telemetry rewrite, and the remaining text
  claimed the report "stays on this device" even in that exact branch,
  where nothing was recorded anywhere; the fallback is restored.
- `templates/AI-Optimized-Templates-2026`'s Word/SOW guides, `visual-standards.md`
  and `deck-storyline.md` cited `writing-method.md` under two invented,
  nonexistent names (`writing-standards.md`, `analysis-method.md`).
  `tools/template_index.py` gained rule R6, which checks that a reference
  guide or template-workshop `.md` citation resolves to a real file — residue
  the binary-asset rules (R1-R4) cannot see.

### Unchanged by design
- The four 2026 template bases and their AI-optimisation (see
  `templates/*/CHANGELOG.md`); Purview labels untouched; SOW heading
  harmonisation and the deck confidentiality slide as in 1.0.0.

## [1.0.0] - 2026-09-03 (as netwoven-deliverables)

- Rebuild on the 2026 templates; three skills; nine output evals passed;
  maximum-effort code review with 13 fixes. See `docs/EVALUATION-v1.md`.
