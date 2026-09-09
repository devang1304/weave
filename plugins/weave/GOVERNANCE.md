# Governance: acceptable use, data handling, and disclosure

This page is the reference for anyone at Netwoven using Weave and for the
team that maintains it. Items marked *draft for counsel* need legal review
before they are relied on externally.

## 1. Acceptable use (one page)

**What Weave is for.** Building and checking Netwoven deliverables: Word
documents, PowerPoint decks, statements of work, and (later) spreadsheets
and collateral, on Netwoven templates, from material you are entitled to
use.

**Approved surfaces.** Cowork in Claude Desktop, Claude Code in Claude
Desktop, and Microsoft Copilot Cowork, each on Netwoven's own tenant or
plan. Public chat tools are not approved for client material.

**What may go in.** Client material Netwoven holds under an NDA or
engagement letter, inside the approved tenant, unless that client's terms
say otherwise (the client profile records this as `ai_permitted`).

**What never goes in.** Credentials, keys, tokens; personal data beyond
names and roles; another party's confidential material; anything marked
"do not share". Weave refuses these and says why.

**Review before send.** A named Netwoven consultant reads every deliverable
before it leaves. Weave records "Draft prepared with Netwoven Weave;
reviewed by ____" and `/weave:review` fills in the reviewer.

**Verify numbers.** Anything Weave flags as inferred, hedged, or unsourced
is checked against the primary source before it ships.

**Report problems.** `/weave:feedback`, or the Weave channel in Teams.

**Acknowledgement.** Staff confirm they have read this page when they
first run `/weave:setup`; the acknowledgement is recorded in the usage list.

## 2. Data handling in the skills

The operative rules are in `references/guardrails.md` and are read by
every skill before it opens a source. In short: refuse credentials and
sensitive personal data; warn on internal candour and other clients'
pricing; stop on restricted labels or an `ai_permitted: false` profile.

## 3. Sensitivity labels

Netwoven's templates carry Purview labels ("Public" on the document
template, "General Business" on the SOW templates). Weave never removes or
lowers a label. Phase 2 reads the label of attached sources and warns or
stops; Phase 3 can apply a higher label from Netwoven's catalog. Until then
the send checklist reminds the author to set the label in Office, because
a signed-price SOW labelled "General Business" is not right.

## 4. Disclosure of AI assistance

Internally: expected and never penalised. Reviewers assess the deliverable,
not the method. Externally: *draft for counsel*, suggested engagement-letter
language:

> Netwoven may use approved AI-assisted drafting tools under Netwoven's
> supervision to prepare deliverables. All deliverables are reviewed and
> approved by a named Netwoven consultant before release. Client materials
> are processed under Netwoven's enterprise agreements, are not used to
> train models, and are retained per the retention terms in Schedule X.

Client profiles carry `disclose_ai: true` when a client requires a visible
note on deliverables.

## 5. Usage data (telemetry) and how to turn it off

Weave records one event per build, review, validation, source read, or
read-back of a Weave file, one event when a Weave skill is invoked, and one
per feedback report:

`event_id, timestamp, plugin_version, surface, type, skill, invocation,
recipe, kind, outcome (pass/fail), hard_fails, warnings, score, duration_s,
source_types, rating (feedback only), text (feedback only), user_hash,
client_hash`

`user_hash` and `client_hash` are salted SHA-256 values; the salt rotates
monthly, so events never join across months. Events never contain prompt
text, file names, paths, titles, client names, or session identifiers.
Feedback events carry no client hash. Clients with `ai_permitted: false`
produce no client hash.

Events are written on the device first (`~/.weave/telemetry/events.jsonl`,
or the plugin data folder) and sent to Netwoven's usage list, a SharePoint
list in the `netwoven_weave` Microsoft 365 group, at the start of the next
session, once the plugin team has configured the list and the user has
signed in through `/weave:setup` (Microsoft sign-in with a device code; the
sign-in is remembered per device). Until both are true, events stay on the
device. Turn passive events off in the plugin settings ("Share anonymised
usage events") or with `WEAVE_TELEMETRY=0`; a feedback report you file
yourself is always recorded. `/weave:setup` shows this notice once. Design
and setup: `docs/TELEMETRY-FEEDBACK-MEMORY.md`.

## 6. Retention and platform settings

Anthropic Team and Enterprise plans do not train on Netwoven's data;
standard retention is 30 days; zero-data-retention is available on request
through the Netwoven account. Recommended admin settings: a short local
transcript retention (`cleanupPeriodDays`), feedback command disabled
(`DISABLE_FEEDBACK_COMMAND=1`) on the Microsoft-hosted variant. Microsoft
Copilot Cowork processes data inside the Microsoft 365 tenant with Anthropic
as a sub-processor. Client profiles and logos live under the user's Weave
data folder with owner-only permissions and are never committed to a repo.

## 7. Preference memory stays on the device

When you edit a Weave-built file by hand and later ask Weave to read it back,
Weave notes the difference: a section you removed, a word you replaced, a
slide you dropped. The same change seen three times becomes a preference
the build skills apply. Those notes live only in your Weave data folder
(`~/.weave/memory/`), hold at most a few words of your edit or one heading,
are never sent anywhere, and are never included in usage events.
`weave.py memory show` lists them, `weave.py memory forget KEY` drops one,
and `WEAVE_MEMORY=0` turns the feature off.

## 8. Security of the plugin itself

Before each release: no secrets in the repository (scanned), dependencies
pinned, template bases hash-verified, user-supplied strings never placed in
regular-expression replacements, no shell string execution, `claude plugin
validate --strict` clean, `docMetadata/LabelInfo.xml` byte-identical to the
base in every generated file. The checklist lives in
`docs/SECURITY-CHECKLIST.md` in the repository.
