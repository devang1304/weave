---
title: Netwoven SOW Templates 2026 (AI Reference Guide)
covers: NW_milestone_SOW_template_2026.dotx, NW_TM_SOW_template_2026.dotx, their _v2.dotx cleanups, NW_SOW_Milestone_Base_2026.docx, NW_SOW_TM_Base_2026.docx
last_verified: 2026-09-02
---

# Netwoven SOW templates 2026

This guide covers both Statement of Work templates: the fixed-fee **Milestone**
variant (MS) and the **Time and Materials** variant (TM). They share one
lineage and about 95 percent of their content; the differences are listed in
§8. Every value was read from the unpacked XML of the originals. Values that
the approved harmonization changes are shown both ways. Anything that needs
a rendered page to confirm is marked "to be confirmed by render".

**Files this guide describes:**

| File | Purpose |
|---|---|
| `NW_milestone_SOW_template_2026.dotx`, `NW_TM_SOW_template_2026.dotx` | Netwoven's originals, unmodified |
| `NW_milestone_SOW_template_2026_v2.dotx`, `NW_TM_SOW_template_2026_v2.dotx` | Cleaned and harmonized (§14) |
| `NW_SOW_Milestone_Base_2026.docx`, `NW_SOW_TM_Base_2026.docx` | `_v2` with the content type flipped so `python-docx` opens them. The `sow` skill's scripts use these (`assets/`) |

Related references: `brand-tokens.md`, `writing-method.md` (SOW register),
`word-template-guide.md` (the general document template, a different
lineage), `visual-standards.md` §10.3 (why fee tables stay tables).

---

## 1. Lineage: not the document template

The SOW templates are **not** derived from `NW_document_template_2026.dotx`.
They share its theme file (stock Office, `accent1=#4472C4`, Calibri Light /
Calibri), its wordmark bytes (`image7.png` here equals the document's
`image1.png`, 236 x 34), its five pillar icons, its `#145CA4` 3pt header and
footer rules, its `#404040` / `#F2F2F2` / `#C00000` cover constructs, its
`#FF0000` dotted confidentiality box, and the same three data-binding
`storeItemID`s. Everything else differs:

- Headings use **redefined built-in `Heading1/2/3`**, not `NW Heading`.
  `NW Heading 1` to `4` exist in the style sheet with zero uses and with
  small caps switched **off** (the opposite of the document template). They
  are removed from the bases; the validator forbids them in a SOW.
- One section, no `titlePg`, so the header and footer print on the cover.
- The cover is an anchored text box, not a `Cover Pages` gallery SDT.
- The TOC is titled "Contents" and has no List of Figures or Tables.
- The confidentiality year is a live `DATE \@ "yyyy"` field.
- 46 (MS) or 35 (TM) Company content controls carry the client name into
  the legal text.

Do not run `new_deliverable_docx.py` on a SOW base: it keys on `NWHeading1`
and the "Document Revision History" page, neither of which exists here.

---

## 2. Brand subset used by the SOWs

| Token | Hex | Where in the SOW |
|---|---|---|
| Netwoven rule blue | `#145CA4` | Header bottom rule and footer top rule (3pt); `Netwoven Table 1` header fill (timeline table); cover borders |
| Heading blue | `#2F5496` | `Heading1/2/3`, `TOC1` (theme `accent1` shade `BF`; never retheme) |
| Heading 2 rule | `#AEAAAA` | `Heading2` top border (kept by harmonization) |
| Navy rule | `#000080` | 1.5pt bottom rule under the shipped `Heading1`; **removed** by harmonization |
| Footer label | `#0070C0` | Bold "Netwoven Confidential" second footer line; roles table header fill |
| Biz Apps header | `#92D050` | Header row of the Biz Apps table |
| PSO Grid header | `#D9D9D9` | Deliverables table header row and first column |
| Products table | `#9CC2E5` borders, `#DEEAF6` bands, `#2E74B5` text | `Grid Table 6 Colorful Accent 5` (theme `accent5` tints) |
| TM TOC2 | `#002060` | T&M variant only, bold level-2 TOC entries |
| Cover | `#404040` pillar cells, `#F2F2F2` labels, `#C00000` and `#145CA4` 3pt borders | Inside the cover text box (both Choice and Fallback copies) |
| Confidentiality box | `#FF0000` dotted 0.5pt | Around the three legal paragraphs |
| Caption / small print | `#44546A` | `Caption`, `No Spacing` |

Fonts as shipped: Segoe UI body; `Heading1/2/3` Segoe UI bold small caps;
`Body Text` **Verdana** (with a bogus `w:cs="Sendnya"` hint); `Netwoven
Table 1/2` header rows **Sitka Banner** (a font-substitution defect); `2-Level
Legal` Times New Roman 10pt; `Normal-Bullet` Arial; checkbox glyphs MS
Gothic. The base changes `Sitka Banner` → `Segoe UI Semibold` and `Sendnya`
→ `Segoe UI`; `Verdana` in `Body Text` is left as shipped (visible, not
approved for change).

---

## 3. Style catalog, post-harmonization

The user approved harmonizing SOW headings to the document template's
values. The table gives the shipped value and the base value.

| Style ID | Shipped (both variants unless noted) | In `_v2` and the bases |
|---|---|---|
| `Heading1` | Segoe UI (via `Normal`) **bold** small caps 18pt `#2F5496`; `pageBreakBefore`, `keepNext`, widow control off; 1.5pt `#000080` bottom rule; after 6pt, single; **MS only** `ind left=720` (0.5in); numbered `%1.` (`numId 56` MS, `numId 48` TM); outline 0 | Segoe UI Semibold (`ascii`/`hAnsi`), `cs` Segoe UI, `eastAsiaTheme majorEastAsia`; bold removed; small caps kept; `#2F5496` kept; 22pt (`sz 44`); rule removed; MS indent removed; page break, keep-next, numbering, spacing, outline kept |
| `Heading2` | bold small caps 16pt `#2F5496`, `#AEAAAA` 0.5pt top rule, before 2pt, numbered `%1.%2` | Segoe UI Semibold 18pt, small caps, `#AEAAAA` rule kept |
| `Heading3` | bold small caps 16pt `#2F5496`, numbered `%1.%2.%3` | Segoe UI Semibold 14pt, small caps **off**, colour `auto` (black) |
| `Heading1Char` to `Heading3Char` | linked character styles mirroring the shipped values | mirrored to the new values |
| `Heading4` | Calibri Light italic `#2F5496` on **`numId 1`**, a different list from `Heading1/2/3` | unchanged. Never use `Heading4` in a SOW: its numbering is a separate list and the validator rejects it |
| `TOCHeading` | based on `Heading1`, outline 9; **MS** has `numPr numId=0` (unnumbered, sections 1 to 9); **TM lacks it**, so the TM TOC title takes number 1 and every section is off by one (2 to 10) | TM gets `numPr numId=0` |
| `NWHeading1` to `NWHeading4` | present, zero uses, small caps off, `u color 145CA4` | removed |
| `Normal` | Segoe UI 11pt, after 8pt, line 1.08 | unchanged |
| `BodyText` | Verdana 11pt, `cs` Sendnya, left indent 0.5in, before 2pt after 3pt, single. Used by the Executive Summary sentences, the MS Scope of Work outline, the Signatures effective-date line | `cs` → Segoe UI; otherwise unchanged |
| `ListParagraph` (alias "Bullet") | Segoe UI 11pt, indent 0.5in, contextual; bullets via direct `numPr` (most often `numId 19`, Symbol bullet) | unchanged |
| `ListBullet` | Calibri 10pt, `numId 18` | unchanged (semi-hidden) |
| `Normal-Bullet` | Arial, `numId 2` (Wingdings arrow), justified | unchanged |
| `2-LevelLegal1` to `6` | Times New Roman 10pt, `numId 14` (`1.`, `1.1`, `(a)`, `(i)`) | kept: referenced from numbering levels (Track 1 correction 8) |
| `Block`, `Block-Body`, `BodyTextIndent` | Signature block lines and "This offer expires" line | unchanged |
| `Caption` | Segoe UI 9pt italic `#44546A` | unchanged |
| `TOC1` | Segoe UI 12pt bold `#2F5496`, tabs 0.31in and right leader 6.49in | unchanged |
| `TOC2` | MS: Segoe UI, `cs` Sendnya, kern 28, tabs 0.61in / 6.49in, indent 0.15in. TM: bold `#002060` | `cs` fixed; colour left as shipped |
| `TOC3` | indent 0.31in | unchanged |
| `paragraph`, `normaltextrun`, `eop`, `pagebreaktextspan`, `ui-provider` (TM) | Web-paste artefacts used only in the Appendix | removed with the Appendix |
| `Comment*`, `Balloon*`, `Revision`, `Mention`, `UnresolvedMention`, `FollowedHyperlink`, `PlainText*`, `Quote*`, `KCITableNormal`, `ListTable4-Accent1`, `PlainTable5`, `TOC4`, `TableofFigures`, `NetwovenTable21` | zero references | removed (zero-reference assertion before each) |
| `PlainTable3` | zero direct uses, but the `basedOn` parent of `NetwovenTable3` | kept (build guard; recorded deviation from the plan's removal list) |

Table styles:

| Style ID | Header row | Body | Borders | Used by |
|---|---|---|---|---|
| `TableGrid` | none | Segoe UI 11pt | `auto` (black) 0.5pt full grid | Fee table (MS 13x4), rate card (TM 30x7) |
| `PSOGrid` | bold, fill `#D9D9D9`, vertically centred, `tblHeader` (repeats on page break), keep with next | 10pt, cell padding 0, `cantSplit` rows | `text1` (black) 0.5pt grid | Deliverables table (MS 2x5, TM 2x4). First column also `#D9D9D9` bold |
| `NetwovenTable1` | fill `#145CA4`, white 12pt, shipped font **Sitka Banner** → base **Segoe UI Semibold** | Segoe UI | `#262626` 0.5pt | Timeline table (Phase / Description / Duration) |
| `NetwovenTable2` | same header fix | banded `#F2F2F2` | `#BFBFBF` | Available, unused |
| `NetwovenTable3` | Calibri bold caps `#145CA4` | banded | none | Available, unused |
| `GridTable6Colorful-Accent5` | bold, 1.5pt `#9CC2E5` bottom border | `#2E74B5` text, `#DEEAF6` bands | `#9CC2E5` 0.5pt | Microsoft Products checklist (26x2) |
| (no style) | fill `#0070C0` | | direct borders | Roles table (8x4) |
| (no style) | fill `#92D050` | | direct borders | Biz Apps table (8x2) |

---

## 4. Page, header and footer

- Letter, 1in margins, header/footer inset 0.4in, one section, page numbers
  start at 1, **no `titlePg`**: the header and footer print on the cover.
- The final `sectPr` carries its own `headerReference` (`header1.xml`) and
  `footerReference` (`footer1.xml`), so deleting body content never loses
  them.
- `header1.xml`: a Title content control (cached `<<Project Name>>`), then
  115 literal spaces, then the wordmark anchored at `800324 x 115301` EMU
  (0.875in by 0.126in); 3pt `#145CA4` bottom border on the paragraph.
- `footer1.xml`, paragraph 1: Company content control (cached `<<Client>>`)
  / centre `ptab` / `PAGE` field / right `ptab` / Publish Date content
  control (date picker, `dateFormat M/d/yyyy` already present, cached
  `MM/DD/YYYY`); 3pt `#145CA4` top border. Paragraph 2: "Netwoven
  Confidential", bold `#0070C0`. The validator requires this footer text.
- `settings.xml` ships `documentProtection edit=readOnly formatting=1
  enforcement=0` (declared but not enforced); removed in the base.
  `updateFields=true` is added.

---

## 5. Placeholder and content-control map

**Every `<<Client>>`, `<<Project Name>>` and `MM/DD/YYYY` in the file is the
cached text of a bound content control**, including the Executive Summary
effective-date sentence and the Signatures effective-date line (Track 1
correction 1). Builders therefore rewrite the cached `w:t` inside every bound
SDT and set the three stores; a plain text substitution would miss nothing
but would also leave the stores stale.

| Control (alias) | Store | storeItemID | MS occurrences | TM occurrences | Cached text |
|---|---|---|---|---|---|
| Title | `core.xml dc:title` (shipped value `<<Project Name>>`) | `{6C3C8BC8-…}` | 3 in body (cover Choice and Fallback, Scope intro) + 1 header | same | `<<Project Name>>` |
| Company | `app.xml Company` | `{6668398D-…}` | 46 in body + 1 footer | 35 in body + 1 footer | `<<Client>>` |
| Publish Date | `customXml/item1.xml PublishDate` | `{55AF091B-…}` | 4 in body (cover x2, Executive Summary, Signatures) + 1 footer | same | `MM/DD/YYYY` |

Body counts include both the `mc:Choice` and the `mc:Fallback` copy of the
cover text boxes.

Other placeholders (not content controls):

| Text | Where | Builder action |
|---|---|---|
| Yellow-highlighted author instructions (31 runs MS, 22 TM) | Executive Summary fill-in sentence, `<<Insert…>>` prompts, "Post Go Live support for X weeks", roles-table note, timeline note, MS total line, MS travel `X%` | `strip_highlighted_instructions`: whole instruction paragraphs deleted; partially highlighted paragraphs handled by rule (Executive Summary → empty `BodyText` slot; MS "Post Go Live…" → deleted; MS plain total line → deleted; MS travel `X%` → `[__]%` until `--expenses-cap`) |
| `<<Insert Glossary Terms as necessary>>` | Glossary (highlighted in TM, not in MS) | deleted |
| `<<Insert, as necessary>>` (MS Project Scope, three lines) | Assumptions › Project Scope | deleted |
| "(Please use the following link for Resource Rate Card and delete this line after use)," (MS) / "Note: Please use the rate chart using the following link, …" (TM) | Project Cost Estimate | deleted together with hyperlink `rId28` (a 2023 SharePoint tenant path) |
| "Please Click the Executive Summary link ( to update Estimate)" | Project Cost Estimate | deleted |
| `<<Put bulleted list of Requirements the project would be addressing>>`, "Scope 1", "Scope 2" (TM) / Milestone outline (MS) | Scope of Work | replaced by real scope |
| "Please Insert Client Logo" | Cover, `Rectangle: Rounded Corners 5` (`1390650 x 600075` EMU, 1.52in by 0.66in, `roundRect`) | `set_client_logo(path)` fills it or `remove_client_logo_box()` removes it (both Choice and Fallback). Without `--client-logo` the box is removed |
| `Author Name` style fields | none in SOWs | n/a |

Cover structure (`body[0]`, one `NoSpacing` paragraph): `Text Box 1`
(`6600825 x 8724900` EMU, 7.22in by 9.54in) holds the title, pillar band,
hero photo `image1.jpeg` (1536 x 1294) and icons; `Text Box 4`
(`1057275 x 685800` EMU) holds "Prepared for: `<<Client>>`" and the date;
the rounded rectangle holds the logo prompt. No MISA badge, no `Cover
Pages` SDT.

---

## 6. Table map with fields and bookmarks

Body indexes refer to the shipped originals.

| Table | Variant / body index | Style | Size | Header row | Fields, bookmarks | Builder helper |
|---|---|---|---|---|---|---|
| Microsoft Products | MS 26, TM 28 | `GridTable6Colorful-Accent5` | 26 x 2 (grid 7105 / 2250) | "Microsoft Products being utilized in this SOW", "Checkbox" | 25 `w14:checkbox` SDTs in column 2 (§7) | `tick_product(name)`; `drop_unticked_products()` unless `--keep-products` |
| Biz Apps | MS 29, TM 31 | none, header fill `#92D050` | 8 x 2 | "Serial No", "Product Name" | row 1 "D365 Customer Engagement", rows 2 to 7 "Product 2" to "Product 7" | `set_biz_apps([...])` (drops unused rows) |
| Deliverables | MS 46 | `PSOGrid` | 2 x 5 | "Scope Section", "Milestone", "Deliverable", "Description", "Format" | second row empty | `clear_sample_rows`, `add_deliverable_row(...)` |
| Deliverables | TM 44 | `PSOGrid` | 2 x 4 | "Scope Section", "Deliverable", "Description", "Format" | sample row "Project Initiation / Infrastructure & Access Requests / … / Document" | same |
| Roles | MS 105, TM 81 | none, header fill `#0070C0` | 8 x 4 | "Role", "Onshore", "Offshore", "Role Description" | rows Program Manager, Project Manager, Technical PM, Architect, Business Analyst, Developer, QA with "X" marks | rows removed by name via `--json-spec` |
| Timeline | TM 136 (in Timeline); **MS 358 (inside the Appendix)** | `NetwovenTable1` | 5 x 3 | "Phase", "Description", "Duration" | four sample phases (Requirements finalization and design; Implementation; UAT and Remediation, Go-Live; Post Go-Live Support) | MS: cloned into the Timeline section in place of the OLE object before the Appendix is stripped (Track 1 correction 7). `add_timeline_row(...)`; `--keep-appendix-timeline-table` keeps the sample rows (see `--help`) |
| Fee table | MS 170 | `TableGrid` | 13 x 4 (grid 1838 / 3617 / 2659 / 1775, width 9889) | "Milestone", "Milestone Description", "Milestone Payment ($)", "Expected Completion" | rows 1 to 6 "Milestone 1" to "Milestone 5", "Milestone X" (`0.00`, row 1 "Week 1"); row 7 blank; row 8 "Sub-Total*" `=SUM(ABOVE) \# "0.00"`; row 9 "Client Discount / Where applicable, else delete / -0.00"; row 10 "License Cost / e.g., Sharegate / .00"; row 11 blank; row 12 "Total* / Total Project Estimate: / 0.00" `=SUM(ABOVE) \# "0.00"` with bookmark **`TE`** | `reset_fee_rows_ms`, `add_fee_row_ms(...)`, `set_fee_cell`, `set_field_result` (writes cached results, keeps fields and the `TE` bookmark) |
| Rate card | TM 155 | `TableGrid` | 30 x 7 (grid 1161 / 3772 / 1112 / 943 / 902 / 1608 / 1571, width 11069 twips = 7.69in, wider than the 6.5in text width as shipped) | "Location", "Netwoven Roles", "No of Resources", "Weekly Hours", "No of Weeks", "Rate (Per Hour)", "Total Estimate"; row 1 sub-header "Preferential ($)", "Costing ($)" | rows 2 to 9 onshore roles (Consultant*, Senior Consultant*, Principal Consultant*, Technical Project Manager / Project Manager*, Technical Architect / Sr. Technical Architect*, Engagement Lead*, Engagement Manager*, Practice Director*), each `=PRODUCT(LEFT) \# "0.00"`; row 10 "Onshore Estimate*" `=SUM(ABOVE)` bookmark **`Onshore_Total`**; row 11 blank; rows 12 to 20 offshore roles (Junior Developer/Admin*, Engineer / Admin / QA *, Senior Engineer / Sr. BA / Sr. QA…, Principal Engineer…, Sr. Principal Engineer…, Associate Technical Architect…, Technical Architect / Technical Project Manager…, Sr. Technical Architect…, Technical Director / Sr. Technical Director…), each `=PRODUCT(LEFT)`; row 21 "Offshore Estimate*" `=SUM(ABOVE)` bookmark **`Offshore_Total`**; row 22 blank; row 23 "Onshore Total*" `REF Onshore_Total`; row 24 "Offshore Total*" `REF Offshore_Total`; row 25 blank (merged); row 26 "Sub Total*" `=SUM(ABOVE)`; row 27 "Client Discount"; row 28 "License Cost"; row 29 "Total Project Estimate*" `=SUM(ABOVE)` bookmark **`Total_Estimate`** | `set_rate_row_tm(role, resources, hours, weeks, rate)`, `drop_unused_rate_rows_tm()`, `recompute_tm_totals()` (writes cached results for every `=PRODUCT`, `=SUM` and `REF` field and keeps all three bookmarks) |

`REF` fields: **the MS template has none** (Track 1 correction 6). Its
Executive Summary total is plain text (`*$ 0.00`) and its "Total Project
Estimate (right click to refresh…)" line is a highlighted plain paragraph;
the hyperlink "Price Estimate" points at the bookmark `_Price_Estimate` on
the Project Cost Estimate heading. The TM template has `REF Total_Estimate`
twice (Executive Summary and the "Total Project Estimate*: $0.00" line under
the rate card) and `REF _Ref145000101` (the Project Cost Estimate heading)
in the Executive Summary. `recompute_tm_totals` refreshes those caches; the
validator checks that every `REF` target bookmark still exists.

Word recalculates `=SUM(ABOVE)` and `=PRODUCT(LEFT)` on field update. How the
blank spacer rows interact with `SUM(ABOVE)` in Word is to be confirmed in
Word by the user; the helpers write the arithmetic result into the cache so
the document is correct even before an update.

---

## 7. Checkbox recipe

The 25 product checkboxes are `w:sdt` elements whose `w:sdtPr` carries a
`w14:checkbox` and whose `w:sdtContent` **wraps the whole table cell
(`w:tc`)**, not a run. python-docx's `row.cells` does not see cells inside
an SDT, so all SOW table code iterates `tr.iter(qn('w:tc'))` with lxml
(Track 1 correction 2).

Shipped values, identical on all 25:

```xml
<w:sdtPr>
  <w:id w:val="…"/>
  <w14:checkbox>
    <w14:checked w14:val="0"/>
    <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
    <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
  </w14:checkbox>
</w:sdtPr>
<w:sdtContent>
  <w:tc> … <w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:cs="Segoe UI" w:hint="eastAsia"/></w:rPr><w:t>☐</w:t></w:r> … </w:tc>
</w:sdtContent>
```

To tick: set `w14:checked/@w14:val` to `1` **and** replace the `w:t` glyph
`☐` (U+2610) with `☒` (U+2612). Word shows the glyph, not the flag, so both
must agree. `tick_product(doc, "SharePoint Online")` matches the product
name in column 1 (exact, then case-insensitive prefix) and does both.

```python
from docx.oxml.ns import qn
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

def tick(sdt):
    sdt.find(f".//{{{W14}}}checked").set(f"{{{W14}}}val", "1")
    for t in sdt.iter(qn("w:t")):
        if t.text == "☐":
            t.text = "☒"
```

Shipped product rows, in order: Azure Active Directory Premium 2; Azure
Active Directory Premium Conditional Access; Exchange Online; Insider Risk
Manager (IRM); Intune; Managed Security Services; Microsoft Defender
Endpoint (MDE); Microsoft Defender for Cloud Apps; Microsoft Defender for
Identity (MDI); Microsoft Defender for Office (MDO); Microsoft Information
Protection; Outlook Mobile; SharePoint Online; Microsoft Teams; Teams Apps;
Teams for Frontline Workers; Teams Meetings; Teams Phone; Teams Rooms; Viva
Connections; Viva Engage; Viva Goals; Viva Insights; Viva Learning; Viva
Topics. (Two names carry a longer suffix in the XML than shown here, "Microsoft
Defender Endpoint (MDE)" and "Microsoft Defender for Identity (MDI)"; match
on the prefix.)

---

## 8. Milestone versus T&M

| Aspect | Milestone (MS) | Time and Materials (TM) |
|---|---|---|
| Choose when | Fixed fee, deliverable acceptance, milestone payments | Hourly rates, weekly hours, monthly invoicing, capacity engagement |
| Glossary | SOW, Change Order ("A supplement document…"), GDC ("…(Outside of USA)") | Same three with slightly different wording, plus "T&M" |
| Executive Summary total | Plain text `*$ 0.00` (no field) | `REF Total_Estimate` field |
| Scope of Work | `BodyText` milestone outline (Milestone 1 / Project Preparation / …, Milestone 2, Milestone X, "Post Go Live support for X weeks") | `<<Put bulleted list…>>` prompt plus "Scope 1", "Scope 2" |
| Deliverables table | `PSOGrid` 2x5 with a Milestone column; empty sample row; a list of sample deliverables per milestone follows the table | `PSOGrid` 2x4; one filled sample row |
| Assumptions H2s | Overall Proposal, Working Environment, Deliverables (4 bullets incl. deemed acceptance after 5 days), Project Scope (Infrastructure & Support, Use Cases & Prototypes, Miscellaneous, Post Go-Live Support, Overall Timeline), CPOR, Govern 365 (6 H2) | Overall Proposal, Working Environment, Deliverables (2 bullets), CPOR, Govern 365 (5 H2; TM headings are lower-case "Claiming Partner of record (cpor)", "Govern 365 promotional offer", title-cased in the base) |
| "Pause / stop" clause | "the current milestone is immediately due for payment in full" | "the current invoice is immediately due for payment in full" |
| Staffing note | "Netwoven US and Offshore resources" | "Netwoven Offshore resources" |
| Timeline table | Lives in the Appendix; cloned into Timeline by the builder | Already in the Timeline section |
| Cost section | Fee table 13x4, bookmark `TE`, travel paragraph with `X%` cap, deemed-approval paragraph ("The above expected completion dates…") | Rate card 30x7, bookmarks `Onshore_Total`, `Offshore_Total`, `Total_Estimate`, `REF` fields; no travel paragraph |
| Payment Schedule | "Netwoven will Invoice <<Client>> based on terms and conditions of the MSA, PSA or SOW…" | "All Invoices will be submitted on last Friday of every month…" |
| Copy drifts fixed in base | "provide with a resource" → "provide a resource" (Staffing) | "This effort for endeavor" → "This effort for this endeavor" (Executive Summary; the sentence is stripped anyway); heading title case; `TOCHeading numId=0` |
| Section numbering as shipped | 1 to 9 | 2 to 10 (fixed in base) |
| Company SDTs | 46 | 35 |
| Highlighted runs | 31 | 22 |
| Rate-card hyperlink (`rId28`) | "…/Rate Cards/2023/Internal Use…" | "…/Rate Cards/2023/Client-Partner Use…" |

Mixed signals (fixed price for some phases, hourly for others) mean one
question to the user before choosing.

---

## 9. Skeleton: keep, optional, strip

Heading text is exact; numbers are the shipped MS numbering (TM as shipped is
one higher; the base fixes it).

| # | Section (Heading1 / Heading2 / Heading3) | Fate |
|---|---|---|
| | Cover text box, Statement of Confidentiality (three paragraphs, live year), page break | Keep |
| | "Contents" TOC SDT (`TOC \o "1-3" \h \z \u`) | Keep; `reset_toc("Contents")` |
| 1 | **Glossary** (three or four `ListParagraph` definitions) | Keep verbatim; add terms only if the SOW uses them; delete the `<<Insert Glossary Terms…>>` line |
| 2 | **Executive Summary** | Fill-in sentence replaced by a real two-to-four-sentence summary in `BodyText`; keep the "This Statement of Work (“SOW”) for services is entered and effective as of…" sentence with its date and Company controls |
| 3 | **Microsoft Products Utilized in Project** ("Solution Stack" M365 Workloads products table; "Biz Apps" table) | Keep; tick products; fill Biz Apps or drop unused rows |
| 4 | **Scope of Request and Deliverables** › Scope of Work · Deliverables · Out of Scope | Fill all three; Out of Scope is empty in the template and must be written |
| 5 | **Assumptions** › Overall Proposal · Working Environment · Deliverables · (MS) Project Scope · Establish M365 Claiming Partner of Record (CPOR) · Govern 365 Promotional Offer | Keep verbatim. Individual H2s can be removed through `--json-spec assumptions_remove` when the user says so (CPOR and Govern 365 are the usual candidates) |
| 6 | **Project Operations** › Netwoven Roles and Responsibilities (table) · `<<Client>>` Roles and Responsibilities (nine `Heading3`: Project Sponsors (Business & IT), Project Decision Maker, Project Coordinator / Manager, Business Users, IT Infrastructure Team, Intranet Support Team, Identity Team, Governance Committee (if applicable), IT Communications Team) · Change Requests · Status Reporting | Keep; trim roles rows and client-role H3s to the engagement; remove the highlighted note above the roles table |
| 7 | **Budget and Timeline** › Timeline · Project Cost Estimate · Payment Schedule · Staffing | Keep; OLE Gantt removed and replaced by the Phase table; fee or rate rows filled; instruction lines removed |
| 8 | **Signatures** (effective-date line, "For <<Client>>" / "For Netwoven, Inc.:" block, "This offer expires 30 days from the publish date of this document.") | Keep verbatim |
| 9 | **Appendix** ("THIS SECTION HAS BEEN CREATED AS A GUIDELINE…", roughly 150 paragraphs of internal scope-assumption libraries naming former clients, a SmartArt `Diagram 19`, and in MS the Phase table) | **Strip unconditionally**, from the `Heading1` "Appendix" paragraph to the final `sectPr`, dropping the `diagrams/*` relationships |

Also stripped by the builder: the embedded Excel Gantt (`w:object`,
`embeddings/Microsoft_Excel_Worksheet.xlsx`, EMF preview `image8.emf`, the
`_MON_…` bookmark), every highlighted instruction, the rate-card hyperlink,
the "Please Click the Executive Summary link" line, and sample content. The
post-save assertion checks that `word/embeddings/`, `word/diagrams/` and
`image8.emf` are absent from the output zip.

---

## 10. Verbatim legal and boilerplate paragraphs

These are quoted exactly from the template into every SOW; the AI never
rewrites them (house punctuation rules do not apply inside them). Identify
each by its opening words. Bracketed notes mark MS/TM differences.

Glossary
- "Statement Of Work (SOW)" definition ("A document that defines all the work management aspects of the project plan…")
- "Change Order (CO)" definition (MS "A supplement document of a Statement of Work…"; TM "A document in lieu of a Statement of Work…")
- "Global Delivery Center (GDC)" definition
- TM only: "T&M" definition ("A term used to denote a “Time and Materials” project…")

Executive Summary
- "This Statement of Work (“SOW”) for services is entered and effective as of MM/DD/YYYY ("Effective Date") by and between <<Client>>."

Scope intro
- "Below is a summary of the expected scope of services that the team will be responsible for providing to <<Client>> for the <<Project Name>> project:"
- "The following deliverables will be provided as part of the current SOW:"

Assumptions
- "The following assumptions and dependencies are based upon our approach and overall project scope."
- Overall Proposal: "<<Client>> will be responsible for the prioritization decisions…"; "All <<Client>> sponsors, stakeholders and technical resources…"; "<<Client>> will provide a project coordinator…"; "The project will be staffed using Netwoven [US and] Offshore resources…"; "If there are any changes at <<Client>> that impacts the project… the current [milestone | invoice] is immediately due for payment in full."
- Working Environment: "Project work will be performed at Netwoven’s [USA & ]GDC locations…"; "Remote access to <<Client>>’ internal network…"; "Appropriate access to M365 [needs to be | is] provided…"; "<<Client>> will provide tools for team collaboration…"
- Deliverables: "Appropriate information is [required to be | ] available from the client…"; "Two reviews of deliverables are included in the scope…"; MS only: "<<Client>> will provide feedback on the deliverables in a timely manner. If feedback is not received within 5 days…" (deemed acceptance, keep); MS only: "Where completion of deliverables relies on <<Client>>…"
- MS Project Scope: "<<Client>> IT will provision all requested servers…"; "Netwoven will work within the confines of <<Client>> security…"; "2 weeks of post go-live support are included in the project…"; "The project will be completed within the provided timeline in this Statement of Work…"
- CPOR: "CPOR enables Netwoven to help <<Client>> optimize the use of Microsoft Online Services…"
- Govern 365: "Netwoven’s product Govern 365 has proved tremendously valuable…" (keeps the `www.govern365.com` hyperlink, `rId25`)

Project Operations
- "The following section defines project operations including roles and responsibilities, change request process, and project reporting."
- Each client-role `Heading3` and its "Responsibilities include:" bullets
- Change Requests: "During this engagement, Netwoven or <<Client>> may determine a change in scope…"; "When a change request is submitted, both parties will review the request…"
- Status Reporting: "Status and health related metrics will be reviewed with <<Client>> on a weekly basis…"

Budget and Timeline
- "The following high-level timeline has been put together for the project. A Detailed [Work Breakdown Structure (WBS) | project plan] will be created during the project initiation phase."
- "The timeline is based on the current understanding of the requirements…"
- MS: "Based on the project scope and deliverables defined earlier, the following costs are estimated…"; TM: "Under this SOW, Netwoven will provide the following resources for the duration of the contract…"
- MS: "Travel and living expenses will be approved in advance and in writing by the <<Client>> Project Manager or Sponsor… shall not exceed X% of total project cost, above." (the `X%` is filled from `--expenses-cap`)
- MS: "The above expected completion dates are estimates… If <<Client>> doesn’t respond within 5 business days… it will be deemed approved as completed and ready for invoicing…" (deemed acceptance, keep)
- Payment Schedule: MS "Netwoven will Invoice <<Client>> based on terms and conditions of the MSA, PSA or SOW…"; TM "All Invoices will be submitted on last Friday of every month…"
- "Note: All payments are due within the agreed-upon payment terms… a late fee of 2% of the outstanding balance…"
- Staffing: "Any named resource discussed or in this Statement of Work is subject to availability…"

Signatures
- "This Statement of Work has been reviewed and approved as per the signatures below."
- "The effective date for this Statement of Work [(“SOW”)] is MM/DD/YYYY"
- "For <<Client>>" / "For Netwoven, Inc.:" block with Name, Signature, Date lines
- "This offer expires 30 days from the publish date of this document."

Confidentiality (cover pages)
- "This document contains information that is proprietary and confidential to Netwoven, Inc. and <<Client>> which shall not be disclosed…"
- "Any other company and product names mentioned are used for identification purposes only and may be trademarks of their respective owners."
- "©2001 - [DATE yyyy] Netwoven, Inc. All rights reserved…"

---

## 11. Quirks

1. **Use `Heading1/2/3` only.** `NW Heading` styles are gone from the base;
   `Heading4` numbers from a different list (`numId 1`).
2. **TM section numbering is off by one as shipped** (TOC title takes
   number 1). Fixed in the base with `TOCHeading numPr numId=0`.
3. **`Body Text` is Verdana**, not Segoe UI, and indented 0.5in. The
   Executive Summary and Signatures lines use it. Left as shipped.
4. **`Netwoven Table 1/2` header font was `Sitka Banner`**, a substitution
   defect; fixed to Segoe UI Semibold (approved). `Netwoven Table 21`
   (`@Yu Gothic UI Semilight`) is a corrupted duplicate and is removed.
5. **The header and footer print on the cover** (no `titlePg`). This is the
   template's design; do not add a first-page header.
6. **Header spacing is 115 literal spaces** between the title control and the
   logo. A long title wraps; keep titles under about 60 characters (to be
   confirmed by render).
7. **The rate card is wider than the text block** (7.69in vs 6.5in) and
   extends into the right margin as shipped. Left as designed; to be
   confirmed by render.
8. **`MM/DD/YYYY` looks like plain text but is a date control.** The footer
   control formats as `M/d/yyyy`; the body controls carry no `dateFormat`
   and show whatever cached text the builder writes (the builder writes the
   same `M/d/yyyy` form).
9. **The confidentiality year is a `DATE` field** and updates itself; do not
   hard-code a year.
10. **Checkbox SDTs wrap table cells**; see §7 for the python-docx pitfall.
11. **Appendix names former clients** (Intel, Ross Stores, Cypress) and
    embeds a 1.08in SmartArt; the validator's forbidden-string list catches
    any survivor.
12. **The rate-card hyperlink leaks a SharePoint tenant path**; removed with
    its paragraph and relationship.
13. **Sensitivity label kept.** `docProps/custom.xml` holds the
    `MSIP_Label_d3e71191-c083-4948-94ca-f99c3ca5a353_*` properties (Name
    "General Business", Method "Standard", SetDate 2020-06-25) and
    `docMetadata/LabelInfo.xml` holds the matching label. The base removes
    only the SharePoint properties around them (Sales Pursuit Doc Type,
    AuthorIds_*, ContentTypeId, taxonomy GUID properties, Proposal Status0,
    `SharedWithUsers` with three staff names, and in TM
    `GrammarlyDocumentId`). The nested `LabelInfo` inside the embedded xlsx
    goes away with the xlsx itself; the builder never edits a label.
14. **Metadata leaks scrubbed**: `dc:creator` / `lastModifiedBy` named
    employee, `dc:subject` "OneDrive For Business (OD4B) Assessment And
    Planning", `lastPrinted` 2018-06-02, MS `word/intelligence2.xml` (empty
    Editor cache), `documentProtection`.
15. **Six `customXml` items**; only `item1` (`CoverPageProperties`,
    `{55AF091B-…}`) is referenced. The dead ones have different numbers and
    GUIDs in MS and TM; identify by `ds:itemID`.
16. **`BodyText` and MS `TOC2` carry `w:cs="Sendnya"`**, a font that does not
    exist; replaced by Segoe UI in the base.

---

## 12. Procedure for an AI (the `sow` skill)

1. Decide the variant from the source (§8 first row). Mixed or unclear:
   ask one question.
2. Run the analysis with the purpose "the reader signs". The pyramid is
   short: what Netwoven will deliver, for how much, by when, under which
   assumptions.
3. Build the base with the script (§13). Pass `--client-logo` if a logo
   file exists; otherwise the logo box is removed.
4. Fill the sections in template order using `Heading1/2/3` only. Never
   invent rates, hours, dates or discounts; every number comes from the
   source or is flagged.
5. Keep every paragraph in §10 verbatim. Never carry the rate-card link or
   the Appendix. Keep the deemed-acceptance clauses (MS).
6. Ensure nothing highlighted, no `<<…>>`, no `MM/DD/YYYY`, no "Please
   Insert Client Logo" survives (validator hard failures).
7. Validate (`--type sow --client "Client, Inc."`). Deliver the flag list in
   chat.

---

## 13. Programmatic recipe

```
python scripts/new_sow_docx.py --type milestone assets/bases/NW_SOW_Milestone_Base_2026.docx out.docx \
  --title "Intranet Modernization" --client "Contoso, Inc." --date 2026-09-15 \
  [--client-logo contoso.png] [--keep-products] [--keep-appendix-timeline-table] \
  [--json-spec spec.json] [--expenses-cap 10]

python scripts/new_sow_docx.py --type tm assets/bases/NW_SOW_TM_Base_2026.docx out.docx \
  --title "Managed Services Q4" --client "Contoso, Inc." --date 2026-09-15 [--json-spec spec.json]
```

The builder asserts the base matches `--type`, rejects `--internal`, strips
in content-keyed order (Appendix with the MS timeline table cloned first,
OLE timeline, highlighted and unhighlighted instructions, sample content),
runs `reset_toc("Contents")`, `set_bound_fields`, logo handling,
`prune_numbering`, `scrub_metadata`, sets `updateFields` and
`app.xml TitlesOfParts`.

`--json-spec` schema (all keys optional): `milestones` (list of
`{name, description, amount, completion}`), `rates` (list of
`{role, location, resources, weekly_hours, weeks, rate}`), `products` (names
to tick), `biz_apps` (names), `deliverables` (rows, MS 5 columns / TM 4 --
see the note in `new_sow_docx.py`'s own docstring), `timeline` (rows),
`expenses_cap` (number), `assumptions_remove` (list of Assumptions H2
titles).

Filling by hand with the helpers:

```python
import docx
from nw_docx_helpers import (set_bound_fields, reset_toc, find_table, add_fee_row_ms,
                             recompute_ms_totals, set_rate_row_tm, recompute_tm_totals,
                             tick_product, set_biz_apps, add_deliverable_row, add_timeline_row,
                             set_client_logo, remove_client_logo_box, strip_highlighted_instructions)

d = docx.Document("out.docx")

tick_product(d, "SharePoint Online")
tick_product(d, "Microsoft Teams")
set_biz_apps(d, ["D365 Customer Engagement"])            # drops the unused "Product N" rows

# add_deliverable_row/add_timeline_row/add_fee_row_ms/set_rate_row_tm all take
# the TABLE element (from find_table), not the Document -- resolve it first.
deliverables_tbl = find_table(d, header_startswith=["Scope Section"])
# MS deliverables row is 5 columns (Scope Section, Milestone, Deliverable,
# Description, Format); TM is 4 (no Milestone column) -- match the base's
# real shape or values silently shift into the wrong cells.
add_deliverable_row(deliverables_tbl,
                    ["Design", "Milestone 1", "Information architecture", "Site map and navigation", "Document"])

timeline_tbl = find_table(d, header_startswith=["Phase"])
add_timeline_row(timeline_tbl, "Discovery", ["Workshops and requirements"], "Weeks 1 to 3")

# Milestone fee table: rows above Sub-Total; cached results written, =SUM(ABOVE) fields and bookmark TE kept
fee_tbl = find_table(d, header_startswith=["Milestone"])
add_fee_row_ms(fee_tbl, "Milestone 1", "Discovery and design", 24000.00, "Week 3")
add_fee_row_ms(fee_tbl, "Milestone 2", "Build and test", 61000.00, "Week 12")
recompute_ms_totals(fee_tbl, d.element.body)

# T&M rate card (TM base only): unused role rows dropped, =PRODUCT(LEFT)/=SUM(ABOVE)/REF caches refreshed
# rate_tbl = find_table(d, header_startswith=["Location"])
# set_rate_row_tm(rate_tbl, "Senior Consultant", "onshore", resources=1, weekly_hours=40, weeks=12, rate=185.00)
# recompute_tm_totals(rate_tbl, d.element.body)

set_client_logo(d, "contoso.png")        # or remove_client_logo_box(d)

# Every in-memory (python-docx) edit above must happen before this save.
d.save("out.docx")

# Bound fields rewrite every <<Client>>, <<Project Name>>, MM/DD/YYYY cache
# and the three stores directly in the saved zip. Takes a FILE PATH, not the
# in-memory Document, and must run AFTER the save above -- reopening `d` and
# re-saving it afterward would silently discard this patch.
set_bound_fields("out.docx", title="Intranet Modernization", company="Contoso, Inc.", date="2026-09-15")
```

All table code uses lxml `tr.iter(qn('w:tc'))`, never `row.cells`, because
of the checkbox SDTs (§7). Only python-docx, python-pptx and lxml are
available at runtime (Python 3.9 syntax).

---

## 14. Changelog per variant (as built by `tools/build_bases.py`)

All eight gates passed for both variants (Stage B whitelist, Stage C
single-part diff, forbidden strings absent in `_v2` and base, LabelInfo
present and byte-identical, builder markers on the base). Rendering gates
were not run; the expected outcome (Stage A identical pages, Stage B page
count within two pages with an identical cover) is to be confirmed by the
user's render pass.

### Stage A, both variants (invisible when rendered)

1. `customXml`: 5 unreferenced items removed with `itemProps`, `.rels`,
   overrides and document rels. MS: `item2` (content type), `item3` (form
   templates), `item4` (bibliography), `item5` (library properties), `item6`
   (taxonomy sync). TM: `item2` (content type), `item3` (taxonomy sync),
   `item4` (form templates), `item5` (library properties), `item6`
   (bibliography). `item1` (`CoverPageProperties`) kept in both.
2. `docProps/custom.xml`: non-MSIP properties removed (MS 9: AuthorIds x3,
   ContentTypeId, Proposal Status0, Sales Pursuit Doc Type, SharedWithUsers,
   two taxonomy GUID properties; TM 10: the same plus GrammarlyDocumentId);
   all 7 `MSIP_Label_d3e71191-…` properties kept (pids renumbered from 2).
   `docMetadata/LabelInfo.xml` byte-identical to the original.
3. MS only: `word/intelligence2.xml` removed with its override and
   relationship (`rId39`).
4. `docProps/core.xml`: creator and lastModifiedBy → "Netwoven",
   `lastPrinted` removed, `dc:subject` cleared (4 fields). `dc:title` keeps
   `<<Project Name>>` for the builder to overwrite.
5. Unused styles removed, MS 28 / TM 29: `KCITableNormal`,
   `NetwovenTable21`, `eop`, `normaltextrun`, `pagebreaktextspan`,
   `BalloonText`, `BalloonTextChar`, `CommentReference`, `CommentText`,
   `CommentTextChar`, `CommentSubject`, `CommentSubjectChar`, `Revision`,
   `Mention`, `UnresolvedMention`, `FollowedHyperlink`, `PlainText`,
   `PlainTextChar`, `Quote`, `QuoteChar`, `ListTable4-Accent1`,
   `PlainTable5`, `TOC4`, `TableofFigures`, `NWHeading1` to `NWHeading4`;
   TM additionally `ui-provider` (not present in MS). **Recorded deviation:
   `PlainTable3` was on the removal list but is the `basedOn` parent of
   `NetwovenTable3`, so the guard kept it in both variants.**
6. `w:rFonts/@w:cs="Sendnya"` → `Segoe UI` (MS 3, TM 4 occurrences).
7. Numbering: MS removed 9 `w:num` (`4, 5, 16, 17, 25, 27, 28, 57, 58`)
   and 9 orphan `w:abstractNum` (`2, 4, 19, 22, 29, 39, 40, 42, 48`); TM
   removed 5 `w:num` (`4, 5, 16, 17, 20`) and 5 orphan `w:abstractNum`
   (`3, 20, 27, 36, 37`). Heading and `2-Level Legal` lists untouched.
8. rsids: MS 3735 / TM 4412 attributes stripped from 6 text parts; 58
   `w:style/w:rsid` elements removed in each; `w:rsids` (MS 4157, TM 4010
   entries) removed from both settings parts.
9. fontTable: 12 entries pruned per variant, in `word/fontTable.xml` and
   its glossary mirror `word/glossary/fontTable.xml` (Yu Gothic Light,
   Sendnya, SimSun, @Yu Gothic UI Semilight, Yu Mincho in both; Aptos and
   Aptos Display in the glossary copy only).
10. `settings.xml`: `documentProtection` (`readOnly`, `enforcement=0`)
    removed.
11. Embedded `word/embeddings/Microsoft_Excel_Worksheet.xlsx`: nested
    `docProps/core.xml` creator and lastModifiedBy → "Netwoven"; the nested
    `LabelInfo.xml` byte-identical.

### Stage B, both variants (approved visible fixes)

Stage B vs Stage A changed exactly `word/document.xml`, `word/fontTable.xml`,
`word/glossary/fontTable.xml`, `word/settings.xml`, `word/styles.xml`.
**Recorded deviation: the Stage B whitelist was extended with
`word/glossary/fontTable.xml`**, because Word mirrors the font table into
the glossary part and the "zero `Sitka Banner` occurrences" gate could not
pass without pruning it too.

12. `Heading1`: `rFonts` ascii/hAnsi Segoe UI Semibold, cs Segoe UI,
    eastAsiaTheme majorEastAsia; `b` and `bCs` removed; `sz`/`szCs` 44; navy
    `#000080` bottom `pBdr` removed; MS `ind left=720` removed (TM had no
    indent); `keepNext`, `pageBreakBefore`, `numPr`, `spacing`, `outlineLvl`
    kept. `Heading1Char` mirrored.
13. `Heading2`: same font change, bold removed, `sz`/`szCs` 36, `#AEAAAA`
    top rule kept, small caps kept. `Heading2Char` mirrored.
14. `Heading3`: same font change, bold and `smallCaps` removed, colour
    `auto`, `sz`/`szCs` 28. `Heading3Char` mirrored.
15. `NetwovenTable1` and `NetwovenTable2` first-row `rFonts` `Sitka Banner`
    → `Segoe UI Semibold` (2), followed by pruning `Sitka Banner` from both
    font tables (2).
16. `settings.xml`: `<w:updateFields w:val="true"/>` inserted before
    `hdrShapeDefaults`.

### Stage B, MS only

17. Copy fix "provide with a resource with similar" → "provide a resource
    with similar" (Staffing, 1 run).

### Stage B, TM only

18. `TOCHeading` gets `<w:numPr><w:numId w:val="0"/></w:numPr>`, fixing the
    section numbering off-by-one.
19. Copy fix "This effort for endeavor is" → "This effort for this endeavor
    is" (Executive Summary, 1 run; the sentence is removed by the builder
    anyway).
20. Heading title case, each in the heading and its cached TOC entry (2 +
    2): "Establish M365 Claiming Partner of record (cpor)" → "Establish M365
    Claiming Partner of Record (CPOR)"; "Govern 365 promotional offer" →
    "Govern 365 Promotional Offer".

### Stage C, both variants

`[Content_Types].xml` main part flipped to `document.main+xml`; every other
part byte-identical to `_v2`; python-docx opens both bases and the builder
markers were found (`Heading1` "Appendix", the fee or rate table with its
fields and bookmarks, 25 checkboxes, the client logo box, the "Contents"
TOC, the "Netwoven Confidential" footer).

Parts diff: Stage A vs original changed 17 parts and removed 16 (MS) or 15
(TM); Stage B vs A changed 5; Stage C vs B changed 1. Hashes: MS source
`2d60b722…`, `_v2` `03885d06…`, base `e91703fe…`; TM source `105a2be2…`,
`_v2` `e32059d8…`, base `64d009bd…`.

### Not changed (by decision)

- Purview/MSIP sensitivity label: `docMetadata/LabelInfo.xml`, its
  relationship and content type, every `MSIP_Label_*` property in
  `docProps/custom.xml`, and the nested `LabelInfo.xml` inside the embedded
  Excel worksheet (the worksheet itself is removed later by the builder).
- Heading numbering formats (`numbering.xml` levels, `numId`s on
  `Heading1` to `Heading9`; `Heading4` stays on `numId 1`).
- `word/theme/theme1.xml` (`Heading1/2/3` and `TOC1` carry
  `themeColor=accent1`).
- The rate-card hyperlink `rId28` and all sample and instruction content
  (builder responsibility, not base cleanup).
- `docProps/app.xml` (`TitlesOfParts` and `Company` are rewritten by the
  builder).
- `Body Text` in Verdana, the 7.69in rate card width, the header on the
  cover, and `TOC2` colours.
