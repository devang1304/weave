---
title: Netwoven Word Document Template 2026 (AI Reference Guide)
covers: NW_document_template_2026.dotx, NW_document_template_2026_v2.dotx, NW_Document_Base_2026.docx
last_verified: 2026-09-02
---

# Netwoven Word document template 2026

This is the machine-readable specification for Netwoven's general document
template. It exists so an AI (or a person) can produce a correctly branded
Word deliverable without reverse-engineering the OOXML each time. Every value
was read from the unpacked template XML. Where a claim depends on a rendered
page that this build has not produced, it is marked "to be confirmed by
render". Where the previous guide (for the 2021 template) was wrong, the
correction is called out explicitly.

**Files this guide describes:**

| File | Purpose |
|---|---|
| `NW_document_template_2026.dotx` | Netwoven's original, unmodified. It is the 2021 template re-saved in Word 2026 (see §1). |
| `NW_document_template_2026_v2.dotx` | Cleaned template. Stage A removes dead metadata with no visible change; Stage B applies the approved configuration fixes listed in §11. |
| `NW_Document_Base_2026.docx` | `_v2` with one content-type string flipped so `python-docx` opens it. Content-identical otherwise. This is the file the scripts use (`assets/bases/NW_Document_Base_2026.docx`). |

Related references: `brand-tokens.md` (colours, fonts, palettes),
`visual-standards.md` §10 (figures, captions, table selection),
`writing-method.md` (prose), `sow-template-guide.md` (the SOW templates
are a different lineage; do not apply this guide to them).

---

## 1. The one thing to understand first

**The 2026 document template is the 2021 template re-saved, not a descendant
of the cleaned `_v2` of 2021.** Styles, theme, numbering, both sections,
header and footer XML, and all 20 media files are byte-identical to the 2021
original. Every 2021 cleanup was reverted (the SharePoint and bibliography
`customXml` parts, `docProps/custom.xml`, the `KCITableNormal` style, the
"Segue UI" typo, personal names in `core.xml`).

Three things did change in the re-save:

1. The confidentiality copyright now reads `©2001 - 2026` (was 2018). The
   year is literal text split across three runs (`©2001 - 20`, `2`, `6`),
   not a field. It will not update itself.
2. `dc:title` is empty and all three Title content controls carry
   `showingPlcHdr`. Combined with a glossary regression (item 3), the header
   renders the text `[Company]` and the cover title renders five spaces until
   Word re-resolves the binding.
3. A Purview sensitivity label was applied (`docMetadata/LabelInfo.xml` and
   the `MSIP_Label_8c8476dd-…` properties, label name "Public",
   `contentBits=0`, non-visual). **This label is kept, untouched, in `_v2`
   and in the base.** It is never removed.

**Correction to the old guide.** The theme (`word/theme/theme1.xml`) is the
stock Office scheme (`accent1=#4472C4`, Calibri Light / Calibri) and the old
guide called it "a dead decoy". It is not dead. `Heading1`, `Heading2`,
`Heading4`, `TOC1`, and therefore `NW Heading 1`, `NW Heading 2` and
`NW Heading 4`, colour their text with `w:themeColor="accent1"
w:themeShade="BF"`, which resolves to `#2F5496` only because `accent1` is
`#4472C4`. `TOC Heading` and the base `Heading1` take the theme major font
(Calibri Light); `No Spacing` and the `Netwoven Table 3` header row take the
theme minor font (Calibri). **Retheming this file would recolour every heading
and change three fonts. Never edit `theme1.xml`.**

---

## 2. Brand tokens used by this template

Full definitions are in `brand-tokens.md`. This table only says where each
token is wired in this file.

| Token | Hex | Where |
|---|---|---|
| Netwoven rule blue | `#145CA4` | `header2` bottom rule and `footer2` top rule, both 3pt (`w:sz="24"`); `Netwoven Table 1` and `2` header fills; `Netwoven Table 3` header text; cover layout-table borders; `NW Heading 1` `w:u w:color` (no underline style, so nothing is drawn) |
| Heading blue | `#2F5496` | `NW Heading 1`, `NW Heading 2`, `TOC1`, `TOC Heading`, base `Heading1/2/4` (theme-derived, see §1) |
| Heading 4 grey | `#3B3838` | `NW Heading 4` (`background2` shade `40`) |
| Heading 2 rule | `#AEAAAA` | `NW Heading 2` top border, 0.5pt |
| Caption grey | `#44546A` | `Caption`, `No Spacing` (`text2`) |
| Grid | `#262626` | `Netwoven Table 1` borders (`text1` tint `D9`) |
| Light grid | `#BFBFBF` | `Plain Table 1` borders inherited by `Netwoven Table 2` |
| Header underline | `#7F7F7F` | `Netwoven Table 3` header bottom border |
| Band | `#F2F2F2` | `Netwoven Table 2` and `3` alternate rows; cover pillar labels |
| Pillar band | `#404040` | Five service-pillar cells on the cover |
| Cover title rule | `#C00000` | 3pt cell border under the cover title |
| Confidentiality border | `#FF0000` | Dotted 0.5pt box around the three legal paragraphs |
| Placeholder | `#808080` | `Placeholder Text` and `Subtle Emphasis` (the "Statement of Confidentiality" heading uses `Subtle Emphasis` with italic switched off) |
| Logo pixels | `#00549F`, `#B30738` | `image1.png` only |

Fonts: Segoe UI (body), Segoe UI Semibold (headings, table 1/2 headers, the
confidentiality and revision-history headings), Calibri Light (`TOC Heading`
via theme), Calibri (`No Spacing`, `Netwoven Table 3` header via theme). The
`fontTable` also lists Symbol, Courier New and Wingdings (bullet glyphs) and
Yu Gothic Light / Yu Mincho (East Asian theme fonts, unused). Segoe UI is not
embedded; see the portability caveat in `brand-tokens.md` §3.

---

## 3. Style catalog (fully resolved)

`docDefaults`: theme minor font (Calibri), 11pt, `spacing after=160`
(8pt), `line=259` (1.08). `Normal` overrides only the Latin font to Segoe UI.
Every row below is the as-rendered result after inheritance.

### 3.1 Paragraph styles to use

| Style ID | Display name | Font | Size | Colour | Case / weight | Paragraph |
|---|---|---|---|---|---|---|
| `Normal` | Normal | Segoe UI | 11pt | auto | regular | after 8pt, line 1.08 |
| `NWHeading1` | NW Heading 1 | Segoe UI Semibold | 22pt | `#2F5496` | small caps | page break before, keep with next, keep lines, before 12pt, after 0, outline 0, numbered `%1` (numId 2 level 0, hanging 0.3in) |
| `NWHeading2` | NW Heading 2 | Segoe UI Semibold | 18pt | `#2F5496` | small caps | 0.5pt `#AEAAAA` top rule, before 2pt, numbered `%1.%2` (hanging 0.4in) |
| `NWHeading3` | NW Heading 3 | Segoe UI Semibold | 14pt | auto (black) | normal case | numbered `%1.%2.%3` (hanging 0.5in) |
| `NWHeading4` | NW Heading 4 | Segoe UI Semibold | 12pt | `#3B3838` | italic (inherited from `Heading4`) | numbered `%1.%2.%3.%4` (hanging 0.6in) |
| `ListParagraph` | List Paragraph | Segoe UI | 11pt | auto | regular | left indent 0.5in, contextual spacing. Bullets come from `numId 7` (Symbol bullet, then Courier New "o", then Wingdings square) |
| `Caption` | caption | Segoe UI | 9pt | `#44546A` | italic | after 10pt, single line |
| `NoSpacing` | No Spacing | **Calibri** (theme minor; no `basedOn`, no font override) | 10pt | `#44546A` | regular | no spacing, single. **Correction:** the old guide said Segoe UI |
| `TOCHeading` | TOC Heading | **Calibri Light** (theme major, via `Heading1`) | 16pt | `#2F5496` | regular | outline level 9. The template's own TOC title paragraph carries a direct `numPr numId=0` so it is not numbered; `reset_toc` deep-copies that `pPr` |
| `TOC1` | toc 1 | Segoe UI | 12pt | `#2F5496` | bold | tabs 0.31in and right dot-leader at 6.49in, after 5pt |
| `TOC2` / `TOC3` / `TOC4` | toc 2 to 4 | Segoe UI | 11pt | auto | regular | indent 0.15in / 0.31in / 0.46in, after 5pt. `TOC4` has no references and is removed in the base |
| `TableofFigures` | table of figures | Segoe UI | 11pt | auto | regular | after 0. Used by both the List of Figures and the List of Tables |
| `Header` / `Footer` | header / footer | Segoe UI | 11pt (runs set 10pt) | auto | regular | centre tab 3.25in, right tab 6.5in, no spacing |

Never apply the base `Heading1` to `Heading4` directly; they exist only as
parents (Calibri Light 16/13/12/11pt, `#2F5496` or `#1F3763`). The validator
flags direct `Heading1` use in a document deliverable.

Character styles: `PlaceholderText` (`#808080`), `SubtleEmphasis` (italic
`#808080`), `Emphasis` (italic), `Hyperlink` (`#0563C1` underline),
`Heading1Char` to `Heading9Char` (linked, unused), `UnresolvedMention`
(unused, removed in the base).

### 3.2 Table styles

| Style ID | Display name | Header row | Body | Borders | Notes |
|---|---|---|---|---|---|
| `NetwovenTable1` | Netwoven Table 1 | fill `#145CA4`, Segoe UI Semibold 12pt white (`b=0`) | plain | `#262626` 0.5pt full grid | Body font Segoe UI, paragraphs single, no spacing |
| `NetwovenTable2` | Netwoven Table 2 | fill `#145CA4`, Segoe UI Semibold 12pt white | odd rows white, even rows `#F2F2F2` | `#BFBFBF` 0.5pt grid (from `PlainTable1`) | Also defines last-row bold with double top border, first/last column bold, vertical band `#F2F2F2` |
| `NetwovenTable3` | Netwoven Table 3 | no fill; **Calibri** (theme `minorHAnsi`) bold ALL CAPS 12pt `#145CA4`; `#7F7F7F` bottom border | `#F2F2F2` banding | none | **Correction:** the old guide implied Segoe UI. First column is bold caps with `#7F7F7F` right border |
| `PlainTable1` / `PlainTable3` | Plain Table 1 / 3 | | | | Parents only; never apply directly |
| `TableGrid`, `PlainTable5`, `KCITableNormal` | | | | | Zero uses; removed in the base |

Selection rule (`visual-standards.md` §10.3): Table 1 default, Table 2 for
eight or more rows, Table 3 for a small quiet inline table.

### 3.3 Numbering

Seven `w:num` definitions. `numId 2` (abstract 5) is the nine-level outline
linked to `Heading1` to `Heading9` (`%1`, `%1.%2`, `%1.%2.%3` with no
trailing dot). Applying an `NW Heading` style numbers the paragraph
automatically. `numId 7` (abstract 0) is the bullet list used by
`ListParagraph`. `numId 1, 3, 4, 5, 6` are numbered and lettered lists used
only inside the disposable instructional section; the builder's
`prune_numbering` drops any `w:num` (then any orphan `w:abstractNum`) left
unreferenced after the strip.

---

## 4. Page, section, header and footer mechanics

- Page Letter (`12240 x 15840` twips), margins 1in, header and footer inset
  0.4in (`576`), line pitch 360.
- **Two sections.** Section 1 is the cover only: `titlePg` on, page numbers
  lower roman starting at 0, and six header/footer references (`even`
  `header1`/`footer1`, `default` `header2`/`footer2`, `first`
  `header3`/`footer3`). The section break sits in the paragraph at body
  index 34, right after the List of Tables. Section 2 (the final `sectPr`)
  carries **no** header/footer references and inherits section 1's; page
  numbers restart at 1.
- `header2` (default): `Header` style paragraph with a 3pt `#145CA4` bottom
  border; the wordmark anchored right (`1076325 x 154305` EMU, 1.18in by
  0.17in); a Title content control (10pt) bound to `dc:title`.
- `footer2` (default): 3pt `#145CA4` top border; Company content control
  (10pt) / centre `ptab` / `PAGE` field / right `ptab` / Publish Date
  content control (date picker, `storeMappedDataAs dateTime`, no
  `dateFormat` in the original; the base adds `M/d/yyyy`).
- `header3`/`footer3` (first page) are blank paragraphs: `titlePg` uses them
  to suppress the header on the cover. Keep them.
- `header1`/`footer1` (even page) are blank and vestigial; `settings.xml` has
  no `evenAndOddHeaders`, so they never show. They stay in the base because
  section 1 references them and removing them buys nothing.
- `settings.xml` in the original has no `updateFields`; the base adds
  `<w:updateFields w:val="true"/>` (before `hdrShapeDefaults`) so the TOC
  and captions refresh on open.

---

## 5. Content controls and the glossary regression

Three values propagate through data-bound content controls (`w:sdt` with
`w:dataBinding`):

| Field | Store | storeItemID | Occurrences (document / header / footer) | Cached text in the original |
|---|---|---|---|---|
| Title | `docProps/core.xml` `dc:title` | `{6C3C8BC8-F283-45AE-878A-BAB7291924A1}` | 2 (cover, Choice and Fallback copies) / 1 (`header2`) / 0 | five spaces on the cover, `[Company]` in the header, all with `showingPlcHdr` |
| Company | `docProps/app.xml` `Company` | `{6668398D-A668-4E3E-A5EB-62B293D839F1}` | 5 (cover x2, confidentiality paragraphs x3) / 0 / 1 | `Company Name Here` |
| Publish Date | `customXml/item1.xml` (`CoverPageProperties/PublishDate`) | `{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}` | 2 (cover, `w:date` with stale `fullDate=2017-12-28T00:00:00Z`) / 0 / 1 | `[Publish Date]` with `showingPlcHdr` |

The placeholder prompts come from `word/glossary/document.xml`. The
re-save collapsed the glossary to **three placeholder docParts that all read
`[Company]`**, one of which (`B25901AB…`) is shared by the cover Title SDT
and the header Title SDT, another (`3E84454D…`) by the body Company SDT and
the footer Company SDT, and the third (`6301F91A…`) by the cover and footer
Publish Date SDTs. Because a docPart is shared, "fixing the text" of one
would change the other. The base therefore **adds two new docParts**
(`[Title]`, `[Publish Date]`) and re-points `header2`'s Title SDT and
`footer2`'s Publish Date SDT to them (Track 1 correction 3), leaving the
original three untouched.

How the builder sets the fields (`nw_docx_helpers.set_bound_fields`): write
`dc:title`, `app.xml Company`, and `item1.xml PublishDate` (ISO date), then
rewrite the cached `w:t` inside **every** bound SDT in document, header and
footer, remove `showingPlcHdr`, and set `w:date/@w:fullDate` on the date
controls. Editing only the store, or only one visible run, leaves LibreOffice
and unrefreshed Word showing stale text.

The cover itself is a `Cover Pages` gallery SDT (`docPartUnique`) in the
first body paragraph (`NoSpacing`), containing two anchored text boxes:
`Text Box 1` (`6858000 x 9144000` EMU, the full 7.5in by 10in cover canvas)
and `Text Box 4` (`2130950 x 500463` EMU, the version and date box). Inside
`Text Box 1` sit two five-row layout tables (one in `mc:Choice`, one in the
VML `mc:Fallback`) with `#404040` and `#F2F2F2` fills and `#145CA4` /
`#C00000` 3pt borders, the hero photo, five pillar icons, and the MISA badge.
Helpers that touch the cover must edit both the Choice and the Fallback copy.

---

## 6. Quirks (read before assuming anything)

1. **Headings are not Netwoven rule blue.** `NW Heading 1/2` are `#2F5496`
   (theme-derived); only rules and table headers use `#145CA4`. Reproduce
   exactly.
2. **NW Heading 3 is black** (`w:color auto`), normal case.
3. **NW Heading 4 is italic**, inherited from `Heading4`, not visible in its
   own XML.
4. **NW Heading 1's underline draws nothing**: `w:u` has a colour but no
   style.
5. **`TOC Heading` is Calibri Light**, not Segoe UI Semibold, because it
   inherits `Heading1`'s theme major font.
6. **`No Spacing` is Calibri 10pt grey**, not Segoe UI (correction).
7. **`Netwoven Table 3`'s header row is Calibri** via `minorHAnsi`
   (correction).
8. **The theme is load-bearing** for heading colour and three fonts. Never
   retheme (correction; the old guide suggested retheming was safe).
9. **The confidentiality year is literal text**, `©2001 - 2026`, split over
   three runs. The builder does not alter it; flag to the user if a different
   year is required.
10. **Title SDT placeholder regression.** `dc:title` empty plus
    `showingPlcHdr` plus shared `[Company]` docParts means the header shows
    `[Company]` in the shipped template. Fixed by the base (§5, §11).
11. **The "Statement of Confidentiality" heading is grey.** It is a `Normal`
    paragraph with `Subtle Emphasis` (`#808080`), italic off, Segoe UI
    Semibold 14pt. Not an `NW Heading`, so it never appears in the TOC.
12. **The template's own instructions still say "Segue UI"** (body index
    104). Text-only typo in the disposable section; fixed in `_v2`.
13. **Vestigial even-page header/footer** (`header1`/`footer1`), see §4.
14. **The sensitivity label is present and kept.** `docProps/custom.xml`
    holds nine `MSIP_Label_8c8476dd-af43-4ac9-b73c-99d49e20c66a_*` properties
    (Name "Public", Method "Privileged", ContentBits 0, SetDate 2021-06-21)
    and `docMetadata/LabelInfo.xml` holds the matching `clbl:label`. The base
    scrubs only the non-label properties around them (ContentTypeId, Order,
    xd_Signature, xd_ProgID, ComplianceAssetId, TemplateUrl,
    _ExtendedDescription, TriggerFlowInfo, MediaServiceImageTags). Never
    describe the label as removed; the validator never flags it.
15. **Six `customXml` items** ship: `item1` is the live
    `CoverPageProperties` store (referenced by `storeItemID`), `item2` to
    `item6` are SharePoint content-type schema, taxonomy sync, form
    templates, library properties, and an empty bibliography, none
    referenced. Identify by `ds:itemID` against `storeItemID`, never by
    number (the SOW templates number them differently).
16. **`core.xml` leaks** a named employee in `dc:creator` and
    `lastModifiedBy`, `dc:subject` "Test Document Subtitle", a 2018
    `lastPrinted`; `app.xml TitlesOfParts` reads "Document Title Here". All
    scrubbed in the base (`creator` → "Netwoven").

---

## 7. Media

| File | Pixels | Role |
|---|---|---|
| `image1.png` | 236 x 34, RGBA | Netwoven wordmark (`#00549F` / `#B30738`). Header (`header2`) and cover. Byte-identical to the SOW `image7.png` and the deck `image5.png` |
| `image2.jpg` | 1536 x 1920 | Cover hero photo (Golden Gate Bridge, Pexels). Swappable per deliverable; the template's own guidance says so |
| `image3.png` to `image7.png` | 64 x 64, palette | Five white pillar icons on `#404040`: Content and Collaboration, Security and Compliance, Modern Applications, Data and Analytics, Cloud Infrastructure and Management |
| `image8.png` | 254 x 168 | "Member of Microsoft Intelligent Security Association" badge, cover only |
| `image9.png` to `image15.png` | 763x210, 763x205, 551x507, 450x263, 560x144, 627x150, 552x639 | Screenshots used only by the "How to use this template" section; removed with it |
| `hdphoto1.wdp` to `hdphoto5.wdp` | | Word picture-effect renditions; internal plumbing |

---

## 8. Document skeleton (body indexes of the shipped file)

| Body index | Content | Fate in a generated deliverable |
|---|---|---|
| 0 | Cover (`Cover Pages` SDT, `NoSpacing` paragraph) | Keep; fields set by `set_bound_fields`; swap `image2.jpg` if a better photo exists |
| 1 to 5 | "Statement of Confidentiality" heading, three legal paragraphs in the red dotted box, page break | Keep for client documents. `--internal` removes the page (Company becomes Netwoven) |
| 7 to 17 | "Document Revision History", "This document was prepared by:", `Author Name`, `Job Position`, `Netwoven Inc.`, a `Netwoven Table 1` 5x4 (Date / Version / Revision Description / Author with rows 0.1 to 1.0), page break | Optional. `--author` / `--role` fill the two lines; unchanged "Author Name" or "Job Position" is a validator warning. The template's own guidance drops this page for SOWs and proposals |
| 19 | `Table of Contents` SDT (`TOC \o "1-3" \h \z \u`), title in `TOCHeading` | Keep for documents over about three pages. `reset_toc("Table of Contents")` rebuilds it as a dirty field that Word fills on open |
| 20 to 29 | "List of Figures" (`TOC \h \z \c "Figure"`) | Removed unless `--keep-figure-lists` |
| 30 to 33 | "List of Tables" (`TOC \h \z \c "Table"`) | Removed unless `--keep-figure-lists` |
| 34 | Section break paragraph (six header/footer refs, `titlePg`) | Keep, always |
| 35 to 42 | `NW Heading 1` to `4` stubs with "Sample Normal Text" | Removed; the first stub is the marker where generated content starts |
| 43 to 123 | "How to use this template" (`NW Heading 1`), its subsections, three sample tables, captioned figures | **Always removed.** Never ship it |
| final `sectPr` | Section 2 properties | Keep |

`--minimal` produces the shortest skeleton the builder supports (cover plus
content; see `--help` for exactly which front-matter pages it drops).

---

## 9. Procedure for an AI

1. Run the analysis (`writing-method.md`) and settle the pyramid outline
   first.
2. Build the base document with the script (§10), passing title, client
   company, and date. Use `--internal` for Netwoven-internal documents.
3. Write content after the skeleton using the style IDs in §3.1: `NW Heading
   1` for key-line sections (each starts a new page), `NW Heading 2/3/4`
   beneath, `Normal` for prose, `ListParagraph` with `numId 7` for bullets,
   Netwoven table styles for every table, `Caption` with `SEQ` fields for
   every figure and table (`visual-standards.md` §10).
4. Never set a heading colour, a table cell fill, or a font by hand inside a
   Netwoven style; the style does it.
5. Do not touch `theme1.xml`, the header and footer parts, `settings.xml`
   beyond what the builder does, or any value in §2.
6. Validate (`python scripts/validate_deliverable.py out.docx --type doc
   [--internal] --client "Client, Inc."`). Rendering (`--render`) needs
   LibreOffice; the user runs that later.
7. Deliver the flag list in chat, never inside the document.

---

## 10. Programmatic recipe

Command line (fixed interface):

```
python scripts/new_deliverable_docx.py assets/bases/NW_Document_Base_2026.docx out.docx \
  --title "Tenant Assessment Report" --company "Contoso, Inc." --date "September 2026" \
  [--internal] [--keep-figure-lists] [--minimal] [--author "Name" --role "Title"]
```

Then fill the body with python-docx. The helpers ship next to the scripts.

```python
import docx
from docx.shared import Inches
from nw_docx_helpers import set_bound_fields, reset_toc

d = docx.Document("out.docx")            # already stripped by the builder

# Content starts after the skeleton; the builder left the final sectPr in place.
d.add_paragraph("Executive summary", style="NW Heading 1")      # auto-numbered "1", new page
d.add_paragraph("Contoso can complete the migration in 14 weeks if governance is fixed first.",
                style="Normal")
d.add_paragraph("Current state", style="NW Heading 2")          # "1.1", grey rule above
for item in ("1,240 site collections", "312 without an owner", "18 TB of content"):
    d.add_paragraph(item, style="List Paragraph")               # bullets come from numId 7 via the builder's list helper

t = d.add_table(rows=1, cols=3)
t.style = d.styles["Netwoven Table 2"]                          # banded, blue header
for c, h in zip(t.rows[0].cells, ("Workload", "Sites", "Size (TB)")):
    c.text = h
row = t.add_row().cells
row[0].text, row[1].text, row[2].text = "SharePoint", "1,240", "12.4"

# If the builder was run with --minimal and you need a TOC after all:
# reset_toc(d, "Table of Contents")

# Re-setting a bound field later (all SDT caches + stores + fullDate):
# set_bound_fields takes a FILE PATH, not the in-memory Document -- save first.
d.save("out.docx")
set_bound_fields("out.docx", title="Tenant Assessment Report", company="Contoso, Inc.", date="2026-09-15")
```

Notes for scripts:

- The final `sectPr` inherits header and footer from section 1. Do not delete
  the section-break paragraph at body index 34 or the header and footer
  vanish. The builder preserves it.
- `python-docx` cannot open the `.dotx`; use the `_Base_2026.docx`.
- Word does not care which template a `.docx` came from; the flipped base
  opens normally in Word.
- The generated file's TOC is a dirty field; Word prompts to update on open
  because `updateFields` is set. LibreOffice resolves it at conversion.

---

## 11. Changelog (as built by `tools/build_bases.py`)

Every operation below was applied and counted by the build; the parts-diff
whitelist, forbidden-string, LabelInfo and builder-marker gates all passed
(8 of 8). Rendering gates were not run (no LibreOffice); the zero-page-diff
claim for Stage A is to be confirmed by the user's render pass. Source
sha256 `758504da…`, `_v2` `08899fcc…`, base `6c6bdc19…`.

Stage A (intended to be invisible when rendered):

1. `customXml`: removed the 5 unreferenced items `item2` to `item6` with
   their `itemProps`, `.rels`, content-type overrides and document
   relationships (SharePoint content-type schema, taxonomy sync, form
   templates, library properties, empty bibliography); kept `item1`
   (`CoverPageProperties`, `{55AF091B-…}`).
2. `docProps/custom.xml`: removed 9 non-MSIP properties (ComplianceAssetId,
   ContentTypeId, MediaServiceImageTags, Order, TemplateUrl,
   TriggerFlowInfo, _ExtendedDescription, xd_ProgID, xd_Signature); kept
   all 7 `MSIP_Label_8c8476dd-…` properties (pids renumbered from 2).
   `docMetadata/LabelInfo.xml` is byte-identical to the original.
3. `docProps/core.xml`: creator and lastModifiedBy → "Netwoven",
   `lastPrinted` removed, `dc:subject` cleared, and **`dc:title` set to
   "Document Title Here"** (5 fields). Recorded deviation: the plan listed
   the title write under Stage B, but `core.xml` sits outside the Stage B
   whitelist, so it is written here; together with the Stage B SDT fix it is
   the approved visible change.
4. Unused styles removed (5): `KCITableNormal`, `PlainTable5`, `TOC4`,
   `TableGrid`, `UnresolvedMention`. Each passed the zero-reference check
   (no `pStyle`/`rStyle`/`tblStyle` use in any text part or numbering, no
   `basedOn`/`link`/`next` from a surviving style).
5. Numbering: 1 `w:num` removed (`numId 1`) and 1 orphan `w:abstractNum`
   (`4`).
6. rsids: 1078 `rsid*` attributes stripped from 10 text parts
   (`w14:paraId`/`textId`, SDT and bookmark ids, `anchorId`/`editId`
   untouched); 44 `w:style/w:rsid` elements removed; `w:rsids` (198
   entries) removed from both settings parts.
7. fontTable: 8 entries pruned with zero occurrences elsewhere. Recorded
   deviation: Word mirrors the font table into the glossary part, so
   `word/glossary/fontTable.xml` was pruned too (Yu Gothic Light, Arial
   Bold, Yu Mincho in both; Aptos and Aptos Display in the glossary copy
   only).

Stage B (approved fixes; Stage B vs Stage A changed exactly
`word/document.xml`, `word/footer2.xml`, `word/glossary/document.xml`,
`word/glossary/styles.xml`, `word/header2.xml`, `word/settings.xml`):

8. Glossary: 2 new placeholder docParts cloned from `B25901AB…`:
   `1AF6B431C7F15556BFCFFFE34AF26578` = `[Title]` and
   `B1C79B5DEAEC504D9AF1819457865F66` = `[Publish Date]`; the existing three
   parts untouched (shared with the body Company SDTs); 2 paragraph styles
   cloned in `glossary/styles.xml`.
9. `header2` Title SDT: placeholder re-pointed to `[Title]`,
   `showingPlcHdr` removed, cached text `[Company]` → "Document Title Here".
   Recorded deviation: the run's `rPr` was replaced by the SDT's `rPr`, so
   the `PlaceholderText` character style (grey `#808080`) no longer applies
   to the header title; it renders in the header's own 10pt Segoe UI.
10. `footer2` Publish Date SDT: placeholder re-pointed to `[Publish Date]`;
    `<w:dateFormat w:val="M/d/yyyy"/>` added as first child of `w:date`.
11. Cover Title SDTs (Choice and Fallback copies): `showingPlcHdr` removed,
    cached text → "Document Title Here" (2).
12. Cover Publish Date SDTs: stale `w:fullDate="2017-12-28T00:00:00Z"`
    dropped (2).
13. "Segue UI" → "Segoe UI" in the instructional paragraph (2 runs).
14. `settings.xml`: `<w:updateFields w:val="true"/>` inserted before
    `hdrShapeDefaults`.
15. fontTable recheck after Stage B: 0 further prunes.

Stage C: `[Content_Types].xml` main part flipped `template.main+xml` →
`document.main+xml`; every other part byte-identical to `_v2`; python-docx
opens the base and all builder markers were found (`NWHeading1` stub at
body[35], "List of Figures" at body[20], section `sectPr` with six
references at body[34], final `sectPr` with none, TOC gallery SDT,
"Statement of Confidentiality" and "Document Revision History" markers,
`<PublishDate/>` self-closing in `item1`, footer `w:date`, `app.xml`
Company).

Parts diff: Stage A vs original changed 20 parts and removed 15; Stage B
vs A changed 6; Stage C vs B changed 1.

### Not changed (by decision)

- Purview/MSIP sensitivity label: `docMetadata/LabelInfo.xml`, its
  relationship and content type, and every `MSIP_Label_*` property in
  `docProps/custom.xml`.
- Even-page `header1.xml` / `footer1.xml` (referenced by the section-1
  `sectPr`).
- Heading numbering formats (`numbering.xml` levels, `numId`s on
  `Heading1` to `Heading9`).
- `word/theme/theme1.xml` (`NW Heading 1/2/4` and `TOC1` carry
  `themeColor=accent1`).
- `docProps/app.xml` (`TitlesOfParts` and `Company` are rewritten by the
  builder).
- The confidentiality year `©2001 - 2026` (literal text) and the cover hero
  photo.

### Recommendation reversed

The 2021 guide suggested that retheming `theme1.xml` to Netwoven colours
would be safe because "nothing reads from it". That was wrong (§1): heading
colours and three fonts resolve through the theme. The recommendation is
withdrawn. Leave the theme exactly as shipped.
