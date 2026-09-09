---
title: Netwoven Brand Tokens (2026 templates)
covers: NW_document_template_2026.dotx, NW_milestone_SOW_template_2026.dotx, NW_TM_SOW_template_2026.dotx, NW_presentation_template_2026.potx and their _Base_2026 derivatives
last_verified: 2026-09-02
---

# Netwoven brand tokens

This file is the single source of truth for every colour, font, and chart
palette the plugin uses. The three template guides (`word-template-guide.md`,
`sow-template-guide.md`, `ppt-template-guide.md`) describe where each token is
wired inside a file; this file decides which token to reach for. Chart palettes
are defined here and nowhere else. Every value below was read from the
template XML or sampled from the logo pixels; nothing is estimated.

Values marked "to be confirmed by render" have been derived from XML but not
yet seen on a rendered page in this build.

---

## 1. The one thing to understand first

Netwoven's four 2026 templates carry three different colour systems that all
look "blue". They are not interchangeable.

| System | Primary value | Lives in | Use it for |
|---|---|---|---|
| Logo pixels | `#00549F` blue, `#B30738` red | The wordmark PNG shipped in all four files (`236x34`, identical bytes) | Nothing but the logo itself. Never recolour a shape to `#00549F` and call it brand blue. Chart palettes borrow both values as categorical colours only. |
| Word chrome | `#145CA4` rules and table headers, `#2F5496` headings | Document and SOW templates | Word deliverables only. `#2F5496` is theme `accent1` at shade `BF`, so the theme is load-bearing for heading colour. Never retheme a Word file. |
| Deck accent | `#00B0F0` cyan, `#32355F` indigo, `#595959` text, `#D0CECE` dividers | Presentation template ("Netwoven Default Theme") | Decks only. The deck keeps `#00B0F0` exactly as designed. Do not "correct" it toward `#145CA4` or `#00549F`. |

Decision table for a new element:

| You are adding | In a Word document or SOW | In a deck |
|---|---|---|
| A table header fill | Use a Netwoven table style (Table 1 or 2 give `#145CA4`) | `#00B0F0` fill, white bold 12pt text (see contrast note in §5) |
| A heading | Use the heading style; never set a colour | Leave the title placeholder black |
| A chart series | Document categorical palette (§4.2) | Deck categorical palette (§4.1) |
| A callout or stat number | `#145CA4` | `#00B0F0` (Three Stat layout already does this) |
| A label or secondary heading | `#2F5496` via style, never direct | `#32355F` |
| A divider line | `#145CA4` 3pt is reserved for header and footer rules; body dividers are `#BFBFBF` 0.5pt | `#D0CECE` 0.75pt |
| Body text | Automatic (black) | `#595959` (already the body default) |
| Something red | `#B30738` sparingly, for risk or the diverging palette | `#B30738` sparingly, same rule |

---

## 2. Colour master table

### 2.1 Shared by all four templates

| Token | Hex | Source | Where it appears |
|---|---|---|---|
| Logo blue | `#00549F` | Sampled, 2333 of the wordmark's opaque pixels | The word "Netwoven" in the wordmark; the 25-year badge in the deck master |
| Logo red | `#B30738` | Sampled, 120 pixels | The swoosh in the wordmark; the badge |
| White | `#FFFFFF` | theme `lt1` | Page and slide backgrounds, text on dark fills |
| Black | `#000000` | theme `dk1` | Body text (Word), titles (deck) |
| Word theme `dk2` / `text2` | `#44546A` | theme1.xml (stock Office in the Word files) | Caption and No Spacing text in Word |
| Word theme `lt2` / `background2` | `#E7E6E6` | theme1.xml | Base for `#AEAAAA` (shade BF) and `#3B3838` (shade 40) |
| Deck theme `dk2` | `#44546A` | deck theme1.xml | Base for the `#D6DCE5` panel on Big Statement (lumMod 20, lumOff 80) |
| Deck theme `lt2` | `#E7E6E6` | deck theme1.xml | Base for `#D0CECE` dividers (lumMod 90) |

### 2.2 Word document and SOW chrome

| Token | Hex | XML derivation | Where it appears |
|---|---|---|---|
| Netwoven rule blue | `#145CA4` | literal `w:color` / `w:fill` | 3pt (`sz=24`) header bottom rule and footer top rule on every interior page; `Netwoven Table 1` and `2` header fills; `Netwoven Table 3` header text; `NW Heading 1` underline colour (draws nothing, no underline style) |
| Heading blue | `#2F5496` | `accent1` (`#4472C4`) with `themeShade=BF` | `Heading1`, `Heading2`, `Heading4`, `NW Heading 1`, `NW Heading 2`, `TOC1`, `TOC Heading`; SOW `Heading1/2/3` |
| Heading 3 base blue | `#1F3763` | `accent1` shade `7F` | Base `Heading3` only; `NW Heading 3` and harmonized SOW `Heading3` override to `auto` (black) |
| Heading 4 grey | `#3B3838` | `background2` shade `40` | `NW Heading 4` |
| Heading 2 rule grey | `#AEAAAA` | `background2` shade `BF` | 0.5pt top rule on `NW Heading 2` and SOW `Heading2` |
| Caption grey | `#44546A` | `text2` | `Caption`, `No Spacing` |
| Table grid | `#262626` | `text1` tint `D9` | `Netwoven Table 1` borders, 0.5pt |
| Light grid | `#BFBFBF` | `background1` shade `BF` | `Plain Table 1` borders inherited by `Netwoven Table 2` |
| Header underline | `#7F7F7F` | `text1` tint `80` | `Netwoven Table 3` header bottom border |
| Row band | `#F2F2F2` | `background1` shade `F2` | Alternate rows in `Netwoven Table 2` and `3`, cover pillar labels |
| Pillar band | `#404040` | literal | The five service-pillar cells on the cover (document and SOW) |
| Cover title rule | `#C00000` | literal | 3pt cell border under the cover title block |
| Confidentiality border | `#FF0000` | literal | 0.5pt dotted box around the three confidentiality paragraphs |
| Placeholder text | `#808080` | `Placeholder Text` style | Unfilled content-control prompts |
| Hyperlink | `#0563C1` | theme `hlink` | Hyperlinks |
| SOW footer label | `#0070C0` | literal | "Netwoven Confidential" footer line; roles-table header fill |
| SOW Biz Apps header | `#92D050` | literal | Biz Apps table header fill |
| SOW PSO Grid header | `#D9D9D9` | `background1` shade `D9` | Deliverables table header and first column |
| SOW products table | `#9CC2E5` borders, `#DEEAF6` bands, `#2E74B5` text | `accent5` tints and shade | `Grid Table 6 Colorful Accent 5` (products checklist) |
| SOW navy rule (removed) | `#000080` | literal | 1.5pt bottom rule under the shipped SOW `Heading1`; removed by the approved harmonization (see `sow-template-guide.md` §3) |
| TM TOC level-2 | `#002060` | literal | `TOC2` in the T&M variant only |

### 2.3 Deck (Netwoven Default Theme)

| Token | Hex | XML derivation | Where it appears |
|---|---|---|---|
| Accent cyan | `#00B0F0` | theme `accent1`, also hardcoded 17 times in layouts | Cover side panel (L1, L3), cover title 48pt, Three Stat numerals 96pt, Three Image Horizontal captions, Two Photo labels, Testimonial On Right bar, Closing Slide band |
| Label indigo | `#32355F` | hardcoded 11 times | Category labels on Five Part Content (L14) and Six Icon (L28) |
| Body grey | `#595959` | `tx1` lumMod 65 lumOff 35 | Body text levels 1 to 5, footer contact text, subtitles |
| Divider grey | `#D0CECE` | `bg2` lumMod 90 | Footer vertical dividers |
| Placeholder prompt grey | `#808080` | `tx1` lumMod 50 lumOff 50 | "Image Holder" prompt text in picture placeholders |
| Item grey | `#404040` | hardcoded | Item text on 2 Photo on right (L7) |
| Panel tint | `#D6DCE5` | `tx2` lumMod 20 lumOff 80 | Left panel on Big Statement with Illustration (L25) |
| Testimonial subtext | `#E5F2F3` | hardcoded | "Position, Company" line on L19/L20 |
| Cover group | `#375991` | hardcoded | A fill inside the cover logo group on L1/L3 |
| Agenda rings | `#88B9BE` | hardcoded 3pt outline | Two decorative rounded rectangles on Agenda (L4) |
| Eyebrow one-off | `#0055A0` | hardcoded | Optional eyebrow on L25. A near miss of logo blue, kept as designed; do not reuse elsewhere |
| Theme accent2 | `#7030A0` | theme | Bar on Testimonial On Left (L20), icon prompt text on Six Icon |
| Theme accent3 | `#7F7F7F` | theme | Unused by layouts |
| Theme accent4 | `#FFC000` | theme | Gradient stop on the instructional slide 1 only |
| Theme accent5 | `#375623` | theme | Unused by layouts |
| Theme accent6 | `#70AD47` | theme | Unused by layouts |

Do not use the theme accent cycle for charts. Position 3 is a neutral grey,
position 5 is a near-black green, and white text fails contrast on positions
1, 4, and 6. Use §4 instead.

---

## 3. Typography

| Role | Document | SOW (as shipped, then harmonized) | Deck |
|---|---|---|---|
| Body | Segoe UI 11pt (`Normal`) | Segoe UI 11pt (`Normal`); `Body Text` style is Verdana 11pt with a 0.5in indent | Segoe UI 20/18/16/14/14pt, `#595959` (master body levels 1 to 5) |
| Headings | Segoe UI Semibold 22/18/14/12pt (`NW Heading 1` to `4`) | Shipped `Heading1/2/3` are Segoe UI bold small caps 18/16/16pt; harmonized to Segoe UI Semibold 22/18/14pt, small caps on 1 and 2, `Heading3` black | Segoe UI 40pt black (master title); Section Header 60pt; cover 48pt cyan |
| Table header | Segoe UI Semibold 12pt white (Table 1, 2); Calibri bold caps 12pt `#145CA4` (Table 3, via `minorHAnsi`) | Same styles; Table 1/2 header font was `Sitka Banner` (substitution defect) and is fixed to Segoe UI Semibold in the base | No table style defined; builder writes Segoe UI bold 12pt white on `#00B0F0` |
| Caption | Segoe UI 9pt italic `#44546A` | same | 12pt `#595959` source line (builder) |
| Small print | `No Spacing` is Calibri 10pt `#44546A` (theme minor font, not Segoe UI) | same | Footer 11pt Segoe UI `#595959` |
| Numbers as visuals | 24pt Segoe UI Semibold `#145CA4` (builder) | not used | Three Stat 96pt `#00B0F0`, letter spacing `-300` |
| Theme fonts | Calibri Light / Calibri (stock Office; reached only by `TOC Heading`, `No Spacing`, Table 3 header) | same | Segoe UI / Segoe UI ("Custom 1") |

Portability caveat. Segoe UI and Segoe UI Semibold are Microsoft fonts and
are not embedded. Windows and Microsoft 365 render them natively. macOS,
Linux, LibreOffice and Google Docs substitute, which shifts line breaks and
small-caps rendering. `brew install --cask font-selawik` plus a fontconfig
alias gives a metric-compatible stand-in for render checks; the user runs this
later.

---

## 4. Chart palettes (defined only here)

Scripts import these from `nw_pptx_helpers.py` (`PALETTE_*` constants), which
must match this table byte for byte.

### 4.1 Deck categorical (`PALETTE_DECK`)

| Order | Hex | Role | Contrast of white text on it |
|---|---|---|---|
| 1 | `#00B0F0` | accent cyan | 2.48:1, fails AA. Use dark labels or place labels outside |
| 2 | `#32355F` | label indigo | 11.60:1 |
| 3 | `#00549F` | logo blue | 7.58:1 |
| 4 | `#B30738` | logo red, reserve for the "bad" or "risk" series | 6.99:1 |
| 5 | `#0080B0` | mid blue, fixed value close to accent1 at 75 percent luminance (`#0084B4`) | 4.45:1, large text only |
| 6 | `#93E2FF` | accent1 lumMod 40 lumOff 60 | fails; use `#32355F` labels (8.05:1) |

Use the palette in order. Two series use 1 and 2; a highlight series uses 1
against everything else in `#D0CECE`.

### 4.2 Document categorical (`PALETTE_DOC`, to be confirmed by render)

| Order | Hex | Role |
|---|---|---|
| 1 | `#145CA4` | Word rule blue |
| 2 | `#B30738` | logo red |
| 3 | `#32355F` | indigo |
| 4 | `#0080B0` | mid blue |
| 5 | `#93E2FF` | light cyan |
| 6 | `#7F7F7F` | neutral grey |

### 4.3 Sequential (`PALETTE_SEQ`)

`#93E2FF` → `#00B0F0` → `#0080B0` → `#32355F` (light to dark, four steps).
For a five-step ramp insert `#005878` (accent1 lumMod 50) between the last
two.

### 4.4 Diverging (`PALETTE_DIV`)

`#B30738` ← `#F2F2F2` → `#00549F`. Red is the negative end; the neutral
midpoint is the Word row-band grey so it reads as "no change" in both media.

### 4.5 Chart chrome (`CHROME_*`)

| Element | Value |
|---|---|
| Axis and legend text | Segoe UI 12pt `#595959` |
| Gridlines | `#D0CECE`, 0.75pt, horizontal only |
| Axis lines | `#D0CECE` |
| Chart title | none (the slide title carries the message) |
| Legend | bottom, or direct labels when six or fewer points |
| Data labels | 12pt `#595959`, `#FFFFFF` only on `#32355F`, `#00549F`, `#B30738` |

### 4.6 python-pptx snippet for per-series colours

```python
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

PALETTE_DECK = ["00B0F0", "32355F", "00549F", "B30738", "0080B0", "93E2FF"]
CHROME_TEXT, CHROME_GRID = "595959", "D0CECE"

def add_chart(slide, x, y, cx, cy, categories, series, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
    data = CategoryChartData()
    data.categories = categories
    for name, values in series:
        data.add_series(name, values)
    gf = slide.shapes.add_chart(chart_type, x, y, cx, cy, data)
    chart = gf.chart
    chart.has_title = False
    chart.font.name, chart.font.size = "Segoe UI", Pt(12)
    chart.font.color.rgb = RGBColor.from_string(CHROME_TEXT)
    if chart_type in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        pts = chart.plots[0].series[0].points
        for i in range(len(categories)):
            pts[i].format.fill.solid()
            pts[i].format.fill.fore_color.rgb = RGBColor.from_string(PALETTE_DECK[i % 6])
    else:
        for i, s in enumerate(chart.plots[0].series):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = RGBColor.from_string(PALETTE_DECK[i % 6])
        va = chart.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = RGBColor.from_string(CHROME_GRID)
        va.format.line.color.rgb = RGBColor.from_string(CHROME_GRID)
        chart.category_axis.format.line.color.rgb = RGBColor.from_string(CHROME_GRID)
    chart.has_legend = len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    return chart
```

Every series gets an explicit `srgbClr`. A series left on the theme cycle is
a validator warning ("chart series without explicit colour").

---

## 5. WCAG contrast table

Ratios computed with the WCAG 2.x relative-luminance formula. AA normal text
needs 4.5:1, AA large text (18pt regular or 14pt bold, and any deck title)
needs 3:1, AAA needs 7:1.

| Foreground on background | Ratio | AA normal | AA large | Use |
|---|---|---|---|---|
| `#FFFFFF` on `#00B0F0` | 2.48:1 | fail | fail | Never for text. The cover title is cyan on white (same 2.48:1) and passes only as a 48pt decorative title; keep it short |
| `#000000` on `#00B0F0` | 8.47:1 | pass | pass | Labels inside cyan tiles |
| `#32355F` on `#00B0F0` | 4.68:1 | pass | pass | Alternative label on cyan |
| `#FFFFFF` on `#32355F` | 11.60:1 | pass | pass | Preferred dark tile |
| `#FFFFFF` on `#00549F` | 7.58:1 | pass | pass | |
| `#FFFFFF` on `#B30738` | 6.99:1 | pass | pass | |
| `#FFFFFF` on `#0080B0` | 4.45:1 | fail | pass | Large text only |
| `#FFFFFF` on `#145CA4` | 6.78:1 | pass | pass | Word table headers |
| `#FFFFFF` on `#2F5496` | 7.42:1 | pass | pass | |
| `#FFFFFF` on `#404040` | 10.37:1 | pass | pass | Cover pillar band |
| `#FFFFFF` on `#0070C0` | 5.15:1 | pass | pass | SOW roles header |
| `#FFFFFF` on `#7030A0` | 8.02:1 | pass | pass | Testimonial On Left bar |
| `#595959` on `#FFFFFF` | 7.00:1 | pass | pass | Deck body text |
| `#595959` on `#F2F2F2` | 6.26:1 | pass | pass | Banded table rows |
| `#44546A` on `#FFFFFF` | 7.71:1 | pass | pass | Captions |
| `#2F5496` on `#FFFFFF` | 7.42:1 | pass | pass | Word headings |
| `#145CA4` on `#FFFFFF` | 6.78:1 | pass | pass | Table 3 header text |
| `#00549F` on `#FFFFFF` | 7.58:1 | pass | pass | |
| `#B30738` on `#FFFFFF` | 6.99:1 | pass | pass | |
| `#32355F` on `#FFFFFF` | 11.60:1 | pass | pass | Deck labels |
| `#0080B0` on `#FFFFFF` | 4.45:1 | fail | pass | Chart fill, not text |
| `#7F7F7F` on `#FFFFFF` | 4.00:1 | fail | pass | Grey series, not text |
| `#00B0F0` on `#FFFFFF` | 2.48:1 | fail | fail | Stat numerals pass only because 96pt is decorative; always pair with a `#000000` label |
| `#D0CECE` on `#FFFFFF` | 1.57:1 | n/a | n/a | Gridlines only, never text |
| `#000000` on `#93E2FF` | 14.56:1 | pass | pass | |
| `#32355F` on `#93E2FF` | 8.05:1 | pass | pass | |
| `#000000` on `#D9D9D9` | 14.88:1 | pass | pass | PSO Grid header |
| `#000000` on `#92D050` | 11.36:1 | pass | pass | Biz Apps header |
| `#E5F2F3` on `#00B0F0` | 2.16:1 | fail | fail | Template's own testimonial subtext on the cyan bar; keep it to one short line, as designed |
| `#FFFFFF` on `#70AD47` | 2.71:1 | fail | fail | Why theme accent6 is banned for charts |
| `#FFFFFF` on `#FFC000` | 1.64:1 | fail | fail | Why theme accent4 is banned for charts |

Rules that follow from the table:

1. Never put white text on `#00B0F0`. Cyan tiles take `#000000` or `#32355F`
   labels; a cyan table header takes white text only at 12pt bold and is a
   builder default that the user should confirm by render; if it reads
   poorly, switch the header fill to `#32355F`.
2. Meaning is never carried by colour alone; pair colour with a label, a
   marker shape, or position.
3. Gridlines and dividers are the only place `#D0CECE` appears.
