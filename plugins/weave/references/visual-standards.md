---
title: Visual Standards
covers: charts, tables, frameworks, stat tiles, icons, and figures for decks on NW_Presentation_Base_2026.pptx and documents on NW_Document_Base_2026.docx, NW_SOW_Milestone_Base_2026.docx, NW_SOW_TM_Base_2026.docx
last_verified: 2026-09-02
---

# Visual standards

This file says how a visual is chosen, sized, coloured, and labelled. Colour
values are the tokens in `brand-tokens.md` §4 (palettes) and §5 (contrast);
they are not repeated here except inside code. Layout names and placeholder
indexes are from `ppt-template-guide.md`. Every geometry value marked "confirm
by render" was derived from the layout XML and has not yet been checked on a
rendered slide in this build.

Units. python-pptx uses EMU. `Inches(1) == 914400`. The slide is 13.333in by
7.5in (`12192000 x 6858000` EMU).

---

## 1. Message first, then comparison, then chart

Decide the message (the slide's action title), name the comparison it makes,
then pick the form. Never start from the data shape.

| Comparison | Message pattern | Chart | Avoid |
|---|---|---|---|
| Component (share of a whole) | "X is 62 percent of Y" | Stacked bar (one bar) or pie with 5 or fewer slices | Pie with more than 5 slices; 3-D pie; donut with the total missing |
| Item (ranking) | "A is larger than B" | Horizontal bar, sorted | Vertical bars with long labels; radar |
| Time series (change) | "X rose 30 percent since 2024" | Line (many periods) or column (few periods) | Area charts; dual axes |
| Frequency (distribution) | "Most sites are under 5 GB" | Column histogram | Line for discrete bins |
| Correlation (relationship) | "Effort tracks item count" | Scatter | Two bar series side by side |

When a table beats a chart:

- The audience needs exact values (fees, dates, hours). SOW fee and rate
  tables are always tables and never charts.
- There are two or fewer data points.
- The comparison is across two dimensions with different units.
- The values are text (owner, status, decision).

---

## 2. Chart mechanics (python-pptx)

Charts go on `Title Only` (L12) inside the content area below.

```python
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

CONTENT = dict(x=Inches(0.92), y=Inches(1.6), cx=Inches(11.5), cy=Inches(5.0))  # confirm by render
SOURCE_Y = Inches(6.55)  # source line baseline, above the footer bar at ~6.9in
```

Rules applied by `nw_pptx_helpers.add_chart` and to be kept when drawing by
hand:

1. Per-series colours from `PALETTE_DECK` in order; per-point colours for pie
   and doughnut. Highlight pattern: series of interest in `#00B0F0`, all
   others `#D0CECE`.
2. No chart title. The slide title is the message.
3. Segoe UI 12pt `#595959` for every text element in the chart.
4. Horizontal gridlines only, `#D0CECE`; no vertical gridlines; no chart
   border; no plot-area fill.
5. Value axis starts at zero for bars and columns. Line charts may truncate
   the axis only when the title says so ("index, 2024 = 100").
6. Direct data labels when six or fewer points; otherwise a bottom legend.
   Labels use `number_format` from §8.
7. Sort bars by value unless the categories have a natural order (time,
   stage).
8. One chart per slide. Two small multiples are allowed only when they share
   axes and the title compares them.
9. Alt text on the graphic frame (see §7).
10. Source line (see §9).

Column and bar charts use gap width 80 (`chart.plots[0].gap_width = 80`) and
no overlap. Line charts use 2.25pt lines, no markers above 12 points, round
markers 7pt below that.

---

## 3. Framework recipes

All recipes draw into the L12 content area (x 0.92in, y 1.6in, w 11.5in,
h 5.0in; confirm by render). Shapes use `MSO_SHAPE` autoshapes with 0.75pt
`#D0CECE` outlines or no outline, Segoe UI text, and the fills named below.
Text inside shapes is 14pt minimum, `#000000` or `#32355F` on light fills,
`#FFFFFF` on `#32355F`, `#00549F`, `#B30738` fills only.

Helper used by every recipe:

```python
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def box(slide, shape, x, y, cx, cy, text="", fill="FFFFFF", line="D0CECE", font_color="000000", size=14, bold=False, align=PP_ALIGN.CENTER):
    s = slide.shapes.add_shape(shape, x, y, cx, cy)
    s.fill.solid(); s.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line: s.line.color.rgb = RGBColor.from_string(line); s.line.width = Pt(0.75)
    else: s.line.fill.background()
    s.shadow.inherit = False
    tf = s.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.1)
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name, r.font.size, r.font.bold = "Segoe UI", Pt(size), bold
    r.font.color.rgb = RGBColor.from_string(font_color)
    return s
```

### 3.1 Two-by-two matrix

Four quadrants of 5.55in by 2.3in with a 0.2in gutter, axis labels outside.

```python
X0, Y0, W, H, G = Inches(1.4), Inches(1.9), Inches(5.45), Inches(2.25), Inches(0.2)
fills = {"tl": "F2F2F2", "tr": "93E2FF", "bl": "F2F2F2", "br": "F2F2F2"}   # highlight one quadrant
cells = {"tl": (X0, Y0), "tr": (X0 + W + G, Y0), "bl": (X0, Y0 + H + G), "br": (X0 + W + G, Y0 + H + G)}
for k, (x, y) in cells.items():
    box(slide, MSO_SHAPE.RECTANGLE, x, y, W, H, labels[k], fill=fills[k], line=None, align=PP_ALIGN.LEFT)
# axis labels: y-axis rotated text box at x 0.92in, x-axis text box at y 6.5in, both 12pt #595959
```

Place items inside quadrants as 12pt text or small circles (0.25in) of
`#32355F`; never more than 12 items on a matrix.

### 3.2 Chevron process

Three to six steps across the full width. For `n` steps, width is
`(11.5 - 0.1 * (n - 1)) / n` inches, height 1.1in at y 2.4in, with a
description box beneath each chevron.

```python
n = len(steps)
w = (11.5 - 0.1 * (n - 1)) / n
for i, (title, desc) in enumerate(steps):
    x = Inches(0.92 + i * (w + 0.1))
    box(slide, MSO_SHAPE.CHEVRON, x, Inches(2.4), Inches(w), Inches(1.1), title,
        fill="00B0F0" if i == highlight else "32355F", line=None, font_color="000000" if i == highlight else "FFFFFF", bold=True)
    box(slide, MSO_SHAPE.RECTANGLE, x, Inches(3.7), Inches(w), Inches(2.2), desc, fill="FFFFFF", line=None, align=PP_ALIGN.LEFT, size=12)
```

Use `MSO_SHAPE.PENTAGON` for the first step so the row starts flat.

### 3.3 Maturity staircase

Four or five ascending blocks from bottom left to top right, each block
2.2in wide; step height 0.9in; the current state in `#00B0F0`, target in
`#32355F`, others `#F2F2F2`.

```python
levels = ["Ad hoc", "Managed", "Defined", "Measured", "Optimised"]
bw, sh = Inches(2.2), Inches(0.9)
for i, name in enumerate(levels):
    x = Inches(0.92) + i * bw
    h = sh * (i + 1)
    y = Inches(6.4) - h
    fill = "00B0F0" if i == current else ("32355F" if i == target else "F2F2F2")
    box(slide, MSO_SHAPE.RECTANGLE, x, y, bw - Inches(0.05), h, name, fill=fill, line=None,
        font_color="FFFFFF" if fill == "32355F" else "000000", bold=True)
```

Label "Today" and "Target" with 12pt `#595959` text boxes above the two
highlighted blocks.

### 3.4 Before and after

Two columns of 5.55in with a 0.4in gutter at x 0.92in and x 6.87in; column
headers "Today" (fill `#F2F2F2`) and "Target state" (fill `#93E2FF`) at
y 1.6in, height 0.6in; three to five paired rows beneath, 0.8in each, joined
by a right arrow (`MSO_SHAPE.RIGHT_ARROW`, 0.3in by 0.3in, fill `#00B0F0`) in
the gutter. Rows must be parallel: the same attribute on both sides.

### 3.5 Timeline and roadmap

Horizontal axis of months or quarters along y 2.3in; a 0.05in `#D0CECE` line
across the full width with period labels above (12pt `#595959`). Swimlanes
of 0.9in height stacked beneath, lane name in a 1.6in left column
(`#F2F2F2`), bars as rounded rectangles (`MSO_SHAPE.ROUNDED_RECTANGLE`)
filled `#32355F` for committed work and `#93E2FF` for tentative work, 12pt
labels inside or to the right. Milestones are `MSO_SHAPE.DIAMOND` 0.25in in
`#B30738` for a decision gate, `#00B0F0` for a delivery. Six lanes maximum.

Compute x from dates:

```python
def x_for(date, start, end, x0=Inches(2.6), width=Inches(9.8)):
    return x0 + int(width * (date - start).days / (end - start).days)
```

### 3.6 Harvey balls

Use for qualitative ratings on a table of options. Draw a 0.3in circle
outline in `#32355F` and overlay a `MSO_SHAPE.PIE` (or `CHORD`) with the same
bounds and adjustment to the quarter value, filled `#32355F`. Zero is an
empty circle; four is a full circle. Always add a legend row explaining the
scale, and never rate more than six options by more than six criteria.

### 3.7 RACI or heat table

Build with `add_table` (or `slide.shapes.add_table`) and fill cells directly.

| Value | Fill | Text |
|---|---|---|
| R (Responsible) | `#32355F` | `#FFFFFF` bold |
| A (Accountable) | `#00549F` | `#FFFFFF` bold |
| C (Consulted) | `#93E2FF` | `#000000` |
| I (Informed) | `#F2F2F2` | `#000000` |
| Heat high / medium / low | `#B30738` / `#93E2FF` / `#F2F2F2` | `#FFFFFF` / `#000000` / `#000000` |

Letters or values stay in the cells; colour never carries the meaning alone.

---

## 4. Stat tiles (`Three Stat`, L26)

| Element | Spec (from the layout) |
|---|---|
| Layout | `Three Stat`, no title placeholder |
| Numerals | idx 10, 15, 17; 96pt Segoe UI, `#00B0F0`, letter spacing `-300`, bottom anchored at y 2.18in, 3.95in wide |
| Labels | idx 11, 14, 16; 21pt Segoe UI black, top anchored at y 4.47in |
| Content | A number with its unit ("42 %", "$1.2M", "3 wks"); the label is a full clause ("sites migrated in wave 1"), 6 words or fewer |
| Rule | Three numbers that together make one point. If the point needs prose, add a 14pt `#595959` text box at y 5.4in, 11.5in wide, one sentence |

For one or two numbers use `Big Statement with Illustration` (L25): statement
idx 16 (72pt), eyebrow idx 14 (21pt, the template's `#0055A0`), picture
idx 13. Use `set_paragraph_no_bullet` on L25 and L26 text so the layout's
zero-width-space bullets stay hidden; do not edit the layout `pPr`.

---

## 5. Icons

- Meaning only. An icon labels a category; it never decorates.
- Monochrome `#32355F` (the template's own label colour), or `#00B0F0` for
  the one highlighted item.
- PNG at 512px or larger (the master's own footer glyphs are 512px PNGs).
  SVG only with a PNG fallback: the template's own SVGs have none and
  LibreOffice may drop them, so the builder never adds new SVG-only
  pictures.
- Same visual family on one slide (all outline or all filled), same size
  (0.66in on `Six Icon`, matching idx 19 to 27).
- Alt text as in §7. Decorative icons are removed rather than described.

---

## 6. Photography

Cover, section headers, testimonials, and closers only. Never on evidence
slides. Landscape, no text baked in, licensed (the Word template's own
guidance points to free stock sources and requires a paid or free licence).
Crop to the placeholder rather than stretching: set `picture.crop_*` or use
`ph(slide, idx).insert_picture(path)` which crops to fill.

---

## 7. Colour and accessibility

1. Palettes only from `brand-tokens.md` §4; any `srgbClr` outside the
   approved set is a validator warning.
2. Contrast per `brand-tokens.md` §5. No white text on `#00B0F0`.
3. Never colour alone: pair with a label, marker, pattern, or position.
4. Alt text on every chart, picture, and framework group. python-pptx exposes
   no property for it, so write the `descr` attribute:

```python
def set_alt_text(shape, text):
    # works for pictures, graphic frames (charts, tables), autoshapes, and groups
    cNvPr = shape._element.xpath("./*[local-name()='nvSpPr' or local-name()='nvPicPr' or local-name()='nvGraphicFramePr' or local-name()='nvGrpSpPr']/*[local-name()='cNvPr']")[0]
    cNvPr.set("descr", text)
    cNvPr.set("name", text[:60])
```

Alt text states the message and the data ("Column chart: migration effort
by workload, SharePoint 48 percent, Teams 32 percent, Exchange 20 percent").

5. Reading order follows z-order. Add title first, visual second, source
   last.
6. Minimum 14pt on slides for anything the audience must read; 12pt only for
   chart chrome, table cells, and source lines. Minimum 9pt in Word captions
   (the `Caption` style).

---

## 8. Number formatting

| Kind | Format | Example |
|---|---|---|
| Counts | thousands separator, no decimals | 12,480 sites |
| Percentages | whole numbers; one decimal only when the difference matters | 18 percent in prose; `18%` inside charts and tables |
| Currency | `$` prefix, thousands separator, no cents above $1,000; `$1.2M` and `$480K` in charts | $48,500; $1.2M |
| Hours and weeks | digits with unit | 320 hours; 12 weeks |
| Dates | Month D, YYYY in prose; `M/d/yyyy` where the template's footer date lives; ISO `YYYY-MM-DD` in PublishDate metadata | September 15, 2026 |
| Ranges | "to" or "through", never a dash | 3 to 5 weeks |
| Rounding | Round to the precision the source supports; never add precision | |
| Chart `number_format` | `#,##0` counts, `0%` shares, `$#,##0` money, `$#,##0.0,,"M"` millions | |

Values quoted from a source keep the source's exact figure (see
`writing-method.md` Phase C).

---

## 9. Source lines

Every chart, table, and stat slide carries a source line: a text box at
x 0.92in, y 6.55in, 11.5in wide, 0.3in high, Segoe UI 10.5pt `#595959`, left
aligned, reading `Source: <system or document>, <date>` and, when relevant,
`n = <count>`. The text "Source:" is the one place a mid-line colon is
allowed on a slide. Estimates say "Netwoven estimate" and the basis.

---

## 10. Word-side figures and tables

### 10.1 Geometry

Letter page, 1in margins, so the text width is 6.5in (9360 twips). Every
figure is 6.5in wide or narrower and centred. Height 3.5in to 4in for a
full-width chart; never taller than 6in so caption and figure stay together.

### 10.2 Captions and fields

Figures are captioned below, tables above, both in the `Caption` style with
a `SEQ` field so Word numbers them and the List of Figures and List of
Tables update:

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_caption(doc, label, text):           # label is "Figure" or "Table"
    p = doc.add_paragraph(style="Caption")
    p.add_run(f"{label} ")
    r = p.add_run()
    for tag, attrs, txt in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                            ("w:instrText", {"xml:space": "preserve"}, f" SEQ {label} \\* ARABIC "),
                            ("w:fldChar", {"w:fldCharType": "separate"}, None),
                            ("w:t", {}, "1"),
                            ("w:fldChar", {"w:fldCharType": "end"}, None)):
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(qn(k) if k.startswith("w:") else k, v)
        if txt: el.text = txt
        r._r.append(el)
    p.add_run(f" {text}")
    return p
```

Caption text is a noun phrase in sentence case ("Site inventory by
workload"), no trailing period. The template's own captions read
"Figure 1 Company Info" with a space, no colon; keep that pattern. Set
`updateFields` in `settings.xml` (the base does) so Word refreshes the
numbers on open.

### 10.3 Table style selection

| Use | Style ID | Why |
|---|---|---|
| Default data table, short | `NetwovenTable1` | Full `#262626` grid, `#145CA4` header |
| Long table, 8 rows or more | `NetwovenTable2` | Banded rows aid scanning |
| Small inline table, quiet look | `NetwovenTable3` | Borderless, `#145CA4` caps header (Calibri via theme minor font) |
| SOW fee and rate tables | `TableGrid`, as shipped | Their `=SUM(ABOVE)` and `=PRODUCT(LEFT)` fields and bookmarks depend on the existing rows; never restyle them |
| SOW deliverables | `PSOGrid`, as shipped | |

Header row text is whatever the style dictates; never bold or recolour cells
by hand in a Netwoven-styled table. Right-align numeric columns. Repeat the
header row on page breaks (`trPr/tblHeader`) for tables over one page.

### 10.4 matplotlib recipe (document palette, 200 dpi)

```python
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

PALETTE_DOC = ["#145CA4", "#B30738", "#32355F", "#0080B0", "#93E2FF", "#7F7F7F"]
CHROME = "#595959"

def bar_png(categories, values, path, highlight=None):
    plt.rcParams.update({"font.family": ["Segoe UI", "Selawik", "DejaVu Sans"], "font.size": 9})
    fig, ax = plt.subplots(figsize=(6.5, 3.6), dpi=200)
    colors = [PALETTE_DOC[0] if highlight is None or i == highlight else "#D0CECE" for i in range(len(values))]
    ax.barh(categories, values, color=colors)
    ax.invert_yaxis()
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color("#D0CECE")
    ax.tick_params(colors=CHROME, length=0)
    ax.xaxis.grid(True, color="#D0CECE", linewidth=0.6); ax.set_axisbelow(True)
    for i, v in enumerate(values):
        ax.text(v, i, f" {v:,}", va="center", color=CHROME, fontsize=9)
    fig.tight_layout()
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
```

Insert with `doc.add_picture(path, width=Inches(6.5))`, centre the paragraph,
then `add_caption(doc, "Figure", "...")`. Alt text in Word is set on
`wp:docPr` (`descr` attribute) of the inline drawing:

```python
pic = doc.add_picture(path, width=Inches(6.5))
docPr = pic._inline.docPr
docPr.set("descr", "Bar chart: sites by workload, SharePoint 1,240, Teams 860, Exchange 410")
```

Multi-series document charts use `PALETTE_DOC` in order; the document
palette is marked "to be confirmed by render" in `brand-tokens.md` until the
user's render pass.

---

## 11. Do-not list

- No 3-D charts, no shadows, no gradients, no bevels.
- No chart titles on slides.
- No dual axes.
- No pie charts with more than five slices, and never two pies for
  comparison.
- No theme accent cycle for series colours.
- No white text on `#00B0F0`.
- No text below 12pt anywhere on a slide, below 9pt in Word.
- No new SVG pictures without a PNG fallback.
- No edits to layout or master XML to fix a one-off slide; draw on the slide.
- No manual formatting of Netwoven-styled Word tables.
- No dashes in ranges or labels ("3 to 5", not "3-5" in prose; axis tick
  labels may use a hyphen only when the source data does).
- No colour as the sole carrier of meaning.
- No decorative photos or icons on evidence slides.
- No screenshots without a crop and a callout.
- No visual without a source line.
