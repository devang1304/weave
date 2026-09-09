---
title: Netwoven PowerPoint Template 2026 (AI Reference Guide)
covers: NW_presentation_template_2026.potx, NW_presentation_template_2026_v2.potx, NW_Presentation_Base_2026.pptx
last_verified: 2026-09-02
---

# Netwoven PowerPoint template 2026

This is the machine-readable specification for Netwoven's 2026 deck template
("Netwoven Default Theme"). It exists so an AI can build a correctly branded
deck without re-deriving the OOXML. Every layout name, placeholder index,
size, colour and geometry below was extracted from the template XML with a
script; the catalog in §3 is exact. Nothing has been rendered in this build,
so layout marks read "proven (to be confirmed by render)" rather than
"proven".

**Files this guide describes:**

| File | Purpose |
|---|---|
| `NW_presentation_template_2026.potx` | Netwoven's original, unmodified |
| `NW_presentation_template_2026_v2.potx` | Cleaned: dead metadata removed, dangling section list removed, one duplicate picture removed, `preserve="1"` added to the last layout. No layout, master, theme, or slide content change (§11) |
| `NW_Presentation_Base_2026.pptx` | `_v2` with the content type flipped so `python-pptx` opens it. The scripts use this file (`assets/bases/NW_Presentation_Base_2026.pptx`) |

Related references: `brand-tokens.md` (colours, palettes, contrast),
`deck-storyline.md` (which layout for which slide), `visual-standards.md`
(charts, frameworks, stat tiles).

---

## 1. The one thing to understand first

**This is not a revision of the 2021 deck. It is a third-party slide chassis
re-skinned for Netwoven.** Shape names inside the layouts are Russian and
German (`Полилиния 20`, `Скругленный прямоугольник 6`, `Текст 3`,
`Рисунок 12`, `Заголовок 1`, `Inhaltsplatzhalter 4`), the placeholder
indexes are non-standard (`1, 2, 10` to `27, 32` to `38`), and only 4 of the
18 media files (the wordmark and three footer glyphs) are shared with the old
template. Everything the old guide said about layout names, `1_Title Only`,
the `Confidentiality Message` layout, the `#F0F2F9` background, Arial, and
`accent1=#4472C4` no longer applies.

Two consequences drive everything else:

1. **The theme is live and Netwoven's.** `clrScheme "Netwoven Default
   Theme"` sets `accent1=#00B0F0`, and the layouts also hard-code `#00B0F0`
   17 times. The master background is white (`bgRef 1001 → bg1`). The font
   scheme ("Custom 1") is Segoe UI for both major and minor. The deck keeps
   `#00B0F0` exactly as designed; it is not the Word chrome blue and is not
   "corrected".
2. **Dispatch on placeholder type and idx from the layout, never on
   position.** `Title Slide` (L1) has **no title placeholder**;
   `placeholders[0]` on it raises. The usable cover is `Title Slide for
   Verticals` (L3). Use `ph(slide, idx)` from `nw_pptx_helpers.py`.

---

## 2. Theme and master text styles

| Slot | Value | Live use |
|---|---|---|
| `dk1` / `lt1` | windowText / window (black / white) | Titles, background |
| `dk2` | `#44546A` | Agenda body text (`tx2`), Big Statement panel tint (lumMod 20 lumOff 80 → `#D6DCE5`) |
| `lt2` | `#E7E6E6` | Footer dividers (lumMod 90 → `#D0CECE`) |
| `accent1` | `#00B0F0` | Cover panel, cover title, stat numerals, captions, Testimonial On Right bar, Closing band |
| `accent2` | `#7030A0` | Testimonial On Left bar, Six Icon "ICON" prompt text |
| `accent3` | `#7F7F7F` | none in layouts |
| `accent4` | `#FFC000` | gradient stop on the instructional slide 1 only |
| `accent5` | `#375623` | none |
| `accent6` | `#70AD47` | none |
| `hlink` / `folHlink` | `#0563C1` / `#954F72` | hyperlinks |
| Fonts | major Segoe UI, minor Segoe UI | Titles `+mj-lt`, body `+mn-lt`; 58 direct `Segoe UI` references in layouts; 20 dead `Calibri Light` references in Closing Slide list styles |
| Format scheme | Office (stock) | |
| Slide size | `12192000 x 6858000` EMU, 13.333in by 7.5in, 16:9 | |

Master `txStyles` (fully resolved):

| Style | Value |
|---|---|
| Title | 40pt, `tx1` (black), `+mj-lt`, left, line spacing 90 percent, no bullet |
| Body level 1 | 20pt, `tx1` lumMod 65 lumOff 35 (`#595959`), bullet `•` in Arial, `marL 228600` `indent -228600` (0.25in hanging), space before 10pt |
| Body levels 2 to 5 | 18 / 16 / 14 / 14pt, same colour and bullet, indents 0.75 / 1.25 / 1.75 / 2.25in, space before 5pt |
| Body levels 6 to 9 | 18pt black (unused) |
| Other | 18pt black, `+mn-lt` |

Master placeholders: title at x 0.92in, y 0.40in, w 11.5in, h 1.11in
(`838200, 365126`); body idx 1 at x 0.92in, y 1.65in, w 11.5in, h 5.11in
(`838200, 1504950`). Layouts that show "(inherit)" in §3 use these.

---

## 3. Layout catalog (all 39, in master order)

Placeholder columns give `type idx` and any override of size, colour, or
font from the layout's `lstStyle` level 1. Geometry is inches from the
layout's `xfrm` (blank = inherited from the master). "Master footer" says
whether the master's footer shapes show (`showMasterSp` is unset on every
layout except L19, L20 and L38, which set `0`). Mark: **proven** means the
builder uses it (to be confirmed by render); **usable** means structurally
sound; **avoid** with the reason.

| # | Layout name (exact) | Placeholders (type idx: notes) | Master footer | Mark |
|---|---|---|---|---|
| L1 | `Title Slide` | pic 15, 16, 17, 18, 19 (five tilted "Image Holder" tiles, 12pt `#808080` prompt); subTitle 1 (18pt `#595959`, x 7.27 y 4.95 w 4.40 h 1.38). **No title placeholder.** Static cyan panel `#00B0F0` at x 11.43 w 5.55 (runs off the right edge); logo group with SVG badge | yes | **avoid**: no title; `BROKEN_LAYOUTS` in the helpers |
| L2 | `Title and Content` | title 0 (40pt, inherit); body 1 (20/18/16/14/14pt, inherit) | yes | **proven**: executive summary, next steps, Confidentiality slide |
| L3 | `Title Slide for Verticals` | ctrTitle 0 (48pt `#00B0F0`, anchor bottom, x 6.51 y 1.23 w 5.44 h 3.80); subTitle 1 (18pt `#595959`, x 6.51 y 5.17 w 5.44 h 1.38); pic 15 (x 0.68 y 3.27 w 2.35 h 3.52), pic 16 (x 3.59 y 1.90 w 2.35 h 3.45). Static cyan panel x 11.43 w 5.44; logo group (`image7.png` + SVG badge, `#375991` fill) at x 0.72 y 0.87 | yes | **proven**: the cover |
| L4 | `Agenda` | body 21 (18pt, not bold, `tx2` `#44546A`, `+mj-lt`, x 0.88 y 1.51 w 5.67 h 4.87); pic 38 (x 7.15 y 0.92 w 4.67 h 4.67, pattern fill). Static text box "Agenda" 40pt at x 0.88 y 0.55; two `#88B9BE` 3pt rounded-rectangle outlines; cyan panel | yes | **usable**: no title placeholder, the word "Agenda" is baked in |
| L5 | `Company Profile` | pic 25 (x 0 y 0 w 13.33 h 4.05, 14.67pt prompt). Static "Netwoven Inc." 28pt and a 13pt boilerplate paragraph ("Netwoven is a trusted Microsoft Solutions Partner who unravels complex business…") | yes | **avoid**: baked marketing text cannot be edited through placeholders |
| L6 | `Photo on Left` | pic 33 (x -1.34 y 0 w 5.53 h 7.5); title 0 (40pt, x 4.72 y 0.40 w 7.70 h 1.15); body 1 (18pt, `buNone`, x 4.72 y 1.71 w 7.70 h 5.05) | yes | usable |
| L7 | `2 Photo on right` | pic 10 (x 8.18 y 1.13 w 4.11 h 2.22), pic 17 (x 8.18 y 4.14); title 0 (x 0.92 y 0.40 w 6.89 h 1.11); body 21 and 24 (20pt bold `#00B0F0` Segoe UI, labels); body 22, 23, 25, 26 (16pt `#404040`, items, 3.49 / 3.40 wide) | yes | usable: two labelled photo call-outs |
| L8 | `Section Header` | title 0 (60pt, anchor bottom, x 0.91 y 1.87 w 11.5 h 3.12); body 1 (24pt `tx1` tint 75, `buNone`, x 0.91 y 5.02 w 11.5 h 1.64); pic 32 (x 9.75 y -0.01 w 3.58 h 3.58) | yes | **proven**: section dividers, "Appendix" |
| L9 | `Two Content` | title 0 (inherit); body 1 (half, x 0.92 y 1.71 w 5.67 h 5.05); body 2 (half, x 6.75 y 1.71 w 5.67 h 5.05) | yes | **proven** |
| L10 | `Three Column` | title 0 (40pt, inherit); body 13, 14, 15 (20/18/16/14/14pt, x 0.92 / 4.89 / 8.85, y 1.74, w 3.56 h 4.98) | yes | **proven** |
| L11 | `Comparison` | title 0 (x 0.92 y 0.40 w 11.5 h 1.06); body 1 and 3 (24pt bold `buNone`, anchor bottom, y 1.48 h 0.90, x 0.92 / 6.75); body 2 (half) and 4 (x 0.92 / 6.75, y 2.41, w 5.64 / 5.67, h 4.36) | yes | **proven** |
| L12 | `Title Only` | title 0 (inherit) | yes | **proven**: charts, tables, frameworks. Content area x 0.92 y 1.6 w 11.5 h 5.0 (to be confirmed by render) |
| L13 | `Title Only (Centered)` | title 0 (inherit geometry, `algn=ctr`) | yes | usable |
| L14 | `Five Part Content` | pic 10 to 14 (x 0.51 / 3.02 / 5.53 / 8.03 / 10.54, y 1.37, w 2.31 h 3.68); title 0 (x 0.92 y 0.40 w 11.5 h 0.76); body 1, 16, 18, 20, 22 (16pt bold `#32355F` Segoe UI, `buNone`, category labels, y ~5.13); body 15, 17, 19, 21, 23 (10pt `#000000`, `buNone`, y ~5.90) | yes | usable, but the 10pt content text is below the 12pt floor; use only for labels |
| L15 | `Text with Right Bottom Photo` | pic 35 (x 7.41 y 0.79 w 5.92 h 5.92); title 0 (40pt Segoe UI, x 0.88 y 0.33 w 4.61 h 0.89); body 21 (14pt `tx2`, x 0.88 y 1.47 w 6.54 h 4.60) | yes | usable |
| L16 | `Text with Right Top Photo` | pic 32 (x 7.41 y -0.01); title 0 and body 21 as L15 | yes | usable |
| L17 | `Text with Left Bottom Photo` | pic 35 (x 0 y 0.79); title 0 (x 5.92 y 0.40 w 4.61); body 21 (x 5.92 y 1.54 w 6.54 h 4.60) | yes | usable |
| L18 | `Text with Left Top Photo` | pic 35 (x 0 y -0.01); title 0 and body 21 as L17 | yes | usable |
| L19 | `Testimonial On Right` | pic 32 (x 0 y 0 w 6.67 h 7.5 photo); pic 33 (x 5.67 y 5.12 w 1.99 h 1.92); pic 34 (customer logo, x 7.67 y 0.34 w 2.08 h 0.93); title 0 (24pt bold, x 7.66 y 1.33 w 4.84 h 0.93); body 21 (16pt `tx2`, quote, x 7.66 y 2.39 w 4.84 h 2.79); body 35 (24pt bold `bg1`, name, on a `#00B0F0` bar x 6.99 y 5.47 w 5.23 h 1.20); body 36 (16pt `#E5F2F3`, position). Two SVG quote glyphs without fallback | **no** | **proven**: the quote layout |
| L20 | `Testimonial On Left` | mirror of L19; bar fill is **`accent2` purple** at x 0.93 y 5.37 | **no** | avoid: purple bar; use L19 |
| L21 | `Blank` | none | yes | **proven**: the template's Thank You slide sits on it |
| L22 | `Two Comparison` | title 0 (x 0.50 y 0.46 w 12.33 h 0.57); body 14, 15 (16pt bold, `+mj-lt`, `buNone`, anchor centre, y 1.31 h 0.67); body 12, 13 (x 0.51 / 6.81, y 2.18, w 6.03 h 4.49) | yes | usable |
| L23 | `Four Content` | title 0 (x 0.51 y 0.36 w 12.33 h 0.57); body 14 to 17 (x 0.51 / 3.66 / 6.81 / 9.96, y 1.12, w 2.88 h 5.59) | yes | usable |
| L24 | `Five Content` | title 0; body 14 to 18 (w 2.20, x 0.51 / 3.05 / 5.58 / 8.11 / 10.64, y 1.14, h 5.58) | yes | usable; columns are narrow |
| L25 | `Big Statement with Illustration` | pic 13 (x 0 y 0 w 3.86 h 6.71, over a `#D6DCE5` panel); body **16** (statement, 72/48/36pt, `+mj-lt`, bullets are U+200B, x 4.71 y 1.50 w 8.11 h 5.21, anchor top); body **14** (eyebrow, 21pt `#0055A0`, `buNone`, anchor bottom, x 4.71 y 0.71 w 8.11 h 0.44). No title placeholder | yes | **proven**: one big number or sentence. Note the idx assignment: statement is 16, eyebrow is 14 |
| L26 | `Three Stat` | body 10, 15, 17 (numerals, 96pt `#00B0F0`, `+mj-lt`, spc -300, anchor bottom, y 2.18 h 1.82, x 0.53 / 4.72 / 8.91, w 3.95); body 11, 14, 16 (labels, 21pt `tx1`, `+mn-lt`, `buNone`, anchor top, y 4.47 h 0.67). No title placeholder | yes | **proven** |
| L27 | `Three Image Horizontal` | title 0 (x 0.51 y 0.36 w 12.33 h 0.57); pic 16, 17, 18 (x 1.33 / 5.51 / 9.69, y 1.54, 2.33 square); body 13, 14, 15 (level 1 20pt bold `#00B0F0` `buNone`; levels 2+ 12pt with U+200B bullets; x 0.51 / 4.70 / 8.88, y 4.21, w 3.96 h 1.96) | yes | usable |
| L28 | `Six Icon` | title 0 (**left column**, x 0.51 y 3.01 w 3.94 h 0.57); body 13, 14, 15 (row 1, y 1.25) and 16, 17, 18 (row 2, y 4.50), x 5.75 / 8.36 / 11.00, w 1.83 h 2.00, level 1 16pt bold `#32355F`, levels 2+ 12pt U+200B; pic 19, 20, 21 (y 0.29) and 25, 26, 27 (y 3.55), 0.66 square, "ICON" prompt 9pt `accent2` | yes | **proven**: four to six parallel points |
| L29 | `Grid Layout 1-2` | title 0 (x 0.51 y 0.36 w 12.33 h 0.57); eyebrows 14, 15, 17 (10.5pt bold all caps spc 70, `+mn-lt`, `buNone`, anchor centre, h 0.33); body 12 (x 0.51 y 1.72 w 6.03 h 4.73), 13 (x 6.81 y 1.72 h 1.97), 16 (x 6.81 y 4.48 h 1.97); body 18/16/14pt | yes | usable |
| L30 | `Grid Layout 1-2-2` | title 0; eyebrows 14, 15, 17, 19, 21; bodies 12 (x 0.51 w 3.95 h 4.73), 13, 16 (x 4.70), 18, 20 (x 8.89), h 1.97 | yes | usable |
| L31 | `Grid Layout 2-1` | title 0; eyebrows 14, 15, 17; bodies 12 (x 0.51 y 1.70 h 1.90), 13 (y 4.47 h 1.96), 16 (x 6.81 h 4.73) | yes | usable |
| L32 | `Grid Layout 2-2` | title 0; eyebrows 14, 15, 18, 19; bodies 12, 13 (x 0.51), 16, 17 (x 6.81), w 6.03, h 1.90 / 1.96 | yes | **proven**: four points |
| L33 | `Grid Layout 2-2-2` | title 0; eyebrows 14, 15, 18, 19, 22, 23; bodies 12, 13 (x 0.51), 16, 17 (x 4.70), 20, 21 (x 8.89), w 3.95 | yes | **proven**: six points |
| L34 | `Content with Caption` | title 0 (32pt, anchor bottom, x 0.92 y 0.50 w 4.30 h 1.75); body 1 (32/28/24/20pt, x 5.67 y 1.08 w 6.75 h 5.33); body 2 (16pt `buNone`, x 0.92 y 2.25 w 4.30 h 4.17) | yes | avoid: 32pt body |
| L35 | `Picture with Caption` | title 0 (32pt); pic 1 (x 5.67 y 1.08 w 6.75 h 5.33); body 2 (16pt `buNone`) | yes | usable for one annotated screenshot |
| L36 | `Title and Vertical Text` | title 0; body 1 vertical | yes | avoid |
| L37 | `Vertical Title and Text` | vertical title 0 (x 9.54 w 2.88 h 6.36); body 1 (x 0.92 w 8.46) | yes | avoid |
| L38 | `Closing Slide` | pic 25 (x 0 y 0 w 13.33 h 5.08). Static: `#00B0F0` band x 0 y 5.08 w 8.33 h 2.42; "Netwoven Inc." 28pt; "4000 Pimlico Drive Suite 114-103 Pleasanton, CA 94588, United States" 16pt white; "+1 877 638 9683", "info@netwoven.com", "netwoven.com" 14.67pt with cyan glyph freeforms; SVG map pin without fallback; dead `Calibri Light` list styles | **no** | **proven**: the template's contact closer |
| L39 | `Only Image` | pic 32 (x 0 y 0 w 13.33 h 7.5) | yes | usable; lacked `preserve="1"` in the original (added in `_v2`) |

Picture placeholder prompt text ("Image Holder", "Click icon to add
picture", "Insert customer logo", "ICON") is layout text; an unfilled
picture placeholder renders that prompt, so the builder removes empty
placeholders (`finalize --prune-empty`) and the validator fails on "Image
Holder" or "Click to edit" surviving in a slide.

---

## 4. Master and footer mechanics

The master (`slideMaster1.xml`) carries the footer as **static shapes**;
there is no `p:hf` element and no footer, date, or slide-number placeholder.
It inherits onto every layout that does not set `showMasterSp="0"` (all
except L19, L20, L38), including the two cover layouts (to be confirmed by
render: the cover's cyan panel overlaps the slide-number box at the right).

| Shape (id) | Geometry (in) | Content |
|---|---|---|
| `Picture 14` (15) | x 1.03 y 7.00 w 1.39 h 0.20 | Wordmark `image5.png` (236 x 34, `#00549F` / `#B30738`) |
| `Group 20` (21) with `Picture 4` (5) and `Picture 18` (19) | x 2.95 y 6.96 w 1.47 h 0.32 | **"25 YEARS OF INNOVATION" badge `image6.png` (930 x 200), placed twice, identical**; `_v2` removes the duplicate id 5 and keeps id 19. The badge is kept exactly as designed (user decision) |
| `Straight Connector 6` (7), `Straight Connector 16` (17) | x 4.87 and 6.41, y 6.90, h 0.43 | Vertical dividers, `bg2` lumMod 90 (`#D0CECE`) |
| `Picture 13` (14) | x 5.17 y 6.99 w 1.03 h 0.34 | Microsoft Solutions Partner badge `image4.png` (166 x 55) |
| `Picture 9` (10), `Rectangle 10` (11) | x 6.60 / 6.87, y 6.99 | Globe glyph `image3.png` (512 x 512), "netwoven.com" 11pt Segoe UI `#595959` |
| `Picture 7` (8), `Rectangle 11` (12) | x 8.18 / 8.56 | Mail glyph `image1.png`, "info@netwoven.com" |
| `Picture 8` (9), `Rectangle 12` (13) | x 10.30 / 10.61 | Phone glyph `image2.png`, "+1 877 638 9683" |
| `TextBox 15` (16) | x 12.45 y 6.96 w 0.54 h 0.29 | `slidenum` field, 11pt `#595959` |

The footer text runs carry a vestigial `a:ea typeface="Roboto"` (a leftover
from the 2021 template's Roboto footer); `_v2` removes it, along with
`Tahoma` `ea` hints on layouts 4 and 15 to 20 and slide 5. Latin text is
Segoe UI throughout.

If Netwoven's phone number or address changes, the master (footer) and
L38 (Closing Slide) must both be edited; nothing is data-bound. `core.xml`
`dc:title` is empty and not bound anywhere; the builder sets it.

---

## 5. Tables, charts, notes

- `tableStyles.xml` defines no styles; its `def` GUID
  `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` is PowerPoint's built-in "Medium
  Style 2 - Accent 1", which would put white text on saturated `#00B0F0`
  (2.48:1 contrast, fails AA). The builder's `add_table` therefore writes
  explicit fills: header `#00B0F0` with white bold 12pt Segoe UI (confirm by
  render; switch to `#32355F` if it reads poorly), bands `#FFFFFF` /
  `#F2F2F2`, body 12pt `#595959`.
- No charts, no notes pages with content (a `notesMaster1.xml` exists with
  the stock "Office Theme" as `theme2.xml`), no SmartArt, no embedded
  workbooks. Charts are created by `add_chart` per `brand-tokens.md` §4.6.

---

## 6. Quirks (confirmed by inspection)

1. **`Title Slide` (L1) has no title placeholder.** Five picture tiles and a
   subtitle. Never use it; the validator fails a deck whose first slide is
   on it.
2. **Non-standard placeholder idx values.** Bodies are idx 10 to 27 and 32
   to 38 depending on layout; `Three Stat` pairs are (10, 11), (15, 14),
   (17, 16); `Big Statement` statement is 16 and eyebrow is 14. Dispatch by
   idx from `LAYOUTS`, never by order.
3. **Zero-width-space bullets.** L25, L27 and L28 set `buChar` to U+200B on
   body levels so the bullet is invisible. Leave the layout `pPr` alone; use
   `set_paragraph_no_bullet` on slide paragraphs if a visible bullet
   appears.
4. **19 SVG pictures have no raster fallback** (`a:blip` without `r:embed`,
   only the `asvg:svgBlip` extension): the partner badge in the cover group
   (L1 and L3, `image8.svg`), two quote glyphs each on L19 and L20
   (`image9.svg`), the map pin on L38 (`image10.svg`), and all 12 ribbon
   pictures on the Thank You slide (`image14.svg` to `image17.svg`).
   LibreOffice may drop them and `python-pptx` `Picture.image` raises on
   them. PowerPoint renders them fine. The optional `--svg-fallback` build
   step (needs LibreOffice) adds PNG fallbacks; the default base records
   the limitation in `build-manifest.json`.
5. **`Agenda` (L4) bakes the word "Agenda"** as a static 40pt text box; there
   is no title placeholder. **`Company Profile` (L5) bakes two paragraphs of
   marketing copy.** Neither can be edited through placeholders.
6. **`Testimonial On Left` (L20) uses a purple `accent2` bar** where L19 uses
   cyan. Kept as designed; use L19.
7. **The L25 eyebrow is `#0055A0`**, a one-off near miss of the logo blue.
   Kept as designed.
8. **The 25-year badge is a dated raster** (`image6.png`) and appears twice
   in the master group. `_v2` removes the duplicate picture only; the badge
   stays exactly as designed (user decision).
9. **`Only Image` (L39) lacked `preserve="1"`**; PowerPoint may delete an
   unpreserved, unused layout on save. `_v2` adds the attribute.
10. **`ppt/presentation.xml` has a `p14:sectionLst`** ("Default Section")
    listing the six slide ids; it dangles when slides are deleted. `_v2`
    removes the extension (Track 1 correction 4).
11. **Two cover layouts, one cyan panel each, both show the master footer**
    (no `showMasterSp="0"`). To be confirmed by render.
12. **`Five Part Content` (L14) content text is 10pt**, below the 12pt
    floor.
13. **Dead `Calibri Light` list styles** on the L38 static text boxes (20
    references); kept in `_v2` because the build could not pixel-verify
    their removal. Harmless: the runs set Segoe UI directly.
14. **Metadata.** Four dead `customXml` items (SharePoint), a
    `docProps/custom.xml` with only `ContentTypeId` and an empty
    `MediaServiceImageTags` (no MSIP properties, so the part is removed), a
    `thumbnail.jpeg` of the instructional slide, `dc:creator` and
    `lastModifiedBy` naming an employee, `app.xml` `TitlesOfParts` listing
    slide titles. **`docMetadata/LabelInfo.xml` (label
    `{d3e71191-c083-4948-94ca-f99c3ca5a353}`, method Standard) is kept
    untouched.**
15. **Slide 4's title is a placeholder with `idx="4294967295"`** (an
    orphaned placeholder), which is why generic title lookups miss it. The
    slide is dropped anyway.

---

## 7. The six shipped slides, and the generated closers

| Slide | Layout | Content | Fate |
|---|---|---|---|
| 1 | `Title and Content` | Title "How to Use This Template…", five bullets ("Intended Use…", "Repurpose Content…", "Keep Content Concise…", "Maintain Branding & Formatting…", "Customize as Needed…"), a text box "* Delete this slide from the presentation upon completion", gradient background with an `accent4` stop | **Drop** (matched by `DROP_MARKERS`) |
| 2 | `Title Slide for Verticals` | Empty title and subtitle, two stock photos (`image11.jpg` 285 x 186, `image12.jpg` 200 x 262) in the picture slots | **Drop**; the builder creates its own cover |
| 3 | `Agenda` | Empty agenda body, photo `image13.jpg` (400 x 447) | **Drop** |
| 4 | `Only Image` | "SECTION DIVIDER" in an orphan title placeholder, brown scrim shapes, a text box "Remove this section, if not needed / or use it to add subtitle. To Set a background image…", animation and transition | **Drop**; dividers use `Section Header` |
| 5 | `Blank` | 12 SVG ribbon pictures and a text box "THANK YOU / FOR YOUR / TIME TODAY" (72pt white Segoe UI, not a title placeholder) | **Keep** as the Thank You closer, tagged via `slide.name`; `--no-thankyou` drops it |
| 6 | `Closing Slide` | Picture placeholder 25 filled with `image18.jpg` (1280 x 488) | **Keep** as the contact closer; `--no-closing` drops it |

The template has **no confidentiality slide or layout**. The builder generates
one:

| Property | Value |
|---|---|
| Layout | `Title and Content` (L2) |
| Title (idx 0) | `Statement of Confidentiality` |
| Body (idx 1) | Three paragraphs, 16pt, `buNone` (`set_paragraph_no_bullet`), colour inherited (`#595959`), 10pt space before each |
| Paragraph 1 | `This document contains information that is proprietary and confidential to Netwoven, Inc. and {client} which shall not be disclosed, transmitted, or duplicated, used in whole or in part for any purpose other than its intended purpose. Any use or disclosure in whole or in part of this information without the express written permission of Netwoven, Inc. and {client} are prohibited.` |
| Paragraph 2 | `Any other company and product names mentioned are used for identification purposes only, and may be trademarks of their respective owners.` |
| Paragraph 3 | `©2001 - {year} Netwoven, Inc.  All rights reserved.  Any use or distribution of these materials without express authorization of Netwoven, Inc. and {client} are strictly prohibited.` |
| `{client}` | `--client` value, verbatim |
| `{year}` | Current year at build time |
| Position | Immediately before the Thank You slide; final order is body slides, Confidentiality, Thank You, Closing |
| Skipped when | `--internal` or `--no-confidentiality` |

The wording is the Word template's, copied verbatim including its grammar
("…are prohibited") and the word "document"; do not edit it.

---

## 8. Media

| File | Pixels | Used by | Role |
|---|---|---|---|
| `image1.png`, `image2.png`, `image3.png` | 512 x 512 | master | Mail, phone, globe footer glyphs |
| `image4.png` | 166 x 55 | master | Microsoft Solutions Partner badge |
| `image5.png` | 236 x 34 | master | Wordmark (same bytes as the Word templates' logo) |
| `image6.png` | 930 x 200 | master (twice) | 25 Years of Innovation badge, `#00549F` / `#B30738` / `#3E3E3E` |
| `image7.png` | 2389 x 361 | L1, L3 | Cover logo strip inside the cover group |
| `image8.svg` | | L1, L3 | Cover partner badge (no fallback) |
| `image9.svg` | | L19, L20 | Quote glyph (no fallback) |
| `image10.svg` | | L38 | Map pin (no fallback) |
| `image11.jpg`, `image12.jpg` | 285 x 186, 200 x 262 | slide 2 | Demo cover photos (dropped with the slide) |
| `image13.jpg` | 400 x 447 | slide 3 | Demo agenda photo (dropped) |
| `image14.svg` to `image17.svg` | | slide 5 | Thank You ribbons (kept, no fallback) |
| `image18.jpg` | 1280 x 488 | slide 6 | Closing slide photo (kept) |
| `docProps/thumbnail.jpeg` | | package | Thumbnail of the instructional slide; removed in `_v2` |

---

## 9. Procedure for an AI

1. Write the ghost deck (`deck-storyline.md` §1) from the pyramid before
   touching the file.
2. `create` the deck with the script (§10): cover on `Title Slide for
   Verticals`, closers tagged, Confidentiality generated, demo slides gone.
3. Add body slides with the helpers, choosing layouts from the map in
   `deck-storyline.md` §5. Fill placeholders by idx. Draw charts and
   frameworks on `Title Only` per `visual-standards.md`.
4. Never edit the master, a layout, or the theme for a one-off need. Never
   add SVG-only pictures. Never use `Title Slide`, `Company Profile`, or
   `Testimonial On Left`.
5. `finalize` the deck: closers reordered last, empty placeholders pruned,
   `app.xml` slide count fixed, core title set.
6. Validate (`--type deck [--internal] --client "Client, Inc."`). Render
   (`--render`) when LibreOffice is available; the user runs that later.

---

## 10. Programmatic recipe

```
python scripts/new_deck_pptx.py create assets/bases/NW_Presentation_Base_2026.pptx out.pptx \
  --title "Tenant Assessment Readout" --subtitle "Findings and recommendations" --client "Contoso, Inc." \
  [--cover-image a.jpg[,b.jpg]] [--internal] [--no-confidentiality] [--no-thankyou] [--no-closing]

python scripts/new_deck_pptx.py finalize out.pptx [--title "..."] [--prune-empty]
```

Between the two commands, add body slides:

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from nw_pptx_helpers import (LAYOUTS, ph, slide_text, add_title_content, add_section_header,
                             add_two_content, add_three_stat, add_table, add_chart,
                             set_paragraph_no_bullet, PALETTE_CATEGORICAL)

prs = Presentation("out.pptx")
layouts = {l.name: l for l in prs.slide_masters[0].slide_layouts}   # exact names from §3

# Executive summary
add_title_content(prs, "Contoso can migrate in 14 weeks once site ownership is fixed",
                  ["312 of 1,240 site collections have no owner",
                   "Ownership cleanup unblocks three of four migration waves",
                   "Decision needed: approve the four-week governance sprint"])

# Section divider
add_section_header(prs, "Current state", "What the inventory shows")

# Evidence slide: chart on Title Only, drawn into the content area (confirm by render)
s = prs.slides.add_slide(layouts["Title Only"])
ph(s, 0).text = "Three workloads carry 80 percent of the content"
# add_chart(slide, kind, categories, series, left, top, width, height, ...);
# series is a dict of {series name: ordered values}, colours come from
# PALETTE_CATEGORICAL automatically -- never pass a custom palette here.
add_chart(s, "column", ["SharePoint", "Teams", "Exchange", "Other"], {"TB": [12.4, 3.1, 1.8, 0.7]},
          Inches(0.92), Inches(1.6), Inches(11.5), Inches(5.0))

# Three Stat: numerals idx 10, 15, 17; labels idx 11, 14, 16
add_three_stat(prs, [("1,240", "site collections"), ("312", "without an owner"), ("18 TB", "of content")])

# Big Statement: statement idx 16, eyebrow idx 14 (from the layout XML)
s = prs.slides.add_slide(layouts["Big Statement with Illustration"])
ph(s, 14).text = "GOVERNANCE"
ph(s, 16).text = "One owner per site before wave 1"
for p in ph(s, 16).text_frame.paragraphs:
    set_paragraph_no_bullet(p)

# Next steps table: header 00B0F0 white bold 12pt, bands FFFFFF / F2F2F2
s = prs.slides.add_slide(layouts["Title Only"])
ph(s, 0).text = "Four actions start the governance sprint next week"
add_table(s, ["Action", "Owner", "Date"],
          [["Approve sprint scope", "Contoso IT Director", "2026-09-19"],
           ["Assign site owners", "Contoso M365 team", "2026-10-03"]])

prs.save("out.pptx")
```

Notes for scripts:

- The closers (Thank You, Closing, generated Confidentiality) are tagged in
  `slide.name`; `finalize` moves them to the end in `CLOSER_TAGS` order and
  falls back to text ("THANK YOU", "TIME TODAY") and layout name ("Closing
  Slide") detection if the tags are missing.
- Deleting slides with python-pptx means dropping the relationship and the
  `sldId`; `_v2` already removed the `p14:sectionLst` so nothing dangles.
- `Picture.image` raises on the SVG-only pictures; use `slide_text` (which
  recurses groups and tables) rather than iterating pictures when reading a
  slide.
- The footer needs no code; it inherits from the master on every body
  slide.

---

## 11. Changelog (as built by `tools/build_bases.py`)

All eight gates passed (Stage B whitelist, Stage C single-part diff,
forbidden strings absent in `_v2` and base, LabelInfo present and
byte-identical, builder markers on the base). Rendering gates were not run;
the zero-pixel claim for Stage A is to be confirmed by the user's render
pass. Source sha256 `48fac854…`, `_v2` `c8eaef1c…`, base `5d36037f…`.

Stage A (no visible change intended):

1. `customXml`: all 4 items removed with `itemProps`, `.rels`, overrides
   and presentation rels (`item1` form templates, `item2` a SharePoint
   core-properties scaffold, `item3` content type, `item4` taxonomy sync);
   none was referenced.
2. `docProps/thumbnail.jpeg` removed with its package relationship (the
   `jpeg` default extension kept for the slide photos).
3. `docProps/custom.xml` removed with its override and relationship: it held
   only `ContentTypeId` and an empty `MediaServiceImageTags`, no MSIP
   properties. **`docMetadata/LabelInfo.xml` kept, byte-identical.**
4. `docProps/core.xml`: creator and lastModifiedBy → "Netwoven" (2);
   `dc:title` left empty for the builder.
5. `presentation.xml`: the `p:ext {521415D9-…}` carrying `p14:sectionLst`
   removed.
6. `slideMaster1.xml`: the duplicate badge `p:pic id=5` removed from
   `grpSp id=21` (both pictures used the same `rId46`; their offsets
   differed by 2151 EMU in y); `id=19` on top kept.
7. Vestigial `<a:ea typeface="Roboto"/>` removed from the master (8) and
   `<a:ea typeface="Tahoma"/>` from layouts 4, 15 to 20 and slide 5 (12).
8. `slideLayout39.xml` (`Only Image`): `preserve="1"` added to the root.

Stage B: no changes (SVG PNG fallbacks were not requested; run
`build_bases.py --only ppt --svg-fallback` after installing LibreOffice to
add them; the limitation is recorded in `build-manifest.json`).

Stage C: `[Content_Types].xml` main part flipped
`presentationml.template.main+xml` → `presentationml.presentation.main+xml`;
every other part byte-identical to `_v2`; python-pptx opens the base; the
39 layouts enumerate by name with the placeholder tuples in §3 (a `p:ph`
without a `type` reports as `OBJECT` in python-pptx); slide 1 matches
`DROP_MARKERS`, slide 5 contains "THANK YOU", slide 6 sits on `Closing
Slide`.

Parts diff: Stage A vs original changed 15 parts (`[Content_Types].xml`,
`_rels/.rels`, `docProps/core.xml`, presentation part and rels, layouts 4,
15 to 20 and 39, the master, slide 5) and removed 14; Stage B vs A changed
0; Stage C vs B changed 1.

### Not changed (by decision)

- Purview/MSIP sensitivity label `docMetadata/LabelInfo.xml`
  (`{d3e71191-…}`, method Standard), its relationship and content type.
- The "25 YEARS OF INNOVATION" badge (`image6.png` on the master).
- `#00B0F0` `accent1` in the theme and its 17 hard-coded copies in layouts.
- `Testimonial On Left` (L20) purple `accent2` bar.
- `Big Statement with Illustration` (L25) `#0055A0` eyebrow colour.
- `Closing Slide` (L38) dead `Calibri Light` list styles (20 references):
  kept, because the zero-pixel condition could not be verified without a
  render.
- `ppt/theme/theme1.xml` and every layout and master shape position.
- The 19 SVG pictures without raster fallback (opt-in Stage B only).

### Discrepancy to reconcile with the helpers

The plan's `LAYOUTS` sketch lists `Big Statement with Illustration` as
"statement 14, eyebrow 16". The layout XML and the build report agree on
the opposite: **idx 16 is the 72pt statement (`Text Placeholder 2`), idx 14
is the 21pt `#0055A0` eyebrow (`Text Placeholder 7`)**. This guide and
`deck-storyline.md` follow the XML; `nw_pptx_helpers.LAYOUTS` should too.
