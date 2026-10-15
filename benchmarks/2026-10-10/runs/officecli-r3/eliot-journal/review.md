# Review: eliot-journal (authoring bucket), OfficeCLI run r3

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes. Three page PNGs were produced.
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes

## Rubric check
| # | Check | Result |
|---|---|---|
| 1 | journal.docx exists | pass |
| 2 | Two-column poem section | pass: s1 `<w:cols w:num="2" w:space="567"/>`, two columns in the render |
| 3 | Poem titles are headings | pass: four Heading 1 paragraphs (Preludes, Hollow Men, Prufrock, Waste Land) |
| 4 | Verse line breaks preserved | pass: `w:br` per line inside stanza paragraphs, and read shows one line per verse line. Text matches the source verbatim, stanza breaks included |
| 5 | Frontispiece with caption | pass: img0 (3.54x2.76in, alt text) and an italic "Frontispiece" caption on the title page |
| 6 | At least two footnotes | pass: three (fn1–fn3), each with real editorial content |
| 7 | Colophon hyperlink | pass: link0 "public domain" goes to https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon single-column | pass: s0 and s2 have no cols |
| 9 | Landscape | pass: every sectPr is 16838x11906 with orient=landscape (A4) |
| 10 | Header and footer | pass: hdr0–2 "A T. S. Eliot Reader". ftr0–2 use live PAGE and NUMPAGES fields, centered at 9pt, and the render shows "Page 1/2/3 of 3" |
| 11 | Garamond throughout | pass: docx:base font="Garamond". Title and Heading1 rFonts are explicit Garamond with no theme attributes, and the render shows serif type everywhere |
| 12 | Clean multi-section render | pass |

## Merits
- Clean three-section structure, all landscape, with the columns only in the poem section.
- The verse is faithful to the source, and stanza breaks are kept (Preludes' last line, Prufrock's closing couplet).
- The footnotes are real, apt glosses, not placeholders.
- The agent caught the trap where theme fonts override an explicit font, so the headings really do render in Garamond.
- Every section has its own header and footer, so the page chrome appears on every page.
- The page-number fields are live, and "of 3" is correct in the render.

## Demerits
- **Minor: body footnote marks are not superscript.** The reference runs use `rStyle FootnoteReference`, but styles.xml defines only Normal, Title and Heading1. The marks print at full size and run into the text: "steams and stamps.1", "make our visit.2", "dried tubers.3" on page 2.
- **Minor: Heading 1 keeps Word's default blue (2F5496).** The serif face is right, but the color looks like a word processor, not a printed anthology.

Not counted against OfficeCLI: docx-cli's `footers list` and `read` show the cached "Page 1 of 1" instead of `{page}`/`{pages}` tokens for these complex fields. That is a docx-cli reader gap. `docx validate` reports only w14:paraId and w14:textId attributes, which are mc:Ignorable noise.

## Frictions
The agent reported several issues with OfficeCLI:
- **Style font kept theme attributes.** `set /styles/Heading1 --prop font=` left the theme attributes in place, so the readback lied. The agent fixed it by replacing rFonts with a raw set. (major)
- **Header/footer on `/` covered only the last sectPr.** Earlier sections got no header or footer, with no warning. The agent fixed it by adding them to each section. (major)
- **The `--grid auto` contact sheet was misleading.** It showed a blank poem page and a sans-serif fallback. (major)
- **`\n` made new paragraphs.** The skill doc says it gives a line break. The agent found that `\v` works, written `\\v` in JSON. (minor)
- **Footer size applied only to the first run.** `size=9pt` sized only the first run. (minor)
- **`view text` drops `w:br`.** Line breaks disappear from the text view. (minor)
- **Screenshot `-o` path.** A relative path resolved against the wrong directory. (minor)
- **Section semantics are unclear.** Which section an `add --type section` break closes is confusing. (minor)
- **Noisy `view issues`.** It flags verse paragraphs and section-break paragraphs. (minor)

The path was a slog: at least three silent traps (theme fonts, header scope, line-break escape) would each have produced a wrong document if the agent had not verified the XML and the per-page renders. The output is correct anyway.
