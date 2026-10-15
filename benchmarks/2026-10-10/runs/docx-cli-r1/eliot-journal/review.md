# Review — eliot-journal (Authoring: columns, verse, footnotes, links, figure)

**Bucket:** authoring (from scratch, no baseline)

## Verdict
- Task success: **success**
- Renders correctly (Word opened it): **true** (3 output pages rendered)
- Layout sound: **true**
- Formatting preserved: **preserved**
- Survived read loop: **true**

## Rubric check
| # | Criterion | Result | Evidence |
|---|---|---|---|
| 1 | journal.docx exists | PASS | file present; `validate` → "valid — 9 parts checked" |
| 2 | Two-column poem section | PASS | AST s1 `columns: 2`; page 2 shows verse in two columns |
| 3 | Poem titles are headings | PASS | p6/p9/p11/p14 style `Heading2` (Preludes (I), The Hollow Men, Prufrock, The Waste Land) |
| 4 | Verse line breaks preserved | PASS (minor note) | each line renders on its own line; implemented as `<w:br/>` hard breaks inside a stanza paragraph rather than one paragraph per line — read shows separate lines, not prose |
| 5 | Frontispiece + caption | PASS | `images list` → img0 (svg, 3.5x2.72in); p4 "Frontispiece" in Caption style, centered, title section |
| 6 | ≥2 footnotes | PASS | fn1 (Prufrock "etherized"), fn2 (Waste Land opening), both non-empty |
| 7 | Hyperlink in colophon | PASS | link0 "public domain" → https://www.gutenberg.org/ebooks/1321 in p17 |
| 8 | Title + colophon single column | PASS | s0 and s2 have no `columns` |
| 9 | Landscape | PASS | all sections `pageOrientation: landscape`, 15840x12240 |
| 10 | Header + footer page numbers | PASS | hdr0-2 "A T. S. Eliot Reader"; ftr0-2 "Page {page} of {pages}"; render shows header top-left and centered "Page N of 3" |
| 11 | Garamond throughout | PASS | `<!-- docx:base font="Garamond" -->`; render shows serif face in body AND headings |
| 12 | Renders cleanly | PASS | Word rendered 3 pages; section transitions visible |

## Merits
- Every rubric item is satisfied and survives `docx read` / `--ast` / list verbs.
- Clean title page: centered title, italic subtitle, editor's note, frontispiece sized to 3.5in, centered caption.
- Two-column verse page in Word, with the footnotes at the foot of their columns; the stanza breaks ("And then the lighting of the lamps.") are kept.
- Garamond reaches the theme-following headings too (`set-default-font` worked); headings are serif in the render.
- Header, footer, landscape and the hyperlink all use first-class structures (not body text). Fields render as "Page 1 of 3".
- `validate` is clean.

## Demerits
- **minor**: Verse lines are soft line breaks inside a stanza paragraph, not one paragraph per line as the rubric literally says. The visual result and the read view are correct, so this passes in substance.
- **minor**: The first section is typed `continuous` (s0 sectionType). This is harmless for the first section but a bit odd. Empty section-carrier paragraphs p5/p16 exist but are invisible.
- **minor (cosmetic)**: The colophon sits alone on page 3, with the page mostly blank. The task asks for a separate section, so this is acceptable. Headings keep the default blue Heading color, which is less "printed anthology" than black, but it is not a requirement.

## Frictions
- Agent reports: the verse hard-break syntax (trailing backslash) and footnote syntax `[^n]` are undocumented in `create --help` and were discovered by probing. Markdown cannot express centering, so an extra batched `edit --alignment center` pass was needed, and the caption did not inherit the image's alignment. `styles --at` does not show the font face, so set-default-font could not be verified without rendering. Reversed arg order (`styles --at X FILE`) dumped the full help into the error hint. The `docx:p pN align=…` note replaces the bare locator, which confused it.
- Read: the result is correct, but the path was a moderate slog of discovery for verse breaks and footnotes. The help text gap for the trailing-backslash hard break is the most important UX fix, because a weaker run would flatten the poems into prose.
