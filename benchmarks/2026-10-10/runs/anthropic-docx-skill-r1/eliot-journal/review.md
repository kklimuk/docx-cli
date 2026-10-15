# Review: eliot-journal (authoring bucket), competitor arm: Anthropic docx skill, r1

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes. Three page PNGs were produced.
- **Layout sound:** yes
- **Formatting preserved:** n/a. The document was authored from scratch, and all requested formatting is present.
- **Survived read loop:** yes. Every intended element is in the render, and docx-cli reads all of them back.

## Rubric check
| # | Criterion | Result | Evidence |
|---|---|---|---|
| 1 | File exists | PASS | `journal.docx` is present |
| 2 | Two-column poem section | PASS | s1 has `<w:cols w:num="2" w:sep="true">`. Page 2 shows two columns with a separator rule |
| 3 | Poem titles are headings | PASS | Preludes (I), The Hollow Men, Prufrock (excerpt) and Waste Land are each a Heading 2 (`##` in the read view) |
| 4 | Verse line breaks preserved | PASS | Each line is its own paragraph (p6–p54), with empty paragraphs as stanza gaps. The text matches `assets/eliot-poems.md` verbatim |
| 5 | Frontispiece and caption | PASS | `images list` returns img0 (PNG fallback of the SVG, 3.75x2.92in, with alt text). The italic "Frontispiece" caption sits below it on page 1 |
| 6 | At least 2 footnotes | PASS | 3 footnotes: fn1 Hollow Men, fn2 Prufrock "etherized", fn3 Waste Land/Chaucer. All have content and render at the page foot |
| 7 | Colophon hyperlink | PASS | link0 "public domain" points to https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon single-column | PASS | s0 and s2 have `w:num="1"` |
| 9 | Landscape | PASS | All three sectPr have `w:orient="landscape"`. The read view shows `orientation="landscape"`, and the renders are wide pages |
| 10 | Running header and footer | PASS (render) | The header reads "A T. S. Eliot Reader" on every page (hdr0–hdr2). The footer renders centered "Page 1 of 3", "Page 2 of 3", "Page 3 of 3" from real PAGE/NUMPAGES complex fields. docx-cli's `footers list` shows `"Page  of"`, not `{page}`/`{pages}` tokens. That is a docx-cli reader gap (see below), not a competitor defect |
| 11 | Garamond throughout | PASS | `<!-- docx:base font="Garamond" size="11pt" -->`. The render shows a serif face in the body, the headings, the header and the footer |
| 12 | Multi-section renders cleanly | PASS | Title page → two-column body (next page) → colophon (next page). No blank pages and no overflow |

## Merits
- All 12 rubric items are met in the Word render. All four poems are included, Waste Land too.
- Verse matches the source line for line, with stanza breaks kept as empty paragraphs.
- The SVG is embedded with a PNG fallback and descriptive alt text, and the caption sits right beneath it.
- There are three useful editor's footnotes, one more than the two required. They are placed on the relevant lines and render at the foot of the column.
- Headers and footers are repeated on each section. The page-count fields are real (`PAGE`/`NUMPAGES`) and Word resolves them correctly.
- The document default is set to Garamond in styles.xml, and the headings use it too.
- The look is clean and anthology-like: a centered title and subtitle, an italic header, and a column separator rule.

## Demerits
- None of note against the rubric.
- (minor, cosmetic) The colophon page is mostly white space and carries a "Colophon" heading the request didn't ask for. This is a reasonable design choice, not a defect.

## Tool observation (docx-cli, not the competitor)
- docx-cli's reader shows the footer as `text="Page  of"`. It drops the PAGE/NUMPAGES complex fields, which here have an empty cached result (the `separate` fldChar is followed directly by `end`). docx-cli should surface these as `{page}`/`{pages}` tokens whatever the cached value. This is a reader gap worth fixing.

## Frictions
The agent reported that the skill doesn't cover several features this request needs: SVG embedding (it needs a PNG fallback, which the agent made itself with `rsvg-convert`), multi-section documents with per-section columns and orientation, footnotes, hyperlinks, PAGE fields, and verse line-break conventions. It also said the soffice render substitutes fonts, so it could not confirm Garamond visually. It filled every gap from prior knowledge of the docx-js API, and the result was correct on the first build. The path was smooth in outcome but relied on knowledge outside the skill. The skill's coverage gaps (SVG fallback, sections and columns) are real UX demerits that a weaker or less-informed agent might not get past.
