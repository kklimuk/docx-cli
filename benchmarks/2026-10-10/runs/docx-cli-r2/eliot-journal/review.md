# Review: eliot-journal (Authoring bucket: columns, verse, footnotes, links, figure)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (4 page PNGs produced and inspected)
- **Layout sound:** true, with minor blemishes (see Demerits)
- **Formatting preserved:** preserved
- **Survived read loop:** true (a fresh `docx read` matches `renders/output/read.md` byte for byte; `docx validate` reports "valid — 9 parts checked")

## Rubric check (criteria.md)
| # | Check | Result |
|---|---|---|
| 1 | journal.docx exists | PASS |
| 2 | Two-column poem section | PASS: AST `s1` has `columns: 2`, `nextPage`; the render shows two columns on page 2 |
| 3 | Poem titles are headings | PASS: p6, p9, p11 and p14 are `Heading2` (Preludes, Hollow Men, Prufrock, Waste Land) |
| 4 | Verse line breaks preserved | PASS in spirit, deviates from the letter: every verse line is on its own line in the render and the read view, but as in-paragraph line breaks (p7 has 23 runs), not one paragraph per line. Stanza breaks are separate paragraphs. The text matches `assets/eliot-poems.md` exactly |
| 5 | Frontispiece and caption | PASS: `images list` shows img0 (svg, 4x3.11in, alt "Frontispiece"). p4 "Frontispiece" uses the Caption style, centered, on page 1 |
| 6 | At least 2 footnotes | PASS: 3 footnotes (Michelangelo, Chaucer/April, Old Guy/straw), all with content and shown at the foot of page 2 |
| 7 | Colophon hyperlink | PASS: link0 "public domain" goes to https://www.gutenberg.org/ebooks/1321 in p17 |
| 8 | First and last sections single-column | PASS: s0 and s2 have no columns |
| 9 | Landscape | PASS: all 3 sections are 15840x12240 `landscape`; there is a `docx:page orientation="landscape"` note |
| 10 | Header and footer | PASS: hdr0-2 "A T. S. Eliot Reader" and ftr0-2 "Page {page} of {pages}". The render shows the header on every page and a centered "Page N of 4" |
| 11 | Garamond throughout | PASS: `<!-- docx:base font="Garamond" -->`. The render shows a Garamond serif in the body, the headings, the header, the footer and the footnotes |
| 12 | Multi-section render | PASS: Word rendered 4 pages, and the section transitions are visible |

## Merits
- Every rubric item is met. The document validates cleanly and survives the read loop.
- The verse is verbatim from the source, and each verse line is visually distinct. The stanza break in Preludes and in Prufrock is kept.
- The title page is clean: a centered Heading 1 title, an italic centered subtitle, the editor's note, a centered frontispiece, and a centered Caption-style caption.
- The two-column body flows correctly, and Word renumbered the footnotes in reading order (1-3) at the foot of page 2.
- Landscape is a real page-setup change on every section. The header and footer are real marginal parts in all three sections, not body text.
- `styles set-default-font --all` gave a true document-wide Garamond, including the headings, so the rubric's "partial" trap (sans-serif headings) is avoided.
- The hyperlink covers exactly the words "public domain" with the correct URL.

## Demerits
- **minor: orphaned fragment on page 3.** The last two lines of The Waste Land ("Earth in forgetful snow, feeding / A little life with dried tubers.") spill alone onto page 3, which is otherwise blank. The colophon section is `nextPage` (the default type for s2), so the colophon then sits alone on page 4. The result is two mostly empty pages. A continuous colophon break, or slightly tighter spacing, would have kept the booklet at 3 pages. The agent did not re-render after its last edit and did not notice. Content is not lost, so layout is judged sound.
- **minor: headings use the default Word blue (#2E74B5).** The title and poem headings are blue, which reads more like a word-processor default than the "printed anthology" look the request asked for. The agent noticed this but did not restyle.
- **minor: verse lines are soft breaks, not separate paragraphs.** The rubric literally asks for "each verse line is its own paragraph". The visible outcome (one line per verse line) is correct, so this is not counted as a failure.
- **cosmetic (tool): footnote ids in the read view are out of order.** The read view shows `[^fn3]` before `[^fn1]`, while Word displays them as 1, 2, 3. This is a docx-cli read-view quirk (ids follow creation order), not an agent error.

## Frictions
The agent reported:
- Markdown has no alignment syntax, so the title and subtitle needed a follow-up `edit --alignment center` (it batched them).
- `images add --alignment center` centered the image but not the caption, which needed another edit. **Tool gap:** the caption should inherit the alignment.
- The `sections` ack echoes the requested range (p5-p14), while the read view shows the shifted applies-to range (p6..p16). This is confusing.
- The default heading color and spacing work against the anthology look, and restyling them is not discoverable from set-default-font.
- An inline `<br>` in `create --from` markdown was dropped silently, with no warning. **Tool gap.**
- Dead ends: a zsh shell mishap (the agent's own fault), and an unexplained exit 1 from one `create --help` call that was not reproduced.

**Read:** the path was mostly smooth. The core features (create --from, sections with columns, footers set with page tokens, footnotes add --anchor, images add --caption, set-default-font --all) worked on the first attempt. Frictions were small polish items: caption alignment, `<br>` silently dropped, and the range shift in the ack. The one real lapse was skipping a final render, which would have exposed the page-3 orphan.
