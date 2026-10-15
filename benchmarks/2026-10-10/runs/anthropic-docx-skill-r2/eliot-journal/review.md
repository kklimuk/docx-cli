# Review: eliot-journal (authoring) — competitor arm (Anthropic docx skill, haiku, r2)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes. Three pages rendered.
- **Layout sound:** yes
- **Formatting preserved:** n/a (authored from scratch)
- **Survived read loop:** yes. All intended content is present in the render and in the `docx read` output.

## Rubric check
| # | Criterion | Result |
|---|---|---|
| 1 | journal.docx exists | PASS |
| 2 | Two-column poem section | PASS: `docx:section s1 cols="2"`; the render shows two columns with a separator rule |
| 3 | Poem titles are headings | PASS: four `##` headings (Heading 2) |
| 4 | Verse line breaks preserved | PASS: one paragraph per line, stanza breaks as in the source |
| 5 | Frontispiece and caption | PASS: img0 PNG, 3.75x2.92in, with an italic "Frontispiece" caption |
| 6 | Two or more footnotes | PASS: fn1–fn4, all non-empty |
| 7 | Colophon hyperlink | PASS: link0 "public domain" goes to https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon single-column | PASS: s0 and s2 have no cols |
| 9 | Landscape | PASS: `docx:page s0 orientation="landscape"`; the pages render wide |
| 10 | Header and footer page numbers | PASS (render): the header appears on every page and the footer reads "Page 1 of 3" through "Page 3 of 3", centered |
| 11 | Garamond throughout | PASS: `docx:base font="Garamond"`; body and headings both render in Garamond |
| 12 | Renders cleanly | PASS |

## Merits
- The three-section structure is correct: title page, then the two-column body, then the colophon. All of it is landscape.
- The headings use a real heading style and render in Garamond.
- The verse lines are intact, and the stanza gaps faithfully match `assets/eliot-poems.md`.
- The four footnotes are scholarly and sit at the foot of each column.
- The external hyperlink has the correct URL and anchor text.
- The frontispiece is rasterized from SVG, centered, has alt text, and carries a caption.
- The PAGE and NUMPAGES fields work in Word.

## Demerits
- **Minor:** The colophon is alone on a mostly blank page 3 because of the nextPage break (see page-003.png). The rubric allows it, but a continuous break would look tighter.
- **Minor:** The running header also prints on the title page. This is consistent with "every page", but a printed anthology would usually suppress it there.

## Tool-side note (docx-cli, not counted against the competitor)
`docx footers list` / `read` reports `text="Page  of"`. docx-js puts `w:fldChar` and `w:instrText` inside the same `w:r` as the `w:t` text, and docx-cli's marginal reader drops those in-run fields instead of emitting `{page}`/`{pages}` tokens. Word renders the fields correctly, so this is a docx-cli reader gap.

## Frictions
- `require('docx')` failed with MODULE_NOT_FOUND from the scenario cwd. SKILL.md says the module is "preinstalled", but it is only installed in the skill's own folder. The agent worked around this with an absolute path.
- docx-js ImageRun rejects SVG, and SKILL.md doesn't cover rasterizing it. The agent used rsvg-convert.
- SKILL.md has no recipes for columns, footnotes, or hyperlinks. The agent wrote these APIs from memory and then checked them in the XML.
- The agent's own soffice relative-path slip cost one retry.
- The create workflow doesn't mention validation. The agent ran validate.py on its own initiative.

**Read:** It was a moderate slog on the setup and image side, with three or four detours, but the authoring itself went through cleanly and the output is complete and correct.
