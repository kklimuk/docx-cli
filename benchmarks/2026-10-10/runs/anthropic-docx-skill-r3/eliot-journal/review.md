# Review: eliot-journal (authoring bucket), competitor arm: Anthropic docx skill, run r3

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (4 landscape pages)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true (all intended content is present in the render and in the docx-cli read)

## Rubric check
| # | Check | Result |
|---|---|---|
| 1 | journal.docx exists | Pass |
| 2 | Two-column poem section | Pass: `docx:section s1 cols="2"`, and page 2 renders two columns |
| 3 | Poem titles are headings | Pass: Heading2 on p5, p9, p11, p15 (all four poems, including the optional Waste Land) |
| 4 | Verse line breaks preserved | Pass, with a minor deviation: lines render separately, but they are `w:br` breaks inside one stanza paragraph, not one paragraph per line |
| 5 | Frontispiece and caption | Pass: img0 (PNG fallback of the SVG, 3.75x2.92in) plus an italic "Frontispiece" caption |
| 6 | At least 2 footnotes | Pass: fn1-fn3, all with content |
| 7 | Hyperlink in colophon | Pass: link0 "public domain" goes to https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon single-column | Pass: s0 and s2 have no cols |
| 9 | Landscape | Pass: `orientation="landscape"`, and the pages render wide |
| 10 | Header and footer | Pass: the header renders "A T. S. Eliot Reader" on every page and the footer renders "Page N of 4", centered. The docx-cli reader shows "Page  of" (see note) |
| 11 | Garamond throughout | Pass: `docx:base font="Garamond"`, and headings render in a serif |
| 12 | Renders cleanly | Pass |

## Merits
- The section structure is exactly right: single-column title, two-column landscape body, single-column colophon.
- The typography is clean. The centered Garamond title, italic subtitle, framed frontispiece and italic caption read like a real anthology.
- All four poems have Heading2 titles, the verse is set line by line with a hanging indent, and stanza gaps are kept (for example, "And then the lighting of the lamps.").
- There are three footnotes, and they render at the foot of each page.
- The hyperlink target and the colophon wording are both exact.
- Page chrome works: an italic running header, and a centered "Page X of Y" built from real PAGE/NUMPAGES fields.

## Demerits
- **Minor:** the verse uses soft line breaks inside one paragraph per stanza, not one paragraph per line. The render and the read view still show separate lines, so the visible intent holds.
- **Minor:** page 3 is mostly empty (two Waste Land lines and a note) before the colophon starts on a new page. This is normal flow from a nextPage section, not a stacked break.
- **Not a competitor defect:** `docx footers list` and `read` show "Page  of". docx-js puts the fldChar/instrText in the same run as the literal "Page ", and docx-cli's reader does not turn that shape into `{page}`/`{pages}` tokens. Word renders the footer correctly. This is a docx-cli reader gap worth following up.

## Frictions
- SKILL.md has no SVG guidance. The agent found `rsvg-convert` on its own and passed a PNG fallback.
- There is no recipe for multi-column sections, mixed section page setup, footnotes, ExternalHyperlink, or PageNumber fields, so the agent relied on remembering the docx-js API.
- `validate.py` failed under `python3 -I` with "No module named helpers" and worked without `-I`.
- The first build dropped "in the" from the colophon. Only a render caught it, because the skill has no step for checking the text against the source.
- soffice.py prints a noisy "Task policy set failed" line on macOS.

**Read:** the result is fully correct, but the path was a moderate slog. Success depended on the agent's own knowledge of the docx-js API, not on guidance from the skill.
