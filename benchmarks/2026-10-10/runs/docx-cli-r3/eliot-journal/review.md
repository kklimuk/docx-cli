# Review: eliot-journal (Authoring: columns, verse, footnotes, links, figure)

**Run:** docx-cli-r3 / haiku

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: 3 page PNGs produced |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived write→read loop | **true** |

## Rubric check

| # | Condition | Result | Evidence |
|---|---|---|---|
| 1 | File exists | PASS | `journal.docx` present |
| 2 | Two-column poem section | PASS | AST: `s1` `columns:2`, `sectionType:nextPage`; page 2 shows two columns |
| 3 | Poem titles are headings | PASS | `##` (Heading 2) for Preludes (I), The Hollow Men (opening), Prufrock (excerpt), The Waste Land (opening) |
| 4 | Verse line breaks preserved | PASS | One line per verse line in read.md and in the render. Lines are hard `<w:br/>` breaks within stanza paragraphs, not separate paragraphs (minor deviation from the literal wording) |
| 5 | Frontispiece and caption | PASS | `images list` → img0 (SVG, alt "Frontispiece", 4×3.11in, centered); p4 "Frontispiece" in style=Caption, centered |
| 6 | ≥2 footnotes | PASS | 3 footnotes (fn1–fn3), all with real content and anchored in the verse |
| 7 | Colophon hyperlink | PASS | link0 "public domain" → https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon single-column | PASS | s0 and s2 have no columns |
| 9 | Landscape | PASS | `docx:page s0 orientation="landscape"`; 15840×12240 |
| 10 | Header and footer | PASS | header "A T. S. Eliot Reader"; footer "Page {page} of {pages}", rendered centered as "Page N of 3" on all pages |
| 11 | Garamond throughout | PASS | `<!-- docx:base font="Garamond" -->`; render shows Garamond in body, title and the blue headings |
| 12 | Clean multi-section render | PASS | Word rendered 3 pages; `docx validate` reports valid (9 parts) |

## Merits

- Every rubric condition is met, and every change shows up on a fresh `docx read` / `read --ast` / `list`.
- The page structure is correct. Page 1 is a single-column title page with a centered title, an italic subtitle, the editor's note, the centered frontispiece and an italic Caption. Page 2 is the two-column verse body, flowing evenly across both columns. Page 3 holds the colophon.
- The footnotes are real editorial glosses, anchored at the right phrases ("you and I", "Michelangelo", "cruellest month"). Word numbers them 1–3 in reading order and places them at the page foot.
- `set-default-font` changed both docDefaults and the theme, so the headings are also Garamond, not Calibri Light. This was the rubric's partial-fail trap, and the agent avoided it.
- The document validates clean against the schema.

## Demerits

- **minor**: The colophon sits alone on page 3. The `s2` break starts a new page, so the final page has a single line (page-003.png). This is acceptable for a colophon, but it looks sparse.
- **minor**: Each verse line is a soft line break inside one paragraph per stanza, not its own paragraph. The output looks right and reads right, but it differs from the rubric's literal "own paragraph" wording.
- **minor**: The header and footer were written as three per-section copies (hdr0–2, ftr0–2) instead of one shared part. The rendered result is identical.
- **cosmetic (read view)**: In read.md the footnote ids follow creation order (`fn1`, `fn3`, `fn2`), while Word numbers them 1, 2, 3. A weak agent could be confused when it maps the read view to the printed output.

## Frictions

The agent reported these frictions and dead ends:

1. **Major: verse line breaks are hard to discover.** With `create --from`, plain newlines in a stanza reflowed into prose. The agent found the trailing-backslash GFM hard break by trial and error, and it is not documented in the `create`/`insert` help. This is a real discoverability gap for the "preserve line breaks" requirement.
2. **Minor: inline footnote markers.** `create --from` turned an inline `[^1]` into its own standalone paragraph. The agent deleted it and used `footnotes add --anchor` instead. This may be a tool bug in GFM footnote import and is worth checking.
3. **Minor: caption alignment.** `images add --caption` did not carry `--alignment center` over to the caption, so a separate `edit --at p4 --alignment center` was needed.
4. **Minor: placeholder paragraph.** An extra `edit --at p3 ""` was needed to clear a placeholder paragraph before inserting the image.
5. **Minor: colophon page and per-section header parts.** These are layout and structure surprises, not failures.

**Read:** The result is correct and polished, but the path was a moderate slog. The agent hit three real dead ends (verse breaks, inline footnotes, caption alignment), each costing extra calls. The verse hard-break discoverability and the `[^n]` import behavior are the two items worth fixing in the tool.
