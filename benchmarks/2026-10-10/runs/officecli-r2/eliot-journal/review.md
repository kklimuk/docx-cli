# Review — eliot-journal (authoring bucket) — OfficeCLI, run r2

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes (3 page PNGs)
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes

## Rubric check
| # | Criterion | Result |
|---|---|---|
| 1 | journal.docx exists | Pass |
| 2 | Two-column poem section | Pass: s1 `<w:cols w:num="2" equalWidth>`; page 2 renders two columns |
| 3 | Poem titles are headings | Pass: Heading1 for all four poems |
| 4 | Verse line breaks preserved | Pass, with a caveat: lines are `<w:br/>` breaks inside one paragraph per poem rather than one paragraph per line; every line renders and reads on its own line |
| 5 | Frontispiece and caption | Pass: img0 (3.54x2.76in, alt text), italic "Frontispiece" caption below it on page 1 |
| 6 | At least two footnotes | Pass: fn1 (Michelangelo/Prufrock) and fn2 (Waste Land's April) |
| 7 | Colophon hyperlink | Pass: link0 "public domain" goes to https://www.gutenberg.org/ebooks/1321 |
| 8 | Title and colophon are single-column | Pass: s0 has cols=1, s2 has no cols setting |
| 9 | Landscape | Pass: all sectPr are 16838x11906 with orient=landscape |
| 10 | Header and footer | Pass: the header "A T. S. Eliot Reader" is on every page. The footer uses real PAGE/NUMPAGES fields and renders "Page N of 3". docx-cli lists the footer text as the cached value "Page 1 of 1"; that is a limit of docx-cli's reader, not an OfficeCLI defect |
| 11 | Garamond throughout | Pass: `<!-- docx:base font="Garamond" -->`, and the render shows serif in the body, title, subtitle and headings |
| 12 | Renders cleanly | Pass: 3 pages and no blank pages. The section changes are visible |

## Merits
- The three-section structure is right: a single-column title page, a two-column verse body and a single-column colophon, all landscape.
- The title page has the centered Title, the Subtitle, the italic editor's note, the frontispiece with alt text, and the caption.
- All four poems are present with Heading1 titles. Verse lines and stanza gaps are intact, and the poems flow side by side on page 2.
- The two editor's footnotes have real content and render at the foot of page 2.
- The colophon hyperlink is correct and styled as a link.
- The running header is on all pages, and the centered field-based "Page X of Y" renders correctly (1/3, 2/3, 3/3).
- Garamond is set in docDefaults and on the styles, with the theme font bindings cleared, so Word really renders Garamond.
- The agent removed its probe and placeholder paragraphs, so the document is clean.

## Demerits
- **Minor — footnote reference marks are not superscript.** Body runs use `rStyle="FootnoteReference"`, but `styles.xml` defines only Normal, Title, Subtitle and Heading1. The style is dangling, so page 2 shows "Talking of Michelangelo.1" and "dried tubers.2" at the text baseline. The marks in the footnote area are superscript.
- **Minor — verse structure.** Each poem is one paragraph with line breaks (p7, p9, p11, p13) rather than one paragraph per verse line, which the rubric's wording prefers. The visual result and the read view both meet the task's requirement.

## Frictions (agent-reported)
- `add --type section` gave the earlier section a portrait page size and no columns, so it needed separate `set` calls (major).
- Header/footer `add` attached only to the final sectPr, silently. The agent fixed this with a raw-set prepend of headerReference/footerReference on the first sectPr (major).
- Setting a style font left `asciiTheme`/`hAnsiTheme` in place, which silently overrides Garamond. It was cleared only by an undocumented empty-value set (major).
- Batch JSON did not decode `\v` the way `--prop` does, so the verse was rewritten with a real U+000B (minor).
- `view issues` flagged verse for a missing first-line indent and flagged the section carrier paragraphs as empty (minor).
- A "Page X of Y" footer took four `add` calls (minor).

**Read:** the result is correct, but the path was a slog. Three of the traps were silent and needed OOXML knowledge to diagnose: the section page setup, the header attachment and the theme-font override. A less careful weak agent would probably have shipped a document with the wrong font and no header on the first pages. The agent's self-reported "partial" was too cautious, because the output meets every rubric item.
