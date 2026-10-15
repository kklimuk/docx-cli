# Review — eliot-journal (authoring) — OfficeCLI, run r1

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes (3 page PNGs)
- **Layout sound:** yes
- **Formatting preserved:** preserved (n/a baseline; authored formatting reads back intact)
- **Survived read loop:** yes

## Rubric check
1. File exists: yes, `journal.docx`.
2. Two-column body: s1 `<w:cols w:num="2">`, and the render shows the poems in two columns on page 2.
3. Headings: Preludes (I), The Hollow Men, The Love Song of J. Alfred Prufrock and The Waste Land (I. The Burial of the Dead) all use `Heading1`.
4. Verse lines: one paragraph per line, with an empty paragraph at each stanza break (p19, p45). The text matches the source.
5. Frontispiece: `img0` (PNG plus an `asvg:svgBlip` SVG), 9 cm wide, centered on page 1, with an italic "Frontispiece" Caption paragraph below it.
6. Footnotes: three with non-empty text (fn1–fn3), rendered at the foot of page 2.
7. Hyperlink: `link0` "public domain" goes to https://www.gutenberg.org/ebooks/1321 and renders as a blue underlined link.
8. Single-column title and colophon: s0 has no cols, and s2 has `cols num=1`.
9. Landscape: every sectPr has `pgSz 15840x12240 orient=landscape`.
10. Header and footer: the header "A T. S. Eliot Reader" appears on all 3 pages. The footer uses real PAGE/NUMPAGES complex fields and renders "Page 1 of 3", "Page 2 of 3" and "Page 3 of 3", centered. s1 inherits the references from s0.
11. Garamond: `docx:base font="Garamond"` is set in docDefaults. The Title, Heading1, Caption and FootnoteText styles name Garamond explicitly with no theme-font attributes. The render shows serif text in the body and the headings.
12. Clean render: yes. Section transitions are visible (title, then columns, then colophon).

## Merits
- Every requirement is met and visible in the Word render. The anthology looks polished.
- Real PAGE/NUMPAGES fields evaluate correctly in Word.
- The SVG is embedded properly, with a PNG fallback and an svgBlip extension.
- The agent caught and fixed the theme-font override on Title/Heading1 by replacing the raw XML, so the headings really are Garamond.
- The agent found that the header/footer references were missing from the first section and fixed it, so the title page also has the header and footer.
- The section/column model is correct: single, then 2 columns, then single.

## Demerits
- **minor**: Footnote 2 is anchored on "Let us go then, you and I," but its text starts "Michelangelo: …" and glosses lines further down. The anchor and the gloss do not match.
- **minor**: The FootnoteReference, FootnoteText and Caption styles were created with `customStyle="true"`, although they are built-in style names. Word tolerates this.
- **minor (cosmetic)**: The "Colophon" heading is left-aligned while the colophon text is centered. Page 3 is almost empty, which is acceptable for a colophon.
- Not an OfficeCLI defect: docx-cli's `headers/footers list` and `read` show the footer as the cached text "Page 1 of 1" rather than `{page}`/`{pages}` tokens. The XML holds live fields, and Word renders them correctly.

## Frictions
The agent reported:
- `\n` in `--prop text` splits into separate paragraphs, not `<w:br/>`, which contradicts the docs.
- A header/footer added on `/` is referenced only by the final sectPr.
- Setting a font on a style leaves the theme rFonts attributes in place, so the theme font still wins in Word. The agent rated this a blocker.
- Built-in footnote and caption styles are not auto-defined.
- Footer field runs did not pick up the size set on the paragraph.
- The ownership model for `add section` and its sectPr is confusing.
- `view issues` gives noisy indent warnings on verse.

Its dead ends were the footnote path, the header references and the theme fonts, each fixed with raw-set XML. Read: the result is correct, but getting there was a moderate slog. Three of the fixes needed raw XML surgery, which a weaker agent might not have managed. The theme-font and section header-reference issues are real OfficeCLI tool bugs.
