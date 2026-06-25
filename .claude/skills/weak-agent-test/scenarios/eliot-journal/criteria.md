# Grading rubric — eliot-journal (JUDGE ONLY)

## Pass conditions

1. **File exists.** `journal.docx` is present in the scenario folder.

2. **Two-column poem section.** `docx read --ast journal.docx` shows a section whose
   `sectPr` has `columns: 2` (or equivalent); OR the render shows the verse flowing in
   two side-by-side columns.

3. **Poem titles are headings.** The AST contains heading-level paragraphs (style
   `Heading 1`/`Heading 2` or equivalent) for each poem title — *Preludes (I)*,
   *The Hollow Men*, *The Love Song of J. Alfred Prufrock*, and optionally *The Waste
   Land* opening.

4. **Verse line breaks preserved.** Each verse line is its own paragraph (not reflowed
   into a single paragraph); confirmed by `docx read` showing separate lines, not
   collapsed prose.

5. **Frontispiece embedded with caption.** `docx images list journal.docx` returns at
   least one image; the surrounding text or adjacent paragraph provides a visible
   caption (e.g. "Frontispiece") in the title section.

6. **At least two footnotes.** `docx footnotes list journal.docx` returns ≥ 2 entries
   with non-empty content.

7. **Hyperlink in colophon.** `docx hyperlinks list journal.docx` returns an entry whose
   `url` is `https://www.gutenberg.org/ebooks/1321` and whose anchor text is "public
   domain" (or contains those words).

8. **Title and colophon are single-column.** The first and last sections have no
   `columns` property (or `columns: 1`) in the AST sectPr.

9. **Landscape orientation.** `docx read --ast journal.docx` shows the trailing
   section's `pageOrientation: "landscape"` with `pageWidth > pageHeight`; OR `docx
   read journal.docx` shows a `<!-- docx:page … orientation="landscape" -->` note.

10. **Running header and footer page numbers.** `docx headers list journal.docx` shows a
    header whose text content contains "A T. S. Eliot Reader". `docx footers list
    journal.docx` shows a footer whose text contains a page-number field — e.g.
    "Page {page} of {pages}" (or similar `{page}`/`{pages}` token form). The render
    shows the title at the top and a page number at the bottom of each page.

11. **Document-wide serif (Garamond) font.** `docx read journal.docx` shows a
    `<!-- docx:base font="Garamond" … -->` note, confirming Garamond is the document
    default. The render shows a serif face in BOTH body text AND poem-title headings. If
    a heading is still rendered in the default sans-serif, only the body font was changed
    (not the document-wide default) — this is a partial fail.

12. **Multi-section structure renders cleanly.** `docx render journal.docx` produces a
    PDF/PNG without errors; column breaks and section transitions are visible on the
    rendered pages.

## How to verify

```bash
# Quick structural checks
docx read --ast journal.docx | jq '.sections[] | {cols: .sectPr.columns, orient: .sectPr.pageOrientation}'
docx footnotes list journal.docx
docx hyperlinks list journal.docx
docx headers list journal.docx
docx footers list journal.docx
docx images list journal.docx

# Full read (check docx:base note, verse lines, headings)
docx read journal.docx

# Render for visual confirmation
docx render journal.docx --output journal-preview.png
```
