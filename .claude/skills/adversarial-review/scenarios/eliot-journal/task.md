# Authoring: columns, verse, footnotes, links, figure

## Task

Author a NEW journal.docx T. S. Eliot poetry reader: a single-column title section
with an embedded **frontispiece figure + caption**, a TWO-COLUMN body section holding
the poems with their titles as headings and verse line breaks preserved (plus at least
two editor's **footnotes**), and a closing single-column colophon section containing a
**hyperlink**.

## Resolution criteria

journal.docx exists; a two-column section actually holds the poems (verify via `docx
read --ast` sectPr columns or the render showing side-by-side columns); poem titles are
headings; verse line breaks are preserved (lines NOT collapsed into one paragraph); the
frontispiece image is embedded with a caption; at least two footnotes are present
(`docx footnotes list`); a hyperlink is present in the colophon (`docx hyperlinks
list`); the title and colophon are single-column. Multi-section structure renders
cleanly.
