# Authoring: columns, verse, footnotes, links, figure

## Task

Author a NEW journal.docx T. S. Eliot poetry reader: a single-column title section
with an embedded **frontispiece figure + caption**, a TWO-COLUMN body section holding
the poems with their titles as headings and verse line breaks preserved (plus at least
two editor's **footnotes**), and a closing single-column colophon section containing a
**hyperlink**. Set the whole document to **landscape** orientation (the wide page suits
the two-column verse). Add a **running header** with the collection title ("A T. S.
Eliot Reader") and **centered page numbers** in the footer ("Page X of Y").

## Resolution criteria

journal.docx exists; a two-column section actually holds the poems (verify via `docx
read --ast` sectPr columns or the render showing side-by-side columns); poem titles are
headings; verse line breaks are preserved (lines NOT collapsed into one paragraph); the
frontispiece image is embedded with a caption; at least two footnotes are present
(`docx footnotes list`); a hyperlink is present in the colophon (`docx hyperlinks
list`); the title and colophon are single-column. The page is **landscape** (`docx read
--ast` shows the trailing section's `pageOrientation: "landscape"` with `pageWidth >
pageHeight`, or `docx read` shows the `docx:page … orientation="landscape"` note).
A **running header** carries the collection title and the **footer shows page numbers**
(`docx headers list` shows a header with the title; `docx footers list` shows a footer
whose text is a page-number field — e.g. "Page {page} of {pages}"; the render shows the
title at the top and a page number at the bottom of each page). Multi-section structure
renders cleanly.
