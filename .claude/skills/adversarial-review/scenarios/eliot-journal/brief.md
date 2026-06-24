# Poetry-journal brief — a T. S. Eliot reader

Author a brand-new `.docx` "poetry journal" dedicated to **T. S. Eliot**, using the
poems in `assets/eliot-poems.md`. There is no template — build it from Markdown/CLI
from scratch. The point is to exercise rich authoring: headings, multi-column
**sections**, verse line breaks, **footnotes**, a **hyperlink**, and an **embedded
figure**.

## Required structure

1. A **title page**: a centered title ("A T. S. Eliot Reader"), a subtitle, and a
   short one-paragraph editor's note — a single-column section. Embed the
   **frontispiece** image `assets/frontispiece.svg` on the title page with a caption
   (e.g. "Frontispiece"). It sizes itself from the SVG; pass a width if you'd like it
   smaller.
2. A **section break** into a **two-column** section that holds the poems, so the verse
   flows in columns like a printed anthology.
3. Each poem gets a **heading** with its title, then the verse. Keep the line breaks of
   the verse intact (lines must not be reflowed into one paragraph). Add at least **two
   footnotes** — short editor's annotations on a word or line (e.g. glossing an allusion
   in *The Waste Land* or a phrase in *Prufrock*).
4. Include at least: *Preludes (I)*, *The Hollow Men* (opening), and *The Love Song of
   J. Alfred Prufrock* (excerpt). Add *The Waste Land* opening if you have room.
5. End with a short single-column **colophon** section: "Compiled June 2026. All poems
   in the public domain." Make the words **"public domain"** a **hyperlink** to
   https://www.gutenberg.org/ebooks/1321 (The Waste Land on Project Gutenberg).
6. Print the journal in **landscape** — set the whole document's page orientation to
   landscape so the wide page carries the two columns comfortably. (Page orientation is
   a section/page-setup property — set it, don't fake it by resizing anything.)

## What "done" looks like

A journal that renders with a single-column title section (carrying the frontispiece
figure + its caption), a two-column body where the poems sit side by side, correct poem
headings, preserved verse line breaks, at least two footnotes, and a closing colophon
whose "public domain" text links out. Save it as **`journal.docx`** in this scenario
folder.
