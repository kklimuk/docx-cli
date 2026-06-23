# Styling fidelity + drawing preservation + paragraph layout

## Task

Fill the Harvard résumé template for the supplied candidate, preserving heading
styles, the right-aligning tab stops on dates, and the bullet lists. Remove the
bracketed [Note: ...] helper text. Leave the [drawing] element alone. Then tidy the
paragraph layout: add space before each section heading and space after each entry
so the cramped template breathes (paragraph spacing — set in place, don't retype).

## Resolution criteria

Name, contact, education, experience, leadership filled from the candidate sheet;
[Note: ...] helpers gone; section headings keep their Heading style; dates still
right-align at the tab; the [drawing] element survives. **Paragraph layout:** the
three section headings have added space-before and the experience/leadership entries
have space-after (visible breathing room in the render; `docx read --ast` shows the
`spacing.before`/`spacing.after` twips, or `docx read` shows the `docx:p …
space-before/space-after` hints). Render should look like a clean, well-spaced résumé.
