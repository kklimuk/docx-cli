# PSA — surgical tracked edit (preservation test)

`professional-services-agreement.docx` is a long, heavily-formatted contract: a SOW
cover page with colored runs, yellow placeholders, gray italic *drafting notes*,
checkbox options, multiple tables, **ten** headers/footers, and an entire attached
"Standard Terms" part. This task is about **precision and preservation**: make two
tiny edits and leave absolutely everything else untouched.

## The only changes you should make

Turn **tracked changes ON** first, then:

1. **Payment Period** cell (label "Payment Period"): replace the placeholder
   `[Fill in payment terms, e.g., 30 days from Customer's receipt of invoice]`
   with: **30 days from Customer's receipt of invoice**
2. **Invoice Period** cell (label "Invoice Period"): replace the placeholder
   `[Fill in cadence of sending invoices, ...]` with: **Monthly**

## Hard constraints

- **Do not** fill in, delete, or reword any other placeholder, drafting note,
  checkbox, heading, table, header, or footer.
- **Do not** delete the gray italic drafting notes or the attached Standard Terms.
- Both edits must be recorded as **tracked changes** (visible redlines), not silent
  replacements.

## What "done" looks like

`track-changes list` shows exactly your two edits and nothing else. A before/after
render is visually identical everywhere except those two cells. Every other
placeholder, the drafting notes, all ten headers/footers, and the full Standard
Terms section are still present and unchanged.
