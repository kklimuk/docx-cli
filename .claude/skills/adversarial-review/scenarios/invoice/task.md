# Table editing + restructure + image replace

## Task

Three jobs on the invoice:

1. Fill the three tables with the supplied company/customer details, line items, and
   totals — keep the table layout and formatting intact. Remove any leftover
   placeholder lines you don't fill (the extra address line, the per-item
   "Description N" sub-lines) so no template text is left behind.
2. The supplier added a **fourth** line item, but the line-items table only has rows
   for three. **Insert a new row** for it (use the table tools — don't overwrite a
   totals row), fill it in, then **set the line-items table's column widths** so the
   Description column is the widest — while keeping the Price/Amount columns wide
   enough that the dollar values don't wrap.
3. Replace the logo in the **top-left corner** with the new company mark at
   `assets/logo.svg`.

Leave the payment mark down in the footer alone.

## Resolution criteria

All placeholder cells replaced with the supplied values and leftover placeholder lines
removed (no "Your City, State Zip" / "Description N" text left behind); totals match
the data sheet. The line-items table holds **four** line items — the fourth added as a
new row via the table tools (not by overwriting a totals row) — and its column widths
were adjusted so Description is widest and no dollar value wraps.
The top-left logo is replaced with the new mark from `assets/logo.svg` (a green
ink-blob) and the footer's payment image is preserved (the doc still has two embedded
images). The three tables remain well-formed. Changes visible via `docx read`.
