# Grading rubric — invoice (JUDGE ONLY)

The agent never sees this file. It is the ground-truth definition of "correct" for the request in `task.md`. Judge the finished document against these checks.

## Pass conditions

### Placeholder fill
- All placeholder text replaced: no `Item 1`, `Item 2`, `Item 3`, `$0.00` (in line-item cells), `Customer name`, `Your Company Name`, `Your City, State Zip`, `Description N`, or similar template strings remain anywhere in the document body.
- Correct values from `brief.md` are present in the right cells:
  - Company header: Northwind Robotics, Inc. / (415) 555-0137 / billing@northwindrobotics.com / 2120 Bryant Street, San Francisco, CA 94110
  - Invoice #: NW-1042 / Invoice date: June 8, 2026 / Due date: July 8, 2026
  - Bill To / Ship To: Acme Health Systems, LLC / ap@acmehealth.com / 88 Market Street, Suite 400, Chicago, IL 60603 (shipping = same)
  - Totals row values: Subtotal $10,100.00 / Discount −$500.00 / Shipping $0.00 / Tax 8.75% $840.00 / Other $0.00 / Total $10,440.00
  - Notes: "Payment due within 30 days. Wire details on request."

### Leftover placeholder lines removed
- The extra address placeholder line (e.g. "Your City, State Zip") is gone — not left as blank text.
- Per-item "Description N" sub-lines that were not used are gone — not left as empty paragraphs inside a cell.
- Verify via `docx read FILE` that no blank or template-text paragraphs remain in the affected table cells.

### Line-items table has four rows
- The table contains exactly **four** data rows (one per line item) plus the header row and any totals/subtotal rows — the fourth row (Calibration kit / 2 / $300.00 / $600.00) was **inserted** as a new row via `docx tables insert-row`, not typed over a totals row.
- Totals rows are intact and their values match the brief.
- Verify with `docx read FILE --ast` that the line-items table's row count is correct and the totals rows still exist.

### Column widths
- The Description column is the widest column in the line-items table (compare `w:tcW` values in `--ast` output or check that `docx tables set-col-widths` was called).
- Price and Amount column widths are sufficient so that `$10,100.00` fits on one line without wrapping (verify via render or by checking that the column width covers the character count at the document's default font size).

### Line-items table formatting (`tables format`)
- **Header-row shading**: the header row's cells carry a non-`auto` `<w:shd w:fill>` — a light grey (e.g. `D9D9D9`/`EEEEEE`/`F2F2F2`, or any grey name the agent passed). Verify via `docx read FILE` (`docx:cell … shading="…"` on the header cells) or the `w:shd` fill in `--ast`. The expected path is `docx tables format --at tN:r0 --shade …` (one call broadcasting over the row); per-cell calls are also fine.
- **Right-aligned money columns**: every cell in the Price and Amount columns has its paragraph right-aligned (`<w:jc w:val="right">`), surfaced as `docx:cell … halign="right"` in `docx read`. Expected path: `docx tables format --at tN:cC --halign right` (column broadcast) for each of the two columns, or an equivalent per-cell/`edit --alignment` pass. Left-aligned dollar columns = fail.
- **Repeat header on page 2**: the header row carries `<w:trPr><w:tblHeader/>`, surfaced as `docx:table … repeat-header="r0"`. Expected path: `docx tables format --at tN:r0 --repeat-header`. (Under track-changes this rides a `[docx-cli]` audit comment, not a revision — that's correct, not a defect.)
- These three are *additive* polish: they must not break the fill/insert-row/width work above (the header row is still the column titles, the four data rows are intact).

### Logo replaced
- The top-left image is the new mark sourced from `assets/logo.svg`.
- The document still contains **exactly two** embedded images (new logo + footer payment mark).
- Verify via `docx images list FILE`: two entries, and the logo entry's dimensions / source match `assets/logo.svg`.
- Footer payment mark is preserved: the footer image relationship is still present and unreferenced relationships were not dropped.

### Formatting preserved
- Font, size, bold/italic on existing cells are unchanged (check a sample run via `--ast`).
- Table borders and shading are intact.
- The "PAID" / payment mark in the footer is untouched.

## How to verify

- `docx read FILE` — scan for remaining placeholder strings; check Notes text; confirm four line items visible.
- `docx read FILE --ast` — inspect table row counts, `w:tcW` column widths, run-level formatting on sample cells.
- `docx images list FILE` — confirm exactly two images; confirm logo image index/rId is different from the original.
- Render before/after with `docx render FILE` and compare pages: logo changed, column widths visible, no wrapping dollar values, footer mark intact.
- Check that `tables insert-row` (not an overwrite) was used: the totals row must still exist below the four item rows.
