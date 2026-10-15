# Review — invoice (Table editing + restructure + image replace) — competitor arm: Anthropic docx skill, run r3

## Verdict
| Check | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: one page rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true**: `docx read` shows every change, including the text-box story |

## Merits
- **Placeholder fill is complete and correct.** The company is Northwind Robotics, Inc. / (415) 555-0137 / billing@northwindrobotics.com / 2120 Bryant Street, San Francisco, CA 94110. The invoice fields are NW-1042 / June 8, 2026 / July 8, 2026, and Bill To / Ship To read Acme Health Systems, LLC. The totals are $10,100.00 / -$500.00 / $0.00 / Tax total (8.75%) $840.00 / $0.00 / $10,440.00, and Notes reads "Payment due within 30 days. Wire details on request." The only `$0.00` values left are the intended Shipping and Other.
- **Leftover lines were removed, not blanked.** The "Your City, State Zip" line is gone (the `t0:r0c0` cell went from 6 to 5 paragraphs), and every "Description N" sub-line is gone from the item cells.
- **The fourth row was inserted above the totals.** The table now has the header, 4 item rows and 6 totals rows (Subtotal…Total, all intact), with gridSpan merges and borders the same as in the baseline.
- **Column widths:** the tblGrid is 4547/1000/1800/2300/1575 twips, and Description (3.16in) is the widest column. In the render, `$10,100.00`, `$10,440.00` and "Tax total (8.75%)" all sit on one line.
- **Table polish:** `repeat-header="r0"` (tblHeader) is now set. The F2F2F2 header shading and right-aligned Price/Amount are present; they were already in the template, and the agent checked them rather than assuming.
- **Stamp box:** it still exists, anchored on the Notes paragraph. It reads **Thank you for your business!** in both the mc:Choice and the mc:Fallback copy (grep finds 2 hits and 0 hits for TEMPLATE / DO NOT SEND). The text is still red C00000, bold and centered, and the yellow fill, red border and position look identical to the baseline.
- **Logo:** `image3.svg` is byte-identical to `assets/logo.svg` (sha1 465d83…), embedded through `asvg:svgBlip` with a PNG fallback, at the same 1.05×1.06in floating position. The render shows the new green mark top-left. The footer PayPal mark `image2.png` is byte-identical to the baseline (sha1 fbdbdd…), and the footer rels are unchanged.
- `docx validate` reports the file as valid.

## Demerits
- **minor**: The package holds 3 media files (`image1.png` PNG fallback, `image2.png` footer, `image3.svg`), not 2. Only two pictures are drawn, and an SVG plus PNG fallback is how Word itself embeds SVG, so this is a counting ambiguity, not a real defect.
- **minor**: docProps/core.xml still holds the template author and lastModifiedBy. This was not requested, and the agent flagged it itself.

## Before/after comparison
- **Visual:** only the intended content changed. The INVOICE title, both upper tables, the header-row fill, the totals box borders, the Notes heading, the stamp box's look and the PayPal footer are all unchanged. The Quantity column narrowed to make room for Price/Amount, with no wrapping.
- **Textual:** the read.md diff shows exactly the fills, the removed placeholder lines, the inserted r4 row, the new widths, `repeat-header="r0"` and the stamp wording. There are no unintended changes.

## Frictions
The agent reported:
- **blocker:** the skill never mentions mc:AlternateContent Choice/Fallback twins. The agent found the stale Fallback only by reading the XML.
- **major:** deleting paragraphs, cloning a row, setting tblHeader and keeping tblGrid and tcW consistent all took hand-written regex XML surgery.
- **major:** there is no SVG recipe; the agent had to know svgBlip from memory and rasterize the fallback with rsvg-convert, which is not a listed dependency.
- **minor:**
  - ambiguity in the image count;
  - soffice.py failed silently and left a stale PDF;
  - noisy `validate.py --author` warnings on untracked edits;
  - no step for inspecting existing formatting;
  - leftover metadata.

Dead ends:
- `merge_runs.py` merged the "Phone: " runs, which broke an exact-match replacement.
- A render that looked unchanged was really a failed conversion.
- The agent first mis-mapped the totals labels to the Quantity column.

**Read:** the result is correct, but the path was a slog. Every structural operation was raw XML surgery with little help from the skill. The result is a success with a UX demerit.
