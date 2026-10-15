# Review — invoice (bucket: docx-cli-r3)

**Scenario:** Table editing + restructure + image replace

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (output page-001.png produced) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** (a fresh `docx read` matches `renders/output/read.md` exactly) |

## Merits

- **Every placeholder filled with the correct value.** Company, phone, email, the one-line address, NW-1042, June 8, 2026, July 8, 2026, Bill To and Ship To (Ship To = billing), and all totals: $10,100.00 / -$500.00 / $0.00 / $840.00 / $0.00 / $10,440.00. Notes now reads "Payment due within 30 days. Wire details on request."
- **Leftover placeholder lines removed.** "Your City, State Zip" (t0:r0c0:p5) is gone, and so are the three grey "Description N" sub-lines. Each item cell now holds a single paragraph, with no empty paragraphs left behind.
- **Placeholder sweep is clean.** `find` returns no matches for `Item`, `Description `, `Customer`, `Your`, `Jan, 1`, `0001`, `TEMPLATE` or `DO NOT SEND`. `$0.00` appears only in t2:r7c3 (Shipping) and t2:r9c3 (Other), and the brief specifies both values.
- **Fourth row was inserted, not overwritten.** t2 now has the header, 4 item rows (the Calibration kit row is r4 with 2 / $300.00 / $600.00) and all 6 totals rows (Subtotal through Total, r5–r10). The row has the same borders, gridSpan and halign as the other item rows.
- **Column widths.** The grid changed from `3.127,0.853,1.542,1.176,1.094in` to `3.611,0.278,0.833,1.389,1.682in`. Description (3.61in) is the widest, Price is 1.39in and Amount is 1.68in. In the render, `$10,100.00` and `$10,440.00` each sit on one line. The merged Quantity cell (gridSpan 2) and the header's gridSpan are intact.
- **Table polish.** Header cells keep `shading="F2F2F2"`, and Price/Amount keep `halign="right"`; both were already in the template, so they were satisfied with no change. The agent added `repeat-header="r0"` on t2.
- **Logo replaced.** img0 is now `image/svg+xml` and `word/media/image1.svg` matches `assets/logo.svg`. The tool stripped a comment and an `<animate>` element; otherwise the file is identical. The old PNG logo is no longer in the document. The footer PayPal mark (`media/image2.png`, rId2 in the footer rels) is untouched and renders. The package has exactly two images: the new logo and the footer mark.
- **Stamp rewritten in place.** `find "Thank you for your business!"` returns `tbx0:p0:0-28`. The box keeps its red bold `#C00000` text, centered alignment, cream fill and red border in the render. `raw get --at p10` shows the new text in both the Choice and the Fallback copies, so the Fallback was synced.
- `docx validate` reports valid (8 parts checked).

## Demerits

- **minor:** The "Tax total" label was not changed to "Tax total (8.75%)". The value $840.00 is correct, and the rubric only checks the value, so this is a judgment call and not a failure.
- **minor (tool, not agent):** `docx images list` returns only img0 and omits the footer image, but the rubric expects "two entries." The document does have both images (confirmed in the media folder and the footer rels). The verb under-reports.

Visual before/after: the only changes are the intended ones. Header, INVOICE title, footer and the stamp box's position and look are all unchanged. Because two sub-lines were removed, the table and Notes sit higher on the page, which is expected. No content was pushed off the page and there is no blank page.

## Frictions

The agent reported these frictions and dead ends:
- `tables format --repeat-header` prints a "layout changed — verify it renders right" hint even though repeating the header moves nothing (noise).
- `images list` omits the footer and header images. To confirm the "two images" requirement, the agent had to inspect `raw get --at ftr0`. This is a real gap in the tool.
- Nothing tells the agent that the header shading and right alignment are already set. It had to check the AST to tell a no-op apart from a missed change.
- `tables insert-row --cells` is comma-split, so "Extended support, 12 mo." can't be passed in one call. The agent worked around it with `edit`.
- `$300.00` gets shell-mangled unless it is single-quoted. The agent handled it, but a weaker run could trip on it.

**Overall:** The path was fairly smooth, with a few small detours caused by real tool gaps: the incomplete images list, the comma split in `--cells`, and the noisy layout hint. The output is correct on every rubric check.
