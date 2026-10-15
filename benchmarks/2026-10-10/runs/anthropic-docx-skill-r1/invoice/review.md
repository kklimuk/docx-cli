# Review: invoice (Table editing + restructure + image replace), Anthropic docx skill arm (haiku), r1

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes. One page rendered.
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes. Every intended change shows in the Word render, and `docx read` also returns it.

## Merits
- Every placeholder is filled with the correct value. Company, phone, email and the one-line address are in place. The invoice number is NW-1042, dated June 8, 2026 and due July 8, 2026. Bill To and Ship To both carry the Acme details, with the shipping address set to the billing address.
- The extra "Your City, State Zip" line is gone. The read view shows `t0:r0c0` going from p5 down to p4, with no blank leftover.
- All three "Description N" sub-lines are gone, and each item cell has exactly one paragraph (`t2:r1c0:p0` and so on).
- The fourth row (Calibration kit / 2 / $300.00 / $600.00) is a new row above the totals block. All six totals rows survive with the correct values: $10,100.00, -$500.00, $0.00, Tax total (8.75%) $840.00, $0.00 and $10,440.00.
- Column widths: the grid changed from `3.127,0.853,1.542,1.176,1.094in` to `3.403,0.694,0.799,1.597,1.25in`. Description is the widest column in every row: 4.10in across its gridSpan=2 in the header and 3.40in in the data rows. No dollar value wraps in the render. Merges (gridSpan) are intact.
- Header row: the grey `F2F2F2` fill is present on every cell. It was already in the template, and the agent kept it. The row now carries `<w:tblHeader/>`, and read shows `repeat-header="r0"`.
- Price and Amount: every cell reads `halign="right"`, including the four item rows and the totals.
- Stamp text box: it still exists (`tbx0` is anchored on p10, wrap square, aligned right) and reads "Thank you for your business!". `find "DO NOT SEND"` returns no matches, and `find "Thank you for your business!"` returns `tbx0:p0:0-28`. The new string appears twice in document.xml, so the Choice and Fallback copies are both updated. Red bold centered formatting is kept, and the box geometry and fill look identical to the baseline.
- Logo: image1.png is a 512×512 PNG rasterized from assets/logo.svg. It renders as the new green blob mark at the original top-left size and position (1.05×1.06in). The footer PayPal mark (image2.png, footer rel rId2) is untouched. There are still exactly two referenced images. The pre-existing unreferenced `image1.svg` is from the template, not added by the agent.
- Notes: "Payment due within 30 days. Wire details on request." is in the previously empty Notes body paragraph, with run properties matched to the body text.
- Fonts, sizes, borders and the footer all look the same as the baseline.

## Demerits
- **minor**: The logo is a raster PNG swap. No svgBlip extension was added and the SVG is not embedded. The visual result is correct, but the vector source is lost.
- **minor**: The old Notes sentence "Thank you for your business." stays as a flow paragraph next to the new stamp "Thank you for your business!", so the wording is duplicated. This paragraph is also the stamp's anchor, and the rubric only asks that the new Notes text be present. The request is ambiguous about whether this sentence was part of the "Notes body". Cosmetic.

## Frictions
- The agent reported several gaps in the skill. It has no SVG-logo route, no recipe for cloning table rows or recomputing grid widths, and no warning about the duplicate mc:AlternateContent Choice/Fallback text. The agent handled the Choice/Fallback pair correctly anyway.
- Dead ends: `validate.py --author` reported every untracked edit as a failure, and the agent had to re-run without it. `timeout` is not available on macOS. The new Notes text first rendered at a smaller size because the empty paragraph had no rPr, which the agent fixed.
- Overall the run was a moderate slog. Everything was hand-written lxml XML surgery on document.xml, but the agent reached a fully correct result.
