# Review: invoice (OfficeCLI, run r1)

**Scenario:** invoice (table editing, restructuring, and an image replace)
**Bucket:** officecli-r1

## Verdict

| Dimension | Result |
|---|---|
| Task success | **partial** |
| Renders correctly (Word opened it) | true |
| Layout sound | true |
| Formatting preserved | preserved (with one small-font defect on the new Notes line) |
| Survived read loop | true |

Almost every requirement is met and verified. The exception is the Notes instruction. The task said to *replace* the Notes body with "Payment due within 30 days. Wire details on request." The agent put that sentence into the blank 8pt spacer paragraph above the existing body. The original body line, "Thank you for your business.", is still there. In the Word render the new sentence shows as tiny 8pt text, with the old template line under it at full size. A client would see this.

## Merits

- **Header and "from" block.** These are all filled: Northwind Robotics, Inc., (415) 555-0137, billing@northwindrobotics.com, and 2120 Bryant Street, San Francisco, CA 94110. The extra "Your City, State Zip" line was removed as a paragraph, not blanked: `t0:r0c0` now has p0–p4 where the baseline had p0–p5.
- **Invoice details.** NW-1042, June 8, 2026 and July 8, 2026 are correct.
- **Bill To / Ship To.** Both are correct, with shipping set to the billing address.
- **Line items.** There are four rows: RX-7 / On-site / Extended support / Calibration kit, with correct quantity, price and amount values. The fourth row was a clone of an existing row, so the totals block (r5–r10) is intact. The per-item "Description N" sub-lines were removed as paragraphs; each item cell now has only p0.
- **Totals.** The values are $10,100.00, -$500.00, $0.00, Tax total (8.75%) $840.00, $0.00 and $10,440.00.
- **Header row.** It is shaded D9D9D9 (the baseline was F2F2F2, so the change is real) and has `<w:tblHeader/>` (`repeat-header="r0"`).
- **Money columns.** Price and Amount are right-aligned (`halign="right"`). The template already had this, and it was kept.
- **Column widths.** Description is the widest column and no dollar value wraps in the render. The widths are unchanged from the template, which already met the requirement.
- **Merges.** All three tables have their gridSpans and borders intact.
- **Stamp box.** It is kept and reads **Thank you for your business!**, bold, red #C00000, centered. The text is updated in BOTH the mc:Choice and mc:Fallback stories (both copies are in document.xml). `find "DO NOT SEND"` returns no matches. `find "Thank you for your business!"` returns `tbx0:p0:0-28`. The box keeps its fill, border and position in the render.
- **Logo.** It is replaced with the SVG from `assets/logo.svg` (the green blob shows in Word). The old placeholder image part was dropped and nothing else references it. The footer PayPal mark (`footer1.xml.rels` rId2 → media/image2.png) is untouched. The document holds two images: the logo and the footer mark.
- **Validity.** `docx validate` reports `valid — 8 parts checked`. The render opened with no errors.

## Demerits

1. **Major: the Notes body was not replaced.** The output read view shows `p9` (8pt) "Payment due within 30 days. Wire details on request." followed by `p10` "Thank you for your business." The task said to replace the Notes body, and the old line is still present. The new text also sits in the template's 8pt spacer paragraph, so it renders at about two-thirds the size of the surrounding body text. The PNG shows a small line under "Notes" with the old full-size sentence below it. The agent noticed the 8pt problem itself in its frictions and chose not to fix it.
2. **Minor: the PNG fallback for the logo is a 1×1 transparent pixel.** `media/image3.png` is 67 bytes, 1×1. Word renders the SVG, but any consumer that relies on the raster fallback (older Word, some converters and previewers) will show a blank logo. docx-cli's `images list` and `read` report only this blank PNG hash.
3. **Minor (cosmetic): the logo part is in an unusual location.** It sits at the package root as `/media/image.svg` and `/media/image3.png`, outside `word/`. This is legal OPC and Word accepts it, but it is non-standard.

No layout defects were introduced. The page is still a single page, and the tables and footer are in place.

## Frictions (agent-reported) and assessment

- The VML Fallback text box could not be addressed. Every path the agent tried failed ("No textbox found ... Available children: textbox(1)"), so it fell back to `raw-set --xpath '//w:pict//w:txbxContent//w:t'`. The agent did keep Choice and Fallback in sync in the end, but by hand.
- `add --type row --from tr[4] --after tr[4]` placed the copy at tr[3], so the rows came out of order. The agent recovered by rewriting rows 3–5 rather than moving the row.
- Setting the logo with `set src=file.svg` auto-generated a 1×1 placeholder PNG fallback. The agent flagged that it could not verify the fallback, and the file confirms it is blank.
- `view text` hides table cell content, so the placeholder sweep needed extra queries.
- The agent spotted the 8pt Notes size but left it as it was.
- A zsh glob expanded an unquoted path (the agent's own error).

**Read:** the path was a slog in places. There were three tool-side dead ends: the Fallback text box path, the misplaced row clone, and the opaque SVG fallback. The agent worked through them carefully and the end state is mostly right. The one real miss, the Notes body, came from the agent misreading which paragraph was the "body", not from a tool limitation.
