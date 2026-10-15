# Review — invoice (OfficeCLI, run r2)

**Scenario:** invoice — Table editing + restructure + image replace
**Bucket:** editing (OfficeCLI comparison)

## Verdict
| Dimension | Result |
| --- | --- |
| Task success | **success** |
| Renders correctly (Word opened it) | **true** — Word produced `renders/output/page-001.png` (the "close exited 1" note is an environment message: no resident session, disk copy current) |
| Layout sound | **true** — single page, tables intact, no wrapping dollar values, stamp box in place |
| Formatting preserved | **preserved** |
| Survived read loop | **true** — a live `docx read` is identical to `renders/output/read.md` |

## Merits
- Every placeholder is filled with the requested values. The header shows Northwind Robotics, Inc. / (415) 555-0137 / billing@northwindrobotics.com / 2120 Bryant Street, San Francisco, CA 94110. The invoice block shows NW-1042 / June 8, 2026 / July 8, 2026. Bill To shows Acme Health Systems, LLC / ap@acmehealth.com / 88 Market Street, Suite 400, Chicago, IL 60603, and Ship To has the same name and address.
- Leftover placeholder lines were removed, not blanked. "Your City, State Zip" (`t0:r0c0:p5`) is gone, and the "Description 1/2/3" sub-lines are gone from every item cell. No empty paragraphs remain in the item cells.
- The line-items table now has 4 data rows. Calibration kit / 2 / $300.00 / $600.00 is a new row (`t2:r4`) above Subtotal, so the totals rows were not typed over. The table went from 10 rows to 11, and the borders on the new row match the other item rows.
- Totals: $10,100.00 / -$500.00 / $0.00 / "Tax total (8.75%)" $840.00 / $0.00 / $10,440.00.
- Notes body reads "Payment due within 30 days. Wire details on request." The stamp's anchor run was kept, so the text box is still anchored on `p10`.
- Stamp box: `docx find "DO NOT SEND"` returns no matches, and `docx find "Thank you for your business!"` returns `tbx0:p0:0-28`. The text appears twice in document.xml, once in the Choice copy and once in the Fallback copy. Bold, red C00000 text and centred alignment are kept. Fill, border and position look the same in the render.
- Column widths: the grid is 4200/1000/1900/2000/2000 twips, and every `tcW` agrees with it, including the gridSpan cells (header Description = 5200, item Quantity = 2900). Description is the widest column in both the header row and the item rows. Price and Amount are 1.389in each, and $10,100.00 / $10,440.00 fit on one line in the render.
- Header row has F2F2F2 shading and Quantity/Price/Amount are right-aligned. Both were already in the template and were kept. `tblHeader` was added, and `read` shows `repeat-header="r0"`.
- Logo: the old placeholder image is gone. `media/image.svg` is byte-identical to `assets/logo.svg` and is referenced through `asvg:svgBlip`. Word renders it as the green mark in the top-left. The footer PayPal mark (`word/media/image2.png`, rId2 in footer1.xml.rels) and the hyperlink are untouched. The package holds two images: the logo and the footer mark.
- `docx validate` reports "valid — 8 parts checked".

## Demerits
- **minor** — The logo's raster fallback is a 1×1 transparent PNG (`media/image3.png`, 67 bytes), referenced by `a:blip r:embed`. Word uses the SVG and renders correctly. A consumer that ignores `svgBlip` (older renderers and some converters) will show a blank logo. `docx images list` and `read` also report the blank PNG's hash (`2b1da20a…png`) as the logo, not the SVG.
- **cosmetic, not scored** — The new parts were written at the package root (`/media/image.svg`, `/media/image3.png`), not under `word/media/`. This is legal OPC and Word accepts it, but it is unconventional.
- **cosmetic, not scored** — The Description item column got narrower than in the template (3.127in → 2.917in). It is still the widest column, so this meets the requirement.

## Frictions (agent-reported)
- `add --from <row> --after <row>` ignored the anchor. It reported "Copied to tr[3]" but placed the clone above Item 2, and a `move --after` was needed to fix the order. This is a real OfficeCLI bug: a wrong position reported with no error.
- `set --prop colWidths` changed `tblGrid` but left the `tcW` values on spanned cells stale. The agent set 11 cell widths by hand. Second OfficeCLI bug: a silent inconsistency.
- When the logo src was set to an SVG, the tool generated a 1×1 transparent PNG fallback, and `set` has no way to supply a real fallback. This is the root of the minor demerit above.
- In resident mode the on-disk zip stays stale until close. One zsh glob misfire on an unquoted `/body/tbl[3]` path. The `issues` view flags template spacers as empty paragraphs.
- The agent noticed by itself that a whole-paragraph text set on Notes would have destroyed the stamp anchor, and used `--find` instead. Good judgement, but the tool gives no guard against this.

**Read:** the result is correct and clean, but getting there was a moderate slog. Two silent-wrong-result tool bugs (row clone position, colWidths/tcW) each cost extra correction passes, and only the agent's careful verification caught them.
