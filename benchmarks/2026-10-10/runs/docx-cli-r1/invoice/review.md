# Review — invoice (Table editing + restructure + image replace)

**Run:** docx-cli-r1 (haiku)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** — one-page output PNG rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** — `docx read` output is byte-identical to the harness read view |

## Merits
- **Placeholders gone, values correct.** Header: Northwind Robotics, Inc. / Phone: (415) 555-0137 / billing@northwindrobotics.com / 2120 Bryant Street, San Francisco, CA 94110. Invoice # NW-1042, June 8, 2026, July 8, 2026. Bill To and Ship To: Acme Health Systems, LLC, ap@acmehealth.com, 88 Market Street, Suite 400, Chicago, IL 60603. Totals: $10,100.00 / -$500.00 / $0.00 / $840.00 / $0.00 / $10,440.00. `find` for `TEMPLATE`, `DO NOT SEND`, `Item `, `Description 1`, `Customer name` and `Your ` all return no matches.
- **Leftover lines removed.** The 6th address line (`Your City, State Zip`) is gone, and so are all three `Description N` sub-lines. The item cells now hold one paragraph each. The empty spacer paragraphs that remain (t0:r0c0:p1, t1 p1, p14) were already in the baseline.
- **Four line items.** Calibration kit / 2 / $300.00 / $600.00 is a new row t2:r4. The totals block (r5–r10) is intact, with its borders.
- **Widths.** The grid is 3.6 / 0.5 / 1.0 / 1.4 / 1.292in, which makes Description clearly the widest. In the render, $10,100.00 and $10,440.00 each fit on one line.
- **Table polish.** The header row still has F2F2F2 shading, which was already in the template. Price and Amount are `halign="right"`, also already set. `repeat-header="r0"` is newly set.
- **Stamp box.** It was rewritten in place: `tbx0:p0` contains "Thank you for your business!" and keeps its bold red #C00000 centered formatting. The box's fill, border and position look the same in the render. The new text is in both the Choice and the Fallback copies (2 occurrences in document.xml).
- **Logo.** img0 is now the SVG mark: a green circle in the render, with a new hash and the same 1.05x1.06in floating geometry. The footer PayPal mark (footer rel → media/image2.png) is untouched, so two images are in use.
- **Notes and validation.** The Notes body is replaced. `docx validate` reports valid.

## Demerits
- **minor — the SVG logo has no raster fallback.** rId6 → `media/image1.svg` is embedded directly, and document.xml has no `svgBlip` extension. Word 365 rendered it fine, but older Word versions may not. The agent flagged this itself; the tool should emit a PNG fallback or warn.
- **minor — the tax label was relabeled** from "Tax total" to "Tax total (8.75%)". This is a reasonable reading of the brief, which writes it that way, and does no harm.

No unintended changes appear in the baseline/output diff beyond the requested fills, the removed lines, the new row, the widths, repeat-header, the stamp text and the logo.

## Frictions
These are the agent's reported frictions:
- `tables set-widths --widths` doesn't accept `in`/`pt` units, so the agent converted to twips by hand.
- `$` in `tables insert-row --cells "…"` was shell-expanded inside double quotes, and the help doesn't warn about it.
- `docx tables --help` exits 1.
- `images list` omits header/footer images, so the "two images" check can't be made from that noun (verified: it lists only img0).
- SVG embed gives no fallback warning.
- `tables format` doesn't say when shading or alignment already match.

**Read:** the path was smooth, with one edit batch plus a few table verbs and no real dead ends. The frictions are paper cuts worth fixing, especially units in `set-widths`, the exit code on `--help`, and footer images missing from `images list`.
