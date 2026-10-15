# Review — invoice (officecli-r3) — bucket: editing (table edit + restructure + image replace + text box)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes. Rendered: true, 1 output page PNG viewed.
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes. A fresh `docx read` is byte-identical to the output read view, and `docx validate` reports `valid — 8 parts checked`.

## Merits
- **All placeholders filled with the correct values.** That covers the company, phone, email, the one-line address, NW-1042, June 8, 2026 / July 8, 2026, Bill To and Ship To (Acme Health Systems, LLC / ap@acmehealth.com / 88 Market Street, Suite 400, Chicago, IL 60603), and the Notes text. A grep of every word/*.xml part finds no `Item N`, `Description N`, `Customer name`, `Your Company` or `Your City`.
- **Leftover placeholder lines removed.** The `Your City, State Zip` paragraph (t0:r0c0:p5) is gone, and the gray `Description N` sub-line is gone from every item cell. No empty paragraphs are left in the item cells.
- **Line-items table now has 4 data rows**, with the totals rows (Subtotal…Total) intact below them as r5–r10. The values are $10,100.00 / -$500.00 / $0.00 / $840.00 / $0.00 / $10,440.00. Calibration kit sits in its own row (r4); no totals row was overwritten.
- **Table structure kept.** The gridSpan merges are unchanged: header Description spans 2 columns, and data Quantity spans 2 columns. Borders are unchanged.
- **Widths:** the tblGrid is 5200/450/1300/2300/1974 twips, and every `tcW` was resynced (5650/1300/2300/1974 in the header, 5200/1750/2300/1974 in the data rows). Description is clearly the widest column. In the Word render, $10,100.00 and $10,440.00 fit on one line.
- **Header row:** it has `shading="F2F2F2"` on all 4 cells, and `<w:tblHeader/>` sits on r0, so `read` shows `repeat-header="r0"`.
- **Price/Amount right-aligned:** `halign="right"` is on every Price and Amount cell, in the header, the data rows and the totals.
- **Logo replaced.** It is a single drawing with docPr "Northwind Robotics logo" and alt text. It uses an `asvg:svgBlip` that points at the SVG, plus a PNG fallback blip. The size is 972000 EMU square (1.06 in, effectively the original 1.05x1.06 in), and it is still top-left, floating, with wrap through. The footer PayPal mark (`word/media/image2.png`, footer1 rId2) is untouched and still renders. There are 2 embedded images: the logo and the footer mark.
- **Stamp box kept and reworded.** `tbx0` is still anchored on the Notes paragraph and reads `Thank you for your business!`. It stays bold, red C00000 and centered, and the fill, border and position are unchanged in the render. Both the mc:Choice and mc:Fallback copies carry the new text (2 occurrences, 2 Fallbacks), and there are 0 `TEMPLATE`/`DO NOT SEND` strings anywhere. `find "Thank you for your business!"` → `tbx0:p0:0-28`.
- **Before/after diff is clean.** The textual diff of the read views shows only the intended changes, and the visual comparison shows the header, footer, INVOICE title and table borders preserved.

## Demerits
- **(minor) Logo drawing re-created rather than swapped in place.** The original custom wrap polygon was regenerated, and the position offset changed a little (the logo sits about 15 px further right/lower in the render). Nobody would notice this visually.
- **(minor) Unconventional part placement.** The new media lives at the package root (`/media/image.svg`, `/media/image3.png`) instead of under `word/media/`. This is valid OPC and Word renders it, but it is atypical.
- **(minor) The narrow 450-twip grid column was kept.** Header "Quantity" (1300) and data Quantity (450+1300) cover different grid columns. That was already true in the template, but the renderer's header "Quantity" now sits slightly offset from the digits below it. This is cosmetic only.

## Frictions (agent-reported)
- `add --type row --from … --after/--before` ignored the position flags; the clone always landed at index 3. The agent recovered with `move --after` and then rewrote the text. That is a real tool bug: position flags were silently ignored.
- There is no image src/replace verb. The agent had to remove and re-add the picture, re-specifying anchor, wrap, position and size, and the wrap polygon was lost.
- Setting colWidths updated tblGrid but not the per-cell `tcW`; the agent resynced them by hand.
- `help docx row` is unknown, because the element is named `table-row`.
- `view issues` flags intentional spacer paragraphs as empty, and the screenshot renderer's font substitution varied between renders.

**Read:** the result is correct, but the path was a moderate slog: three major tool gaps, all of which the agent worked around. These are tool UX/bug problems, not output defects.
