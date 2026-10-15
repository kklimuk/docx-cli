# Review — invoice (table editing + restructure + image replace) — arm: Anthropic docx skill (haiku, r2)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (1-page output PNG produced)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true (all intended content present in the render and in `docx read`)

## Merits
- Every placeholder family replaced with the right values: header (Northwind Robotics, Inc. / (415) 555-0137 / billing@northwindrobotics.com / 2120 Bryant Street, San Francisco, CA 94110), invoice # NW-1042, June 8, 2026 / July 8, 2026, Bill To / Ship To (Acme Health Systems, LLC, ap@acmehealth.com, 88 Market Street, Suite 400, Chicago, IL 60603), and the Notes text "Payment due within 30 days. Wire details on request."
- Leftover lines removed: the "Your City, State Zip" line is gone (t0:r0c0 now ends at p4), and all three "Description N" sub-lines are gone. The item cells now hold a single paragraph each, with no empty paragraphs left behind.
- A 4th item row (Calibration kit / 2 / $300.00 / $600.00) was inserted as a new row (t2 now has r0–r10). The totals rows are intact below it: Subtotal $10,100.00, Discount -$500.00, Shipping $0.00, Tax $840.00, Other $0.00, Total $10,440.00.
- The header row carries `<w:tblHeader/>` (`repeat-header="r0"` in read). The header shading F2F2F2 and the right alignment of Price/Amount were already present in the template, and both are preserved (`halign="right"` on every Price/Amount cell).
- Description is the widest column and no dollar value wraps ($10,100.00 and $10,440.00 each sit on one line in the render).
- Stamp text box kept with its look unchanged (red border, cream fill, position, bold red centered run). It now reads "Thank you for your business!" in both the mc:Choice and mc:Fallback copies (2 occurrences in document.xml), and TEMPLATE / DO NOT SEND appear nowhere.
- Logo: the top-left image bytes were replaced with a 600x600 PNG rasterized from assets/logo.svg; the render shows the green ink mark. The package still has exactly two media files, and the footer PayPal mark is intact in the render.
- The file passes `docx validate` (8 parts checked). Tables, borders and fonts match the baseline.

## Demerits
- **minor** — The agent changed the label "Tax total" to "Tax total (8.75%)". In the render that label wraps to two lines in the narrow label column. This is a cosmetic wrap of a label, not a dollar value, and the task literally names the line "Tax total (8.75%)", so it is defensible. It is still a small visual regression against the baseline's single-line row.
- **minor** — The logo is a rasterized PNG rather than an embedded SVG (svgBlip). It is visually correct; it is a fidelity note, not a failure.
- (Not a defect) Column widths were left unchanged because Description was already the widest. The rubric's width conditions are met in the render.

## Before/after comparison
- Visual: only content changed: the logo, text, a 4th row, the stamp wording and the Notes text. Everything else matches the baseline: the INVOICE title, table borders, header shading, footer and stamp box styling. The page height shrank slightly because the sub-lines were removed. Nothing overflowed and the document stays on 1 page.
- Textual diff of the read views: exactly the expected lines changed (image hash, t0/t1 fills, `repeat-header="r0"` added, item rows refilled plus one new row, totals, Notes, stamp). No unintended changes.

## Frictions
- The agent reports that SKILL.md has no table guidance: row cloning, tblHeader, cell jc and column widths all had to be worked out from raw OOXML. It counts this as a blocker-level gap in the skill.
- It reports no warning that text boxes carry two copies of the story (mc:Choice and mc:Fallback). The agent handled this correctly anyway.
- It reports no image-replace or SVG recipe. It used `rsvg-convert`, which is outside the skill's listed dependencies.
- It had to check the formatting that was already in place by reading raw XML, and had to check wrapping by rendering with soffice and pdftoppm.
- Dead end: it cloned a row after filling it, which broke the clone; it fixed this by cloning before filling.
- Read: the result is correct, but the path was a slog of hand-written lxml scripts with no tool support. A strong UX demerit for the skill; the output itself is clean.
