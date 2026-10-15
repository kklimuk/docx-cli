# Review — invoice (docx-cli, run r2)

**Scenario:** invoice — Table editing + restructure + image replace (editing bucket)

## Verdict

| Dimension | Result |
| --- | --- |
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (Word render produced page-001.png) |
| Layout sound | **true** |
| Formatting preserved | **preserved** (one minor spacing deviation, see Demerits) |
| Survived write→read loop | **true** (fresh `docx read` is byte-identical to the harness's output read.md) |

## Merits

- **All placeholders filled with the right values.** Company/phone/email/address, Invoice # NW-1042, June 8, 2026 / July 8, 2026, Bill To + Ship To (Acme Health Systems, LLC; ap@acmehealth.com; 88 Market Street, Suite 400, Chicago, IL 60603), all totals ($10,100.00 / -$500.00 / $0.00 / $840.00 / $0.00 / $10,440.00), and the Notes text match the brief exactly.
- **The placeholder sweep is clean.** `docx find` for `Item`, `Description 1`, `Customer name`, `Your`, `Jan, 1`, `TEMPLATE`, and `DO NOT SEND` all return no matches.
- **Leftover lines are gone.** The "Your City, State Zip" line was removed (the address fits on t0:r0c0:p3), the second customer-name line moved up into place, and all three "Description N" sub-lines are gone. The item cells now hold one paragraph each.
- **Fourth row inserted, not overwritten.** t2 now has the header, 4 item rows (r1–r4, with Calibration kit / 2 / $300.00 / $600.00 in r4), and all six totals rows (r5–r10). The gridSpan="2" Quantity cells and the totals-box borders are intact.
- **Column widths are right.** The widths are `3,0.5,1.5,1.6,1.2in`. In item rows, Description is 3in, Quantity (spanning two grid columns) is 2in, Price is 1.6in, and Amount is 1.2in, so Description is the widest. In the render, `$10,100.00`, `$10,440.00` and `Tax total (8.75%)` each fit on one line.
- **Table polish is in place.** `repeat-header="r0"` is now on the docx:table hint. The header cells carry `shading="F2F2F2"`, and every Price and Amount cell has `halign="right"`. The shading and right alignment were already in the template, which is acceptable under the rubric.
- **Logo swapped.** img0 is now `image/svg+xml` (hash 7d680d…, from assets/logo.svg) at the same 1.05x1.06in size and float position. The package holds exactly two media parts: image1.svg (the new logo) and image2.png (the footer PayPal mark). footer1.xml.rels still has the image and hyperlink relationships, and the footer mark shows in the render.
- **Stamp box rewritten in place.** `find "Thank you for your business!"` → `tbx0:p0:0-28`. The box keeps its red C00000 bold, centered text and its cream fill with red border at the same position. The string appears twice in document.xml, which means the mc:Choice and mc:Fallback copies are in sync.
- Baseline and output renders match apart from the intended content changes: the header, the INVOICE title, the footer and the table borders are unchanged.

## Demerits

- **minor: two spacer paragraphs removed beyond the placeholder lines.** The agent deleted the empty spacer paragraph between "Your Company Name" and "Phone:" (old t0:r0c0:p1) and the empty spacers under "Bill To:" / "Ship To:" (old t1:r0cX:p1). These were template spacing, not placeholder text, so the header block and Bill To block now sit tighter than the template design (visible in the before/after render: "Phone:" now directly under the company name). The result still looks clean, and the task's "remove leftover placeholder lines" is ambiguous here.
- **minor: the totals label changed from "Tax total" to "Tax total (8.75%)".** This follows the brief's wording, so it is defensible, not a defect of substance.
- **minor (tool): `images list` reports only img0.** It doesn't list the footer image, so the CLI can't confirm the rubric's "exactly two images" check. I verified it from the package parts and the footer rels instead.

## Frictions (agent-reported) and assessment

- **major:** The totals labels live in the Price grid column because of the gridSpan. The first `tables set-widths` (Description 3.6in) made "Tax total (8.75%)" wrap with no warning, and the agent needed a render to find it and a second width pass to fix it. This is a real tool gap: set-widths doesn't map visible columns to grid columns and doesn't warn about wraps.
- **minor:** The Markdown read omits empty paragraphs, so the spacer p9 and blank cell lines only showed up in `--ast`. Empty cell paragraphs inside tables do show as bare locators, as the baseline read shows, but body-level blanks do not.
- **minor:** Re-applying halign right and the F2F2F2 shading, which were already set, printed the same "changed" ack. A "no change" signal would help.
- **minor:** `images list` omits header and footer images.
- **minor:** Verifying the wrap took three renders.
- **dead end:** A `$300.00` argument is shell-expanded inside double quotes, so the agent needed single quotes. It handled this.

**Assessment:** a moderate slog that reached a fully correct result. The agent batched the text edits (one edit batch of 31 changes plus one delete batch), which is the efficient path. The main time sink was the table width and wrap interaction, which the CLI can't surface without a render.
