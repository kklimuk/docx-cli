# Invoice data — fill out invoice.docx

Fill the attached `invoice.docx` with the real values below **and swap its logo**.
Keep the table layout and the formatting intact — only the text content and the logo
change.

The template has more placeholder *lines* than some sections need (e.g. the "from"
address spans two lines but the real address fits on one; each line item has a
"Description N" sub-line). **Fill the lines you use and remove the leftover
placeholder lines** so no template text is left behind — clear a line by giving it
empty text (`{ "at": "pN", "text": "" }` in an `edit --batch`, or `docx edit --at pN
--text ""`).

## Header / "from"

- Company: **Northwind Robotics, Inc.**
- Phone: **(415) 555-0137**
- Email: **billing@northwindrobotics.com**
- Address: **2120 Bryant Street, San Francisco, CA 94110**
- Invoice #: **NW-1042**
- Invoice date: **June 8, 2026**
- Due date: **July 8, 2026**

## Bill To / Ship To

- Customer: **Acme Health Systems, LLC**
- Customer email: **ap@acmehealth.com**
- Billing address: **88 Market Street, Suite 400, Chicago, IL 60603**
- Shipping address: same as billing.

## Line items

There are **four** line items, but the table only has rows for three. Fill the three
existing item rows, then **insert a new row** for the fourth (use the table tools —
don't type over a totals row), and fill it. After they're in, **set the line-items
table's column widths** so Description is the widest column — while keeping the Price
and Amount columns wide enough to show the dollar values (e.g. `$10,100.00`) on one
line without wrapping.

| Description | Quantity | Price | Amount |
| --- | --- | --- | --- |
| RX-7 sensor module | 4 | $1,250.00 | $5,000.00 |
| On-site integration (per day) | 3 | $900.00 | $2,700.00 |
| Extended support, 12 mo. | 1 | $1,800.00 | $1,800.00 |
| Calibration kit | 2 | $300.00 | $600.00 |

## Totals

- Subtotal: **$10,100.00**
- Discount: **-$500.00**
- Shipping: **$0.00**
- Tax total (8.75%): **$840.00**
- Other: **$0.00**
- **Total: $10,440.00**

## Notes

Replace the "Notes" body with: *Payment due within 30 days. Wire details on request.*

## Logo

Replace the placeholder logo in the **top-left corner** with the new company mark at
`assets/logo.svg`. Swap that one image only — don't add a second logo, and **don't
touch the small payment mark in the footer** (it must stay). Tip: list the document's
images first to find the right one to target.

## What "done" looks like

All `Item 1/2/3`, `$0.00`, `Customer name`, `Your Company Name`, etc. placeholders
are replaced with the values above — and leftover placeholder lines (the extra
address line, the "Description N" sub-lines) are removed, not left behind; the three
tables are still present and unbroken; the line-items columns are sized so Description
is widest and no dollar value wraps; the top-left logo is now the new mark from
`assets/logo.svg` (a green ink-blob); and the footer's payment mark is untouched (the
document still has two embedded images).
