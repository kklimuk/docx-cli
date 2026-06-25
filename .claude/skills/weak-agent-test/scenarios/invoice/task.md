# Update our invoice template with real data and a new logo

We have an invoice template (`invoice.docx`) that needs to be filled out with the real deal data before we send it to the client. Fill it in with the values below and swap its logo. Keep the table layout and the formatting intact — only the text content and the logo change.

One thing to watch for: the template has more placeholder lines than some sections need (for example, the "from" address spans two lines but the real address fits on one; each line item has a sub-line below it). Fill the lines you use and remove the leftover placeholder lines so no template text is left behind. A half-filled form with empty placeholder lines dangling is not done.

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

There are **four** line items, but the table only has rows for three. Fill the three existing item rows, then add a new row for the fourth one (without overwriting the totals row) and fill it. After they're all in, make the Description column the widest — while keeping the Price and Amount columns wide enough to show dollar values like `$10,100.00` on one line without wrapping.

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

Replace the placeholder logo in the **top-left corner** with the new company mark at `assets/logo.svg`. Swap that one image only — don't add a second logo, and **don't touch the small payment mark in the footer** (it must stay). The document should still have two embedded images when you're done (the new logo + the footer mark).

## What done looks like

All `Item 1/2/3`, `$0.00`, `Customer name`, `Your Company Name`, and similar placeholder text is replaced with the values above. Leftover placeholder lines (the extra address line, the per-item sub-lines) are removed, not left dangling. The three tables are still present and unbroken. The line-items table has four filled rows, the Description column is the widest, and no dollar value wraps to a second line. The top-left logo is the new mark from `assets/logo.svg`, and the footer's payment mark is untouched.
