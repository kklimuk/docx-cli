# Company's-counsel finalization

You represent **Nimbus Freight Systems, Inc.** (the "Company"). Opposing counsel
returned `contract-redlined.docx` with tracked changes and comments. Finalize it.

First, see what's pending:
- `docx track-changes list FILE` — the tracked revisions (each has a `tcN` id)
- `docx comments list FILE` — the comments (each has a `cN` id)

## Tracked changes — accept some, reject others

**Accept** these revisions (the Company agrees):
- **Payment terms** changed from Net 90 → **Net 30**.
- **Non-compete** shortened from five (5) years → **one (1) year**.

**Reject** these revisions (restore the Company's original draft):
- The **liability-cap** change — the Contractor raised the Company's **$100** cap;
  keep the original $100.
- The **deletion of §11 (Personal Guarantee)** — the Company wants the guarantee, so
  reject the deletion (the section stays).

Tip: `track-changes list` ids (`tcN`) renumber after each accept/reject — re-list
between operations, or address one change at a time with `--at tcN`. (Each text
revision shows as a `del` + an `ins` pair; act on both halves of a change.)

## Comments — reply to all, resolve the addressed ones

- **Reply** to each of the four comments with a one-line Company position (e.g.
  "Agreed — revised." or "Declined — standard term.").
- **Resolve** the two the Company has fully addressed: the **auto-renewal** comment
  and the **arbitration-fees** comment. Leave the **IP** and **indemnification**
  comments open (still under negotiation).

## What "done" looks like

The accepted revisions are clean text, the rejected ones are back to the original,
every comment has a reply, and the auto-renewal + arbitration comments are resolved.
`docx track-changes list` and `docx comments list` reflect this.
