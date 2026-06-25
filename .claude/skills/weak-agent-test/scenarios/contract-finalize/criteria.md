# Grading rubric — contract-finalize (JUDGE ONLY)

## Pass conditions

### Tracked changes

1. **Net 90 → Net 30** accepted: the payment-terms paragraph contains "Net 30" as clean
   text; no pending deletion or insertion for that phrase remains in
   `docx track-changes list FILE`.
2. **Non-compete 5 yr → 1 yr** accepted: the non-compete clause reads "one (1) year" as
   clean text; no pending revision for that phrase.
3. **Liability-cap change rejected**: the liability-cap paragraph still reads "$100"
   (the Company's original wording); the Contractor's proposed increase is gone (no
   pending insertion for the raised cap, and the Company's "$100" deletion is
   un-deleted).
4. **§11 Personal Guarantee deletion rejected**: the Personal Guarantee section is
   present in the document body (not removed); no pending deletion covers it.
5. **Indemnification pPrChange accepted**: `docx track-changes list FILE` returns no
   entry with `type: "paragraph-format"` on the indemnification clause. The 1.5-line
   spacing (or whatever spacing the Contractor proposed) is applied as live formatting
   with no pending-change marker.
6. **No other pending changes**: after the five operations above, `docx track-changes
   list FILE` returns an empty list (no stray revisions left).

### Comments

7. **All four comments have replies**: `docx comments list FILE` shows four top-level
   comments, each with at least one reply entry whose `parentId` matches the parent
   comment's `id`.
8. **Reply regression guard — parentId + non-empty anchor**: every reply object in
   `docx comments list FILE` satisfies BOTH:
   - `parentId` is present and non-null.
   - `anchor.startBlockId` is a non-empty string (not `""`).
   (This guards against the issue where `comments reply` wrote to `comments.xml` but
   never anchored the reply, causing Word to silently delete it on next save.)
9. **Reply regression guard — read-view visibility**: `docx read FILE --comments`
   shows every reply as its own `[^cN]` footnote with a `↳ cN` thread marker. A reply
   absent from the default read view fails the write→read invariant.
10. **auto-renewal comment resolved**: `docx comments list FILE` shows the auto-renewal
    comment with `resolved: true` (or equivalent resolved status).
11. **arbitration-fees comment resolved**: same — marked resolved.
12. **IP comment still open**: the IP-ownership comment has `resolved: false` (or no
    resolved flag).
13. **indemnification comment still open**: the indemnification comment has
    `resolved: false` (or no resolved flag).

## How to verify

```
# Pending tracked changes (should be empty on a passing run):
docx track-changes list contract-redlined.docx

# Comments with replies and resolved status:
docx comments list contract-redlined.docx

# Full read view — replies must appear as [^cN] footnotes with ↳ cN markers:
docx read contract-redlined.docx --comments
```
