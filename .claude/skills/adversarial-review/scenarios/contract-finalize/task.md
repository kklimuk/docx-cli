# Legal review: accept/reject revisions + resolve comments

## Task

The Contractor's counsel returned this redlined draft (`contract-redlined.docx`) with
tracked changes (text redlines **and** a paragraph-formatting change) and margin
comments. Act as the **Company's** counsel finalizing it: **accept** the revisions the
Company will agree to, **reject** the ones it won't, **reply** to every comment with
the Company's position, and **resolve** the comments you've addressed. Edit the file
in place.

## Resolution criteria

In the finished file: the accepted revisions are applied as clean text (no redline
left), and the rejected ones are reverted to the original wording; the tracked
**paragraph-formatting change** (the change on the indemnification clause) is
accepted (no longer pending — the 1.5 line spacing stays, the marker is gone); every
comment has a Company reply; and the comments the Company addressed are marked
resolved. Verify via `docx track-changes list` (the accepted/rejected revisions and
the pPrChange no longer appear as pending) and `docx comments list` (replies present,
the addressed comments resolved).

**Replies must be real, Word-visible comments — not sidecar orphans** (regression
guard for issue #1, where `comments reply` wrote the reply to `comments.xml` but never
anchored it, so Word silently deleted it on the next save). Check all three of these
on the finished file:

1. Every reply in `docx comments list` has a `parentId` AND a non-empty anchor
   (`anchor.startBlockId` is a real locator, not `""`).
2. `docx read FILE --comments` shows every reply as its own `[^cN]` footnote with the
   `↳ cN` thread marker — a reply missing from the default read view fails the
   write→read invariant.
