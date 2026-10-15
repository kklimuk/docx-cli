# Review — contract-finalize (bucket: editing / review-finalize), docx-cli-r3

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true** (3 output pages rendered) |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

## Verification performed
- `docx track-changes list` returns `no tracked changes` (rubric 1–6).
- `docx comments list` shows four top-level comments (c0–c3), each with one reply (c4–c7). Each reply has a matching `parentId` and a non-empty `anchor.startBlockId` (p3, p12, p22, p30). `resolved` is true on c0 and c3 and false on c1 and c2 (rubric 7, 8, 10–13).
- `docx read --comments` shows the replies as `[^c4]`–`[^c7]` with `↳ c0`…`↳ c3` markers, and the resolved roots carry a ✓ (rubric 9).
- `docx validate` reports `valid — 10 parts checked`.
- Baseline vs output read.md diff: the liability cap goes from `$100,000` back to `one hundred dollars ($100)`, and `11. Personal Guarantee.` plus its guarantee paragraph (p28) are reinstated, which renumbers Dispute Resolution through Miscellaneous to 12–15. There are no other text changes. Net 30, one (1) year and line-spacing 1.5 on p22 were already the accepted-view text in the baseline and are now clean.
- Visual comparison: the baseline shows the redline balloons (Deleted: Net 90…, five (5) years, $100, the Personal Guarantee deletion). The output has no revision balloons and only comment and reply balloons. The header, footer, the maroon-header schedule table, the 1.5-spaced indemnification clause and the signature block are all unchanged.

## Merits
- Every rubric check passes. All 10 revisions were decided correctly: Net 30 and one (1) year are accepted, the $100 cap and the §11 Personal Guarantee are restored, and the indemnification spacing is accepted as live formatting.
- The agent used the batched review recipe: a dry-run, then one `track-changes apply --accept … --reject …`, then one `comments reply --batch` with `resolve:true` on c0 and c3.
- Replies are attributed to "Nimbus Freight Systems, Inc." and each states a one-line Company position.
- The restored §11 keeps its numbering, bold heading and body indent, and the later sections renumber cleanly.

## Demerits
- **minor (tool UX):** The read view and `track-changes list` disagree. `read --current` has no `[^tcN]` markers for formatting-only or paragraph-mark revisions (tc4, tc8), and the list does not group the pure deletion of §11 (tc7, tc8, tc9) under one handle. The agent had to assemble the handles by hand. The output is not affected.
- **minor (tool UX):** The `apply --dry-run` preview labels actions with XML verbs ("unwrap", "delete"), not "accept"/"reject". It even calls an accepted pPrChange "delete".
- **minor (tool UX):** `comments reply` silently defaults the author to "Reviewer". The agent worked around it with a per-line `author` key.

## Frictions (agent-reported)
- §11 is a pure deletion with no rev handle, so the agent had to find all three tcNs itself (major, per the agent).
- The default reply author is misleading (minor).
- The dry-run action labels are confusing (minor).
- The read-view footnote markers do not match the list handles (minor).

**Read:** the path was mostly smooth. The agent finished in a few batched calls with a dry-run safety check. The real slog was reconciling handles for the multi-part §11 deletion across two views that don't agree. It is worth giving a pure deletion block a single grouped handle and marking every tcN in the read view.
