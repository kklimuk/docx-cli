# Review: contract-finalize (Anthropic docx skill, run r3)

**Scenario:** contract-finalize. Legal review: accept or reject tracked changes, then reply to and resolve comments. **Bucket:** editing (review workflow). **Arm:** competitor (Anthropic docx skill, haiku agent).

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: 3 output pages rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

All 13 rubric conditions pass.

## Evidence against the rubric

1. **Net 90 → Net 30 accepted.** Page 1, §3 shows "All payments are Net 30 from the Company's acceptance" as clean text, with no redline and no "Deleted:" balloon (the baseline had one).
2. **Non-compete accepted.** Page 2, §6 shows "For one (1) year after the Agreement ends" as clean text.
3. **Liability cap rejected.** Page 2, §9 reads "one hundred dollars ($100)". The "$100,000" insertion is gone. The read-view diff shows only this text changing in p24.
4. **§11 Personal Guarantee deletion rejected.** On page 2 the "11. Personal Guarantee." heading and its body are back. Later sections renumber to 12 to 15, which matches the original pre-deletion numbering.
5. **Indemnification pPrChange accepted.** The read view shows p22 with `line-spacing="1.5"` and no pending change. The render shows 1.5-line spacing.
6. **No pending changes.** `docx track-changes list` returns "No tracked changes".
7. **All four comments have replies.** Replies c4 to c7 are authored by "Nimbus Freight Systems, Inc." and have parentId c0, c1, c2 and c3.
8. **Reply anchors.** Every reply has a non-empty `anchor.startBlockId` (p3, p12, p22, p30), the same as its parent.
9. **Read-view visibility.** `docx read --comments` shows `[^c4]` to `[^c7]` with `↳ c0` to `↳ c3` thread markers.
10. **Auto-renewal resolved.** c0 and its reply c4 have `resolved: true`. Word greys out the NF2 reply as done.
11. **Arbitration fees resolved.** c3 and its reply c7 have `resolved: true`. Word greys out NF8.
12. **IP comment open.** c1 and c5 have `resolved: false`.
13. **Indemnification comment open.** c2 and c6 have `resolved: false`.

### Before and after comparison
- **Visual:** the header, footer, title, table (including its colour fill), list numbering style, signature block and page count (3) all match the baseline. The only differences are the intended ones: redline balloons gone, the guarantee section restored, reply balloons added, and resolved threads greyed out.
- **Textual:** the baseline-to-output read-view diff contains only the liability-cap text (back to $100), the restored §11 heading and body, and the renumbering of §12 to §15. Accept-view text elsewhere is unchanged.

## Merits
- Every accept and reject decision is correct, including the hardest one: un-deleting the §11 paragraph-mark deletion so the list item and its body came back with the correct numbering.
- The pPrChange was accepted correctly, keeping line=360 as live formatting.
- Replies are properly threaded. They use their own comment ids, nested range markers sharing the parent's anchor, and commentsExtended parent links. Word renders them as replies under the right parents.
- Resolve state is correct per thread. The two closed threads are resolved and the two open ones are not.
- Reply wording is sensible and professional ("Declined — standard term; arbitration fees stay as drafted.").
- The agent flagged that the payment table cell still says "Net 90 after final acceptance". That cell was untracked and out of scope, and flagging it rather than silently changing it was a good call.

## Demerits
- **Minor:** the reply comments in comments.xml are pretty-printed, with whitespace between elements inside `<w:comment>`. Word ignores it and renders the replies cleanly. However, docx-cli `comments list` shows the reply `text` with leading and trailing newlines and spaces, so a downstream reader sees noisy text. This is a well-formedness nit, not a visible defect.
- **Minor (not graded):** the payment table still says "Net 90 after final acceptance", which conflicts with the accepted Net 30 clause. This is out of the task's scope and the agent reported it.

## Frictions
The agent reported:
- (major) There is no selective accept or reject helper. `accept_changes.py` is accept-all only, so the agent hand-wrote regex edits on `w:ins`, `w:del`, `pPrChange` and the paragraph-mark `w:del`.
- (major) `comment.py` has no resolve option. The agent hand-wrote `w15:commentEx` entries with `done="1"`, including root entries that did not exist in the source.
- (major) `validate.py --author` reports correct rejections as "Document text doesn't match", so a correct result looks like a failure.
- (minor) `comment.py` fails under `python3 -I` (ModuleNotFoundError: office).
- (minor) Nesting the reply's range markers and reference runs is manual.
- (minor) The source document's paraIds are ≥ 0x80000000, so the validator fails on the pristine file.

**Assessment:** the output is correct, but the path was a slog of raw-XML surgery. The skill has no primitive for selective accept/reject or for resolving a comment, and its validator gave misleading failures on correct work. That the result came out right reflects careful agent work, not tool support. This is a significant UX demerit for the skill, but it does not affect task success.
