# Review: contract-finalize (OfficeCLI, r1)

**Scenario:** contract-finalize. Legal review: accept/reject tracked changes, then reply to and resolve comments.
**Bucket:** editing (tracked-change finalize + comment threading)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: 3 output pages rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true**: text, revisions, and replies all read back. Resolved state was confirmed in the OOXML because the docx-cli reader does not surface it (see notes). |

## Rubric check
1. **Net 30 accepted.** p7 reads "Net 30 from the Company's acceptance" as clean text. The baseline showed an insertion plus "Deleted: Net 90 from the Company". PASS
2. **Non-compete accepted.** p18 reads "For one (1) year after the Agreement ends" as clean text. PASS
3. **Liability cap rejected.** p24 reads "one hundred dollars ($100)". The baseline accepted view showed "$100,000", and both halves of that change are gone. PASS
4. **§11 Personal Guarantee deletion rejected.** The "11. Personal Guarantee." heading and its body paragraph are present. Dispute Resolution is now §12, and the later sections are renumbered through §15, which matches the original draft. PASS
5. **Indemnification pPrChange accepted.** p22 has `line-spacing="1.5"` (`w:line="360"`), and no pPrChange is left. The render shows the 1.5 spacing with no change bar. PASS
6. **No pending changes.** `docx track-changes list` returns "No tracked changes", and document.xml contains 0 `w:ins`/`w:del`/`w:pPrChange`/`w:rPrChange` elements. PASS
7. **Every comment has a reply.** c4→c0, c5→c1, c6→c2 and c7→c3 each carry a `parentId`. PASS
8. **Replies are anchored.** Every reply has a `startBlockId` (p3, p12, p22, p30). Each reply also has its own `commentRangeStart`/`End`/`Reference` in document.xml, nested around the parent's range. PASS
9. **Read-view visibility.** `read --comments` shows c4 to c7 as `[^cN]` footnotes with `↳ c0` through `↳ c3` thread markers. PASS
10. **Auto-renewal resolved, and** 11. **arbitration resolved.** In commentsExtended.xml, paraId 97CB30E2 (c0) and EC14166F (c3) have `w15:done="true"`. Those paraIds match the comment paragraphs, and the part is wired by a commentsExtended relationship and a content-type override. PASS
12. **IP comment open, and** 13. **indemnification comment open.** c1 and c2 have `done="false"`. PASS

## Merits
- All five revision decisions are correct. Both halves of each paired text revision were handled, and the §11 paragraph-mark deletion was handled too.
- The rejected changes restore the original wording exactly: "$100", and the full §11 text with correct renumbering.
- The indemnification spacing change is accepted as live formatting. The 1.5-line spacing shows in the render.
- The replies are proper threaded replies. Each has a `paraIdParent` in commentsExtended, its own anchored range, and a Company author ("Nimbus Freight Systems"). Word renders each reply in a balloon beside its parent.
- The resolve flags are correct: done=true only on the auto-renewal and arbitration threads.
- Nothing else changed. The baseline-to-output read.md diff shows only p24 ($100) and the restored §11 with the downstream renumbering. The header, footer, table, signature block, and other text are all unchanged.

## Demerits
- (minor, not attributable to the output) `docx comments list` prints no `resolved` field for c0 or c3, although `w15:done="true"` is set correctly in the file. This looks like a docx-cli reader gap. One possibility is that the commentsExtended relationship uses an absolute Target (`/word/commentsExtended.xml`), which the reader may not resolve. This is a docx-cli bug and is not counted against OfficeCLI. It is still worth a follow-up in docx-cli.
- (minor, cosmetic) The c0 reply text says "Declined" even though the thread is marked resolved. The task permits any one-line position, so this is not a rubric failure.
- (note) Word's PDF output draws the resolved threads (JC1/NF2 and JC7/NF8) the same way as the open ones. This is normal for Word's print markup, and the file-level flags are correct.

## Frictions
The agent reported these frictions:
- **Major:** the ins/del halves of a revision are listed separately with no pairing key. Accepting one half leaves the document half-redlined and gives no error. The agent caught this only by re-running the query.
- Replying works through `add --type comment --prop parentId=N`, which is not discoverable from the comment help.
- The add-comment acknowledgement gives a path, not a commentId.
- The `help docx revision` output is long and leads with paraMark detail.
- `validate` has no baseline mode, so the 9 errors already in the original file look like the agent's errors.

It also hit two dead ends: half-accepting a pair, and dumping the raw document XML because the text view hides deleted text.

**Read on the path:** the agent got a correct result, but the path was a moderate slog. Revision pairing is a real trap in OfficeCLI's tool UX: a less careful agent would have left half-redlines. Reply threading had to be found by digging.
