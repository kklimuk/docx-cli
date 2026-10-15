# Review — contract-finalize (OfficeCLI, run r3)

**Scenario:** contract-finalize — Legal review: accept/reject tracked changes, then reply to and resolve comments
**Bucket:** editing (review workflow)

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** |
| Renders correctly (Word opened it) | **true**: Word rendered 3 output pages |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** (one docx-cli reader gap, described below; not an OfficeCLI defect) |

## Rubric check

1. **Net 90 → Net 30 accepted.** PASS. p7 reads "Net 30 from the Company's acceptance" as clean text. `document.xml` contains 0 `w:ins`/`w:del`/`w:pPrChange`/`w:rPrChange` elements. The table row "Net 90 after final acceptance" was never part of the redline and is the same in the baseline.
2. **Non-compete accepted.** PASS. p18 reads "For one (1) year after the Agreement ends" as clean text.
3. **Liability cap rejected.** PASS. p24 reads "one hundred dollars ($100)". The "$100,000" insertion is gone. The read-view diff against the baseline shows this as the only change in the paragraph.
4. **§11 Personal Guarantee deletion rejected.** PASS. The heading (p27) and body (p28) are back. The numbering runs 11 Personal Guarantee, then 12 Dispute Resolution through 15 Miscellaneous, and the Word render shows the same order.
5. **Indemnification pPrChange accepted.** PASS. p22 has live `<w:spacing w:after="80" w:line="360" w:lineRule="auto"/>`, so 1.5 spacing shows in the render. No pPrChange remains.
6. **No other pending changes.** PASS. `docx track-changes list` prints "No tracked changes". The output render has no redline markup or Deleted balloons.
7. **All four comments have replies.** PASS. c5→c1, c6→c2, c7→c3 and c8→c0 each carry `parentId`. Word shows each NFS reply threaded under its JC comment.
8. **Reply anchors.** PASS. Every reply has a non-empty `anchor.startBlockId` (p3/p12/p22/p30). Each reply has its own `commentRangeStart`/`End`/`Reference` in `document.xml`, nested around the parent's range.
9. **Read-view visibility.** PASS. `docx read --comments` shows every reply as its own `[^cN]` footnote with a `↳ cN` marker.
10. **Auto-renewal resolved.** PASS in the file. `commentsExtended.xml` has `w15:paraId="97CB30E2" w15:done="true"`, and c0's paragraph has paraId 97CB30E2.
11. **Arbitration-fees resolved.** PASS in the file. `w15:paraId="EC14166F" w15:done="true"` (c3).
12. **IP comment open.** PASS. `81F00F8C` has `done="false"`.
13. **Indemnification comment open.** PASS. `F6B5A501` has `done="false"`.

**Note on 10–11:** `docx comments list` does not show `resolved: true` for c0 and c3. The cause is in docx-cli: `#readExtended` in `src/core/ast/document/comments.tsx` accepts only `"1"` and `"0"` for `w15:done`. OfficeCLI wrote `"true"`, which is a valid ST_OnOff value. The resolved state is correct in the file, so this is a docx-cli bug, not an OfficeCLI failure (per the comparison instructions).

## Merits

- All five revision operations are correct, with zero leftover revision markup in `document.xml`.
- The rejected §11 deletion restored both the heading paragraph mark and the body. Numbering re-flowed correctly to 15 sections.
- The pPrChange accept kept the Contractor's 1.5-line spacing as live formatting.
- Replies are properly threaded: `w15:paraIdParent` is set and each reply has its own range markers around the parent's anchor. Word renders the four replies threaded.
- The resolved and open flags match the task exactly.
- Nothing else changed. The visual and textual diffs against the baseline show only the intended edits; the header, footer, table and signature block are identical.
- The reply wording is reasonable: "Declined — standard term." for the open IP and indemnification threads, and "Noted — addressed; Section N stands as drafted." for the two resolved threads.

## Demerits

- **minor:** `commentsExtended.xml` keeps an orphan `<w15:commentEx w15:paraId="00100090" w15:paraIdParent="97CB30E2">` from a reply the agent deleted (old c4). No comment has that paraId. Word ignores it and the file rendered, but it is stray metadata.
- **minor:** Comment ids now have a gap (c0–c3, then c5–c8), left over from the deleted and re-added reply. This is cosmetic.
- **info (docx-cli bug, not charged to OfficeCLI):** docx-cli reads `w15:done="true"` as not resolved, so `comments list` and `read` hide the resolved state.

## Frictions

The agent reported these:
- **major:** `set --prop text` on a reply silently dropped its parentId threading. The agent noticed, deleted the reply, and re-added it with `parentId=0`.
- **minor:** Reply anchoring is undocumented; a reply needs an explicit body path.
- **minor:** `validate` reports schema errors that were mostly already in the baseline (paraId MaxExclusive, textId 00000000, duplicate Heading1). The agent was not sure whether the 4 textId errors were new.
- **positive:** Accept/reject by `/revision[@id=N]` worked cleanly, and the ids stayed stable.

**Read:** The revision half went smoothly. The comment half needed one real dead end: editing a reply's text silently un-threaded it, and the agent had to delete and re-add the reply. That was a real tool trap, which the agent caught by checking its work. It left a small orphan entry in commentsExtended. Overall this was moderate friction with a correct result.
