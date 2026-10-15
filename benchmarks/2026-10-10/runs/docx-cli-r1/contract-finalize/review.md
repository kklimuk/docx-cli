# Review — contract-finalize (editing bucket) — docx-cli r1, haiku

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** yes (3 output pages rendered)
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes

The run was finished, not abandoned. The harness logged the self-report as "lost", but the agent's final message contains a complete fenced json block. That points to a parsing problem on the harness side, not an incomplete run. Everything below was checked against the actual file.

## Rubric check
| # | Check | Result | Evidence |
|---|---|---|---|
| 1 | Net 90 → Net 30 accepted | PASS | p7 reads "Net 30 from the Company's acceptance" as clean text, with no redline in the render. The table cell "Net 90 after final acceptance" is the original, untouched text in both baseline and output. |
| 2 | Non-compete 1 yr accepted | PASS | p18 reads "For one (1) year after the Agreement ends" as clean text. |
| 3 | Liability cap rejected | PASS | p24 reads "one hundred dollars ($100)". The $100,000 insertion is gone. |
| 4 | §11 deletion rejected | PASS | "11. Personal Guarantee." (p27/p28) is present, and the sections after it renumber to 12–15. |
| 5 | Indemnification pPrChange accepted | PASS | p22 has a live `line-spacing="1.5"`. The render shows 1.5 spacing with no change bar. |
| 6 | No pending changes | PASS | `track-changes list` → "No tracked changes". |
| 7 | All 4 comments have replies | PASS | c4→c0, c5→c1, c6→c2, c7→c3. |
| 8 | parentId + non-empty anchor | PASS | Every reply has a parentId and anchors to p3, p12, p22 or p30. |
| 9 | Replies visible in read view | PASS | `read --comments` shows [^c4]…[^c7], each with a `↳ cN` marker. |
| 10 | Auto-renewal resolved | PASS | c0 has `resolved: true`. |
| 11 | Arbitration resolved | PASS | c3 has `resolved: true`. |
| 12 | IP still open | PASS | c1 has `resolved: false`. |
| 13 | Indemnification still open | PASS | c2 has `resolved: false`. |

`docx validate` → valid (10 parts checked).

## Merits
- All five revision decisions came out correct, applied in one `track-changes apply` call.
- The §11 rejection fully restores the heading and body paragraph. List numbering (11–15) re-flows correctly in both the Word render and the read view.
- The indemnification spacing change is now live formatting: 1.5 spacing is visible in the render and shown as a `line-spacing` hint in the read view.
- Every reply is threaded and anchored. Word renders each one as a reply bubble next to its parent.
- The resolve state is correct for all four threads.
- The textual diff between baseline and output read views shows only the intended changes (§9 cap, §11 restored, renumbering). Header, footer, tables and signature block are unchanged, and the visual comparison confirms it.
- The agent caught the invented arbitration concession it had written in its own reply and corrected it (deleted that reply, re-posted "Agreed — revised.").

## Demerits
- **minor**: The replies are attributed to "Reviewer" (initials R), not to the Company. The task says the replies should state the Company's position. The text does that, but the author label does not identify the Company. The agent noticed this and did not fix it.
- **minor**: The c3 reply says "Agreed — revised.", but the arbitration clause (p30) was not revised. It still says the Contractor pays all fees in advance. The agent used the task's example wording, so this does not violate the rubric. It is a semantically hollow response on the record, though.

## Frictions
The agent reported these frictions and dead ends:
- (dead end) A `comments reply --batch` call wrote an invented legal concession for c3. The agent fixed it with `comments delete` and a fresh reply.
- (major) Replies default to the author "Reviewer", and the batch ack (bare c4..c7) does not show the author.
- (minor) The §11 paragraph deletion shows up as three separate handles (tc7/tc8/tc9) with nothing grouping them.
- (minor) `track-changes apply --dry-run` prints raw JSON.
- (minor) The pPrChange line in the list doesn't identify which clause it belongs to.
- (minor) The `apply` confirmation ("10 changes") doesn't separate accepts from rejects.

Overall the path was fairly smooth. The core work took one batched apply plus one batched reply. The only detour was the self-inflicted reply-content fix. The author default and the ungrouped paragraph-deletion handles are real UX gaps in the tool, not bugs.
