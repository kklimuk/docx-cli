# Review — contract-finalize (bucket: editing / review-finalize, docx-cli r2, haiku)

## Verdict
| Dimension | Result |
|---|---|
| Task success | **success** (13/13 rubric checks pass) |
| Renders correctly (Word opened it) | **true**: 3 output pages rendered |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived write→read loop | **true** |

## Rubric check
1. Net 30 accepted: p7 reads "Net 30 from the Company's acceptance" as clean text.
2. Non-compete accepted: p18 reads "For one (1) year after the Agreement ends" as clean text.
3. Liability cap rejected: p24 reads "one hundred dollars ($100)", and the $100,000 insertion is gone.
4. §11 deletion rejected: "11. Personal Guarantee." plus its body (p27/p28) are back, and the later sections renumber to 12–15.
5. Indemnification pPrChange accepted: p22 has `line-spacing="1.5"` live, and the render shows the 1.5 spacing with no change bar.
6. `track-changes list`: "No tracked changes".
7–8. Four replies, c4–c7, have parentId c0–c3, and each anchor.startBlockId matches its parent's paragraph (p3/p12/p22/p30).
9. `read --comments` shows [^c4]…[^c7], each with a "↳ cN" marker.
10–13. c0 and c3 are resolved:true; c1 and c2 are resolved:false.

`docx validate`: valid (10 parts).

## Merits
- One `track-changes apply` call (accept rev0 rev1 tc4; reject rev2 tc7 tc8 tc9) settled all five decisions correctly.
- The replies are anchored and threaded. In Word they appear as blue replies under each red parent comment.
- Baseline vs output, visual: the header, footer, the milestone table, the signature table and all other paragraphs match. The only differences are the intended ones: redlines are gone, §11 is restored, the following sections are renumbered, and the reply balloons are added.
- Baseline vs output, text diff of read.md: only p24 ($100) and the §11 restoration with its renumbering differ. No unintended text change.
- The agent noticed it had written reply wording the document does not support and corrected it.

## Demerits
- **Minor**: the reply positions are the agent's own wording, and all four replies say "Declined", including the replies on the two resolved threads (c4 "Declined — auto-renewal stays as drafted; standard term."). The task allowed this ("Declined — standard term." was one of its examples), so it is not a failure.
- **Minor, outside scope**: table cell t0:r1c2 still says "Net 90 after final acceptance", which no longer matches the accepted Net 30 body text. This was in the original document and is outside the task. The agent flagged it and correctly left it alone.

## Frictions
- Agent-reported: the task does not say which position to take on each comment, so the agent had to draft the positions itself.
- Agent-reported: `track-changes list` mixes rev and tc handles, and tc8 is a bare paragraph-mark delete that the agent had to infer belongs with the §11 deletion (tc7/tc9). A grouping hint would help.
- Dead end: the first reply batch had inaccurate wording, so the agent deleted c4–c7 and re-posted them. The error was the agent's, not the tool's.
- Overall the path was smooth: tracked changes were settled in a single batched call, and replies went through comments reply --batch with resolve. The only rework was self-inflicted reply wording.
