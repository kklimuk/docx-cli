# Review — contract-markup (bucket: editing / legal review: redlining + commenting) — run docx-cli-r3

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success**, all 5 rubric conditions pass |
| Renders correctly (Word opened it) | **true**: 4 output PNGs; `docx validate` reports valid, 10 parts |
| Layout sound | **true** |
| Formatting preserved | **preserved** |
| Survived read loop | **true** |

### Rubric check
1. **Changes are tracked.** Pass. `track-changes list` reports 92 tracked changes (46 logical), and the doc toggle is `on` (`<!-- docx:track-changes on -->`).
2. **At least 3 tracked text edits.** Pass, all 4 are done:
   - §3 Net 90 → Net 30 (p7 rev0–rev8). The schedule table cells were also matched (rev9–rev14).
   - §6 non-compete → 12-month non-solicit, with liquidated damages struck and the heading retitled (rev15–rev26).
   - §8 mutual, fault-based indemnity subject to the §9 cap (rev27–rev35).
   - §9 mutual 3-month-fees cap and mutual consequential-damages waiver (rev36–rev45).
3. **§9 line spacing increased.** Pass. p24 carries `line-spacing="1.5"` in read. The render shows §9 visibly opened up while the other clauses are unchanged.
4. **At least 4 anchored comments.** Pass. There are 9 comments, all with non-empty anchors. They cover all 5 flagged clauses: IP (c1, p10), §7 counsel waiver (c3, p20), §11 guarantee (c6, p28), §12 arbitration (c7, p30) and §13 amendment-by-URL (c8, p32). There are also 4 objection notes on the redlined clauses (c0, c2, c4, c5).
5. **IP sub-points set to lower-roman.** Pass. In the AST, p11–p14 have `list.format = "lower-roman"`, and all 15 top-level items stay `decimal`. The render shows (i)–(iv) where the baseline had (a)–(d), and the parenthesis style is kept.

## Merits
- All four unacceptable terms were redlined as real tracked replacements. The rewrites are substantive and legally coherent: mutual, fault-based, and capped.
- The agent also updated the Payment Timing table cells to match the new payment terms, as tracked edits. That is good attention to consistency.
- Comments carry a clear objection rationale and the author name "Mercer Design Studio". All are anchored to the correct clause paragraphs.
- `lists set --format lower-roman` changed only level 1 (the IP sub-points). Top-level clause numbering, the heading spacing (p2 spacing before 200 / after 60 is intact in the AST) and the list numbering are all preserved.
- The §9 spacing was applied untracked, which the rubric allows. No other paragraph's formatting changed.
- Headers, footers, the table shading and borders, and the signature block all match the baseline.
- The document validates, and Word renders it with the markup balloons.

## Demerits
- **(minor, tool/read view)** The markdown read view still prints the IP sub-points as `1. 2. 3. 4.` and has no `format="lower-roman"` hint on level-1 items, so only `--ast` shows the change. That breaks the "visible in Markdown" goal for this formatting fix. The rubric acknowledges this behavior, so it is not scored against the output.
- **(minor, tool display)** `track-changes list` labels pure insertions and deletions with a `¶` placeholder (`rev1 ¶ → "payable"`, `rev4 "set off, deduct, or" → ¶`, `rev44 "not" → ¶`). These look like paragraph-mark changes but are not. The word-level fragmentation (10 edits became 46 revisions) also makes review harder.
- **(cosmetic)** p2's read hint changed from `docx:p p2 space-before="10pt" space-after="3pt"` to `docx:list p2 format="decimal"`. The AST shows the spacing is unchanged, so this is a read-view hint artifact of the numbering change, not a document change.
- Page 4 of the output render is Word's comment-balloon overflow page from the markup view. It is not a layout defect.

## Frictions (agent-reported) and assessment
- **Major (agent):** the markdown read view hides the lower-roman format. Confirmed: this is a real write-read-loop visibility gap for list formats on non-zero levels.
- **Ambiguous `lists set --format` help** ("the whole list"). The agent ran a scratch test to confirm that only level 1 changes. That was sensible but cost extra turns. The help text and the ack should name the affected level.
- **Fragmented `track-changes list` with misleading `¶` labels.** Confirmed in the listing above.
- **No tracking or author echo in the edit batch ack.** Minor UX gap.
- **Dead end on a mismatch the agent thought it saw:** the agent concluded that "the document actually renders them as decimal 1-4". This is false: the baseline Word render shows (a)–(d). The agent was misled by the markdown read view, which prints ordered lists as decimal. The criteria describe this as expected, but it is a concrete case where the markdown view actively misinforms a weak agent about the document's state. The agent still did the right thing.

**Read on the path:** The agent reached a correct, high-quality result efficiently, using batched edits and comments. Its only real detours came from the markdown read view hiding list formats and from vague `lists set` help, which led to a scratch-copy verification and a wrong belief about the original numbering.
