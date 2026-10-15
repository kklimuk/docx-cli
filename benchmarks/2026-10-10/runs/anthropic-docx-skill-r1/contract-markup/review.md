# Review: contract-markup (anthropic-docx-skill, r1)

**Scenario:** contract-markup (legal review: redlining and commenting). Bucket: editing.
**Arm:** competitor (Anthropic docx skill, haiku agent)

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** (all 5 rubric conditions pass) |
| Renders correctly (Word opened it) | true (4 output pages; page 4 is Word's comment-overflow page in markup view) |
| Layout sound | true |
| Formatting preserved | preserved |
| Survived read loop | true |

## Rubric check

1. **Changes are tracked.** Pass. `docx track-changes list` reports 23 tracked changes (14 logical). The toggle is off, which the rubric allows.
2. **At least 3 tracked text edits.** Pass. All four clauses are redlined:
   - §3 payment: rev0 "Net 90 from the Company's acceptance, which… sole discretion" → "Net 30 from the Company's receipt of a correct invoice". The schedule-table cell also changes 90 → 30 (rev1).
   - §6 non-compete: "five (5) years" → "six (6) months"; "anywhere in the world" → "within the United States"; the any-industry scope is narrowed to the design services; ", former, or prospective" is deleted; the $50,000 liquidated-damages sentence is deleted.
   - §8 indemnity: ", including claims arising from the Company's own negligence" → ", except to the extent caused by the negligence or willful misconduct of the Company".
   - §9 liability: the cap is mutual at 12 months' fees; the "Contractor's liability… unlimited" sentence is deleted; the consequential-damages exclusion is mutual.
3. **§9 line spacing.** Pass. tc22 is a tracked pPrChange with `spacing.line ·→360, lineRule auto`. The read view shows `line-spacing="1.5"` on p24, and the render shows visibly wider spacing in §9.
4. **At least 4 anchored comments.** Pass. There are 9 comments by "Mercer Design Studio". All five flagged clauses have one: §4 IP (p10), §11 guarantee (p28), §12 arbitration (p30), §7 counsel waiver (p20) and §13 amendment (p32). The four redlined clauses also each carry an explanatory comment. Every comment has a non-empty anchor.
5. **IP sub-points in lower-roman.** Pass. In numbering.xml, abstractNum 2 level 1 is now `lowerRoman`, `(%2)`. The render shows (i)–(iv) where the baseline showed (a)–(d). The top-level 1.–15. numbering is unchanged. abstractNum 2 is the clause list itself, so the only list affected is its own level 1. docx-cli's `--ast` does not emit a `list.format` field for this list. That is a gap in docx-cli's reader, not a competitor failure, and the render is authoritative.

## Merits

- All four unacceptable terms are redlined with surgical, word-level `w:del`/`w:ins`, not whole-paragraph replacement. The redlines read cleanly in Word's markup balloons.
- The agent also conformed the payment-schedule table cell (Net 90 → Net 30) as a tracked change.
- The §9 spacing change is a real tracked pPrChange, which goes beyond the rubric's requirement.
- The comments are substantive and specific, with proposed alternatives, and are attributed to Mercer Design Studio.
- The textual diff between the baseline and output read views shows changes ONLY in p7, t0:r1c2, p18, p22 and p24. Nothing outside the intended edits changed. The header, footer, table styling, signature block and clause numbering all match the baseline.
- The accepted-view wording is grammatical and coherent.

## Demerits

- **minor:** The comments are anchored to whole paragraphs, not specific phrases. This is acceptable, but less precise than a phrase anchor.
- **minor:** The table cell now reads "Net 30 after final acceptance", while the §3 body says "from the Company's receipt of a correct invoice". The trigger is slightly inconsistent, though the agent's comment says the schedule is "conformed".
- **minor:** Each comment run references `rStyle w:val="CommentReference"`, but styles.xml defines no such style, so the reference dangles. Word tolerates it and the render is fine. The agent itself flagged this as friction.
- **minor:** comments.xml is pretty-printed, with whitespace between elements. Word renders it correctly. docx-cli's `comments list` surfaces that whitespace in the `text` field, which is a quirk of docx-cli's reader and is not counted against the competitor.
- **observation:** The §7 counsel-waiver comment asks that the sentence be struck, but it is left unredlined. That matches the brief, which says to flag §7, not edit it.

## Frictions

The agent reported these:
- The skill has no tracked-change authoring helper. The agent hand-wrote a run-splitting redline script (scratch/redline.py).
- The skill gives no guidance on pPrChange or on how to change numbering.
- comment.py emits a CommentReference rStyle without checking that the style exists.
- Comment markers must be placed by hand. A naive regex put `commentRangeStart` inside `pPrChange`, and validate.py's error ("This element is not expected") did not locate the problem.
- merge_runs.py fails under `python3 -I` with ModuleNotFoundError.
- The agent hit two script reruns: a duplicated argument, and a wording slip ("anywhere in the United States") that it caught on the accepted-view check.

**Read on the path:** a slog. The result is correct and clean, but the agent had to build its own redlining tool and debug marker placement by trial and error. That is heavy lifting for a weak agent, and the result depended on careful self-verification. The UX demerit is significant even though the output passes.
