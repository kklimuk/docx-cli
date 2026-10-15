# Review: contract-markup (bucket: editing, redlining + commenting). Arm: Anthropic docx skill (haiku), run r2

## Verdict
- **Task success:** success. All 5 rubric conditions pass.
- **Renders correctly (Word opened it):** yes. 4 output pages; page 4 is Word's normal comment-overflow page for a long balloon.
- **Layout sound:** yes
- **Formatting preserved:** preserved
- **Survived read loop:** yes

## Rubric check
1. **Tracked changes present.** `docx track-changes list` shows 10 revisions (5 logical replaces), all by author "Mercer Design Studio". PASS
2. **At least 3 tracked clause edits.** All four were made:
   - §3: "Net 90 from the Company's acceptance, which the Company may grant or withhold in its sole discretion" became "Net 30 from the Company's receipt of a correct invoice". The schedule row changed from "Net 90 after final acceptance" to "Net 30 after final acceptance".
   - §6: the 5-year worldwide non-compete and the $50k liquidated damages became a 6-month non-solicit limited to customers the Contractor dealt with directly.
   - §8: "including claims arising from the Company's own negligence" became "except to the extent caused by the negligence or willful misconduct of the Company".
   - §9: the cap is now mutual (fees paid or payable over 12 months), and the consequential-damages exclusion applies to both parties.

   PASS
3. **§9 spacing.** p24 shows `line-spacing="1.5"` (w:line=360, untracked). The render shows the opened-up lines. PASS
4. **At least 4 anchored comments.** There are 9 comments in total, each with an anchor. Five cover the "flag these" items: IP (p12), guarantee (p28), arbitration (p30), §7 counsel waiver (p20) and §13 amendment (p32). Four explain the redlines. PASS
5. **IP sub-points in lower-roman.** The render shows (i)–(iv) where the baseline had (a)–(d). numbering.xml has abstractNum ilvl 1 set to `lowerRoman`, and the top-level 1–15 numbering is unchanged. docx-cli's `--ast` doesn't print a `list.format` field for this document, which is a limit of docx-cli's reader, not a competitor defect. The render is authoritative here. PASS

## Merits
- All four must-change clauses are redlined as real w:del/w:ins with a consistent author. The edits are surgical: §3's set-off sentence and §8's trailing sentence keep their original text.
- The schedule table's Net 90 was caught and redlined too, so the table matches the clause.
- There are 9 well-written, clause-anchored comments, and the five discussion items are clearly marked "For discussion".
- The numbering change was scoped correctly: only ilvl 1 changed, and the agent checked that it doesn't affect the headings.
- The §9 spacing change is visible in the render, and the rest of the document is unchanged.
- The visual diff against the baseline shows no unintended changes to the header, footer, table styling, signature block or numbering.

## Demerits
- **minor:** In the schedule row, the edit "Net 30 after final acceptance" keeps the "final acceptance" condition, which is a little weaker than the clause's "receipt of a correct invoice".
- **minor:** The set-off/withholding row ("Unlimited / Sole discretion") and the set-off sentence in §3 were left alone and not flagged. The agent noted this itself. The task didn't require it.
- **minor:** comments.xml is pretty-printed. Word renders the comments cleanly, but other readers (for example `docx comments list`) pick up the inter-element whitespace in the comment `text` ("\n  \n ...").
- **minor (cosmetic):** The comment-author prefix in the balloons renders in a serif font, which suggests the comment paragraphs lack the standard CommentText/annotation styling. This has no functional impact.
- **minor:** The IP comment is anchored only to sub-point (ii), not to the whole IP section, although its text covers (ii) and (iii).

## Frictions
The agent reported:
- No recipe for a tracked replacement inside a run, so it hand-wrote a splitter for before / w:del / w:ins / after.
- comment.py doesn't place the anchor markers, so it hand-wrote commentRangeStart/End and commentReference.
- accept_changes.py gave misleading heading numbering on the untouched original, so it switched to pandoc for verification.
- There's no guidance on numbering.xml, or on whether to track formatting changes.
- A regex failure on w:spacing attributes.

**Read:** the result is correct, but getting there was a slog. The agent had to write raw OOXML tooling (tracked-run splitting, comment anchoring, numbering edits) by hand, which the skill doesn't provide. It's a strong outcome reached by a high-effort, error-prone path.
