# Review — contract-markup (Legal review: redlining + commenting), editing bucket

## Verdict

| Dimension | Result |
|---|---|
| Task success | **success** (all 5 rubric conditions pass) |
| Renders correctly (Word opened it) | **true**: 4 output pages rendered |
| Layout sound | **true**: page 4 is Word's comment-balloon overflow in markup view, not a document defect |
| Formatting preserved | **preserved** |
| Survived read loop | **true**: a fresh `docx read` matches `renders/output/read.md` byte for byte; `validate` reports the file clean |

### Rubric check
1. **Tracked changes**: `track-changes list` reports 50 changes (26 logical). The document toggle is `on`.
2. **At least 3 tracked clause edits**: 4 of 4 are present:
   - §3 Net 90 → Net 30 from receipt of a correct invoice, in p7 and also in the table cell t0:r1c2.
   - §6 non-compete: 5 years and worldwide scope → 12 months, limited to the field of work, no liquidated damages.
   - §8: indemnity now limited to the Contractor's negligence and carves out the Company's own negligence.
   - §9: mutual cap, and the $100 cap is struck.
3. **§9 spacing**: p24 (and the p23 heading) have `spacing.line=360 auto`. This is recorded as tracked format changes tc34/tc49, and the read view shows `line-spacing="1.5"`. It is visibly roomier on page 2 of the render.
4. **At least 4 anchored comments**: there are 9 (c0–c8). All 5 flag-only items are covered: §4 IP (p10), §7 counsel waiver (p20), §11 guarantee (p28), §12 arbitration (p30), §13 URL amendment (p32). Each comment is anchored across the full clause text.
5. **Lower-roman IP sub-points**: in the AST, p11–p14 have `list.format = "lower-roman"`. numbering.xml has a lvlOverride on numId 2 that sets ilvl 1 to `lowerRoman`. Word renders the items as (i)–(iv), and the top-level 1–15 numbering is unchanged.

## Merits
- All four unacceptable clauses are redlined as real `<w:ins>`/`<w:del>` revisions, visible in the Word render. The payment change also reached the payment-timing table cell.
- The comment text is clear and persuasive. Notes on the redlined clauses explain each change. Residual concerns (set-off right, uncapped indemnity survival) are raised rather than silently left.
- Spacing and numbering were applied cleanly. The top-level numbering, table, header, footer and signature block are untouched.
- The baseline-vs-output diff of read.md is limited to p7, t0:r1c2, p18, p22, p24 and hint lines. There are no unintended text changes.
- The workflow was efficient: one `edit --batch` and one `comments --batch`.

## Demerits
- **Major (tool, read view):** the lower-roman change is invisible in the default Markdown view. The sub-points still read `1. … <!-- docx:list p11 continues -->` with no `format="lower-roman"` hint, so only `--ast` reveals it. This breaks the spirit of "visible in Markdown" for a weak agent verifying its own work.
- **Minor (tool, read view):** after `lists set`, p2's hint changed from `docx:p p2 space-before="10pt" space-after="3pt"` to `docx:list p2 format="decimal"`. The spacing hint dropped out of the read view. The AST and the render are unchanged.
- **Minor (tool):** the word-level redline produces tracked changes that are only whitespace: a single deleted or inserted space (`<w:delText> </w:delText>`). `track-changes list` shows these as `¶` (for example `rev6 replace p18 ¶ → "directly"`), which wrongly suggests a paragraph-mark change. Word shows empty "Deleted:" balloons for them (page 2). The non-compete redline is choppy for the counterparty.
- **Minor (content):** §8 still says "uncapped and survives indefinitely". Comment c2 flags it. The task allowed this.

## Frictions
The agent reported:
- No `format="lower-roman"` hint in the read view (confirmed above).
- Ambiguous help for `lists set --format` on a multi-level list, and a dry run that doesn't name the levels or paragraphs it will change.
- A fragmented word-level redline list (confirmed above; made worse by the `¶` display of space-only fragments).
- `edit --batch --dry-run` previews only locators.
- No way to address a clause by its section number, so the agent mapped §9 to p23/p24 by hand.
- No guidance on the final state of the tracking toggle.

**Read:** a smooth path with no real dead ends. The friction was in verification (the house-style fix wasn't visible in the default view, so the agent had to use `--ast`), not in execution.
