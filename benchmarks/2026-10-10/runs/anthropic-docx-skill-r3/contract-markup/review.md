# Review — contract-markup (bucket: editing / redline + comment) — arm: Anthropic docx skill (haiku), r3

## Verdict
- **Task success:** success (all 5 rubric conditions pass)
- **Renders correctly (Word opened it):** true — 3 document pages + 1 Word comment-overflow page
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true (all content present in render; docx-cli also reads tracked changes + comments)

## Rubric check
1. **Changes tracked** — PASS. `docx track-changes list` shows 25 tracked changes (16 logical), visible as redlines in the Word render.
2. **≥3 tracked text edits** — PASS, all 4: §3 Net 90→Net 30 (body p7 + schedule cell t0:r1c2), §6 five years/worldwide non-compete replaced with a 12-month customer non-solicit and the $50,000 liquidated-damages sentence deleted (rev2, rev3, tc12), §8 indemnity limited to Contractor fault with an explicit Company-negligence exclusion (rev4, rev5), §9 cap made mutual at 12 months' fees (rev6–rev8).
3. **§9 line spacing** — PASS. p24 (and heading p23) have tracked pPrChange spacing ·→360 auto; read shows `line-spacing="1.5"`. The render shows the opened-up §9 body and a "Formatted: Line spacing: 1.5 lines" balloon.
4. **≥4 anchored comments** — PASS, all 5 flagged clauses are covered: IP (c4 on p10), §7 counsel waiver (c7 on p20), §11 guarantee (c5 on p28), §12 arbitration (c6 on p30), §13 amendment (c8 on p32). There are also comments on the 4 redlined clauses, for 9 comments in all.
5. **IP sub-points lower-roman** — PASS in the render. The sub-points show (i)–(iv) with the same hanging indent as the original (a)–(d). The XML has a new abstractNum 3 (lowerRoman, "(%1)") / numId 3, and the change is tracked as pPrChange ("Formatted: Numbered … i, ii, iii"). Top-level clauses still number 1–15. Note: docx-cli's `--ast` shows no `list.format` field for these paragraphs, and its markdown drops the `docx:list` hint. That is a quirk in docx-cli's reader, not a failure by the competitor.

## Merits
- All four unacceptable terms were redlined with substantive, legally sensible rewrites. The Net-30 change was also made in the payment-schedule table row.
- Every wording change is tracked (w:ins/w:del). There are no silent edits: the accepted-view diff against the baseline shows changes only in p7, t0:r1c2, p18, p22 and p24.
- 9 comments, each anchored to the whole relevant clause body. Each states clearly what is objected to and why, and they are authored as "Mercer Design Studio".
- The formatting fixes (1.5 spacing, roman numerals) were applied and even tracked as pPrChange. Original indentation and parenthesized label style are kept.
- Header, footer, table shading, signature block and clause numbering are unchanged. The layout is intact, and the extra page holds only Word's comment-overflow markup.

## Demerits
- **minor**: The agent's summary says 8 comments, but there are 9. This is a small self-report inaccuracy.
- **minor**: comments.xml was written pretty-printed. Word ignores the whitespace (balloons render cleanly), but docx-cli's `comments list` shows the text padded with newlines and indentation, so a non-Word consumer would see whitespace noise.
- **minor** (not a competitor fault): docx-cli's read view shows the re-numbered IP sub-points as plain "1.–4." with no list hint, and its AST has no format field for them. Word renders (i)–(iv) correctly.

## Frictions
The agent reported:
- Hand-writing numbering.xml for the a→i switch (major)
- comment.py printing marker snippets without placing them (major)
- No pPrChange docs
- A missing CommentReference style
- PYTHONPATH needed for merge_runs.py under `python3 -I`
- Having to write its own tracked find/replace helper

Its dead ends:
- ModuleNotFoundError under isolated mode
- A misplaced commentRangeStart inside pPrChange, caused by its own regex and caught by inspection
- macOS lacking a `timeout` binary
- pandoc's flattened numbering looking like a regression

Read: the result is excellent, but the path was a slog of hand-built XML (numbering, pPrChange, comment markers, redline helper). One self-inflicted structural bug was caught only by manual inspection.
