# Review: contract-markup (OfficeCLI, r1)

**Bucket:** editing. Legal review: redlining and commenting.

## Verdict
- **Task success:** success. All 5 rubric conditions pass.
- **Renders correctly (Word opened it):** true. Word produced 3 output pages. The reported render "error" (`close exited 1 … nothing to close`) is an OfficeCLI session no-op, not a document defect.
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Rubric check
1. **Changes are tracked.** Pass. `docx track-changes list` shows 13 tracked changes (7 logical): rev0 p7, rev1 p18, tc4 p18 deletion, rev2 p22, rev3/rev4/rev5 p24. All are authored "Mercer Design Studio". The XML has 7 `<w:del>` and 6 `<w:ins>`. The document toggle is off, which the rubric allows.
2. **At least 3 tracked text edits.** Pass, with all 4 done:
   - §3: Net 90 at sole discretion became Net 30 from the invoice date.
   - §6: five years worldwide became six months of direct competition, and the $50,000 liquidated-damages sentence is a tracked deletion.
   - §8: the Company's own-negligence clause was replaced with a Contractor-fault-only indemnity that excludes Company negligence.
   - §9: the cap is now mutual and the consequential-damages exclusion applies to both parties.
3. **§9 line spacing.** Pass. The AST for p24 shows `spacing.line=360, lineRule=auto`, and `read` prints `line-spacing="1.5"`. The change is untracked, which the rubric allows. The render shows §9 visibly opened up.
4. **At least 4 anchored comments.** Pass, with 10 comments, each anchored to a whole clause paragraph:
   - Flagged items: IP (b) on p12 and IP (c) on p13, §7 on p20, §11 on p28, §12 on p30, §13 on p32. All 5 flagged topics are covered.
   - The 4 redlined clauses (p7, p18, p22, p24) also carry comments explaining the reason.
5. **IP sub-points in lower Roman.** Pass. In the output, `numbering.xml` abstractNum 2, ilvl 1 changed from `lowerLetter` to `lowerRoman`, and `lvlText "(%2)"` is unchanged. Level 0 is still `decimal`. Word renders (i)–(iv), and the top-level 1.–15. numbering is untouched.
   - The docx-cli `read --ast` does not output `list.format` for this list in either the baseline or the output. That is a gap in the docx-cli reader, not an OfficeCLI failure. The rendered glyphs and the raw XML confirm the change.

## Merits
- All 4 "must change" terms are redlined as clean, minimal del/ins pairs. The deletion of the liquidated-damages sentence is its own tracked deletion.
- The replacement wording is sensible and fixes each objection, for example "fees paid in the preceding month for both parties" and "neither party is liable" for consequential damages.
- There are 10 well-reasoned comments that cover every flagged item, including both IP sub-points (b) and (c). All are attributed to Mercer Design Studio.
- The Roman-numeral change was done correctly at the numbering definition, not as literal text. Only `numFmt` changed.
- §9 spacing is set to exactly 1.5 lines, applied untracked, which is cleanest for house-style changes.
- Everything outside the targets is preserved: header, footer, payment table, signature block, pagination (3 pages), and all other clauses. The baseline/output read diff shows only p7, p18, p22 and p24 changed.
- The agent's summary honestly lists language it left alone: §3 set-off and §8 "uncapped".

## Demerits
- **Minor:** every comment anchors to the whole clause paragraph rather than the offending phrase. Comments are still correctly pinned to the right clause. This is a limit of the tool's `comments add` anchoring, which offers no phrase anchor.
- **Minor (pre-existing, not introduced):** a duplicate `Heading1` styleId exists in `styles.xml` in both the baseline and the output. The agent correctly left it alone, but OfficeCLI `view issues` crashes on it.
- **Minor (observation):** the §3 payment table still says "Net 90 after final acceptance" in the Payment Timing cell, so it is inconsistent with the redlined Net 30 text. The task did not ask for this, but a careful reviewer would have redlined it too.

## Frictions
The agent reported:
- **Major:** the numbering path syntax is inconsistent. `get --depth` lists `abstractNum[N]/lvl[M]`, but `set` needs `abstractNum[@id=N]/level[L]` (0-based) with `format`, not `numFmt`, and the listed path gives a misleading "not found" error.
- **Major:** `view issues` and `validate` crash or complain on the pre-existing duplicate Heading1 style.
- **Minor:** there is no phrase anchor for comments.
- **Minor:** the help does not say that a tracked find/replace with an empty replacement leaves a stray space unless the find string includes the leading space.
- **Minor:** no hint about the document-level tracking toggle.

Dead ends:
1. The agent first treated the (a)–(d) labels as literal text.
2. The `lvl[2]` path from the listing failed.

**Read:** the result is clean and fully correct, but the path to it was a moderate slog on the numbering change, because the OfficeCLI path forms in the listing and in `set` don't match. The redlines and comments themselves went smoothly.
