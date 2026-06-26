# Grading rubric — contract-markup (JUDGE ONLY)

## Pass conditions

1. **Tracked changes enabled** — the document's revision-tracking flag is on (verifiable
   via `docx track-changes list`; the list is non-empty and the document setting is
   active).

2. **≥ 3 tracked text edits or deletions** — at least three of the following egregious
   terms are revised as tracked changes (visible redlines, not silent edits):
   - §6 Non-Competition: five-year / worldwide scope reduced or struck
   - §9 Limitation of Liability: the asymmetric $100 cap made mutual or raised
   - §3 Compensation: Net-90 changed to Net-30 (or similar)
   - §8 Indemnification: Company's-own-negligence carve-out added

   Verified by: `docx track-changes list` shows ≥ 3 entries of type `insertion` or
   `deletion` touching the relevant clause paragraphs.

3. **≥ 1 tracked paragraph-formatting change on §9** — the §9 Limitation of Liability
   paragraph is set to 1.5 line spacing (and/or space-after) as a tracked revision.
   This must appear as a `pPrChange` entry in `docx track-changes list` (type
   `formatting`), not as a silent reformat.

4. **≥ 4 anchored comments** — comments attached to text in at least four of:
   - §4 IP over-assignment
   - §11 Personal guarantee
   - §12 Company-selected arbitration / one-way fee-shifting
   - §7 Waiver of independent counsel
   - §13 Unilateral amendment by URL posting

   Verified by: `docx comments list` returns ≥ 4 entries, each with a non-empty
   `anchor` field pointing to text in the relevant clause.

5. **IP sub-points reformatted to lowercase Roman numerals** — the lettered
   sub-points under the Intellectual Property clause (the level-1 list items: "all
   work product…", "all intellectual property…", "all of the Contractor's
   pre-existing…", "all feedback…") are switched from lower-alpha (a/b/c/d) to
   lower-roman (i/ii/iii/iv). The top-level clause numbering is left untouched.

   Verified by: `docx read contract.docx --ast` shows those level-1 list
   paragraphs with `list.format` = `"lower-roman"` (the natural path is
   `docx lists set --at <sub-point pN> --format lower-roman`). Note the markdown
   body renders ordered items as decimal regardless, so the `--ast` `list.format`
   (or a `<!-- docx:list … format="lower-roman" -->` hint) is the signal, not the
   rendered glyph.

## How to verify

```
# 1. Check tracked changes (redlines + formatting revision)
docx track-changes list contract.docx

# 2. Check anchored comments
docx comments list contract.docx

# 3. Spot-check a specific clause paragraph for the pPrChange
docx read contract.docx --ast | grep -A5 pPrChange

# 4. Confirm the IP sub-points are lower-roman (level-1 list items)
docx read contract.docx --ast | grep -o '"format":"lower-roman"'
```

All five conditions must pass for the scenario to be scored green.
