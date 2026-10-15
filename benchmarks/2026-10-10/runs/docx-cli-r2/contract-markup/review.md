# Review — contract-markup (editing bucket) — docx-cli r2 (haiku)

## Verdict
- **Task success:** success (all 5 rubric conditions pass)
- **Renders correctly (Word opened it):** true — 4 output pages rendered (page 4 is Word's markup-view comment-overflow page, not document content)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Rubric check
1. **Changes tracked** — PASS. `track-changes list`: 66 tracked changes (33 logical); doc toggle `docx:track-changes on`.
2. **>= 3 tracked clause edits** — PASS, all four: §3 Net 90 -> Net 30 from receipt of correct invoice (p7, rev0–rev9) plus table rows t0:r1c2 / t0:r5c1 / t0:r5c2 (rev10–13); §6 five-year worldwide non-compete + $50k LDs struck, replaced by 6-month customer non-solicit, heading retitled (rev14–23); §8 carve-out "except to the extent caused by the negligence or willful misconduct of the Company…" (rev24–25); §9 cap made mutual, $100 and one-sided consequential waiver struck (rev26–32).
3. **§9 spacing** — PASS. AST p24 `spacing.line=360 lineRule=auto` (also on heading p23); read shows `line-spacing="1.5"`; render shows visibly opened lines.
4. **>= 4 anchored comments** — PASS, 9 comments, all anchored (author "Mercer Design Studio"): IP p10 (c1), §7 counsel waiver p20 (c3), §11 guarantee p28 (c6), §12 arbitration p30 (c7), §13 amendment p32 (c8) — all five flag items — plus explanatory comments on the 4 redlined clauses.
5. **IP sub-points lower-roman** — PASS. AST p11–p14 `list.format = "lower-roman"`; render shows (i)–(iv). Top-level clauses still render 1–15 in Word.

## Merits
- All four unacceptable terms redlined as real tracked insertions/deletions with sensible, lawyerly replacement wording; the payment-schedule table rows were also updated consistently with §3.
- All five "flag for discussion" items received clear, studio-voiced comments anchored to the right clause paragraph.
- §9 spacing applied directly (1.5 / line=360 auto), untracked as allowed.
- IP sub-points switched to i–iv; top-level numbering 1–15 visually intact after repair.
- Baseline vs output diff shows ONLY intended changes: header, footer, table shading/borders, signature block, other clauses unchanged.
- Everything is retrievable on re-read (markdown, AST, track-changes list, comments list).

## Demerits
- **minor** — Numbering structure changed: the top-level clause list, originally one numId (2), is now split into numId 2 (§1–4), numId 3 (IP sub-points), numId 4 (§5–15, `start=5`). It renders correctly, but sections 5–15 no longer continue the list automatically (inserting a clause before §5 would not renumber them). Caused by the `lists set --restart` workaround. Read now shows `docx:list p15 start="5" format="decimal"` where baseline showed a plain paragraph hint.
- **minor** — Redlines are word-level fragmented (66 revisions / 33 logical for 8 edited paragraphs/cells); p18 and p7 are hard for the counterparty to read in the margin balloons (render page 1–2). It's acceptable redlining practice, but noisy.
- **minor** — Comments anchor whole paragraphs rather than the offending phrase. Acceptable.

## Frictions (agent-reported) and assessment
- **Tool bug (major):** `lists set --at p11 --format lower-roman` would change the whole shared numId 2, renumbering every top-level clause, because the IP sub-items share the abstract list with the clauses. There's no level-scoped option. The agent found this on a scratch copy.
- **Tool bug / UX (major):** `lists set --restart` splits the list and cascades: §5–15 restarted at 1. The agent had to repair it with `lists set --at p15 --restart --start 5`. No warning names the paragraphs whose numbering changed.
- Markdown renders roman sub-items as `1.`/`2.`. Only the `renders="i, ii, iii, …"` hint shows the format.
- `--from/--to` slice shows `start=5` mid-list, which is misleading.
- Span `--text ""` deletion semantics under tracking aren't documented in `edit --help`.
- Overall: the redlining and commenting path was smooth. The numerals house-style fix was a slog with two dead ends, and it only worked because the agent tested on a scratch copy. It points to a real `lists set` gap: the tool can't scope a change to one level of a shared multi-level list.
