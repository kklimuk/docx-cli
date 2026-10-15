# Review — contract-markup (bucket: editing / legal redline + comments) — OfficeCLI r3

## Verdict
- **Task success:** success (all 5 rubric conditions pass)
- **Renders correctly (Word opened it):** true — 4 output pages rendered (page 4 is Word's comment-balloon overflow page for a truncated comment, not document content)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Rubric check
1. **Changes tracked** — PASS. `docx track-changes list` reports 18 tracked changes (10 logical), real `<w:ins>`/`<w:del>` with author "Mercer Design Studio". Doc toggle is off (allowed).
2. **>= 3 tracked text edits** — PASS, all four egregious terms redlined:
   - §3: "Net 90 from the Company's acceptance, which ... sole discretion" -> "Net 30 from the Company's receipt of a correct invoice" (rev0), plus schedule cell t0:r1c2 "Net 90 after final acceptance" -> "Net 30 after receipt of a correct invoice" (rev1).
   - §6: "five (5) years" -> "six (6) months", worldwide/any-industry -> "within the United States ... directly competes with the services ..." (rev2, rev3), $50,000 liquidated-damages sentence struck (tc8).
   - §8: "including claims arising from the Company's own negligence" -> "excluding claims to the extent arising from the Company's own negligence or willful misconduct" (rev4).
   - §9: cap made mutual, "$100" lesser-of removed, "Contractor's liability ... unlimited" struck, consequential-damages waiver made mutual (rev5, rev6, tc15, rev7).
3. **§9 line spacing** — PASS. p24 has `w:line="360" w:lineRule="auto"`; read shows `line-spacing="1.5"` on p24 (heading p23 also set to 1.5, harmless). Applied untracked (permitted). Render shows visibly opened-up §9 body.
4. **>= 4 anchored comments** — PASS. 10 comments, all anchored to full paragraphs: §3 (p7), §6 (p18), §8 (p22), §9 (p24), IP (b)/(c) (p12, p13), §7 (p20), §11 (p28), §12 (p30), §13 (p32). All five flagged clauses covered.
5. **IP sub-points lower-roman** — PASS. numbering.xml abstractNum 2 ilvl 1 is now `numFmt lowerRoman`, `lvlText "(%2)"`; ilvl 0 stays decimal. Word render shows (i)-(iv) with top-level 1..15 untouched. (docx-cli's `--ast` does not surface `list.format` for this list — a docx-cli reader gap, not an OfficeCLI defect; verified via raw OOXML + render.)

## Merits
- Every wording change is a genuine tracked revision with a sensible author; nothing changed silently (baseline vs output read-view diff touches only p7, t0:r1c2, p18, p22, p24 — exactly the redlined clauses).
- Redlines are surgical (phrase-level replace) rather than whole-paragraph rewrites, so the opposing side sees precisely what moved; the schedule table was conformed to the new payment term.
- Comments are substantive, labelled MUST CHANGE / FLAG FOR DISCUSSION, and cover all nine target clauses.
- House-style numerals done by changing the level format in place; top-level numbering intact.
- Header, footer, table shading/borders, signature block, and pagination preserved; Word opens and renders cleanly.

## Demerits
- (minor) Comments are anchored to whole paragraphs rather than the specific offending phrase — acceptable, slightly less precise.
- (minor) 1.5 spacing also applied to the §9 heading paragraph (not requested, harmless).
- (minor, not OfficeCLI's fault) docx-cli `read --ast` omits `list.format`, so the rubric's suggested grep finds nothing; verified via numbering.xml and the render instead.

## Frictions
- `view issues` crashes on a duplicate `Heading1` styleId that is pre-existing in the source (major UX gap in OfficeCLI's proofing check).
- `validate` reports the pre-existing duplicate styleId without distinguishing it from new errors.
- Numbering level addressing (`level[L]` = 0-based ilvl) not documented; one not-found dead end.
- Empty-replacement tracked deletion needs the leading space in the find string; author must be passed on every call.
- Read: the path was mostly smooth — tracked find/replace with curly quotes worked first try and no raw XML was needed; the dead ends were small detours, not a slog.
