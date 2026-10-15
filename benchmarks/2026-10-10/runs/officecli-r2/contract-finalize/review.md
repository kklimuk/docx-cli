# contract-finalize — OfficeCLI run r2 (bucket: editing / review finalize)

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (3 pages rendered)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Merits
- **Tracked changes:** all handled correctly. Net 30 and "one (1) year" are clean text, the liability cap reads "one hundred dollars ($100)" again, and §11 Personal Guarantee is restored (heading and body; later sections renumber to 12–15 as in the original draft). `docx track-changes list` returns "No tracked changes", and document.xml has 0 `w:ins`/`w:del`/`pPrChange`/`rPrChange`.
- **Indemnification spacing:** accepted as live formatting. p22 carries `w:spacing line="360"` (1.5x), `read` shows `line-spacing="1.5"`, and the render has the wide spacing with no change bar.
- **Replies:** four threaded replies (c4–c7) with `parentId` c0–c3. Each reply has its own comment range around the parent's paragraph, so `anchor.startBlockId` is p3/p12/p22/p30 (not empty). Each shows in `docx read --comments` as `[^cN]` with `↳ cN`. Word renders every reply directly under its parent in the margin.
- **Resolve state:** correct in OOXML. `commentsExtended.xml` marks the auto-renewal root (97CB30E2 = c0) and the arbitration root (EC14166F = c3) `w15:done="true"`. The IP (c1) and indemnification (c2) roots are `done="false"`.
- **Before/after:** the baseline-vs-output markdown diff shows only the liability-cap restore and the §11 restore with its renumbering. Visually, the header, footer, payment table (including the original "Net 90 after final acceptance" cell, which was never a revision) and the signature block are unchanged.

## Demerits
- **(minor) Reply wording:** the auto-renewal reply says "Agreed — revised." although the renewal clause text was not changed; the arbitration reply says the same. The wording comes from the task's own example, so this is not a rubric failure, but the substance is slightly inaccurate.
- **(minor, docx-cli reader gap, not an OfficeCLI defect):** `docx comments list` shows no `resolved: true` for c0/c3, and `--open` still lists all 8 comments. OfficeCLI's relationship for commentsExtended uses an absolute Target (`/word/commentsExtended.xml`, Id `R560341e293d74cad`). That is valid OPC, but docx-cli apparently fails to resolve it, so it never sees the `w15:done` flags. Rubric items 10–13 were verified directly in the XML instead.

## Frictions
- The agent needed 11 sequential `set /revision[@id=N] --prop revision.action=...` calls. There is no batch option and no way to act on a paired deletion and insertion together: accepting deletion id 0 left insertion id 1 pending.
- There is no reply verb. A reply requires `add --type comment --prop parentId=N` on the parent's anchor paragraph.
- Resolving with `done=true` applies to the root only; the replies stay `done=false`.
- `view text` does not show pending insertions or deletions, so the agent fell back to queries and raw XML.
- `validate` reports schema messages in comments.xml that were already present in the original file.
- It briefly mistook the "11." list number for literal text.
- **Read:** a moderate slog through verbose per-id operations, but no destructive dead ends, and the result is correct.
