# Review — contract-finalize (bucket: editing / review-finalize) — arm: Anthropic docx skill (haiku), r1

## Verdict
- **Task success:** success
- **Renders correctly (Word opened it):** true (3 output pages rendered)
- **Layout sound:** true
- **Formatting preserved:** preserved
- **Survived read loop:** true

## Merits
- All five revision decisions are correct and clean: "Net 30" (p7) and "one (1) year" (p18) are plain text; liability cap reads "one hundred dollars ($100)" (p24) with no trace of $100,000; §11 Personal Guarantee heading + body restored, downstream sections renumber back to 12–15 exactly as the Company's original draft; indemnification paragraph keeps 1.5-line spacing as live formatting (render shows the wider spacing, `docx read` shows `line-spacing="1.5"`, no pPrChange).
- `docx track-changes list` → "No tracked changes" (criteria 1–6 pass).
- `docx comments list` → four Contractor comments c0–c3, each with exactly one Company reply (c4–c7) with correct `parentId` and non-empty `anchor.startBlockId` (criteria 7–8).
- `docx read --comments` shows every reply as its own `[^cN]` footnote with `↳ cN` (criterion 9).
- Resolution state exactly right: c0 (auto-renewal) + c3 (arbitration) and their replies resolved; c1 (IP) and c2 (indemnification) open (criteria 10–13). Word's render shows the resolved threads greyed and the open ones active, replies threaded under their parents.
- Visual before/after diff: header, footer, payment table, signature block, list numbering and fonts unchanged; only the intended text/markup changed.

## Demerits
- **minor** — comments.xml reply entries were written pretty-printed (newlines/indentation between `<w:p>`/`<w:r>` elements). Word ignores this (the render shows clean reply text), but docx-cli's `comments list` surfaces the whitespace in the reply `text` field. Cosmetic; a reader-side quirk more than a competitor defect.
- **minor** — the agent rewrote the four pre-existing comment paraIds (cleared the top bit) to satisfy the skill validator; consistent across comments/commentsExtended/commentsIds, harmless in the render, but an unrequested mutation of opposing counsel's metadata.
- **minor (content judgment)** — auto-renewal reply reads "Declined — … standard term; resolved." — a plausible Company position; arbitration reply similar. Acceptable per the task's examples.

## Frictions
- Agent reported: no skill command for selective accept/reject (only accept-all via LibreOffice) → hand-edited document.xml with a regex script; no resolve/reply helper (hand-set `w15:done`, hand-placed nested comment range markers); `comment.py` failed under `python3 -I` (ModuleNotFoundError: office); `validate.py --auto-repair` flagged paraId ≥ 0x80000000 but did not repair it, forcing a manual fix across three parts; author defaults to "Claude".
- Read: a slog — the correct result was reached mostly by raw XML surgery the skill does not tool, a high-risk path for a weak agent, but the output is correct.
