# src/cli — command surfaces

Each command is a kebab-case folder. `index.ts(x)` is the public surface: `export async function run(args: string[]): Promise<number>`. Sub-files for shared helpers; no cross-feature imports unless centrally exposed via `@core/*`.

## Adding a CLI command

Add a folder here, register it in the COMMANDS map in [index.ts](index.ts), add tests under `tests/cli/`, document in README and the root CLAUDE.md. If it emits OOXML, the file is `.tsx` and uses JSX components from `@core/jsx`.

## Hyperlink/image edits emit audit comments under track-changes

OOXML has no `w:hyperlinkChange` / `w:drawingChange` element — Word silently bypasses tracking for hyperlink edits and image swaps. We compromise: when `<w:trackChanges/>` is on, `hyperlinks add/replace/delete` and `images replace` each auto-emit a `[docx-cli] …` comment anchored to the affected span/run, attributed via the `--author` chain. The mutation itself stays silent (no fake `<w:ins>`/`<w:del>` — OOXML has no honest construct for it). Helpers live in [comments/helpers.tsx](comments/helpers.tsx): `emitAuditComment`, `findContainingParagraph`, `findElementOffsetsInParagraph`, `addCommentMarkersAroundRun`. When track-changes is off, no comment is emitted.

`images delete` is the exception — deleting an image is honest content removal, so under tracking it wraps the drawing's run in a real `<w:del>` (accept removes it, reject restores it; the media part is kept until accept). It does **not** emit an audit comment. Replace stays audit-comment because swapping bytes has no tracked-change construct, but removal does.
