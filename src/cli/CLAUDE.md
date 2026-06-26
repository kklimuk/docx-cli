# src/cli — command surfaces

Each command is a kebab-case folder. `index.ts(x)` is the public surface: `export async function run(args: string[]): Promise<number>`. Sub-files for shared helpers; no cross-feature imports unless centrally exposed via `@core/*`.

## Adding a CLI command

Add a folder here, register it in the COMMANDS map in [index.ts](index.ts), add tests under `tests/cli/`, document in README and the root CLAUDE.md. If it emits OOXML, the file is `.tsx` and uses JSX components from `@core/jsx`.

## read-markdown structural annotations

`read` surfaces structural facts the GFM body can't show as `<!-- docx:TYPE … -->`
HTML comments, built via `formatNote` / `htmlAttr` in [read/annotations.ts](read/annotations.ts).
These are read-time VISIBILITY hints — the importer DROPS them all (the structure
survives edits in place, `read --ast` is lossless, and the authoring verbs manage
it), emitted **deviation-only** (only what differs from the document default).
Today: `docx:section` (rendered at the section's START, with an
`applies-to="pX..pY (below)"` scope on deviating sections), `docx:page`,
`docx:table` (`widths`/`borders` plus the formatting `tables format` authors —
`align`/`style`, and the row-level `repeat-header`/`row-heights` that have no GFM
row slot so they ride the table note keyed by `rR`), per-cell `docx:cell` hints
(`gridSpan`/`vMerge`/`shading`/`vAlign`/`borders`, plus a `halign` derived from a
uniform non-default cell-paragraph alignment — cell text alignment is otherwise
invisible in a GFM cell), a head
`docx:track-changes on` when the doc's tracking toggle is enabled,
`docx:header`/`docx:footer` notes (the content in a `text` attr — so the importer's
`docx:` drop can't re-inject it into the body — with fields as `{page}`/`{date}`/…
tokens; `type` attr only for `first`/`even`). Section annotations render at the
section's START (where its content begins), not the sectPr's physical end, so each
reads right before the content it governs (`computeSectionStarts`/`renderSectionStart`
in `read/markdown.ts`). **`docx:section` is the one annotation that's deviation-only
in its ATTRIBUTES but not its presence: EVERY section emits a `<!-- docx:section sN -->`
marker — a contentless default-single-column section and the trailing mandatory
sectPr included — so the read consistently flags where each section begins (and the
`sN` an agent addresses); `cols`/`type`/`applies-to` appear only when they deviate.** A UNIFORM marginal (same across every section) rides the head
block instead; a marginal whose text DIFFERS by section renders at that section's
start. And
`docx:layout` on a tab-aligned paragraph that wraps — either inside a multi-column
section (tab stops wrap mid-line there) or a line whose trailing content sits on a
right-edge LEFT tab so a long value overflows the margin (the résumé
`San`/`Francisco` split). The latter warn names its cure: `edit --at pN --tabs
right` swaps the LEFT tab for a RIGHT tab flush at the margin (in `cli/edit/tabs.ts`,
emitted via `ParagraphOptions.tabs`), so the content right-aligns instead of
wrapping. Both are render-only breaks Markdown can't show. **`docx sections`
page-setup now CURES the right-edge-tab case for you**: changing margins/size
(doc-wide, or on a single-section doc) auto-applies that same `--tabs right` cure
to every fragile right-edge tab — calibrated to the OLD margins, it would wrap at
the new width — so the agent doesn't have to act on the hint (weak agents dismiss
it as "informational"). The hint still fires for cases page-setup can't reach (a
hand-set tab, a multi-section per-section edit).
The per-paragraph `docx:p` note carries `style`/`align` plus direct paragraph
spacing/indent (`space-before`/`space-after`/`line-spacing`/`indent-left`/
`indent-right`/`first-line`/`hanging`, in points/inches/multiples so an agent can
read a value and re-apply it through the matching `edit`/`insert` flag). Like
`docx:cell`, the note carries the locator as its leading token, so a paragraph
WITH a `docx:p` note does NOT also get the bare `<!-- pN -->` locator — that would
duplicate `pN`. Plain paragraphs (no note) keep the bare locator. Every paragraph
shows its addressable `pN` exactly once, either bare or as the note's leading token.
Full contract in [src/core/markdown/CLAUDE.md](../core/markdown/CLAUDE.md). New
structural annotation? Use `formatNote`, keep it deviation-only, and remember the
importer won't reconstruct it — NO comment is parse-back, not even `docx:base`.

## Hyperlink/image edits emit audit comments under track-changes

OOXML has no `w:hyperlinkChange` / `w:drawingChange` element — Word silently bypasses tracking for hyperlink edits and image swaps. We compromise: when `<w:trackChanges/>` is on, `hyperlinks add/replace/delete` and `images replace` each auto-emit a `[docx-cli] …` comment anchored to the affected span/run, attributed via the `--author` chain. The mutation itself stays silent (no fake `<w:ins>`/`<w:del>` — OOXML has no honest construct for it). The entry point is `new Comments(document).addAudit(anchor, { body, author, date })` (in [@core/comments](../core/comments/index.tsx)); the lower-level marker helpers `findContainingParagraph`, `findElementOffsetsInParagraph`, `addCommentMarkersAroundRun` live in [@core/comments/markers](../core/comments/markers.tsx). When track-changes is off, no comment is emitted.

`images delete` is the exception — deleting an image is honest content removal, so under tracking it wraps the drawing's run in a real `<w:del>` (accept removes it, reject restores it; the media part is kept until accept). It does **not** emit an audit comment. Replace stays audit-comment because swapping bytes has no tracked-change construct, but removal does.
