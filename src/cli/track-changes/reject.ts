import { runApply } from "./apply";

const HELP = `docx track-changes reject — reject tracked changes (revert to pre-change state)

Usage:
  docx track-changes reject FILE --at tcN [options]
  docx track-changes reject FILE --all [options]

Rejecting an insertion (<w:ins>) or move destination (<w:moveTo>) removes the
element and its content (it shouldn't have arrived at this location).
Rejecting a deletion (<w:del>) or move source (<w:moveFrom>) unwraps the
wrapper and converts <w:delText> back to <w:t>, so the text reappears as
plain runs.

Section-property revisions (<w:sectPrChange>): reject restores the prior-state
snapshot — the live section's columns/type are replaced with the values that
were in effect before the tracked edit.

Paragraph-mark trackings (<w:ins>/<w:del> inside <w:pPr><w:rPr>): rejecting
a paragraph-mark insertion removes the entire owning paragraph (the inserted
break disappears — for sentinels created by "insert --section" this also
removes the section break the sentinel was carrying). Rejecting a
paragraph-mark deletion just removes the marker (the paragraph stays).

moveFrom and moveTo are processed independently. To fully undo a move, target
both halves (or use --all). The runtime treats them as paired only by their
shared revision id, not by atomic accept/reject.

Out of scope: formatting changes (<w:rPrChange>/<w:pPrChange>). These aren't
modeled in the AST today and --all silently skips them.

Options:
  --at tcN          Reject a single tracked change by id
  --all             Reject every tracked change (mutually exclusive with --at)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -h, --help        Show this help

Examples:
  docx track-changes reject doc.docx --at tc0
  docx track-changes reject doc.docx --all
  docx track-changes reject doc.docx --all --dry-run
`;

export async function run(args: string[]): Promise<number> {
	return runApply(args, "reject", HELP);
}
