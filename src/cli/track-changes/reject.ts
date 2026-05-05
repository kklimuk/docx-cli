import { runApply } from "./apply";

const HELP = `docx track-changes reject — reject tracked changes (revert to pre-change state)

Usage:
  docx track-changes reject FILE --at tcN [options]
  docx track-changes reject FILE --all [options]

Rejecting an insertion (<w:ins>) removes the element and its inserted text.
Rejecting a deletion (<w:del>) unwraps the wrapper and converts <w:delText>
back to <w:t>, so the text reappears as plain runs.

Out of scope: tracked paragraph marks (<w:rPr><w:del/></w:rPr> inside <w:pPr>),
formatting changes (<w:rPrChange>/<w:pPrChange>), and tracked moves (<w:moveFrom>/
<w:moveTo>). These aren't modeled in the AST today and --all silently skips them.

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
