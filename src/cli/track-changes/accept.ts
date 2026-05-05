import { runApply } from "./apply";

const HELP = `docx track-changes accept — accept tracked changes (incorporate into the doc)

Usage:
  docx track-changes accept FILE --at tcN [options]
  docx track-changes accept FILE --all [options]

Accepting an insertion (<w:ins>) unwraps the wrapper — the inserted text
becomes plain runs. Accepting a deletion (<w:del>) removes the element and
its <w:delText> entirely (the text disappears for real).

Out of scope: tracked paragraph marks (<w:rPr><w:ins/></w:rPr> inside <w:pPr>),
formatting changes (<w:rPrChange>/<w:pPrChange>), and tracked moves (<w:moveFrom>/
<w:moveTo>). These aren't modeled in the AST today and --all silently skips them.

Options:
  --at tcN          Accept a single tracked change by id
  --all             Accept every tracked change (mutually exclusive with --at)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -h, --help        Show this help

Examples:
  docx track-changes accept doc.docx --at tc0
  docx track-changes accept doc.docx --all
  docx track-changes accept doc.docx --all --dry-run
`;

export async function run(args: string[]): Promise<number> {
	return runApply(args, "accept", HELP);
}
