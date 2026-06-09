import { describeForms } from "@core";
import { runApply } from "./apply";

const AT_FORMS = describeForms(["trackedChange"], "                    ");

const HELP = `docx track-changes accept — accept tracked changes (incorporate into the doc)

Usage:
  docx track-changes accept FILE --at tcN [options]
  docx track-changes accept FILE --all [options]

Accepting an insertion (<w:ins>) or move destination (<w:moveTo>) unwraps the
wrapper — the content becomes plain runs at this location. Accepting a
deletion (<w:del>) or move source (<w:moveFrom>) removes the element and its
<w:delText> entirely (the text disappears for real).

Section-property revisions (<w:sectPrChange>): accept drops the prior-state
snapshot, leaving the live section properties in place.

Paragraph-mark trackings (<w:ins>/<w:del> inside <w:pPr><w:rPr>): accepting
a paragraph-mark insertion just removes the marker (the inserted paragraph
stays as a regular paragraph). Accepting a paragraph-mark deletion merges
the owning paragraph with the next paragraph — the next paragraph's runs
are appended to this one and the next paragraph is removed (per ECMA-376
§17.13.5.4).

Out of scope: formatting changes (<w:rPrChange>/<w:pPrChange>). These aren't
modeled in the AST today and --all silently skips them.

Target (one required, mutually exclusive):
  --at tcN          Accept a tracked change by id. Repeat for multiple ids
                    (--at tc1 --at tc2 --at tc3) — all targets are resolved
                    against the pre-mutation tree, so renumbering during the
                    batch is not a concern. Supports:
${AT_FORMS}
                    See \`docx info locators\`.
  --all             Accept every tracked change.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints {ok:true, operation, path,
  applied}. --dry-run prints the preview {operation, dryRun, path, applied}.
  Errors print {code, error, hint?} with a nonzero exit. Discover ids with
  \`docx track-changes list FILE\`.

Examples:
  docx track-changes accept doc.docx --at tc0
  docx track-changes accept doc.docx --at tc1 --at tc3 --at tc5
  docx track-changes accept doc.docx --all
  docx track-changes accept doc.docx --all --dry-run
`;

export async function run(args: string[]): Promise<number> {
	return runApply(args, "accept", HELP);
}
