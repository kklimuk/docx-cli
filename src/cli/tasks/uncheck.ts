import { toggleTask } from "./toggle";

const HELP = `docx tasks uncheck — clear an existing task's checkbox (☐)

Usage:
  docx tasks uncheck FILE --at pN [options]

Locator (required):
  --at pN           The task-list paragraph to uncheck. It must already be a GFM
                    task item (a leading <w:sdt><w14:checkbox/> SDT). Discover
                    ids with \`docx read FILE\` (task lines render as - [ ] / - [x]).

Options:
  --track           Record the toggle as a tracked change even when the
                    document's track-changes toggle is off. Surfaces via
                    \`docx track-changes list\` as a checkboxToggle revision.
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Examples:
  docx tasks uncheck doc.docx --at p3
  docx tasks uncheck doc.docx --at p3 --track --author "Reviewer"

Output:
  Prints a one-line confirmation on success (exit 0) — an in-place toggle shifts
  nothing, so the pN is unchanged. --verbose prints {ok:true, operation:"edit",
  …}. Errors print {code, error, hint?} with a nonzero exit.
`;

export async function run(args: string[]): Promise<number> {
	return toggleTask(args, false, HELP);
}
