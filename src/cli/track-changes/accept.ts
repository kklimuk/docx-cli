import { describeForms } from "@core";
import { runApply } from "./run-apply";

const AT_FORMS = describeForms(["trackedChange"], "                    ");

const HELP = `docx track-changes accept — accept tracked changes (incorporate into the doc)

Usage:
  docx track-changes accept FILE --at tcN [options]
  docx track-changes accept FILE --all [options]

Examples:
  docx track-changes list doc.docx                 # get the tcN / revN handles
  docx track-changes accept doc.docx --at rev0     # one del+ins pair, one call
  docx track-changes accept doc.docx --at tc1 --at tc3 --at tc5
  docx track-changes accept doc.docx --all
  docx track-changes accept doc.docx --all --dry-run

Accepting works exactly like Word: an accepted insertion stays as plain
text; an accepted deletion disappears for real; an accepted section/
paragraph-property change keeps the new properties; an accepted
paragraph-mark deletion merges the paragraph with the next one. To accept
some and reject the rest in ONE call, use \`docx track-changes apply\`.

Out of scope: run-formatting changes Word tracked (bold/color tweaks) aren't
modeled; --all silently skips them.

Target (one required, mutually exclusive):
  --at tcN          Accept a tracked change by id. Repeat for multiple ids
                    (--at tc1 --at tc2 --at tc3) — all targets are resolved
                    against the pre-mutation tree, so renumbering during the
                    batch is not a concern. Supports:
${AT_FORMS}
                    See \`docx info locators\`.
  --at revN         Accept a del+ins REPLACE pair in one call (both halves of one
                    logical change). \`list\` tags the two tcNs with a shared
                    "group": "revN"; addressing the revN saves the accept-relist-
                    accept ping-pong (tcN ids renumber after each single accept).
  --all             Accept every tracked change.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would change; do not write the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints
  {ok:true, operation, path, applied}. --dry-run prints the preview
  {operation, dryRun, path, applied}. Errors print {code, error, hint?} with
  a nonzero exit. Discover ids with \`docx track-changes list FILE\`.
`;

export async function run(args: string[]): Promise<number> {
	return runApply(args, "accept", HELP);
}
