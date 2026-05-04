import packageJson from "../../package.json" with { type: "json" };
import { writeStdout } from "./respond";

export const VERSION = packageJson.version;

const TOP_HELP = `docx ${VERSION} — read, edit, and comment on .docx files

Usage:
  docx <command> [options]

Commands:
  create    FILE  Create a new minimal .docx
  read      FILE  Print AST as JSON
  edit      FILE  Replace text/properties at a locator
  insert    FILE  Insert a block or run
  delete    FILE  Remove a block or run range
  find      FILE QUERY  Find text spans, return locators
  replace   FILE PATTERN REPL  Substitute text spans (sed for docx)
  comments  …     Add, reply, resolve, delete, list comments
  images    …     Extract, replace, list images
  track-changes FILE on|off  Toggle tracked-changes mode
  info      …     Reference material (schema, locators)

Run "docx <command> --help" for command-specific help.

Environment:
  DOCX_AUTHOR    Default author for comments and tracked-change attribution
  DOCX_CLI_NOW   Override the timestamp used for tracked changes (test only)

Tracked changes:
  When <w:trackChanges/> is set in the doc (toggle via "docx track-changes
  FILE on"), insert/edit/delete/replace automatically emit <w:ins>/<w:del>
  markers attributed to $DOCX_AUTHOR.
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
