import { writeStdout } from "./respond";

export const VERSION = "0.1.0";

const TOP_HELP = `docx ${VERSION} — read, edit, and comment on .docx files

Usage:
  docx <command> [options]

Commands:
  create    FILE  Create a new minimal .docx
  read      FILE  Print AST as JSON
  edit      FILE  Replace text/properties at a locator
  insert    FILE  Insert a block or run
  delete    FILE  Remove a block or run range
  comments  …     Add, reply, resolve, delete, restore, list comments
  images    …     Extract, replace, list images
  track-changes FILE on|off  Toggle tracked-changes mode
  schema          Dump AST JSON Schema
  locators        Dump locator grammar reference

Run "docx <command> --help" for command-specific help.

Environment:
  DOCX_AUTHOR  Default --author for "comments add" / "comments reply"
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
