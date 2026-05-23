import packageJson from "../../package.json" with { type: "json" };
import { writeStdout } from "./respond";

export const VERSION = packageJson.version;

const TOP_HELP = `docx ${VERSION} — read, edit, and comment on .docx files

Usage:
  docx <command> [options]

Commands:
  create    FILE  Create a new minimal .docx
  read      FILE  Render body as Markdown (--ast for JSON)
  edit      FILE  Replace text/properties at a locator
  insert    FILE  Insert a block or run
  delete    FILE  Remove a block or run range
  find      FILE QUERY  Find text spans, return locators
  replace   FILE PATTERN REPL  Substitute text spans (sed for docx)
  wc        FILE [LOCATOR]  Count words in the doc or a slice
  outline   FILE  List headings as a hierarchical tree
  comments  …     Add, reply, resolve, delete, list comments
  footnotes …     Add, edit, delete, list footnotes
  endnotes  …     Add, edit, delete, list endnotes
  images    …     Extract, replace, list images
  hyperlinks …    Add, list, replace, delete hyperlinks
  track-changes …  Toggle (FILE on|off), list / accept / reject changes
  info      …     Reference material (schema, locators)

It is highly recommended to agents to run "docx info locators" to understand their capabilities.
Run "docx <command> --help" for command-specific help.

Environment:
  DOCX_AUTHOR    Default author for comments and tracked-change attribution
  DOCX_CLI_NOW   Override the timestamp used for tracked changes (test only)

Tracked changes:
  Toggle:    docx track-changes FILE on|off
  Inventory: docx track-changes list FILE  (JSON array of { id, kind, author,
             date, revisionId, blockId, text }; sectPrChange entries also
             include { prior, current } with the section properties on each
             side of the edit)
  Accept:    docx track-changes accept FILE (--at tcN | --all)
  Reject:    docx track-changes reject FILE (--at tcN | --all)

  When <w:trackChanges/> is set, insert/edit/delete/replace automatically emit
  <w:ins>/<w:del> markers attributed via --author (or $DOCX_AUTHOR, or "docx-cli").
  edit --at sN under tracking emits <w:sectPrChange> snapshots on the section.

  Accept/reject handle: run-level ins/del/moveFrom/moveTo, sectPrChange
  (snapshot drop or restore), and paragraph-mark ins/del (accept-del merges
  with the next paragraph; reject-ins removes the entire owning paragraph).
  Out of scope: rPrChange / pPrChange formatting revisions.

  "docx read" surfaces them as CriticMarkup:
    {++inserted++}[^tcN]   {--deleted--}[^tcN]
  with a [^tcN]: definition appendix. Switch view with --accepted (post-accept:
  drop del, ins as plain) or --baseline (pre-change: drop ins, del as plain).
  "docx wc" mirrors the same --accepted / --baseline flags.
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
