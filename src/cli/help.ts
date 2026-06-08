import packageJson from "../../package.json" with { type: "json" };
import { writeStdout } from "./respond";

export const VERSION = packageJson.version;

const TOP_HELP = `docx ${VERSION} — read, edit, and comment on .docx files

Usage:
  docx <command> [options]

Commands (each one-liner names capabilities you'd otherwise miss; see <command> --help):
  create    FILE  Create a new .docx (--from PATH.md | --from - builds from Markdown; --force to overwrite)
  read      FILE  Render as Markdown with pN locators; --from/--to to slice, --accepted (default)/--current/--baseline tracked views, --comments, --ast for JSON-AST
  edit      FILE  Replace or strip text/formatting at pN, pN:S-E, pN-pM, sN, eqN, or table-cell locators (--clear to strip formatting, --track to redline, --batch for many edits in one read)
  insert    FILE  Insert a paragraph, image, table, equation, code block, markdown, or structural break (--after/--before LOCATOR; --track; --batch for many inserts in one read)
  delete    FILE  Remove a paragraph, range, table, or section break (--at LOCATOR; --track for tracked deletion)
  find      FILE [QUERY]  Find spans by text, OR by formatting (--highlight/--color/--bold/--italic/--underline); returns locators for --at
  replace   FILE PATTERN REPL  Substitute text spans, sed-style (--regex, --track to redline, --dry-run to preview, --batch for a multi-pattern script)
  wc        FILE [LOCATOR]  Count words in the doc or a slice (--accepted/--baseline/--current tracked view, --json)
  outline   FILE  List headings as a locator tree (pN feeds --at / read --from; --style-prefix, --json)
  render    FILE  Visual page verification — render each page as PNG/JPG via Word or LibreOffice
  comments  …     Add (--at LOCATOR | --anchor PHRASE | --batch), reply, resolve (--unset to reopen), delete, list (--thread cN)
  footnotes …     Add (--at | --anchor PHRASE), edit, delete, list footnotes (--text/--runs/--markdown bodies)
  endnotes  …     Add (--at | --anchor PHRASE), edit, delete, list endnotes (--text/--runs/--markdown bodies)
  images    …     Add (--caption "Figure 1: …" for a captioned figure), extract, replace, delete, list images
  hyperlinks …    Add, list, replace, delete hyperlinks (add uses --url; replace uses --with)
  tables    …     Restructure tables — insert/delete rows & columns, merge/unmerge, set widths, borders
  track-changes …  Toggle (FILE on|off); list / accept / reject revisions; "read" shows them as CriticMarkup
  info      …     Reference material, no FILE needed (schema for read --ast, locator grammar)

It is highly recommended to run "docx info locators" and "docx info schema" (neither needs a FILE) to understand the addressing model and AST.
Run "docx <command> --help" for command-specific help.

VERIFY LAYOUT VISUALLY: "docx read" shows text/structure as Markdown but NOT how the
page looks — multi-column sections, page breaks, image sizing, and where content lands
on the page don't appear there. After authoring or inserting layout-affecting content,
render to images and look: "docx render FILE --out pages/" writes page-001.png, … which
you can read. Adjust and re-render until it looks right. (Note: a multi-column section's
columns apply to the content BEFORE the section break, not after — see "docx insert --help".)

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
  <w:ins>/<w:del> markers attributed via --author (or $DOCX_AUTHOR, or "Reviewer").
  Pass --track to any insert/edit/delete/replace (and the tables verbs / images
  delete) to record just that one invocation as tracked even when the toggle is
  off; when the toggle is already on, tracking is automatic.
  edit --at sN under tracking emits <w:sectPrChange> snapshots on the section.

  Accept/reject handle: run-level ins/del/moveFrom/moveTo, sectPrChange
  (snapshot drop or restore), paragraph-mark ins/del (accept-del merges with the
  next paragraph; reject-ins removes the owning paragraph), the table-structural
  revisions (rowIns/rowDel/cellIns/cellDel/tblGridChange/tblPrChange/tcPrChange),
  and checkboxToggle. Out of scope: rPrChange / pPrChange formatting revisions.

  "docx read" default view is --accepted (clean text: insertions inlined,
  deletions dropped). Add --current for CriticMarkup:
    {++inserted++}[^tcN]   {--deleted--}[^tcN]
  with a [^tcN]: definition appendix; --baseline renders the pre-change text.
  "docx wc" mirrors the same --accepted / --baseline / --current flags.
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
