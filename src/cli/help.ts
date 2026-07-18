import packageJson from "../../package.json" with { type: "json" };
import { writeStdout } from "./respond";

export const VERSION = packageJson.version;

/** The shared "reach for the dedicated noun" tail on the `insert` and `edit`
 * one-liners — one home, so adding/renaming a noun is a single edit. */
const NOUN_REDIRECTS = `For a CODE block use "docx code", an EQUATION "docx equations", a TASK checkbox "docx tasks", an IMAGE "docx images", a TABLE "docx tables", for LAYOUT "docx sections".`;

const TOP_HELP = `docx ${VERSION} — read, edit, and comment on .docx files

Usage:
  docx <command> [options]

Commands (each one-liner names capabilities you'd otherwise miss; see <command> --help):
  create    FILE  Create a new .docx (--from PATH.md | --from - builds from Markdown; --text-file for literal text; --force to overwrite)
  read      FILE  Render as Markdown with pN locators; --from/--to to slice; --comments for comment bodies; --current shows tracked changes inline; --ast for lossless JSON.
  insert    FILE  Insert content at a locator (--after/--before LOCATOR; --track; --batch for many inserts in one read). ${NOUN_REDIRECTS}
  find      FILE [QUERY]  Find content by text, OR by formatting; returns locators for \`insert\`, \`replace\`, \`edit\`, \`delete\`.
  replace   FILE PATTERN REPL  Replace content, sed-style. KEEPS the run's formatting and any tabs (--regex, --track to redline, --dry-run to preview, --batch for a multi-pattern fill).
  edit      FILE  Replace or strip content/formatting at a locator (--clear to strip formatting, --track to redline, --batch for many edits in one read). ${NOUN_REDIRECTS}
  delete    FILE  Remove content at a locator (--at LOCATOR; --track for tracked deletion; --batch to remove many in one read)
  outline   FILE  List headings as a locator tree (pN feeds --at / read --from; --style-prefix, --json)
  wc        FILE [LOCATOR]  Count words in the doc or a slice (--accepted/--baseline/--current tracked view, --json)
  sections  FILE  Multi-column layout, section breaks & PAGE SETUP — columns on a range (--at pN-pM --columns N); page margins/orientation/size for the WHOLE document ("--margins 0.5" with NO --at sets every section) or one section (--at sN). The ONLY way to do columns; insert does not.
  render    FILE  Visual page verification: render each page as PNG/JPG via Word or LibreOffice
  diff      FILE  Change verification: Show what changed vs another version. A git-style unified diff of the read view. Snapshot BEFORE editing ("cp doc.docx doc.orig.docx"), then "diff doc.docx --against doc.orig.docx" (--against also takes a saved read output, or - for stdin / a git branch)
  styles    …     Author or inspect styles. List/describe styles (--used, --at ID); "styles set --at Heading1 --color 1F4E79 --bold" restyles every heading; "styles create" mints one; "set-default-font" sets the doc font — the catalog isn't in the body
  code      …     Author or replace a syntax-highlighted code block ("code add FILE --after pN --code-file snippet.py --language go" / "code edit --at pN")
  equations …     Insert or edit a LaTeX equation ("equations add FILE --after pN --equation x^2" / "equations edit --at eqN")
  tasks     …     Author or toggle a GFM task-list checkbox ("tasks add FILE --after pN --text '…' [--checked]" / "tasks check --at pN" / "tasks uncheck --at pN")
  comments  …     Add (--at LOCATOR | --anchor PHRASE | --batch), reply, resolve (--unset to reopen), delete, list (--thread cN)
  footnotes …     Add (--at | --anchor PHRASE), edit, delete, list footnotes (--text/--runs/--markdown bodies)
  endnotes  …     Add (--at | --anchor PHRASE), edit, delete, list endnotes (--text/--runs/--markdown bodies)
  headers   …     Set/list/clear page headers — --text, page numbers, date, fields; --first-page/--even; default = whole document
  footers   …     Set/list/clear page footers — e.g. "footers set FILE --page-number --of-pages" for "Page X of Y"
  images    …     Add (--caption "Figure 1: …" for a captioned figure), extract, replace, delete, list images
  hyperlinks …    Add, list, replace, delete hyperlinks (add uses --url; replace uses --with)
  tables    …     Create a table (create --after pN --rows N --cols M), then restructure — insert/delete rows & columns, merge/unmerge, set widths, borders, format
  lists     FILE  Renumber a numbered list — "lists set --at pN --start 5" / "--format upper-roman" / "--restart" / "--continue"
  track-changes …  Toggle (on|off FILE); list / accept / reject; "apply" finalizes a whole review (accept some + reject the rest) in ONE call; "read --current" shows changes inline
  raw       …     LAST-RESORT escape hatch for OOXML no verb above covers. DO NOT USE THIS unless you've tried everything else.
  validate  FILE  Schema-check the document against the bundled ECMA-376 transitional XSDs (per-part errors; exit 0 = clean)
  info      …     Reference material, no FILE needed (schema for read --ast, locator grammar)

It is HIGHLY RECOMMENDED to run \`docx info locators\` to understand the addressing model.
Run "docx <command> --help" for command-specific help.

BATCH MANY CHANGES AFTER ONE READ: locator ids are positional and shift after structural
edits (insert/delete/section changes), so going one-at-a-time forces a re-read to refresh
ids after each. Miss the change and the next command lands on the wrong block or errors
BLOCK_NOT_FOUND. Skip all that: edit / insert / replace / delete and comments
(add/resolve/delete) all take --batch FILE.jsonl (one JSON change per line; "-" reads
stdin). Every locator addresses the document AS READ, so ids stay valid across the whole
batch — one read, one write, no re-reading between changes. See "<command> --help".

VERIFY LAYOUT VISUALLY, ONLY WHEN LAYOUT IS THE QUESTION: "docx read" is the source of
truth for CONTENT, so if you filled text, replaced placeholders, edited cells, or added
comments / tracked changes, "read" plus the write→read loop already prove it — do NOT
render (each render spins up Word and is slow). Render only for what Markdown can't show:
multi-column sections, page/section breaks, image sizing/placement, table geometry — and
then ONCE at the end (not after every edit), or one final time if you're genuinely unsure
it looks right: \`docx render FILE --out pages/\` writes page-001.png, … which you can read.
(To put text in columns, name the range: "docx sections --at pN-pM --columns N" — it inserts the bounding breaks so the columns land on exactly that range. A raw section break's columns apply to the content BEFORE it, which is why insert no longer takes --section.)

Environment:
  DOCX_AUTHOR    Default author for comments and tracked-change attribution
  DOCX_CLI_NOW   Override the timestamp used for tracked changes (test only)
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
