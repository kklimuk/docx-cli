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
  insert    FILE  Insert a paragraph, image, table, equation, code block, markdown, or page break (--after/--before LOCATOR; --track; --batch for many inserts in one read). For COLUMN layout use "docx sections", not insert.
  delete    FILE  Remove a paragraph, range, table, or section break (--at LOCATOR; --track for tracked deletion; --batch to remove many in one read)
  find      FILE [QUERY]  Find spans by text, OR by formatting (--highlight/--color/--bold/--italic/--underline); returns locators for --at
  replace   FILE PATTERN REPL  Substitute text spans, sed-style (--regex, --track to redline, --dry-run to preview, --batch for a multi-pattern script)
  wc        FILE [LOCATOR]  Count words in the doc or a slice (--accepted/--baseline/--current tracked view, --json)
  outline   FILE  List headings as a locator tree (pN feeds --at / read --from; --style-prefix, --json)
  sections  FILE  Multi-column layout & section breaks — put a paragraph range in N columns (--at pN-pM --columns N) or recount a section (--at sN). The ONLY way to do columns; insert does not.
  styles    FILE  List the styles you can apply (--used for ones in use; --at ID to describe one) — the catalog isn't in the body
  render    FILE  Visual page verification — render each page as PNG/JPG via Word or LibreOffice
  comments  …     Add (--at LOCATOR | --anchor PHRASE | --batch), reply, resolve (--unset to reopen), delete, list (--thread cN)
  footnotes …     Add (--at | --anchor PHRASE), edit, delete, list footnotes (--text/--runs/--markdown bodies)
  endnotes  …     Add (--at | --anchor PHRASE), edit, delete, list endnotes (--text/--runs/--markdown bodies)
  images    …     Add (--caption "Figure 1: …" for a captioned figure), extract, replace, delete, list images
  hyperlinks …    Add, list, replace, delete hyperlinks (add uses --url; replace uses --with)
  tables    …     Restructure tables — insert/delete rows & columns, merge/unmerge, set widths, borders
  track-changes …  Toggle (on|off FILE); list / accept / reject revisions; "read" shows them as CriticMarkup
  info      …     Reference material, no FILE needed (schema for read --ast, locator grammar)

It is highly recommended to run "docx info locators" and "docx info schema" (neither needs a FILE) to understand the addressing model and AST.
Run "docx <command> --help" for command-specific help.

BATCH MANY CHANGES IN ONE READ: filling a form or applying many edits? Don't go
one-at-a-time — edit / insert / replace / delete and comments (add/resolve/delete)
all take --batch FILE.jsonl (one JSON change per line; "-" reads stdin). Every locator
addresses the document AS READ, so ids stay valid across the whole batch — one
read, one write, no re-reading between changes. See "<command> --help".

VERIFY LAYOUT VISUALLY — ONLY WHEN LAYOUT IS THE QUESTION: "docx read" is the source of
truth for CONTENT, so if you filled text, replaced placeholders, edited cells, or added
comments / tracked changes, "read" plus the write→read loop already prove it — do NOT
render (each render spins up Word and is slow). Render only for what Markdown can't show:
multi-column sections, page/section breaks, image sizing/placement, table geometry — and
then ONCE at the end (not after every edit), or one final time if you're genuinely unsure
it looks right: "docx render FILE --out pages/" writes page-001.png, … which you can read.
(To put text in columns, name the range: "docx sections --at pN-pM --columns N" — it inserts the bounding breaks so the columns land on exactly that range. A raw section break's columns apply to the content BEFORE it, which is why insert no longer takes --section.)

Environment:
  DOCX_AUTHOR    Default author for comments and tracked-change attribution
  DOCX_CLI_NOW   Override the timestamp used for tracked changes (test only)

Tracked changes (full detail: "docx track-changes --help"):
  docx track-changes on|off FILE  ·  list  ·  accept/reject (--at tcN | --all)
  With tracking on, edit/insert/delete/replace auto-emit <w:ins>/<w:del>
  (attributed via --author / $DOCX_AUTHOR); pass --track to any one of them (plus
  the tables verbs / images delete) to redline just that invocation. "docx read"
  shows revisions as CriticMarkup with --current ({++ins++}/{--del--}), --baseline
  for the pre-change text, or --accepted (default, clean); "docx wc" mirrors them.
`;

export async function printTopHelp(): Promise<void> {
	await writeStdout(TOP_HELP);
}
