import { runStylesCreate } from "./create";
import { runStylesRead } from "./read";
import { runStylesSet } from "./set";
import { runSetDefaultFont } from "./set-default-font";

const HELP = `docx styles — inspect, edit, and create the styles a document can apply

Usage:
  docx styles FILE [--used] [--at STYLEID] [--json]     # list / describe (read)
  docx styles --catalog [--json]                        # built-ins, no FILE
  docx styles set FILE --at STYLEID [formatting]        # edit a style definition
  docx styles create FILE STYLEID [--type …] [formatting]   # define a new style
  docx styles set-default-font FILE "Font Name" [--size N] [--all]

The style catalog lives in word/styles.xml, not the document body — so unlike
everything else, you can't see it by reading the doc. Read it to learn what
\`--style NAME\` values exist (for \`insert --style\` / \`edit --style\`); edit a
definition to restyle every paragraph that uses it ("make all Heading 1s green").

Subcommands:
  (read)             docx styles FILE — list defined styles; --at describes one;
                     --used filters to applied styles; --catalog lists built-ins
  set                Change an existing style's formatting/metadata in place
  create             Define a new custom paragraph or character style
  set-default-font   Set the document-wide default font (styles.xml + theme)

Read options:
  --at STYLEID   Describe one style (id, type, name, basedOn, full formatting)
  --used         List only the styles actually applied somewhere in the body
  --catalog      List the built-in styles docx-cli can apply on demand (Title,
                 Subtitle, Heading1–9, Quote, IntenseQuote, Code, …). No FILE.
  --json         Structured output (a JSON array for the list; object for --at)
  -h, --help     Show this help

Examples:
  docx styles report.docx                 # styles defined in this doc
  docx styles report.docx --at Heading1   # what does Heading1 look like?
  docx styles --catalog                    # built-ins you can apply via --style
  docx styles set report.docx --at Heading1 --color 1F4E79 --size 16 --bold
  docx styles create report.docx Callout --color C00000 --bold --size 12
  docx styles set-default-font report.docx "Times New Roman"

See \`docx styles set --help\`, \`docx styles create --help\`, and
\`docx styles set-default-font --help\` for each mutator's details.
`;

export async function run(args: string[]): Promise<number> {
	// `styles` is read-by-default; the three mutating subverbs branch off the first
	// positional. Keeps the read command from going dual-natured while each mutator
	// inherits --dry-run / -o / the ack confirmation as a normal mutator.
	if (args[0] === "set-default-font") return runSetDefaultFont(args.slice(1));
	if (args[0] === "set") return runStylesSet(args.slice(1));
	if (args[0] === "create") return runStylesCreate(args.slice(1));
	return runStylesRead(args, HELP);
}
