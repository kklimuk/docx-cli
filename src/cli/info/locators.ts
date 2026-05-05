import { parseArgs } from "util";
import { EXIT, fail, writeStdout } from "../respond";

const HELP = `docx info locators — print the locator grammar reference

Usage:
  docx info locators [options]

Options:
  --json       Print as JSON
  -h, --help   Show this help

Examples:
  docx info locators
  docx info locators --json | jq '.entityLocators'
`;

const REFERENCE = `LOCATOR GRAMMAR

Block locators:
  pN              Paragraph N (e.g., p3)
  tN              Table N
  sN              Section break N
  tT:rRcC:pK      Paragraph K of cell (R,C) in table T

Span locators (within a single paragraph):
  pN:S-E          Characters S..E of paragraph N (inclusive start, exclusive end)
  tT:rRcC:pK:S-E  Characters S..E of a cell paragraph

Range locators (across blocks):
  pN:S-pM:E       From char S of paragraph N to char E of paragraph M

Entity locators:
  cN              Comment id (e.g., c0)
  imgN            Image id (e.g., img2)
  linkN           Hyperlink id (e.g., link0)
  tN:rRcC         Cell at row R, column C of table tN

Examples:
  p3              -> the entire paragraph p3
  p3:5-20         -> characters 5..20 of p3
  p3:5-p5:10      -> from char 5 of p3 to char 10 of p5
  c1              -> comment c1
  img0            -> image img0
  link0           -> hyperlink link0
  t0:r1c2         -> cell at row 1, col 2 of table t0
  t0:r1c2:p0      -> first paragraph of that cell
  t0:r1c2:p0:5-10 -> chars 5..10 of that paragraph

Notes:
  Block ids are positional and shift after structural edits — re-read
  between non-trivial mutations. Mutating commands print before/after
  locator info in their JSON ack so agents can chain operations.
`;

const JSON_REFERENCE = {
	blockLocators: {
		paragraph: { syntax: "pN", example: "p3" },
		table: { syntax: "tN", example: "t0" },
		sectionBreak: { syntax: "sN", example: "s0" },
		cellParagraph: { syntax: "tT:rRcC:pK", example: "t0:r1c2:p0" },
	},
	spanLocator: {
		syntax: "pN:S-E",
		example: "p3:5-20",
		semantics:
			"Characters S..E within paragraph N (start inclusive, end exclusive)",
		cellSyntax: "tT:rRcC:pK:S-E",
	},
	rangeLocator: {
		syntax: "pN:S-pM:E",
		example: "p3:5-p5:10",
		semantics: "From char S of paragraph N to char E of paragraph M",
	},
	entityLocators: {
		comment: { syntax: "cN", example: "c1" },
		image: { syntax: "imgN", example: "img0" },
		hyperlink: { syntax: "linkN", example: "link0" },
		cell: { syntax: "tN:rRcC", example: "t0:r1c2" },
		nestedCell: { syntax: "tN:rRcC:pK", example: "t0:r1c2:p0" },
	},
	notes: [
		"Block ids are positional and shift after structural edits.",
		"Re-read between non-trivial mutations.",
		"Mutating commands print before/after locator info in their JSON ack.",
	],
};

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				json: { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	if (parsed.values.json) {
		await writeStdout(`${JSON.stringify(JSON_REFERENCE, null, 2)}\n`);
		return EXIT.OK;
	}

	await writeStdout(REFERENCE);
	return EXIT.OK;
}
