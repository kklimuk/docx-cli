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
  pN-pM           Whole-paragraph range pN..pM (inclusive) — used by
                  edit, delete, wc to operate on a contiguous span as a unit
  pN:S-pM:E       From char S of paragraph N to char E of paragraph M

Entity locators:
  cN              Comment id (e.g., c0)
  imgN            Image id (e.g., img2)
  linkN           Hyperlink id (e.g., link0)
  eqN             Equation id (e.g., eq3) — the Nth <m:oMath> / <m:oMathPara>
                  in document order. Surfaces with run.latex = the
                  reconstructed LaTeX (see core/equation).
  tcN             Tracked change id (e.g., tc0) — the Nth revision wrapper
                  in document order. Includes run-level <w:ins>/<w:del>/
                  <w:moveFrom>/<w:moveTo>, section-property revisions
                  <w:sectPrChange>, and paragraph-mark <w:ins>/<w:del>
                  (a self-closing element inside <w:pPr><w:rPr>).
  tN:rRcC         Cell at row R, column C of table tN

Table structure locators (for the "docx tables" verbs):
  tN:rR           Row R of table tN (e.g. delete-row)
  tN:cC           Column C of table tN (e.g. delete-column)
  tN:rR1cC1-rR2cC2   Rectangular cell region, top-left to bottom-right by
                  logical row/column (e.g. merge)

Examples:
  p3              -> the entire paragraph p3
  p3:5-20         -> characters 5..20 of p3
  p3-p7           -> paragraphs p3..p7 inclusive (as a unit)
  p3:5-p5:10      -> from char 5 of p3 to char 10 of p5
  s0              -> first section break (typically the front-matter section)
  c1              -> comment c1
  img0            -> image img0
  link0           -> hyperlink link0
  tc0             -> first tracked change in the document
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
	blockRangeLocator: {
		syntax: "pN-pM",
		example: "p3-p7",
		semantics:
			"Whole paragraphs pN..pM inclusive, as a unit. Accepted by edit, delete, and wc; endpoints must share a parent (no cross-cell ranges).",
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
		equation: {
			syntax: "eqN",
			example: "eq0",
			notes:
				"Equation in document order; surfaces with `run.latex` (reconstructed LaTeX) and `run.display` (inline vs $$…$$).",
		},
		trackedChange: {
			syntax: "tcN",
			example: "tc0",
			notes:
				"Covers run-level <w:ins>/<w:del>/<w:moveFrom>/<w:moveTo>, <w:sectPrChange>, and paragraph-mark <w:ins>/<w:del>.",
		},
		cell: { syntax: "tN:rRcC", example: "t0:r1c2" },
		nestedCell: { syntax: "tN:rRcC:pK", example: "t0:r1c2:p0" },
	},
	tableStructureLocators: {
		row: { syntax: "tN:rR", example: "t0:r1" },
		column: { syntax: "tN:cC", example: "t0:c2" },
		cellRange: { syntax: "tN:rR1cC1-rR2cC2", example: "t0:r0c0-r1c1" },
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
