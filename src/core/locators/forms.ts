/** The single source of truth for locator-form documentation. `info locators`
 *  renders its REFERENCE / JSON_REFERENCE from this registry, and every
 *  locator-taking command builds its accepted-forms help block via
 *  `describeForms(keys)` — so the docs can't drift from each other. The forms
 *  themselves are parsed in `parse.ts`; this file only describes them. */

export type LocatorFormKey =
	| "paragraph"
	| "table"
	| "section"
	| "span"
	| "blockRange"
	| "crossSpan"
	| "cell"
	| "cellParagraph"
	| "cellSpan"
	| "tableRow"
	| "tableColumn"
	| "cellRange"
	| "comment"
	| "image"
	| "hyperlink"
	| "equation"
	| "trackedChange"
	| "footnote"
	| "endnote";

export type LocatorForm = {
	syntax: string;
	example: string;
	summary: string;
};

export const LOCATOR_FORMS: Record<LocatorFormKey, LocatorForm> = {
	paragraph: { syntax: "pN", example: "p3", summary: "whole paragraph N" },
	table: { syntax: "tN", example: "t0", summary: "whole table N" },
	section: { syntax: "sN", example: "s0", summary: "section break N" },
	span: {
		syntax: "pN:S-E",
		example: "p3:5-20",
		summary: "chars S..E of paragraph N (start inclusive, end exclusive)",
	},
	blockRange: {
		syntax: "pN-pM",
		example: "p3-p7",
		summary: "paragraphs N..M inclusive, as one unit",
	},
	crossSpan: {
		syntax: "pN:S-pM:E",
		example: "p3:5-p5:10",
		summary: "from char S of paragraph N to char E of paragraph M",
	},
	cell: {
		syntax: "tN:rRcC",
		example: "t0:r1c2",
		summary: "cell at row R, column C of table N",
	},
	cellParagraph: {
		syntax: "tN:rRcC:pK",
		example: "t0:r1c2:p0",
		summary: "paragraph K of that cell (chainable to any nesting depth)",
	},
	cellSpan: {
		syntax: "tN:rRcC:pK:S-E",
		example: "t0:r1c2:p0:5-10",
		summary: "chars S..E of a cell paragraph",
	},
	tableRow: {
		syntax: "tN:rR",
		example: "t0:r1",
		summary: "row R of table N",
	},
	tableColumn: {
		syntax: "tN:cC",
		example: "t0:c2",
		summary: "column C of table N",
	},
	cellRange: {
		syntax: "tN:rR1cC1-rR2cC2",
		example: "t0:r0c0-r1c1",
		summary: "rectangular cell region, top-left to bottom-right",
	},
	comment: { syntax: "cN", example: "c0", summary: "comment id" },
	image: {
		syntax: "imgN",
		example: "img0",
		summary: "image id (document order)",
	},
	hyperlink: {
		syntax: "linkN",
		example: "link0",
		summary: "hyperlink id (document order)",
	},
	equation: {
		syntax: "eqN",
		example: "eq0",
		summary: "equation id (document order; surfaces run.latex)",
	},
	trackedChange: {
		syntax: "tcN",
		example: "tc0",
		summary: "tracked-change id (document order)",
	},
	footnote: { syntax: "fnN", example: "fn0", summary: "footnote id" },
	endnote: { syntax: "enN", example: "en0", summary: "endnote id" },
};

/** Render an aligned `syntax  summary` block for a command's accepted subset,
 *  for embedding in a HELP string. Syntax column is padded to the widest entry
 *  so the summaries line up. `indent` prefixes every line. */
export function describeForms(keys: LocatorFormKey[], indent = "  "): string {
	const rows = keys.map((key) => LOCATOR_FORMS[key]);
	const width = Math.max(...rows.map((form) => form.syntax.length));
	return rows
		.map((form) => `${indent}${form.syntax.padEnd(width)}  ${form.summary}`)
		.join("\n");
}
