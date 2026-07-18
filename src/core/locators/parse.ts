export type Locator =
	| { kind: "block"; blockId: string }
	| { kind: "blockSpan"; blockId: string; start: number; end: number }
	| { kind: "blockRange"; startBlockId: string; endBlockId: string }
	| {
			kind: "range";
			start: { blockId: string; offset: number };
			end: { blockId: string; offset: number };
	  }
	| { kind: "comment"; commentId: string }
	| { kind: "image"; imageId: string }
	| { kind: "hyperlink"; hyperlinkId: string }
	| { kind: "trackedChange"; trackedChangeId: string }
	| { kind: "equation"; equationId: string }
	| { kind: "footnote"; footnoteId: string }
	| { kind: "endnote"; endnoteId: string }
	| { kind: "marginal"; marginalId: string }
	| {
			kind: "cell";
			tableId: string;
			row: number;
			col: number;
			inner?: Locator;
	  }
	| { kind: "tableRow"; tableId: string; row: number }
	| { kind: "tableColumn"; tableId: string; col: number }
	| {
			kind: "cellRange";
			tableId: string;
			start: { row: number; col: number };
			end: { row: number; col: number };
	  };

export class LocatorParseError extends Error {
	constructor(
		public input: string,
		message: string,
	) {
		super(`Invalid locator "${input}": ${message}`);
		this.name = "LocatorParseError";
	}
}

const BLOCK_RE = /^(p|t|s)(\d+)$/;
const SPAN_RE = /^p(\d+):(\d+)-(\d+)$/;
// Each side is a body paragraph (`p3`) or a cell paragraph (`t0:r1c2:p0`) —
// the form `find` prints for a spanning match, so it pipes into `comments add
// --at`. Both sides must share a container (validated after the match).
const RANGE_RE =
	/^((?:t\d+:r\d+c\d+:)?p\d+):(\d+)-((?:t\d+:r\d+c\d+:)?p\d+):(\d+)$/;
const BLOCK_RANGE_RE = /^p(\d+)-p(\d+)$/;
const COMMENT_RE = /^c(\d+)$/;
const IMAGE_RE = /^img(\d+)$/;
const LINK_RE = /^link(\d+)$/;
const TRACKED_CHANGE_RE = /^tc(\d+)$/;
const EQUATION_RE = /^eq(\d+)$/;
const FOOTNOTE_RE = /^fn(\d+)$/;
const ENDNOTE_RE = /^en(\d+)$/;
const MARGINAL_RE = /^(hdr|ftr)(\d+)$/;
const CELL_RANGE_RE = /^t(\d+):r(\d+)c(\d+)-r(\d+)c(\d+)$/;
const CELL_RE = /^t(\d+):r(\d+)c(\d+)(?::(.+))?$/;

/** True when a locator carries a cell segment (`…:rRcC`, nested chains
 *  included) — i.e. it addresses a cell or content inside one. Lives beside
 *  `CELL_RE` so it moves with the grammar; the raw splice paths use it to
 *  apply Word's "a cell ends with a paragraph" rule to the spliced parent. */
export function isCellScopedLocator(locator: string): boolean {
	return /:r\d+c\d+(?::|$)/.test(locator);
}

/** True for an `rIdN` relationship handle. The rels family rides `--at`
 *  alongside block locators (an `rId` names the relationship a body reference
 *  points at), so its recognition lives here in the grammar rather than as
 *  scattered inline regexes across the raw CLI verbs. */
export function isRelationshipLocator(locator: string): boolean {
	return /^rId\d+$/.test(locator);
}

/** True for a `hdrN` / `ftrN` marginal handle — the page header/footer id that
 *  `headers list`/`footers list` report, `read` carries on the `docx:header`/
 *  `docx:footer` hint, and `raw get` / `headers`/`footers --at` resolve. Unlike
 *  block locators it addresses content in a SEPARATE part (`word/footer1.xml`),
 *  not the body, so the CLI routes it through the marginal resolver. */
export function isMarginalLocator(locator: string): boolean {
	return MARGINAL_RE.test(locator);
}
const TABLE_ROW_RE = /^t(\d+):r(\d+)$/;
const TABLE_COLUMN_RE = /^t(\d+):c(\d+)$/;

export function parseLocator(input: string): Locator {
	const trimmed = input.trim();
	if (trimmed.length === 0) throw new LocatorParseError(input, "empty");

	const commentMatch = trimmed.match(COMMENT_RE);
	if (commentMatch) {
		return { kind: "comment", commentId: `c${commentMatch[1]}` };
	}

	const imageMatch = trimmed.match(IMAGE_RE);
	if (imageMatch) return { kind: "image", imageId: `img${imageMatch[1]}` };

	const linkMatch = trimmed.match(LINK_RE);
	if (linkMatch) {
		return { kind: "hyperlink", hyperlinkId: `link${linkMatch[1]}` };
	}

	const trackedChangeMatch = trimmed.match(TRACKED_CHANGE_RE);
	if (trackedChangeMatch) {
		return {
			kind: "trackedChange",
			trackedChangeId: `tc${trackedChangeMatch[1]}`,
		};
	}

	const equationMatch = trimmed.match(EQUATION_RE);
	if (equationMatch) {
		return { kind: "equation", equationId: `eq${equationMatch[1]}` };
	}

	const footnoteMatch = trimmed.match(FOOTNOTE_RE);
	if (footnoteMatch) {
		return { kind: "footnote", footnoteId: `fn${footnoteMatch[1]}` };
	}

	const endnoteMatch = trimmed.match(ENDNOTE_RE);
	if (endnoteMatch) {
		return { kind: "endnote", endnoteId: `en${endnoteMatch[1]}` };
	}

	// MARGINAL_RE is anchored ^(hdr|ftr)\d+$, so a match IS the whole id.
	if (MARGINAL_RE.test(trimmed)) {
		return { kind: "marginal", marginalId: trimmed };
	}

	// Checked before the cell forms: a cell-prefixed cross-paragraph range
	// (`t0:r0c0:p0:2-t0:r0c0:p1:3`) would otherwise be swallowed by CELL_RE,
	// whose recursive inner parse can't read the dashed right-hand side.
	const rangeMatch = trimmed.match(RANGE_RE);
	if (rangeMatch) {
		const [, startBlock, startCapture, endBlock, endCapture] = rangeMatch;
		if (
			startBlock &&
			endBlock &&
			cellPrefixOf(startBlock) !== cellPrefixOf(endBlock)
		) {
			throw new LocatorParseError(
				input,
				"range endpoints must share a container — both body paragraphs, or paragraphs of the SAME table cell",
			);
		}
		const startOffset = Number(startCapture);
		const endOffset = Number(endCapture);
		validateOffsets(input, startOffset, endOffset, startBlock !== endBlock);
		return {
			kind: "range",
			start: { blockId: startBlock ?? "", offset: startOffset },
			end: { blockId: endBlock ?? "", offset: endOffset },
		};
	}

	const cellRangeMatch = trimmed.match(CELL_RANGE_RE);
	if (cellRangeMatch) {
		const [, tableIndex, startRow, startCol, endRow, endCol] = cellRangeMatch;
		return {
			kind: "cellRange",
			tableId: `t${tableIndex}`,
			start: { row: Number(startRow), col: Number(startCol) },
			end: { row: Number(endRow), col: Number(endCol) },
		};
	}

	const cellMatch = trimmed.match(CELL_RE);
	if (cellMatch) {
		const [, tableIndex, rowIndex, columnIndex, rest] = cellMatch;
		const result: Locator = {
			kind: "cell",
			tableId: `t${tableIndex}`,
			row: Number(rowIndex),
			col: Number(columnIndex),
		};
		if (rest) result.inner = parseLocator(rest);
		return result;
	}

	const tableRowMatch = trimmed.match(TABLE_ROW_RE);
	if (tableRowMatch) {
		const [, tableIndex, rowIndex] = tableRowMatch;
		return {
			kind: "tableRow",
			tableId: `t${tableIndex}`,
			row: Number(rowIndex),
		};
	}

	const tableColumnMatch = trimmed.match(TABLE_COLUMN_RE);
	if (tableColumnMatch) {
		const [, tableIndex, columnIndex] = tableColumnMatch;
		return {
			kind: "tableColumn",
			tableId: `t${tableIndex}`,
			col: Number(columnIndex),
		};
	}

	const blockRangeMatch = trimmed.match(BLOCK_RANGE_RE);
	if (blockRangeMatch) {
		const [, startIndex, endIndex] = blockRangeMatch;
		const startNum = Number(startIndex);
		const endNum = Number(endIndex);
		if (endNum < startNum) {
			throw new LocatorParseError(input, "end paragraph precedes start");
		}
		return {
			kind: "blockRange",
			startBlockId: `p${startIndex}`,
			endBlockId: `p${endIndex}`,
		};
	}

	const spanMatch = trimmed.match(SPAN_RE);
	if (spanMatch) {
		const [, paragraphIndex, startCapture, endCapture] = spanMatch;
		const start = Number(startCapture);
		const end = Number(endCapture);
		validateOffsets(input, start, end, false);
		return { kind: "blockSpan", blockId: `p${paragraphIndex}`, start, end };
	}

	const blockMatch = trimmed.match(BLOCK_RE);
	if (blockMatch) {
		const [, prefix, idx] = blockMatch;
		return { kind: "block", blockId: `${prefix}${idx}` };
	}

	throw new LocatorParseError(input, "unrecognized syntax");
}

/** The `tN:rRcC` prefix of a cell-paragraph id, `""` for a body paragraph —
 *  the container a cross-paragraph range must not leave. */
function cellPrefixOf(blockId: string): string {
	const lastColon = blockId.lastIndexOf(":");
	return lastColon === -1 ? "" : blockId.slice(0, lastColon);
}

function validateOffsets(
	input: string,
	start: number,
	end: number,
	crossBlock: boolean,
): void {
	if (start < 0 || end < 0) {
		throw new LocatorParseError(input, "offsets must be non-negative");
	}
	if (!crossBlock && end < start) {
		throw new LocatorParseError(input, "end offset precedes start");
	}
}
