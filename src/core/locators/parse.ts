export type Locator =
	| { kind: "block"; blockId: string }
	| { kind: "blockSpan"; blockId: string; start: number; end: number }
	| {
			kind: "range";
			start: { blockId: string; offset: number };
			end: { blockId: string; offset: number };
	  }
	| { kind: "comment"; commentId: string }
	| { kind: "image"; imageId: string }
	| { kind: "hyperlink"; hyperlinkId: string }
	| { kind: "trackedChange"; trackedChangeId: string }
	| {
			kind: "cell";
			tableId: string;
			row: number;
			col: number;
			inner?: Locator;
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
const RANGE_RE = /^p(\d+):(\d+)-p(\d+):(\d+)$/;
const COMMENT_RE = /^c(\d+)$/;
const IMAGE_RE = /^img(\d+)$/;
const LINK_RE = /^link(\d+)$/;
const TRACKED_CHANGE_RE = /^tc(\d+)$/;
const CELL_RE = /^t(\d+):r(\d+)c(\d+)(?::(.+))?$/;

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

	const rangeMatch = trimmed.match(RANGE_RE);
	if (rangeMatch) {
		const [, startBlock, startCapture, endBlock, endCapture] = rangeMatch;
		const startOffset = Number(startCapture);
		const endOffset = Number(endCapture);
		validateOffsets(input, startOffset, endOffset, startBlock !== endBlock);
		return {
			kind: "range",
			start: { blockId: `p${startBlock}`, offset: startOffset },
			end: { blockId: `p${endBlock}`, offset: endOffset },
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
