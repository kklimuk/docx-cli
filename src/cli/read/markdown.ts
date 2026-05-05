import {
	type Block,
	type Comment,
	type Doc,
	type Footnote,
	flattenParagraphs,
	type Locator,
	LocatorParseError,
	type Paragraph,
	parseLocator,
	type Run,
	type Table,
	type TableCell,
	type TextRun,
	type TrackedChange,
} from "@core";

export type MarkdownView = "current" | "accepted" | "baseline";

export type MarkdownOptions = {
	from?: string;
	to?: string;
	view?: MarkdownView;
	showComments?: boolean;
};

export class MarkdownLocatorError extends Error {
	constructor(
		public input: string,
		message: string,
	) {
		super(message);
		this.name = "MarkdownLocatorError";
	}
}

type CommentIndex = {
	endingsByRun: Map<string, string[]>;
	spanText: Map<string, string>;
	orderedIds: string[];
};

type RenderContext = {
	options: MarkdownOptions;
	commentIndex: CommentIndex;
	referencedFootnoteIds: Set<string>;
	referencedEndnoteIds: Set<string>;
	referencedTrackedChanges: Map<string, TrackedChange>;
};

export function renderMarkdown(
	doc: Doc,
	options: MarkdownOptions = {},
): string {
	const blocks = sliceBlocks(doc.blocks, options.from, options.to);
	const commentIndex = options.showComments
		? buildCommentIndex(blocks, options)
		: emptyCommentIndex();
	const ctx: RenderContext = {
		options,
		commentIndex,
		referencedFootnoteIds: new Set(),
		referencedEndnoteIds: new Set(),
		referencedTrackedChanges: new Map(),
	};

	const parts: string[] = [];
	for (const block of blocks) {
		const rendered = renderBlock(block, ctx);
		if (rendered !== null) parts.push(rendered);
	}
	const definitions: string[] = [];
	if (options.showComments) {
		const commentFootnotes = renderCommentFootnotes(commentIndex, doc.comments);
		if (commentFootnotes.length > 0) definitions.push(commentFootnotes);
	}
	const footnoteDefs = renderNoteDefinitions(
		doc.footnotes,
		ctx.referencedFootnoteIds,
	);
	if (footnoteDefs.length > 0) definitions.push(footnoteDefs);
	const endnoteDefs = renderNoteDefinitions(
		doc.endnotes,
		ctx.referencedEndnoteIds,
	);
	if (endnoteDefs.length > 0) definitions.push(endnoteDefs);
	const trackedChangeDefs = renderTrackedChangeFootnotes(
		ctx.referencedTrackedChanges,
	);
	if (trackedChangeDefs.length > 0) definitions.push(trackedChangeDefs);
	if (definitions.length > 0) parts.push(definitions.join("\n"));
	if (parts.length === 0) return "";
	return `${parts.join("\n\n")}\n`;
}

function emptyCommentIndex(): CommentIndex {
	return {
		endingsByRun: new Map(),
		spanText: new Map(),
		orderedIds: [],
	};
}

function isRunVisible(run: TextRun, view: MarkdownView): boolean {
	const kind = run.trackedChange?.kind;
	if (!kind) return true;
	if (view === "accepted" && kind === "del") return false;
	if (view === "baseline" && kind === "ins") return false;
	return true;
}

function buildCommentIndex(
	blocks: Block[],
	options: MarkdownOptions,
): CommentIndex {
	const view = options.view ?? "current";
	const lastSlot = new Map<string, string>();
	const spanText = new Map<string, string>();
	const orderedIds: string[] = [];

	for (const paragraph of flattenParagraphs(blocks)) {
		paragraph.runs.forEach((run, index) => {
			if (run.type !== "text") return;
			if (!isRunVisible(run, view)) return;
			for (const commentId of run.comments ?? []) {
				if (!spanText.has(commentId)) orderedIds.push(commentId);
				spanText.set(commentId, (spanText.get(commentId) ?? "") + run.text);
				lastSlot.set(commentId, slotKey(paragraph.id, index));
			}
		});
	}

	const endingsByRun = new Map<string, string[]>();
	for (const commentId of orderedIds) {
		const slot = lastSlot.get(commentId);
		if (!slot) continue;
		const list = endingsByRun.get(slot) ?? [];
		list.push(commentId);
		endingsByRun.set(slot, list);
	}

	return { endingsByRun, spanText, orderedIds };
}

function slotKey(paragraphId: string, runIndex: number): string {
	return `${paragraphId}#${runIndex}`;
}

function renderBlock(block: Block, ctx: RenderContext): string | null {
	if (block.type === "paragraph") return renderParagraph(block, ctx);
	if (block.type === "table") return renderTable(block, ctx);
	if (block.type === "sectionBreak") return `--- <!-- ${block.id} -->`;
	return null;
}

function renderParagraph(
	paragraph: Paragraph,
	ctx: RenderContext,
): string | null {
	const body = renderRuns(paragraph.id, paragraph.runs, ctx);
	if (body.length === 0) return null;
	const prefix = paragraphPrefix(paragraph);
	return `${prefix}${body} <!-- ${paragraph.id} -->`;
}

function paragraphPrefix(paragraph: Paragraph): string {
	const headingLevel = headingLevelFor(paragraph.style);
	if (headingLevel !== null) return `${"#".repeat(headingLevel)} `;
	if (paragraph.list) {
		const indent = "  ".repeat(paragraph.list.level);
		return `${indent}- `;
	}
	return "";
}

function headingLevelFor(style: string | undefined): number | null {
	if (!style) return null;
	if (!style.startsWith("Heading")) return null;
	const remainder = style.slice("Heading".length).trim();
	if (remainder === "") return 1;
	const parsed = Number(remainder);
	if (!Number.isInteger(parsed) || parsed < 1 || parsed > 6) return null;
	return parsed;
}

function renderRuns(
	paragraphId: string,
	runs: Run[],
	ctx: RenderContext,
): string {
	const view = ctx.options.view ?? "current";
	const visibleEntries: { run: Run; originalIndex: number }[] = [];
	runs.forEach((run, index) => {
		if (run.type === "text" && !isRunVisible(run, view)) return;
		visibleEntries.push({ run, originalIndex: index });
	});

	let out = "";
	let cursor = 0;
	while (cursor < visibleEntries.length) {
		const entry = visibleEntries[cursor];
		if (!entry) {
			cursor++;
			continue;
		}
		const { run } = entry;
		if (run.type === "text") {
			let lookahead = cursor + 1;
			while (lookahead < visibleEntries.length) {
				const next = visibleEntries[lookahead];
				if (!next || next.run.type !== "text") break;
				if (!sameDecoration(run, next.run)) break;
				lookahead++;
			}
			const segment = visibleEntries.slice(cursor, lookahead);
			const segmentRuns = segment.map((entry) => entry.run as TextRun);
			out += renderTextSegment(segmentRuns, view);
			if (view === "current") {
				for (const segmentRun of segmentRuns) {
					if (segmentRun.trackedChange) {
						ctx.referencedTrackedChanges.set(
							segmentRun.trackedChange.id,
							segmentRun.trackedChange,
						);
					}
				}
			}
			out += commentEndingsFor(paragraphId, segment, ctx.commentIndex);
			cursor = lookahead;
			continue;
		}
		if (run.type === "image") {
			const alt = sanitizeAltText(run.alt ?? run.id);
			out += `![${alt}](${run.id})`;
		} else if (run.type === "break") {
			if (run.kind === "line") out += "<br>";
		} else if (run.type === "tab") {
			out += "\t";
		} else if (run.type === "equation") {
			const escaped = run.text.replace(/`/g, "\\`");
			out += `\`equation: ${escaped}\``;
		} else if (run.type === "footnoteRef") {
			if (run.kind === "footnote") ctx.referencedFootnoteIds.add(run.id);
			else ctx.referencedEndnoteIds.add(run.id);
			out += `[^${run.id}]`;
		} else if (run.type === "chart") {
			out += `\`[${run.kind}]\``;
		}
		cursor++;
	}
	return out;
}

function commentEndingsFor(
	paragraphId: string,
	segment: { originalIndex: number }[],
	commentIndex: CommentIndex,
): string {
	if (commentIndex.endingsByRun.size === 0) return "";
	let out = "";
	for (const entry of segment) {
		const ids = commentIndex.endingsByRun.get(
			slotKey(paragraphId, entry.originalIndex),
		);
		if (!ids) continue;
		for (const commentId of ids) out += `[^${commentId}]`;
	}
	return out;
}

function sameDecoration(a: TextRun, b: TextRun): boolean {
	return (
		(a.bold ?? false) === (b.bold ?? false) &&
		(a.italic ?? false) === (b.italic ?? false) &&
		(a.strike ?? false) === (b.strike ?? false) &&
		(a.underline ?? "") === (b.underline ?? "") &&
		(a.color ?? "") === (b.color ?? "") &&
		(a.highlight ?? "") === (b.highlight ?? "") &&
		a.hyperlink?.id === b.hyperlink?.id &&
		a.trackedChange?.id === b.trackedChange?.id &&
		sameCommentSet(a.comments, b.comments)
	);
}

function sameCommentSet(
	left: string[] | undefined,
	right: string[] | undefined,
): boolean {
	const a = left ?? [];
	const b = right ?? [];
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) {
		if (a[i] !== b[i]) return false;
	}
	return true;
}

function renderTextSegment(runs: TextRun[], view: MarkdownView): string {
	const text = runs.map((run) => run.text).join("");
	if (text.length === 0) return "";
	const first = runs[0];
	if (!first) return "";
	let out = text;
	if (first.bold) out = `**${out}**`;
	if (first.italic) out = `*${out}*`;
	if (first.strike) out = `~~${out}~~`;
	if (first.underline) out = `<u>${out}</u>`;
	const color = colorAttrFor(first.color);
	if (color) out = `<span style="color:${color}">${out}</span>`;
	const highlight = highlightCssFor(first.highlight);
	if (highlight)
		out = `<span style="background-color:${highlight}">${out}</span>`;
	if (first.hyperlink) {
		const target = first.hyperlink.url ?? `#${first.hyperlink.anchor ?? ""}`;
		out = `[${out}](${target})`;
	}
	if (view === "current" && first.trackedChange) {
		const marker = first.trackedChange.kind === "ins" ? "++" : "--";
		out = `{${marker}${out}${marker}}[^${first.trackedChange.id}]`;
	}
	return out;
}

function colorAttrFor(value: string | undefined): string | null {
	if (!value) return null;
	const lowered = value.toLowerCase();
	if (lowered === "auto") return null;
	if (/^[0-9a-f]{6}$/.test(lowered)) return `#${lowered}`;
	return value;
}

/** Map an OOXML highlight value (`yellow`/`darkBlue`/...) to a CSS color. The
 * names match CSS color keywords once camelCase is folded to lowercase. */
function highlightCssFor(value: string | undefined): string | null {
	if (!value) return null;
	if (value === "none") return null;
	return value.replace(/[A-Z]/g, (letter) => letter.toLowerCase());
}

function renderTable(table: Table, ctx: RenderContext): string | null {
	if (table.rows.length === 0) return null;
	const colCount = Math.max(...table.rows.map((row) => row.cells.length));
	if (colCount === 0) return null;
	const renderedRows = table.rows.map((row) => {
		const cells: string[] = [];
		for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
			const cell = row.cells[columnIndex];
			cells.push(cell ? renderCell(cell, ctx) : "");
		}
		return cells;
	});
	const lines: string[] = [];
	const headerRow = renderedRows[0];
	if (!headerRow) return null;
	lines.push(rowToLine(headerRow));
	lines.push(`| ${Array(colCount).fill("---").join(" | ")} |`);
	for (let rowIndex = 1; rowIndex < renderedRows.length; rowIndex++) {
		const row = renderedRows[rowIndex];
		if (row) lines.push(rowToLine(row));
	}
	return lines.join("\n");
}

function rowToLine(cells: string[]): string {
	return `| ${cells.join(" | ")} |`;
}

function renderCell(cell: TableCell, ctx: RenderContext): string {
	const parts: string[] = [];
	for (const block of cell.blocks) {
		if (block.type === "paragraph") {
			const body = renderRuns(block.id, block.runs, ctx);
			if (body.length === 0) continue;
			parts.push(`${body} <!-- ${block.id} -->`);
			continue;
		}
		if (block.type === "table") {
			parts.push(renderNestedTable(block, ctx));
		}
	}
	return escapeCell(parts.join("<br>"));
}

function renderNestedTable(table: Table, ctx: RenderContext): string {
	const rows = table.rows.map((row) =>
		row.cells.map((cell) => renderCell(cell, ctx)).join(" / "),
	);
	return rows.join(" // ");
}

function escapeCell(text: string): string {
	return text.replace(/\|/g, "\\|");
}

function sanitizeAltText(text: string): string {
	return text.replace(/[\r\n]+/g, " ").replace(/\]/g, "\\]");
}

function renderCommentFootnotes(
	commentIndex: CommentIndex,
	comments: Comment[],
): string {
	if (commentIndex.orderedIds.length === 0) return "";
	const byId = new Map(comments.map((comment) => [comment.id, comment]));
	const sorted = [...commentIndex.orderedIds].sort(commentIdCompare);
	const lines: string[] = [];
	for (const commentId of sorted) {
		const comment = byId.get(commentId);
		if (!comment) continue;
		const span = commentIndex.spanText.get(commentId) ?? "";
		lines.push(formatFootnote(comment, span));
	}
	return lines.join("\n");
}

function renderNoteDefinitions(
	notes: Footnote[],
	referenced: Set<string>,
): string {
	if (referenced.size === 0) return "";
	const byId = new Map(notes.map((note) => [note.id, note]));
	const sorted = [...referenced].sort(noteIdCompare);
	const lines: string[] = [];
	for (const id of sorted) {
		const note = byId.get(id);
		const body = (note?.text ?? "").replace(/\s+/g, " ").trim();
		lines.push(`[^${id}]: ${body}`);
	}
	return lines.join("\n");
}

function renderTrackedChangeFootnotes(
	referenced: Map<string, TrackedChange>,
): string {
	if (referenced.size === 0) return "";
	const sorted = [...referenced.values()].sort((a, b) =>
		trackedChangeIdCompare(a.id, b.id),
	);
	const lines: string[] = [];
	for (const change of sorted) {
		const kind = change.kind === "ins" ? "insertion" : "deletion";
		const author = change.author || "unknown";
		const meta = change.date ? `${author} (${change.date})` : author;
		lines.push(`[^${change.id}]: ${kind} by ${meta}`);
	}
	return lines.join("\n");
}

function trackedChangeIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /^tc(\d+)$/);
}

function noteIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /(\d+)$/);
}

function commentIdCompare(left: string, right: string): number {
	return numericIdCompare(left, right, /^c(\d+)$/);
}

function numericIdCompare(
	left: string,
	right: string,
	pattern: RegExp,
): number {
	const leftMatch = left.match(pattern);
	const rightMatch = right.match(pattern);
	if (leftMatch?.[1] && rightMatch?.[1]) {
		return Number(leftMatch[1]) - Number(rightMatch[1]);
	}
	return left.localeCompare(right);
}

function formatFootnote(comment: Comment, spanText: string): string {
	const quoted = quoteSpan(spanText);
	const reply = comment.parentId ? ` ↳ ${comment.parentId}` : "";
	const resolved = comment.resolved ? "✓ " : "";
	const body = comment.text.replace(/\s+/g, " ").trim();
	return `[^${comment.id}]: ${quoted} — ${resolved}${comment.author} (${comment.date})${reply}: ${body}`;
}

function quoteSpan(text: string): string {
	const collapsed = text.replace(/\s+/g, " ").trim();
	const escaped = collapsed.replace(/"/g, '\\"');
	return `"${escaped}"`;
}

function sliceBlocks(
	blocks: Block[],
	from: string | undefined,
	to: string | undefined,
): Block[] {
	if (!from && !to) return blocks;
	const fromId = from ? blockIdForLocator(from, "from") : null;
	const toId = to ? blockIdForLocator(to, "to") : null;
	const fromIndex = fromId ? blocks.findIndex((b) => b.id === fromId) : 0;
	if (from && fromId && fromIndex === -1) {
		throw new MarkdownLocatorError(
			from,
			`--from ${from} not found at document top level`,
		);
	}
	const toIndex = toId
		? blocks.findIndex((b) => b.id === toId)
		: blocks.length - 1;
	if (to && toId && toIndex === -1) {
		throw new MarkdownLocatorError(
			to,
			`--to ${to} not found at document top level`,
		);
	}
	if (toIndex < fromIndex) return [];
	return blocks.slice(fromIndex, toIndex + 1);
}

function blockIdForLocator(input: string, position: "from" | "to"): string {
	let parsed: Locator;
	try {
		parsed = parseLocator(input);
	} catch (err) {
		if (err instanceof LocatorParseError) {
			throw new MarkdownLocatorError(input, err.message);
		}
		throw err;
	}
	switch (parsed.kind) {
		case "block":
			return parsed.blockId;
		case "blockSpan":
			return parsed.blockId;
		case "range":
			return position === "from" ? parsed.start.blockId : parsed.end.blockId;
		case "cell":
			return parsed.tableId;
		case "comment":
		case "image":
		case "hyperlink":
		case "trackedChange":
			throw new MarkdownLocatorError(
				input,
				`--${position} does not accept a ${parsed.kind} locator — use a paragraph or table locator`,
			);
	}
}
