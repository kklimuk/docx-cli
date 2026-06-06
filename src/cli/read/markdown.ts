import {
	type Block,
	type Body,
	type Comment,
	type Footnote,
	flattenParagraphs,
	type Locator,
	LocatorParseError,
	type Paragraph,
	parseLocator,
	type Run,
	type Table,
	type TableCell,
	type TableRow,
	type TextRun,
	type TrackedChange,
} from "@core";
import {
	codeBlockLanguageFromStyleId,
	isCodeBlockStyleId,
} from "@core/code-block";
import { extensionForImageMime } from "@core/image/formats";

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
	doc: Body,
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
	let cursor = 0;
	while (cursor < blocks.length) {
		const block = blocks[cursor];
		if (!block) {
			cursor++;
			continue;
		}
		// Collapse a run of CodeBlock paragraphs into one fenced GFM block.
		// Walk forward as long as adjacent blocks are CodeBlock paragraphs;
		// emit one ```...``` for the group. Locator comments live on the fence
		// (start/end) rather than per-line, so the block reads cleanly.
		if (isCodeBlockParagraph(block)) {
			let lookahead = cursor + 1;
			while (lookahead < blocks.length) {
				const next = blocks[lookahead];
				if (!next || !isCodeBlockParagraph(next)) break;
				lookahead++;
			}
			const group = blocks.slice(cursor, lookahead) as Paragraph[];
			parts.push(renderCodeBlockGroup(group, ctx));
			cursor = lookahead;
			continue;
		}
		const rendered = renderBlock(block, ctx);
		if (rendered !== null) parts.push(rendered);
		cursor++;
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
	if (view === "accepted" && (kind === "del" || kind === "moveFrom"))
		return false;
	if (view === "baseline" && (kind === "ins" || kind === "moveTo"))
		return false;
	return true;
}

function buildCommentIndex(
	blocks: Block[],
	options: MarkdownOptions,
): CommentIndex {
	const view = options.view ?? "accepted";
	const lastSlot = new Map<string, string>();
	const spanText = new Map<string, string>();
	const orderedIds: string[] = [];

	for (const paragraph of flattenParagraphs(blocks)) {
		paragraph.runs.forEach((run, index) => {
			// Text runs are the primary carrier; equation runs also need to
			// pick up comments (audit-comment fallback for tracked equation
			// edits anchors comment ranges on the `<m:oMath>` itself).
			const comments = runComments(run);
			if (!comments) return;
			if (run.type === "text" && !isRunVisible(run, view)) return;
			const spanContribution =
				run.type === "text"
					? run.text
					: run.type === "equation"
						? run.text
						: "";
			for (const commentId of comments) {
				if (!spanText.has(commentId)) orderedIds.push(commentId);
				spanText.set(
					commentId,
					(spanText.get(commentId) ?? "") + spanContribution,
				);
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

/** Comment IDs attached to a run, regardless of run type. Today text runs
 *  and equation runs are the only carriers — extend here if other run types
 *  start carrying comment anchors. */
function runComments(run: Run): string[] | undefined {
	if (run.type === "text") return run.comments;
	if (run.type === "equation") return run.comments;
	return undefined;
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

function isCodeBlockParagraph(block: Block): block is Paragraph {
	return block.type === "paragraph" && isCodeBlockStyleId(block.style);
}

/** Collapse a run of CodeBlock paragraphs into a GFM fenced block. Token
 *  formatting (the colors lowlight applied on insert) gets stripped — the
 *  fenced rendering loses syntax-highlighting fidelity but stays a faithful
 *  source code representation, which is the right trade-off for a markdown
 *  view. Locator comments mark the fence start/end rather than each line.
 *  The language tag on the opening fence comes from the first paragraph's
 *  `CodeBlock-LANG` pStyle suffix (or empty for the bare `CodeBlock`).
 *
 *  Tracked-change references inside the group's runs (someone edited a line
 *  under tracking) are collected into `ctx.referencedTrackedChanges` so the
 *  current-view footnote appendix still surfaces them — even though their
 *  CriticMarkup wrappers are stripped from the fenced rendering itself. */
function renderCodeBlockGroup(
	paragraphs: Paragraph[],
	ctx: RenderContext,
): string {
	if (ctx.options.view !== "baseline" && ctx.options.view !== "accepted") {
		// `current` view: collect tracked-change refs so [^tcN] definitions
		// still render in the footnote appendix.
		for (const paragraph of paragraphs) {
			for (const run of paragraph.runs) {
				if (run.type === "text" && run.trackedChange) {
					ctx.referencedTrackedChanges.set(
						run.trackedChange.id,
						run.trackedChange,
					);
				}
			}
		}
	}
	const lines = paragraphs.map((paragraph) =>
		paragraph.runs
			.filter(
				(run): run is TextRun =>
					run.type === "text" && typeof run.text === "string",
			)
			.map((run) => run.text)
			.join(""),
	);
	const firstId = paragraphs[0]?.id ?? "";
	const lastId = paragraphs[paragraphs.length - 1]?.id ?? firstId;
	const language = codeBlockLanguageFromStyleId(paragraphs[0]?.style) ?? "";
	const openComment = `<!-- ${firstId} -->`;
	const closeComment = firstId === lastId ? "" : ` <!-- ${lastId} -->`;
	return [
		`\`\`\`${language}${openComment}`,
		...lines,
		`\`\`\`${closeComment}`,
	].join("\n");
}

function renderParagraph(
	paragraph: Paragraph,
	ctx: RenderContext,
): string | null {
	const body = renderRuns(paragraph.id, paragraph.runs, ctx);
	if (body.length === 0) return null;
	const prefix = paragraphPrefix(paragraph);
	// Display equations need to be on their own line for KaTeX-based renderers
	// (Obsidian, VS Code preview, etc.) to recognize `$$…$$` as display math.
	// Putting the locator after a space on the same line confuses the parser
	// — it sees the trailing `$` as an unmatched math-mode toggle.
	const separator = isDisplayEquationOnly(body) ? "\n" : " ";
	return `${prefix}${body}${separator}<!-- ${paragraph.id} -->`;
}

/** True if the rendered body is a single display-math expression (`$$…$$`)
 *  with no other content — the case where we put the locator on its own
 *  line so KaTeX-based markdown renderers process the math correctly. */
function isDisplayEquationOnly(body: string): boolean {
	const trimmed = body.trim();
	if (!trimmed.startsWith("$$") || !trimmed.endsWith("$$")) return false;
	// Reject `$$X$$Y$$Z$$` (two separate display equations) — only single ones.
	const inner = trimmed.slice(2, -2);
	return !inner.includes("$$");
}

function paragraphPrefix(paragraph: Paragraph): string {
	// Blockquote prefix comes before everything else. `> ` repeats per
	// nesting depth; the AST reader fills `quoteDepth` from `pStyle="Quote"`
	// / `pStyle="QuoteListParagraph"` plus the paragraph's `<w:ind w:left>`
	// value. Markdown stitches adjacent `> ` lines back into one logical
	// blockquote on re-parse.
	const quotePrefix = paragraph.quoteDepth
		? "> ".repeat(paragraph.quoteDepth)
		: "";
	const headingLevel = headingLevelFor(paragraph.style);
	if (headingLevel !== null) {
		return `${quotePrefix}${"#".repeat(headingLevel)} `;
	}
	if (paragraph.list) {
		const indent = "  ".repeat(paragraph.list.level);
		// GFM ordered-list numbering auto-increments client-side; using `1. `
		// for every item is the idiomatic representation.
		const marker = paragraph.list.ordered ? "1. " : "- ";
		const task =
			paragraph.taskState === "checked"
				? "[x] "
				: paragraph.taskState === "unchecked"
					? "[ ] "
					: "";
		return `${quotePrefix}${indent}${marker}${task}`;
	}
	return quotePrefix;
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
	const view = ctx.options.view ?? "accepted";
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
			// Content-addressed URL: `<sha256>.<ext>`. The walker on the
			// import side (`@core/markdown::preloadImages`) recognizes the
			// shape and reuses the existing media part by hash instead of
			// shelling out to `loadImageSource` — that's what makes
			// `read → edit → write` round-trip without re-fetching images.
			// `imageById.values()` in `Body` shares the same hash → rId
			// mapping the walker queries. Mirrors the on-disk naming
			// convention used by `docx images extract`.
			const extension = extensionForImageMime(run.contentType) ?? "bin";
			out += `![${alt}](${run.hash}.${extension})`;
		} else if (run.type === "break") {
			if (run.kind === "line") out += "<br>";
		} else if (run.type === "tab") {
			out += "\t";
		} else if (run.type === "equation") {
			// `$…$` for inline, `$$…$$` for display. The walker in
			// `@core/equation` reconstructed `run.latex` from the OMML subtree;
			// `run.text` is the legacy plaintext fallback for fully-degraded
			// (unrecognized) equations — use it only when the LaTeX walker
			// returned empty.
			const body = run.latex.length > 0 ? run.latex : run.text;
			out += run.display ? `$$${body}$$` : `$${body}$`;
			// Append any comment endings that close on this equation run
			// (audit comments from tracked equation edits anchor here).
			out += commentEndingsFor(
				paragraphId,
				[{ originalIndex: cursor }],
				ctx.commentIndex,
			);
		} else if (run.type === "noteRef") {
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
		(a.runStyle ?? "") === (b.runStyle ?? "") &&
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
	// Backticks INSIDE other formatting per GFM precedence — `**`x`**` is bold
	// code; `**x**` inside backticks would be literal asterisks. Skip Code runs
	// inside a fenced code block (callers strip runStyle on those — see
	// `renderCodeBlockGroup`); this branch handles `runStyle: "Code"` only when
	// it's still meaningful (inline code spans in a normal paragraph).
	if (first.runStyle === "Code") out = `\`${out}\``;
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
		const marker = criticMarkerFor(first.trackedChange.kind);
		out = `{${marker}${out}${marker}}[^${first.trackedChange.id}]`;
	}
	return out;
}

/** CriticMarkup doesn't have a native "moved" marker, so we render moveTo
 * with the same `++` markers as an insertion (the text appears at this
 * location in the accepted view) and moveFrom with the same `--` as a
 * deletion (the text leaves this location). The footnote definition
 * carries the precise kind so a reader can distinguish move vs. ins/del. */
function criticMarkerFor(kind: TrackedChange["kind"]): "++" | "--" {
	if (kind === "ins" || kind === "moveTo") return "++";
	return "--";
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
	// Whole-row tracked changes are filtered by view: accepted drops deleted
	// rows, baseline drops inserted ones. Cell-level (column) tracked changes
	// can't be represented in a GFM table, so they're left in place.
	const view = ctx.options.view ?? "accepted";
	const rows = table.rows.filter((row) => isRowVisible(row, view));
	if (rows.length === 0) return null;
	const colCount = Math.max(...rows.map((row) => row.cells.length));
	if (colCount === 0) return null;
	const renderedRows = rows.map((row) => {
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

function isRowVisible(row: TableRow, view: MarkdownView): boolean {
	const kind = row.trackedChange?.kind;
	if (view === "accepted" && kind === "rowDel") return false;
	if (view === "baseline" && kind === "rowIns") return false;
	return true;
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
		const kind = trackedChangeLabelFor(change.kind);
		const author = change.author || "unknown";
		const meta = change.date ? `${author} (${change.date})` : author;
		lines.push(`[^${change.id}]: ${kind} by ${meta}`);
	}
	return lines.join("\n");
}

function trackedChangeLabelFor(kind: TrackedChange["kind"]): string {
	if (kind === "ins") return "insertion";
	if (kind === "del") return "deletion";
	if (kind === "moveTo") return "moveTo";
	return "moveFrom";
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
		case "blockRange":
			return position === "from" ? parsed.startBlockId : parsed.endBlockId;
		case "cell":
			return parsed.tableId;
		case "comment":
		case "image":
		case "hyperlink":
		case "trackedChange":
		case "equation":
		case "footnote":
		case "endnote":
		case "tableRow":
		case "tableColumn":
		case "cellRange":
			throw new MarkdownLocatorError(
				input,
				`--${position} does not accept a ${parsed.kind} locator — use a paragraph or table locator`,
			);
	}
}
