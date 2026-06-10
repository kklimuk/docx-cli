import type {
	Blockquote,
	Heading,
	List,
	ListItem,
	Paragraph as MdParagraph,
	Table as MdTable,
	TableRow as MdTableRow,
	PhrasingContent,
	Root,
	RootContent,
} from "mdast";
import type { Math as MdMath } from "mdast-util-math";
import { HorizontalRule } from "../blocks";
import { buildCodeBlockParagraphs, ensureCodeBlockStyles } from "../code-block";
import { latexToOmml } from "../equation";
import { w } from "../jsx";
import type { NullableXmlNode, XmlNode } from "../parser";
import { Table, TableCell, TableRow } from "../table";
import { TaskCheckbox } from "../task-list";
import { MarkdownImportError } from "./errors";
import { type WalkContext, walkInline } from "./inline";

/** Drive every block in the mdast root into OOXML, returning a flat
 * `XmlNode[]` ready to splice into `<w:body>` (or a parent `<w:tc>`). The
 * lens does pre-walk provisioning (image fetch, footnote-id minting) before
 * calling here so this stays synchronous. */
export function walkRoot(root: Root, ctx: WalkContext): XmlNode[] {
	return walkBlocks(root.children, ctx);
}

function walkBlocks(
	nodes: readonly RootContent[],
	ctx: WalkContext,
): XmlNode[] {
	const out: XmlNode[] = [];
	for (const node of nodes) {
		for (const block of walkBlock(node, ctx)) out.push(block);
	}
	return out;
}

function walkBlock(node: RootContent, ctx: WalkContext): XmlNode[] {
	switch (node.type) {
		case "paragraph":
			return [paragraphBlock(node, ctx, undefined)];
		case "heading":
			return [headingBlock(node, ctx)];
		case "list":
			return listBlocks(node, ctx, 0, null);
		case "blockquote":
			return blockquoteBlocks(node, ctx);
		case "code":
			ensureCodeBlockStyles(ctx.document, node.lang ?? undefined);
			return buildCodeBlockParagraphs(node.value, node.lang ?? undefined);
		case "thematicBreak":
			return [<HorizontalRule />];
		case "table":
			return [tableBlock(node, ctx)];
		case "math":
			return [mathParagraph(node)];
		case "html":
			// Raw block HTML (stray tags, `<!-- pN -->` locators, and the
			// `docx:section` / `docx:page` / `docx:table` visibility annotations
			// `read` emits) — drop. These are read-time hints, not parse-back: the
			// importer never reconstructs structure from them (`--ast` is the
			// lossless view; edit-in-place + `docx sections`
			// manage layout). Agents author through markdown features.
			return [];
		case "footnoteDefinition":
			// The lens pre-collects footnote definitions and registers their
			// bodies into `footnotes.xml` before the block walk runs. Nothing
			// emits inline here.
			return [];
		case "definition":
			// Link/image reference defs — remark inlined them; we never see
			// them here as orphans.
			return [];
		case "yaml":
			// YAML frontmatter — drop. (Could surface into `core.xml` properties
			// in a follow-up, but that's out of S8 scope.)
			return [];
		default:
			throw new MarkdownImportError(
				"USAGE",
				`Unsupported markdown block: ${(node as { type: string }).type}`,
			);
	}
}

function paragraphBlock(
	node: MdParagraph,
	ctx: WalkContext,
	options: { style?: string; leftIndentTwips?: number } = {},
): XmlNode {
	const children = walkInline(node.children, ctx);
	return (
		<w.p>
			<ParagraphProperties
				style={options.style}
				leftIndentTwips={options.leftIndentTwips}
			/>
			{children}
		</w.p>
	);
}

function headingBlock(node: Heading, ctx: WalkContext): XmlNode {
	const styleId = `Heading${node.depth}` as const;
	ctx.document.ensureStyles().ensureStyle(styleId);
	const children = walkInline(node.children, ctx);
	return (
		<w.p>
			<w.pPr>
				<w.pStyle w-val={styleId} />
			</w.pPr>
			{children}
		</w.p>
	);
}

/** Walk a markdown blockquote into a flat sequence of `<w:p>` quoted-paragraph
 * emissions. OOXML has no `<w:blockquote>` container — quote treatment is
 * per-paragraph via `pStyle="Quote"` (or `QuoteListParagraph` for nested
 * lists) plus a `<w:ind w:left>` whose value encodes nesting depth at
 * `720 * depth` twips. The AST reader recovers `paragraph.quoteDepth` from
 * those two signals; the markdown renderer prepends one `> ` per depth.
 *
 * **Supported nested content**: paragraphs, lists (bullet / ordered / task),
 * and nested blockquotes. These all carry the quote framing on round-trip.
 *
 * **Unsupported nested content**: code blocks, tables, math, headings, and
 * thematic breaks inside a blockquote intentionally **escape** — we emit
 * them at top level with no quote framing, breaking the quote at that point.
 * On round-trip they come back outside the quote. This is documented per
 * the v0.12 design (see [src/core/markdown/CLAUDE.md](./CLAUDE.md)). */
function blockquoteBlocks(
	node: Blockquote,
	ctx: WalkContext,
	depth = 1,
): XmlNode[] {
	ctx.document.ensureStyles().ensureStyle("Quote");
	const out: XmlNode[] = [];
	for (const child of node.children) {
		if (child.type === "paragraph") {
			out.push(
				paragraphBlock(child, ctx, {
					style: "Quote",
					leftIndentTwips: 720 * depth,
				}),
			);
			continue;
		}
		if (child.type === "blockquote") {
			for (const block of blockquoteBlocks(child, ctx, depth + 1)) {
				out.push(block);
			}
			continue;
		}
		if (child.type === "list") {
			for (const block of listBlocks(child, ctx, 0, null, depth)) {
				out.push(block);
			}
			continue;
		}
		// Escape: code / table / math / heading / HR / definition / html /
		// footnoteDefinition / yaml inside a blockquote emit at top level
		// without the quote framing. They break the surrounding quote at
		// that point — adjacent quoted content before and after surfaces as
		// separate blockquotes on round-trip.
		for (const block of walkBlock(child, ctx)) out.push(block);
	}
	return out;
}

function mathParagraph(node: MdMath): XmlNode {
	try {
		// `latexToOmml(_, true)` returns `<m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>`,
		// which sits as a direct `<w:p>` child (the same shape `insert --equation
		// --display` emits via `EquationParagraph` in `core/insert/index.tsx`).
		const omml = latexToOmml(node.value, true);
		return <w.p>{omml}</w.p>;
	} catch (error) {
		throw new MarkdownImportError(
			"USAGE",
			`Could not parse display math: ${error instanceof Error ? error.message : String(error)}`,
			"Check the LaTeX. We accept the temml dialect (KaTeX-compatible).",
		);
	}
}

/** Walk an mdast `list` into a flat sequence of `<w:p>` list-paragraphs.
 *
 * `parentNumId === null` is the outermost call: we allocate a fresh numId via
 * `ensureNumbering().allocate(kind)` and provision the `ListParagraph` style.
 * Nested lists inherit `numId` from their parent so the same list-id cascades
 * through indentation levels, and `level` increments. */
function listBlocks(
	node: List,
	ctx: WalkContext,
	level: number,
	parentNumId: number | null,
	quoteDepth = 0,
): XmlNode[] {
	const kind = node.ordered ? "ordered" : "bullet";
	const numbering = ctx.document.ensureNumbering();
	// mdast's `start` is only meaningful on ordered lists. Forward it as the
	// startOverride so `10. tenth` survives intact; bullet lists fall through
	// to the default (1), which still emits a `<w:lvlOverride>` to prevent
	// Word from auto-continuing the previous list.
	const start = node.ordered ? (node.start ?? 1) : 1;
	const numId = parentNumId ?? numbering.allocate(kind, start);
	ctx.document.ensureStyles().ensureStyle("ListParagraph");
	if (quoteDepth > 0) {
		ctx.document.ensureStyles().ensureStyle("QuoteListParagraph");
	}
	const out: XmlNode[] = [];
	for (const item of node.children) {
		for (const block of listItemBlocks(item, ctx, level, numId, quoteDepth))
			out.push(block);
	}
	return out;
}

function listItemBlocks(
	item: ListItem,
	ctx: WalkContext,
	level: number,
	numId: number,
	quoteDepth: number,
): XmlNode[] {
	const out: XmlNode[] = [];
	let firstParagraphSeen = false;
	for (const child of item.children) {
		if (child.type === "paragraph") {
			const taskState =
				!firstParagraphSeen && item.checked != null
					? item.checked
						? ("checked" as const)
						: ("unchecked" as const)
					: undefined;
			out.push(
				listParagraphBlock(child, ctx, {
					numId,
					level,
					taskState,
					quoteDepth,
				}),
			);
			firstParagraphSeen = true;
			continue;
		}
		if (child.type === "list") {
			for (const block of listBlocks(
				child,
				ctx,
				level + 1,
				numId,
				quoteDepth,
			)) {
				out.push(block);
			}
			continue;
		}
		// Other block content (code, blockquote) inside a listItem — emit
		// without list-paragraph framing so it visually associates with the
		// item but isn't shoehorned into list semantics.
		for (const block of walkBlock(child, ctx)) out.push(block);
	}
	return out;
}

function listParagraphBlock(
	node: MdParagraph,
	ctx: WalkContext,
	options: {
		numId: number;
		level: number;
		taskState: "checked" | "unchecked" | undefined;
		quoteDepth: number;
	},
): XmlNode {
	const children = walkInline(node.children, ctx);
	const pStyle =
		options.quoteDepth > 0 ? "QuoteListParagraph" : "ListParagraph";
	// Word silently ignores a paragraph-level `<w:ind w:left>` on a numbered
	// paragraph when `w:hanging` isn't also specified — it falls back to the
	// numbering lvl's pPr (`left=300, hanging=240` at ilvl=0). To shift the
	// whole bullet+body pair right by the quote indent we emit BOTH values
	// explicitly: body at `300*(level+1) + 720*depth`, bullet at body minus
	// 240 (matching the lvl's hanging). These constants track the indents in
	// `AbstractNum` in core/ast/document/numbering.tsx; bump them together if
	// you change one.
	const indent =
		options.quoteDepth > 0
			? {
					left: 300 * (options.level + 1) + 720 * options.quoteDepth,
					hanging: 240,
				}
			: undefined;
	return (
		<w.p>
			<w.pPr>
				<w.pStyle w-val={pStyle} />
				<w.numPr>
					<w.ilvl w-val={String(options.level)} />
					<w.numId w-val={String(options.numId)} />
				</w.numPr>
				{indent && (
					<w.ind
						w-left={String(indent.left)}
						w-hanging={String(indent.hanging)}
					/>
				)}
			</w.pPr>
			{options.taskState !== undefined && (
				<TaskCheckbox checked={options.taskState === "checked"} />
			)}
			{options.taskState !== undefined && (
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}> </w.t>
				</w.r>
			)}
			{children}
		</w.p>
	);
}

/** GFM tables map to `<Table>` with even column widths. We don't bold the
 * header row — markdown's delimiter row carries alignment hints but no "this
 * is a header" semantic that translates to Word's `tblLook firstRow`. Leave
 * styling to a follow-up if asked. */
function tableBlock(node: MdTable, ctx: WalkContext): XmlNode {
	const colCount = Math.max(...node.children.map((row) => row.children.length));
	if (colCount === 0) {
		throw new MarkdownImportError("USAGE", "Markdown table has zero columns");
	}
	const grid = evenColumnWidths(colCount);
	const rows = node.children.map((row) => tableRowNode(row, ctx, colCount));
	return <Table grid={grid}>{rows}</Table>;
}

function tableRowNode(
	row: MdTableRow,
	ctx: WalkContext,
	colCount: number,
): XmlNode {
	const cells: XmlNode[] = [];
	for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
		const cell = row.children[columnIndex];
		const cellChildren: PhrasingContent[] = cell ? cell.children : [];
		cells.push(
			<TableCell>
				<w.p>{walkInline(cellChildren, ctx)}</w.p>
			</TableCell>,
		);
	}
	return <TableRow>{cells}</TableRow>;
}

function ParagraphProperties({
	style,
	leftIndentTwips,
}: {
	style: string | undefined;
	leftIndentTwips: number | undefined;
}): NullableXmlNode {
	if (!style && leftIndentTwips === undefined) return null;
	// Schema child order (CT_PPrBase, §17.3.1.26): pStyle → … → ind → jc. We
	// only emit pStyle + ind here so order is fixed by construction.
	return (
		<w.pPr>
			{style && <w.pStyle w-val={style} />}
			{leftIndentTwips !== undefined && (
				<w.ind w-left={String(leftIndentTwips)} />
			)}
		</w.pPr>
	);
}

function evenColumnWidths(cols: number): number[] {
	// Mirror `evenGrid` in `core/table/index.tsx` — 9360 twips total = 6.5"
	// at letter paper with default 1" margins. Last column absorbs the
	// rounding remainder so the row sums to the page width exactly.
	const total = 9360;
	const base = Math.floor(total / cols);
	const remainder = total - base * cols;
	return Array.from({ length: cols }, (_, index) =>
		index === cols - 1 ? base + remainder : base,
	);
}
