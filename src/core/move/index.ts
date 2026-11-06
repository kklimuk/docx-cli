import type { Document } from "../ast/document";
import { LocatorResolveError } from "../ast/document/body";
import { CellTargetError } from "../locators/cell-target-error";
import { LocatorParseError } from "../locators/parse";
import type { XmlNode } from "../parser";
import { CellInsertionCursor, ensureCellEndsWithParagraph } from "../table";
import { type MoveDestination, resolveMoveDestination } from "./destination";
import {
	assertDestinationOutsideSource,
	assertMovableSource,
	assertNotNoOp,
	assertSameStory,
	assertSourcesDisjoint,
	destinationInside,
} from "./guards";
import { MoveError } from "./move-error";
import { type MoveSource, resolveMoveSource } from "./source";

export { MoveError };

/** Cross-cutting lens over "relocate existing blocks." Stateless — a
 *  back-reference to the document only. The move is a pure splice of the LIVE
 *  source nodes, so everything inside them (comment ranges, note references,
 *  bookmarks, fields, content controls, drawings, text boxes, `dcx:raw`
 *  blocks) rides along untouched — nothing round-trips through Markdown.
 *
 *  Two phases so the CLI can preview: {@link plan} resolves and validates every
 *  request against the document AS READ (it mutates nothing), {@link apply}
 *  performs the splices in request order. Untracked only — the tracked shape
 *  (`<w:moveFrom>`/`<w:moveTo>`) isn't written yet, so the CLI refuses a move
 *  while tracking is on. */
export class Move {
	constructor(private readonly document: Document) {}

	/** Resolve every request's source and destination to live node refs, then
	 *  run every refusal — before anything moves. Throws a {@link MoveError}
	 *  (with `entry` set when there is more than one request). */
	plan(requests: MoveRequest[]): PlannedMove[] {
		const planned = requests.map((request, entry) =>
			withEntry(requests.length > 1 ? entry : undefined, () =>
				this.planOne(request),
			),
		);
		if (planned.length > 1) assertBatchConsistent(planned);
		return planned;
	}

	/** Splice every planned move in order and return the nodes that now carry
	 *  the moved content (a moved paragraph that filled an empty cell's
	 *  mandatory paragraph is reported as that paragraph). */
	apply(planned: PlannedMove[]): XmlNode[] {
		const placed: XmlNode[] = [];
		const cursors = new Map<XmlNode, CellInsertionCursor>();
		const lastPlacedAfter = new Map<XmlNode, XmlNode>();
		const touchedContainers = new Set<XmlNode[]>();
		for (const { source, destination } of planned) {
			detach(source);
			touchedContainers.add(source.parent);
			if (destination.kind === "cell") {
				placed.push(...placeInCell(destination, source, cursors));
				continue;
			}
			placeBesideBlock(destination, source.nodes, lastPlacedAfter);
			touchedContainers.add(destination.parent);
			placed.push(...source.nodes);
		}
		for (const cursor of cursors.values()) cursor.ensureTerminalParagraph();
		for (const container of touchedContainers) {
			this.keepContainerValid(container);
		}
		return placed;
	}

	private planOne(request: MoveRequest): PlannedMove {
		const source = resolveMoveSource(
			this.document,
			request.at,
			request.subtree ?? false,
		);
		assertMovableSource(source);
		const destination = resolveMoveDestination(
			this.document,
			request.to,
			request.side,
		);
		assertSameStory(source, destination);
		assertDestinationOutsideSource(source, destination);
		assertNotNoOp(source, destination);
		return { source, destination };
	}

	/** A cell or text-box story must still hold a paragraph (and a cell must
	 *  END with one) after its blocks moved out — Word rejects a bare
	 *  `<w:tc/>`. The body is exempt: a move never takes content out of it
	 *  wholesale, and its trailing `<w:sectPr>` is never moved. */
	private keepContainerValid(container: XmlNode[]): void {
		if (container === this.document.body.body.children) return;
		ensureCellEndsWithParagraph(container);
	}
}

/** One relocation: the block(s) at `at` go `side` the block (or bare cell) at
 *  `to`. `subtree` widens a heading `at` to the heading plus everything under
 *  it, up to the next heading of the same or higher level. */
export type MoveRequest = {
	at: string;
	to: string;
	side: "before" | "after";
	subtree?: boolean;
};

export type PlannedMove = { source: MoveSource; destination: MoveDestination };

/** Run `body`, translating every locator/cell error it raises into a
 *  {@link MoveError} stamped with the batch entry index. */
function withEntry<T>(entry: number | undefined, body: () => T): T {
	try {
		return body();
	} catch (error) {
		if (error instanceof MoveError) {
			throw new MoveError(error.code, error.message, error.hint, entry);
		}
		if (error instanceof CellTargetError) {
			throw new MoveError(error.code, error.message, error.hint, entry);
		}
		if (error instanceof LocatorParseError) {
			throw new MoveError("INVALID_LOCATOR", error.message, undefined, entry);
		}
		if (error instanceof LocatorResolveError) {
			throw new MoveError("BLOCK_NOT_FOUND", error.message, undefined, entry);
		}
		throw error;
	}
}

/** Batch-only refusals: every locator addresses the document AS READ, so two
 *  entries may not claim the same block, and no entry may anchor on a block
 *  some entry is moving (where it lands would depend on entry order). */
function assertBatchConsistent(planned: PlannedMove[]): void {
	for (const [entry, move] of planned.entries()) {
		withEntry(entry, () => {
			for (const earlier of planned.slice(0, entry)) {
				assertSourcesDisjoint(earlier.source, move.source);
			}
			for (const [owner, other] of planned.entries()) {
				if (!destinationInside(move.destination, other.source)) continue;
				throw new MoveError(
					"USAGE",
					`${move.destination.locator} is moved by entry ${owner} (${other.source.locator}), so it can't also anchor a move`,
					"Anchor on a block no entry moves, or fold adjacent moves into one range (pN-pM).",
				);
			}
		});
	}
}

/** Take the source's nodes out of their parent by live `indexOf` (earlier
 *  moves may have shifted their siblings; the node refs stay valid). */
function detach(source: MoveSource): void {
	for (const node of source.nodes) {
		const index = source.parent.indexOf(node);
		if (index !== -1) source.parent.splice(index, 1);
	}
}

/** Splice beside a block anchor by its live index. Several moves anchored
 *  AFTER the same block stack in entry order (`lastPlacedAfter` is the
 *  per-anchor offset `insert --batch` uses); BEFORE stacks naturally. */
function placeBesideBlock(
	destination: Extract<MoveDestination, { kind: "block" }>,
	nodes: XmlNode[],
	lastPlacedAfter: Map<XmlNode, XmlNode>,
): void {
	const { anchor, parent, side } = destination;
	if (side === "before") {
		parent.splice(parent.indexOf(anchor), 0, ...nodes);
		return;
	}
	const previous = lastPlacedAfter.get(anchor);
	const after =
		previous && parent.includes(previous) ? previous : (anchor as XmlNode);
	parent.splice(parent.indexOf(after) + 1, 0, ...nodes);
	const last = nodes[nodes.length - 1];
	if (last) lastPlacedAfter.set(anchor, last);
}

/** Splice at a bare cell's start/end through one cursor per cell, so stacked
 *  moves keep entry order. The first move into a cell that held only Word's
 *  mandatory empty paragraph fills that paragraph rather than leaving a blank
 *  line above (or below) the moved content. */
function placeInCell(
	destination: Extract<MoveDestination, { kind: "cell" }>,
	source: MoveSource,
	cursors: Map<XmlNode, CellInsertionCursor>,
): XmlNode[] {
	const { cell } = destination;
	let cursor = cursors.get(cell.node);
	// An earlier entry may have moved that empty paragraph out; reuse only a
	// paragraph still in the cell.
	const candidate = cursor ? null : destination.reusableParagraph;
	const reuse = candidate && cell.parent.includes(candidate) ? candidate : null;
	if (!cursor) {
		cursor = CellInsertionCursor.open(cell);
		cursors.set(cell.node, cursor);
	}
	return cursor.insert(source.nodes, destination.side, reuse);
}
