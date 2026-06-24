import type { XmlNode } from "../../parser";
import { flattenImageRuns } from "../text";
import type {
	Block,
	Comment,
	DocProperties,
	Marginal,
	Note,
	TrackedChangeKind,
} from "../types";

export class Body {
	schemaVersion = 1 as const;
	path: string;
	properties: DocProperties;
	blocks: Block[];
	comments: Comment[];
	footnotes: Note[];
	endnotes: Note[];
	headers: Marginal[];
	footers: Marginal[];

	body: XmlNode;
	blockReferences: Map<string, BlockReference> = new Map();
	imageById: Map<string, ImageReference> = new Map();
	hyperlinkById: Map<string, HyperlinkReference> = new Map();
	equationReferences: Map<string, EquationReference> = new Map();

	constructor(init: {
		path: string;
		properties: DocProperties;
		blocks: Block[];
		comments: Comment[];
		footnotes: Note[];
		endnotes: Note[];
		headers: Marginal[];
		footers: Marginal[];
		body: XmlNode;
	}) {
		this.path = init.path;
		this.properties = init.properties;
		this.blocks = init.blocks;
		this.comments = init.comments;
		this.footnotes = init.footnotes;
		this.endnotes = init.endnotes;
		this.headers = init.headers;
		this.footers = init.footers;
		this.body = init.body;
	}

	*iterateBlocks(): IterableIterator<Block> {
		yield* iterateBlocks(this.blocks);
	}

	findBlockById(blockId: string): Block | null {
		for (const block of this.iterateBlocks()) {
			if (block.id === blockId) return block;
		}
		return null;
	}

	resolveBlock(blockId: string): BlockReference {
		const reference = this.blockReferences.get(blockId);
		if (!reference) {
			throw new LocatorResolveError(
				{ kind: "block", blockId },
				`Block not found: ${blockId}`,
			);
		}
		return reference;
	}

	resolveBlockRange(
		startBlockId: string,
		endBlockId: string,
	): BlockRangeReference {
		const start = this.resolveBlock(startBlockId);
		const end = this.resolveBlock(endBlockId);
		if (start.parent !== end.parent) {
			throw new LocatorResolveError(
				{ kind: "blockRange", startBlockId, endBlockId },
				`Range endpoints ${startBlockId} and ${endBlockId} live in different containers — they must be siblings`,
			);
		}
		const startIndex = start.parent.indexOf(start.node);
		const endIndex = end.parent.indexOf(end.node);
		if (startIndex === -1 || endIndex === -1) {
			throw new LocatorResolveError(
				{ kind: "blockRange", startBlockId, endBlockId },
				"Range endpoint became detached from its parent (stale block reference)",
			);
		}
		if (endIndex < startIndex) {
			throw new LocatorResolveError(
				{ kind: "blockRange", startBlockId, endBlockId },
				`Range ${startBlockId}-${endBlockId} runs backwards — ${endBlockId} appears before ${startBlockId} in document order`,
			);
		}
		return { parent: start.parent, startIndex, endIndex };
	}

	findBodyChildren(): XmlNode[] {
		return this.body.children;
	}

	listBlockIds(): string[] {
		return [...this.blockReferences.keys()];
	}

	listImageIds(): string[] {
		return [...this.imageById.keys()];
	}

	/** Find an image by its SHA-256 content hash. Used by the markdown
	 * importer to round-trip `![alt](<sha256>.<ext>)` references without
	 * re-fetching: when the reader emits an image, it identifies the bytes
	 * by their hash; on import we look the hash up against the target
	 * doc's media and reuse the existing relationship if it matches.
	 * Returns `undefined` if no image in this body carries that hash. */
	findImageByHash(hash: string): ImageReference | undefined {
		for (const run of flattenImageRuns(this.blocks)) {
			if (run.hash === hash) return this.imageById.get(run.id);
		}
		return undefined;
	}

	listHyperlinkIds(): string[] {
		return [...this.hyperlinkById.keys()];
	}

	listEquationIds(): string[] {
		return [...this.equationReferences.keys()];
	}

	toJSON(): object {
		return {
			schemaVersion: this.schemaVersion,
			path: this.path,
			properties: this.properties,
			blocks: this.blocks,
			comments: this.comments,
			footnotes: this.footnotes,
			endnotes: this.endnotes,
			headers: this.headers,
			footers: this.footers,
		};
	}
}

/** Recursive block iterator that descends into table cells. Stand-alone helper
 *  so callers with just `Block[]` (no `Body` instance) can share the walker. */
export function* iterateBlocks(blocks: Block[]): IterableIterator<Block> {
	for (const block of blocks) {
		yield block;
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					yield* iterateBlocks(cell.blocks);
				}
			}
		}
	}
}

export class LocatorResolveError extends Error {
	constructor(
		public locator: { kind: string; [key: string]: unknown },
		message: string,
	) {
		super(message);
		this.name = "LocatorResolveError";
	}
}

export type BlockReference = { node: XmlNode; parent: XmlNode[] };

export type BlockRangeReference = {
	parent: XmlNode[];
	startIndex: number;
	endIndex: number;
};

export type ImageReference = {
	relationshipId: string;
	partName: string;
	contentType: string;
};

export type HyperlinkReference = {
	node: XmlNode;
	parent: XmlNode[];
	relationshipId?: string;
};

export type EquationReference = {
	/** The `<m:oMath>` element (or its `<m:oMathPara>` parent for display
	 *  equations; the walker handles both — see `core/equation/read.ts`). */
	node: XmlNode;
	parent: XmlNode[];
	blockId: string;
	display: boolean;
	/** Reconstructed LaTeX from the reader's `ommlToLatex(node)` pass. Cached
	 *  here so `edit --at eqN --display`/`--inline` mode-toggles work for
	 *  equations inside table cells too — walking `view.doc.blocks` to find
	 *  the run's `latex` field misses cell paragraphs. */
	latex: string;
};

// Exported here (not on `CommentsView`) so `Body` types don't depend on
// `CommentsView`, which in turn imports a `CommentReference` for its own
// `commentReferences` map. The shape is trivial enough that either home is
// fine; keeping it next to the other reference types is the pragmatic call.
export type CommentReference = { node: XmlNode; parent: XmlNode[] };

export type TrackedChangeReference = {
	node: XmlNode;
	parent: XmlNode[];
	blockId: string;
	/** Explicit kind for revisions whose kind can't be derived from the wrapper
	 * tag alone. Two cases:
	 *  - Ambiguous tag: a row revision is a `<w:ins>`/`<w:del>` inside
	 *    `<w:trPr>`, sharing its tag with run-level changes — same for
	 *    cellIns/cellDel inside `<w:tcPr>`.
	 *  - Structural pattern: `checkboxToggle` points at the SDT itself; the
	 *    inner `<w:ins>`/`<w:del>` glyph pair is recognized via
	 *    `findCheckboxToggle` and never surfaces as its own reference.
	 * Set for rowIns/rowDel/cellIns/cellDel and checkboxToggle; absent for
	 * run-level changes (kind derived from tag). */
	kind?: TrackedChangeKind;
	/** Set for rowIns/rowDel: the owning `<w:tr>` and the table's children, so
	 * accept/reject can remove the whole row. (Paragraph-mark scope is NOT
	 * stored — it's derivable from `blockId` via `Body.blockReferences`.) */
	tableRow?: XmlNode;
	tableRowParent?: XmlNode[];
	/** Set for cellIns/cellDel: the owning `<w:tc>` and the row's children. */
	tableCell?: XmlNode;
	tableCellParent?: XmlNode[];
};
