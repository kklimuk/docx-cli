import type { Document } from "../ast/document";
import type { BlockRangeReference, BlockReference } from "../ast/document/body";
import type { Run, SectionType } from "../ast/types";
import { Paragraph, type ParagraphOptions } from "../blocks";
import { buildCodeBlockParagraphs, ensureCodeBlockStyles } from "../code-block";
import type { XmlNode } from "../parser";
import { applyColumns, applySectionType, wrapSectPrChange } from "../sections";
import { flipCheckboxTracked, flipCheckboxUntracked } from "../task-list";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "../track-changes";
import {
	applyFormattingPreservingEdit,
	applyTrackedRangeReplace,
	applyUntrackedRangeReplace,
	assertParagraphOnlyTrackedRange,
	TrackedRangeConflictError,
} from "../track-changes/replace";

/** Cross-cutting lens over "edit an existing block." Stateless — each method
 * takes an already-resolved `BlockReference` (or `BlockRangeReference`) plus
 * a spec, then provisions any styles it needs, dispatches between tracked
 * vs untracked machinery, and mutates the document in place. Throws
 * `EditError(code, message, hint?)` for domain failures (wrong locator
 * tag, tracked-range conflict with a non-paragraph block, no-op edits). */
export class Edit {
	constructor(private document: Document) {}

	section(
		blockRef: BlockReference,
		spec: { columns?: number; sectionType?: SectionType },
		opts: { authorFlag?: string } = {},
	): void {
		if (blockRef.node.tag !== "w:sectPr") {
			throw new EditError(
				"BLOCK_NOT_FOUND",
				`Section locator did not resolve to a section break`,
			);
		}
		if (this.document.isTrackChangesEnabled()) {
			wrapSectPrChange(
				blockRef.node,
				new TrackChanges(this.document).mintMeta(opts.authorFlag),
			);
		}
		applyColumns(blockRef.node, spec.columns);
		applySectionType(blockRef.node, spec.sectionType);
	}

	taskToggle(
		blockRef: BlockReference,
		checked: boolean,
		opts: { authorFlag?: string } = {},
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--task requires a paragraph locator; got a non-paragraph block",
			);
		}
		const tracked = this.document.isTrackChangesEnabled();
		const ok = tracked
			? flipCheckboxTracked(
					blockRef.node,
					checked,
					makeMetaMinter(this.document, opts.authorFlag),
				)
			: flipCheckboxUntracked(blockRef.node, checked);
		if (!ok) {
			throw new EditError(
				"USAGE",
				"--task requires a task-list paragraph (one with a leading <w:sdt><w14:checkbox/></w:sdt>)",
				"Use `docx read FILE --ast` to inspect; convert a plain bullet to a task by replacing the paragraph via `--runs`.",
			);
		}
	}

	paragraph(
		blockRef: BlockReference,
		spec: ParagraphContentSpec,
		opts: { authorFlag?: string; noFormatting?: boolean } = {},
	): void {
		const targetIndex = blockRef.parent.indexOf(blockRef.node);
		if (targetIndex === -1) {
			throw new EditError(
				"BLOCK_NOT_FOUND",
				"Block reference is stale (parent does not contain it)",
			);
		}

		this.document
			.ensureStyles()
			.ensureReferencedStyle(spec.paragraphOptions.style);
		if (spec.kind === "runs") {
			this.document.ensureStyles().ensureReferencedRunStyles(spec.runs);
		}

		const tracked = this.document.isTrackChangesEnabled();

		if (canPreserveFormatting(spec, opts.noFormatting ?? false)) {
			applyFormattingPreservingEdit(
				this.document,
				blockRef.node,
				spec.text,
				spec.paragraphOptions,
				opts.authorFlag,
				tracked,
			);
			return;
		}

		if (spec.kind === "code") {
			ensureCodeBlockStyles(this.document, spec.language);
		}
		const newParagraphs = buildNewParagraphs(spec);

		if (tracked) {
			applyTrackedRangeReplace(
				this.document,
				blockRef.parent,
				targetIndex,
				targetIndex,
				newParagraphs,
				opts.authorFlag,
			);
		} else {
			applyUntrackedRangeReplace(
				blockRef.parent,
				targetIndex,
				targetIndex,
				newParagraphs,
			);
		}
	}

	/** Range replace: `pN-pM`. No formatting preservation (Word's empirical
	 * model for paragraph-range replace is "del all old, ins all new"; no
	 * cross-paragraph LCS, and we match it). Rejects tracked ranges that span
	 * a non-paragraph block (most commonly a table) because the tracked-range
	 * walker injects `<w:pPr>` into every span block, which would corrupt
	 * `<w:tbl>`. */
	range(
		rangeRef: BlockRangeReference,
		spec: ParagraphContentSpec,
		opts: { authorFlag?: string } = {},
	): void {
		this.document
			.ensureStyles()
			.ensureReferencedStyle(spec.paragraphOptions.style);
		if (spec.kind === "runs") {
			this.document.ensureStyles().ensureReferencedRunStyles(spec.runs);
		}
		if (spec.kind === "code") {
			ensureCodeBlockStyles(this.document, spec.language);
		}

		const tracked = this.document.isTrackChangesEnabled();
		if (tracked) {
			try {
				assertParagraphOnlyTrackedRange(rangeRef);
			} catch (error) {
				if (error instanceof TrackedRangeConflictError) {
					throw new EditError(
						"TRACKED_CHANGE_CONFLICT",
						error.message,
						error.hint,
					);
				}
				throw error;
			}
		}

		const newParagraphs = buildNewParagraphs(spec);
		if (tracked) {
			applyTrackedRangeReplace(
				this.document,
				rangeRef.parent,
				rangeRef.startIndex,
				rangeRef.endIndex,
				newParagraphs,
				opts.authorFlag,
			);
		} else {
			applyUntrackedRangeReplace(
				rangeRef.parent,
				rangeRef.startIndex,
				rangeRef.endIndex,
				newParagraphs,
			);
		}
	}
}

/** The paragraph-content specs that produce one or more new paragraphs.
 * Shared between `Edit.paragraph` (single block) and `Edit.range` (block
 * range); equation/task/section have their own method signatures. The
 * `markdown-blocks` variant carries pre-built XmlNodes from a prior
 * `new MarkdownImport(document).blocks(source)` — the CLI does the async
 * parse before calling into the lens, so the lens stays synchronous. */
export type ParagraphContentSpec =
	| {
			kind: "text";
			text: string;
			format: TextFormatting;
			paragraphOptions: ParagraphOptions;
	  }
	| { kind: "runs"; runs: Run[]; paragraphOptions: ParagraphOptions }
	| {
			kind: "code";
			content: string;
			language?: string;
			paragraphOptions: ParagraphOptions;
	  }
	| {
			kind: "markdown-blocks";
			blocks: XmlNode[];
			paragraphOptions: ParagraphOptions;
	  };

type TextFormatting = {
	color?: string;
	bold?: boolean;
	italic?: boolean;
};

/** Domain error from `Edit.*`. `code` is a literal subset of the CLI's
 * `ErrorCode` union so callers can `return fail(err.code, err.message,
 * err.hint)` directly — no cast, full type-check coverage. */
export type EditErrorCode =
	| "USAGE"
	| "BLOCK_NOT_FOUND"
	| "TRACKED_CHANGE_CONFLICT";

export class EditError extends Error {
	constructor(
		public code: EditErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "EditError";
	}
}

/** The formatting-preservation path applies only to `--text` (not `--runs`,
 *  which already lets the agent specify per-run formatting). It also bows
 *  out when the agent passed any explicit run-level format flag — those
 *  apply uniformly to the new paragraph, which conflicts with per-token
 *  inheritance. `--no-formatting` is the explicit opt-out. */
function canPreserveFormatting(
	spec: ParagraphContentSpec,
	noFormatting: boolean,
): spec is Extract<ParagraphContentSpec, { kind: "text" }> {
	if (noFormatting) return false;
	if (spec.kind !== "text") return false;
	const format = spec.format;
	if (format.color || format.bold || format.italic) return false;
	return true;
}

/** Build the new paragraph(s) for a paragraph-content spec. Text/runs produce
 *  a single paragraph; code produces one paragraph per source line via
 *  `buildCodeBlockParagraphs`. The single-anchor edit path routes a multi-
 *  paragraph result through `applyTrackedRangeReplace` / `applyUntrackedRangeReplace`
 *  with `startIndex === endIndex` (M=1, N=K), so multi-line code lands cleanly. */
function buildNewParagraphs(spec: ParagraphContentSpec): XmlNode[] {
	if (spec.kind === "code") {
		return buildCodeBlockParagraphs(spec.content, spec.language);
	}
	if (spec.kind === "text") {
		return [
			<Paragraph
				text={spec.text}
				{...spec.paragraphOptions}
				{...(spec.format.color ? { color: spec.format.color } : {})}
				{...(spec.format.bold ? { bold: true as const } : {})}
				{...(spec.format.italic ? { italic: true as const } : {})}
			/>,
		];
	}
	if (spec.kind === "markdown-blocks") {
		// Pre-built by the CLI via `MarkdownImport.blocks(...)`. The lens does
		// nothing else — the markdown walker has already provisioned styles,
		// allocated list numIds, registered footnote bodies, and minted image
		// rels on the document. We just splice these blocks where the locator
		// pointed.
		return spec.blocks;
	}
	return [<Paragraph runs={spec.runs} {...spec.paragraphOptions} />];
}

function makeMetaMinter(
	document: Document,
	authorFlag: string | undefined,
): () => TrackedMeta {
	const allocator = new TrackChanges(document).createAllocator();
	const author = resolveAuthor(authorFlag);
	const date = resolveDate();
	return () => ({ author, date, revisionId: allocator.next() });
}
