import type { Document } from "../ast/document";
import type { BlockRangeReference, BlockReference } from "../ast/document/body";
import type { Run, SectionType } from "../ast/types";
import {
	applyParagraphOptionsInPlace,
	insertPprChildInOrder,
	Paragraph,
	type ParagraphOptions,
} from "../blocks";
import { buildCodeBlockParagraphs, ensureCodeBlockStyles } from "../code-block";
import { Comments } from "../comments";
import { clearFormatting as clearRunFormatting } from "./clear-formatting";

export { CLEARABLE_ATTRS, resolveClearTags } from "./clear-formatting";

import {
	extractCommentMarkers,
	type ParagraphCommentMarker,
	paragraphTextLength,
	reanchorCommentMarkers,
} from "../comments/markers";
import { replaceSpanInParagraph, type TrackedReplaceOptions } from "../find";
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
		opts: { authorFlag?: string; track?: boolean } = {},
	): void {
		if (blockRef.node.tag !== "w:sectPr") {
			throw new EditError(
				"BLOCK_NOT_FOUND",
				`Section locator did not resolve to a section break`,
			);
		}
		if (opts.track ?? this.document.isTrackChangesEnabled()) {
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
		opts: { authorFlag?: string; track?: boolean } = {},
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--task requires a paragraph locator; got a non-paragraph block",
			);
		}
		const tracked = opts.track ?? this.document.isTrackChangesEnabled();
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
		opts: { authorFlag?: string; noFormatting?: boolean; track?: boolean } = {},
	): XmlNode {
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

		const tracked = opts.track ?? this.document.isTrackChangesEnabled();

		// Lift any comment range markers out of the old paragraph BEFORE its
		// content is rebuilt, so they can be re-anchored to the new content
		// instead of collapsing to a zero-length range (the orphaned-comment bug).
		const commentMarkers = extractCommentMarkers(blockRef.node);

		if (canPreserveFormatting(spec, opts.noFormatting ?? false)) {
			applyFormattingPreservingEdit(
				this.document,
				blockRef.node,
				spec.text,
				spec.paragraphOptions,
				opts.authorFlag,
				tracked,
			);
			this.reanchorComments(blockRef.node, commentMarkers);
			// Mutated in place — the same <w:p> node is the result.
			return blockRef.node;
		}

		if (spec.kind === "code") {
			ensureCodeBlockStyles(this.document, spec.language);
		}
		const newParagraphs = buildNewParagraphs(spec);
		inheritParagraphFormattingIfPlain(
			blockRef.node,
			newParagraphs,
			spec.paragraphOptions.style,
		);
		const anchorTarget = newParagraphs[0];
		if (anchorTarget?.tag === "w:p") {
			this.reanchorComments(anchorTarget, commentMarkers);
		} else {
			this.resolveComments(commentMarkers);
		}

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
		// The spliced-in first paragraph is the result (a following clear in a
		// combined content+clear edit targets this node, not the replaced one).
		return anchorTarget ?? blockRef.node;
	}

	/** Style-only edit: re-apply paragraph properties (`--style`/`--alignment`)
	 *  in place, keeping every existing run. The symmetric "restyle without
	 *  retyping" to the content path's `--style` ride-along — Word's `<w:pPrChange>`
	 *  isn't modeled, so (like the ride-along) this applies directly rather than as
	 *  a tracked revision, regardless of the document's track-changes toggle. */
	paragraphProperties(
		blockRef: BlockReference,
		options: ParagraphOptions,
	): XmlNode {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--style/--alignment alone restyle a paragraph; this locator is not a paragraph.",
			);
		}
		this.document.ensureStyles().ensureReferencedStyle(options.style);
		applyParagraphOptionsInPlace(blockRef.node.children, options);
		return blockRef.node;
	}

	/** Character-span replace: `pN:S-E` (or a cell paragraph `tN:rRcC:pK:S-E`).
	 * Replaces exactly the text in `[start, end)` with `replacement`, leaving the
	 * paragraph's `<w:pPr>` and every other run untouched. The replacement run
	 * inherits the `<w:rPr>` of the run at the span start (so font/size/color/etc.
	 * survive) — this is the keystone that lets `find → edit --at <span>` work
	 * without rewriting the whole paragraph. Reuses `replaceSpanInParagraph`, the
	 * same machinery `replace` uses; under tracking the cut is `<w:del>` and the
	 * replacement `<w:ins>`. Offsets are accepted-view, matching `find`'s output. */
	span(
		blockRef: BlockReference,
		span: { start: number; end: number },
		replacement: string,
		opts: { authorFlag?: string; track?: boolean } = {},
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"A character-span locator (pN:S-E) edits text inside a paragraph; this locator does not resolve to a paragraph.",
			);
		}
		const length = paragraphTextLength(blockRef.node, "accepted");
		if (span.end > length) {
			throw new EditError(
				"INVALID_LOCATOR",
				`Span ${span.start}-${span.end} is out of range (the paragraph has ${length} characters)`,
				'Run `docx find FILE "phrase"` to get an exact span locator.',
			);
		}
		const tracked: TrackedReplaceOptions | undefined =
			(opts.track ?? this.document.isTrackChangesEnabled())
				? {
						meta: {
							author: resolveAuthor(opts.authorFlag),
							date: resolveDate(),
						},
						allocator: new TrackChanges(this.document).createAllocator(),
					}
				: undefined;
		replaceSpanInParagraph(
			blockRef.node,
			span,
			replacement,
			tracked,
			"accepted",
		);
	}

	/** Re-place the comment markers snapshotted before an edit so they bracket
	 *  the rebuilt paragraph; any comment whose anchor text is entirely gone
	 *  (empty new paragraph) is marked resolved instead. */
	private reanchorComments(
		paragraph: XmlNode,
		markers: ParagraphCommentMarker[],
	): void {
		if (markers.length === 0) return;
		const orphaned = reanchorCommentMarkers(paragraph, markers, "current");
		this.resolveComments(
			markers.filter((marker) => orphaned.includes(marker.id)),
		);
	}

	/** Mark the comments behind these markers resolved (used when an edit
	 *  removes the anchor's content and there's nothing left to bracket). */
	private resolveComments(markers: ParagraphCommentMarker[]): void {
		const ids = [...new Set(markers.map((marker) => marker.id))].filter((id) =>
			this.document.comments?.findById(id),
		);
		if (ids.length > 0) new Comments(this.document).resolve(ids, true);
	}

	/** Strip run-level formatting (the `tags` set names `<w:rPr>` child elements)
	 *  from a whole paragraph (`span` null) or just the runs overlapping a
	 *  character span — keeping the text. The inverse of authoring formatting;
	 *  pairs with `find --highlight … | edit --clear highlight`. Mutates rPr in
	 *  place so unmodelled run properties survive. */
	clearFormatting(
		blockRef: BlockReference,
		span: { start: number; end: number } | null,
		tags: Set<string>,
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--clear requires a paragraph or character-span locator",
			);
		}
		clearRunFormatting(blockRef.node, span, tags);
	}

	/** Like `clearFormatting` but targets a paragraph node directly. Used by the
	 *  combined content+clear edit, where the content step may have spliced in a
	 *  fresh paragraph node (so the original blockRef is stale): `paragraph()`
	 *  returns the resulting node and we clear THAT. The message differs from the
	 *  locator path's: here the locator WAS a paragraph; it's the new content
	 *  (e.g. `--markdown` that produced a table) that isn't clearable. */
	clearFormattingNode(
		node: XmlNode,
		span: { start: number; end: number } | null,
		tags: Set<string>,
	): void {
		if (node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--clear can't apply: the new content isn't a single paragraph",
				"Drop --clear (a table/structural block has no run formatting to strip), or clear separately with `edit --at <pN> --clear …` after the content edit.",
			);
		}
		clearRunFormatting(node, span, tags);
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
		opts: { authorFlag?: string; track?: boolean } = {},
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

		const tracked = opts.track ?? this.document.isTrackChangesEnabled();
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
	| "INVALID_LOCATOR"
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
	// Tabs/newlines are fine: the preserve-path emitter splits them into
	// <w:tab/>/<w:br/> within each rPr-bearing run, so a "**Name**⇥date" line
	// keeps its per-segment formatting instead of flattening to one plain run.
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

/** Decision 2 (style preservation): when a whole-paragraph edit replaces a
 *  paragraph with a single PLAIN paragraph (no explicit `--style`, and the new
 *  content didn't set its own `<w:pStyle>` — e.g. `--markdown "Q3 Title"` on a
 *  Heading, or `--runs`/`--text --bold`), inherit the old paragraph's style so
 *  re-titling a heading keeps it a heading. `--text` without overrides already
 *  preserves style via the formatting-preserving path; this covers the other
 *  paths. Markdown that carries its own block style (a `#` heading, a list) sets
 *  `<w:pStyle>` itself, so the guard below leaves it alone. */
function inheritParagraphFormattingIfPlain(
	oldParagraph: XmlNode,
	newParagraphs: XmlNode[],
	explicitStyle: string | undefined,
): void {
	if (newParagraphs.length !== 1) return;
	const newParagraph = newParagraphs[0];
	if (!newParagraph || newParagraph.tag !== "w:p") return;
	// Content that brings its own block style (a markdown `#` heading, a list)
	// owns its <w:pPr> — don't overwrite it with the old paragraph's.
	if (newParagraph.findChild("w:pPr")?.findChild("w:pStyle")) return;
	const oldPpr = oldParagraph.findChild("w:pPr");
	if (!oldPpr) return;

	// Base the replacement's paragraph properties on the OLD <w:pPr> (a clone, in
	// its valid child order) so DIRECT formatting survives a fresh-run replace —
	// alignment (w:jc), spacing, indent, keepNext, list membership, etc. Only the
	// style was carried before, which silently dropped centering on `edit --text`.
	const merged = oldPpr.clone();
	// If the caller passed an explicit --style, applyParagraphOptions will set it;
	// drop the inherited one so it doesn't fight.
	if (explicitStyle) {
		merged.children = merged.children.filter(
			(child) => child.tag !== "w:pStyle",
		);
	}
	// Overlay any pPr children the new paragraph already set (explicit --alignment
	// etc.) so caller options win over the inherited ones.
	const newPpr = newParagraph.findChild("w:pPr");
	if (newPpr) {
		for (const child of newPpr.children) {
			const index = merged.children.findIndex((own) => own.tag === child.tag);
			// Replace a same-tag inherited child in place (keeps its valid slot), or
			// splice a new one at its canonical CT_PPr position — never push to the
			// end, where it would land AFTER the paragraph-mark <w:rPr> the clone
			// carries and trip Word's "unreadable content" repair (e.g. <w:jc> from
			// --alignment).
			if (index >= 0) merged.children[index] = child;
			else insertPprChildInOrder(merged, child);
		}
		newParagraph.children[newParagraph.children.indexOf(newPpr)] = merged;
	} else {
		newParagraph.children.unshift(merged);
	}
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
