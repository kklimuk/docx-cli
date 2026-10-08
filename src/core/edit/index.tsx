import type { Document } from "../ast/document";
import type { BlockRangeReference, BlockReference } from "../ast/document/body";
import type { Run, SectionType } from "../ast/types";
import {
	applyParagraphOptionsInPlace,
	applyParagraphOptionsToBlocks,
	ensureParagraphProperties,
	hasParagraphProperties,
	inheritRprChild,
	injectPprChange,
	isInheritableRunProperty,
	Paragraph,
	type ParagraphOptions,
	priorPprChildren,
	wrapPprChange,
} from "../blocks";
import { buildCodeBlockParagraphs, ensureCodeBlockStyles } from "../code-block";
import { Comments } from "../comments";
import { clearFormatting as clearRunFormatting } from "./clear-formatting";
import {
	type RunFormat,
	setFormatting as setRunFormatting,
} from "./set-formatting";

export { CLEARABLE_ATTRS, resolveClearTags } from "./clear-formatting";
export type { RunFormat } from "./set-formatting";

import {
	extractCommentMarkers,
	type ParagraphCommentMarker,
	paragraphTextLength,
	reanchorCommentMarkers,
} from "../comments/markers";
import { matchDocumentLook } from "../dominant-formatting";
import { replaceSpanInParagraph, type TrackedReplaceOptions } from "../find";
import { continueList } from "../lists/continue-list";
import { hasTextBox } from "../mc";
import {
	bringsNewBlockStructure,
	inheritParagraphFormattingIfPlain,
} from "../paragraph-inheritance";
import {
	partitionParagraphRuns,
	TRACKED_CHANGE_WRAPPER_TAGS,
	XmlNode,
} from "../parser";
import {
	applyColumns,
	applyPageGeometry,
	applySectionType,
	type PageGeometry,
	wrapSectPrChange,
} from "../sections";
import { flipCheckboxTracked, flipCheckboxUntracked } from "../task-list";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "../track-changes";
import { watchPendingRevisions } from "../track-changes/pending-authors";
import { paragraphMarkRunRpr } from "../track-changes/preserve-formatting";
import {
	applyFormattingPreservingEdit,
	applyTrackedRangeReplace,
	applyUntrackedRangeReplace,
	assertParagraphOnlyTrackedRange,
	liftObjectRuns,
	reattachObjectRuns,
	TrackedRangeConflictError,
} from "../track-changes/replace";
import { markdownAsTextEdit } from "./markdown-as-text-edit";

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
		spec: { columns?: number; sectionType?: SectionType } & PageGeometry,
		opts: { authorFlag?: string; track?: boolean } = {},
	): void {
		if (blockRef.node.tag !== "w:sectPr") {
			throw new EditError(
				"BLOCK_NOT_FOUND",
				`Section locator did not resolve to a section break`,
			);
		}
		if (opts.track ?? this.document.isTrackChangesEnabled()) {
			// One snapshot captures the WHOLE prior sectPr (cols/type/pgSz/pgMar), so
			// any combination of the mutations below records as a single sectPrChange.
			wrapSectPrChange(
				blockRef.node,
				new TrackChanges(this.document).mintMeta(opts.authorFlag),
			);
		}
		applyColumns(blockRef.node, spec.columns);
		applySectionType(blockRef.node, spec.sectionType);
		applyPageGeometry(blockRef.node, {
			pageSize: spec.pageSize,
			orientation: spec.orientation,
			margins: spec.margins,
		});
	}

	taskToggle(
		blockRef: BlockReference,
		checked: boolean,
		opts: { authorFlag?: string; track?: boolean } = {},
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"tasks check/uncheck requires a paragraph locator; got a non-paragraph block",
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
				"tasks check/uncheck requires a task-list paragraph (one with a leading <w:sdt><w14:checkbox/></w:sdt>)",
				"Use `docx read FILE` to spot task lines (- [ ] / - [x]); author a new task with `docx tasks add`.",
			);
		}
	}

	paragraph(
		blockRef: BlockReference,
		spec: ParagraphContentSpec,
		opts: {
			authorFlag?: string;
			noFormatting?: boolean;
			track?: boolean;
			/** The caller restyles the returned paragraph's runs afterward (ride-
			 *  along `--size`/`--clear`): a markdown edit then replaces the whole
			 *  paragraph, so the restyle touches only the NEW runs — on the word-
			 *  diff path it would also restyle the kept old words, untracked. */
			restyleFollows?: boolean;
		} = {},
	): ParagraphEditResult {
		const targetIndex = blockRef.parent.indexOf(blockRef.node);
		if (targetIndex === -1) {
			throw new EditError(
				"BLOCK_NOT_FOUND",
				"Block reference is stale (parent does not contain it)",
			);
		}
		// Snapshot BEFORE the rebuild: the single-author content rebuild resolves
		// other authors' pending revisions without review (issue #17), so the
		// caller can warn about whichever ones didn't survive.
		const listResolvedAuthors = watchPendingRevisions(
			blockRef.parent,
			targetIndex,
			targetIndex,
			opts.authorFlag,
		);

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

		const textEdit = wordDiffText(blockRef.node, spec, opts, tracked);
		if (textEdit !== null) {
			applyFormattingPreservingEdit(
				this.document,
				blockRef.node,
				textEdit,
				spec.paragraphOptions,
				opts.authorFlag,
				tracked,
			);
			this.reanchorComments(blockRef.node, commentMarkers);
			// Mutated in place — the same <w:p> node is the result.
			return { node: blockRef.node, resolvedAuthors: listResolvedAuthors() };
		}

		if (spec.kind === "code") {
			ensureCodeBlockStyles(this.document, spec.language);
		}
		const newParagraphs = buildNewParagraphs(spec);
		// The replacement is built from content that never saw the old paragraph's
		// text box / image / embed, so lift those runs out first and put them back
		// on the live result below — the text-box invariant ("the box rides its
		// anchor RUN") holds on this path exactly as on the `--text` diff path.
		// Pictures the replacement re-states stay with the old content instead.
		const objectRuns = liftObjectRuns(
			blockRef.node,
			restatedPictures(newParagraphs),
		);
		// Run inheritance goes FIRST: its `bringsNewBlockStructure` guard must
		// see only the structure the new content brought (a markdown `#` heading,
		// a list item), not the old paragraph's `<w:pStyle>` that the pPr
		// inheritance below is about to stamp on. In the other order every
		// replacement in a styled paragraph (the résumé's BodyText contact line)
		// looked like it "owned" its style and inherited no run font — the fresh
		// runs fell back to the style chain (Times New Roman) while `read` still
		// showed them bare under `docx:base font="Calibri"`.
		//
		// `runs` is the explicit, byte-precise surface — the caller states each
		// run's rPr, so inheriting the old paragraph's would silently override
		// their choices (a `{"bold": false}` run can't opt out, since the emitter
		// drops the falsy child and the inherited `<w:b/>` fills the gap).
		if (!opts.noFormatting && spec.kind !== "code" && spec.kind !== "runs") {
			inheritCommonRunFormatting(blockRef.node, newParagraphs);
		}
		inheritParagraphFormattingIfPlain(
			blockRef.node,
			newParagraphs,
			spec.paragraphOptions.style,
		);
		applyMarkdownRideAlong(spec, newParagraphs);
		continueList(this.document, blockRef.node, newParagraphs);
		// Whatever the rPr inheritance above didn't donate (a `#` heading on a
		// body line brings new structure, so it inherits no rPr) takes the look
		// that was AT this position — the replaced paragraph's rendered face and
		// explicit size — before the document's dominant face/size.
		if (!opts.noFormatting && authorsFreshText(spec)) {
			matchDocumentLook(this.document, newParagraphs, {
				position: blockRef.node,
			});
		}
		const anchorTarget = newParagraphs[0];
		if (anchorTarget?.tag === "w:p") {
			this.reanchorComments(anchorTarget, commentMarkers);
		} else {
			this.resolveComments(commentMarkers);
		}

		if (tracked) {
			// Paragraph properties riding along with the content edit (style/alignment/
			// spacing/indent/tabs) are a tracked revision: snapshot the OLD paragraph's
			// prior `<w:pPr>` into a `<w:pPrChange>` on the new paragraph's pPr so reject
			// restores it. The fresh pPr already holds the NEW props, so we can't use
			// `wrapPprChange` (which snapshots current children) — supply the prior
			// snapshot explicitly via `injectPprChange`. `applyTrackedRangeReplace`'s
			// `replacePPr` then carries this pPr (marker included) onto the live node.
			if (
				hasParagraphProperties(spec.paragraphOptions) &&
				anchorTarget?.tag === "w:p"
			) {
				injectPprChange(
					ensureParagraphProperties(anchorTarget),
					priorPprChildren(blockRef.node.findChild("w:pPr")),
					new TrackChanges(this.document).mintMeta(opts.authorFlag),
				);
			}
			applyTrackedRangeReplace(
				this.document,
				blockRef.parent,
				targetIndex,
				targetIndex,
				newParagraphs,
				opts.authorFlag,
			);
			// Under tracking the OLD node stays as the live "transition" paragraph
			// (new content appended inside `<w:ins>`), so the untracked carriers go
			// back onto it — after the replace, or they'd be swept into the `<w:del>`.
			reattachObjectRuns(blockRef.node, objectRuns);
		} else {
			if (objectRuns.leading.length > 0 || objectRuns.trailing.length > 0) {
				reattachObjectRuns(objectCarrierTarget(newParagraphs), objectRuns);
			}
			applyUntrackedRangeReplace(
				blockRef.parent,
				targetIndex,
				targetIndex,
				newParagraphs,
			);
		}
		// The first NEW paragraph is the result: untracked it's the spliced-in
		// node; tracked, its runs are the very nodes now inside the transition's
		// `<w:ins>`. A following clear/format in a combined content+clear edit
		// must target only those — never the transition itself, whose `<w:del>`
		// holds the original runs reject restores.
		return {
			node: anchorTarget ?? blockRef.node,
			resolvedAuthors: listResolvedAuthors(),
		};
	}

	/** Properties-only edit: re-apply paragraph properties (`--style`/`--alignment`/
	 *  `--space-*`/`--line-spacing`/`--indent-*`/`--tabs`) in place, keeping every
	 *  existing run — the "restyle without retyping" twin of the content path's
	 *  ride-along. Under track-changes (the doc toggle or `opts.track`), the prior
	 *  `<w:pPr>` is snapshotted into a `<w:pPrChange>` BEFORE the mutation, so the
	 *  change is a real tracked revision (accept drops the marker, reject restores
	 *  the prior pPr) — empirically the shape Word emits for ANY paragraph-property
	 *  change. Mirrors `Edit.section`/`wrapSectPrChange`. */
	paragraphProperties(
		blockRef: BlockReference,
		options: ParagraphOptions,
		opts: { authorFlag?: string; track?: boolean } = {},
	): XmlNode {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"--style/--alignment alone restyle a paragraph; this locator is not a paragraph.",
			);
		}
		this.document.ensureStyles().ensureReferencedStyle(options.style);
		if (opts.track ?? this.document.isTrackChangesEnabled()) {
			wrapPprChange(
				ensureParagraphProperties(blockRef.node),
				new TrackChanges(this.document).mintMeta(opts.authorFlag),
			);
		}
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

	/** Set run-level formatting (bold/italic/underline/color/highlight/font/size/
	 *  …) on a whole paragraph (`span` null) or just the runs overlapping a
	 *  character span — keeping the text. The inverse of `clearFormatting`: where
	 *  clear strips an `<w:rPr>` child, set adds/replaces it (find-or-creating the
	 *  rPr, splicing children in CT_RPr order). Like clear, it mutates rPr in place
	 *  so unmodelled run properties survive, and — like `paragraphProperties` and
	 *  clear — it applies DIRECTLY regardless of the track-changes toggle: Word's
	 *  `<w:rPrChange>` isn't modeled, so a formatting change is never recorded as a
	 *  tracked revision (see `src/cli/track-changes` — rPrChange/pPrChange are
	 *  out of scope for accept/reject). */
	setFormatting(
		blockRef: BlockReference,
		span: { start: number; end: number } | null,
		format: RunFormat,
	): void {
		if (blockRef.node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"Run-formatting flags require a paragraph or character-span locator",
			);
		}
		setRunFormatting(blockRef.node, span, format);
	}

	/** Like `setFormatting` but targets a paragraph node directly — used by the
	 *  combined content+format edit, where the content step may have spliced in a
	 *  fresh paragraph node (so the original blockRef is stale): `paragraph()` /
	 *  `span()` returns the resulting node and we format THAT. */
	setFormattingNode(
		node: XmlNode,
		span: { start: number; end: number } | null,
		format: RunFormat,
	): void {
		if (node.tag !== "w:p") {
			throw new EditError(
				"USAGE",
				"Run formatting can't apply: the new content isn't a single paragraph",
				"Drop the formatting flags (a table/structural block has no runs to format), or set them separately with `edit --at <pN> …` after the content edit.",
			);
		}
		setRunFormatting(node, span, format);
	}

	/** Range replace: `pN-pM`. No formatting preservation (Word's empirical
	 * model for paragraph-range replace is "del all old, ins all new"; no
	 * cross-paragraph LCS, and we match it). Rejects tracked ranges that span
	 * a non-paragraph block (most commonly a table) because the tracked-range
	 * walker injects `<w:pPr>` into every span block, which would corrupt
	 * `<w:tbl>`. Returns the other authors whose pending revisions the replace
	 * resolved without review (issue #17), for the caller to warn about. */
	range(
		rangeRef: BlockRangeReference,
		spec: ParagraphContentSpec,
		opts: { authorFlag?: string; track?: boolean; noFormatting?: boolean } = {},
	): string[] {
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

		const listResolvedAuthors = watchPendingRevisions(
			rangeRef.parent,
			rangeRef.startIndex,
			rangeRef.endIndex,
			opts.authorFlag,
		);
		const newParagraphs = buildNewParagraphs(spec);
		applyMarkdownRideAlong(spec, newParagraphs);
		const firstReplaced = rangeRef.parent[rangeRef.startIndex];
		if (!opts.noFormatting && authorsFreshText(spec)) {
			matchDocumentLook(this.document, newParagraphs, {
				position: firstReplaced?.tag === "w:p" ? firstReplaced : undefined,
			});
		}
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
		return listResolvedAuthors();
	}
}

/** `Edit.paragraph`'s result: the paragraph node the content edit produced,
 *  and the other authors whose pending revisions the rebuild resolved. */
export type ParagraphEditResult = { node: XmlNode; resolvedAuthors: string[] };

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

/** The text `Edit.paragraph` applies as a word-level diff
 *  (`applyFormattingPreservingEdit`), or null when the edit replaces the whole
 *  paragraph. `--text` qualifies per `canPreserveFormatting`; a single-paragraph
 *  markdown rewrite that changes no visible formatting is a TEXT edit too, so a
 *  clause rewrite redlines only the changed words and unchanged words keep their
 *  runs (see `markdownAsTextEdit`) — unless the caller restyles the result
 *  afterward (`restyleFollows`), which would then restyle the kept old words
 *  untracked. */
function wordDiffText(
	paragraph: XmlNode,
	spec: ParagraphContentSpec,
	opts: { noFormatting?: boolean; restyleFollows?: boolean },
	tracked: boolean,
): string | null {
	if (tracked && hasPendingRevisions(paragraph)) return null;
	if (opts.noFormatting) return null;
	if (canPreserveFormatting(spec, false)) return spec.text;
	if (spec.kind !== "markdown-blocks" || opts.restyleFollows) return null;
	return markdownAsTextEdit(paragraph, spec.blocks);
}

/** Whether a paragraph already carries pending tracked insertions/deletions/
 *  moves. The `--text` word-diff rebuilds a paragraph from its VISIBLE runs, so
 *  under tracking it would unwrap a prior `<w:ins>` into plain text and drop a
 *  prior `<w:del>` outright — a second tracked `--text` edit of a clause baked
 *  the first rewrite into the baseline and erased the counterparty's original
 *  wording, unrecoverable by reject (haiku contract-markup, 2026.10.01 r1). Such
 *  a paragraph takes the whole-paragraph tracked replace instead, which keeps
 *  every prior deletion inside its `<w:del>`: a coarser redline that keeps the
 *  original wording. A text box's story is its own set of paragraphs, so its
 *  revisions don't make the anchor paragraph "pending". */
function hasPendingRevisions(paragraph: XmlNode): boolean {
	for (const child of paragraph.children) {
		if (TRACKED_CHANGE_WRAPPER_TAGS.has(child.tag)) return true;
		if (child.tag === "w:txbxContent") continue;
		if (child.children.length > 0 && hasPendingRevisions(child)) return true;
	}
	return false;
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

/** Content whose runs the agent didn't specify byte-for-byte — markdown and
 * `--text` — so they take the document's look. `--runs` states its own rPr and
 * a code block owns its monospace. */
function authorsFreshText(spec: ParagraphContentSpec): boolean {
	return spec.kind === "markdown-blocks" || spec.kind === "text";
}

/** Build the new paragraph(s) for a paragraph-content spec. Text/runs produce
 *  a single paragraph; code produces one paragraph per source line via
 *  `buildCodeBlockParagraphs`. The single-anchor edit path routes a multi-
 *  paragraph result through `applyTrackedRangeReplace` / `applyUntrackedRangeReplace`
 *  with `startIndex === endIndex` (M=1, N=K), so multi-line code lands cleanly. */
function buildNewParagraphs(spec: ParagraphContentSpec): XmlNode[] {
	if (spec.kind === "code") {
		return buildCodeBlockParagraphs(
			spec.content,
			spec.language,
			spec.paragraphOptions,
		);
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
		// Pre-built by the CLI via `MarkdownImport.blocks(...)` — the markdown
		// walker has already provisioned styles, allocated list numIds, registered
		// footnote bodies, and minted image rels on the document. Its ride-along
		// paragraph properties land later, in `applyMarkdownRideAlong`.
		return spec.blocks;
	}
	return [<Paragraph runs={spec.runs} {...spec.paragraphOptions} />];
}

/** A markdown source's ride-along paragraph options (see
 *  `applyParagraphOptionsToBlocks`); other content kinds apply theirs when built. */
function applyMarkdownRideAlong(
	spec: ParagraphContentSpec,
	paragraphs: XmlNode[],
): void {
	if (spec.kind === "markdown-blocks") {
		applyParagraphOptionsToBlocks(paragraphs, spec.paragraphOptions);
	}
}

/** Which of the old paragraph's objects `liftObjectRuns` should leave with the
 *  old content. Markdown CAN express a picture (`![alt](sha256.ext)` — the
 *  read → edit round-trip re-emits the same media part), so when the
 *  replacement carries a `<w:drawing>` the caller re-stated the paragraph's
 *  pictures: the old ones stay behind — discarded with the replaced node, or
 *  wrapped in the tracked `<w:del>` so reject restores them — instead of coming
 *  back as a duplicate `img0`. Only a plain picture qualifies: a chart,
 *  SmartArt, shape, or text box (`<w:drawing>` or not) has no Markdown form, so
 *  it is always lifted and returned. */
function restatedPictures(
	newParagraphs: XmlNode[],
): (object: XmlNode) => boolean {
	const restates = newParagraphs.some((block) =>
		block.findDescendant("w:drawing"),
	);
	return (object) =>
		restates &&
		object.tag === "w:drawing" &&
		object.findDescendant("a:blip") !== undefined &&
		!hasTextBox(object);
}

/** Replacement runs inherit the run formatting COMMON to every visible run of
 *  the replaced paragraph — the intersection of their `<w:rPr>` children (an
 *  8pt Arial form cell whose fill-in span is also underlined contributes the
 *  font/size/color, not the underline). Without this, a whole-paragraph
 *  `--markdown` (or `--text` + run flags) replacement emits bare runs that
 *  fall back to docDefaults — the filled MNDA term cells rendered 11pt Calibri
 *  inside an 8pt Arial form. `<w:highlight>` is never inherited: highlight
 *  marks a placeholder-to-fill, and re-stamping it on the filled value would
 *  recreate the todo marker the edit just resolved. Runs that carry their OWN
 *  rPr (markdown `**bold**`, `--bold`) keep every child they set and gain only
 *  the inherited ones they don't. */
function inheritCommonRunFormatting(
	oldParagraph: XmlNode,
	newParagraphs: XmlNode[],
): void {
	// Enumerate visible runs via the wrapper-aware partition — text lives inside
	// `<w:hyperlink>`/`<w:ins>`/… on redlined or linked paragraphs, and a flat
	// `findChildren("w:r")` would see none of it, silently skipping inheritance
	// (the MNDA font-drift defect resurfacing on exactly those paragraphs).
	const textRuns = partitionParagraphRuns(oldParagraph).runs.filter((run) =>
		run.findChild("w:t"),
	);
	// An EMPTY paragraph (a blank form cell, a template's mandatory cell
	// paragraph) has no runs to take the common formatting from — the formatting
	// Word applies when you type into it lives on the paragraph MARK
	// (`<w:pPr><w:rPr>`). The `--text` diff path already falls back to it
	// (`paragraphMarkRunRpr`); without the same fallback here, `edit --markdown`
	// into the MNDA's empty signature cells wrote bare runs that rendered in the
	// theme font at 12pt instead of the cell's Arial 9pt (batch-3, two runs).
	const firstRpr =
		textRuns[0]?.findChild("w:rPr") ??
		(textRuns.length === 0 ? paragraphMarkRunRpr(oldParagraph) : null);
	if (!firstRpr) return;
	const otherSignatureSets = textRuns
		.slice(1)
		.map(
			(run) =>
				new Set(
					(run.findChild("w:rPr")?.children ?? []).map((child) =>
						XmlNode.serialize([child]),
					),
				),
		);
	const template = firstRpr.clone();
	template.children = template.children.filter((child) => {
		// Never inherit a placeholder-fill (`<w:highlight>`) or tracked-revision
		// (`<w:rPrChange>`) marker — the shared `isInheritableRunProperty` policy.
		if (!isInheritableRunProperty(child)) return false;
		const signature = XmlNode.serialize([child]);
		return otherSignatureSets.every((set) => set.has(signature));
	});
	if (template.children.length === 0) return;
	for (const paragraph of newParagraphs) {
		if (paragraph.tag !== "w:p") continue;
		// A paragraph that brings NEW block structure (a markdown `#` heading on a
		// body line, a list item on a plain paragraph) owns its look through its
		// style — stamping the replaced paragraph's direct rPr onto its runs would
		// defeat that style. Same structure as before (a bullet retyped as a
		// bullet) inherits like plain text. Shared guard with the pPr pass so the
		// two agree on what "plain" means.
		if (bringsNewBlockStructure(paragraph, oldParagraph)) continue;
		// Apply through the same wrapper-aware partition, so a run minted inside a
		// markdown link's `<w:hyperlink>` inherits the font like its siblings.
		for (const run of partitionParagraphRuns(paragraph).runs) {
			const own = run.findChild("w:rPr");
			if (!own) {
				run.children.unshift(template.clone());
				continue;
			}
			for (const child of template.children) inheritRprChild(own, child);
		}
	}
}

/** The paragraph lifted object runs go back onto for an untracked replace: the
 *  first `<w:p>` of the replacement. When the replacement has none (markdown
 *  that built only a table, or nothing at all), a bare paragraph is prepended
 *  to carry them — dropping the box because the new content happened to be a
 *  table is the silent loss this guards against. Only called when there ARE
 *  lifted objects, so an object-free edit's block list is never padded. */
function objectCarrierTarget(newParagraphs: XmlNode[]): XmlNode {
	const first = newParagraphs.find((block) => block.tag === "w:p");
	if (first) return first;
	const carrier = new XmlNode("w:p");
	newParagraphs.unshift(carrier);
	return carrier;
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
