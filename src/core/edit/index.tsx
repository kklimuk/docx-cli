import type { Document } from "../ast/document";
import type { BlockRangeReference, BlockReference } from "../ast/document/body";
import type { Run, SectionType } from "../ast/types";
import {
	applyParagraphOptionsInPlace,
	ensureParagraphProperties,
	hasParagraphProperties,
	injectPprChange,
	insertRprChildInOrder,
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
import { replaceSpanInParagraph, type TrackedReplaceOptions } from "../find";
import { readListContext } from "../insert";
import { w } from "../jsx";
import {
	inheritParagraphFormattingIfPlain,
	paragraphOwnsBlockStructure,
} from "../paragraph-inheritance";
import { partitionParagraphRuns, XmlNode } from "../parser";
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
		// `runs` is the explicit, byte-precise surface — the caller states each
		// run's rPr, so inheriting the old paragraph's would silently override
		// their choices (a `{"bold": false}` run can't opt out, since the emitter
		// drops the falsy child and the inherited `<w:b/>` fills the gap).
		if (!opts.noFormatting && spec.kind !== "code" && spec.kind !== "runs") {
			inheritCommonRunFormatting(blockRef.node, newParagraphs);
		}
		continueHostList(this.document, blockRef.node, newParagraphs);
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
		// Pre-built by the CLI via `MarkdownImport.blocks(...)`. The lens does
		// nothing else — the markdown walker has already provisioned styles,
		// allocated list numIds, registered footnote bodies, and minted image
		// rels on the document. We just splice these blocks where the locator
		// pointed.
		return spec.blocks;
	}
	return [<Paragraph runs={spec.runs} {...spec.paragraphOptions} />];
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
	const firstRpr = textRuns[0]?.findChild("w:rPr");
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
		// A paragraph that brought its OWN block structure (a markdown `#` heading,
		// a list item) owns its look through its style — stamping the replaced
		// paragraph's direct rPr onto its runs would defeat that style (an 8pt
		// Arial form cell replaced with `## Heading` would render the heading at
		// 8pt Arial). Skip it via the shared `paragraphOwnsBlockStructure` guard so
		// the two inheritance passes agree on what "plain" means.
		if (paragraphOwnsBlockStructure(paragraph)) continue;
		// Apply through the same wrapper-aware partition, so a run minted inside a
		// markdown link's `<w:hyperlink>` inherits the font like its siblings.
		for (const run of partitionParagraphRuns(paragraph).runs) {
			const own = run.findChild("w:rPr");
			if (!own) {
				run.children.unshift(template.clone());
				continue;
			}
			for (const child of template.children) {
				if (own.findChild(child.tag)) continue;
				insertRprChildInOrder(own, child.clone());
			}
		}
	}
}

/** When a whole-paragraph edit replaces a LIST ITEM with markdown that brings
 *  its own list, the markdown walker has already minted a FRESH numId — a
 *  brand-new list. Dropped mid-list, that restarts numbering at the split
 *  point and desynchronizes everything after it in Word. An agent rewriting
 *  one clause with `--markdown "1. …"` means "replace this item's content,"
 *  not "start a new list" — so re-point the new items at the HOST paragraph's
 *  numId (same list kind only), nesting their levels under the host's. The
 *  minted numId goes unreferenced, which is harmless (see the relationship
 *  invariant: orphans are safe, dangling references are not). */
function continueHostList(
	document: Document,
	oldParagraph: XmlNode,
	newParagraphs: XmlNode[],
): void {
	const host = readListContext(oldParagraph);
	if (!host) return;
	const numbering = document.numbering;
	if (!numbering) return;
	const hostFormat = numbering.getFormat(String(host.numId), host.level);
	// `numFmt="none"` is an unnumbered list — not something a fresh ordered/bullet
	// list should silently continue into (matches `Lists.isOrdered`, which also
	// excludes "none"). Only a real bullet or ordered host is continuable.
	if (!hostFormat || hostFormat === "none") return;
	const hostKind = hostFormat === "bullet" ? "bullet" : "ordered";
	for (const paragraph of newParagraphs) {
		if (paragraph.tag !== "w:p") continue;
		const numPr = paragraph.findChild("w:pPr")?.findChild("w:numPr");
		const numIdNode = numPr?.findChild("w:numId");
		if (!numPr || !numIdNode) continue;
		const fresh = Number(numIdNode.getAttribute("w:val") ?? "0");
		if (!Number.isFinite(fresh) || fresh <= 0 || fresh === host.numId) {
			continue;
		}
		const freshFormat = numbering.getFormat(String(fresh), 0);
		const freshKind = freshFormat === "bullet" ? "bullet" : "ordered";
		// A bullet list replacing an ordered item (or vice versa) is a deliberate
		// kind change — keep the fresh list.
		if (freshKind !== hostKind) continue;
		numIdNode.setAttribute("w:val", String(host.numId));
		const ilvlNode = numPr.findChild("w:ilvl");
		const freshLevel = Number(ilvlNode?.getAttribute("w:val") ?? "0");
		const shifted = Math.min(
			(Number.isFinite(freshLevel) ? freshLevel : 0) + host.level,
			8,
		);
		if (ilvlNode) ilvlNode.setAttribute("w:val", String(shifted));
		else if (shifted > 0) {
			// CT_NumPr order: <w:ilvl> precedes <w:numId>.
			numPr.children.unshift(<w.ilvl w-val={String(shifted)} />);
		}
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
