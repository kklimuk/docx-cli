import temml from "temml";
import type { Document } from "../ast/document";
import { m } from "../jsx";
import { XmlNode } from "../parser";
import {
	resolveAuthor,
	resolveDate,
	TrackChanges,
	type TrackedMeta,
} from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import { mathmlToOmml } from "./mathml-to-omml";

/** Cross-cutting lens over the document's equations. Constructed at call
 * sites with `new Equations(document)`; holds only a back-reference. Reads from
 * `document.body.equationReferences` (populated by the AST reader) and splices
 * a recompiled `<m:oMath>` / `<m:oMathPara>` back in place. Tracking-aware:
 * when `<w:trackChanges/>` is on, the edit lands as a paired
 * `<w:del>OLD</w:del><w:ins>NEW</w:ins>` next to each other in the same
 * parent. */
export class Equations {
	constructor(private document: Document) {}

	/** List every equation in document order with its current LaTeX, mode,
	 * and the block id of the paragraph containing it. */
	list(): Array<{
		id: string;
		latex: string;
		display: boolean;
		blockId: string;
	}> {
		const out: Array<{
			id: string;
			latex: string;
			display: boolean;
			blockId: string;
		}> = [];
		for (const [id, reference] of this.document.body.equationReferences) {
			out.push({
				id,
				latex: reference.latex,
				display: reference.display,
				blockId: reference.blockId,
			});
		}
		return out;
	}

	/** Recompile an existing equation. `latex` undefined means "keep the
	 * cached LaTeX" (a pure display-mode toggle); `display` undefined means
	 * "keep the cached mode". Throws `EquationNotFoundError` if the id
	 * doesn't resolve, `EquationStaleError` if the cached reference is no
	 * longer in its parent, or `EquationParseError` if temml rejects the
	 * LaTeX. Under track-changes, emits a paired `<w:del>OLD</w:del>
	 * <w:ins>NEW</w:ins>` instead of an in-place splice. */
	edit(
		id: string,
		options: { latex?: string; display?: boolean; author?: string } = {},
	): void {
		const reference = this.document.body.equationReferences.get(id);
		if (!reference) throw new EquationNotFoundError(id);

		const latex = options.latex ?? reference.latex;
		const display = options.display ?? reference.display;
		const omml = latexToOmml(latex, display);

		const index = reference.parent.indexOf(reference.node);
		if (index === -1) throw new EquationStaleError(id);

		if (this.document.isTrackChangesEnabled()) {
			const allocator = new TrackChanges(this.document).createAllocator();
			const author = resolveAuthor(options.author);
			const date = resolveDate();
			const delMeta: TrackedMeta = {
				author,
				date,
				revisionId: allocator.next(),
			};
			const insMeta: TrackedMeta = {
				author,
				date,
				revisionId: allocator.next(),
			};
			reference.parent.splice(
				index,
				1,
				<Del meta={delMeta}>{[reference.node]}</Del>,
				<Ins meta={insMeta}>{[omml]}</Ins>,
			);
		} else {
			reference.parent.splice(index, 1, omml);
		}
	}
}

/** Compile a LaTeX equation string to an OMML subtree ready to splice into
 *  a `<w:p>`. Inline equations are wrapped in `<m:oMath>`; display equations
 *  in `<m:oMathPara><m:oMath>…</m:oMath></m:oMathPara>` (the
 *  display-paragraph wrapper Word, Pandoc, and LibreOffice all emit for
 *  centered block equations).
 *
 *  Pipeline: LaTeX → MathML (via `temml`, MIT-licensed, broad coverage) →
 *  OMML (our own [mathml-to-omml.tsx](./mathml-to-omml.tsx) adapter). We
 *  picked `temml` for the LaTeX-parsing problem since it's the genuinely
 *  hard part — thousands of macros, environments, package commands. For the
 *  MathML → OMML step we control the schema subset we encounter and own
 *  the adapter outright.
 *
 *  Throws `EquationParseError` for malformed LaTeX (temml's error messages
 *  are descriptive). Caller is responsible for catching and producing a
 *  user-friendly CLI error. */
export function latexToOmml(latex: string, display = false): XmlNode {
	const mathml = renderMathml(latex, display);
	const parsed = XmlNode.parse(mathml);
	const mathRoot = XmlNode.findRoot(parsed, "math");
	if (!mathRoot) {
		throw new EquationParseError(
			`temml returned no <math> root for LaTeX: ${latex.slice(0, 80)}`,
			latex,
		);
	}
	const children = mathmlToOmml(mathRoot);
	const oMath = <m.oMath>{children}</m.oMath>;
	return display ? <m.oMathPara>{oMath}</m.oMathPara> : oMath;
}

/** JSX component form for callers building paragraphs via the emitter
 *  pipeline.
 *
 *  @public Staged for the S8 markdown walker — when it compiles `$…$` /
 *  `$$…$$` from a markdown source it'll emit `<Equation latex>` directly,
 *  alongside its other JSX components. The CLI (`insert --equation` /
 *  `edit --at eqN --equation`) calls `latexToOmml` directly instead of
 *  going through this wrapper. */
export function Equation({
	latex,
	display = false,
}: {
	latex: string;
	display?: boolean;
}): XmlNode {
	return latexToOmml(latex, display);
}

/** Error thrown when temml can't parse the LaTeX. The `latex` field holds
 *  the original input for caller-side error messages. */
export class EquationParseError extends Error {
	constructor(
		message: string,
		public latex: string,
	) {
		super(message);
		this.name = "EquationParseError";
	}
}

/** Error thrown by `Equations.edit` when the equation id doesn't resolve in
 *  `document.body.equationReferences`. */
export class EquationNotFoundError extends Error {
	constructor(public id: string) {
		super(`Equation not found: ${id}`);
		this.name = "EquationNotFoundError";
	}
}

/** Error thrown by `Equations.edit` when the cached equation reference is
 *  no longer in its parent (e.g. the paragraph was edited between
 *  read-time and edit-time). */
export class EquationStaleError extends Error {
	constructor(public id: string) {
		super(`Equation ${id} reference is stale (parent does not contain it)`);
		this.name = "EquationStaleError";
	}
}

function renderMathml(latex: string, displayMode: boolean): string {
	try {
		return temml.renderToString(latex, {
			displayMode,
			// `\ref`/`\eqref` are render-time client-side only in temml; we don't
			// emit equation numbering anyway. `strict: false` lets temml accept a
			// few common LaTeX shortcuts (`\cfrac{}{}` etc.) without throwing on
			// non-strict shapes — same default Pandoc uses.
			strict: false,
		});
	} catch (err) {
		const message = err instanceof Error ? err.message : String(err);
		throw new EquationParseError(message, latex);
	}
}
