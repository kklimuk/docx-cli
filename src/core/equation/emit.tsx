import temml from "temml";
import { m } from "../jsx";
import { XmlNode } from "../parser";
import { mathmlToOmml } from "./mathml-to-omml";

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
