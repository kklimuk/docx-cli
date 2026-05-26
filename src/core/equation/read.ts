import type { XmlNode } from "../parser";
import {
	collectPlainText,
	ELEMENT_HANDLERS,
	isFormatSuppressed,
} from "./handlers";
import { stripCommandWrap } from "./latex";

/** Walk an `<m:oMath>` or `<m:oMathPara>` subtree and reconstruct LaTeX.
 *
 *  The walker handles the OMML constructs observed across three reference
 *  fixtures (Microsoft Word for Mac, Pandoc, LibreOffice) plus the
 *  commonly-used elements from the rest of the OMML schema. For constructs
 *  it doesn't recognize, it falls back **per-subtree** to plaintext
 *  concatenation — an unfamiliar inner construct degrades just that piece,
 *  not the whole equation.
 *
 *  Presentation-only wrappers Word emits (`<m:ctrlPr>`, `<m:*Pr>` carrying
 *  only `<m:ctrlPr>`, `<w:rFonts ascii="Cambria Math">` on math runs) are
 *  ignored; they carry no semantic structure. The list lives below as
 *  `PRESENTATION_TAGS` and matches what we cataloged in the empirical
 *  Word probe (see [tests/fixtures/equations-word.docx](
 *  ../../../tests/fixtures/equations-word.docx)). */
export function ommlToLatex(root: XmlNode): string {
	const inner = root.tag === "m:oMathPara" ? root.findChild("m:oMath") : root;
	if (!inner) return "";
	// Trim disambiguation spaces (added by `escapeAndMap` after every mapped
	// command like `\alpha `) where the next character is unambiguously NOT
	// a letter — at the end of the equation, before close-brackets, before
	// other delimiters. `\xi)` parses identically to `\xi )` so the space
	// is just visual noise that prevents exact round-trips.
	return convertChildren(inner)
		.replace(/[ \t]+$/, "")
		.replace(/(\\[A-Za-z]+) ([)\]])/g, "$1$2");
}

/** Walk every semantic child of `node`, concatenate the LaTeX. Skips
 *  presentation-only `*Pr` wrappers and `<m:ctrlPr>` siblings. Greedy-groups
 *  consecutive math runs that share the same formatting (bold, color,
 *  strike, size) into a single wrapping LaTeX command — `\boldsymbol{x+y}`
 *  instead of `\boldsymbol{x}\boldsymbol{+}\boldsymbol{y}`. Exported so
 *  per-element handlers in [handlers.ts](./handlers.ts) can recurse without
 *  importing from each other (the dispatcher lives here; handlers live there). */
export function convertChildren(node: XmlNode | undefined): string {
	if (!node) return "";
	const children = node.children;
	let out = "";
	let i = 0;
	while (i < children.length) {
		const child = children[i] as XmlNode;
		if (PRESENTATION_TAGS.has(child.tag)) {
			i++;
			continue;
		}
		const fmt = extractRunFormat(child);
		if (fmt) {
			let endIdx = i + 1;
			let chunk = convertElement(child);
			while (endIdx < children.length) {
				const next = children[endIdx] as XmlNode;
				if (PRESENTATION_TAGS.has(next.tag)) {
					endIdx++;
					continue;
				}
				const nextFmt = extractRunFormat(next);
				if (!nextFmt || !runFormatsEqual(fmt, nextFmt)) break;
				chunk += convertElement(next);
				endIdx++;
			}
			// If we're inside a script-bubble-up scope and this chunk's format
			// matches the bubbled wrap, skip the per-chunk wrap — the outer
			// `\textcolor`/`\boldsymbol`/etc. already covers it.
			out += isFormatSuppressed(fmt) ? chunk : wrapRunChunk(chunk, fmt);
			i = endIdx;
			continue;
		}
		out += convertElement(child);
		i++;
	}
	return out;
}

function convertElement(node: XmlNode): string {
	const handler = ELEMENT_HANDLERS[node.tag];
	if (handler) return handler(node);
	// Unknown construct — defensive plaintext fallback so we corrupt just
	// this subtree's structure, not the surrounding LaTeX.
	return collectPlainText(node);
}

// ---------------------------------------------------------------------------
// Run-level formatting detection (color / bold / strike / size).
// ---------------------------------------------------------------------------

type RunFormat = {
	bold?: boolean;
	color?: string;
	sizeHalfPoints?: number;
	strike?: boolean;
};

/** Read formatting properties off a math run's `<m:rPr>`. Returns
 *  `undefined` when the run carries no detectable format we need to wrap
 *  (so the caller doesn't enter the grouping branch). */
function extractRunFormat(node: XmlNode): RunFormat | undefined {
	if (node.tag !== "m:r") return undefined;
	const rPr = node.findChild("m:rPr");
	if (!rPr) return undefined;
	const sty = rPr.findChild("m:sty")?.getAttribute("m:val");
	const bold = sty === "b" || sty === "bi";
	const wRPr = rPr.findChild("w:rPr");
	const color = wRPr?.findChild("w:color")?.getAttribute("w:val");
	const sizeRaw = wRPr?.findChild("w:sz")?.getAttribute("w:val");
	const sizeHalfPoints =
		sizeRaw && Number.isFinite(Number(sizeRaw)) ? Number(sizeRaw) : undefined;
	const strike = wRPr?.findChild("w:strike") !== undefined;
	if (!bold && !color && sizeHalfPoints === undefined && !strike) {
		return undefined;
	}
	return {
		bold: bold || undefined,
		color,
		sizeHalfPoints,
		strike: strike || undefined,
	};
}

function runFormatsEqual(a: RunFormat, b: RunFormat): boolean {
	return (
		Boolean(a.bold) === Boolean(b.bold) &&
		(a.color ?? "") === (b.color ?? "") &&
		a.sizeHalfPoints === b.sizeHalfPoints &&
		Boolean(a.strike) === Boolean(b.strike)
	);
}

/** Apply the LaTeX wrapping commands for the formatting on this chunk.
 *  Order matters for readability: size outermost, then color, strike,
 *  bold innermost so `{\Large \textcolor{red}{\cancel{\boldsymbol{x}}}}`
 *  reads cleanly. */
function wrapRunChunk(body: string, fmt: RunFormat): string {
	// Trim trailing disambiguation space (from `escapeAndMap` mapping the
	// last char to a `\command `) — when it lands right before our closing
	// `}` the space is just noise.
	let out = body.replace(/[ \t]+$/, "");
	if (fmt.bold) {
		// If the body already carries `\boldsymbol{…}` wraps (because bold-
		// italic Greek codepoints decode to `\boldsymbol{\alpha}` via the
		// math-alpha tables), strip the inner wraps — the outer one handles
		// the whole span, so `\boldsymbol{\boldsymbol{\alpha}+\boldsymbol{\beta}}`
		// collapses to `\boldsymbol{\alpha+\beta}`. Uses brace-depth tracking
		// so `\boldsymbol{X^{2}}` (nested braces) unwraps correctly too.
		out = `\\boldsymbol{${stripCommandWrap(out, "boldsymbol")}}`;
	}
	if (fmt.strike) out = `\\cancel{${out}}`;
	if (fmt.color) {
		const named = COLOR_NAME_BY_HEX[fmt.color.toUpperCase()];
		out = named
			? `\\textcolor{${named}}{${out}}`
			: `\\textcolor[HTML]{${fmt.color}}{${out}}`;
	}
	if (fmt.sizeHalfPoints !== undefined) {
		const sizeCommand = SIZE_COMMAND_BY_HALF_POINTS[fmt.sizeHalfPoints];
		out = sizeCommand
			? `{${sizeCommand} ${out}}`
			: `{\\fontsize{${fmt.sizeHalfPoints / 2}pt}{${fmt.sizeHalfPoints / 2}pt}\\selectfont ${out}}`;
	}
	return out;
}

/** Reverse-lookup table: when a hex color matches a known LaTeX/xcolor
 *  named color, emit `\textcolor{red}{…}` instead of the hex form. Covers
 *  the standard `xcolor` set (dvipsnames + a few x11names) — about 30 of
 *  the most common colors in everyday docs. Keys are uppercase 6-hex
 *  digit strings without `#`. */
const COLOR_NAME_BY_HEX: Record<string, string> = {
	"000000": "black",
	FFFFFF: "white",
	FF0000: "red",
	"008000": "green", // \textcolor{green} = #008000 in xcolor (pure 00FF00 is lime)
	"00FF00": "lime",
	"0000FF": "blue",
	"00FFFF": "cyan",
	FF00FF: "magenta",
	FFFF00: "yellow",
	"808080": "gray",
	D3D3D3: "lightgray",
	A9A9A9: "darkgray",
	A52A2A: "brown",
	"808000": "olive",
	FFA500: "orange",
	FFC0CB: "pink",
	"800080": "purple",
	"008080": "teal",
	EE82EE: "violet",
	"4B0082": "indigo",
	DC143C: "crimson",
	FFD700: "gold",
	FA8072: "salmon",
	FF7F50: "coral",
	"40E0D0": "turquoise",
	FFFFF0: "ivory",
	F0E68C: "khaki",
	DDA0DD: "plum",
	D2B48C: "tan",
	"000080": "navy",
	"800000": "maroon",
	C0C0C0: "silver",
};

/** Half-point sizes (Word's `<w:sz>` unit) mapped to LaTeX size commands.
 *  Values chosen to match temml's `\Large` (1.44em ≈ 22*1.44 = 32 hp) /
 *  `\small` (0.9em ≈ 22*0.9 ≈ 20 hp) output relative to a 22 hp baseline. */
const SIZE_COMMAND_BY_HALF_POINTS: Record<number, string> = {
	16: "\\tiny",
	18: "\\scriptsize",
	20: "\\small",
	22: "\\normalsize",
	24: "\\large",
	29: "\\Large",
	32: "\\Large",
	35: "\\LARGE",
	41: "\\huge",
	49: "\\Huge",
};

/** Wrapper elements that carry presentation-only metadata. The walker steps
 *  past them — their semantic siblings (the actual base / exponent /
 *  numerator / etc.) are visited via the handler that owns the parent. */
const PRESENTATION_TAGS = new Set<string>([
	"m:sSupPr",
	"m:sSubPr",
	"m:sSubSupPr",
	"m:sPrePr",
	"m:fPr",
	"m:radPr",
	"m:naryPr",
	"m:limLowPr",
	"m:limUppPr",
	"m:funcPr",
	"m:accPr",
	"m:barPr",
	"m:groupChrPr",
	"m:dPr",
	"m:mPr",
	"m:eqArrPr",
	"m:phantPr",
	"m:boxPr",
	"m:borderBoxPr",
	"m:rPr",
	"m:ctrlPr",
	"w:rPr",
	"m:oMathParaPr",
]);
