import type { XmlNode } from "../parser";
import {
	convertGrouped,
	escapeAndMap,
	KNOWN_OPERATORS,
	promoteOperatorName,
} from "./latex";
import { convertChildren } from "./read";

/** Map from OMML element tag to its LaTeX-emitting handler. The walker's
 *  dispatcher consults this; missing tags fall through to a plaintext
 *  fallback (per-subtree degradation). Adding support for a new OMML
 *  element: add a `handleXxx` here, list the corresponding `*Pr` wrapper in
 *  [read.ts](./read.ts) `PRESENTATION_TAGS`, and add a test case. */
export const ELEMENT_HANDLERS: Record<string, (node: XmlNode) => string> = {
	"m:r": handleRun,
	"m:e": handleE,
	"m:sSup": handleSSup,
	"m:sSub": handleSSub,
	"m:sSubSup": handleSSubSup,
	"m:sPre": handleSPre,
	"m:sPreSup": handleSPreSup,
	"m:f": handleFraction,
	"m:rad": handleRadical,
	"m:nary": handleNary,
	"m:limLow": handleLimLow,
	"m:limUpp": handleLimUpp,
	"m:func": handleFunc,
	"m:acc": handleAccent,
	"m:bar": handleBar,
	"m:groupChr": handleGroupChr,
	"m:d": handleDelimiter,
	"m:m": handleMatrix,
	"m:eqArr": handleEqArr,
	"m:phant": handlePhantom,
	"m:box": handleBox,
	"m:borderBox": handleBox,
};

/** Plaintext fallback for unknown constructs — mirrors the historical
 *  `collectMathText` so an unfamiliar OMML element renders as concatenated
 *  text rather than corrupting the surrounding LaTeX. Walks all `<m:t>` /
 *  `<m:delText>` descendants in document order. */
export function collectPlainText(node: XmlNode): string {
	let out = "";
	for (const ch of node.children) {
		if (ch.tag === "m:t" || ch.tag === "m:delText") {
			out += ch.collectText();
			continue;
		}
		out += collectPlainText(ch);
	}
	return out;
}

// ---------------------------------------------------------------------------
// Token / leaf handlers
// ---------------------------------------------------------------------------

/** `<m:r>` math run. Pandoc marks function-operator names (`sin`, `lim`,
 *  `log`, …) with `<m:rPr><m:sty m:val="p"/>` (plain/upright style) — LaTeX
 *  typesets those as `\sin`/`\lim`/`\log` so they're upright rather than
 *  italic. We mirror that:
 *    - upright text matching a known operator → `\name `
 *    - any other upright text containing letters → `\text{…}` (preserves
 *      the upright typesetting on re-write; bare `pH` would render italic)
 *    - everything else → escaped + Unicode-mapped LaTeX */
function handleRun(node: XmlNode): string {
	const rPr = node.findChild("m:rPr");
	const isPlain = rPr?.findChild("m:sty")?.getAttribute("m:val") === "p";
	// A run with no formatting and just a literal space is our SpaceRun shape
	// (writer emits this for `<mspace>` with non-zero width — `\,`, `\quad`,
	// etc.). Round-tripping the bare space loses it through temml on the next
	// pass (temml drops literal whitespace in math mode), so emit `\,` to
	// re-trigger the same mspace.
	if (!rPr) {
		const t = node.findChild("m:t");
		const text = t?.collectText() ?? "";
		// A run with no formatting and only literal spaces is our SpaceRun
		// shape (writer emits 1 space for `\,`, 4 for `\quad`, 8 for `\qquad`,
		// etc.). Map count back to the closest LaTeX spacing command.
		if (text.length > 0 && /^ +$/.test(text)) {
			if (text.length >= 6) return "\\qquad ";
			if (text.length >= 3) return "\\quad ";
			return "\\, ";
		}
	}
	let out = "";
	for (const child of node.children) {
		if (child.tag !== "m:t" && child.tag !== "m:delText") continue;
		const raw = child.collectText();
		if (isPlain && KNOWN_OPERATORS.has(raw)) {
			out += `\\${raw} `;
			continue;
		}
		if (isPlain && /[A-Za-z]/.test(raw)) {
			out += `\\text{${raw.replace(/([\\${}])/g, "\\$1")}}`;
			continue;
		}
		out += escapeAndMap(raw);
	}
	return out;
}

/** `<m:e>` is the generic "element" wrapper inside other constructs; forward
 *  its semantic children. */
function handleE(node: XmlNode): string {
	return convertChildren(node);
}

// ---------------------------------------------------------------------------
// Super/sub/scripts
// ---------------------------------------------------------------------------

function handleSSup(node: XmlNode): string {
	const e = child(node, "m:e");
	const sup = child(node, "m:sup");
	return scriptBubbledOrPlain([e, sup], (parts) => {
		const [base, supStr] = parts;
		return `${convertGrouped(base ?? "")}^${convertGrouped(supStr ?? "")}`;
	});
}

function handleSSub(node: XmlNode): string {
	const e = child(node, "m:e");
	const sub = child(node, "m:sub");
	return scriptBubbledOrPlain([e, sub], (parts) => {
		const [base, subStr] = parts;
		return `${convertGrouped(base ?? "")}_${convertGrouped(subStr ?? "")}`;
	});
}

function handleSSubSup(node: XmlNode): string {
	const e = child(node, "m:e");
	const sub = child(node, "m:sub");
	const sup = child(node, "m:sup");
	return scriptBubbledOrPlain([e, sub, sup], (parts) => {
		const [base, subStr, supStr] = parts;
		return `${convertGrouped(base ?? "")}_${convertGrouped(subStr ?? "")}^${convertGrouped(supStr ?? "")}`;
	});
}

/** Emit a script (sSup/sSub/sSubSup) inside a single shared formatting wrap
 *  when ALL leaf runs across base + sup/sub carry the SAME formatting (color,
 *  bold, strike, size). Without this, `\textcolor{red}{\mathbb{E}[X^2]}` reads
 *  back as `\textcolor{red}{\mathbb{E}[}{\textcolor{red}{X}}^{\textcolor{red}{2}}\textcolor{red}{]}`
 *  — three separate red spans split by the script structure, which trips up
 *  strict KaTeX-based markdown renderers. Bubbling the wrap out collapses the
 *  output to `\textcolor{red}{\mathbb{E}[X^2]}` and lets the convertChildren
 *  grouping merge the bracket pieces with the script. */
function scriptBubbledOrPlain(
	parts: Array<XmlNode | undefined>,
	render: (rendered: string[]) => string,
): string {
	const shared = sharedSubtreeFormat(parts);
	if (!shared) {
		const rendered = parts.map((p) => convertChildren(p));
		return render(rendered);
	}
	// Render each part INSIDE the shared format scope by skipping the wrap
	// (the leaf runs still emit raw LaTeX; convertChildren's grouping would
	// re-wrap them, so we walk children with a flag suppressed via the
	// `suppressFormatWrap` context).
	const rendered = parts.map((p) =>
		withSuppressedFormat(shared, () => convertChildren(p)),
	);
	return wrapWithFormat(render(rendered), shared);
}

type SharedFormat = {
	bold?: boolean;
	color?: string;
	sizeHalfPoints?: number;
	strike?: boolean;
};

/** Walk every leaf `<m:r>` under `parts`. If they all carry the same
 *  formatting (and at least one carries SOMETHING), return that format —
 *  otherwise undefined. Empty subtrees count as "no format constraint",
 *  so a fully-formatted base + plain sub still bubbles if the base is the
 *  only thing carrying format. */
function sharedSubtreeFormat(
	parts: Array<XmlNode | undefined>,
): SharedFormat | undefined {
	let agreed: SharedFormat | undefined;
	let sawAny = false;
	for (const part of parts) {
		if (!part) continue;
		for (const run of collectRuns(part)) {
			const fmt = extractRunFormatFromR(run);
			if (!fmt) return undefined; // unformatted run breaks the shared wrap
			sawAny = true;
			if (agreed === undefined) {
				agreed = fmt;
			} else if (!sameFormat(agreed, fmt)) {
				return undefined;
			}
		}
	}
	return sawAny ? agreed : undefined;
}

function collectRuns(node: XmlNode): XmlNode[] {
	const out: XmlNode[] = [];
	const walk = (n: XmlNode): void => {
		if (n.tag === "m:r") {
			out.push(n);
			return;
		}
		for (const c of n.children) walk(c);
	};
	walk(node);
	return out;
}

function extractRunFormatFromR(run: XmlNode): SharedFormat | undefined {
	const rPr = run.findChild("m:rPr");
	if (!rPr) return undefined;
	const sty = rPr.findChild("m:sty")?.getAttribute("m:val");
	const bold = sty === "b" || sty === "bi";
	const wRPr = rPr.findChild("w:rPr");
	const color = wRPr?.findChild("w:color")?.getAttribute("w:val");
	const sizeRaw = wRPr?.findChild("w:sz")?.getAttribute("w:val");
	const sizeHalfPoints =
		sizeRaw && Number.isFinite(Number(sizeRaw)) ? Number(sizeRaw) : undefined;
	const strike = wRPr?.findChild("w:strike") !== undefined;
	if (!bold && !color && sizeHalfPoints === undefined && !strike)
		return undefined;
	return {
		bold: bold || undefined,
		color,
		sizeHalfPoints,
		strike: strike || undefined,
	};
}

function sameFormat(a: SharedFormat, b: SharedFormat): boolean {
	return (
		Boolean(a.bold) === Boolean(b.bold) &&
		(a.color ?? "") === (b.color ?? "") &&
		a.sizeHalfPoints === b.sizeHalfPoints &&
		Boolean(a.strike) === Boolean(b.strike)
	);
}

/** Wrap `body` in the LaTeX commands that correspond to the shared format.
 *  Mirrors `wrapRunChunk` in read.ts but with body trimming. */
function wrapWithFormat(body: string, fmt: SharedFormat): string {
	let out = body;
	if (fmt.bold) out = `\\boldsymbol{${out}}`;
	if (fmt.strike) out = `\\cancel{${out}}`;
	if (fmt.color) {
		// Prefer named-color form if recognized — keeps the output readable.
		const named = SHARED_COLOR_NAMES[fmt.color.toUpperCase()];
		out = named
			? `\\textcolor{${named}}{${out}}`
			: `\\textcolor[HTML]{${fmt.color}}{${out}}`;
	}
	return out;
}

// Small subset of `read.ts`'s named-color table — sufficient for the common
// cases that show up in scripts. The full table is in `read.ts` and used
// when the bubble-up doesn't fire (per-run wrapping path).
const SHARED_COLOR_NAMES: Record<string, string> = {
	"000000": "black",
	FFFFFF: "white",
	FF0000: "red",
	"0000FF": "blue",
	"008000": "green",
	"00FF00": "lime",
	FFA500: "orange",
	"800080": "purple",
	"808080": "gray",
	FFFF00: "yellow",
	"00FFFF": "cyan",
	FF00FF: "magenta",
};

let suppressedFormat: SharedFormat | undefined;

function withSuppressedFormat<T>(fmt: SharedFormat, fn: () => T): T {
	const previous = suppressedFormat;
	suppressedFormat = fmt;
	try {
		return fn();
	} finally {
		suppressedFormat = previous;
	}
}

/** Exported for the convertChildren grouping in [read.ts](./read.ts) — when
 *  it picks up a run, it checks whether the run's format matches what the
 *  caller already suppressed (bubble-up scope) and skips re-wrapping in that
 *  case. */
export function isFormatSuppressed(fmt: SharedFormat): boolean {
	if (!suppressedFormat) return false;
	return sameFormat(suppressedFormat, fmt);
}

/** Pre-subscript (common in chemistry / isotopes: `${}^{14}_{6}C$`). */
function handleSPre(node: XmlNode): string {
	const sub = convertGrouped(convertChildren(child(node, "m:sub")));
	const e = convertChildren(child(node, "m:e"));
	return `{}_${sub}${e}`;
}

function handleSPreSup(node: XmlNode): string {
	const sub = convertGrouped(convertChildren(child(node, "m:sub")));
	const sup = convertGrouped(convertChildren(child(node, "m:sup")));
	const e = convertChildren(child(node, "m:e"));
	return `{}_${sub}^${sup}${e}`;
}

// ---------------------------------------------------------------------------
// Fractions and roots
// ---------------------------------------------------------------------------

function handleFraction(node: XmlNode): string {
	const num = convertChildren(child(node, "m:num")).trimEnd();
	const den = convertChildren(child(node, "m:den")).trimEnd();
	const fType = node
		.findChild("m:fPr")
		?.findChild("m:type")
		?.getAttribute("m:val");
	// `<m:type m:val="noBar"/>` is OMML's shape for `\binom{n}{k}` — Word and
	// our own emitter both wrap this in `<m:d>` with parentheses, so the
	// enclosing `\left(`/`\right)` comes naturally from `handleDelimiter`.
	if (fType === "noBar") return `\\binom{${num}}{${den}}`;
	if (fType === "skw") return `{}^{${num}}/_{${den}}`; // skewed (\nicefrac-ish)
	return `\\frac{${num}}{${den}}`;
}

/** `<m:rad>` with optional `<m:degHide m:val="1"/>` → `\sqrt{x}`; else
 *  `\sqrt[n]{x}`. An empty `<m:deg/>` is treated as degHide. */
function handleRadical(node: XmlNode): string {
	const radPr = node.findChild("m:radPr");
	const degHide = radPr?.findChild("m:degHide")?.getAttribute("m:val") === "1";
	const body = convertChildren(child(node, "m:e")).trimEnd();
	if (degHide) return `\\sqrt{${body}}`;
	const degNode = node.findChild("m:deg");
	const deg = degNode ? convertChildren(degNode).trimEnd() : "";
	if (!deg) return `\\sqrt{${body}}`;
	return `\\sqrt[${deg}]{${body}}`;
}

// ---------------------------------------------------------------------------
// N-ary operators and limits
// ---------------------------------------------------------------------------

function handleNary(node: XmlNode): string {
	const naryPr = node.findChild("m:naryPr");
	const chr = naryPr?.findChild("m:chr")?.getAttribute("m:val") ?? "∫";
	const cmd = NARY_OP_BY_CHR[chr] ?? `\\${chr}`;
	const subHide = naryPr?.findChild("m:subHide")?.getAttribute("m:val") === "1";
	const supHide = naryPr?.findChild("m:supHide")?.getAttribute("m:val") === "1";
	const subNode = subHide ? undefined : child(node, "m:sub");
	const supNode = supHide ? undefined : child(node, "m:sup");
	const sub = subNode ? convertChildren(subNode).trimEnd() : "";
	const sup = supNode ? convertChildren(supNode).trimEnd() : "";
	const body = convertChildren(child(node, "m:e"));
	let out = cmd;
	if (sub) out += `_{${sub}}`;
	if (sup) out += `^{${sup}}`;
	if (body) out += ` ${body.trimEnd()}`;
	return out;
}

/** `<m:limLow>` / `<m:limUpp>` cover two shapes:
 *   - Function with a limit: `\lim_{x→0}`, `\sup_{i}`, `\inf_{x∈X}` — the base
 *     is a function operator name (lim / sup / inf / max / …). We emit
 *     `\name_{lim}` / `\name^{lim}` so the function-operator promotion picks up.
 *   - Arbitrary `\underset{X}{base}` / `\overset{X}{base}` — when the base
 *     isn't a function name, the user wants the label CENTERED above/below
 *     the base, not raised as a sub/sup. AMSmath provides `\underset` /
 *     `\overset` for this; we route to them.
 *
 *  Detection: if `promoteOperatorName` changes the base (turns `lim` into
 *  `\lim`), we know the base is an operator and the `_{}`/`^{}` form is
 *  the right rendering. Otherwise we use `\underset` / `\overset`. */
function handleLimLow(node: XmlNode): string {
	const rawBase = convertChildren(child(node, "m:e")).trimEnd();
	const promoted = promoteOperatorName(rawBase);
	const lim = convertChildren(child(node, "m:lim")).trimEnd();
	if (promoted !== rawBase || usesSubSupForm(rawBase)) {
		return `${promoted}_{${lim}}`;
	}
	return `\\underset{${lim}}{${rawBase}}`;
}

function handleLimUpp(node: XmlNode): string {
	const rawBase = convertChildren(child(node, "m:e")).trimEnd();
	const promoted = promoteOperatorName(rawBase);
	const lim = convertChildren(child(node, "m:lim")).trimEnd();
	if (promoted !== rawBase || usesSubSupForm(rawBase)) {
		return `${promoted}^{${lim}}`;
	}
	return `\\overset{${lim}}{${rawBase}}`;
}

/** True when the base of a `<m:limLow>` / `<m:limUpp>` is a command that
 *  conventionally takes a sub/sup label rather than an `\underset`/`\overset`
 *  wrap — function operators (`\lim`, `\sup`, …) and the labeled-brace
 *  family (`\underbrace`, `\overbrace`, `\overparen`, …). Anything else is
 *  arbitrary content and gets the `\underset`/`\overset` form. */
function usesSubSupForm(text: string): boolean {
	const opMatch = text.match(/^\\([A-Za-z]+)\s?$/);
	if (opMatch && KNOWN_OPERATORS.has(opMatch[1] ?? "")) return true;
	const braceMatch = text.match(/^\\([A-Za-z]+)\{/);
	return !!braceMatch && BRACE_COMMANDS.has(braceMatch[1] ?? "");
}

/** Brace decoration commands that take a sub/sup label (`\underbrace{X}_{Y}`,
 *  `\overbrace{X}^{Y}`, …). Listed here so `handleLimLow`/`Upp` know to
 *  preserve the conventional LaTeX form rather than wrapping in
 *  `\underset`/`\overset`. */
const BRACE_COMMANDS = new Set([
	"underbrace",
	"overbrace",
	"underparen",
	"overparen",
	"undergroup",
	"overgroup",
]);

/** `<m:func><m:fName>lim</m:fName><m:e>{body}</m:e></m:func>` — function
 *  application. When `<m:fName>` resolves to a known operator (lim/sin/cos/
 *  log/etc.), we promote it to `\name` so it typesets as a math operator
 *  rather than italic letters. Trim trailing whitespace on fName because
 *  operator-name promotion already adds a `\name ` trailing space, and the
 *  template literal adds another — collapsing those keeps `\ln 2` from
 *  becoming `\ln  2` (double space). */
function handleFunc(node: XmlNode): string {
	const fNameNode = child(node, "m:fName");
	const body = convertChildren(child(node, "m:e")).trimEnd();
	let fName = fNameNode ? convertChildren(fNameNode).trimEnd() : "";
	fName = promoteOperatorName(fName);
	return body ? `${fName} ${body}` : fName;
}

// ---------------------------------------------------------------------------
// Decorations
// ---------------------------------------------------------------------------

/** `<m:acc><m:accPr><m:chr m:val="̂"/></m:accPr><m:e>x</m:e></m:acc>` →
 *  `\hat{x}` etc. The combining-diacritic char identifies the accent. */
function handleAccent(node: XmlNode): string {
	const chr =
		node.findChild("m:accPr")?.findChild("m:chr")?.getAttribute("m:val") ?? "";
	const body = convertChildren(child(node, "m:e")).trimEnd();
	const cmd = ACCENT_CMD_BY_CHR[chr] ?? "\\hat";
	return `${cmd}{${body}}`;
}

function handleBar(node: XmlNode): string {
	const pos = node
		.findChild("m:barPr")
		?.findChild("m:pos")
		?.getAttribute("m:val");
	const body = convertChildren(child(node, "m:e")).trimEnd();
	return pos === "bot" ? `\\underline{${body}}` : `\\overline{${body}}`;
}

/** `<m:groupChr>` for over/underbrace and similar grouping decorations. */
function handleGroupChr(node: XmlNode): string {
	const props = node.findChild("m:groupChrPr");
	const chr = props?.findChild("m:chr")?.getAttribute("m:val") ?? "";
	const pos = props?.findChild("m:pos")?.getAttribute("m:val") ?? "top";
	const body = convertChildren(child(node, "m:e")).trimEnd();
	const cmd = GROUP_CHR_CMD_BY_CHR[chr];
	if (cmd) return `${cmd}{${body}}`;
	return pos === "bot" ? `\\underbrace{${body}}` : `\\overbrace{${body}}`;
}

// ---------------------------------------------------------------------------
// Delimiters and tabular structures
// ---------------------------------------------------------------------------

/** `<m:d>` delimited expression. May hold multiple `<m:e>` children separated
 *  by the `<m:sepChr>` attribute (defaults to `,`). We always emit
 *  `\left…\right…` for stretchy delimiters — Pandoc reads either bare or
 *  `\left/\right` forms.
 *
 *  Special shorthands when the delimiter wraps a single semantic child:
 *  - `(<m:f noBar>)` → `\binom{n}{k}` (OMML's representation of `\binom`
 *    keeps the parens in `<m:d>` because Word's renderer relies on them)
 *  - `(<m:m>)` / `[<m:m>]` / `|<m:m>|` / `‖<m:m>‖` → `\begin{pmatrix}…\end{pmatrix}`
 *    and friends (matrix shorthand environments). */
function handleDelimiter(node: XmlNode): string {
	const dPr = node.findChild("m:dPr");
	const beg = dPr?.findChild("m:begChr")?.getAttribute("m:val") ?? "(";
	const end = dPr?.findChild("m:endChr")?.getAttribute("m:val") ?? ")";
	const sep = dPr?.findChild("m:sepChr")?.getAttribute("m:val") ?? ",";
	const elements = node.findChildren("m:e");
	const single = elements.length === 1 ? elements[0] : undefined;
	const sole = single ? nonPresentationChildren(single) : [];
	const onlyChildTag = sole.length === 1 ? sole[0]?.tag : undefined;

	if (onlyChildTag === "m:f" && beg === "(" && end === ")") {
		const fraction = sole[0] as XmlNode;
		const fType = fraction
			.findChild("m:fPr")
			?.findChild("m:type")
			?.getAttribute("m:val");
		if (fType === "noBar") {
			const num = convertChildren(fraction.findChild("m:num")).trimEnd();
			const den = convertChildren(fraction.findChild("m:den")).trimEnd();
			return `\\binom{${num}}{${den}}`;
		}
	}

	if (onlyChildTag === "m:m") {
		const shorthand = MATRIX_SHORTHAND_BY_PAIR.get(`${beg}${end}`);
		if (shorthand) {
			return renderMatrixRows(sole[0] as XmlNode, shorthand);
		}
	}

	const body = elements.map((e) => convertChildren(e)).join(sep);
	return `\\left${normalizeDelimiter(beg)}${body}\\right${normalizeDelimiter(end)}`;
}

/** Matrix shorthand: `(…)` → `\begin{pmatrix}…\end{pmatrix}` and friends.
 *  The shorthand collapses one `<m:d>` + `<m:m>` pair we'd otherwise emit
 *  as `\left(\begin{matrix}…\end{matrix}\right)`. */
const MATRIX_SHORTHAND_BY_PAIR = new Map<string, string>([
	["()", "pmatrix"],
	["[]", "bmatrix"],
	["{}", "Bmatrix"],
	["||", "vmatrix"],
	["‖‖", "Vmatrix"],
]);

function renderMatrixRows(matrix: XmlNode, environment: string): string {
	const rows = matrix.findChildren("m:mr").map((row) => {
		const cells = row.findChildren("m:e").map((e) => convertChildren(e));
		return cells.join(" & ");
	});
	return `\\begin{${environment}}${rows.join(" \\\\ ")}\\end{${environment}}`;
}

/** A node's children minus the presentation-only metadata wrappers — for
 *  detecting "this `<m:e>` holds exactly one structural child" without
 *  tripping over `<m:ctrlPr>` etc. */
function nonPresentationChildren(node: XmlNode): XmlNode[] {
	return node.children.filter((c) => !PRESENTATION_TAGS_FOR_PROBE.has(c.tag));
}

const PRESENTATION_TAGS_FOR_PROBE = new Set([
	"m:ctrlPr",
	"m:rPr",
	"w:rPr",
	"#text",
]);

/** Matrix `<m:m>` containing `<m:mr>` rows of `<m:e>` cells. The surrounding
 *  delimiter (if any) is on a parent `<m:d>`; we emit a bare `\begin{matrix}`
 *  and let the parent's `\left(`/`\right)` wrap it for pmatrix-like output. */
function handleMatrix(node: XmlNode): string {
	const rows = node.findChildren("m:mr").map((row) => {
		const cells = row.findChildren("m:e").map((e) => convertChildren(e));
		return cells.join(" & ");
	});
	return `\\begin{matrix}${rows.join(" \\\\ ")}\\end{matrix}`;
}

/** `<m:eqArr>` aligned equation array — Pandoc uses it for `\begin{aligned}`,
 *  `\begin{cases}`, and similar multi-line layouts. We render as `aligned`;
 *  consumers (Pandoc) reformat to whichever environment they prefer. */
function handleEqArr(node: XmlNode): string {
	const rows = node.findChildren("m:e").map((e) => convertChildren(e));
	return `\\begin{aligned}${rows.join(" \\\\ ")}\\end{aligned}`;
}

// ---------------------------------------------------------------------------
// Miscellaneous
// ---------------------------------------------------------------------------

function handlePhantom(_node: XmlNode): string {
	// `<m:phant>` is invisible-spacing. No reliable LaTeX equivalent that
	// round-trips; dropping is benign in rendered output.
	return "";
}

function handleBox(node: XmlNode): string {
	return `\\boxed{${convertChildren(node).trimEnd()}}`;
}

// ---------------------------------------------------------------------------
// Tables shared by handlers
// ---------------------------------------------------------------------------

/** N-ary operator chars → LaTeX command. Word and Pandoc emit the Unicode
 *  operator in `<m:naryPr><m:chr m:val>`. */
const NARY_OP_BY_CHR: Record<string, string> = {
	"∑": "\\sum",
	"∏": "\\prod",
	"∐": "\\coprod",
	"∫": "\\int",
	"∬": "\\iint",
	"∭": "\\iiint",
	"∮": "\\oint",
	"∯": "\\oiint",
	"∰": "\\oiiint",
	"⋂": "\\bigcap",
	"⋃": "\\bigcup",
	"⨁": "\\bigoplus",
	"⨂": "\\bigotimes",
	"⨅": "\\bigsqcap",
	"⨆": "\\bigsqcup",
};

/** Combining-diacritic chars (U+0300–U+036F) → LaTeX accent command. */
/** Producers spread the same logical accent across two Unicode encodings:
 *  combining-diacritic codepoints (U+0300–U+036F, what Word emits) and
 *  spacing/overscript codepoints (U+00AF, U+203E, etc., what Pandoc tends to
 *  emit). We register both forms for each accent so either producer's output
 *  round-trips. */
const ACCENT_CMD_BY_CHR: Record<string, string> = {
	// Combining marks (U+0300-U+036F)
	"̂": "\\hat",
	"̃": "\\tilde",
	"̄": "\\bar",
	"̇": "\\dot",
	"̈": "\\ddot",
	"⃛": "\\dddot",
	"̌": "\\check",
	"̆": "\\breve",
	"̀": "\\grave",
	"́": "\\acute",
	"⃗": "\\vec",
	"⃡": "\\overleftrightarrow",
	// Spacing / overscript variants (Pandoc's preference)
	"^": "\\hat",
	"~": "\\tilde",
	"¯": "\\bar", // U+00AF MACRON
	"‾": "\\bar", // U+203E OVERLINE
	"˙": "\\dot",
	"¨": "\\ddot",
	ˇ: "\\check",
	"˘": "\\breve",
	"`": "\\grave",
	"´": "\\acute",
	"→": "\\vec",
};

/** Group-character chars (overbrace/underbrace/etc.) → LaTeX command. The
 *  stretchy / wide-accent chars (`^`, `~`, `→`, `‾`, `¯`) live here too —
 *  our writer emits `<m:groupChr>` instead of `<m:acc>` when temml marks
 *  the over-char as stretchy, so reading those back as `\widehat` /
 *  `\widetilde` / `\overrightarrow` / `\overline` (rather than the
 *  non-stretchy `\hat` / `\tilde` / `\vec` / `\bar`) preserves the
 *  intent. */
const GROUP_CHR_CMD_BY_CHR: Record<string, string> = {
	"⏞": "\\overbrace",
	"⏟": "\\underbrace",
	"⏜": "\\overparen",
	"⏝": "\\underparen",
	"⏠": "\\overgroup",
	"⏡": "\\undergroup",
	"⃗": "\\overrightarrow",
	"⃖": "\\overleftarrow",
	"⃐": "\\overarc",
	// Wide / stretchy accent forms (the writer routes here when stretchy=true).
	"^": "\\widehat",
	"~": "\\widetilde",
	"→": "\\overrightarrow",
	"‾": "\\overline",
	"¯": "\\overline",
};

/** Map an OMML delimiter character to a LaTeX-friendly form. LaTeX accepts
 *  literal `(` `)` `[` `]` `\{` `\}` `|`; for invisible (empty val) we emit
 *  `.` (LaTeX's invisible delimiter, e.g. `\left.` for one-sided braces). */
function normalizeDelimiter(chr: string): string {
	if (!chr) return ".";
	if (chr === "{" || chr === "}") return `\\${chr}`;
	if (chr === "‖") return "\\|";
	return chr;
}

function child(node: XmlNode, tag: string): XmlNode | undefined {
	return node.findChild(tag);
}
