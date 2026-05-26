import { m, w } from "../jsx";
import type { XmlNode } from "../parser";
import { isMathTextNoise } from "./latex";

/** Convert a MathML root (`<math>`) to an array of OMML siblings — one per
 *  semantic top-level child. The caller wraps the result in `<m:oMath>` or
 *  `<m:oMathPara>` (the JSX `<Equation>` component in [emit.tsx](./emit.tsx)
 *  does this).
 *
 *  Temml emits MathML close enough to the Core profile that we get clean
 *  output for the constructs we cover. Patterns we recognize specifically:
 *
 *   - `<msubsup><mo>∑</mo>…</msubsup> body` → `<NaryOperator>` (the
 *     big-operator with absorbed body — temml emits the operator and its
 *     scripts in `<msubsup>`, then the body as the next sibling)
 *   - `<mrow><mo fence>(</mo> X <mo fence>)</mo></mrow>` → `<Delimited>`
 *     (temml's pmatrix / bmatrix / \left(…\right) shape)
 *   - `<mover>x <mo>^</mo></mover>` → `<OverAccent>` when the over-char is
 *     in the accent set and non-stretchy; stretchy → `<GroupCharacter>`
 *     (which the reader maps to `\widehat`/`\overrightarrow`/etc.)
 *   - `<munder><munder>X<mo>⏟</mo></munder>Y</munder>` (labeled brace) →
 *     `<m:limLow>` wrapping `<GroupCharacter>` so `\underbrace{X}_{Y}`
 *     survives
 *   - `<mfrac linethickness="0px">` → `<Fraction noBar>` for `\binom`
 *   - `<mtable class="tml-jot">` → `<EquationArray>` (`<m:eqArr>`) so
 *     `\begin{aligned}` survives; ordinary `<mtable>` → `<Matrix>`
 *   - `<mtable class="tml-tageqn">` (3-cell tag layout) → equation OMML
 *     followed by a trailing tag text run for `\tag{N}` */
export function mathmlToOmml(root: XmlNode): XmlNode[] {
	currentFormat = {}; // reset between top-level calls
	return convertSiblings(root.children);
}

// ---------------------------------------------------------------------------
// Run-formatting context — temml expresses `\boldsymbol{…}`, `\textcolor{…}{…}`,
// `\Large …`, `\cancel{…}` etc. as parent elements (mrow / mstyle / menclose)
// with styling attributes that should apply to every descendant text run. We
// thread the active formatting through a module-level context that descender
// helpers push to (and pop from) as they enter / leave a styled scope. The
// context is consulted by `MathRun` constructions to emit the right OMML
// run-property children.
// ---------------------------------------------------------------------------

type FormatContext = {
	bold?: boolean;
	color?: string;
	sizeHalfPoints?: number;
	strike?: boolean;
};

let currentFormat: FormatContext = {};

/** Run `fn` with `extra` merged into the active format. Restores prior
 *  state on exit (even if `fn` throws). */
function withFormat<T>(extra: FormatContext, fn: () => T): T {
	const previous = currentFormat;
	currentFormat = { ...previous, ...extra };
	try {
		return fn();
	} finally {
		currentFormat = previous;
	}
}

/** Merge ambient `currentFormat` with formatting picked up from `node`'s
 *  own attributes (`style="color:…;font-weight:bold;…"`, `mathvariant`,
 *  `mathsize`). Pass the result to a `MathRun` so the OMML run carries
 *  the right `<m:sty>` / `<w:rPr>` children. */
function runFormat(node: XmlNode): FormatContext {
	return { ...currentFormat, ...readNodeFormat(node) };
}

/** Parse formatting attributes off a MathML element. Returns only the
 *  fields temml's emits; leaves others undefined for inheritance. */
function readNodeFormat(node: XmlNode): FormatContext {
	const out: FormatContext = {};
	const style = node.getAttribute("style");
	if (style) {
		const color = /color\s*:\s*(#?[A-Fa-f0-9]{3,8})/.exec(style);
		if (color?.[1]) out.color = normalizeColorHex(color[1]);
		if (/font-weight\s*:\s*(?:bold|[6-9]\d\d)/.test(style)) out.bold = true;
		const ts = /text-decoration[^;]*line-through/.exec(style);
		if (ts) out.strike = true;
	}
	const variant = node.getAttribute("mathvariant");
	if (variant && /bold/.test(variant)) out.bold = true;
	const mathsize = node.getAttribute("mathsize");
	if (mathsize) {
		const ems = Number.parseFloat(mathsize);
		if (Number.isFinite(ems)) {
			// Word's `<w:sz>` is in HALF-points; default math is ~11pt = 22 half-pt.
			out.sizeHalfPoints = Math.round(ems * 22);
		}
	}
	return out;
}

/** Strip zero-width and variation-selector characters from MathML text
 *  before it lands in an OMML `<m:t>`. Word renders variation selectors
 *  (`⊕` + VS15 from temml's `\oplus`) as missing-glyph boxes — silently
 *  stripping them keeps the rendered math clean. */
function cleanMathText(text: string): string {
	let out = "";
	for (const ch of text) if (!isMathTextNoise(ch)) out += ch;
	return out;
}

/** Normalize CSS color values to a 6-hex-digit uppercase string (no `#`),
 *  which is what `<w:color w:val>` expects. Accepts `#ffaabb`, `ffaabb`,
 *  `#fab` (short), and a few common CSS names. */
function normalizeColorHex(raw: string): string {
	const cleaned = raw.replace(/^#/, "").toUpperCase();
	if (/^[0-9A-F]{6}$/.test(cleaned)) return cleaned;
	if (/^[0-9A-F]{3}$/.test(cleaned)) {
		return cleaned
			.split("")
			.map((c) => c + c)
			.join("");
	}
	return cleaned;
}

// ---------------------------------------------------------------------------
// Components — pure `props → XmlNode` OMML builders. Each emits a single
// OMML element from already-converted child OMML.
// ---------------------------------------------------------------------------

function MathRun({
	text,
	upright,
	bold,
	color,
	sizeHalfPoints,
	strike,
}: {
	text: string;
	upright: boolean;
	bold?: boolean;
	color?: string;
	sizeHalfPoints?: number;
	strike?: boolean;
}): XmlNode {
	// Math-style: combine upright/bold for `<m:sty>`. Per ECMA-376 §22.1.3:
	// `p`=plain (upright), `b`=bold (italic by default in math), `bi`=bold-italic.
	// Italic is the default for letters; we don't need `i` explicitly.
	let styVal: "p" | "b" | "bi" | undefined;
	if (bold && upright) styVal = "bi";
	else if (bold) styVal = "b";
	else if (upright) styVal = "p";

	const hasWRPr =
		color !== undefined || sizeHalfPoints !== undefined || strike === true;
	if (styVal === undefined && !hasWRPr) {
		return (
			<m.r>
				<m.t>{text}</m.t>
			</m.r>
		);
	}
	return (
		<m.r>
			<m.rPr>
				{styVal && <m.sty m-val={styVal} />}
				{hasWRPr && (
					<w.rPr>
						{color !== undefined && <w.color w-val={color} />}
						{sizeHalfPoints !== undefined && (
							<w.sz w-val={String(sizeHalfPoints)} />
						)}
						{strike && <w.strike />}
					</w.rPr>
				)}
			</m.rPr>
			<m.t>{text}</m.t>
		</m.r>
	);
}

/** A run carrying one or more literal spaces — `xml:space="preserve"` keeps
 *  the whitespace through the fast-xml-builder round-trip. Used for
 *  `<mspace>` and other thin-space MathML constructs that temml emits
 *  between tokens. The `count` defaults to 1 (`\,`-equivalent thin space);
 *  higher counts approximate `\quad` (4 spaces) / `\qquad` (8 spaces). */
function SpaceRun({ count = 1 }: { count?: number }): XmlNode {
	return (
		<m.r>
			<m.t {...{ "xml:space": "preserve" }}>{" ".repeat(count)}</m.t>
		</m.r>
	);
}

function Superscript({
	base,
	sup,
}: {
	base: XmlNode[];
	sup: XmlNode[];
}): XmlNode {
	return (
		<m.sSup>
			<m.e>{base}</m.e>
			<m.sup>{sup}</m.sup>
		</m.sSup>
	);
}

function Subscript({
	base,
	sub,
}: {
	base: XmlNode[];
	sub: XmlNode[];
}): XmlNode {
	return (
		<m.sSub>
			<m.e>{base}</m.e>
			<m.sub>{sub}</m.sub>
		</m.sSub>
	);
}

function SubSuperscript({
	base,
	sub,
	sup,
}: {
	base: XmlNode[];
	sub: XmlNode[];
	sup: XmlNode[];
}): XmlNode {
	return (
		<m.sSubSup>
			<m.e>{base}</m.e>
			<m.sub>{sub}</m.sub>
			<m.sup>{sup}</m.sup>
		</m.sSubSup>
	);
}

/** `<m:f>` fraction — bar by default. `noBar` flips on `<m:type m:val="noBar"/>`
 *  inside `<m:fPr>`, the OMML shape for `\binom{n}{k}`. Wrapped in `<m:d>`
 *  with parentheses by the caller (the fence-detection dispatcher does this
 *  naturally since temml emits the `\binom` parens as `<mo fence>` siblings). */
function Fraction({
	num,
	den,
	noBar = false,
}: {
	num: XmlNode[];
	den: XmlNode[];
	noBar?: boolean;
}): XmlNode {
	return (
		<m.f>
			{noBar && (
				<m.fPr>
					<m.type m-val="noBar" />
				</m.fPr>
			)}
			<m.num>{num}</m.num>
			<m.den>{den}</m.den>
		</m.f>
	);
}

/** Square root — `<m:rad>` with `<m:degHide/>` so renderers draw a plain √.
 *  The reader recognizes this shape and emits `\sqrt{…}`. */
function SquareRoot({ body }: { body: XmlNode[] }): XmlNode {
	return (
		<m.rad>
			<m.radPr>
				<m.degHide m-val="1" />
			</m.radPr>
			<m.deg />
			<m.e>{body}</m.e>
		</m.rad>
	);
}

function NthRoot({
	degree,
	body,
}: {
	degree: XmlNode[];
	body: XmlNode[];
}): XmlNode {
	return (
		<m.rad>
			<m.deg>{degree}</m.deg>
			<m.e>{body}</m.e>
		</m.rad>
	);
}

function OverAccent({ chr, body }: { chr: string; body: XmlNode[] }): XmlNode {
	// Translate spacing/standalone Unicode chars temml uses (`→`, `^`, `~`)
	// to their COMBINING accent form (U+20D7, U+0302, U+0303) — Word renders
	// the combining diacritics as proper accents over the base; the spacing
	// forms render as tiny floating glyphs alongside the letter.
	const accentChr = ACCENT_CHR_TO_COMBINING[chr] ?? chr;
	return (
		<m.acc>
			<m.accPr>
				<m.chr m-val={accentChr} />
			</m.accPr>
			<m.e>{body}</m.e>
		</m.acc>
	);
}

/** Spacing → combining-diacritic translations for accent characters. OMML's
 *  `<m:chr>` for `<m:acc>` is expected to be a combining mark (U+0300–U+036F
 *  range) so Word's renderer composes it over the base letter. The keys are
 *  temml's MathML output chars; the values are the OMML-canonical forms. */
const ACCENT_CHR_TO_COMBINING: Record<string, string> = {
	"→": "⃗", // U+2192 → U+20D7  (\vec)
	"^": "̂", // U+005E → U+0302  (\hat)
	"~": "̃", // U+007E → U+0303  (\tilde)
	"¯": "̄", // U+00AF → U+0304  (\bar)
	"‾": "̄", // U+203E → U+0304  (\bar, alt encoding)
	"˙": "̇", // U+02D9 → U+0307  (\dot)
	"¨": "̈", // U+00A8 → U+0308  (\ddot)
	"˘": "̆", // U+02D8 → U+0306  (\breve)
	ˇ: "̌", // U+02C7 → U+030C  (\check)
	"`": "̀", // U+0060 → U+0300  (\grave)
	"´": "́", // U+00B4 → U+0301  (\acute)
};

function GroupCharacter({
	chr,
	pos,
	body,
}: {
	chr: string;
	pos: "top" | "bot";
	body: XmlNode[];
}): XmlNode {
	return (
		<m.groupChr>
			<m.groupChrPr>
				<m.chr m-val={chr} />
				<m.pos m-val={pos} />
			</m.groupChrPr>
			<m.e>{body}</m.e>
		</m.groupChr>
	);
}

function Delimited({
	open,
	close,
	body,
}: {
	open: string;
	close: string;
	body: XmlNode[];
}): XmlNode {
	return (
		<m.d>
			<m.dPr>
				<m.begChr m-val={open} />
				<m.endChr m-val={close} />
			</m.dPr>
			<m.e>{body}</m.e>
		</m.d>
	);
}

function Matrix({ rows }: { rows: XmlNode[][][] }): XmlNode {
	const colCount = Math.max(1, ...rows.map((r) => r.length));
	return (
		<m.m>
			<m.mPr>
				<m.mcs>
					<m.mc>
						<m.mcPr>
							<m.count m-val={String(colCount)} />
							<m.mcJc m-val="center" />
						</m.mcPr>
					</m.mc>
				</m.mcs>
			</m.mPr>
			{rows.map((row) => (
				<m.mr>
					{row.map((cell) => (
						<m.e>{cell}</m.e>
					))}
				</m.mr>
			))}
		</m.m>
	);
}

/** `<m:eqArr>` aligned equation array — the OMML shape for `\begin{aligned}`,
 *  `\begin{cases}`, etc. One `<m:e>` per row (cells flattened with implicit
 *  alignment marks; OMML doesn't model the `&` column boundary that LaTeX
 *  does, so each row's cells are concatenated). */
function EquationArray({ rows }: { rows: XmlNode[][] }): XmlNode {
	return (
		<m.eqArr>
			{rows.map((row) => (
				<m.e>{row}</m.e>
			))}
		</m.eqArr>
	);
}

/** `<m:bar>` over/underline — `pos="top"` for `\overline`, `"bot"` for
 *  `\underline`. Drawn as a single horizontal rule, unlike `<m:groupChr>`
 *  whose char extends. */
function Bar({ pos, body }: { pos: "top" | "bot"; body: XmlNode[] }): XmlNode {
	return (
		<m.bar>
			<m.barPr>
				<m.pos m-val={pos} />
			</m.barPr>
			<m.e>{body}</m.e>
		</m.bar>
	);
}

/** `<m:borderBox>` enclosing border — OMML's shape for `\boxed{…}`. (`<m:box>`
 *  exists too but it's semantic grouping only; `<m:borderBox>` is the one that
 *  draws the visible rectangle in Word.) */
function Box({ body }: { body: XmlNode[] }): XmlNode {
	return (
		<m.borderBox>
			<m.e>{body}</m.e>
		</m.borderBox>
	);
}

function Phantom({ body }: { body: XmlNode[] }): XmlNode {
	return (
		<m.phant>
			<m.e>{body}</m.e>
		</m.phant>
	);
}

function NaryOperator({
	chr,
	sub,
	sup,
	body,
}: {
	chr: string;
	sub: XmlNode[] | undefined;
	sup: XmlNode[] | undefined;
	body: XmlNode[];
}): XmlNode {
	return (
		<m.nary>
			<m.naryPr>
				<m.chr m-val={chr} />
				<m.limLoc m-val="undOvr" />
				{!sub && <m.subHide m-val="1" />}
				{!sup && <m.supHide m-val="1" />}
			</m.naryPr>
			{sub ? <m.sub>{sub}</m.sub> : <m.sub />}
			{sup ? <m.sup>{sup}</m.sup> : <m.sup />}
			<m.e>{body}</m.e>
		</m.nary>
	);
}

/** `<m:func>` — OMML's semantic representation for "function applied to
 *  argument". Word uses this to insert the conventional thin space between
 *  the function name and its arg (so `\ln 2` renders as `ln 2`, not `ln2`).
 *  This is the standard shape for `\sin x`, `\cos θ`, `\log K`, `\ln Q`,
 *  `\log_{10} N` (where fName is a script over the operator), etc. */
function FunctionApply({
	fName,
	arg,
}: {
	fName: XmlNode;
	arg: XmlNode[];
}): XmlNode {
	return (
		<m.func>
			<m.fName>{fName}</m.fName>
			<m.e>{arg}</m.e>
		</m.func>
	);
}

/** Build the plain-styled run that represents a bare function-name (`ln`,
 *  `sin`, etc.) inside an `<m:fName>`. */
function operatorNameRun(text: string): XmlNode {
	return (
		<m.r>
			<m.rPr>
				<m.sty m-val="p" />
			</m.rPr>
			<m.t>{text}</m.t>
		</m.r>
	);
}

/** `<m:limLow>` / `<m:limUpp>` — a base with a label below/above. Used for
 *  labeled `\underbrace{X}_{Y}` (limLow wrapping a brace groupChr) and the
 *  `\lim_{x→0}` shape. */
function LimitLayout({
	pos,
	base,
	limit,
}: {
	pos: "low" | "upp";
	base: XmlNode[];
	limit: XmlNode[];
}): XmlNode {
	const Tag = pos === "low" ? m.limLow : m.limUpp;
	return (
		<Tag>
			<m.e>{base}</m.e>
			<m.lim>{limit}</m.lim>
		</Tag>
	);
}

// ---------------------------------------------------------------------------
// Dispatcher — walks MathML, builds OMML via the components above.
// ---------------------------------------------------------------------------

/** Walk a sibling list, emitting OMML and consuming siblings when an n-ary
 *  operator absorbs the following body. Returns an array because some MathML
 *  elements (mrow, math) flatten to multiple OMML siblings. */
function convertSiblings(children: XmlNode[]): XmlNode[] {
	const out: XmlNode[] = [];
	for (let i = 0; i < children.length; i++) {
		const child = children[i];
		if (!child) continue;
		// Function-application separator at the OUTER level: temml emits this
		// after a scripted operator (e.g. `\log_{10} N` produces
		// `<msub>log,10</msub><mo>⁡</mo><mspace/><mi>N</mi>` at the same
		// nesting depth). Pop the previously-emitted node as the function
		// name, consume any mspace, then wrap with the next sibling as arg.
		if (isFunctionApplicationOperator(child)) {
			let argIndex = i + 1;
			while (children[argIndex]?.tag === "mspace") argIndex++;
			const argNode = children[argIndex];
			const previous = out[out.length - 1];
			if (
				argNode &&
				!isPostFunctionTerminator(argNode) &&
				previous &&
				looksLikeFunctionName(previous)
			) {
				out.pop();
				out.push(
					<FunctionApply fName={previous} arg={convertElement(argNode)} />,
				);
				i = argIndex;
				continue;
			}
			// Fallback: drop the function-app op and any trailing thin-space.
			if (children[i + 1]?.tag === "mspace") i++;
			continue;
		}
		// Function-name pattern: an INNER mrow temml wraps around a known
		// operator with thin-space padding (`[mspace, mi=ln, mo⁡, mspace]`)
		// followed by the function's argument as the next sibling at the
		// outer level. Combine into `<m:func>` so Word renders the
		// conventional thin space between name and arg.
		const funcName = wrappedOperatorName(child);
		if (funcName) {
			const next = children[i + 1];
			if (next && !isPostFunctionTerminator(next)) {
				out.push(
					<FunctionApply
						fName={operatorNameRun(funcName)}
						arg={convertElement(next)}
					/>,
				);
				i++;
				continue;
			}
			// No suitable arg follows — emit just the operator name (it'll
			// stand alone, e.g. `\sin` at end of expression).
			out.push(<MathRun text={funcName} upright={true} />);
			continue;
		}
		// N-ary pattern: a sub/sup-scripted big-operator absorbs the next
		// sibling as its body.
		if (isNaryScript(child)) {
			const next = children[i + 1];
			out.push(buildNary(child, next));
			if (next) i++;
			continue;
		}
		// Norm pattern: an `<mi>‖</mi>` opens a norm; consume through the
		// matching close into a `<m:d>` with `‖` delimiters. Temml emits the
		// vertical bars as `<mi>`, not `<mo>`, so the generic fence path
		// doesn't catch them.
		if (isNormBar(child)) {
			const close = findNormBarClose(children, i + 1);
			if (close > i) {
				out.push(
					<Delimited
						open="‖"
						close="‖"
						body={convertSiblings(children.slice(i + 1, close))}
					/>,
				);
				i = close;
				continue;
			}
		}
		// Fenced expression pattern: opening fence-mo, content, closing fence-mo
		// within an mrow.
		if (isFenceOpen(child)) {
			const close = findFenceClose(children, i + 1);
			if (close > i) {
				out.push(
					<Delimited
						open={child.collectText()}
						close={children[close]?.collectText() ?? ""}
						body={convertSiblings(children.slice(i + 1, close))}
					/>,
				);
				i = close;
				continue;
			}
		}
		// mhchem empty-base subscript pattern: `<msub><mrow/><n/></msub>`
		// follows a sibling that should become the subscript base. mhchem
		// surrounds each element with `<mspace>` (per ECMA-376 §22's
		// expectation that visible chemistry is space-padded), so we skip
		// over any trailing SpaceRuns first to reach the real previous atom.
		if (hasEmptyBaseScript(child) && out.length > 0) {
			while (out.length > 0 && isEmittedSpaceRun(out[out.length - 1])) {
				out.pop();
			}
			const previous = out.pop();
			if (previous) {
				out.push(wrapEmptyBaseScript(child, previous));
				continue;
			}
		}
		out.push(...convertElement(child));
	}
	return out;
}

/** Lift one MathML element into its OMML form. Returns an array because
 *  some elements (mrow, math) flatten. Accepts `undefined` for convenient
 *  destructuring of fixed-arity children.
 *
 *  Formatting attributes on the element (`style="color:…"`, `mathvariant`,
 *  `mathsize`) are pushed into the `currentFormat` context for the duration
 *  of the conversion so they reach descendant token runs. `<mfrac
 *  style="color:#ff0000">` colors both numerator and denominator. */
function convertElement(node: XmlNode | undefined): XmlNode[] {
	if (!node) return [];
	const run = () => {
		const handler = ELEMENT_HANDLERS[node.tag];
		if (handler) return handler(node);
		if (node.tag === "#text") return [];
		return convertSiblings(node.children);
	};
	const fmt = readNodeFormat(node);
	if (Object.keys(fmt).length === 0) return run();
	return withFormat(fmt, run);
}

const ELEMENT_HANDLERS: Record<string, (node: XmlNode) => XmlNode[]> = {
	math: (node) => convertSiblings(node.children),
	mrow: handleMrow,
	mstyle: (node) => convertSiblings(node.children),
	semantics: (node) =>
		convertSiblings(node.children.filter((c) => c.tag !== "annotation")),
	mspace: (node) => {
		// Temml emits `<mspace></mspace>` (no width) as a zero-width disambiguator
		// between adjacent identifiers (e.g., `\Delta G`); only explicit-width
		// forms (`\,`, `\quad`, `\;`) are real visual space.
		const width = node.getAttribute("width") ?? "";
		if (!width || /^0(?:[a-z]+)?$/.test(width.trim())) return [];
		// Convert width to a count of literal spaces so `\,` (0.1667em) stays
		// thin while `\quad` (1em) and `\qquad` (2em) render visibly wider in
		// Word. The default space char is ~0.25em wide in Cambria Math, so
		// rounding `em / 0.25` gives a reasonable visual approximation.
		const ems = Number.parseFloat(width);
		const spaces = Number.isFinite(ems)
			? Math.max(1, Math.round(ems / 0.25))
			: 1;
		return [<SpaceRun count={spaces} />];
	},
	// Token leaves — a multi-char `<mi>` is a function name (upright); a
	// single-char one is a variable (italic). Per-element `style="color:…"`
	// or `mathvariant` overrides flow through `runFormat`. Variation selectors
	// and other zero-width noise temml emits (e.g., `⊕` + VS15 for `\oplus`)
	// would render as missing-glyph boxes in Word, so we strip them on the
	// way out.
	mi: (node) => {
		const text = cleanMathText(node.collectText());
		if (!text) return [];
		const fmt = runFormat(node);
		// `mathvariant="normal"` forces upright (e.g., `\mathrm{abc}`).
		const upright =
			text.length > 1 || node.getAttribute("mathvariant") === "normal";
		return [<MathRun text={text} upright={upright} {...fmt} />];
	},
	mn: (node) => {
		const text = cleanMathText(node.collectText());
		if (!text) return [];
		return [<MathRun text={text} upright={true} {...runFormat(node)} />];
	},
	mo: (node) => {
		const text = cleanMathText(node.collectText());
		if (!text || INVISIBLE_OPERATORS.has(text)) return [];
		return [<MathRun text={text} upright={true} {...runFormat(node)} />];
	},
	mtext: (node) => {
		const text = cleanMathText(node.collectText());
		if (!text) return [];
		return [<MathRun text={text} upright={true} {...runFormat(node)} />];
	},
	mphantom: (node) => [<Phantom body={convertSiblings(node.children)} />],
	// Scripts
	msup: (node) => {
		const [base, sup] = nonTextChildren(node, 2);
		return [
			<Superscript base={convertElement(base)} sup={convertElement(sup)} />,
		];
	},
	msub: (node) => {
		const [base, sub] = nonTextChildren(node, 2);
		return [
			<Subscript base={convertElement(base)} sub={convertElement(sub)} />,
		];
	},
	msubsup: (node) => {
		const [base, sub, sup] = nonTextChildren(node, 3);
		return [
			<SubSuperscript
				base={convertElement(base)}
				sub={convertElement(sub)}
				sup={convertElement(sup)}
			/>,
		];
	},
	// Fractions and roots
	mfrac: (node) => {
		const [num, den] = nonTextChildren(node, 2);
		const noBar = node.getAttribute("linethickness") === "0px";
		return [
			<Fraction
				num={convertElement(num)}
				den={convertElement(den)}
				noBar={noBar}
			/>,
		];
	},
	msqrt: (node) => [<SquareRoot body={convertSiblings(node.children)} />],
	mroot: (node) => {
		// MathML order: <mroot>{base}{degree}</mroot> — opposite of `\sqrt[d]{x}`.
		const [base, degree] = nonTextChildren(node, 2);
		return [
			<NthRoot degree={convertElement(degree)} body={convertElement(base)} />,
		];
	},
	// Over / under decorations
	mover: handleMover,
	munder: handleMunder,
	munderover: (node) => {
		const [base, under, over] = nonTextChildren(node, 3);
		// Nest under-then-over so both decorations survive — OMML doesn't
		// model "both" in a single element, so we layer two `<m:limLow>` /
		// `<m:limUpp>` indirectly via the over/under group chars.
		const underChr = under?.collectText() ?? "";
		const overChr = over?.collectText() ?? "";
		const inner: XmlNode = (
			<GroupCharacter chr={underChr} pos="bot" body={convertElement(base)} />
		);
		return [<GroupCharacter chr={overChr} pos="top" body={[inner]} />];
	},
	// Tabular structures
	mtable: handleMtable,
	mfenced: (node) => [
		<Delimited
			open={node.getAttribute("open") ?? "("}
			close={node.getAttribute("close") ?? ")"}
			body={convertSiblings(node.children)}
		/>,
	],
	// Enclosing notations: temml emits `\underline`/`\overline` via `<menclose>`
	// with `notation="bottom"`/`"top"`.
	menclose: handleMenclose,
};

// ---------------------------------------------------------------------------
// Element handlers that need conditional logic richer than a one-liner.
// ---------------------------------------------------------------------------

/** `<mrow>` flattens its children, EXCEPT when temml uses it as the `\boxed`
 *  carrier (a styled mrow with `border:1px solid` inline style — there's no
 *  dedicated MathML element for boxed expressions), or as a thin-space
 *  padding wrapper around a single operator/identifier (`[mspace, mo, mspace]`
 *  for the text-math boundary `=`; `[mspace, mi-operator, mo⁡, mspace]` for
 *  `\ln`, `\sin`, etc.). Those paddings are layout decoration temml inserts
 *  and would round-trip as user-visible `\,` whitespace if we kept them. */
function handleMrow(node: XmlNode): XmlNode[] {
	const style = node.getAttribute("style") ?? "";
	if (style.includes("border:1px solid")) {
		return [<Box body={convertSiblings(node.children)} />];
	}
	const collapsed = collapseSpacePaddedOperator(node);
	if (collapsed) return collapsed;
	// Inherited styling (font-weight, color, mathsize) is pushed into
	// `currentFormat` by `convertElement` before this handler runs, so the
	// descendant token runs pick it up automatically — no extra withFormat
	// needed here.
	return convertSiblings(node.children);
}

/** When the mrow's content (filtering out `#text` and `<mspace>`) is just
 *  one token — an operator `<mo>` or a known function-name `<mi>` (possibly
 *  followed by the FUNCTION APPLICATION invisible op) — emit only that
 *  token, dropping the mspace padding. Returns `undefined` when the mrow
 *  has richer content (then the caller flattens normally). */
function collapseSpacePaddedOperator(mrow: XmlNode): XmlNode[] | undefined {
	const meaningful = mrow.children.filter(
		(c) => c.tag !== "#text" && c.tag !== "mspace",
	);
	if (meaningful.length === 0 || meaningful.length > 2) return undefined;
	const first = meaningful[0];
	if (!first) return undefined;
	if (meaningful.length === 2) {
		const second = meaningful[1];
		if (!second || !isFunctionApplicationOperator(second)) return undefined;
	}
	if (first.tag === "mo") {
		// Single-operator wrapping (the `\text{}=…` boundary). Forward the mo
		// through its element handler so attributes (form, stretchy) still
		// drive `<m:d>` fence detection if relevant — but the mrow's mspaces
		// are dropped.
		return convertElement(first);
	}
	if (first.tag === "mi") {
		const text = first.collectText();
		if (KNOWN_OPERATORS_FOR_WRAPPING.has(text)) {
			return [<MathRun text={text} upright={true} />];
		}
	}
	return undefined;
}

/** Return the wrapped function-name when the mrow has the shape temml emits
 *  for `\ln`, `\sin`, etc. — `[mspace?, mi=name, mo⁡?, mspace?]` where `name`
 *  is a known function operator. The function-application dispatcher uses
 *  this to detect the operator-name + argument pattern across siblings. */
function wrappedOperatorName(node: XmlNode): string | undefined {
	if (node.tag !== "mrow") return undefined;
	const meaningful = node.children.filter(
		(c) => c.tag !== "#text" && c.tag !== "mspace",
	);
	if (meaningful.length === 0 || meaningful.length > 2) return undefined;
	const first = meaningful[0];
	if (!first || first.tag !== "mi") return undefined;
	const text = first.collectText();
	if (!KNOWN_OPERATORS_FOR_WRAPPING.has(text)) return undefined;
	if (meaningful.length === 2) {
		const second = meaningful[1];
		if (!second || !isFunctionApplicationOperator(second)) return undefined;
	}
	return text;
}

/** True if `next` is an element that should NOT be absorbed as a function
 *  argument — operators, fences, separators. `\sin + x` shouldn't consume
 *  the `+` as `\sin{+}`; `\ln, x` shouldn't consume the comma. The argument
 *  is any "value-like" element (mi, mn, mfrac, msqrt, mrow with content,
 *  scripts, etc.). */
function isPostFunctionTerminator(next: XmlNode): boolean {
	if (next.tag === "mo") return true;
	if (next.tag === "mspace") return true;
	return false;
}

/** When the FUNCTION APPLICATION marker fires at the OUTER level, we pop the
 *  previously-emitted node as the function name. This guard checks that the
 *  pop target is plausible — a bare operator run, or a script over an
 *  operator (the `\log_{10}` shape). Without this, an unrelated previous
 *  node would get swallowed into `<m:func>` and render incorrectly. */
function looksLikeFunctionName(node: XmlNode): boolean {
	if (node.tag === "m:r") {
		// Plain-styled run is the standard function-name shape.
		return (
			node.findChild("m:rPr")?.findChild("m:sty")?.getAttribute("m:val") === "p"
		);
	}
	if (
		node.tag === "m:sSub" ||
		node.tag === "m:sSup" ||
		node.tag === "m:sSubSup"
	) {
		// Script whose base is a plain run — `\log_{10}`, `\sin^{-1}`, etc.
		const base = node.findChild("m:e");
		const baseRun = base?.findChild("m:r");
		return (
			baseRun?.findChild("m:rPr")?.findChild("m:sty")?.getAttribute("m:val") ===
			"p"
		);
	}
	return false;
}

/** Mirror of the reader's KNOWN_OPERATORS — when temml wraps these in an
 *  mrow with thin-space padding, we collapse to just the operator name run. */
const KNOWN_OPERATORS_FOR_WRAPPING = new Set([
	"lim",
	"liminf",
	"limsup",
	"sin",
	"cos",
	"tan",
	"sec",
	"csc",
	"cot",
	"arcsin",
	"arccos",
	"arctan",
	"sinh",
	"cosh",
	"tanh",
	"coth",
	"log",
	"ln",
	"lg",
	"exp",
	"max",
	"min",
	"sup",
	"inf",
	"det",
	"gcd",
	"dim",
	"ker",
	"hom",
	"deg",
	"arg",
	"Pr",
]);

/** `<mover>` may carry an accent (non-stretchy → `<m:acc>`), a wide accent
 *  (stretchy → `<m:groupChr>` so the reader maps to `\widehat`/`\widetilde`/
 *  `\overrightarrow`), or a labeled overbrace (nested mover with `⏞` →
 *  `<m:limUpp>` wrapping a `<GroupCharacter>`). Order matters: detect the
 *  nested-label case first. */
function handleMover(node: XmlNode): XmlNode[] {
	const [base, over] = nonTextChildren(node, 2);
	if (!base || !over) return convertSiblings(node.children);

	// Labeled \overbrace: <mover><mover>X<mo>⏞</mo></mover>Y</mover>
	const innerGroup = labeledBraceInner(base, "mover");
	if (innerGroup) {
		return [
			<LimitLayout
				pos="upp"
				base={[innerGroup]}
				limit={convertElement(over)}
			/>,
		];
	}

	// Chemistry arrow / spacer-over: <mover><mo>→</mo><mspace/></mover>
	// (mhchem). The over contributes no visible glyph — emit just the base.
	if (over.tag === "mspace" || isWhitespaceText(over)) {
		return convertElement(base);
	}

	// Arbitrary-content `\overset{X}{base}` / `\stackrel{X}{base}` — when the
	// over is structured (mrow / mtext / mfrac) or multi-char text, it's a
	// label not an accent char. Route to `<m:limUpp>` which can hold any OMML
	// content as its `<m:lim>`; reader emits `\overset{…}{…}` for non-function
	// bases.
	if (isArbitraryLabel(over)) {
		return [
			<LimitLayout
				pos="upp"
				base={convertElement(base)}
				limit={convertElement(over)}
			/>,
		];
	}

	const overChr = over.collectText();
	const body = convertElement(base);
	const stretchy = over.getAttribute("stretchy") === "true";

	if (stretchy && WIDE_OVER_CHARS.has(overChr)) {
		return [<GroupCharacter chr={overChr} pos="top" body={body} />];
	}
	if (ACCENT_OVER_CHARS.has(overChr)) {
		return [<OverAccent chr={overChr} body={body} />];
	}
	return [<GroupCharacter chr={overChr} pos="top" body={body} />];
}

/** `<munder>` mirrors `<mover>` — including the labeled-underbrace nested
 *  shape (`<munder><munder>X<mo>⏟</mo></munder>Y</munder>`) and the arbitrary
 *  `\underset{X}{base}` shape. */
function handleMunder(node: XmlNode): XmlNode[] {
	const [base, under] = nonTextChildren(node, 2);
	if (!base || !under) return convertSiblings(node.children);

	const innerGroup = labeledBraceInner(base, "munder");
	if (innerGroup) {
		return [
			<LimitLayout
				pos="low"
				base={[innerGroup]}
				limit={convertElement(under)}
			/>,
		];
	}

	if (isArbitraryLabel(under)) {
		return [
			<LimitLayout
				pos="low"
				base={convertElement(base)}
				limit={convertElement(under)}
			/>,
		];
	}

	const chr = under.collectText();
	return [<GroupCharacter chr={chr} pos="bot" body={convertElement(base)} />];
}

/** True if the over/under content is a "label" rather than a single
 *  decoration char (`^`, `~`, `→`, `⏞`, …). Labels route through
 *  `<m:limUpp>`/`<m:limLow>` instead of `<m:acc>`/`<m:groupChr>`. Anything
 *  that's not in the accent / wide-accent / brace character tables is a
 *  label by default — that catches `\overset{?}{=}` (single non-accent
 *  char) and `\overset{\text{cat}}{\to}` (multi-char structured content). */
function isArbitraryLabel(node: XmlNode): boolean {
	// Structured wrappers always indicate a label.
	if (node.tag === "mtext" || node.tag === "mrow" || node.tag === "mfrac")
		return true;
	if (
		node.tag === "msup" ||
		node.tag === "msub" ||
		node.tag === "msubsup" ||
		node.tag === "msqrt"
	)
		return true;
	const text = node.collectText();
	if (text.length > 1) return true;
	// Single char: it's a label unless it's recognized as a decoration.
	return (
		!ACCENT_OVER_CHARS.has(text) &&
		!WIDE_OVER_CHARS.has(text) &&
		!BRACE_CHARS.has(text)
	);
}

/** When the base of an outer `<munder>`/`<mover>` is itself the SAME element
 *  with a brace char as its decoration, this is temml's labeled-brace shape:
 *  the inner element is the visual brace; the outer adds the label. Returns
 *  the GroupCharacter for the inner brace so the caller can wrap in
 *  limLow/limUpp. */
function labeledBraceInner(
	candidate: XmlNode,
	parentTag: "munder" | "mover",
): XmlNode | undefined {
	if (candidate.tag !== parentTag) return undefined;
	const [innerBase, innerDeco] = nonTextChildren(candidate, 2);
	if (!innerBase || !innerDeco) return undefined;
	const chr = innerDeco.collectText();
	if (!BRACE_CHARS.has(chr)) return undefined;
	const pos = parentTag === "munder" ? "bot" : "top";
	return (
		<GroupCharacter chr={chr} pos={pos} body={convertElement(innerBase)} />
	);
}

/** `<mtable>` has three shapes we care about: aligned (class="tml-jot",
 *  `\begin{aligned}` → `<m:eqArr>`), tagged equation (class="tml-tageqn",
 *  `\tag{N}` → equation + trailing tag run), and plain (everything else →
 *  `<m:m>` matrix). */
function handleMtable(node: XmlNode): XmlNode[] {
	const tableClass = node.getAttribute("class") ?? "";
	const rows = node.findChildren("mtr");

	// Tag detection: temml puts `class="tml-tageqn"` on the `<mtr>`, not the
	// `<mtable>`. The shape is a single row with `[spacer, equation, tag]`.
	const firstRowClass = rows[0]?.getAttribute("class") ?? "";
	if (firstRowClass.includes("tml-tageqn")) {
		return handleTaggedEquation(node);
	}

	const cells = rows.map((row) =>
		row.findChildren("mtd").map((cell) => convertSiblings(cell.children)),
	);

	if (tableClass.includes("tml-jot")) {
		// Aligned: flatten each row's cells (eqArr has no column model).
		const flatRows = cells.map((row) => row.flat());
		return [<EquationArray rows={flatRows} />];
	}

	return [<Matrix rows={cells} />];
}

/** Temml emits `\tag{N}` as a 3-cell single-row `<mtable>`:
 *  `[spacer, equation, <mtext class="tml-tag">(N)</mtext>]`. We extract the
 *  equation as the leading siblings and append the tag as a trailing text
 *  run — close enough to the rendered layout that downstream consumers see
 *  `equation   (N)` on the same line. */
function handleTaggedEquation(table: XmlNode): XmlNode[] {
	const row = table.findChild("mtr");
	if (!row) return [];
	const cells = row.findChildren("mtd");
	if (cells.length < 2) return convertSiblings(table.children);
	// The equation lives in the middle cell; the tag in the last.
	const equationCell = cells[Math.floor(cells.length / 2)];
	const tagCell = cells[cells.length - 1];
	const equationContent = equationCell
		? convertSiblings(equationCell.children)
		: [];
	const tagText = tagCell?.collectText().trim() ?? "";
	if (!tagText) return equationContent;
	return [
		...equationContent,
		<SpaceRun />,
		<MathRun text={tagText} upright={true} />,
	];
}

/** `<menclose notation="bottom">` → `<m:bar pos="bot">` (\underline);
 *  `notation="top"` → `<m:bar pos="top">` (\overline);
 *  `notation="box"` (and a few synonyms) → `<m:box>` (\boxed);
 *  `notation="updiagonalstrike"`/`"downdiagonalstrike"`/`"horizontalstrike"`
 *  → propagate `strike` formatting to descendant runs (`\cancel`, `\bcancel`,
 *  `\sout`). OMML has no dedicated "strike through math" element, so we use
 *  Word's `<w:strike/>` inside each affected math run. Filter out temml's
 *  decoration `<mrow class="tml-cancel …"></mrow>` placeholder. */
function handleMenclose(node: XmlNode): XmlNode[] {
	const notation = node.getAttribute("notation") ?? "";
	const childrenForBody = node.children.filter(
		(c) =>
			!(c.tag === "mrow" && /tml-cancel/.test(c.getAttribute("class") ?? "")),
	);
	const hasStrike =
		notation.includes("updiagonalstrike") ||
		notation.includes("downdiagonalstrike") ||
		notation.includes("horizontalstrike");
	if (hasStrike) {
		return withFormat({ strike: true }, () => convertSiblings(childrenForBody));
	}
	const body = convertSiblings(childrenForBody);
	if (notation.includes("bottom")) return [<Bar pos="bot" body={body} />];
	if (notation.includes("top")) return [<Bar pos="top" body={body} />];
	if (
		notation.includes("box") ||
		notation.includes("roundedbox") ||
		notation.includes("circle")
	) {
		return [<Box body={body} />];
	}
	return body;
}

const INVISIBLE_OPERATORS = new Set([
	"⁡", // FUNCTION APPLICATION
	"⁢", // INVISIBLE TIMES
	"⁣", // INVISIBLE SEPARATOR
	"⁤", // INVISIBLE PLUS
]);

/** Specifically the FUNCTION APPLICATION (U+2061) — temml emits this between
 *  a function-operator name and its argument, usually followed by an mspace
 *  thin-space. Our operator-name promotion already adds a separator space on
 *  read, so the pair is redundant; the dispatcher uses this to consume it. */
function isFunctionApplicationOperator(node: XmlNode): boolean {
	return node.tag === "mo" && node.collectText() === "⁡";
}

// ---------------------------------------------------------------------------
// N-ary recognition (sum, integral, product, etc.)
// ---------------------------------------------------------------------------

/** Recognize the n-ary pattern temml emits: msubsup/msub/mover/munder/
 *  munderover whose FIRST non-text child is a big-operator `<mo>`. */
function isNaryScript(node: XmlNode): boolean {
	if (
		node.tag !== "msubsup" &&
		node.tag !== "msub" &&
		node.tag !== "msup" &&
		node.tag !== "munderover" &&
		node.tag !== "munder" &&
		node.tag !== "mover"
	)
		return false;
	const first = nonTextChildren(node, 1)[0];
	if (!first || first.tag !== "mo") return false;
	return BIG_OP_CHARS.has(first.collectText());
}

function buildNary(script: XmlNode, body: XmlNode | undefined): XmlNode {
	const children = nonTextChildren(script, 3);
	const op = children[0];
	const opChr = op?.collectText() ?? "";
	let subNode: XmlNode | undefined;
	let supNode: XmlNode | undefined;
	if (script.tag === "msubsup" || script.tag === "munderover") {
		subNode = children[1];
		supNode = children[2];
	} else if (script.tag === "msub" || script.tag === "munder") {
		subNode = children[1];
	} else if (script.tag === "msup" || script.tag === "mover") {
		supNode = children[1];
	}
	return (
		<NaryOperator
			chr={opChr}
			sub={subNode ? convertElement(subNode) : undefined}
			sup={supNode ? convertElement(supNode) : undefined}
			body={body ? convertElement(body) : []}
		/>
	);
}

const BIG_OP_CHARS = new Set([
	"∑",
	"∏",
	"∐",
	"∫",
	"∬",
	"∭",
	"∮",
	"∯",
	"∰",
	"⋂",
	"⋃",
	"⨁",
	"⨂",
	"⨅",
	"⨆",
]);

// ---------------------------------------------------------------------------
// mhchem empty-base subscript (`\ce{H2O}` → `H<empty>_{2}O`).
// ---------------------------------------------------------------------------

/** mhchem emits subscripts by placing an empty `<mrow/>` as the base of an
 *  `<msub>`/`<msup>`, expecting the visible base to come from the previous
 *  sibling. */
function hasEmptyBaseScript(node: XmlNode): boolean {
	if (node.tag !== "msub" && node.tag !== "msup" && node.tag !== "msubsup")
		return false;
	const [base] = nonTextChildren(node, 1);
	if (!base) return false;
	if (base.tag !== "mrow") return false;
	return nonTextChildren(base, 1).length === 0;
}

function wrapEmptyBaseScript(node: XmlNode, base: XmlNode): XmlNode {
	const [, second, third] = nonTextChildren(node, 3);
	if (node.tag === "msub") {
		return <Subscript base={[base]} sub={convertElement(second)} />;
	}
	if (node.tag === "msup") {
		return <Superscript base={[base]} sup={convertElement(second)} />;
	}
	return (
		<SubSuperscript
			base={[base]}
			sub={convertElement(second)}
			sup={convertElement(third)}
		/>
	);
}

// ---------------------------------------------------------------------------
// Norm bars (temml emits `\|` as `<mi>‖</mi>`, not `<mo>`, so the generic
// fence detection misses them).
// ---------------------------------------------------------------------------

function isNormBar(node: XmlNode): boolean {
	if (node.tag !== "mi") return false;
	return node.collectText() === "‖";
}

function findNormBarClose(children: XmlNode[], start: number): number {
	for (let i = start; i < children.length; i++) {
		const child = children[i];
		if (child && isNormBar(child)) return i;
	}
	return -1;
}

// ---------------------------------------------------------------------------
// Fenced expressions (parens / brackets / braces wrapping content)
// ---------------------------------------------------------------------------

function isFenceOpen(node: XmlNode): boolean {
	if (node.tag !== "mo") return false;
	if (node.getAttribute("fence") !== "true") return false;
	// Atomic parens / brackets — temml marks `stretchy="false"` on
	// non-`\left/\right` fences. Skipping the `<m:d>` wrap keeps round-trips
	// honest: `(x_i - \mu)^2` reads back as itself, not
	// `{\left(x_i-\mu\right)}^2`. `\left(…\right)` content carries
	// `stretchy="true"` and still goes through Delimited.
	if (node.getAttribute("stretchy") === "false") return false;
	if (node.getAttribute("form") === "prefix") return true;
	return OPEN_FENCES.has(node.collectText());
}

function isFenceClose(node: XmlNode): boolean {
	if (node.tag !== "mo") return false;
	if (node.getAttribute("fence") !== "true") return false;
	if (node.getAttribute("stretchy") === "false") return false;
	if (node.getAttribute("form") === "postfix") return true;
	return CLOSE_FENCES.has(node.collectText());
}

function findFenceClose(children: XmlNode[], start: number): number {
	let depth = 0;
	for (let i = start; i < children.length; i++) {
		const child = children[i];
		if (!child) continue;
		// Symmetric delimiters (`|`, `‖`) appear in both OPEN_FENCES and
		// CLOSE_FENCES — for those, `form="postfix"` is the unambiguous "this
		// is a close" signal, so check it first.
		const form = child.getAttribute("form");
		if (form === "postfix" && isFenceClose(child)) {
			if (depth === 0) return i;
			depth--;
			continue;
		}
		if (isFenceOpen(child) && form !== "postfix") {
			depth++;
		} else if (isFenceClose(child)) {
			if (depth === 0) return i;
			depth--;
		}
	}
	return -1;
}

const OPEN_FENCES = new Set(["(", "[", "{", "⟨", "⌈", "⌊", "|", "‖"]);
const CLOSE_FENCES = new Set([")", "]", "}", "⟩", "⌉", "⌋", "|", "‖"]);

// ---------------------------------------------------------------------------
// Accent / wide-accent character sets
// ---------------------------------------------------------------------------

/** Over-chars that we recognize as accents (so we emit `<OverAccent>` rather
 *  than `<GroupCharacter>`). Includes both temml's preferred spacing/
 *  overscript codepoints and the combining-diacritic forms Word emits — same
 *  bidirectional coverage as the reader's `ACCENT_CMD_BY_CHR` in
 *  [handlers.ts](./handlers.ts). */
const ACCENT_OVER_CHARS = new Set([
	"^",
	"~",
	"¯",
	"‾",
	"˙",
	"¨",
	"ˇ",
	"˘",
	"`",
	"´",
	"→",
	"̂",
	"̃",
	"̄",
	"̇",
	"̈",
	"̌",
	"̆",
	"̀",
	"́",
	"⃗",
	"⃡",
	"⃛",
]);

/** Wide / stretchy variants — temml marks these with `stretchy="true"`. The
 *  reader's `GROUP_CHR_CMD_BY_CHR` maps each back to `\widehat` / `\widetilde`
 *  / `\overrightarrow` / `\overline`, distinguishing them from the
 *  non-stretchy `\hat` / `\tilde` / `\vec` / `\bar`. */
const WIDE_OVER_CHARS = new Set(["^", "~", "→", "‾", "¯"]);

/** Brace chars that mark a labeled-brace pattern (`\underbrace` / `\overbrace`
 *  with subscript or superscript labels). The decoration char itself doesn't
 *  matter for the outer dispatch — we look at the inner-mover/munder shape. */
const BRACE_CHARS = new Set(["⏞", "⏟", "⏜", "⏝", "⏠", "⏡"]);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Return the first `n` non-text children — MathML element-content models
 *  don't allow mixed content, so the `#text` nodes from XML pretty-printing
 *  are pure formatting and should be skipped. */
function nonTextChildren(node: XmlNode, n: number): XmlNode[] {
	const out: XmlNode[] = [];
	for (const child of node.children) {
		if (child.tag === "#text") continue;
		out.push(child);
		if (out.length === n) break;
	}
	return out;
}

function isWhitespaceText(node: XmlNode): boolean {
	if (node.tag !== "mtext") return false;
	return node.collectText().trim() === "";
}

/** Recognize a previously-emitted SpaceRun by its tag-and-shape signature
 *  rather than by reference equality — JSX-built XmlNodes are fresh each call
 *  and equality wouldn't survive the dispatcher. */
function isEmittedSpaceRun(node: XmlNode | undefined): boolean {
	if (!node || node.tag !== "m:r") return false;
	const t = node.findChild("m:t");
	if (!t) return false;
	return t.collectText() === " ";
}
