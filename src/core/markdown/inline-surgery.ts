import type { PhrasingContent, Root } from "mdast";
import type { Plugin } from "unified";
import { visit } from "unist-util-visit";
import { XmlNode } from "../parser";

/** Remark transformer: post-parse, rewrite every phrasing-children array,
 * gathering CriticMarkup (`{++…++}` / `{--…--}`) and Pandoc bracketed spans
 * (`[text]{.class key="value"}`) into parent nodes that carry the phrasing
 * content *between* their markers.
 *
 * Operating on whole sibling arrays — not individual `text` nodes the way the
 * old text-split plugin did — is what fixes the straddling trap: `{++**bold**++}`
 * parses to `[text("{++"), strong, text("++}")]`, and gathering between the
 * matched markers keeps the `strong` child. Markers that don't balance, and
 * spans whose attributes parse to nothing, degrade to their literal text —
 * never a throw, never lost content (the "unsupported doesn't break" invariant).
 *
 * `visit` runs pre-order; replacing `node.children` in the visitor makes `visit`
 * descend into the rewritten array, so nested markers (`**{++x++}**`) are caught
 * when their wrapper is visited. Re-reducing an already-reduced array is a
 * no-op: a degraded close always precedes its degraded open within one array,
 * so the second pass can't pair them. */
export const remarkInlineSurgery: Plugin<[], Root> = () => (tree) => {
	visit(tree, (node) => {
		const mutable = node as unknown as MutableNode;
		if (!Array.isArray(mutable.children)) return;
		// Two independent reductions over the same sibling array: text-marker
		// surgery (CriticMarkup + Pandoc spans) and HTML-element surgery
		// (<span>/<mark>/<sup>/<sub>/<u> — what `read` now emits). `visit` is
		// pre-order and descends into the replaced array, so nested wrappers are
		// reduced when their parent is visited.
		mutable.children = gatherHtmlSpans(transformChildren(mutable.children));
	});
};

/** Reduce one phrasing-children array. Fast-path bails when no text child holds
 * a marker character; any unexpected failure restores the original array. */
function transformChildren(children: MutableNode[]): MutableNode[] {
	const hasMarker = children.some(
		(child) => child.type === "text" && MARKER_CHAR.test(child.value ?? ""),
	);
	if (!hasMarker) return children;
	try {
		const tokens = tokenize(children);
		if (
			!tokens.some((token) => token.kind === "open" || token.kind === "close")
		) {
			return children;
		}
		return reduce(tokens);
	} catch {
		return children;
	}
}

/** Flatten a children array into a token stream: contiguous `text` values are
 * scanned for markers; every non-text node passes through as an opaque atom
 * (so `inlineCode`/`strong`/`link`/math are never looked inside — code spans
 * are excluded for free). */
function tokenize(children: MutableNode[]): Token[] {
	const tokens: Token[] = [];
	let buffer = "";
	for (const child of children) {
		if (child.type === "text") {
			buffer += child.value ?? "";
			continue;
		}
		flushLiteral(buffer, tokens);
		buffer = "";
		tokens.push({ kind: "atom", node: child });
	}
	flushLiteral(buffer, tokens);
	return tokens;
}

function flushLiteral(text: string, tokens: Token[]): void {
	if (text.length === 0) return;
	let cursor = 0;
	for (const match of text.matchAll(MARKER)) {
		const index = match.index;
		if (index > cursor) {
			tokens.push({ kind: "literal", value: text.slice(cursor, index) });
		}
		const raw = match[0];
		if (raw === "{++") tokens.push({ kind: "open", marker: "ins", raw });
		else if (raw === "{--") tokens.push({ kind: "open", marker: "del", raw });
		else if (raw === "++}") tokens.push({ kind: "close", marker: "ins", raw });
		else if (raw === "--}") tokens.push({ kind: "close", marker: "del", raw });
		else if (raw === "[") tokens.push({ kind: "open", marker: "span", raw });
		else tokens.push({ kind: "close", marker: "span", raw, attrs: match[1] });
		cursor = index + raw.length;
	}
	if (cursor < text.length) {
		tokens.push({ kind: "literal", value: text.slice(cursor) });
	}
}

/** Linear scan with an explicit open-marker stack. A close pops the nearest
 * matching open and builds its node; anything that can't be matched/built is
 * re-emitted as literal text so no content is dropped. */
function reduce(tokens: Token[]): MutableNode[] {
	const root: MutableNode[] = [];
	const stack: Frame[] = [];
	const top = (): MutableNode[] =>
		stack.length > 0 ? (stack[stack.length - 1] as Frame).children : root;

	for (const token of tokens) {
		if (token.kind === "literal") {
			pushText(top(), token.value);
			continue;
		}
		if (token.kind === "atom") {
			top().push(token.node);
			continue;
		}
		if (token.kind === "open") {
			stack.push({ marker: token.marker, raw: token.raw, children: [] });
			continue;
		}
		// A `]{…}` whose attributes parse to nothing recognized isn't really a
		// span close — emit it literally and keep scanning, so a span's content
		// may itself contain `]{…}` sequences and a *later* real close still
		// matches (otherwise `[a]{x}b]{color=…}` would close early at `]{x}`,
		// dropping the color and leaking literal brackets).
		let attributes: SpanAttributes | undefined;
		if (token.marker === "span") {
			attributes = parseAttrs(token.attrs ?? "");
			if (Object.keys(attributes).length === 0) {
				pushText(top(), token.raw);
				continue;
			}
		}
		const matchIndex = findMatchingOpen(stack, token.marker);
		if (matchIndex === -1) {
			pushText(top(), token.raw);
			continue;
		}
		// Improper nesting: spill any frames opened above the match as literal.
		while (stack.length - 1 > matchIndex) {
			spill(stack.pop() as Frame, top());
		}
		const frame = stack.pop() as Frame;
		top().push(buildNode(frame, attributes));
	}
	// Unclosed openers degrade to literal in opening order.
	while (stack.length > 0) {
		spill(stack.pop() as Frame, top());
	}
	return root;
}

function findMatchingOpen(stack: Frame[], marker: Marker): number {
	for (let index = stack.length - 1; index >= 0; index--) {
		if ((stack[index] as Frame).marker === marker) return index;
	}
	return -1;
}

/** Re-emit an unmatched open frame as its literal marker followed by its
 * accumulated children, into `destination`. */
function spill(frame: Frame, destination: MutableNode[]): void {
	pushText(destination, frame.raw);
	appendMerged(destination, frame.children);
}

/** Append children to `destination`, folding text nodes into the trailing text
 * node so spilled literals don't leave adjacent fragments (spurious runs). */
function appendMerged(
	destination: MutableNode[],
	children: MutableNode[],
): void {
	for (const child of children) {
		if (child.type === "text") pushText(destination, child.value ?? "");
		else destination.push(child);
	}
}

/** Build the parent node for a matched frame. Span attributes are pre-parsed
 * and guaranteed non-empty by the caller (an empty-attr `]{…}` never reaches a
 * match — it's treated as literal text). */
function buildNode(
	frame: Frame,
	attributes: SpanAttributes | undefined,
): MutableNode {
	if (frame.marker === "ins") {
		return { type: "criticInsert", children: frame.children };
	}
	if (frame.marker === "del") {
		return { type: "criticDelete", children: frame.children };
	}
	return {
		type: "bracketedSpan",
		attributes: attributes ?? {},
		children: frame.children,
	};
}

/** Append text to a children array, merging into a trailing `text` node so the
 * output stays free of adjacent text fragments. */
function pushText(destination: MutableNode[], value: string): void {
	if (value.length === 0) return;
	const last = destination[destination.length - 1];
	if (last && last.type === "text") {
		last.value = (last.value ?? "") + value;
		return;
	}
	destination.push({ type: "text", value });
}

// ── HTML-element surgery ──────────────────────────────────────────────────
// `read` emits run formatting as HTML (`<span style>`, `<mark>`, `<sup>`,
// `<sub>`, `<u>`) because that's what a markdown reader actually renders. remark
// leaves inline HTML as FLAT tokens — `<span …>` and `</span>` are separate
// `html` siblings with the wrapped markdown between them — so we re-pair them
// here, the same way the text-marker reducer above pairs `{++…++}` / `[…]{…}`.
// Each matched pair becomes a `bracketedSpan` the inline walker already overlays.

/** Tags `read` emits to carry run formatting; any other inline HTML (locator
 * comments, unknown tags) passes through and is dropped downstream, so its text
 * survives unformatted. */
const HTML_FORMAT_TAGS: ReadonlySet<string> = new Set([
	"span",
	"mark",
	"sup",
	"sub",
	"u",
	// `read` emits these for emphasis on whitespace-only runs (markdown `** **`
	// mis-parses); both the short and long forms are accepted on import.
	"b",
	"strong",
	"i",
	"em",
	"s",
]);

/** Re-pair `read`'s formatting tags across the flat `html` tokens remark leaves
 * them as. Fast-path bails when no recognized formatting tag is present; any
 * unexpected failure restores the original array (the "unsupported doesn't
 * break" invariant). */
function gatherHtmlSpans(children: MutableNode[]): MutableNode[] {
	const present = children.some(
		(child) =>
			child.type === "html" &&
			HTML_FORMAT_TAGS.has(htmlTagName(child.value ?? "") ?? ""),
	);
	if (!present) return children;
	try {
		return reduceHtml(children);
	} catch {
		return children;
	}
}

/** Linear scan with an open-tag stack (mirrors `reduce` for text markers). A
 * close pops the nearest matching open and wraps its gathered children in a
 * `bracketedSpan`; unmatched tags / unclosed opens degrade so no content drops. */
function reduceHtml(children: MutableNode[]): MutableNode[] {
	const root: MutableNode[] = [];
	const stack: HtmlFrame[] = [];
	const top = (): MutableNode[] =>
		stack.length > 0 ? (stack[stack.length - 1] as HtmlFrame).children : root;

	for (const child of children) {
		const tag = child.type === "html" ? classifyTag(child.value ?? "") : null;
		if (!tag) {
			top().push(child);
			continue;
		}
		if (tag.kind === "open") {
			stack.push({ name: tag.name, attributes: tag.attributes, children: [] });
			continue;
		}
		const matchIndex = findMatchingHtmlOpen(stack, tag.name);
		if (matchIndex === -1) {
			top().push(child); // stray close — leave raw (dropped downstream)
			continue;
		}
		// Improper nesting: spill any frames opened above the match as their
		// (unformatted) content so nothing is lost.
		while (stack.length - 1 > matchIndex) {
			spillHtml(stack.pop() as HtmlFrame, top());
		}
		const frame = stack.pop() as HtmlFrame;
		top().push({
			type: "bracketedSpan",
			attributes: frame.attributes,
			children: frame.children,
		});
	}
	while (stack.length > 0) spillHtml(stack.pop() as HtmlFrame, top());
	return root;
}

function findMatchingHtmlOpen(stack: HtmlFrame[], name: string): number {
	for (let index = stack.length - 1; index >= 0; index--) {
		if ((stack[index] as HtmlFrame).name === name) return index;
	}
	return -1;
}

/** Drop an unmatched open frame's tag but keep its gathered children (the tag
 * would be dropped downstream anyway; this preserves the content). */
function spillHtml(frame: HtmlFrame, destination: MutableNode[]): void {
	for (const child of frame.children) destination.push(child);
}

/** Classify a lone HTML tag string as an open/close of a known formatting tag,
 * or null for anything else (comments, unknown tags, non-tags). Open tags get
 * their attributes mapped to `SpanAttributes`. */
function classifyTag(
	raw: string,
):
	| { kind: "open"; name: string; attributes: SpanAttributes }
	| { kind: "close"; name: string }
	| null {
	const trimmed = raw.trim();
	const close = trimmed.match(/^<\/([a-zA-Z][a-zA-Z0-9]*)\s*>$/);
	if (close) {
		const name = (close[1] as string).toLowerCase();
		return HTML_FORMAT_TAGS.has(name) ? { kind: "close", name } : null;
	}
	const open = trimmed.match(/^<([a-zA-Z][a-zA-Z0-9]*)((?:\s[^>]*)?)\/?>$/);
	if (!open) return null;
	const name = (open[1] as string).toLowerCase();
	if (!HTML_FORMAT_TAGS.has(name)) return null;
	return {
		kind: "open",
		name,
		attributes: htmlTagToAttributes(name, open[2] ?? ""),
	};
}

/** The lowercased tag name of a lone open/close tag string, or null when the
 * string isn't a single tag (HTML comment, text, …). */
function htmlTagName(raw: string): string | null {
	const match = raw
		.trim()
		.match(/^<\/?([a-zA-Z][a-zA-Z0-9]*)(?:\s[^>]*)?\/?>$/);
	return match ? (match[1] as string).toLowerCase() : null;
}

/** Map a recognized formatting tag + its raw attribute string to the run
 * formatting it carries. Semantic tags imply their property (`<mark>` →
 * highlight, `<sup>` → superscript); `<span>`/`<u>` read their CSS + `data-*`
 * via the project's XML parser. */
function htmlTagToAttributes(name: string, rawAttrs: string): SpanAttributes {
	if (name === "b" || name === "strong") return { bold: true };
	if (name === "i" || name === "em") return { italic: true };
	if (name === "s") return { strike: true };
	const node = parseTagAttributes(name, rawAttrs);
	if (name === "mark") {
		return { highlight: node?.getAttribute("data-highlight") ?? "yellow" };
	}
	if (name === "sup") return { vertAlign: "superscript" };
	if (name === "sub") return { vertAlign: "subscript" };
	if (name === "u") {
		const attributes: SpanAttributes = {
			underline: node?.getAttribute("data-underline") ?? "single",
		};
		const color = node?.getAttribute("data-underline-color");
		if (color) attributes.underlineColor = color;
		return attributes;
	}
	// span
	const attributes: SpanAttributes = {};
	if (!node) return attributes;
	const style = node.getAttribute("style");
	if (style) applyCssStyle(attributes, style);
	const theme = node.getAttribute("data-color-theme");
	if (theme) attributes.colorTheme = theme;
	const tint = node.getAttribute("data-color-theme-tint");
	if (tint) attributes.colorThemeTint = tint;
	const shade = node.getAttribute("data-color-theme-shade");
	if (shade) attributes.colorThemeShade = shade;
	return attributes;
}

/** Parse a single tag's attributes with the project's XML parser (re-closing it
 * so it's well-formed). Returns undefined on any parse failure — the caller then
 * treats the tag as attribute-less rather than throwing. */
function parseTagAttributes(
	name: string,
	rawAttrs: string,
): XmlNode | undefined {
	try {
		return XmlNode.parse(`<${name}${rawAttrs}/>`)[0];
	} catch {
		return undefined;
	}
}

/** Fold the CSS-expressible run props out of a `style="…"` declaration block
 * into `attributes`. The inverse of `wrapRunFormatting` in cli/read/markdown.ts
 * — keep the two property mappings in sync (the read↔import contract). */
function applyCssStyle(attributes: SpanAttributes, style: string): void {
	for (const declaration of style.split(";")) {
		const colon = declaration.indexOf(":");
		if (colon === -1) continue;
		const prop = declaration.slice(0, colon).trim().toLowerCase();
		const value = declaration.slice(colon + 1).trim();
		if (value.length === 0) continue;
		if (prop === "color") attributes.color = stripHash(value);
		else if (prop === "background-color") attributes.shade = stripHash(value);
		else if (prop === "font-family") attributes.font = stripQuotes(value);
		else if (prop === "font-size") {
			const points = Number.parseFloat(value.replace(/pt$/i, ""));
			if (Number.isFinite(points)) {
				attributes.sizeHalfPoints = Math.round(points * 2);
			}
		} else if (prop === "font-variant" && value === "small-caps") {
			attributes.smallCaps = true;
		} else if (prop === "text-transform" && value === "uppercase") {
			attributes.allCaps = true;
		}
	}
}

function stripHash(value: string): string {
	return value.startsWith("#") ? value.slice(1) : value;
}

type HtmlFrame = {
	name: string;
	attributes: SpanAttributes;
	children: MutableNode[];
};

/** Parse a Pandoc attribute block (`#id .class key="value"`). Classes map to
 * the boolean/enum properties they stand for; key/value pairs map by name.
 * Quoted values may contain spaces (font names); unknown tokens are ignored.
 * Validation of enum values (highlight, underline) happens in the walker, so a
 * bad value surfaces as a clean USAGE error rather than corrupt OOXML. */
function parseAttrs(raw: string): SpanAttributes {
	const attributes: SpanAttributes = {};
	for (const match of raw.matchAll(ATTR)) {
		const key = match[1] as string;
		const rawValue = match[2];
		if (key.startsWith("#")) continue;
		if (key.startsWith(".")) {
			applyClass(attributes, key.slice(1));
			continue;
		}
		if (rawValue === undefined) continue;
		applyKeyValue(attributes, key, stripQuotes(rawValue));
	}
	return attributes;
}

function applyClass(attributes: SpanAttributes, className: string): void {
	if (className === "underline") attributes.underline ??= "single";
	else if (className === "smallcaps") attributes.smallCaps = true;
	else if (className === "allcaps") attributes.allCaps = true;
	else if (className === "sup") attributes.vertAlign = "superscript";
	else if (className === "sub") attributes.vertAlign = "subscript";
}

function applyKeyValue(
	attributes: SpanAttributes,
	key: string,
	value: string,
): void {
	switch (key) {
		case "color":
			attributes.color = value;
			break;
		case "colorTheme":
			attributes.colorTheme = value;
			break;
		case "colorThemeTint":
			attributes.colorThemeTint = value;
			break;
		case "colorThemeShade":
			attributes.colorThemeShade = value;
			break;
		case "highlight":
			attributes.highlight = value;
			break;
		case "shade":
			attributes.shade = value;
			break;
		case "underline":
			attributes.underline = value;
			break;
		case "underlineColor":
			attributes.underlineColor = value;
			break;
		case "vertAlign":
			attributes.vertAlign = value;
			break;
		case "font":
			attributes.font = value;
			break;
		case "size": {
			const points = Number.parseFloat(value.replace(/pt$/i, ""));
			if (Number.isFinite(points)) {
				attributes.sizeHalfPoints = Math.round(points * 2);
			}
			break;
		}
	}
}

function stripQuotes(value: string): string {
	const quoted =
		(value.startsWith('"') && value.endsWith('"')) ||
		(value.startsWith("'") && value.endsWith("'"));
	return quoted ? value.slice(1, -1) : value;
}

/** Any marker character — cheap gate before the full tokenizer runs. */
const MARKER_CHAR = /[[\]{}]/;

/** `{++` `++}` `{--` `--}` `]{attrs}` `[`. Lone `]` and `}` are literal; a `[`
 * with no matching `]{…}` degrades back to literal. The attrs capture forbids
 * `[`/`]` so a `]{…}` can never swallow a second close marker (which would suck
 * the prose between two spans into the attr block and drop it) and so prose
 * brackets don't span across an unrelated later `]{…}`. */
const MARKER = /\{\+\+|\+\+\}|\{--|--\}|\]\{([^}[\]]*)\}|\[/g;

/** `#id`, `.class`, or `key=value` (value optionally quoted, may hold spaces). */
const ATTR = /([.#]?[\w-]+)(?:=("[^"]*"|'[^']*'|\S+))?/g;

type Marker = "ins" | "del" | "span";

type Token =
	| { kind: "literal"; value: string }
	| { kind: "atom"; node: MutableNode }
	| { kind: "open"; marker: Marker; raw: string }
	| { kind: "close"; marker: Marker; raw: string; attrs?: string };

type Frame = { marker: Marker; raw: string; children: MutableNode[] };

/** Loose view of an mdast node for in-place rewriting — we only touch `type`,
 * `value`, `children`, and (for spans) `attributes`. */
type MutableNode = {
	type: string;
	value?: string;
	children?: MutableNode[];
	attributes?: SpanAttributes;
};

/** CriticMarkup insertion `{++…++}` — a PARENT node holding the phrasing
 * content between the markers, so `{++**bold**++}` keeps its `strong` child.
 * The inline walker wraps the children in `<w:ins>` (tracking on) or splats
 * them as plain runs (tracking off). */
export interface CriticInsert {
	type: "criticInsert";
	children: PhrasingContent[];
}

/** CriticMarkup deletion `{--…--}` — mirror of {@link CriticInsert}; with
 * tracking off the children are dropped entirely (removed content). */
export interface CriticDelete {
	type: "criticDelete";
	children: PhrasingContent[];
}

/** Pandoc bracketed span `[text]{.class key="value"}` — carries the parsed
 * run-formatting attributes plus the phrasing content they apply to. The walker
 * overlays `attributes` onto the inherited `InlineFormat` and recurses. Not
 * exported: consumers reach it structurally via the `mdast` augmentation. */
interface BracketedSpan {
	type: "bracketedSpan";
	attributes: SpanAttributes;
	children: PhrasingContent[];
}

/** Run-formatting a `[…]{…}` span can carry. Mirrors the formatting subset of
 * `InlineFormat`/`TextRun`; the walker spreads it onto the inherited format.
 * Theme tint/shade stay raw hex strings so the value round-trips byte-exact. */
export interface SpanAttributes {
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	color?: string;
	colorTheme?: string;
	colorThemeTint?: string;
	colorThemeShade?: string;
	highlight?: string;
	shade?: string;
	underline?: string;
	underlineColor?: string;
	vertAlign?: string;
	smallCaps?: boolean;
	allCaps?: boolean;
	font?: string;
	sizeHalfPoints?: number;
}

declare module "mdast" {
	interface PhrasingContentMap {
		criticInsert: CriticInsert;
		criticDelete: CriticDelete;
		bracketedSpan: BracketedSpan;
	}
	interface RootContentMap {
		criticInsert: CriticInsert;
		criticDelete: CriticDelete;
		bracketedSpan: BracketedSpan;
	}
}
