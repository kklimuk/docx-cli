import type { Nodes, Root } from "mdast";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import remarkParse from "remark-parse";
import { unified } from "unified";

// A node's source-offset span. Every mdast node carries this, but the union also
// includes our own CriticMarkup nodes (which don't), so we read it structurally
// rather than widening to mdast's `Nodes`.
type Positioned = {
	position?: { start?: { offset?: number }; end?: { offset?: number } };
};

// The plugin list MUST mirror `parseToMdast` in import.tsx — the mask classifies
// text-vs-construct exactly as the importer's parser does, so what `read` leaves
// unescaped is precisely what the importer reads back as literal text. We parse
// only (no `runSync`): the inline-surgery transformer re-pairs our OWN emitted
// HTML wrappers and never reclassifies raw run content, so it's irrelevant here.
const rawParser = unified().use(remarkParse).use(remarkGfm).use(remarkMath);

// Every char CommonMark lets a backslash escape. A `\` before one of these is an
// escape; escaping anything else (a letter, a space) is a no-op the parser drops.
const ASCII_PUNCTUATION = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~";

// Anything that could BEGIN a construct. None present → no parse needed.
const MAYBE_CONSTRUCT = /[\\`*_~[\]<>&$]/;

/**
 * A per-character escape mask over `content`: `mask[i] === true` means
 * `content[i]` must be backslash-escaped so `read`'s Markdown round-trips through
 * the importer verbatim. The authority is the importer's OWN parser, not a set of
 * hand-rolled rules: we parse the content and escape exactly the punctuation the
 * parser would CONSUME into a construct (emphasis, math, a link, an HTML tag, a
 * footnote, …), leaving every character it keeps as literal text untouched. That
 * is why a `[ x ]` checkbox or a lone `$5` stays clean while a paired `$…$` or a
 * `[label](url)` gets escaped — there is no pairing/flanking/link logic here to
 * drift from what remark actually does.
 *
 * Two things the parser HIDES by decoding (so the position scan alone can't see
 * them) are marked explicitly: a character reference (`&amp;` → `&`) escapes its
 * `&`, and a backslash escape (`\*` → `*`) escapes its `\`.
 *
 * `hasEquation` escapes every `$`: an inline/display equation elsewhere in the
 * same scope emits a real `$`/`$$` that would pair with an otherwise-lone `$`,
 * and that equation's `$` isn't part of `content` for the parser to see.
 */
export function inlineEscapeMask(
	content: string,
	hasEquation = false,
): boolean[] {
	const mask: boolean[] = new Array(content.length).fill(false);
	// Fast path: prose with no construct-starting character can't parse to
	// anything but literal text, so skip the parse (the overwhelmingly common case).
	if (!hasEquation && !MAYBE_CONSTRUCT.test(content)) return mask;

	const literal = literalTextMask(content);
	for (let i = 0; i < content.length; i++) {
		const char = content[i];
		if (char === undefined) continue;
		// A punctuation char the parser did NOT keep as literal text is a construct
		// delimiter (a `*`, `$`, `[`, `<`, `` ` ``, …) — escape it. Non-punctuation
		// is never escapable, so a link's URL letters are left alone (escaping the
		// bracket already neutralizes the link).
		if (!literal[i] && ASCII_PUNCTUATION.includes(char)) mask[i] = true;
	}
	markEntityStarts(content, mask);
	markBackslashEscapes(content, mask);
	if (hasEquation) {
		for (let i = 0; i < content.length; i++) {
			if (content[i] === "$") mask[i] = true;
		}
	}
	return mask;
}

/** Mark every offset the parser reproduces as literal text — i.e. NOT escapable.
 *
 * `text` nodes mark their whole span. The subtlety is the two constructs whose
 * content is a wide `value` string rather than a child `text` node: `inlineMath`
 * (`$…$`) and `inlineCode` (`` `…` ``). Every OTHER construct (link, emphasis,
 * image, html) keeps its interior as a child `text` node, so the span scan
 * already escapes only its delimiters. For these two we mark the INTERIOR literal
 * and leave just the boundary `$`/backtick non-literal: escaping the boundary
 * breaks the construct, so interior punctuation (a `[placeholder]` that fell
 * between two paired `$` across a table cell) needn't be touched.
 *
 * A text node whose `value` was decoded from its source span (an entity or
 * backslash escape) is still marked literal here — `markEntityStarts` /
 * `markBackslashEscapes` re-flag the one trigger character so the rest stays clean. */
function literalTextMask(content: string): boolean[] {
	const literal: boolean[] = new Array(content.length).fill(false);
	const tree = rawParser.parse(content) as Root;
	visit(tree, (node) => {
		if (node.type === "text") {
			markSpan(content, literal, node, 0, 0);
		} else if (node.type === "inlineMath") {
			markSpan(content, literal, node, 1, 1); // exclude the boundary `$`
		} else if (node.type === "inlineCode") {
			const [lead, trail] = backtickRunLengths(content, node);
			markSpan(content, literal, node, lead, trail);
		}
	});
	return literal;
}

/** Mark `[start + lead, end - trail)` of `node`'s source span literal. */
function markSpan(
	content: string,
	literal: boolean[],
	node: Positioned,
	lead: number,
	trail: number,
): void {
	const start = node.position?.start?.offset;
	const end = node.position?.end?.offset;
	if (start === undefined || end === undefined) return;
	const from = Math.max(0, start + lead);
	const to = Math.min(content.length, end - trail);
	for (let i = from; i < to; i++) literal[i] = true;
}

/** The length of the opening and closing backtick fences of an `inlineCode` node
 * (`` `x` `` → [1, 1], `` ``x`` `` → [2, 2]) so its code interior is marked literal. */
function backtickRunLengths(
	content: string,
	node: Positioned,
): [number, number] {
	const start = node.position?.start?.offset ?? 0;
	const end = node.position?.end?.offset ?? 0;
	let lead = 0;
	while (start + lead < end && content[start + lead] === "`") lead++;
	let trail = 0;
	while (end - trail > start + lead && content[end - trail - 1] === "`")
		trail++;
	return [lead, trail];
}

function visit(node: Nodes, fn: (node: Nodes) => void): void {
	fn(node);
	const children = (node as { children?: Nodes[] }).children;
	if (Array.isArray(children)) {
		for (const child of children) visit(child, fn);
	}
}

const ENTITY_REFERENCE = /&(?:#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z][A-Za-z0-9]*);/g;

function markEntityStarts(content: string, mask: boolean[]): void {
	for (const match of content.matchAll(ENTITY_REFERENCE)) {
		if (match.index !== undefined) mask[match.index] = true;
	}
}

function markBackslashEscapes(content: string, mask: boolean[]): void {
	for (let i = 0; i < content.length; i++) {
		if (content[i] !== "\\") continue;
		const next = content[i + 1];
		// A `\` before punctuation (or at the very end, where it could merge into a
		// following `<mark>`/`**` the emitter appends) is itself an escape — preserve
		// the literal backslash by escaping it. `C:\Users` (`\` before a letter) is
		// inert, so it stays clean.
		if (next === undefined || ASCII_PUNCTUATION.includes(next)) mask[i] = true;
	}
}
