import type { Root } from "mdast";
import type { Data, Plugin } from "unified";

/** Remark plugin: footnote-shaped READ ANNOTATIONS are hints, not notes.
 *
 * `read --current` labels every tracked change `{++…++}[^tcN]` and appends a
 * `[^tcN]: insertion by …` definition; `read --comments` does the same with
 * `[^cN]`. Real notes always read as `[^fnN]`/`[^enN]`, so those two shapes are
 * names only a read-time annotation uses — and, like every `<!-- … -->` hint,
 * they are DROPPED on import (root CLAUDE.md: "Comments are never anything but
 * hints"). Without this, a pasted `--current` line turned each label into a real
 * footnote reading "deletion by Reviewer …", and a label with no definition
 * (a single line handed to `edit`) stayed literal and swallowed the CriticMarkup
 * after it (`]{++45++}` read as a Pandoc-span close).
 *
 * Two halves, because GFM only parses `[^x]` as a reference when a `[^x]:`
 * definition exists:
 *   - a micromark TEXT construct, tried before GFM's own footnote call, claims
 *     every `[^tcN]`/`[^cN]` reference — defined or not — as a `readHintLabel`
 *     node. Being a real parser construct, it is escape- and code-aware for free:
 *     `\[^tc0]` and `` `[^tc0]` `` stay literal. `escape.ts` registers the plugin
 *     too, so `read` escapes a literal `[^tc0]` in document text and it
 *     round-trips instead of vanishing;
 *   - a transformer (ordered before `remarkInlineSurgery`) removes those nodes and
 *     every hint-named `footnoteDefinition`, re-merging the text around a removed
 *     label so CriticMarkup next to it tokenizes exactly as if it were never there.
 *
 * Only the exact shapes `tc<digits>` / `c<digits>` match (case-insensitive, as
 * GFM normalizes footnote identifiers); a user's `[^note]` or `[^fn1]` is a real
 * footnote. */
export const remarkReadHintLabels: Plugin<[], Root> = function () {
	const data = this.data();
	data.micromarkExtensions = [
		...(data.micromarkExtensions ?? []),
		hintLabelSyntax,
	];
	data.fromMarkdownExtensions = [
		...(data.fromMarkdownExtensions ?? []),
		hintLabelFromMarkdown,
	];
	return (tree) => {
		dropHintNodes(tree as unknown as MutableNode);
	};
};

/** Is a (GFM-normalized, lowercased) footnote identifier a read annotation? */
const HINT_IDENTIFIER = /^(?:tc|c)\d+$/;

const HINT_TOKEN = "readHintLabel";

const CODE = {
	leftBracket: 91,
	rightBracket: 93,
	caret: 94,
	zero: 48,
	nine: 57,
	upperC: 67,
	upperT: 84,
	lowerC: 99,
	lowerT: 116,
} as const;

/** Tokenize `[^tc<digits>]` / `[^c<digits>]` at a `[`. Anything else is `nok`, so
 * the `[` falls through to GFM's footnote call / link label untouched. */
function tokenizeHintLabel(
	effects: Effects,
	ok: State,
	nok: State,
): State | undefined {
	let digits = 0;
	return start;

	function start(code: Code): State | undefined {
		effects.enter(HINT_TOKEN);
		effects.consume(code);
		return caret;
	}

	function caret(code: Code): State | undefined {
		if (code !== CODE.caret) return nok(code);
		effects.consume(code);
		return prefix;
	}

	function prefix(code: Code): State | undefined {
		if (code === CODE.lowerT || code === CODE.upperT) {
			effects.consume(code);
			return afterT;
		}
		if (code === CODE.lowerC || code === CODE.upperC) {
			effects.consume(code);
			return number;
		}
		return nok(code);
	}

	function afterT(code: Code): State | undefined {
		if (code !== CODE.lowerC && code !== CODE.upperC) return nok(code);
		effects.consume(code);
		return number;
	}

	function number(code: Code): State | undefined {
		if (code !== null && code >= CODE.zero && code <= CODE.nine) {
			digits++;
			effects.consume(code);
			return number;
		}
		if (code !== CODE.rightBracket || digits === 0) return nok(code);
		effects.consume(code);
		effects.exit(HINT_TOKEN);
		return ok;
	}
}

// The micromark/from-markdown extension types come from packages we don't
// depend on directly; derive them from remark-parse's `Data` augmentation and
// state the tokenizer's tiny structural surface locally.
type MicromarkExtension = NonNullable<Data["micromarkExtensions"]>[number];
type FromMarkdownExtension = Exclude<
	NonNullable<Data["fromMarkdownExtensions"]>[number],
	unknown[]
>;
type Code = number | null;
type State = (code: Code) => State | undefined;
type Effects = {
	enter(type: string): unknown;
	exit(type: string): unknown;
	consume(code: Code): void;
};

const hintLabelSyntax = {
	text: {
		[CODE.leftBracket]: {
			name: HINT_TOKEN,
			add: "before",
			tokenize: tokenizeHintLabel,
		},
	},
} as unknown as MicromarkExtension;

const hintLabelFromMarkdown = {
	enter: {
		[HINT_TOKEN](this: CompileContext, token: unknown) {
			this.enter({ type: HINT_TOKEN }, token);
		},
	},
	exit: {
		[HINT_TOKEN](this: CompileContext, token: unknown) {
			this.exit(token);
		},
	},
} as unknown as FromMarkdownExtension;

type CompileContext = {
	enter(node: { type: string }, token: unknown): void;
	exit(token: unknown): void;
};

/** Remove every hint label + hint-named footnote definition (and, for a
 * case-variant reference GFM claimed first, its `footnoteReference`), folding
 * the text on either side of a removed node back into one `text` node. */
function dropHintNodes(node: MutableNode): void {
	if (!Array.isArray(node.children)) return;
	const kept: MutableNode[] = [];
	for (const child of node.children) {
		if (isHintNode(child)) continue;
		const previous = kept[kept.length - 1];
		if (child.type === "text" && previous?.type === "text") {
			previous.value = (previous.value ?? "") + (child.value ?? "");
			if (previous.position && child.position) {
				previous.position = {
					start: previous.position.start,
					end: child.position.end,
				};
			}
			continue;
		}
		kept.push(child);
	}
	node.children = kept;
	for (const child of kept) dropHintNodes(child);
}

function isHintNode(node: MutableNode): boolean {
	if (node.type === HINT_TOKEN) return true;
	if (node.type !== "footnoteDefinition" && node.type !== "footnoteReference")
		return false;
	return HINT_IDENTIFIER.test(node.identifier ?? "");
}

type MutableNode = {
	type: string;
	value?: string;
	identifier?: string;
	children?: MutableNode[];
	position?: { start: unknown; end: unknown };
};
