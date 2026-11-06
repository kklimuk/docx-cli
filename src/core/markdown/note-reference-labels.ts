import type { Root } from "mdast";
import type { Data, Plugin } from "unified";
import type { NoteKind } from "../notes";

/** Remark plugin: a `[^fnN]` / `[^enN]` copied from `read` addresses the
 * document's EXISTING note N.
 *
 * `read` prints every real note reference as `[^fnN]`/`[^enN]` inline and the
 * note's body as a `[^fnN]: …` definition at the end of the read. An agent
 * following "copy a line from the read view, change it, pass it back" copies
 * the line WITHOUT that definition — and GFM only parses `[^x]` as a reference
 * when a `[^x]:` definition exists, so the label used to land as the literal
 * characters `[^fn1]` (and the note itself was lost on the next accept).
 *
 * Two halves, like `remarkReadHintLabels`:
 *   - a micromark TEXT construct, tried before GFM's own footnote call, claims
 *     every `[^fn<digits>]` / `[^en<digits>]`. Being a parser construct it is
 *     escape- and code-aware for free (`\[^fn1]` and `` `[^fn1]` `` stay
 *     literal), and `escape.ts` registers the plugin too, so `read` escapes a
 *     LITERAL `[^fn1]` in document text and it round-trips as text;
 *   - a transformer resolves each claimed label: when the same source DEFINES
 *     it (`[^fn1]: body` — a whole read pasted into `create`), it goes back to
 *     being an ordinary GFM `footnoteReference`, so the definition mints a note
 *     exactly as before; otherwise it becomes an `existingNoteReference` the
 *     inline walker resolves against the target document's notes part. */
export const remarkNoteReferenceLabels: Plugin<[], Root> = function () {
	const data = this.data();
	data.micromarkExtensions = [
		...(data.micromarkExtensions ?? []),
		noteLabelSyntax,
	];
	data.fromMarkdownExtensions = [
		...(data.fromMarkdownExtensions ?? []),
		noteLabelFromMarkdown,
	];
	return (tree) => {
		const root = tree as unknown as MutableNode;
		resolveNoteLabels(root, definedIdentifiers(root));
	};
};

/** An inline reference to note `numericId` of `kind` that already exists in the
 * target document (`[^fn1]` with no `[^fn1]:` definition in the source). */
export type ExistingNoteReference = {
	type: "existingNoteReference";
	kind: NoteKind;
	numericId: string;
	/** The label as written (`fn1`), for error messages. */
	label: string;
};

const NOTE_LABEL_TOKEN = "noteReferenceLabel";

const CODE = {
	leftBracket: 91,
	rightBracket: 93,
	caret: 94,
	zero: 48,
	nine: 57,
	upperE: 69,
	upperF: 70,
	upperN: 78,
	lowerE: 101,
	lowerF: 102,
	lowerN: 110,
} as const;

/** Tokenize `[^fn<digits>]` / `[^en<digits>]` at a `[`. Anything else is `nok`,
 * so the `[` falls through to the hint-label construct, GFM's footnote call, or
 * a link label untouched. */
function tokenizeNoteLabel(
	effects: Effects,
	ok: State,
	nok: State,
): State | undefined {
	let digits = 0;
	return start;

	function start(code: Code): State | undefined {
		effects.enter(NOTE_LABEL_TOKEN);
		effects.consume(code);
		return caret;
	}

	function caret(code: Code): State | undefined {
		if (code !== CODE.caret) return nok(code);
		effects.consume(code);
		return prefix;
	}

	function prefix(code: Code): State | undefined {
		const isPrefix =
			code === CODE.lowerF ||
			code === CODE.upperF ||
			code === CODE.lowerE ||
			code === CODE.upperE;
		if (!isPrefix) return nok(code);
		effects.consume(code);
		return afterPrefix;
	}

	function afterPrefix(code: Code): State | undefined {
		if (code !== CODE.lowerN && code !== CODE.upperN) return nok(code);
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
		effects.exit(NOTE_LABEL_TOKEN);
		return ok;
	}
}

/** Every footnote definition identifier in the source, lowercased the way GFM
 * normalizes them. */
function definedIdentifiers(root: MutableNode): Set<string> {
	const out = new Set<string>();
	visit(root, (node) => {
		if (node.type === "footnoteDefinition" && node.identifier) {
			out.add(node.identifier.toLowerCase());
		}
	});
	return out;
}

function resolveNoteLabels(node: MutableNode, defined: Set<string>): void {
	if (!Array.isArray(node.children)) return;
	node.children = node.children.map((child) =>
		child.type === NOTE_LABEL_TOKEN ? resolveNoteLabel(child, defined) : child,
	);
	for (const child of node.children) resolveNoteLabels(child, defined);
}

/** `[^fn1]` → a GFM reference when the source defines it, else a reference to
 * the document's existing note. */
function resolveNoteLabel(
	node: MutableNode,
	defined: Set<string>,
): MutableNode {
	const label = (node.label ?? "").toLowerCase();
	if (defined.has(label)) {
		return {
			type: "footnoteReference",
			identifier: label,
			label,
			position: node.position,
		};
	}
	const reference: ExistingNoteReference = {
		type: "existingNoteReference",
		kind: label.startsWith("fn") ? "footnote" : "endnote",
		numericId: String(Number(label.slice(2))),
		label,
	};
	return { ...reference, position: node.position };
}

function visit(node: MutableNode, visitor: (node: MutableNode) => void): void {
	visitor(node);
	if (!Array.isArray(node.children)) return;
	for (const child of node.children) visit(child, visitor);
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

const noteLabelSyntax = {
	text: {
		[CODE.leftBracket]: {
			name: NOTE_LABEL_TOKEN,
			add: "before",
			tokenize: tokenizeNoteLabel,
		},
	},
} as unknown as MicromarkExtension;

const noteLabelFromMarkdown = {
	enter: {
		[NOTE_LABEL_TOKEN](this: CompileContext, token: unknown) {
			this.enter({ type: NOTE_LABEL_TOKEN }, token);
		},
	},
	exit: {
		[NOTE_LABEL_TOKEN](this: CompileContext, token: unknown) {
			const node = this.stack[this.stack.length - 1];
			// `[^fn1]` → `fn1`.
			if (node) node.label = this.sliceSerialize(token).slice(2, -1);
			this.exit(token);
		},
	},
} as unknown as FromMarkdownExtension;

type CompileContext = {
	stack: MutableNode[];
	enter(node: { type: string }, token: unknown): void;
	exit(token: unknown): void;
	sliceSerialize(token: unknown): string;
};

type MutableNode = {
	type: string;
	label?: string;
	identifier?: string;
	children?: MutableNode[];
	position?: unknown;
};

declare module "mdast" {
	interface PhrasingContentMap {
		existingNoteReference: ExistingNoteReference;
	}
	interface RootContentMap {
		existingNoteReference: ExistingNoteReference;
	}
}
