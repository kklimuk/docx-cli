import type { FindView } from "../find";
import { isWrapperVisibleInView } from "../find/replace-span";
import {
	alternateContentBranch,
	collectTextBoxContents,
	isAlternateContent,
} from "../mc";
import {
	equationWidth,
	inlineMarkerWidth,
	isEquation,
	isRunBearingWrapper,
	wrapperContentNode,
	type XmlNode,
} from "../parser";

/** Every field in a story, read straight off the XML — no AST involved.
 *
 *  A COMPLEX field is three zero-width `<w:fldChar>` markers that form one
 *  unit — `begin`, an optional `separate`, `end` — with its instruction in the
 *  `<w:instrText>` runs between begin and separate and its cached result in the
 *  runs between separate and end. The markers can sit in different runs,
 *  different wrappers, even different paragraphs (a TOC opens in its first
 *  entry and closes after its last), and fields nest, so they're paired with a
 *  stack per story — the body, each text box's `<w:txbxContent>`, each
 *  header/footer/note/comment. A SIMPLE field (`<w:fldSimple>`) is one
 *  wrapper holding its result runs.
 *
 *  Each marker carries its position in the offset space every span walker
 *  uses (`collectRunSlots`/`sumVisibleTextLength` in find/replace-span.tsx):
 *  the ordinal of its paragraph in document order plus the paragraph-local
 *  offset in `view`. Markers the stack can't pair (an `end` with no open
 *  field, a `begin` never closed) are `strays` — a damaged file. */
export function parseFields(
	storyRoots: XmlNode[],
	view: FindView = "accepted",
): ParsedFields {
	const state: WalkState = {
		view,
		ordinal: 0,
		fields: [],
		strays: [],
		stack: [],
		stories: [],
		ordinals: new Map(),
	};
	for (const root of storyRoots) walkStory(state, root);
	// Text-box stories are queued as they're met (inside an anchor's run) and
	// parsed afterward, each with its own stack — a field never spans a story.
	while (state.stories.length > 0) {
		const story = state.stories.shift();
		if (story) walkStory(state, story);
	}
	return {
		fields: state.fields,
		strays: state.strays,
		ordinals: state.ordinals,
	};
}

function walkStory(state: WalkState, root: XmlNode): void {
	const outer = state.stack;
	state.stack = [];
	walkBlocks(state, root.children);
	for (const open of state.stack) {
		state.strays.push(open.begin);
	}
	state.stack = outer;
}

/** Block-level descent in document order: paragraphs, then anything that
 *  holds paragraphs (tables, rows, cells, content controls, custom XML). A
 *  block-level `<mc:AlternateContent>` is read through its ONE branch. */
function walkBlocks(state: WalkState, children: XmlNode[]): void {
	for (const child of children) {
		if (child.isText) continue;
		if (child.tag === "w:p") {
			walkParagraph(state, child);
			continue;
		}
		if (isAlternateContent(child)) {
			const branch = alternateContentBranch(child);
			if (branch) walkBlocks(state, branch.children);
			continue;
		}
		walkBlocks(state, child.children);
	}
}

function walkParagraph(state: WalkState, paragraph: XmlNode): void {
	const cursor: Cursor = {
		paragraph,
		ordinal: state.ordinal++,
		offset: 0,
	};
	state.ordinals.set(paragraph, cursor.ordinal);
	walkRunContainer(state, cursor, paragraph, { visible: true, measured: true });
}

/** Walk a run container (the paragraph, or a run-bearing wrapper's content
 *  node) the way the span walkers do: runs by `runTextLength`, equations by
 *  their token width, visible wrappers descended, hidden ones contributing no
 *  width. Anything else (a run-level content control, custom XML) is opaque to
 *  the walkers — zero width — but its markers still count for pairing. */
function walkRunContainer(
	state: WalkState,
	cursor: Cursor,
	container: XmlNode,
	scope: Scope,
): void {
	for (const child of container.children) {
		if (child.isText) continue;
		if (child.tag === "w:r") {
			walkRun(state, cursor, container, child, scope);
			continue;
		}
		if (isEquation(child)) {
			if (counts(scope)) cursor.offset += equationWidth(child, state.view);
			continue;
		}
		if (child.tag === "w:fldSimple") {
			walkSimpleField(state, cursor, container, child, scope);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			walkRunContainer(state, cursor, wrapperContentNode(child), {
				...scope,
				visible: scope.visible && isWrapperVisibleInView(child.tag, state.view),
			});
			continue;
		}
		if (child.tag === "w:pPr" || child.tag === "w:rPr") continue;
		walkRunContainer(state, cursor, child, { ...scope, measured: false });
	}
}

/** Whether text met in this scope advances the offset: visible in the view
 *  AND inside containers the span walkers measure. */
const counts = (scope: Scope): boolean => scope.visible && scope.measured;

function walkSimpleField(
	state: WalkState,
	cursor: Cursor,
	container: XmlNode,
	node: XmlNode,
	scope: Scope,
): void {
	const start = position(cursor);
	const resultStart = cursor.offset;
	walkRunContainer(state, cursor, node, scope);
	state.fields.push({
		kind: "simple",
		node,
		container,
		paragraph: cursor.paragraph,
		start,
		end: position(cursor),
		visible: scope.visible,
		instruction: (node.getAttribute("w:instr") ?? "").trim(),
		result: resultText(node),
		resultWidth: cursor.offset - resultStart,
	});
}

function walkRun(
	state: WalkState,
	cursor: Cursor,
	container: XmlNode,
	run: XmlNode,
	scope: Scope,
): void {
	const { visible } = scope;
	for (const child of run.children) {
		if (child.isText) continue;
		if (child.tag === "w:fldChar") {
			const marker: FieldMarker = {
				node: child,
				run,
				container,
				paragraph: cursor.paragraph,
				position: position(cursor),
				visible,
			};
			placeMarker(state, marker, child.getAttribute("w:fldCharType"));
			continue;
		}
		if (child.tag === "w:instrText" || child.tag === "w:delInstrText") {
			const open = state.stack.at(-1);
			if (open && !open.separate) {
				open.instructionRuns.push({
					node: child,
					run,
					container,
					paragraph: cursor.paragraph,
					position: position(cursor),
					visible,
				});
				open.instructionText += child.collectText();
			}
			continue;
		}
		if (child.tag === "w:t" || child.tag === "w:delText") {
			const text = child.collectText();
			for (const open of state.stack) {
				if (open.separate) open.resultText += text;
			}
			if (counts(scope)) cursor.offset += text.length;
			continue;
		}
		if (child.tag === "w:rPr") continue;
		for (const story of collectTextBoxContents(child)) {
			state.stories.push(story);
		}
		if (counts(scope)) cursor.offset += inlineMarkerWidth(child);
	}
}

/** Pair a `<w:fldChar>` with the open field it belongs to. */
function placeMarker(
	state: WalkState,
	marker: FieldMarker,
	type: string | undefined,
): void {
	if (type === "begin") {
		state.stack.push({
			begin: marker,
			instructionRuns: [],
			instructionText: "",
			resultText: "",
		});
		return;
	}
	const open = state.stack.at(-1);
	if (!open) {
		state.strays.push(marker);
		return;
	}
	if (type === "separate") {
		if (open.separate) state.strays.push(marker);
		else open.separate = marker;
		return;
	}
	if (type !== "end") return;
	state.stack.pop();
	state.fields.push({
		kind: "complex",
		begin: open.begin,
		...(open.separate ? { separate: open.separate } : {}),
		end: marker,
		instructionRuns: open.instructionRuns,
		instruction: open.instructionText.trim(),
		result: open.resultText,
	});
}

function position(cursor: Cursor): FieldPosition {
	return { ordinal: cursor.ordinal, offset: cursor.offset };
}

function resultText(node: XmlNode): string {
	let out = "";
	for (const child of node.children) {
		if (child.tag === "w:t" || child.tag === "w:delText") {
			out += child.collectText();
			continue;
		}
		if (!child.isText) out += resultText(child);
	}
	return out;
}

/** Whether any field markup sits anywhere under `nodes` — the cheap test that
 *  lets a mutator skip `parseFields` (a whole-document walk) for the common
 *  paragraph with no field at all. */
export function holdsFieldMarkup(nodes: XmlNode[]): boolean {
	return nodes.some(
		(node) =>
			FIELD_TAGS.has(node.tag) ||
			[...FIELD_TAGS].some((tag) => node.findDescendant(tag) !== undefined),
	);
}

const FIELD_TAGS = new Set([
	"w:fldChar",
	"w:fldSimple",
	"w:instrText",
	"w:delInstrText",
]);

/** Every paragraph that holds one of the field's markers (its fldChars and
 *  instruction runs, or the `<w:fldSimple>` itself) — the paragraphs a
 *  rebuild or removal can't touch without touching the field. */
export function fieldParagraphs(field: Field): Set<XmlNode> {
	if (field.kind === "simple") return new Set([field.paragraph]);
	return new Set(fieldMarkers(field).map((marker) => marker.paragraph));
}

/** A complex field's markers in document order: begin, instruction runs,
 *  separate, end. */
export function fieldMarkers(field: ComplexField): FieldMarker[] {
	return [
		field.begin,
		...field.instructionRuns,
		...(field.separate ? [field.separate] : []),
		field.end,
	];
}

/** The field's kind, as Word names it: the instruction's first word (`REF`,
 *  `PAGE`, `TOC`, `HYPERLINK`, …), upper-cased. */
export function fieldType(field: Field): string {
	return (field.instruction.split(/\s+/)[0] ?? "").toUpperCase() || "FIELD";
}

/** Whether any of the field's markers is visible in the view the positions
 *  were measured in — a field wholly inside a pending deletion has no width
 *  and nothing for an edit to carry. */
export function isFieldVisible(field: Field): boolean {
	if (field.kind === "simple") return field.visible;
	return fieldMarkers(field).some((marker) => marker.visible);
}

/** `a` before `b` in document order (`-1`), at the same point (`0`), after (`1`). */
export function comparePositions(a: FieldPosition, b: FieldPosition): number {
	if (a.ordinal !== b.ordinal) return a.ordinal < b.ordinal ? -1 : 1;
	if (a.offset !== b.offset) return a.offset < b.offset ? -1 : 1;
	return 0;
}

export type ParsedFields = {
	fields: Field[];
	strays: FieldMarker[];
	/** Each walked paragraph's document-order ordinal — how a caller turns a
	 *  `pN:S-E` span into the positions fields are measured in. */
	ordinals: Map<XmlNode, number>;
};

export type Field = ComplexField | SimpleField;

export type ComplexField = {
	kind: "complex";
	begin: FieldMarker;
	separate?: FieldMarker;
	end: FieldMarker;
	instructionRuns: FieldMarker[];
	instruction: string;
	/** The cached result's text (every view — for naming the field). */
	result: string;
};

export type SimpleField = {
	kind: "simple";
	node: XmlNode;
	container: XmlNode;
	paragraph: XmlNode;
	start: FieldPosition;
	end: FieldPosition;
	visible: boolean;
	instruction: string;
	result: string;
	resultWidth: number;
};

/** A zero-width field marker (`<w:fldChar>` or `<w:instrText>`), with the run
 *  that holds it and the node whose child list holds that run. */
export type FieldMarker = {
	node: XmlNode;
	run: XmlNode;
	container: XmlNode;
	paragraph: XmlNode;
	position: FieldPosition;
	/** Inside no view-hidden wrapper. */
	visible: boolean;
};

export type FieldPosition = { ordinal: number; offset: number };

type OpenField = {
	begin: FieldMarker;
	separate?: FieldMarker;
	instructionRuns: FieldMarker[];
	instructionText: string;
	resultText: string;
};

type Scope = { visible: boolean; measured: boolean };

type Cursor = { paragraph: XmlNode; ordinal: number; offset: number };

type WalkState = {
	view: FindView;
	ordinal: number;
	fields: Field[];
	strays: FieldMarker[];
	stack: OpenField[];
	stories: XmlNode[];
	ordinals: Map<XmlNode, number>;
};
