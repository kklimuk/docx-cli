import type { Run, TextRun } from "./ast/types";
import { w } from "./jsx";
import { type NullableXmlNode, XmlNode } from "./parser";
import { TaskCheckbox } from "./task-list";

/** A `<w:p>`. Pass `runs` for full control, or `text` (+ optional run-level
 * formatting) for a single-run paragraph. `runs` wins if both are given; with
 * neither, an empty paragraph is emitted.
 *
 * When `taskState` is set, prepends a `<w:sdt><w14:checkbox/></w:sdt>` content
 * control plus a space run before the user content — the shape Word, Pandoc,
 * and LibreOffice all emit for GFM task list items. The detection mirror in
 * `core/ast/read.ts::detectLeadingCheckbox` strips this prefix back out on
 * read, so a paragraph round-trips losslessly. */
export function Paragraph({
	style,
	alignment,
	list,
	taskState,
	text,
	runs,
	...formatting
}: ParagraphOptions &
	Partial<EmittableTextFormatting> & {
		text?: string;
		runs?: Run[];
	}): XmlNode {
	const resolvedRuns: Run[] = runs ?? [
		{ type: "text", text: text ?? "", ...formatting },
	];
	return (
		<w.p>
			<ParagraphProperties options={{ style, alignment, list }} />
			{taskState && <TaskCheckbox checked={taskState === "checked"} />}
			{taskState && (
				<w.r>
					<w.t {...{ "xml:space": "preserve" }}> </w.t>
				</w.r>
			)}
			{resolvedRuns.map((run) => (
				<RunElement run={run} />
			))}
		</w.p>
	);
}

export type ParagraphOptions = {
	style?: string;
	alignment?: "left" | "center" | "right" | "justify";
	list?: { level: number; numId: number };
	taskState?: "checked" | "unchecked";
};

type EmittableTextFormatting = Pick<
	TextRun,
	| "color"
	| "highlight"
	| "bold"
	| "italic"
	| "underline"
	| "strike"
	| "font"
	| "sizeHalfPoints"
>;

/** A paragraph that belongs to a numbered or bulleted list. Sets pStyle to
 * "ListParagraph" (the canonical Word style for list items — caller is
 * responsible for ensuring it via `view.ensureStyles().ensureStyle("ListParagraph")`) and
 * emits the `<w:numPr>` reference. `numId` must come from
 * `view.ensureNumbering().allocate(kind)` in [src/core/numbering.tsx]. */
export function ListParagraph({
	numId,
	level,
	alignment,
	taskState,
	text,
	runs,
}: {
	numId: number;
	level: number;
	alignment?: "left" | "center" | "right" | "justify";
	taskState?: "checked" | "unchecked";
	text?: string;
	runs?: Run[];
}): XmlNode {
	return (
		<Paragraph
			style="ListParagraph"
			alignment={alignment}
			list={{ level, numId }}
			taskState={taskState}
			text={text}
			runs={runs}
		/>
	);
}

/** A GFM task list item — a `<ListParagraph>` whose body is prefixed with a
 * Word checkbox content control. Sugar over `ListParagraph` with `taskState`
 * set; the SDT prefix is materialized inside `Paragraph` itself.
 *
 * @public Staged for the S8 markdown walker's `- [ ]` / `- [x]` task list
 * items; CLI authoring today goes through `insert --task` / `edit --task`
 * which build the paragraph via `Paragraph({ list, taskState })` directly. */
export function TaskListItem({
	numId,
	level,
	checked,
	text,
	runs,
}: {
	numId: number;
	level: number;
	checked: boolean;
	text?: string;
	runs?: Run[];
}): XmlNode {
	return (
		<ListParagraph
			numId={numId}
			level={level}
			taskState={checked ? "checked" : "unchecked"}
			text={text}
			runs={runs}
		/>
	);
}

/** Emits a run as fresh OOXML. The fresh-emission path supports text/break/tab
 * — the run types an agent can plausibly hand-author via `--runs '[...]'`.
 * Round-tripped runs that we surface but can't re-emit (image, equation,
 * footnoteRef, chart) are silently dropped: their underlying XML lives in the
 * source document and is never re-serialized through this path, so a
 * `read | jq | edit` workflow degrades gracefully instead of crashing. */
export function RunElement({ run }: { run: Run }): NullableXmlNode {
	if (run.type === "text") return <TextRunElement run={run} />;
	if (run.type === "break") {
		return (
			<w.r>
				<w.br w-type={run.kind === "line" ? undefined : run.kind} />
			</w.r>
		);
	}
	if (run.type === "tab") {
		return (
			<w.r>
				<w.tab />
			</w.r>
		);
	}
	return null;
}

/** Mutate a paragraph's children list in place to apply `--style` /
 *  `--alignment` to its `<w:pPr>`. Creates a pPr if absent. Used by the
 *  tracked-edit helpers in `core/track-changes/` and any caller that wants to
 *  swap paragraph properties without rebuilding the whole paragraph. The
 *  `rebuilt` list is the paragraph's `children` array — usually already
 *  filtered to non-run nodes by the caller. */
export function applyParagraphOptionsInPlace(
	rebuilt: XmlNode[],
	options: ParagraphOptions,
): void {
	if (!options.style && !options.alignment) return;
	let pPr = rebuilt.find((child) => child.tag === "w:pPr");
	if (!pPr) {
		pPr = new XmlNode("w:pPr");
		rebuilt.unshift(pPr);
	}
	if (options.style) {
		const existingStyle = pPr.findChild("w:pStyle");
		if (existingStyle) {
			existingStyle.setAttribute("w:val", options.style);
		} else {
			const styleNode = new XmlNode("w:pStyle", { "w:val": options.style });
			pPr.children.unshift(styleNode);
		}
	}
	if (options.alignment) {
		const existingJc = pPr.findChild("w:jc");
		if (existingJc) {
			existingJc.setAttribute("w:val", options.alignment);
		} else {
			pPr.children.push(new XmlNode("w:jc", { "w:val": options.alignment }));
		}
	}
}

/** A paragraph rendered as a horizontal rule — empty body with a bottom border.
 * Word renders this as a thin line spanning the page width.
 *
 * @public Staged for the S8 markdown walker's `---` thematic breaks; not yet
 * wired into a CLI verb, so it has no internal caller today. */
export function HorizontalRule(): XmlNode {
	return (
		<w.p>
			<w.pPr>
				<w.pBdr>
					<w.bottom w-val="single" w-sz="6" w-space="1" w-color="auto" />
				</w.pBdr>
			</w.pPr>
		</w.p>
	);
}

function TextRunElement({ run }: { run: TextRun }): XmlNode {
	return (
		<w.r>
			<RunProperties run={run} />
			<w.t {...{ "xml:space": "preserve" }}>{run.text}</w.t>
		</w.r>
	);
}

const FORMATTING_KEYS = [
	"color",
	"highlight",
	"bold",
	"italic",
	"underline",
	"strike",
	"font",
	"sizeHalfPoints",
	"runStyle",
] as const satisfies readonly (keyof TextRun)[];

function RunProperties({ run }: { run: TextRun }): NullableXmlNode {
	const isEmpty = FORMATTING_KEYS.every((key) => run[key] == null);
	if (isEmpty) return null;
	// `<w:rStyle>` comes FIRST in <w:rPr> per ECMA-376 §17.3.2.28 (CT_RPr child
	// order: rStyle → rFonts → b → i → strike → color → sz → highlight → u → …).
	// Word and LibreOffice both tolerate other orderings, but matching the
	// schema keeps validators happy and round-trips minimal-diff against Word.
	return (
		<w.rPr>
			{run.runStyle && <w.rStyle w-val={run.runStyle} />}
			{run.color && <w.color w-val={run.color} />}
			{run.highlight && <w.highlight w-val={run.highlight} />}
			{run.bold && <w.b />}
			{run.italic && <w.i />}
			{run.underline && <w.u w-val={run.underline} />}
			{run.strike && <w.strike />}
			{run.font && <w.rFonts w-ascii={run.font} w-hAnsi={run.font} />}
			{run.sizeHalfPoints !== undefined && (
				<w.sz w-val={String(run.sizeHalfPoints)} />
			)}
		</w.rPr>
	);
}

function ParagraphProperties({
	options,
}: {
	options: ParagraphOptions;
}): NullableXmlNode {
	if (!options.style && !options.alignment && !options.list) return null;
	// Schema order (CT_PPrBase, ECMA-376 §17.3.1.26): pStyle → numPr → jc.
	return (
		<w.pPr>
			{options.style && <w.pStyle w-val={options.style} />}
			{options.list && (
				<w.numPr>
					<w.ilvl w-val={String(options.list.level)} />
					<w.numId w-val={String(options.list.numId)} />
				</w.numPr>
			)}
			{options.alignment && <w.jc w-val={options.alignment} />}
		</w.pPr>
	);
}
