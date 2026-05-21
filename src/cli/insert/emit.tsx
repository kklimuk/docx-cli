import type { Run, TextRun } from "@core";
import { w } from "@core/jsx";
import type { NullableXmlNode, XmlNode } from "@core/parser";

export type ParagraphOptions = {
	style?: string;
	alignment?: "left" | "center" | "right" | "justify";
	list?: { level: number; numId: number };
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

export type ParagraphProps = ParagraphOptions &
	(
		| ({ text: string; runs?: never } & Partial<EmittableTextFormatting>)
		| { runs: Run[]; text?: never }
	);

/** A paragraph rendered as a horizontal rule — empty body with a bottom border.
 * Word renders this as a thin line spanning the page width. Intended for the
 * S8 markdown walker's `---` rules; not yet wired into any CLI verb. */
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

export function Paragraph(props: ParagraphProps): XmlNode {
	const { style, alignment, list } = props;
	const runs: Run[] =
		"runs" in props && props.runs
			? props.runs
			: [textRunFromProps(props as ParagraphProps & { text: string })];
	return (
		<w.p>
			<ParagraphProperties options={{ style, alignment, list }} />
			{runs.map((run) => (
				<RunElement run={run} />
			))}
		</w.p>
	);
}

export type ListParagraphProps = {
	numId: number;
	level: number;
	alignment?: "left" | "center" | "right" | "justify";
} & ({ text: string } | { runs: Run[] });

/** A paragraph that belongs to a numbered or bulleted list. Sets pStyle to
 * "ListParagraph" (the canonical Word style for list items — caller is
 * responsible for ensuring it via `ensureStyle(view, "ListParagraph")`) and
 * emits the `<w:numPr>` reference. The numId must come from
 * `allocateNum(view, kind)` in [src/core/numbering.tsx]. */
export function ListParagraph(props: ListParagraphProps): XmlNode {
	const baseProps = {
		style: "ListParagraph",
		alignment: props.alignment,
		list: { level: props.level, numId: props.numId },
	} as const;
	if ("runs" in props) {
		return <Paragraph {...baseProps} runs={props.runs} />;
	}
	return <Paragraph {...baseProps} text={props.text} />;
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
] as const satisfies readonly (keyof TextRun)[];

function RunProperties({ run }: { run: TextRun }): NullableXmlNode {
	const isEmpty = FORMATTING_KEYS.every((key) => run[key] == null);
	if (isEmpty) return null;
	return (
		<w.rPr>
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

function textRunFromProps(props: ParagraphProps & { text: string }): TextRun {
	const {
		style: _style,
		alignment: _alignment,
		list: _list,
		text,
		runs: _runs,
		...formatting
	} = props;
	return { type: "text", text, ...formatting };
}
