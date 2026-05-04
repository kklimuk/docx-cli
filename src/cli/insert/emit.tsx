import type { Run, TextRun } from "@core";
import { w } from "@core/jsx";
import type { NullableXmlNode, XmlNode } from "@core/parser";

export type ParagraphOptions = {
	style?: string;
	alignment?: "left" | "center" | "right" | "justify";
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

export function Paragraph(props: ParagraphProps): XmlNode {
	const { style, alignment } = props;
	const runs: Run[] =
		"runs" in props && props.runs
			? props.runs
			: [textRunFromProps(props as ParagraphProps & { text: string })];
	return (
		<w.p>
			<ParagraphProperties options={{ style, alignment }} />
			{runs.map((run) => (
				<RunElement run={run} />
			))}
		</w.p>
	);
}

export function RunElement({ run }: { run: Run }): XmlNode {
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
	throw new Error(
		`Cannot emit run of type "${run.type}" — image emission lives in the images command.`,
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
	if (!options.style && !options.alignment) return null;
	return (
		<w.pPr>
			{options.style && <w.pStyle w-val={options.style} />}
			{options.alignment && <w.jc w-val={options.alignment} />}
		</w.pPr>
	);
}

function textRunFromProps(props: ParagraphProps & { text: string }): TextRun {
	const {
		style: _style,
		alignment: _alignment,
		text,
		runs: _runs,
		...formatting
	} = props;
	return { type: "text", text, ...formatting };
}
