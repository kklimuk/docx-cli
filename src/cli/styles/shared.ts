import type { RunFormat } from "@core";
import type { ParagraphOptions } from "@core/blocks";
import {
	hasRunFormatFlags,
	parseRunFormat,
	parseSpacingIndentFlags,
} from "../parse-helpers";

/** The run + paragraph formatting flags shared by `styles set` and `styles create`
 *  — the SAME vocabulary `edit` uses, so an agent's muscle memory transfers from a
 *  one-off span edit to editing the style definition. `--tabs` is deliberately
 *  absent: a tab stop needs the section content width to resolve a right-edge stop,
 *  which a style definition has no context for. */
export const STYLE_FORMAT_FLAGS = {
	bold: { type: "boolean" },
	italic: { type: "boolean" },
	underline: { type: "boolean" },
	strike: { type: "boolean" },
	caps: { type: "boolean" },
	smallcaps: { type: "boolean" },
	superscript: { type: "boolean" },
	subscript: { type: "boolean" },
	color: { type: "string" },
	font: { type: "string" },
	size: { type: "string" },
	highlight: { type: "string" },
	shade: { type: "string" },
	alignment: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
} as const;

/** The formatting block of a `styles set`/`create` invocation. */
export type StyleFormatting = {
	runFormat?: RunFormat;
	paragraphOptions?: ParagraphOptions;
};

type StyleFormattingError = { error: string; hint?: string };

const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);

/** Documentation/error blurb naming the paragraph-only flags. */
const PARAGRAPH_FLAG_LABEL =
	"--alignment/--space-before/--space-after/--line-spacing/--indent-left/--indent-right/--first-line/--hanging";

/** Parse the shared formatting flags into a `RunFormat` + `ParagraphOptions` slice
 *  for a style definition, or a `{ error }` the caller turns into a `fail()`.
 *  Reuses the exact `edit` parsers (`parseRunFormat`, `parseSpacingIndentFlags`) so
 *  a style's `--bold`/`--size`/`--space-before` behaves identically to a body edit.
 *  `styleType` is the style's real `w:type` and gates the paragraph-level props:
 *  ONLY a paragraph style carries a `<w:pPr>`, so alignment/spacing/indent on a
 *  character/table/numbering style is a USAGE error, not a silent no-op (a table
 *  style's pPr is the per-cell default, which `set` deliberately doesn't author). */
export function parseStyleFormatting(
	values: Record<string, unknown>,
	styleType: string,
): StyleFormatting | StyleFormattingError {
	const out: StyleFormatting = {};

	const runFormat = parseRunFormat(values);
	if (runFormat && "error" in runFormat) return runFormat;
	if (runFormat) out.runFormat = runFormat;

	const paragraphOptions: ParagraphOptions = {};
	const alignment = values.alignment as string | undefined;
	if (alignment !== undefined) {
		if (!ALIGNMENTS.has(alignment)) {
			return {
				error: `Invalid --alignment: ${alignment}`,
				hint: "Use left, center, right, or justify.",
			};
		}
		paragraphOptions.alignment = alignment as ParagraphOptions["alignment"];
	}

	const spacingIndent = parseSpacingIndentFlags(values);
	if ("error" in spacingIndent) return spacingIndent;
	if (spacingIndent.spacing) paragraphOptions.spacing = spacingIndent.spacing;
	if (spacingIndent.indent) paragraphOptions.indent = spacingIndent.indent;

	const hasParagraphFlags = Object.keys(paragraphOptions).length > 0;
	if (hasParagraphFlags && styleType !== "paragraph") {
		return {
			error: `Paragraph formatting (${PARAGRAPH_FLAG_LABEL}) applies to paragraph styles, not ${styleType} styles.`,
			hint: "Drop those flags, or target/create a paragraph style (--type paragraph).",
		};
	}
	if (hasParagraphFlags) out.paragraphOptions = paragraphOptions;

	return out;
}

/** True when the invocation carries any run OR paragraph formatting flag — used to
 *  reject a `styles set` that changes nothing (no formatting and no metadata). */
export function hasStyleFormattingFlags(
	values: Record<string, unknown>,
): boolean {
	if (hasRunFormatFlags(values)) return true;
	return [
		"alignment",
		"space-before",
		"space-after",
		"line-spacing",
		"indent-left",
		"indent-right",
		"first-line",
		"hanging",
	].some((flag) => values[flag] !== undefined);
}
