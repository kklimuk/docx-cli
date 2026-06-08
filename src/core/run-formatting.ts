/** Closed OOXML enums for the three run-format properties that have a fixed
 * value set. Hex colors, theme tokens, and font names are unconstrained (Word
 * degrades unknown values gracefully); only these three, if given a bad value,
 * produce schema-invalid XML — so every ingress that builds a run from
 * untrusted input (`[text]{attrs}` spans AND `--runs` JSON) validates them. */

/** `ST_HighlightColor` — 16 named highlight colors (no hex form). */
export const HIGHLIGHT_VALUES: ReadonlySet<string> = new Set([
	"black",
	"blue",
	"cyan",
	"darkBlue",
	"darkCyan",
	"darkGray",
	"darkGreen",
	"darkMagenta",
	"darkRed",
	"darkYellow",
	"green",
	"lightGray",
	"magenta",
	"red",
	"white",
	"yellow",
]);

/** `ST_Underline` — 18 underline patterns. */
export const UNDERLINE_VALUES: ReadonlySet<string> = new Set([
	"single",
	"words",
	"double",
	"thick",
	"dotted",
	"dottedHeavy",
	"dash",
	"dashedHeavy",
	"dashLong",
	"dashLongHeavy",
	"dotDash",
	"dashDotHeavy",
	"dotDotDash",
	"dashDotDotHeavy",
	"wave",
	"wavyHeavy",
	"wavyDouble",
	"none",
]);

/** `ST_VerticalAlignRun`. */
export const VERTALIGN_VALUES: ReadonlySet<string> = new Set([
	"superscript",
	"subscript",
	"baseline",
]);

export type RunFormatEnums = {
	highlight?: string;
	underline?: string;
	vertAlign?: string;
};

/** The first enum-valued run-format field with an out-of-range value, or null
 * if all are valid/absent. Callers format their own error (markdown import
 * throws `MarkdownImportError`; the `--runs` path returns a USAGE failure). */
export function firstInvalidRunFormat(
	format: RunFormatEnums,
): { field: string; value: string; valid: string } | null {
	if (format.highlight && !HIGHLIGHT_VALUES.has(format.highlight)) {
		return {
			field: "highlight",
			value: format.highlight,
			valid:
				"one of the 16 OOXML highlight names (yellow, green, cyan, darkBlue, …); for an arbitrary hex background use shade",
		};
	}
	if (format.underline && !UNDERLINE_VALUES.has(format.underline)) {
		return {
			field: "underline",
			value: format.underline,
			valid: "an OOXML underline style (single, double, dotted, wave, …)",
		};
	}
	if (format.vertAlign && !VERTALIGN_VALUES.has(format.vertAlign)) {
		return {
			field: "vertAlign",
			value: format.vertAlign,
			valid: "superscript or subscript",
		};
	}
	return null;
}
