import { Fonts } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

export const FONT_HELP = `docx styles set-default-font — set the document-wide default font

Usage:
  docx styles set-default-font FILE "Font Name" [--size N] [--all] [options]

A document font lives in TWO places at once — word/styles.xml (<w:docDefaults>)
and the theme font scheme (word/theme/theme1.xml, major + minor) — and setting
only one silently loses to the other. This sets both, so body text AND
theme-following headings adopt the font. Styles/runs that pin their OWN font
(e.g. a code block's monospace, a deliberately-Arial run) are preserved; pass
--all to repoint those too.

Theme-following headings only adopt the font if the document HAS a theme part
(every doc this CLI creates does; Word docs do too). The ack's "themeUpdated"
is false for the rare theme-less doc — there the body changes but headings keep
their fallback font; run with --all to force the headings onto the new font.

Options:
  --size N           Also set the default font size, in points (e.g. 12).
  --all              Repoint EVERY explicit font — styles, body runs, and notes —
                     onto FONT too, for a guaranteed-uniform document (overrides
                     even code monospace and per-run font choices).
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write the file
  -v, --verbose      Print the full success ack JSON
  -h, --help         Show this help

Examples:
  docx styles set-default-font report.docx "Times New Roman"
  docx styles set-default-font report.docx "Calibri" --size 11
  docx styles set-default-font report.docx "Georgia" --all
`;

export async function runSetDefaultFont(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{ size: { type: "string" }, all: { type: "boolean" }, ...SAVE_FLAGS },
		FONT_HELP,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(FONT_HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", FONT_HELP);
	const fontName = parsed.positionals[1];
	if (!fontName) {
		return fail(
			"USAGE",
			'Missing FONT name (e.g. "Times New Roman")',
			FONT_HELP,
		);
	}

	let sizeHalfPoints: number | undefined;
	const sizeRaw = parsed.values.size as string | undefined;
	if (sizeRaw !== undefined) {
		// Strict decimal only — `Number.parseFloat` would silently accept trailing
		// garbage ("11pt"→11, "1e3"→1000, "12abc"→12) and write a wrong size.
		if (
			!/^\s*\d+(\.\d+)?\s*$/.test(sizeRaw) ||
			Number.parseFloat(sizeRaw) <= 0
		) {
			return fail(
				"USAGE",
				`--size must be a positive number of points, got "${sizeRaw}"`,
			);
		}
		sizeHalfPoints = Math.round(Number.parseFloat(sizeRaw) * 2);
	}

	const all = Boolean(parsed.values.all);
	const outputPath = parsed.values.output as string | undefined;
	const dryRun = Boolean(parsed.values["dry-run"]);

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const result = await new Fonts(document).setDefault(fontName, {
		sizeHalfPoints,
		all,
	});

	if (dryRun) {
		await respond({
			operation: "styles.set-default-font",
			dryRun: true,
			path: filePath,
			font: fontName,
			...(sizeHalfPoints !== undefined ? { sizePt: sizeHalfPoints / 2 } : {}),
			all,
			themeUpdated: result.themeUpdated,
			// Preview the decision-relevant fact: which styles stay off-font (or,
			// under --all, how many fonts would be repointed) — same as the real run.
			...(all
				? { repointed: result.repointed }
				: { explicitStyles: result.explicitStyles }),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	await document.save(outputPath);

	// Tell the agent what kept its own font (so "why is the heading still blue
	// Calibri?" has an answer) and how to force it.
	const count = result.explicitStyles.length;
	const leftover =
		!all && count > 0
			? `${count === 1 ? "1 style keeps" : `${count} styles keep`} their own font (${formatStyleList(result.explicitStyles)}); pass --all to override them too.`
			: undefined;

	await respondAck(
		{
			ok: true,
			operation: "styles.set-default-font",
			path: outputPath ?? filePath,
			font: fontName,
			...(sizeHalfPoints !== undefined ? { sizePt: sizeHalfPoints / 2 } : {}),
			themeUpdated: result.themeUpdated,
			all,
			...(all
				? { repointed: result.repointed }
				: { explicitStyles: result.explicitStyles }),
		},
		leftover,
	);
	return EXIT.OK;
}

function formatStyleList(ids: string[]): string {
	const shown = ids.slice(0, 5).join(", ");
	return ids.length > 5 ? `${shown}, …` : shown;
}
