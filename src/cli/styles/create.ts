import { isBaselineStyle } from "@core";
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
import { parseStyleFormatting, STYLE_FORMAT_FLAGS } from "./shared";

export const CREATE_HELP = `docx styles create — define a new custom style

Usage:
  docx styles create FILE STYLEID [--type paragraph|character] [formatting]
                                  [--name "Display Name"] [--based-on STYLEID]
                                  [--next STYLEID] [options]

Mints a brand-new style in word/styles.xml that \`insert --style STYLEID\` /
\`edit --style STYLEID\` can then apply. STYLEID is the internal id (no spaces —
use --name for the human label). Paragraph styles default to --based-on Normal
and --next Normal. To change a style that already exists (including built-ins
like Heading1), use \`docx styles set\` instead.

  --type paragraph|character   Style kind (default: paragraph). Character styles
                               take run formatting only.

Run formatting:
  --bold --italic --underline --strike --caps --smallcaps
  --superscript | --subscript
  --color HEX   --font NAME   --size PT   --highlight NAME   --shade HEX

Paragraph formatting (paragraph styles only):
  --alignment left|center|right|justify
  --space-before PT  --space-after PT   --line-spacing N|single|1.5|double
  --indent-left IN   --indent-right IN  --first-line IN | --hanging IN

Metadata:
  --name TEXT        Display name (defaults to STYLEID)
  --based-on STYLEID Parent style (paragraph default: Normal)
  --next STYLEID     Style for the next paragraph (paragraph default: Normal)

Options:
  -o, --output PATH  Write to PATH instead of overwriting FILE
  --dry-run          Print what would change; do not write the file
  -v, --verbose      Print the full success ack JSON
  -h, --help         Show this help

Examples:
  docx styles create report.docx Callout --name "Callout" --color C00000 --bold --size 12
  docx styles create report.docx Lead --based-on Normal --italic --space-after 12
  docx styles create report.docx KbdKey --type character --font "Consolas" --shade EEEEEE
`;

export async function runStylesCreate(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			type: { type: "string" },
			name: { type: "string" },
			"based-on": { type: "string" },
			next: { type: "string" },
			...STYLE_FORMAT_FLAGS,
			...SAVE_FLAGS,
		},
		CREATE_HELP,
	);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(CREATE_HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const { values } = parsed;
	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", CREATE_HELP);
	const styleId = parsed.positionals[1];
	if (!styleId) {
		return fail(
			"USAGE",
			"Missing STYLEID (the new style's internal id)",
			CREATE_HELP,
		);
	}
	if (/\s/.test(styleId)) {
		return fail(
			"USAGE",
			`STYLEID can't contain whitespace: "${styleId}"`,
			'Use a single token (e.g. "Callout") and pass the human label via --name "My Callout".',
		);
	}

	const typeRaw = (values.type as string | undefined) ?? "paragraph";
	if (typeRaw !== "paragraph" && typeRaw !== "character") {
		return fail(
			"USAGE",
			`--type must be paragraph or character, got "${typeRaw}"`,
			CREATE_HELP,
		);
	}
	const type = typeRaw;

	// Treat an explicit empty metadata flag (`--based-on ""`) as absent, so the
	// paragraph default (Normal) still applies and we never write a dangling
	// `w:val=""`. A style based on ITSELF is an invalid inheritance cycle.
	const name = (values.name as string | undefined) || undefined;
	const basedOn = (values["based-on"] as string | undefined) || undefined;
	const next = (values.next as string | undefined) || undefined;
	if (basedOn === styleId) {
		return fail(
			"USAGE",
			`A style can't be based on itself ("${styleId}")`,
			"Point --based-on at a different (parent) style, or drop it (it defaults to Normal).",
		);
	}

	const formatting = parseStyleFormatting(values, type);
	if ("error" in formatting) {
		return fail("USAGE", formatting.error, formatting.hint);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const styles = document.ensureStyles();

	// A built-in id isn't "created" — it's customized. Route the agent to `set`,
	// which auto-provisions the baseline def and then applies the overrides.
	if (isBaselineStyle(styleId)) {
		return fail(
			"USAGE",
			`"${styleId}" is a built-in style`,
			`Use \`docx styles set --at ${styleId}\` to customize it (it's provisioned automatically).`,
		);
	}
	if (styles.hasStyle(styleId)) {
		return fail(
			"USAGE",
			`A style with id "${styleId}" already exists`,
			`Use \`docx styles set --at ${styleId}\` to change it.`,
		);
	}

	styles.createStyle({
		styleId,
		type,
		name,
		basedOn,
		next,
		runFormat: formatting.runFormat,
		paragraphOptions: formatting.paragraphOptions,
	});

	const dryRun = Boolean(parsed.values["dry-run"]);
	const outputPath = values.output as string | undefined;

	if (dryRun) {
		await respond({
			operation: "styles.create",
			dryRun: true,
			id: styleId,
			type,
		});
		return EXIT.OK;
	}

	await document.save(outputPath);
	await respondAck(
		{
			ok: true,
			operation: "styles.create",
			path: outputPath ?? filePath,
			id: styleId,
			type,
		},
		`Apply it with \`--style ${styleId}\` on insert/edit.`,
	);
	return EXIT.OK;
}
