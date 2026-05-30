import { buildBlankPackage } from "@core/create";
import {
	EXIT,
	fail,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx create — create a new minimal .docx

Usage:
  docx create FILE [options]

Options:
  --title TEXT     Document title
  --author TEXT    Document author (default: $DOCX_AUTHOR)
  --text TEXT      Seed first paragraph with this text
  --force          Overwrite if file exists
  -v, --verbose    Print the success ack JSON (default: silent on success)
  -h, --help       Show this help

Examples:
  docx create out.docx
  docx create out.docx --title "Spec" --author "Claude" --text "First paragraph."

For a doc that opens with a code block, chain create with insert:
  docx create out.docx
  docx insert out.docx --after p0 --code-file snippet.py --language python
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			title: { type: "string" },
			author: { type: "string" },
			text: { type: "string" },
			force: { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) {
		return fail("USAGE", "Missing FILE argument", HELP);
	}

	if ((await Bun.file(path).exists()) && !parsed.values.force) {
		return fail(
			"USAGE",
			`File already exists: ${path}`,
			"Pass --force to overwrite.",
		);
	}

	const author =
		(parsed.values.author as string | undefined) ?? Bun.env.DOCX_AUTHOR ?? "";
	const title = (parsed.values.title as string | undefined) ?? "";
	const text = parsed.values.text as string | undefined;

	const pkg = buildBlankPackage({ path, title, author, text });
	await pkg.save();

	await respondAck({
		ok: true,
		operation: "create",
		path,
		bytes: Bun.file(path).size,
		blocks: 1,
	});
	return EXIT.OK;
}
