import { writeAtomic } from "@core/package";
import JSZip from "jszip";
import { parseArgs } from "util";
import { EXIT, fail, respond, writeStdout } from "../respond";
import {
	CONTENT_TYPES,
	corePropertiesXml,
	DOCUMENT_RELS,
	documentXml,
	ROOT_RELS,
} from "./template";

const HELP = `docx create — create a new minimal .docx

Usage:
  docx create FILE [options]

Options:
  --title TEXT     Document title
  --author TEXT    Document author (default: $DOCX_AUTHOR)
  --text TEXT      Seed first paragraph with this text
  --force          Overwrite if file exists
  -h, --help       Show this help

Examples:
  docx create out.docx
  docx create out.docx --title "Spec" --author "Claude" --text "First paragraph."
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				title: { type: "string" },
				author: { type: "string" },
				text: { type: "string" },
				force: { type: "boolean" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (e) {
		return fail("USAGE", e instanceof Error ? e.message : String(e), HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

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
	const now = new Date().toISOString();

	const zip = new JSZip();
	zip.file("[Content_Types].xml", CONTENT_TYPES);
	zip.file("_rels/.rels", ROOT_RELS);
	zip.file("word/document.xml", documentXml(text));
	zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
	zip.file("docProps/core.xml", corePropertiesXml({ title, author, now }));

	const buf = await zip.generateAsync({
		type: "uint8array",
		compression: "DEFLATE",
		compressionOptions: { level: 6 },
	});
	await writeAtomic(path, buf);

	await respond({
		ok: true,
		operation: "create",
		path,
		bytes: buf.length,
		blocks: 1,
	});
	return EXIT.OK;
}
