import { writeAtomic } from "@core/package";
import JSZip from "jszip";
import { parseArgs } from "util";
import { EXIT, fail, respondAck, setVerboseAck, writeStdout } from "../respond";
import { CANONICAL_PARTS } from "./canonical-parts";
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
				verbose: { type: "boolean", short: "v" },
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
	const now = new Date().toISOString();

	const zip = new JSZip();
	zip.file("[Content_Types].xml", CONTENT_TYPES);
	zip.file("_rels/.rels", ROOT_RELS);
	zip.file("word/document.xml", documentXml(text));
	zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
	zip.file("docProps/core.xml", corePropertiesXml({ title, author, now }));
	for (const part of Object.values(CANONICAL_PARTS)) {
		zip.file(part.zipPath, part.body);
	}

	const buf = await zip.generateAsync({
		type: "uint8array",
		compression: "DEFLATE",
		compressionOptions: { level: 6 },
	});
	await writeAtomic(path, buf);

	await respondAck({
		ok: true,
		operation: "create",
		path,
		bytes: buf.length,
		blocks: 1,
	});
	return EXIT.OK;
}
