import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import JSZip from "jszip";
import { runCli, tempWorkspace } from "./harness";

// Shared building blocks for the CLI tests. Verb-specific assertions stay in the
// per-verb file; the boilerplate every file kept re-deriving (read the markdown
// or raw document.xml, copy a fixture, list tracked-change kinds) lives here.

/** `docx read` (default markdown view) → stdout. */
export async function readMarkdown(path: string): Promise<string> {
	return (await runCli("read", path)).stdout;
}

/** Raw `word/document.xml` — for asserting on XML we don't model in the AST. */
export async function readDocumentXml(path: string): Promise<string> {
	const pkg = await Pkg.open(path);
	return await pkg.readText("word/document.xml");
}

/** The `kind`s reported by `track-changes list`, in order. */
export async function trackedKinds(path: string): Promise<string[]> {
	const result = await runCli("track-changes", "list", path);
	return (result.parsed as Array<{ kind: string }>).map(
		(change) => change.kind,
	);
}

/** A fresh, mutable temp copy of a committed fixture (so tests never write to
 *  `tests/fixtures/` in place). */
export async function freshFixture(
	label: string,
	fixturePath: string,
): Promise<string> {
	const docPath = join(tempWorkspace(label), "doc.docx");
	await Bun.write(docPath, Bun.file(fixturePath));
	return docPath;
}

/** Build a fresh document with one paragraph followed by a blank table. */
export async function newTableDoc(
	label: string,
	rows = 1,
	cols = 2,
	initialText = "Before",
): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", initialText);
	await runCli(
		"tables",
		"create",
		docPath,
		"--after",
		"p0",
		"--rows",
		String(rows),
		"--cols",
		String(cols),
	);
	return docPath;
}

/** A minimal .docx whose `<w:body>` is exactly `bodyXml` — for shapes no CLI
 * verb can author (raw content controls, grid-shifted rows, unmodeled cell
 * children). Returns the path. */
export async function buildRawDoc(
	bodyXml: string,
	label: string,
): Promise<string> {
	const docPath = join(
		mkdtempSync(join(tmpdir(), `docx-cli-${label}-`)),
		"out.docx",
	);
	const zip = new JSZip();
	zip.file(
		"[Content_Types].xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
	);
	zip.file(
		"_rels/.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
	);
	zip.file(
		"word/document.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`,
	);
	zip.file(
		"word/_rels/document.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
	);
	await Bun.write(
		docPath,
		await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
	);
	return docPath;
}
