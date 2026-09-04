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

/** The XML Word writes for a floating text box — both branches of the
 * `<mc:AlternateContent>` pair, the same story in each. Shared by the CLI
 * text-box tests and the `text-boxes.docx` fixture builder. */
const NS =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';

export function storyXml(lines: string[]): string {
	return lines
		.map(
			(line) => `<w:p><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`,
		)
		.join("");
}

/** The paragraph Word writes for a floating text box, both branches. */
/** Options: `leadingText` puts text in the SAME `<w:r>` before the shape (the
 * shape Word writes when a box is anchored mid-run) and `docPrId` keeps two
 * boxes' drawing ids distinct. */
export function wordTextBoxParagraphXml(
	lines: string[],
	options: { leadingText?: string; docPrId?: number } = {},
): string {
	const story = storyXml(lines);
	const lead = options.leadingText
		? `<w:t xml:space="preserve">${options.leadingText}</w:t>`
		: "";
	const docPrId = options.docPrId ?? 1;
	return (
		`<w:p ${NS}><w:r>${lead}<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>` +
		`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
		`<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>` +
		`<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
		`<wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>` +
		`<wp:docPr id="${docPrId}" name="Text Box ${docPrId}"/><wp:cNvGraphicFramePr/>` +
		`<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
		`<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
		`<wps:txbx><w:txbxContent>${story}</w:txbxContent></wps:txbx>` +
		`<wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>` +
		`</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
		`<mc:Fallback><w:pict><v:rect style="position:absolute;margin-left:0;margin-top:0;width:216pt;height:72pt;z-index:251659264">` +
		`<v:textbox><w:txbxContent>${story}</w:txbxContent></v:textbox></v:rect></w:pict></mc:Fallback>` +
		`</mc:AlternateContent></w:r></w:p>`
	);
}
