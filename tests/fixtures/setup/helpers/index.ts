import { CANONICAL_PARTS } from "@core/create";
import type JSZip from "jszip";

/**
 * Shared boilerplate for hand-rolled fixture scripts. Centralizes the
 * Word-canonical package shape so individual fixtures only have to declare
 * their body content + extras (comments.xml, footnotes.xml, etc.).
 *
 * Why this exists: a hand-rolled docx must include the canonical parts that
 * Word treats as required (styles, settings, fontTable, webSettings, theme,
 * app) — without them Word shows an "unreadable content / recover?" prompt.
 * It must also avoid inter-element whitespace inside <w:body> and
 * <cp:coreProperties>, where whitespace becomes illegal character data. See
 * CLAUDE.md's Architectural Invariants for the full rule.
 */

export type ExtraPart = {
	/** Zip path under package root, e.g. "word/comments.xml" */
	partName: string;
	/** Content-Type override */
	contentType: string;
	/** Relationship type URL */
	relationshipType: string;
	/** Target relative to word/ (the document part) */
	target: string;
	/** Raw XML body */
	body: string;
};

/** Stripped to a single line so Word's reader doesn't see inter-element
 * whitespace as illegal character data. Preserves whitespace inside text
 * nodes (collapses only `\n` + following indent, not interior spaces). */
function normalize(prettyXml: string): string {
	return prettyXml.replace(/\n\s*/g, "");
}

const DOC_NAMESPACES = [
	`xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
	`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`,
	`xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"`,
	`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`,
	`xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`,
	`xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"`,
	`xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`,
	`xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`,
	`xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"`,
	`mc:Ignorable="w14 w15 w16se wp14"`,
].join(" ");

/** Letter-paper sectPr matching what `docx create` emits. Callers can paste
 * this into their body if they want the same default layout. */
export const DEFAULT_SECTPR = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`;

/** Wrap pretty body content (everything that goes inside <w:body>) in the
 * canonical <w:document>. The result has the body collapsed to a single line
 * so Word doesn't see inter-element whitespace as illegal character data. */
export function wrapDocument(prettyBody: string): string {
	const body = normalize(prettyBody);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${DOC_NAMESPACES}><w:body>${body}</w:body></w:document>`;
}

/** Single-line core.xml. The dcterms dates carry the xsi:type Word emits. */
export function buildCoreProps(opts: {
	title?: string;
	author?: string;
	created?: string;
	modified?: string;
}): string {
	const title = escapeXml(opts.title ?? "");
	const author = escapeXml(opts.author ?? "");
	const created = opts.created ?? "2026-05-21T00:00:00Z";
	const modified = opts.modified ?? created;
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${title}</dc:title><dc:creator>${author}</dc:creator><cp:lastModifiedBy>${author}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${modified}</dcterms:modified></cp:coreProperties>`;
}

export function buildContentTypes(extras: ExtraPart[] = []): string {
	const canonicalOverrides = Object.values(CANONICAL_PARTS).map(
		(part) =>
			`\t<Override PartName="/${part.zipPath}" ContentType="${part.contentType}"/>`,
	);
	const extraOverrides = extras.map(
		(part) =>
			`\t<Override PartName="/${part.partName}" ContentType="${part.contentType}"/>`,
	);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
\t<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
\t<Default Extension="xml" ContentType="application/xml"/>
\t<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
\t<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
${[...canonicalOverrides, ...extraOverrides].join("\n")}
</Types>`;
}

export function buildRootRels(): string {
	const packageScoped = Object.values(CANONICAL_PARTS).filter(
		(part) => part.scope === "package",
	);
	const lines = [
		`\t<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`,
		`\t<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
		...packageScoped.map(
			(part, index) =>
				`\t<Relationship Id="rId${3 + index}" Type="${part.relationshipType}" Target="${part.target}"/>`,
		),
	];
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${lines.join("\n")}
</Relationships>`;
}

export function buildDocumentRels(extras: ExtraPart[] = []): string {
	const documentScoped = Object.values(CANONICAL_PARTS).filter(
		(part) => part.scope === "document",
	);
	const canonical = documentScoped.map(
		(part, index) =>
			`\t<Relationship Id="rId${index + 1}" Type="${part.relationshipType}" Target="${part.target}"/>`,
	);
	const extraLines = extras.map(
		(part, index) =>
			`\t<Relationship Id="rId${documentScoped.length + index + 1}" Type="${part.relationshipType}" Target="${part.target}"/>`,
	);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${[...canonical, ...extraLines].join("\n")}
</Relationships>`;
}

/** Write the canonical word-internal parts (styles/settings/font/web/theme)
 * and the canonical docProps/app.xml into the zip. The caller is responsible
 * for the package-scoped boilerplate (Content_Types, rels) via the build* helpers
 * above, and for any extra parts. */
export function addCanonicalParts(zip: JSZip): void {
	for (const part of Object.values(CANONICAL_PARTS)) {
		zip.file(part.zipPath, part.body);
	}
}

function escapeXml(text: string): string {
	return text
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");
}
