import { cp, dc, dcterms, w } from "@core/jsx";
import { XmlNode } from "@core/parser";
import { CANONICAL_PARTS } from "./canonical-parts";

const DOC_NAMESPACES = {
	"xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	"xmlns:r":
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
	"xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
	"xmlns:wp":
		"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
	"xmlns:wp14":
		"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
	"xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
	"xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
	"xmlns:w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
	"mc:Ignorable": "w14 w15 w16se wp14",
};

export const CONTENT_TYPES = buildContentTypes();
export const ROOT_RELS = buildRootRels();
export const DOCUMENT_RELS = buildDocumentRels();

export function documentXml(text: string | undefined): string {
	return serializeWithDeclaration(<DocumentBody text={text} />);
}

export function corePropertiesXml(options: {
	title?: string;
	author?: string;
	now: string;
}): string {
	return serializeWithDeclaration(
		<CoreProperties
			title={options.title ?? ""}
			author={options.author ?? ""}
			now={options.now}
		/>,
	);
}

function DocumentBody({ text }: { text: string | undefined }): XmlNode {
	return (
		<w.document {...DOC_NAMESPACES}>
			<w.body>
				{text !== undefined ? (
					<w.p>
						<w.r>
							<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
						</w.r>
					</w.p>
				) : (
					<w.p />
				)}
				<DefaultSectionProperties />
			</w.body>
		</w.document>
	);
}

/** Letter paper (8.5×11"), 1" margins. Matches Word's "Blank Document" default
 * so the layout looks familiar when a user opens the file. */
function DefaultSectionProperties(): XmlNode {
	return (
		<w.sectPr>
			{XmlNode.element("w:pgSz", { "w:w": "12240", "w:h": "15840" })}
			{XmlNode.element("w:pgMar", {
				"w:top": "1440",
				"w:right": "1440",
				"w:bottom": "1440",
				"w:left": "1440",
				"w:header": "720",
				"w:footer": "720",
				"w:gutter": "0",
			})}
			{XmlNode.element("w:cols", { "w:space": "720" })}
			{XmlNode.element("w:docGrid", { "w:linePitch": "360" })}
		</w.sectPr>
	);
}

function CoreProperties({
	title,
	author,
	now,
}: {
	title: string;
	author: string;
	now: string;
}): XmlNode {
	const namespaceAttrs = {
		"xmlns:cp":
			"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
		"xmlns:dc": "http://purl.org/dc/elements/1.1/",
		"xmlns:dcterms": "http://purl.org/dc/terms/",
		"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
	};
	return (
		<cp.coreProperties {...namespaceAttrs}>
			<dc.title>{title}</dc.title>
			<dc.creator>{author}</dc.creator>
			<cp.lastModifiedBy>{author}</cp.lastModifiedBy>
			<dcterms.created xsi-type="dcterms:W3CDTF">{now}</dcterms.created>
			<dcterms.modified xsi-type="dcterms:W3CDTF">{now}</dcterms.modified>
		</cp.coreProperties>
	);
}

function buildContentTypes(): string {
	const documentOverrides = Object.values(CANONICAL_PARTS)
		.map(
			(part) =>
				`\t<Override PartName="/${part.zipPath}" ContentType="${part.contentType}"/>`,
		)
		.join("\n");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
${documentOverrides}
</Types>`;
}

function buildRootRels(): string {
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

function buildDocumentRels(): string {
	const documentParts = Object.values(CANONICAL_PARTS).filter(
		(part) => part.scope === "document",
	);
	const lines = documentParts.map(
		(part, index) =>
			`\t<Relationship Id="rId${index + 1}" Type="${part.relationshipType}" Target="${part.target}"/>`,
	);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${lines.join("\n")}
</Relationships>`;
}

function serializeWithDeclaration(root: XmlNode): string {
	return XmlNode.serializeWithDeclaration([root]);
}
