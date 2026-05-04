import { cp, dc, dcterms, w } from "@core/jsx";
import { XmlNode } from "@core/parser";

export const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
	<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

export const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

export const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;

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
	const namespaceAttrs = {
		"xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	};
	return (
		<w.document {...namespaceAttrs}>
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
				<w.sectPr />
			</w.body>
		</w.document>
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

function serializeWithDeclaration(root: XmlNode): string {
	return XmlNode.serializeWithDeclaration([root]);
}
