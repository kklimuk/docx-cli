import { describe, expect, test } from "bun:test";
import type { DocView } from "@core/ast/doc-view";
import { buildDoc } from "@core/ast/read";
import {
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
} from "@core/ast/text";
import type { Doc, Paragraph } from "@core/ast/types";
import { XmlNode } from "@core/parser";

function buildSyntheticView(bodyXml: string): Doc {
	const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`;
	const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
	const typesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`;
	const view: DocView = {
		pkg: undefined as unknown as DocView["pkg"],
		documentTree: XmlNode.parse(documentXml),
		relationshipsTree: XmlNode.parse(relsXml),
		contentTypesTree: XmlNode.parse(typesXml),
		doc: undefined as unknown as Doc,
		blockReferences: new Map(),
		commentReferences: new Map(),
		imagesByRelationshipId: new Map(),
		imageById: new Map(),
		hyperlinksByRelationshipId: new Map(),
		hyperlinkById: new Map(),
		trackedChangeReferences: new Map(),
		equationReferences: new Map(),
	};
	view.doc = buildDoc(view, "synthetic.docx");
	return view.doc;
}

function paragraph(doc: Doc, index: number): Paragraph {
	const block = doc.blocks[index];
	if (!block || block.type !== "paragraph") {
		throw new Error(`expected paragraph at index ${index}`);
	}
	return block;
}

describe("text view helpers — tracked moves", () => {
	const movedDoc =
		`<w:p>` +
		`<w:r><w:t xml:space="preserve">prefix </w:t></w:r>` +
		`<w:moveFrom w:id="1" w:author="A" w:date="2026-05-05T00:00:00Z">` +
		`<w:r><w:delText>moved</w:delText></w:r>` +
		`</w:moveFrom>` +
		`<w:r><w:t xml:space="preserve"> middle </w:t></w:r>` +
		`<w:moveTo w:id="2" w:author="A" w:date="2026-05-05T00:00:00Z">` +
		`<w:r><w:t>destination</w:t></w:r>` +
		`</w:moveTo>` +
		`<w:r><w:t xml:space="preserve"> suffix</w:t></w:r>` +
		`</w:p>`;

	test("default (current) view includes both moveFrom and moveTo content", () => {
		const doc = buildSyntheticView(movedDoc);
		expect(paragraphText(paragraph(doc, 0))).toBe(
			"prefix moved middle destination suffix",
		);
	});

	test("accepted view drops moveFrom but keeps moveTo", () => {
		const doc = buildSyntheticView(movedDoc);
		expect(paragraphTextAccepted(paragraph(doc, 0))).toBe(
			"prefix  middle destination suffix",
		);
	});

	test("baseline view keeps moveFrom but drops moveTo", () => {
		const doc = buildSyntheticView(movedDoc);
		expect(paragraphTextBaseline(paragraph(doc, 0))).toBe(
			"prefix moved middle  suffix",
		);
	});
});
