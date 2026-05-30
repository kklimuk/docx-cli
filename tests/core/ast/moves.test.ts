import { describe, expect, test } from "bun:test";
import { Document } from "@core/ast/document";
import type { Body } from "@core/ast/document/body";
import {
	paragraphText,
	paragraphTextAccepted,
	paragraphTextBaseline,
} from "@core/ast/text";
import type { Paragraph } from "@core/ast/types";

function buildSyntheticView(bodyXml: string): Body {
	return Document.fromXml({
		documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`,
	}).body;
}

function paragraph(doc: Body, index: number): Paragraph {
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
