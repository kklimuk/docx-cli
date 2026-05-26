import { describe, expect, test } from "bun:test";
import type { DocView } from "@core/ast/doc-view";
import { buildDoc } from "@core/ast/read";
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

function firstParagraph(doc: Doc): Paragraph {
	const block = doc.blocks[0];
	if (!block || block.type !== "paragraph") {
		throw new Error("expected paragraph as first block");
	}
	return block;
}

describe("readRun (Bug A)", () => {
	test("A1: <w:tab/> + <w:t> in one <w:r> emits two runs in document order", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:tab/><w:t xml:space="preserve">Expires </w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(2);
		expect(runs[0]).toEqual({ type: "tab" });
		expect(runs[1]).toMatchObject({ type: "text", text: "Expires " });
	});

	test("A2: text / tab / text in one <w:r> emits three runs in order", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(3);
		expect(runs[0]).toMatchObject({ type: "text", text: "foo" });
		expect(runs[1]).toEqual({ type: "tab" });
		expect(runs[2]).toMatchObject({ type: "text", text: "bar" });
	});

	test("A3: <w:rPr> formatting applies to every TextRun emitted from the same <w:r>", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(3);
		expect(runs[0]).toMatchObject({ type: "text", text: "foo", bold: true });
		expect(runs[1]).toEqual({ type: "tab" });
		expect(runs[2]).toMatchObject({ type: "text", text: "bar", bold: true });
	});

	test("A4: <w:r> wrapped in <w:ins> propagates trackedChange to every TextRun", () => {
		const doc = buildSyntheticView(
			`<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:tab/><w:t xml:space="preserve">Expires </w:t></w:r></w:ins></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(2);
		expect(runs[0]).toEqual({ type: "tab" });
		const textRun = runs[1];
		if (!textRun || textRun.type !== "text") {
			throw new Error("expected text run at index 1");
		}
		expect(textRun.text).toBe("Expires ");
		expect(textRun.trackedChange).toBeDefined();
		expect(textRun.trackedChange?.kind).toBe("ins");
		expect(textRun.trackedChange?.author).toBe("A");
	});

	test("paragraph offset accounting: subsequent <w:r> sees the right starting offset", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:r><w:tab/><w:t>foo</w:t></w:r>` +
				`<w:commentRangeStart w:id="0"/>` +
				`<w:r><w:t>bar</w:t></w:r>` +
				`<w:commentRangeEnd w:id="0"/>` +
				`</w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		// foo (3) + tab does not contribute to text offset, so commentRangeStart sees offset 3
		expect(runs).toHaveLength(3);
		const lastRun = runs[2];
		if (!lastRun || lastRun.type !== "text") {
			throw new Error("expected text run at index 2");
		}
		expect(lastRun.text).toBe("bar");
		expect(lastRun.comments).toEqual(["c0"]);
	});
});

describe("readRun — additional inline children", () => {
	test("<w:noBreakHyphen/> folds into the surrounding TextRun as U+2011", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>author</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(1);
		expect(runs[0]).toMatchObject({ type: "text", text: "co‑author" });
	});

	test("<w:softHyphen/> folds into the surrounding TextRun as U+00AD", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>hyper</w:t><w:softHyphen/><w:t>active</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(1);
		expect(runs[0]).toMatchObject({ type: "text", text: "hyper­active" });
	});

	test("<w:sym> with Symbol font decodes to Greek glyph", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>angle </w:t><w:sym w:font="Symbol" w:char="44"/><w:t>x</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(1);
		expect(runs[0]).toMatchObject({ type: "text", text: "angle Δx" });
	});

	test("<w:cr/> emits a line BreakRun", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>line1</w:t><w:cr/><w:t>line2</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(3);
		expect(runs[0]).toMatchObject({ type: "text", text: "line1" });
		expect(runs[1]).toEqual({ type: "break", kind: "line" });
		expect(runs[2]).toMatchObject({ type: "text", text: "line2" });
	});

	test("<w:ptab/> emits a TabRun", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>foo</w:t><w:ptab w:relativeTo="margin" w:alignment="left" w:leader="none"/><w:t>bar</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(3);
		expect(runs[1]).toEqual({ type: "tab" });
	});

	test("<w:pict> and <w:object> emit ChartRun placeholders", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>before</w:t><w:pict/></w:r><w:r><w:object/><w:t>after</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs.filter((r) => r.type === "chart")).toHaveLength(2);
	});

	test("offset advances by 1 for noBreakHyphen — paragraph offsets stay aligned with AST text", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>op</w:t></w:r>` +
				`<w:commentRangeStart w:id="0"/>` +
				`<w:r><w:t>!</w:t></w:r>` +
				`<w:commentRangeEnd w:id="0"/>` +
				`</w:p>`,
		);
		const paragraph = firstParagraph(doc);
		// "co" + nbhyphen + "op" = 5 chars, so the "!" run starts at offset 5
		// and the comment span anchor (after the comment-range markers attach)
		// is what's tested elsewhere — here we just confirm the AST text is right.
		const text = paragraph.runs
			.filter((r) => r.type === "text")
			.map((r) => (r.type === "text" ? r.text : ""))
			.join("");
		expect(text).toBe("co‑op!");
		// 'c', 'o', U+2011, 'o', 'p', '!' = 6 chars
		expect(text.length).toBe(6);
	});
});

describe("walkRunContainer — paragraph-level wrappers", () => {
	test("<w:fldSimple> contents are surfaced as plain runs", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">Today is </w:t></w:r>` +
				`<w:fldSimple w:instr=" DATE \\@ &quot;yyyy-MM-dd&quot;">` +
				`<w:r><w:t>2026-05-05</w:t></w:r>` +
				`</w:fldSimple>` +
				`</w:p>`,
		);
		const text = firstParagraph(doc)
			.runs.filter((r) => r.type === "text")
			.map((r) => (r.type === "text" ? r.text : ""))
			.join("");
		expect(text).toBe("Today is 2026-05-05");
	});

	test("<w:smartTag> contents are surfaced as plain runs", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:r><w:t xml:space="preserve">Met </w:t></w:r>` +
				`<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="PersonName">` +
				`<w:r><w:t>Alice</w:t></w:r>` +
				`</w:smartTag>` +
				`<w:r><w:t xml:space="preserve"> today.</w:t></w:r>` +
				`</w:p>`,
		);
		const text = firstParagraph(doc)
			.runs.filter((r) => r.type === "text")
			.map((r) => (r.type === "text" ? r.text : ""))
			.join("");
		expect(text).toBe("Met Alice today.");
	});

	test("<w:moveFrom> attaches a trackedChange with kind 'moveFrom'", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:moveFrom w:id="1" w:author="A" w:date="2026-05-05T00:00:00Z">` +
				`<w:r><w:delText xml:space="preserve">moved out </w:delText></w:r>` +
				`</w:moveFrom>` +
				`</w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(1);
		const run = runs[0];
		if (!run || run.type !== "text") throw new Error("expected text run");
		expect(run.text).toBe("moved out ");
		expect(run.trackedChange?.kind).toBe("moveFrom");
		expect(run.trackedChange?.author).toBe("A");
	});

	test("<w:moveTo> attaches a trackedChange with kind 'moveTo'", () => {
		const doc = buildSyntheticView(
			`<w:p>` +
				`<w:moveTo w:id="2" w:author="A" w:date="2026-05-05T00:00:00Z">` +
				`<w:r><w:t xml:space="preserve">moved here </w:t></w:r>` +
				`</w:moveTo>` +
				`</w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		const run = runs[0];
		if (!run || run.type !== "text") throw new Error("expected text run");
		expect(run.trackedChange?.kind).toBe("moveTo");
	});
});
