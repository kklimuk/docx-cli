import { describe, expect, test } from "bun:test";
import { Document } from "@core/ast/document";
import type { Body } from "@core/ast/document/body";
import type { Paragraph } from "@core/ast/types";

function buildSyntheticView(bodyXml: string): Body {
	return Document.fromXml({
		documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`,
	}).body;
}

function firstParagraph(doc: Body): Paragraph {
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

	test("A4: <w:r> wrapped in <w:ins> propagates trackedChange to every run (tab included)", () => {
		const doc = buildSyntheticView(
			`<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:tab/><w:t xml:space="preserve">Expires </w:t></w:r></w:ins></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs).toHaveLength(2);
		// The tab is one offset character inside the insertion, so it must carry the
		// same tracked-change decoration as the text — otherwise find/replace would
		// count it as visible in the baseline view where the insertion is hidden.
		const tabRun = runs[0];
		if (!tabRun || tabRun.type !== "tab") {
			throw new Error("expected tab run at index 0");
		}
		expect(tabRun.trackedChange?.kind).toBe("ins");
		expect(tabRun.trackedChange?.author).toBe("A");
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

// ---------------------------------------------------------------------------
// mc:AlternateContent (ECMA-376 Part 3) — resolved at EVERY level the spec
// allows a wrapper (block, paragraph, run), nested included — and text boxes.
// Issue #4: Word writes every modern shape/text box as an AlternateContent
// pair, which the walkers dropped wholesale, so a text box was invisible to
// read/find/replace with nothing said about it.
// ---------------------------------------------------------------------------

const WPS_TEXT_BOX = (paragraphs: string) =>
	`<w:drawing><wp:anchor><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:wrapSquare wrapText="bothSides"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>${paragraphs}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
const VML_TEXT_BOX = (paragraphs: string) =>
	`<w:pict><v:rect><v:textbox><w:txbxContent>${paragraphs}</w:txbxContent></v:textbox></v:rect></w:pict>`;
const BOX_PARAGRAPHS = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>CONFIDENTIAL</w:t></w:r></w:p><w:p><w:r><w:t>Acme internal use only.</w:t></w:r></w:p>`;

function wordTextBoxParagraph(paragraphs = BOX_PARAGRAPHS): string {
	return (
		`<w:p><w:r><w:t>Anchor text. </w:t><mc:AlternateContent>` +
		`<mc:Choice Requires="wps">${WPS_TEXT_BOX(paragraphs)}</mc:Choice>` +
		`<mc:Fallback>${VML_TEXT_BOX(paragraphs)}</mc:Fallback>` +
		`</mc:AlternateContent></w:r></w:p>`
	);
}

describe("mc:AlternateContent — run level (Word's shape/text-box shape)", () => {
	test("a Word text box reads as ONE TextBoxRun with tbxN:pK paragraphs", () => {
		const doc = buildSyntheticView(wordTextBoxParagraph());
		const runs = firstParagraph(doc).runs;
		expect(runs.map((run) => run.type)).toEqual(["text", "textBox"]);
		const box = runs[1];
		if (!box || box.type !== "textBox") throw new Error("expected a text box");
		expect(box.id).toBe("tbx0");
		expect(box.floating).toBe(true);
		expect(box.wrap).toBe("square");
		expect(box.align).toBe("right");
		expect(box.blocks.map((block) => block.id)).toEqual(["tbx0:p0", "tbx0:p1"]);
		const first = box.blocks[0];
		if (!first || first.type !== "paragraph") throw new Error("paragraph");
		expect(first.runs[0]).toMatchObject({ text: "CONFIDENTIAL", bold: true });
	});

	test("Choice wins: the Fallback twin does NOT surface a second box", () => {
		const doc = buildSyntheticView(wordTextBoxParagraph());
		expect([...doc.textBoxReferences.keys()]).toEqual(["tbx0"]);
		const reference = doc.textBoxReferences.get("tbx0");
		expect(reference?.anchorBlockId).toBe("p0");
		// The story the reader walked is the CHOICE copy — the one Word reads
		// (the wps shape's story, not the VML Fallback twin).
		const box = firstParagraph(doc).runs[1];
		if (!box || box.type !== "textBox") throw new Error("expected a text box");
		expect(reference?.blocks).toBe(box.blocks);
	});

	test("story paragraphs are addressable: blockReferences parent is the txbxContent child list", () => {
		const doc = buildSyntheticView(wordTextBoxParagraph());
		const reference = doc.blockReferences.get("tbx0:p1");
		expect(reference).toBeDefined();
		const story = doc.textBoxReferences.get("tbx0");
		expect(reference?.parent).toBe(story?.node.children);
		expect(reference?.node.tag).toBe("w:p");
	});

	test("a Fallback-only (VML) text box is read too", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r>${VML_TEXT_BOX(BOX_PARAGRAPHS)}</w:r></w:p>`,
		);
		const box = firstParagraph(doc).runs[0];
		if (!box || box.type !== "textBox") throw new Error("expected a text box");
		expect(box.floating).toBeUndefined();
		expect(box.blocks).toHaveLength(2);
	});

	test("a non-text-box shape under a Choice still surfaces as a [shape] placeholder", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp/></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>`,
		);
		expect(firstParagraph(doc).runs[0]).toEqual({
			type: "chart",
			kind: "shape",
		});
	});

	test("text either side of the wrapper is untouched; the box is zero-width", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>before</w:t><mc:AlternateContent><mc:Choice Requires="wps">${WPS_TEXT_BOX("<w:p><w:r><w:t>boxed</w:t></w:r></w:p>")}</mc:Choice></mc:AlternateContent><w:t>after</w:t></w:r></w:p>`,
		);
		const runs = firstParagraph(doc).runs;
		expect(runs.map((run) => run.type)).toEqual(["text", "textBox", "text"]);
	});

	test("tbxN is global in document order — a box inside a table cell keeps its own counter", () => {
		const doc = buildSyntheticView(
			wordTextBoxParagraph() +
				`<w:tbl><w:tr><w:tc>${wordTextBoxParagraph("<w:p><w:r><w:t>in cell</w:t></w:r></w:p>")}</w:tc></w:tr></w:tbl>`,
		);
		expect([...doc.textBoxReferences.keys()]).toEqual(["tbx0", "tbx1"]);
		expect(doc.textBoxReferences.get("tbx1")?.anchorBlockId).toBe("t0:r0c0:p0");
		expect(doc.blockReferences.has("tbx1:p0")).toBe(true);
	});

	test("a group shape holding two text boxes yields two stories", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wpg:wgp><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>one</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>two</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>`,
		);
		expect([...doc.textBoxReferences.keys()]).toEqual(["tbx0", "tbx1"]);
	});

	test("a text box nested inside a text box chains under its own tbxN", () => {
		const inner = wordTextBoxParagraph(
			"<w:p><w:r><w:t>inner</w:t></w:r></w:p>",
		);
		const doc = buildSyntheticView(wordTextBoxParagraph(inner));
		expect([...doc.textBoxReferences.keys()]).toEqual(["tbx0", "tbx1"]);
		expect(doc.textBoxReferences.get("tbx1")?.anchorBlockId).toBe("tbx0:p0");
		expect(doc.blockReferences.has("tbx1:p0")).toBe(true);
	});
});

describe("mc:AlternateContent — paragraph level (wrapping runs)", () => {
	test("text under the first Choice is read; the Fallback is not", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>a </w:t></w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent><w:r><w:t> z</w:t></w:r></w:p>`,
		);
		const text = firstParagraph(doc)
			.runs.map((run) => (run.type === "text" ? run.text : ""))
			.join("");
		expect(text).toBe("a choice z");
	});

	test("a tracked change inside a Choice registers with the Choice's child list as parent", () => {
		const doc = Document.fromXml({
			documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins></mc:Choice></mc:AlternateContent></w:p><w:sectPr/></w:body></w:document>`,
		});
		const reference = doc.trackedChangeReferences.get("tc0");
		expect(reference).toBeDefined();
		expect(reference?.parent[0]?.tag).toBe("w:ins");
		expect(reference?.blockId).toBe("p0");
	});
});

describe("mc:AlternateContent — block level (wrapping paragraphs)", () => {
	test("paragraphs inside a block-level Choice get sequential pN ids and a branch-local parent", () => {
		const doc = buildSyntheticView(
			`<w:p><w:r><w:t>first</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>second</w:t></w:r></w:p><w:p><w:r><w:t>third</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>old</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:t>fourth</w:t></w:r></w:p>`,
		);
		expect(doc.blocks.map((block) => block.id)).toEqual([
			"p0",
			"p1",
			"p2",
			"p3",
			"s0",
		]);
		const second = doc.blockReferences.get("p1");
		expect(second?.parent).toHaveLength(2);
		expect(second?.parent[1]?.collectText()).toBe("third");
	});

	test("a wrapper nested inside a Choice resolves recursively", () => {
		const doc = buildSyntheticView(
			`<mc:AlternateContent><mc:Choice Requires="a"><mc:AlternateContent><mc:Choice Requires="b"><w:p><w:r><w:t>deep</w:t></w:r></w:p></mc:Choice></mc:AlternateContent></mc:Choice></mc:AlternateContent>`,
		);
		const paragraph = firstParagraph(doc);
		expect(paragraph.id).toBe("p0");
		expect(paragraph.runs[0]).toMatchObject({ text: "deep" });
	});
});

describe("a drawing holding BOTH a picture and a text box", () => {
	test("registers the image (stable imgN) and the story", () => {
		const doc = Document.fromXml({
			documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData><wpg:wgp><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>caption</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:sectPr/></w:body></w:document>`,
			relationshipsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/a.png"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/b.png"/></Relationships>`,
		}).body;
		expect([...doc.imageById.keys()]).toEqual(["img0", "img1"]);
		const first = firstParagraph(doc);
		expect(first.runs.map((run) => run.type)).toEqual(["image", "textBox"]);
		expect(doc.textBoxReferences.get("tbx0")?.anchorBlockId).toBe("p0");
	});
});
