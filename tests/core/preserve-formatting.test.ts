import { describe, expect, test } from "bun:test";
import { XmlNode } from "../../src/core/parser";
import {
	buildTrackedRuns,
	buildUntrackedRuns,
	diffTokens,
	extractOldTokens,
	paragraphMarkRunRpr,
	tokenize,
} from "../../src/core/track-changes/preserve-formatting";

/**
 * Empty-paragraph rPr inheritance: when `edit --text` fills a paragraph that
 * has NO runs (a blank styled table cell), the new run must inherit the
 * paragraph-mark rPr (`<w:pPr><w:rPr>`) — the formatting Word applies when you
 * type into an empty styled cell — instead of emitting a bare, default-font run.
 */

function paragraph(xml: string): XmlNode {
	const [node] = XmlNode.parse(xml);
	if (!node) throw new Error("failed to parse paragraph");
	return node;
}

/** Run the untracked formatting-preserving fill the way `Edit.paragraph` does. */
function fillUntracked(para: XmlNode, newText: string): string {
	const ops = diffTokens(extractOldTokens(para), tokenize(newText));
	return XmlNode.serialize(buildUntrackedRuns(ops, paragraphMarkRunRpr(para)));
}

const ARIAL_9PT =
	'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';

describe("paragraphMarkRunRpr", () => {
	test("extracts the paragraph-mark rPr as a clean run rPr", () => {
		const rPr = paragraphMarkRunRpr(
			paragraph(`<w:p><w:pPr>${ARIAL_9PT}</w:pPr></w:p>`),
		);
		expect(rPr).not.toBeNull();
		const xml = XmlNode.serialize([rPr as XmlNode]);
		expect(xml).toContain('w:ascii="Arial"');
		expect(xml).toContain('w:val="18"');
	});

	test("strips ALL paragraph-mark-only markers (ins/del/moveFrom/moveTo/rPrChange) — none are valid in a run rPr", () => {
		const rPr = paragraphMarkRunRpr(
			paragraph(
				`<w:p><w:pPr><w:rPr>` +
					`<w:ins w:id="1" w:author="X" w:date="Z"/>` +
					`<w:del w:id="2" w:author="X" w:date="Z"/>` +
					`<w:moveFrom w:id="3" w:author="X" w:date="Z"/>` +
					`<w:moveTo w:id="4" w:author="X" w:date="Z"/>` +
					`<w:rPrChange w:id="5" w:author="X" w:date="Z"><w:rPr/></w:rPrChange>` +
					`<w:sz w:val="18"/>` +
					`</w:rPr></w:pPr></w:p>`,
			),
		);
		const xml = XmlNode.serialize([rPr as XmlNode]);
		for (const marker of [
			"w:ins",
			"w:del",
			"w:moveFrom",
			"w:moveTo",
			"w:rPrChange",
		]) {
			expect(xml).not.toContain(marker);
		}
		expect(xml).toContain('w:val="18"');
	});

	test("returns null when there is no paragraph-mark rPr", () => {
		expect(paragraphMarkRunRpr(paragraph("<w:p/>"))).toBeNull();
		expect(
			paragraphMarkRunRpr(paragraph("<w:p><w:pPr><w:rPr/></w:pPr></w:p>")),
		).toBeNull();
	});
});

describe("empty-paragraph fill inherits the paragraph-mark rPr", () => {
	test("untracked: a blank styled cell fill picks up Arial 9pt", () => {
		const xml = fillUntracked(
			paragraph(`<w:p><w:pPr>${ARIAL_9PT}</w:pPr></w:p>`),
			"Dana Okonkwo",
		);
		expect(xml).toContain("Dana Okonkwo");
		expect(xml).toContain('w:ascii="Arial"');
		expect(xml).toContain('w:val="18"');
	});

	test("tracked: the inserted run inside <w:ins> also inherits Arial 9pt", () => {
		const para = paragraph(`<w:p><w:pPr>${ARIAL_9PT}</w:pPr></w:p>`);
		const ops = diffTokens(extractOldTokens(para), tokenize("Dana Okonkwo"));
		const xml = XmlNode.serialize(
			buildTrackedRuns(
				ops,
				() => ({
					author: "Reviewer",
					date: "2026-01-01T00:00:00Z",
					revisionId: 1,
				}),
				paragraphMarkRunRpr(para),
			),
		);
		expect(xml).toContain("<w:ins");
		expect(xml).toContain('w:ascii="Arial"');
		expect(xml).toContain('w:val="18"');
	});

	test("a paragraph with no styling still fills as a bare run (no regression)", () => {
		const xml = fillUntracked(paragraph("<w:p/>"), "plain");
		expect(xml).toContain("plain");
		expect(xml).not.toContain("<w:rPr>");
	});
});

describe("non-empty paragraphs ignore the paragraph-mark fallback (no regression)", () => {
	test("inserted words inherit from the existing run, not the paragraph mark", () => {
		// Run is bold; the paragraph mark declares Arial. Editing must keep
		// inheriting the run's bold for the new word, not switch to the mark's rPr.
		const para = paragraph(
			`<w:p><w:pPr>${ARIAL_9PT}</w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p>`,
		);
		const xml = fillUntracked(para, "Hello world");
		expect(xml).toContain("world");
		expect(xml).toContain("<w:b/>");
		// The mark's Arial/9pt must NOT leak onto the run (it had its own rPr).
		expect(xml).not.toContain('w:ascii="Arial"');
	});
});
