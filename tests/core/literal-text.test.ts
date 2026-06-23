import { describe, expect, test } from "bun:test";
import { literalParagraphs } from "@core";
import { XmlNode } from "@core/parser";

/** The text of each `<w:p>` (its `<w:t>` runs joined) — enough to assert the
 *  newline → paragraph split without reaching into the node shape. */
function paragraphTexts(nodes: XmlNode[]): string[] {
	return nodes.map((node) =>
		[...XmlNode.serialize([node]).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
			.map((match) => match[1])
			.join(""),
	);
}

describe("literalParagraphs (parser-free splitting)", () => {
	test("every newline starts a new paragraph", () => {
		expect(paragraphTexts(literalParagraphs("alpha\nbravo"))).toEqual([
			"alpha",
			"bravo",
		]);
	});

	test("an interior blank line becomes an empty paragraph", () => {
		expect(paragraphTexts(literalParagraphs("alpha\n\nbravo"))).toEqual([
			"alpha",
			"",
			"bravo",
		]);
	});

	test("a single trailing newline does not mint a stray paragraph", () => {
		expect(paragraphTexts(literalParagraphs("alpha\nbravo\n"))).toEqual([
			"alpha",
			"bravo",
		]);
	});

	test("CRLF / CR line endings normalize to one paragraph per line", () => {
		expect(paragraphTexts(literalParagraphs("alpha\r\nbravo\r"))).toEqual([
			"alpha",
			"bravo",
		]);
	});

	test("empty input yields one empty paragraph (never zero blocks)", () => {
		expect(paragraphTexts(literalParagraphs(""))).toEqual([""]);
	});

	test("content GFM would mangle is kept verbatim", () => {
		expect(
			paragraphTexts(
				literalParagraphs("3. note\n*x* and _y_\nhttps://example.com\n{++z++}"),
			),
		).toEqual(["3. note", "*x* and _y_", "https://example.com", "{++z++}"]);
	});

	test("paragraph options apply to every emitted paragraph", () => {
		const xml = XmlNode.serialize(
			literalParagraphs("one\ntwo", { style: "Quote" }),
		);
		expect([...xml.matchAll(/<w:pStyle w:val="Quote"/g)]).toHaveLength(2);
	});
});
