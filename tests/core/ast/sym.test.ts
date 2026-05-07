import { describe, expect, test } from "bun:test";
import { decodeSym } from "@core/ast/sym";

describe("decodeSym", () => {
	test("Symbol font: ASCII codepoints map to Greek/math glyphs", () => {
		expect(decodeSym("Symbol", "61")).toBe("α");
		expect(decodeSym("Symbol", "70")).toBe("π");
		expect(decodeSym("Symbol", "44")).toBe("Δ");
		expect(decodeSym("Symbol", "AE")).toBe("→");
	});

	test("Symbol font lookup is case-insensitive", () => {
		expect(decodeSym("symbol", "61")).toBe("α");
		expect(decodeSym("SYMBOL", "61")).toBe("α");
	});

	test("Wingdings: PUA codepoints map to Unicode equivalents", () => {
		expect(decodeSym("Wingdings", "F0FC")).toBe("✓");
		expect(decodeSym("Wingdings", "F0FE")).toBe("✗");
		expect(decodeSym("Wingdings", "F0E1")).toBe("→");
	});

	test("Wingdings: ASCII alias of a PUA codepoint also maps", () => {
		expect(decodeSym("Wingdings", "FC")).toBe("✓");
	});

	test("ZapfDingbats: known check/cross marks decode", () => {
		expect(decodeSym("ZapfDingbats", "33")).toBe("✓");
		expect(decodeSym("Zapf Dingbats", "37")).toBe("✗");
	});

	test("unknown font falls back to literal codepoint decode", () => {
		expect(decodeSym("Some Custom Font", "41")).toBe("A");
		expect(decodeSym("Some Custom Font", "2660")).toBe("♠");
	});

	test("malformed hex returns empty string", () => {
		expect(decodeSym("Symbol", "")).toBe("");
		expect(decodeSym("Symbol", "ZZZ")).toBe("");
	});
});
