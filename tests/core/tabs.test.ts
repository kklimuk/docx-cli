import { describe, expect, test } from "bun:test";
import { normalizeTabAlign } from "@core/blocks";

// `normalizeTabAlign` is the single source of tab-alignment classification shared
// by the AST reader, the `docx:layout` wrap detector, and the page-setup reflow
// cure — so a stop is flagged and cured consistently (they were diverging on the
// LTR-aware aliases `start`/`end` and a missing/empty `w:val`).
describe("normalizeTabAlign", () => {
	test("LTR aliases collapse: start→left, end→right", () => {
		expect(normalizeTabAlign("start")).toBe("left");
		expect(normalizeTabAlign("end")).toBe("right");
	});

	test("missing or empty w:val → left (Word's default)", () => {
		expect(normalizeTabAlign(undefined)).toBe("left");
		expect(normalizeTabAlign("")).toBe("left");
	});

	test("canonical and other vals pass through unchanged", () => {
		for (const value of [
			"left",
			"center",
			"right",
			"decimal",
			"bar",
			"clear",
		]) {
			expect(normalizeTabAlign(value)).toBe(value);
		}
	});
});
