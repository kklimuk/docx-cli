import { describe, expect, test } from "bun:test";
import { LocatorParseError, parseLocator } from "../../../src/core/locators";

describe("parseLocator", () => {
	test("parses block ids", () => {
		expect(parseLocator("p3")).toEqual({ kind: "block", blockId: "p3" });
		expect(parseLocator("t0")).toEqual({ kind: "block", blockId: "t0" });
		expect(parseLocator("s1")).toEqual({ kind: "block", blockId: "s1" });
		expect(parseLocator("cellp2")).toEqual({
			kind: "block",
			blockId: "cellp2",
		});
	});

	test("parses comment and image ids", () => {
		expect(parseLocator("c1")).toEqual({ kind: "comment", commentId: "c1" });
		expect(parseLocator("img0")).toEqual({ kind: "image", imageId: "img0" });
	});

	test("parses span within a paragraph", () => {
		expect(parseLocator("p3:5-20")).toEqual({
			kind: "blockSpan",
			blockId: "p3",
			start: 5,
			end: 20,
		});
	});

	test("parses cross-block range", () => {
		expect(parseLocator("p3:5-p5:10")).toEqual({
			kind: "range",
			start: { blockId: "p3", offset: 5 },
			end: { blockId: "p5", offset: 10 },
		});
	});

	test("parses table cell with optional inner locator", () => {
		expect(parseLocator("t0:r1c2")).toEqual({
			kind: "cell",
			tableId: "t0",
			row: 1,
			col: 2,
		});
		expect(parseLocator("t0:r1c2:p0")).toEqual({
			kind: "cell",
			tableId: "t0",
			row: 1,
			col: 2,
			inner: { kind: "block", blockId: "p0" },
		});
	});

	test("rejects malformed input", () => {
		expect(() => parseLocator("")).toThrow(LocatorParseError);
		expect(() => parseLocator("abc")).toThrow(LocatorParseError);
		expect(() => parseLocator("p")).toThrow(LocatorParseError);
		expect(() => parseLocator("p3:5-")).toThrow(LocatorParseError);
		expect(() => parseLocator("p3:20-5")).toThrow(LocatorParseError);
		expect(() => parseLocator("p3:-5-20")).toThrow(LocatorParseError);
	});
});
