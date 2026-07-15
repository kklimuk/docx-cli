import { describe, expect, test } from "bun:test";
import { diffHunks, diffStats, renderUnified } from "@core/diff";

/** Newline-terminated text from an array of lines (read output always ends `\n`). */
const doc = (lines: string[]): string => `${lines.join("\n")}\n`;

describe("diffHunks", () => {
	test("identical texts produce no hunks", () => {
		expect(diffHunks("a\nb\nc\n", "a\nb\nc\n")).toEqual([]);
	});

	test("pure insertion", () => {
		const hunks = diffHunks("a\nb\n", "a\nnew\nb\n");
		expect(hunks).toHaveLength(1);
		expect(
			hunks[0]?.lines.filter((l) => l.kind === "insert").map((l) => l.text),
		).toEqual(["new"]);
		expect(hunks[0]?.lines.some((l) => l.kind === "delete")).toBe(false);
	});

	test("pure deletion", () => {
		const hunks = diffHunks("a\ngone\nb\n", "a\nb\n");
		expect(
			hunks[0]?.lines.filter((l) => l.kind === "delete").map((l) => l.text),
		).toEqual(["gone"]);
	});

	test("a change is a delete plus an insert", () => {
		const kinds = diffHunks("a\nold\nb\n", "a\nnew\nb\n")[0]?.lines.map(
			(l) => l.kind,
		);
		expect(kinds).toContain("delete");
		expect(kinds).toContain("insert");
	});

	test("far-apart changes split into separate hunks", () => {
		const before = doc(["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]);
		const after = doc(["A", "b", "c", "d", "e", "f", "g", "h", "i", "J"]);
		expect(diffHunks(before, after)).toHaveLength(2);
	});

	test("changes within 2×context merge into one hunk", () => {
		expect(
			diffHunks(doc(["a", "b", "c", "d", "e"]), doc(["A", "b", "c", "d", "E"])),
		).toHaveLength(1);
	});

	test("a change on the first line clamps context at the file start", () => {
		const hunk = diffHunks("a\nb\nc\n", "A\nb\nc\n")[0];
		expect(hunk?.oldStart).toBe(1);
	});
});

describe("diffStats", () => {
	test("counts hunks, additions, and removals", () => {
		const stats = diffStats(diffHunks("a\nb\nc\n", "a\nX\nc\nY\n"));
		expect(stats).toEqual({ hunks: 1, added: 2, removed: 1 });
	});

	test("identical texts are all zero", () => {
		expect(diffStats(diffHunks("a\n", "a\n"))).toEqual({
			hunks: 0,
			added: 0,
			removed: 0,
		});
	});
});

describe("renderUnified", () => {
	test("empty string when identical", () => {
		expect(renderUnified(diffHunks("a\n", "a\n"))).toBe("");
	});

	test("emits ---/+++ headers and @@ hunks", () => {
		const out = renderUnified(diffHunks("a\nb\n", "a\nc\n"), {
			oldLabel: "OLD",
			newLabel: "NEW",
		});
		expect(out).toContain("--- OLD");
		expect(out).toContain("+++ NEW");
		expect(out).toContain("@@");
	});

	test("word-refines an isolated changed line pair", () => {
		const out = renderUnified(
			diffHunks("the quick brown fox\n", "the quick red fox\n"),
		);
		expect(out).toContain("the quick [-brown-]{+red+} fox");
	});

	test("word-refines equal-length adjacent change runs pairwise", () => {
		const out = renderUnified(
			diffHunks(
				doc(["ctx", "row one A", "row two A", "tail"]),
				doc(["ctx", "row one B", "row two B", "tail"]),
			),
		);
		expect(out).toContain("row one [-A-]{+B+}");
		expect(out).toContain("row two [-A-]{+B+}");
	});

	test("unequal-length change runs stay plain (rows added)", () => {
		const changed = renderUnified(
			diffHunks(doc(["ctx", "X", "tail"]), doc(["ctx", "Y", "Z", "tail"])),
		)
			.split("\n")
			.filter(
				(l) => /^[+-]/.test(l) && !l.startsWith("+++") && !l.startsWith("---"),
			);
		expect(changed).toEqual(["-X", "+Y", "+Z"]);
	});

	test("wordDiff:false gives plain +/- lines", () => {
		const out = renderUnified(diffHunks("old line\n", "new line\n"), {
			wordDiff: false,
		});
		expect(out).toContain("-old line");
		expect(out).toContain("+new line");
		expect(out).not.toContain("[-");
	});

	test("keeps HTML tags atomic — never splits a tag mid-markup", () => {
		const out = renderUnified(
			diffHunks("x <mark>OLD</mark> y\n", "x <mark>NEW</mark> y\n"),
		);
		expect(out).toContain("x <mark>[-OLD-]{+NEW+}</mark> y");
	});

	test("preserves a bare < in a word-refined line (not part of a tag)", () => {
		const out = renderUnified(
			diffHunks(
				doc(["ctx", "value 5 < 10 x", "tail"]),
				doc(["ctx", "value 5 < 20 x", "tail"]),
			),
		);
		expect(out).toContain("value 5 < [-10-]{+20+} x");
	});
});
