import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { normalizeReadMarkers } from "../../src/cli/diff/markers";
import { runCli, spawnCliStdin, tempWorkspace } from "./harness";
import { freshFixture } from "./helpers";

const MINIMAL = "tests/fixtures/minimal.docx";

/** The snapshot workflow the command is built around: a pristine baseline copy
 *  next to an editable working copy. Returns [workingPath, baselinePath]. */
async function snapshot(label: string): Promise<[string, string]> {
	const working = await freshFixture(label, MINIMAL);
	const baseline = join(tempWorkspace(`${label}-base`), "orig.docx");
	await Bun.write(baseline, Bun.file(MINIMAL));
	return [working, baseline];
}

describe("docx diff — basic", () => {
	test("shows an edit as a change, with baseline/current labels and a stat header", async () => {
		const [working, baseline] = await snapshot("diff-basic");
		// A single-token marker so it appears intact in the word-refined line
		// (`{+ZZTOPMARKER+}`) rather than split across paired words.
		await runCli("edit", working, "--at", "p0", "--text", "ZZTOPMARKER");
		const res = await runCli("diff", working, "--against", baseline);
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("ZZTOPMARKER");
		expect(res.stdout).toContain("(baseline)");
		expect(res.stdout).toContain("(current)");
		expect(res.stdout).toMatch(/^# \d+ hunk/m);
	});

	test("no differences → message on stdout, exit 0", async () => {
		const [working, baseline] = await snapshot("diff-nodiff");
		const res = await runCli("diff", working, "--against", baseline);
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("No differences");
	});

	test("self-compare (same path) is called out plainly", async () => {
		const [working] = await snapshot("diff-self");
		const res = await runCli("diff", working, "--against", working);
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("to itself");
	});

	test("--json emits structured hunks and stats", async () => {
		const [working, baseline] = await snapshot("diff-json");
		await runCli("edit", working, "--at", "p0", "--text", "JSON CHANGE");
		const res = await runCli("diff", working, "--against", baseline, "--json");
		expect(res.exitCode).toBe(0);
		const parsed = res.parsed as {
			hunks: unknown[];
			stats: { hunks: number; added: number; removed: number };
		};
		expect(Array.isArray(parsed.hunks)).toBe(true);
		expect(parsed.stats.hunks).toBeGreaterThan(0);
		expect(parsed.stats.added + parsed.stats.removed).toBeGreaterThan(0);
	});
});

describe("docx diff — baseline input forms", () => {
	test("compares against a saved `docx read` text file", async () => {
		const [working, baseline] = await snapshot("diff-textfile");
		const savedRead = join(tempWorkspace("diff-savedread"), "old.md");
		await Bun.write(savedRead, (await runCli("read", baseline)).stdout);
		await runCli("edit", working, "--at", "p0", "--text", "SAVEDREADMARKER");
		const res = await runCli("diff", working, "--against", savedRead);
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("SAVEDREADMARKER");
	});

	test("reads a saved read output from stdin (`--against -`)", async () => {
		const [working, baseline] = await snapshot("diff-stdin");
		const savedRead = (await runCli("read", baseline)).stdout;
		await runCli("edit", working, "--at", "p0", "--text", "STDINMARKER");
		const res = await spawnCliStdin(
			savedRead,
			"diff",
			working,
			"--against",
			"-",
		);
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("STDINMARKER");
	});
});

describe("docx diff — errors & guards", () => {
	test("missing FILE → USAGE (exit 2)", async () => {
		const res = await runCli("diff");
		expect(res.exitCode).toBe(2);
		expect((res.parsed as { code: string }).code).toBe("USAGE");
	});

	test("missing --against → USAGE (exit 2)", async () => {
		const [working] = await snapshot("diff-noagainst");
		const res = await runCli("diff", working);
		expect(res.exitCode).toBe(2);
		expect((res.parsed as { code: string }).code).toBe("USAGE");
	});

	test("`-` as FILE is rejected (only --against reads stdin)", async () => {
		const res = await runCli("diff", "-", "--against", "x.docx");
		expect(res.exitCode).toBe(2);
		expect((res.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--against file not found → FILE_NOT_FOUND (exit 3)", async () => {
		const [working] = await snapshot("diff-missing-base");
		const res = await runCli("diff", working, "--against", "no-such.docx");
		expect(res.exitCode).toBe(3);
		expect((res.parsed as { code: string }).code).toBe("FILE_NOT_FOUND");
	});

	test("--from with a raw-text --against → USAGE (can't slice text)", async () => {
		const [working] = await snapshot("diff-from-text");
		const savedRead = join(tempWorkspace("diff-from-textmd"), "old.md");
		await Bun.write(savedRead, "some read text\n");
		const res = await runCli(
			"diff",
			working,
			"--against",
			savedRead,
			"--from",
			"p0",
		);
		expect(res.exitCode).toBe(2);
		expect((res.parsed as { code: string }).code).toBe("USAGE");
	});

	test("--help exits 0 and leads with the snapshot workflow", async () => {
		const res = await runCli("diff", "--help");
		expect(res.exitCode).toBe(0);
		expect(res.stdout).toContain("cp doc.docx doc.orig.docx");
		expect(res.stdout).toContain("not a patch");
	});
});

describe("normalizeReadMarkers (anti-thrash normalization)", () => {
	test("strips bare locator comments entirely", () => {
		expect(
			normalizeReadMarkers("hi <!-- p3 --> there <!-- t2:r5c0:p0 -->"),
		).toBe("hi there");
	});

	test("reduces a docx:cell locator to its stable table anchor, keeping attributes", () => {
		expect(
			normalizeReadMarkers('x <!-- docx:cell t2:r5c0 shading="FFFF00" -->'),
		).toBe('x <!-- docx:cell t2 shading="FFFF00" -->');
	});

	test("drops a leading paragraph token in a docx:p annotation", () => {
		expect(normalizeReadMarkers('<!-- docx:p p7 style="Heading1" -->')).toBe(
			'<!-- docx:p style="Heading1" -->',
		);
	});

	test("keeps stable section and table anchors", () => {
		const kept = '<!-- docx:section s0 --> <!-- docx:table t1 widths="1in" -->';
		expect(normalizeReadMarkers(kept)).toBe(kept);
	});

	test("a row renumber (r2→r3) normalizes to the SAME string — the core anti-thrash guarantee", () => {
		const beforeRow =
			'A <!-- docx:cell t2:r2c0 borders="single" --> B <!-- t2:r2c1:p0 -->';
		const afterRow =
			'A <!-- docx:cell t2:r3c0 borders="single" --> B <!-- t2:r3c1:p0 -->';
		expect(normalizeReadMarkers(beforeRow)).toBe(
			normalizeReadMarkers(afterRow),
		);
	});

	test("a shading change survives normalization (formatting deltas still diff)", () => {
		const beforeFmt = '<!-- docx:cell t2:r5c0 borders="single" -->';
		const afterFmt =
			'<!-- docx:cell t2:r5c0 shading="FFFF00" borders="single" -->';
		expect(normalizeReadMarkers(beforeFmt)).not.toBe(
			normalizeReadMarkers(afterFmt),
		);
	});
});
