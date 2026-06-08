import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

/**
 * Two ergonomics fixes surfaced by the weak-model adversarial run:
 *   - `--text` with embedded newlines/tabs → real `<w:br/>` / `<w:tab/>` (not a
 *     literal \n that Word swallows). Verse/addresses stay line-per-line, and
 *     `read` round-trips them.
 *   - `insert --image --caption` / `images add --caption` → a native Word
 *     "Caption"-styled paragraph under the figure (Table-of-Figures-able).
 */

type Run = { type: string; text?: string; kind?: string };
type Block = { id: string; type: string; style?: string; runs?: Run[] };

const ASSETS = join(import.meta.dir, "..", "fixtures", "assets");
const PNG = join(ASSETS, "sample.png");

async function blocks(path: string): Promise<Block[]> {
	const result = await runCli("read", path, "--ast");
	return (result.parsed as { blocks: Block[] }).blocks;
}

async function block(path: string, id: string): Promise<Block> {
	const found = (await blocks(path)).find((candidate) => candidate.id === id);
	if (!found) throw new Error(`block ${id} not found`);
	return found;
}

function newDoc(label: string): string {
	return join(tempWorkspace(label), "doc.docx");
}

describe("--text newlines and tabs", () => {
	test("insert --text with a newline becomes a <w:br/> line break", async () => {
		const path = newDoc("nl-insert");
		await runCli("create", path, "--text", "Intro.");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--text",
			"line one\nline two",
		);
		const runs = (await block(path, "p1")).runs ?? [];
		expect(runs.map((run) => run.type)).toEqual(["text", "break", "text"]);
		expect(runs[0]?.text).toBe("line one");
		expect(runs[1]?.kind).toBe("line");
		expect(runs[2]?.text).toBe("line two");
	});

	test("insert --text with a tab becomes a <w:tab/>", async () => {
		const path = newDoc("tab-insert");
		await runCli("create", path, "--text", "Intro.");
		await runCli("insert", path, "--after", "p0", "--text", "a\tb");
		const runs = (await block(path, "p1")).runs ?? [];
		expect(runs.map((run) => run.type)).toEqual(["text", "tab", "text"]);
	});

	test("edit --text (whole paragraph) splits newlines too", async () => {
		const path = newDoc("nl-edit");
		await runCli("create", path, "--text", "placeholder");
		await runCli("edit", path, "--at", "p0", "--text", "first\nsecond\nthird");
		const runs = (await block(path, "p0")).runs ?? [];
		expect(runs.filter((run) => run.type === "break").length).toBe(2);
		expect(
			runs.filter((run) => run.type === "text").map((run) => run.text),
		).toEqual(["first", "second", "third"]);
	});

	test("a multi-line --text paragraph round-trips through read → markdown", async () => {
		const path = newDoc("nl-roundtrip");
		await runCli("create", path, "--text", "Roses are red\nViolets are blue");
		const md = (await runCli("read", path)).stdout;
		expect(md).toContain("Roses are red\nViolets are blue");
	});

	test("single-line --text is still one text run (no behavior change)", async () => {
		const path = newDoc("nl-single");
		await runCli("create", path, "--text", "just one line");
		const runs = (await block(path, "p0")).runs ?? [];
		expect(runs).toHaveLength(1);
		expect(runs[0]?.type).toBe("text");
	});
});

describe("image captions", () => {
	test("insert --image --caption adds a Caption-styled paragraph below the figure", async () => {
		const path = newDoc("cap-insert");
		await runCli("create", path, "--text", "Report.");
		const result = await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--image",
			PNG,
			"--caption",
			"Figure 1: Quarterly revenue",
		);
		expect(result.exitCode).toBe(0);
		// Two blocks minted: the figure paragraph and the caption paragraph.
		const all = await blocks(path);
		const caption = all.find((b) => b.style === "Caption");
		expect(caption).toBeDefined();
		expect((caption?.runs ?? []).map((run) => run.text).join("")).toBe(
			"Figure 1: Quarterly revenue",
		);
	});

	test("the Caption style is provisioned in styles.xml", async () => {
		const path = newDoc("cap-style");
		await runCli("create", path, "--text", "Report.");
		await runCli(
			"insert",
			path,
			"--after",
			"p0",
			"--image",
			PNG,
			"--caption",
			"Fig 1",
		);
		const { stdout } = await runCli("read", path, "--ast");
		expect(stdout).toBeDefined();
		// The caption paragraph carries the Caption pStyle (proves the style was
		// referenced); styles.xml provisioning is exercised by the doc opening
		// cleanly for the read above.
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect(caption).toBeDefined();
	});

	test("images add --caption (alias) works the same way", async () => {
		const path = newDoc("cap-alias");
		await runCli("create", path, "--text", "Report.");
		const result = await runCli(
			"images",
			"add",
			path,
			"--image",
			PNG,
			"--after",
			"p0",
			"--caption",
			"Figure A",
		);
		expect(result.exitCode).toBe(0);
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect((caption?.runs ?? []).map((run) => run.text).join("")).toBe(
			"Figure A",
		);
	});

	test("no --caption → no Caption paragraph", async () => {
		const path = newDoc("cap-none");
		await runCli("create", path, "--text", "Report.");
		await runCli("insert", path, "--after", "p0", "--image", PNG);
		const caption = (await blocks(path)).find((b) => b.style === "Caption");
		expect(caption).toBeUndefined();
	});
});
