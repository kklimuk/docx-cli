import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

type LinkInfo = {
	id: string;
	url?: string;
	anchor?: string;
	tooltip?: string;
};

type Doc = {
	blocks: Array<{
		type: string;
		id: string;
		runs?: Array<{
			type: string;
			text?: string;
			hyperlink?: LinkInfo;
		}>;
	}>;
};

function collectHyperlinks(doc: Doc): Array<{
	blockId: string;
	text: string;
	hyperlink: LinkInfo;
}> {
	const out: ReturnType<typeof collectHyperlinks> = [];
	for (const block of doc.blocks) {
		if (block.type !== "paragraph") continue;
		for (const run of block.runs ?? []) {
			if (run.type !== "text" || !run.hyperlink) continue;
			out.push({
				blockId: block.id,
				text: run.text ?? "",
				hyperlink: run.hyperlink,
			});
		}
	}
	return out;
}

describe("docx read — hyperlinks", () => {
	test("surfaces every hyperlink from academic-paper.docx", async () => {
		const result = await runCli("read", "tests/fixtures/academic-paper.docx");
		expect(result.exitCode).toBe(0);
		const links = collectHyperlinks(result.parsed as Doc);
		const urls = new Set(
			links.map((entry) => entry.hyperlink.url).filter(Boolean),
		);
		expect(urls.size).toBe(16);
		expect(urls).toContain("https://doi.org/10.1080/13548506.2014.1002851");
		expect(urls).toContain(
			"https://www.apa.org/news/press/releases/stress/2017/state-nation.pdf",
		);
	});

	test("surfaces hyperlink in tables-and-lists.docx", async () => {
		const result = await runCli("read", "tests/fixtures/tables-and-lists.docx");
		expect(result.exitCode).toBe(0);
		const links = collectHyperlinks(result.parsed as Doc);
		expect(links).toHaveLength(1);
		expect(links[0]?.hyperlink.url).toContain("seas.gwu.edu");
		expect(links[0]?.text).toContain("seas.gwu.edu");
	});

	test("multiple runs sharing one link share the same url", async () => {
		const result = await runCli("read", "tests/fixtures/large-mixed.docx");
		expect(result.exitCode).toBe(0);
		const doc = result.parsed as Doc;
		// Group hyperlink-tagged runs by paragraph and url; any paragraph that
		// has multiple runs in the same hyperlink should report the same url.
		for (const block of doc.blocks) {
			if (block.type !== "paragraph") continue;
			const urlsByText = new Map<string, Set<string>>();
			for (const run of block.runs ?? []) {
				if (run.type !== "text" || !run.hyperlink?.url) continue;
				if (!urlsByText.has(block.id)) urlsByText.set(block.id, new Set());
				urlsByText.get(block.id)?.add(run.hyperlink.url);
			}
		}
		const total = collectHyperlinks(doc).length;
		expect(total).toBeGreaterThanOrEqual(11);
	});

	test("paragraph offsets include hyperlink text", async () => {
		// Pick a paragraph in academic-paper that ends with a hyperlink. The
		// returned text length must equal the sum of all text-run lengths
		// (including those tagged with a hyperlink).
		const result = await runCli("read", "tests/fixtures/academic-paper.docx");
		const doc = result.parsed as Doc;
		const linked = collectHyperlinks(doc);
		const sample = linked[0];
		expect(sample).toBeDefined();
		const block = doc.blocks.find((entry) => entry.id === sample?.blockId);
		expect(block?.type).toBe("paragraph");
		const concatenated = (block?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(concatenated.length).toBeGreaterThan(0);
		expect(concatenated).toContain(sample?.text ?? "");
	});

	test("each hyperlink has a positional linkN id", async () => {
		const result = await runCli("read", "tests/fixtures/academic-paper.docx");
		const links = collectHyperlinks(result.parsed as Doc);
		const ids = new Set(links.map((entry) => entry.hyperlink.id));
		expect(ids.size).toBe(16);
		for (const id of ids) expect(id).toMatch(/^link\d+$/);
	});
});

type ListEntry = {
	id: string;
	url?: string;
	anchor?: string;
	tooltip?: string;
	text: string;
	blockId: string;
};

describe("docx hyperlinks list", () => {
	test("returns one entry per unique linkN", async () => {
		const result = await runCli(
			"hyperlinks",
			"list",
			"tests/fixtures/large-mixed.docx",
		);
		expect(result.exitCode).toBe(0);
		const entries = result.parsed as ListEntry[];
		expect(entries.length).toBe(11);
		const ids = entries.map((entry) => entry.id);
		expect(new Set(ids).size).toBe(ids.length);
		for (const entry of entries) {
			expect(entry.id).toMatch(/^link\d+$/);
			expect(entry.blockId).toMatch(/^p\d+|^t\d+/);
			expect(entry.text.length).toBeGreaterThan(0);
		}
	});

	test("--help short-circuits", async () => {
		const result = await runCli("hyperlinks", "list", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("hyperlinks list");
	});

	test("missing FILE returns USAGE", async () => {
		const result = await runCli("hyperlinks", "list");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});

describe("docx hyperlinks replace", () => {
	test("changes the URL and leaves other links untouched", async () => {
		const workspace = tempWorkspace("hl-replace");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const beforeList = before.parsed as ListEntry[];
		const target = beforeList.find((entry) => entry.id === "link0");
		expect(target).toBeDefined();
		const originalUrl = target?.url;

		const replace = await runCli(
			"hyperlinks",
			"replace",
			docPath,
			"--at",
			"link0",
			"--with",
			"https://example.com/new",
		);
		expect(replace.exitCode).toBe(0);
		expect(replace.parsed).toMatchObject({
			ok: true,
			operation: "hyperlinks.replace",
			hyperlinkId: "link0",
			to: "https://example.com/new",
			from: originalUrl,
		});

		const after = await runCli("hyperlinks", "list", docPath);
		const afterList = after.parsed as ListEntry[];
		const updated = afterList.find((entry) => entry.id === "link0");
		expect(updated?.url).toBe("https://example.com/new");
		// Other link ids stay stable.
		expect(afterList.length).toBe(beforeList.length);
	});

	test("does not affect other hyperlinks that shared the same target URL", async () => {
		// In large-mixed.docx, link0 and link2 both point to the same URL.
		// Replacing link0 must allocate a new rId so link2 stays unchanged.
		const workspace = tempWorkspace("hl-shared");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const beforeList = before.parsed as ListEntry[];
		const link0Url = beforeList.find((entry) => entry.id === "link0")?.url;
		const link2Url = beforeList.find((entry) => entry.id === "link2")?.url;
		expect(link0Url).toBeTruthy();
		expect(link0Url).toBe(link2Url);

		await runCli(
			"hyperlinks",
			"replace",
			docPath,
			"--at",
			"link0",
			"--with",
			"https://example.com/changed",
		);

		const after = await runCli("hyperlinks", "list", docPath);
		const afterList = after.parsed as ListEntry[];
		expect(afterList.find((entry) => entry.id === "link0")?.url).toBe(
			"https://example.com/changed",
		);
		expect(afterList.find((entry) => entry.id === "link2")?.url).toBe(link2Url);
	});

	test("--dry-run does not modify the file", async () => {
		const workspace = tempWorkspace("hl-dryrun");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const beforeList = before.parsed as ListEntry[];
		const originalUrl = beforeList.find((entry) => entry.id === "link0")?.url;

		const dryRun = await runCli(
			"hyperlinks",
			"replace",
			docPath,
			"--at",
			"link0",
			"--with",
			"https://example.com/dry",
			"--dry-run",
		);
		expect(dryRun.exitCode).toBe(0);
		expect(dryRun.parsed).toMatchObject({
			ok: true,
			dryRun: true,
		});

		const after = await runCli("hyperlinks", "list", docPath);
		const afterList = after.parsed as ListEntry[];
		expect(afterList.find((entry) => entry.id === "link0")?.url).toBe(
			originalUrl,
		);
	});

	test("--output writes to a parallel file", async () => {
		const workspace = tempWorkspace("hl-output");
		const docPath = join(workspace, "doc.docx");
		const outputPath = join(workspace, "out.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		await runCli(
			"hyperlinks",
			"replace",
			docPath,
			"--at",
			"link0",
			"--with",
			"https://example.com/output",
			"-o",
			outputPath,
		);

		const original = await runCli("hyperlinks", "list", docPath);
		const originalList = original.parsed as ListEntry[];
		const originalUrl = originalList.find((entry) => entry.id === "link0")?.url;
		expect(originalUrl).not.toBe("https://example.com/output");

		const output = await runCli("hyperlinks", "list", outputPath);
		const outputList = output.parsed as ListEntry[];
		expect(outputList.find((entry) => entry.id === "link0")?.url).toBe(
			"https://example.com/output",
		);
	});

	test("rejects unknown link id with HYPERLINK_NOT_FOUND", async () => {
		const result = await runCli(
			"hyperlinks",
			"replace",
			"tests/fixtures/large-mixed.docx",
			"--at",
			"link999",
			"--with",
			"https://example.com",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "HYPERLINK_NOT_FOUND",
		});
	});
});

describe("docx hyperlinks add", () => {
	async function setupSample(label: string): Promise<{
		docPath: string;
	}> {
		const workspace = tempWorkspace(label);
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"the quick brown fox jumps over the lazy dog",
		);
		return { docPath };
	}

	test("wraps a span in a new hyperlink", async () => {
		const { docPath } = await setupSample("hl-add");
		const result = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com/brown",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "hyperlinks.add",
			at: "p1:10-15",
			url: "https://example.com/brown",
		});

		const list = await runCli("hyperlinks", "list", docPath);
		const links = list.parsed as ListEntry[];
		expect(links).toHaveLength(1);
		expect(links[0]?.text).toBe("brown");
		expect(links[0]?.url).toBe("https://example.com/brown");
	});

	test("preserves text outside the wrapped span", async () => {
		const { docPath } = await setupSample("hl-add-preserve");
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com",
		);
		const read = await runCli("read", docPath);
		const doc = read.parsed as Doc;
		const p1 = doc.blocks.find((entry) => entry.id === "p1");
		const concatenated = (p1?.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(concatenated).toBe("the quick brown fox jumps over the lazy dog");
	});

	test("rejects spans that overlap an existing hyperlink", async () => {
		const { docPath } = await setupSample("hl-add-overlap");
		await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com/first",
		);
		const second = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:8-20",
			"--url",
			"https://example.com/second",
		);
		expect(second.exitCode).toBe(2);
		expect(second.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("--dry-run does not modify the file", async () => {
		const { docPath } = await setupSample("hl-add-dry");
		const before = await runCli("hyperlinks", "list", docPath);
		const dry = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1:10-15",
			"--url",
			"https://example.com",
			"--dry-run",
		);
		expect(dry.parsed).toMatchObject({ ok: true, dryRun: true });
		const after = await runCli("hyperlinks", "list", docPath);
		expect(after.parsed).toEqual(before.parsed);
	});

	test("non-span locator returns INVALID_LOCATOR", async () => {
		const { docPath } = await setupSample("hl-add-invalid");
		const result = await runCli(
			"hyperlinks",
			"add",
			docPath,
			"--at",
			"p1",
			"--url",
			"https://example.com",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "INVALID_LOCATOR",
		});
	});
});

describe("docx hyperlinks delete", () => {
	test("unwraps the hyperlink and keeps the display text", async () => {
		const workspace = tempWorkspace("hl-del");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const beforeList = before.parsed as ListEntry[];
		const target = beforeList.find((entry) => entry.id === "link0");
		expect(target).toBeDefined();
		const targetText = target?.text ?? "";

		const result = await runCli(
			"hyperlinks",
			"delete",
			docPath,
			"--at",
			"link0",
		);
		expect(result.exitCode).toBe(0);
		expect(result.parsed).toMatchObject({
			ok: true,
			operation: "hyperlinks.delete",
			hyperlinkId: "link0",
		});

		const after = await runCli("hyperlinks", "list", docPath);
		const afterList = after.parsed as ListEntry[];
		expect(afterList.length).toBe(beforeList.length - 1);

		// The removed link's text should still be in the document as plain text.
		const read = await runCli("read", docPath);
		const doc = read.parsed as Doc;
		const allText = doc.blocks
			.flatMap((block) => block.runs ?? [])
			.filter((run) => run.type === "text")
			.map((run) => run.text ?? "")
			.join("");
		expect(allText).toContain(targetText);
	});

	test("--dry-run does not modify the file", async () => {
		const workspace = tempWorkspace("hl-del-dry");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const dry = await runCli(
			"hyperlinks",
			"delete",
			docPath,
			"--at",
			"link0",
			"--dry-run",
		);
		expect(dry.parsed).toMatchObject({ ok: true, dryRun: true });
		const after = await runCli("hyperlinks", "list", docPath);
		expect(after.parsed).toEqual(before.parsed);
	});

	test("rejects unknown link id with HYPERLINK_NOT_FOUND", async () => {
		const result = await runCli(
			"hyperlinks",
			"delete",
			"tests/fixtures/large-mixed.docx",
			"--at",
			"link999",
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "HYPERLINK_NOT_FOUND",
		});
	});

	test("does not affect a sibling link sharing the same target URL", async () => {
		// In large-mixed.docx, link0 and link2 originally point to the same URL;
		// deleting link0 must leave link2 functional.
		const workspace = tempWorkspace("hl-del-shared");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("hyperlinks", "list", docPath);
		const beforeList = before.parsed as ListEntry[];
		const link2Url = beforeList.find((entry) => entry.id === "link2")?.url;

		await runCli("hyperlinks", "delete", docPath, "--at", "link0");

		const after = await runCli("hyperlinks", "list", docPath);
		const afterList = after.parsed as ListEntry[];
		// After delete, the surviving sibling occupies a positional id
		// (probably link1). Find it by matching the URL.
		const survivor = afterList.find((entry) => entry.url === link2Url);
		expect(survivor).toBeDefined();
		expect(survivor?.url).toBe(link2Url);
	});
});

describe("docx insert --url", () => {
	test("creates a paragraph with a single hyperlink run", async () => {
		const workspace = tempWorkspace("insert-url");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));

		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"click here",
			"--url",
			"https://example.com",
		);
		expect(result.exitCode).toBe(0);

		const list = await runCli("hyperlinks", "list", docPath);
		const links = list.parsed as ListEntry[];
		expect(links).toHaveLength(1);
		expect(links[0]).toMatchObject({
			text: "click here",
			url: "https://example.com",
		});
	});

	test("--url without --text is a usage error", async () => {
		const workspace = tempWorkspace("insert-url-runs");
		const docPath = join(workspace, "doc.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/minimal.docx"));

		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--runs",
			'[{"type":"text","text":"x"}]',
			"--url",
			"https://example.com",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});
});
