import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx create + read", () => {
	test("creates a minimal docx and reads it back", async () => {
		const workspace = tempWorkspace("create-read");
		const docPath = join(workspace, "out.docx");

		const create = await runCli(
			"create",
			docPath,
			"--title",
			"Test",
			"--author",
			"Tester",
			"--text",
			"Hello world",
		);
		expect(create.exitCode).toBe(0);
		expect(create.parsed).toMatchObject({
			ok: true,
			operation: "create",
			path: docPath,
		});

		const read = await runCli("read", docPath, "--ast");
		expect(read.exitCode).toBe(0);
		const doc = read.parsed as {
			properties: { title: string; author: string };
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		expect(doc.properties.title).toBe("Test");
		expect(doc.properties.author).toBe("Tester");
		const firstParagraph = doc.blocks.find(
			(block) => block.type === "paragraph",
		);
		expect(firstParagraph?.runs?.[0]?.text).toBe("Hello world");
	});

	test("escapes hostile XML in title/author/text round-trip", async () => {
		const workspace = tempWorkspace("escape");
		const docPath = join(workspace, "out.docx");

		await runCli(
			"create",
			docPath,
			"--title",
			'Has "quotes" & <stuff>',
			"--author",
			"<script>",
			"--text",
			"Body: <input> & 'data'",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as {
			properties: { title: string; author: string };
			blocks: Array<{ type: string; runs?: Array<{ text: string }> }>;
		};
		expect(doc.properties.title).toBe('Has "quotes" & <stuff>');
		expect(doc.properties.author).toBe("<script>");
		const paragraph = doc.blocks.find((block) => block.type === "paragraph");
		expect(paragraph?.runs?.[0]?.text).toBe("Body: <input> & 'data'");
	});

	test("read on missing file returns not-found error", async () => {
		const result = await runCli("read", "/tmp/does-not-exist.docx");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			code: "FILE_NOT_FOUND",
		});
	});

	test("create rejects existing file without --force", async () => {
		const workspace = tempWorkspace("force");
		const docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "first");
		const second = await runCli("create", docPath, "--text", "second");
		expect(second.exitCode).toBe(2);
		expect(second.parsed).toMatchObject({ code: "USAGE" });
	});
});

const WIDE_PNG = "tests/fixtures/assets/sample.png";

async function blankDoc(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--text", "Body.")).exitCode).toBe(0);
	return path;
}

async function imageWidthEmu(path: string): Promise<number> {
	const read = await runCli("read", path, "--ast");
	const blocks = (
		read.parsed as {
			blocks: Array<{ runs?: Array<{ type: string; widthEmu?: number }> }>;
		}
	).blocks;
	const image = blocks
		.flatMap((block) => block.runs ?? [])
		.find((run) => run.type === "image");
	return image?.widthEmu ?? 0;
}

describe("Ergonomics", () => {
	test("create writes to the positional FILE and rejects -o", async () => {
		const workspace = tempWorkspace("create-pos");
		const path = join(workspace, "a.docx");
		expect((await runCli("create", path, "--text", "Hi")).exitCode).toBe(0);
		expect(await Bun.file(path).exists()).toBe(true);
		const bad = await runCli("create", join(workspace, "b.docx"), "-o", "x");
		expect(bad.exitCode).toBe(2);
	});

	test("inserted image is clamped to the page content width", async () => {
		const path = await blankDoc("img-clamp");
		expect(
			(await runCli("insert", path, "--after", "p0", "--image", WIDE_PNG))
				.exitCode,
		).toBe(0);
		const cx = await imageWidthEmu(path);
		// US-Letter content width with 1" margins = 6.5in = 5,943,600 EMU.
		expect(cx).toBeLessThanOrEqual(5_943_600);
		expect(cx).toBeGreaterThan(0);
	});

	test("explicit --width overrides the clamp", async () => {
		const path = await blankDoc("img-width");
		expect(
			(
				await runCli(
					"insert",
					path,
					"--after",
					"p0",
					"--image",
					WIDE_PNG,
					"--width",
					"8",
				)
			).exitCode,
		).toBe(0);
		expect(await imageWidthEmu(path)).toBe(8 * 914400);
	});

	test("images add is an alias for insert --image", async () => {
		const path = await blankDoc("img-add");
		expect(
			(
				await runCli(
					"images",
					"add",
					path,
					"--image",
					WIDE_PNG,
					"--after",
					"p0",
				)
			).exitCode,
		).toBe(0);
		const list = await runCli("images", "list", path);
		expect((list.parsed as unknown[]).length).toBe(1);
	});

	test("footnotes add --anchor drops the reference after the phrase", async () => {
		const workspace = tempWorkspace("fn-anchor");
		const src = join(workspace, "s.md");
		await Bun.write(src, "Revenue was 4.2M this quarter, up sharply.\n");
		const path = join(workspace, "f.docx");
		expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
		expect(
			(
				await runCli(
					"footnotes",
					"add",
					path,
					"--anchor",
					"4.2M",
					"--text",
					"Source: close.",
				)
			).exitCode,
		).toBe(0);
		expect((await runCli("read", path)).stdout).toContain("4.2M[^fn1]");
	});

	test("footnotes add --anchor with no match errors", async () => {
		const workspace = tempWorkspace("fn-nomatch");
		const src = join(workspace, "s.md");
		await Bun.write(src, "Nothing to cite here.\n");
		const path = join(workspace, "f.docx");
		expect((await runCli("create", path, "--from", src)).exitCode).toBe(0);
		const result = await runCli(
			"footnotes",
			"add",
			path,
			"--anchor",
			"absent",
			"--text",
			"x",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("MATCH_NOT_FOUND");
	});
});

describe("docx create --text-file (literal body)", () => {
	test("seeds the body with literal paragraphs, replacing the placeholder", async () => {
		const workspace = tempWorkspace("create-literal");
		const docPath = join(workspace, "out.docx");
		const notes = join(workspace, "notes.txt");
		// GFM-hostile prose: ordered-list marker + CriticMarkup must survive verbatim.
		await Bun.write(notes, "3. Reviewer note\nSecond line {++keep++}\n");

		const create = await runCli("create", docPath, "--text-file", notes);
		expect(create.exitCode).toBe(0);
		expect(create.parsed).toMatchObject({
			ok: true,
			operation: "create",
			blocks: 2, // the seed placeholder was replaced, not appended to
		});

		const read = await runCli("read", docPath, "--ast");
		const paragraphs = (
			read.parsed as {
				blocks: Array<{ type: string; runs?: Array<{ text?: string }> }>;
			}
		).blocks.filter((block) => block.type === "paragraph");
		expect(paragraphs).toHaveLength(2);
		expect(paragraphs[0]?.runs?.map((run) => run.text ?? "").join("")).toBe(
			"3. Reviewer note",
		);
		expect(paragraphs[1]?.runs?.map((run) => run.text ?? "").join("")).toBe(
			"Second line {++keep++}",
		);
	});

	test("rejects more than one content source", async () => {
		const workspace = tempWorkspace("create-mutex");
		const docPath = join(workspace, "out.docx");
		const notes = join(workspace, "n.txt");
		await Bun.write(notes, "x");

		const result = await runCli(
			"create",
			docPath,
			"--text",
			"a",
			"--text-file",
			notes,
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("USAGE");
	});
});
