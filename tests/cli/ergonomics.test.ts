import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

const WIDE_PNG = "tests/fixtures/assets/sample.png";

async function blankDoc(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const path = join(workspace, "out.docx");
	expect((await runCli("create", path, "--text", "Body.")).exitCode).toBe(0);
	return path;
}

function extentCx(xml: string): number {
	const match = xml.match(/<wp:extent cx="(\d+)"/);
	return match ? Number(match[1]) : 0;
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
		const cx = extentCx(await readDocumentXml(path));
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
		expect(extentCx(await readDocumentXml(path))).toBe(8 * 914400);
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

async function readDocumentXml(path: string): Promise<string> {
	const proc = Bun.spawn(["unzip", "-p", path, "word/document.xml"], {
		stdout: "pipe",
	});
	return await new Response(proc.stdout).text();
}
