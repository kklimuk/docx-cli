import { describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { parsePagesSpec } from "../../src/cli/render/parse-pages";
import { runCli, tempWorkspace } from "./harness";

/** True when LibreOffice is installed (the docx → PDF half of the
 * pipeline). The PDF → PNG half is bundled via `@hyzyla/pdfium`, so it's
 * always available. End-to-end tests gate on this so they auto-skip on
 * machines without LibreOffice (matches the LibreOffice round-trip
 * integration suite's pattern). Pure-logic tests (--pages parsing,
 * USAGE errors) always run. */
const LIBREOFFICE_AVAILABLE = await detectLibreOffice();

async function detectLibreOffice(): Promise<boolean> {
	const probe = async (cmd: string[]): Promise<boolean> => {
		const proc = Bun.spawn(cmd, { stdout: "pipe", stderr: "pipe" });
		return (await proc.exited) === 0;
	};
	return (
		(await probe(["/bin/sh", "-c", "command -v soffice"])) ||
		existsSync("/Applications/LibreOffice.app/Contents/MacOS/soffice")
	);
}

describe("docx render — argument parsing (no engine needed)", () => {
	test("missing FILE → USAGE error", async () => {
		const result = await runCli("render");
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("non-existent FILE → FILE_NOT_FOUND", async () => {
		const result = await runCli("render", "/does/not/exist.docx");
		expect(result.exitCode).toBeGreaterThan(0);
		expect(result.parsed).toMatchObject({ code: "FILE_NOT_FOUND" });
	});

	test("--dpi out of range → USAGE", async () => {
		const path = join(tempWorkspace("render-dpi"), "x.docx");
		await runCli("create", path, "--text", "X");
		const result = await runCli("render", path, "--dpi", "10");
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--format bogus → USAGE", async () => {
		const path = join(tempWorkspace("render-fmt"), "x.docx");
		await runCli("create", path, "--text", "X");
		const result = await runCli("render", path, "--format", "tiff");
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--pages bogus → USAGE", async () => {
		const path = join(tempWorkspace("render-pages"), "x.docx");
		await runCli("create", path, "--text", "X");
		const result = await runCli("render", path, "--pages", "abc");
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("--engine bogus → USAGE", async () => {
		const path = join(tempWorkspace("render-engine"), "x.docx");
		await runCli("create", path, "--text", "X");
		const result = await runCli("render", path, "--engine", "abiword");
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});
});

describe("docx render — pages spec parser", () => {
	test("single page", () => {
		expect(parsePagesSpec("3")).toEqual({ first: 3, last: 3 });
	});
	test("range", () => {
		expect(parsePagesSpec("1-5")).toEqual({ first: 1, last: 5 });
	});
	test("descending → error", () => {
		expect(parsePagesSpec("5-1")).toMatch(/descending/);
	});
	test("discontinuous → not-yet-supported error", () => {
		expect(parsePagesSpec("1,3,5")).toMatch(/discontinuous/);
	});
	test("non-positive → error", () => {
		expect(parsePagesSpec("0")).toMatch(/positive/);
	});
	test("empty → error", () => {
		expect(parsePagesSpec("")).toMatch(/empty/);
	});
});

describe.skipIf(!LIBREOFFICE_AVAILABLE)(
	"docx render — end-to-end (libreoffice)",
	() => {
		test("renders a fixture into one PNG per page", async () => {
			const workspace = tempWorkspace("render-e2e");
			const docPath = join(workspace, "out.docx");
			const outDir = join(workspace, "pages");
			await runCli("create", docPath, "--text", "Page one.");
			const result = await runCli(
				"render",
				docPath,
				"--engine",
				"libreoffice",
				"--out",
				outDir,
			);
			expect(result.exitCode).toBe(0);
			const ack = result.parsed as {
				ok: boolean;
				engine: string;
				output: string;
				pages: string[];
			};
			expect(ack.ok).toBe(true);
			expect(ack.engine).toBe("libreoffice");
			expect(ack.output).toBe(outDir);
			expect(ack.pages.length).toBeGreaterThanOrEqual(1);
			for (const pagePath of ack.pages) {
				expect(existsSync(pagePath)).toBe(true);
				expect(pagePath).toMatch(/page-\d{3}\.png$/);
			}
		});

		test("--pages range produces a subset", async () => {
			// Build a multi-page doc by inserting page breaks.
			const workspace = tempWorkspace("render-pages-subset");
			const docPath = join(workspace, "out.docx");
			const outDir = join(workspace, "pages");
			await runCli("create", docPath, "--text", "Page one.");
			await runCli("insert", docPath, "--after", "p0", "--page-break");
			await runCli("insert", docPath, "--after", "p1", "--text", "Page two.");
			await runCli("insert", docPath, "--after", "p2", "--page-break");
			await runCli("insert", docPath, "--after", "p3", "--text", "Page three.");

			const result = await runCli(
				"render",
				docPath,
				"--engine",
				"libreoffice",
				"--out",
				outDir,
				"--pages",
				"1-2",
			);
			const ack = result.parsed as { ok: boolean; pages: string[] };
			expect(ack.ok).toBe(true);
			expect(ack.pages.length).toBe(2);
		});

		test("--format jpg emits jpgs", async () => {
			const workspace = tempWorkspace("render-jpg");
			const docPath = join(workspace, "out.docx");
			const outDir = join(workspace, "pages");
			await runCli("create", docPath, "--text", "JPG test.");
			const result = await runCli(
				"render",
				docPath,
				"--engine",
				"libreoffice",
				"--out",
				outDir,
				"--format",
				"jpg",
			);
			const ack = result.parsed as { ok: boolean; pages: string[] };
			expect(ack.ok).toBe(true);
			expect(ack.pages[0]).toMatch(/\.jpg$/);
		});
	},
);
