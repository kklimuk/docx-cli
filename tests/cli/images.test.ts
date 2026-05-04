import { beforeAll, describe, expect, test } from "bun:test";
import { mkdirSync } from "node:fs";
import { join } from "node:path";
import { runCli, tempWorkspace } from "./harness";

describe("docx images", () => {
	let docPath: string;
	let workspace: string;

	beforeAll(async () => {
		workspace = tempWorkspace("images");
		docPath = join(workspace, "with-images.docx");
		await Bun.write(docPath, Bun.file("tests/fixtures/large-mixed.docx"));
	});

	test("list returns all 15 images with hashes and dimensions", async () => {
		const result = await runCli("images", "list", docPath);
		expect(result.exitCode).toBe(0);
		const images = result.parsed as Array<{
			id: string;
			hash: string;
			contentType: string;
			widthEmu?: number;
		}>;
		expect(images).toHaveLength(15);
		for (const image of images) {
			expect(image.id).toMatch(/^img\d+$/);
			expect(image.hash).toMatch(/^[a-f0-9]{64}$/);
			expect(image.contentType).toBeTruthy();
		}
	});

	test("extract --to writes hashed files and returns manifest", async () => {
		const outDir = join(workspace, "extracted");
		mkdirSync(outDir, { recursive: true });
		const result = await runCli("images", "extract", docPath, "--to", outDir);
		const payload = result.parsed as {
			manifest: Array<{ id: string; path: string; bytes: number }>;
		};
		expect(payload.manifest).toHaveLength(15);
		for (const entry of payload.manifest) {
			const file = Bun.file(entry.path);
			expect(await file.exists()).toBe(true);
			expect(entry.bytes).toBeGreaterThan(0);
		}
	});

	test("extract --id pulls a single image only", async () => {
		const outDir = join(workspace, "single");
		mkdirSync(outDir, { recursive: true });
		const result = await runCli(
			"images",
			"extract",
			docPath,
			"--to",
			outDir,
			"--id",
			"img2",
		);
		const payload = result.parsed as {
			manifest: Array<{ id: string }>;
		};
		expect(payload.manifest).toHaveLength(1);
		expect(payload.manifest[0]?.id).toBe("img2");
	});

	test("replace swaps bytes preserving the partName for matching format", async () => {
		const replaceWorkspace = tempWorkspace("replace");
		const replaceDoc = join(replaceWorkspace, "doc.docx");
		await Bun.write(replaceDoc, Bun.file("tests/fixtures/large-mixed.docx"));

		const before = await runCli("images", "list", replaceDoc);
		const beforeList = before.parsed as Array<{ id: string; hash: string }>;
		const target = beforeList.find((image) => image.id === "img0");
		expect(target).toBeDefined();

		const replacementSrc = beforeList.find(
			(image) => image.id !== "img0" && image.hash !== target?.hash,
		);
		expect(replacementSrc).toBeDefined();
		const extractDir = join(replaceWorkspace, "extracted");
		mkdirSync(extractDir, { recursive: true });
		await runCli(
			"images",
			"extract",
			replaceDoc,
			"--to",
			extractDir,
			"--id",
			replacementSrc?.id ?? "img1",
		);
		const replacementPath = join(extractDir, `${replacementSrc?.hash}.jpg`);
		expect(await Bun.file(replacementPath).exists()).toBe(true);

		const result = await runCli(
			"images",
			"replace",
			replaceDoc,
			"--at",
			"img0",
			"--with",
			replacementPath,
		);
		expect(result.exitCode).toBe(0);

		const after = await runCli("images", "list", replaceDoc);
		const afterList = after.parsed as Array<{ id: string; hash: string }>;
		const updated = afterList.find((image) => image.id === "img0");
		expect(updated?.hash).toBe(replacementSrc?.hash);
	});

	test("replace rejects bad image MIME types", async () => {
		const fakeWorkspace = tempWorkspace("fake");
		const fakeDoc = join(fakeWorkspace, "doc.docx");
		await Bun.write(fakeDoc, Bun.file("tests/fixtures/large-mixed.docx"));
		const fakeFile = join(fakeWorkspace, "fake.txt");
		await Bun.write(fakeFile, "not an image");
		const result = await runCli(
			"images",
			"replace",
			fakeDoc,
			"--at",
			"img0",
			"--with",
			fakeFile,
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("replace returns image-not-found for unknown id", async () => {
		const extractDir = join(workspace, "for-not-found");
		mkdirSync(extractDir, { recursive: true });
		await runCli(
			"images",
			"extract",
			docPath,
			"--to",
			extractDir,
			"--id",
			"img0",
		);
		const realImage = (
			await Array.fromAsync(new Bun.Glob("*.jpg").scan({ cwd: extractDir }))
		)[0];
		expect(realImage).toBeDefined();
		const result = await runCli(
			"images",
			"replace",
			docPath,
			"--at",
			"img99",
			"--with",
			join(extractDir, realImage ?? ""),
		);
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "IMAGE_NOT_FOUND",
		});
	});
});
