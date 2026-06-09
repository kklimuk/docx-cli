import { beforeAll, describe, expect, test } from "bun:test";
import { mkdirSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { Document, XmlNode } from "@core";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { readMarkdown, trackedKinds } from "./helpers";

const FIXTURES = join(import.meta.dir, "..", "fixtures");

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
		const manifest = result.parsed as Array<{
			id: string;
			path: string;
			bytes: number;
		}>;
		expect(manifest).toHaveLength(15);
		for (const entry of manifest) {
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
			"--at",
			"img2",
		);
		const manifest = result.parsed as Array<{ id: string }>;
		expect(manifest).toHaveLength(1);
		expect(manifest[0]?.id).toBe("img2");
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
			"--at",
			replacementSrc?.id ?? "img1",
		);
		const replacementPath = join(extractDir, `${replacementSrc?.hash}.jpeg`);
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
		expect(result.parsed).toMatchObject({ code: "USAGE" });
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
			"--at",
			"img0",
		);
		const realImage = (
			await Array.fromAsync(new Bun.Glob("*.jpeg").scan({ cwd: extractDir }))
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
			code: "IMAGE_NOT_FOUND",
		});
	});
});

const ASSETS = join(import.meta.dir, "..", "fixtures", "assets");
const PNG_PATH = join(ASSETS, "sample.png"); // 96×64
const JPG_PATH = join(ASSETS, "sample.jpg"); // 120×80
const HEIC_PATH = join(ASSETS, "sample.heic"); // 96×64, transcoded to JPEG
const EMU_PER_PIXEL = 9525;
const EMU_PER_INCH = 914400;

type ImageRun = {
	type: string;
	id: string;
	contentType: string;
	widthEmu?: number;
	heightEmu?: number;
	alt?: string;
	trackedChange?: { kind: string };
};

async function imageRuns(docPath: string): Promise<ImageRun[]> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as {
		blocks: Array<{ runs?: ImageRun[] }>;
	};
	return doc.blocks
		.flatMap((block) => block.runs ?? [])
		.filter((run) => run.type === "image");
}

async function newDoc(label: string): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", "Before");
	return docPath;
}

async function dataUri(path: string, mime: string): Promise<string> {
	const bytes = await Bun.file(path).bytes();
	return `data:${mime};base64,${Buffer.from(bytes).toString("base64")}`;
}

describe("docx insert --image", () => {
	test("embeds a PNG from a file path: media part, relationship, content-type", async () => {
		const docPath = await newDoc("img-png");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			PNG_PATH,
			"--alt",
			"A sample",
		);
		expect(result.exitCode).toBe(0);

		const pkg = await Pkg.open(docPath);
		const media = pkg
			.listParts()
			.filter((part) => part.startsWith("word/media/"));
		expect(media).toContain("word/media/image1.png");

		const rels = await pkg.readText("word/_rels/document.xml.rels");
		expect(rels).toContain(
			'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',
		);
		expect(rels).toContain('Target="media/image1.png"');

		const contentTypes = await pkg.readText("[Content_Types].xml");
		expect(contentTypes).toContain(
			'<Default Extension="png" ContentType="image/png"/>',
		);
	});

	test("reads back native pixel dimensions and alt text", async () => {
		const docPath = await newDoc("img-dims");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			PNG_PATH,
			"--alt",
			"A sample",
		);
		const [image] = await imageRuns(docPath);
		expect(image?.contentType).toBe("image/png");
		expect(image?.widthEmu).toBe(96 * EMU_PER_PIXEL);
		expect(image?.heightEmu).toBe(64 * EMU_PER_PIXEL);
		expect(image?.alt).toBe("A sample");
	});

	test("embeds a JPEG from a data: URI", async () => {
		const docPath = await newDoc("img-datauri");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			await dataUri(JPG_PATH, "image/jpeg"),
		);
		const pkg = await Pkg.open(docPath);
		expect(pkg.listParts()).toContain("word/media/image1.jpeg");
		const [image] = await imageRuns(docPath);
		expect(image?.contentType).toBe("image/jpeg");
		expect(image?.widthEmu).toBe(120 * EMU_PER_PIXEL);
		expect(image?.heightEmu).toBe(80 * EMU_PER_PIXEL);
	});

	test("--width alone scales height to preserve aspect ratio", async () => {
		const docPath = await newDoc("img-width");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			PNG_PATH,
			"--width",
			"1.5",
		);
		const [image] = await imageRuns(docPath);
		// 96×64 at 1.5" wide → 1.0" tall.
		expect(image?.widthEmu).toBe(Math.round(1.5 * EMU_PER_INCH));
		expect(image?.heightEmu).toBe(EMU_PER_INCH);
	});

	test("--width and --height together are both honored verbatim", async () => {
		const docPath = await newDoc("img-wh");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			PNG_PATH,
			"--width",
			"2",
			"--height",
			"3",
		);
		const [image] = await imageRuns(docPath);
		expect(image?.widthEmu).toBe(2 * EMU_PER_INCH);
		expect(image?.heightEmu).toBe(3 * EMU_PER_INCH);
	});

	test("two inserts mint distinct part names and relationship ids", async () => {
		const docPath = await newDoc("img-two");
		await runCli("insert", docPath, "--after", "p0", "--image", PNG_PATH);
		await runCli("insert", docPath, "--after", "p1", "--image", JPG_PATH);
		const pkg = await Pkg.open(docPath);
		const media = pkg
			.listParts()
			.filter((part) => part.startsWith("word/media/image"));
		expect(media.sort()).toEqual([
			"word/media/image1.png",
			"word/media/image2.jpeg",
		]);
		expect((await imageRuns(docPath)).length).toBe(2);
	});

	test("transcodes a HEIC file to JPEG before embedding", async () => {
		const docPath = await newDoc("img-heic");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			HEIC_PATH,
			"--alt",
			"Student photo",
		);
		expect(result.exitCode).toBe(0);

		// The embedded part is a JPEG, not HEIC (Word can't render HEIC).
		const pkg = await Pkg.open(docPath);
		expect(pkg.listParts()).toContain("word/media/image1.jpeg");
		expect(pkg.listParts().some((part) => part.endsWith(".heic"))).toBe(false);

		const [image] = await imageRuns(docPath);
		expect(image?.contentType).toBe("image/jpeg");
		// Dimensions survive the transcode (source HEIC is 96×64).
		expect(image?.widthEmu).toBe(96 * EMU_PER_PIXEL);
		expect(image?.heightEmu).toBe(64 * EMU_PER_PIXEL);
	});

	test("transcodes HEIC supplied as a data: URI", async () => {
		const docPath = await newDoc("img-heic-datauri");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			await dataUri(HEIC_PATH, "image/heic"),
		);
		const pkg = await Pkg.open(docPath);
		expect(pkg.listParts()).toContain("word/media/image1.jpeg");
		expect((await imageRuns(docPath))[0]?.contentType).toBe("image/jpeg");
	});

	test("refuses to fetch from a loopback/private address (SSRF gate)", async () => {
		// An agent steered into `http://localhost:.../` or the cloud-metadata
		// endpoint must not turn this CLI into an outbound proxy. Verify both the
		// hostname-that-resolves-to-loopback case and the IP-literal case.
		const pngBytes = await Bun.file(PNG_PATH).bytes();
		const server = Bun.serve({
			port: 0,
			fetch: () =>
				new Response(pngBytes, { headers: { "content-type": "image/png" } }),
		});
		try {
			const docPath = await newDoc("img-http-private");
			const loopback = await runCli(
				"insert",
				docPath,
				"--after",
				"p0",
				"--image",
				`http://localhost:${server.port}/logo.png`,
			);
			expect(loopback.exitCode).toBe(1);
			expect((loopback.parsed as { code: string }).code).toBe("IMAGE_SOURCE");

			const metadata = await runCli(
				"insert",
				docPath,
				"--after",
				"p0",
				"--image",
				"http://169.254.169.254/latest/meta-data/",
			);
			expect(metadata.exitCode).toBe(1);
			expect((metadata.parsed as { code: string }).code).toBe("IMAGE_SOURCE");
		} finally {
			await server.stop(true);
		}
	});

	test("a missing file is a clean IMAGE_SOURCE error", async () => {
		const docPath = await newDoc("img-missing");
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			join(ASSETS, "does-not-exist.png"),
		);
		expect(result.exitCode).toBe(1);
		expect((result.parsed as { code: string }).code).toBe("IMAGE_SOURCE");
	});

	test("an image with no parseable dimensions requires explicit sizing", async () => {
		const svg =
			'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>';
		const svgUri = `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
		const docPath = await newDoc("img-svg");

		const noDims = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			svgUri,
		);
		expect(noDims.exitCode).toBe(2);
		expect((noDims.parsed as { code: string }).code).toBe("USAGE");

		const withDims = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			svgUri,
			"--width",
			"2",
			"--height",
			"2",
		);
		expect(withDims.exitCode).toBe(0);
		const pkg = await Pkg.open(docPath);
		expect(pkg.listParts()).toContain("word/media/image1.svg");
	});

	test("dry-run reports without writing a media part", async () => {
		const docPath = await newDoc("img-dry");
		const before = (await Pkg.open(docPath)).listParts().length;
		const result = await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--image",
			PNG_PATH,
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		const after = (await Pkg.open(docPath)).listParts().length;
		expect(after).toBe(before);
	});

	test("inserting an image under track-changes marks the run as inserted", async () => {
		const docPath = await newDoc("img-tracked");
		await runCli("track-changes", docPath, "on");
		await runCli("insert", docPath, "--after", "p0", "--image", PNG_PATH);
		const [image] = await imageRuns(docPath);
		expect(image?.trackedChange?.kind).toBe("ins");
	});
});

async function mediaParts(docPath: string): Promise<string[]> {
	const pkg = await Pkg.open(docPath);
	return pkg
		.listParts()
		.filter((part) => part.startsWith("word/media/image"))
		.sort();
}

async function docWithImages(label: string, count: number): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	await runCli("create", docPath, "--text", "Before");
	for (let index = 0; index < count; index++) {
		await runCli(
			"insert",
			docPath,
			"--after",
			`p${index}`,
			"--image",
			index % 2 === 0 ? PNG_PATH : JPG_PATH,
		);
	}
	return docPath;
}

describe("docx images delete", () => {
	test("removes the image run and prunes the now-unreferenced part", async () => {
		const docPath = await docWithImages("del-one", 1);
		expect(await mediaParts(docPath)).toEqual(["word/media/image1.png"]);

		const result = await runCli("images", "delete", docPath, "--at", "img0");
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { pruned: boolean }).pruned).toBe(true);

		expect(await imageRuns(docPath)).toHaveLength(0);
		expect(await mediaParts(docPath)).toEqual([]);

		const pkg = await Pkg.open(docPath);
		const rels = await pkg.readText("word/_rels/document.xml.rels");
		expect(rels).not.toContain("/relationships/image");
	});

	test("keeps the part when the rId is still referenced by non-drawing content", async () => {
		// Word commonly references the same image rId from a VML fallback our
		// drawing walker can't see. Deleting the drawing must NOT prune the part,
		// or the surviving reference would dangle and corrupt the document.
		const docPath = await docWithImages("del-vml", 1);
		const document = await Document.open(docPath);
		const rId = document.body.imageById.get("img0")?.relationshipId;
		expect(rId).toBeDefined();
		const body = XmlNode.findRoot(
			document.documentTree,
			"w:document",
		)?.findChild("w:body");
		body?.children.push(
			new XmlNode("w:p", {}, [
				new XmlNode("w:r", {}, [
					new XmlNode("w:pict", {}, [
						new XmlNode("v:imagedata", { "r:id": rId as string }),
					]),
				]),
			]),
		);
		await document.save();

		const result = await runCli("images", "delete", docPath, "--at", "img0");
		expect((result.parsed as { pruned: boolean }).pruned).toBe(false);

		const pkg = await Pkg.open(docPath);
		expect(await mediaParts(docPath)).toEqual(["word/media/image1.png"]);
		expect(await pkg.readText("word/_rels/document.xml.rels")).toContain(
			rId as string,
		);
	});

	test("deletes only the targeted occurrence, keeping the others", async () => {
		const docPath = await docWithImages("del-multi", 3);
		expect(await imageRuns(docPath)).toHaveLength(3);

		await runCli("images", "delete", docPath, "--at", "img1");
		const remaining = await imageRuns(docPath);
		expect(remaining).toHaveLength(2);
		// img1 was the JPEG; the two PNGs survive.
		expect(await mediaParts(docPath)).toEqual([
			"word/media/image1.png",
			"word/media/image3.png",
		]);
	});

	test("deleting one image leaves the other images' parts intact", async () => {
		const docPath = await docWithImages("del-others", 2);
		await runCli("images", "delete", docPath, "--at", "img0");
		expect(await mediaParts(docPath)).toEqual(["word/media/image2.jpeg"]);
		expect(await imageRuns(docPath)).toHaveLength(1);
	});

	test("unknown id is a clean IMAGE_NOT_FOUND", async () => {
		const docPath = await docWithImages("del-missing", 1);
		const result = await runCli("images", "delete", docPath, "--at", "img9");
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("IMAGE_NOT_FOUND");
	});

	test("dry-run reports without removing the part", async () => {
		const docPath = await docWithImages("del-dry", 1);
		const result = await runCli(
			"images",
			"delete",
			docPath,
			"--at",
			"img0",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await mediaParts(docPath)).toEqual(["word/media/image1.png"]);
		expect(await imageRuns(docPath)).toHaveLength(1);
	});

	describe("under track-changes", () => {
		test("wraps the run in a tracked deletion, keeping the part", async () => {
			const docPath = await docWithImages("del-tracked", 1);
			await runCli("track-changes", docPath, "on");
			await runCli("images", "delete", docPath, "--at", "img0");

			// The image still exists, now marked as a tracked deletion.
			const [image] = await imageRuns(docPath);
			expect(image?.trackedChange?.kind).toBe("del");
			// Media part is retained until the change is accepted.
			expect(await mediaParts(docPath)).toEqual(["word/media/image1.png"]);
		});

		test("accepting the deletion removes the image", async () => {
			const docPath = await docWithImages("del-accept", 1);
			await runCli("track-changes", docPath, "on");
			await runCli("images", "delete", docPath, "--at", "img0");
			await runCli("track-changes", "accept", docPath, "--all");
			expect(await imageRuns(docPath)).toHaveLength(0);
		});

		test("rejecting the deletion restores the image", async () => {
			const docPath = await docWithImages("del-reject", 1);
			await runCli("track-changes", docPath, "on");
			await runCli("images", "delete", docPath, "--at", "img0");
			await runCli("track-changes", "reject", docPath, "--all");
			const restored = await imageRuns(docPath);
			expect(restored).toHaveLength(1);
			expect(restored[0]?.trackedChange).toBeUndefined();
		});
	});
});

// The "matching format" replace above keeps the part name; replacing with a
// DIFFERENT extension exercises the part-rename + rel-retarget + Content_Types
// branch, which is exactly where a dangling rId would corrupt the file.
describe("docx images replace — format/extension change (PNG → JPG)", () => {
	test("renames the media part, rewrites the rel Target + [Content_Types], leaves no dangling rId", async () => {
		const docPath = await newDoc("img-replace-ext");
		await runCli("insert", docPath, "--after", "p0", "--image", PNG_PATH);
		expect(await mediaParts(docPath)).toEqual(["word/media/image1.png"]);

		const result = await runCli(
			"images",
			"replace",
			docPath,
			"--at",
			"img0",
			"--with",
			JPG_PATH,
		);
		expect(result.exitCode).toBe(0);

		// The media part is renamed to the new extension; the .png is gone.
		expect(await mediaParts(docPath)).toEqual(["word/media/image1.jpeg"]);

		const pkg = await Pkg.open(docPath);
		const rels = await pkg.readText("word/_rels/document.xml.rels");
		expect(rels).toContain('Target="media/image1.jpeg"');
		expect(rels).not.toContain("image1.png");

		const contentTypes = await pkg.readText("[Content_Types].xml");
		expect(contentTypes).toMatch(/Extension="jpe?g"/);

		// Every image rId referenced by the body still resolves to a relationship.
		const documentXml = await pkg.readText("word/document.xml");
		for (const match of documentXml.matchAll(/r:embed="(rId\d+)"/g)) {
			expect(rels).toContain(`Id="${match[1]}"`);
		}

		// The AST and the list verb both report the new format.
		const [image] = await imageRuns(docPath);
		expect(image?.contentType).toBe("image/jpeg");
		const list = await runCli("images", "list", docPath);
		expect(
			(list.parsed as Array<{ contentType: string }>)[0]?.contentType,
		).toBe("image/jpeg");
	});
});

describe("docx images — -o parallel write", () => {
	test("images delete -o writes to the output and leaves the source byte-unchanged", async () => {
		const src = await newDoc("img-o");
		await runCli("insert", src, "--after", "p0", "--image", PNG_PATH);
		const before = await Bun.file(src).bytes();
		const out = join(tempWorkspace("img-o-out"), "out.docx");

		const result = await runCli(
			"images",
			"delete",
			src,
			"--at",
			"img0",
			"-o",
			out,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(out);
		expect(await Bun.file(src).bytes()).toEqual(before);

		// The image is gone from the output but still present in the source.
		expect((await runCli("images", "list", out)).parsed).toEqual([]);
		expect(
			((await runCli("images", "list", src)).parsed as unknown[]).length,
		).toBe(1);
	});
});

// The "under track-changes" cases above drive tracking via the global toggle;
// the per-invocation --track flag (toggle OFF) is a distinct code path.
describe("docx images delete — --track forces tracking with the toggle off", () => {
	test("--track records a tracked deletion and keeps the media part until accept", async () => {
		const docPath = await newDoc("img-delete-track");
		await runCli("insert", docPath, "--after", "p0", "--image", PNG_PATH);

		const result = await runCli(
			"images",
			"delete",
			docPath,
			"--at",
			"img0",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(docPath)).toContain("del");
		// The drawing is wrapped in <w:del>, not removed — the media part stays
		// until the deletion is accepted.
		expect(await mediaParts(docPath)).toHaveLength(1);
	});
});

// Read-side visibility: `![](hash)` proves an image exists but not "6in wide,
// past the margin". A trailing docx:image note carries size (always) +
// float/wrap/align/overflow (deviation-only). Read-time hints; importer drops.
describe("images surface size + placement as docx:image hints", () => {
	test("an inline image trails size in inches; a small in-bounds one shows only size", async () => {
		const md = await readMarkdown(join(FIXTURES, "images.docx"));
		expect(md).toContain('<!-- docx:image img0 size="1x0.67in" -->');
		// A ~1in inline image is in-bounds and inline → no float/wrap/align/overflow.
		expect(md).not.toMatch(/docx:image img0[^>]*(float|wrap|align|overflow)/);
	});

	test("a floating image surfaces float/wrap/align (markdown + --ast)", async () => {
		const path = join(FIXTURES, "large-mixed.docx");
		const md = await readMarkdown(path);
		expect(md).toMatch(
			/docx:image img0[^>]*float="yes"[^>]*wrap="topAndBottom"[^>]*align="absolute"/,
		);
		const ast = (await runCli("read", path, "--ast")).parsed as {
			blocks: Array<{ runs?: Array<Record<string, unknown>> }>;
		};
		const image = ast.blocks
			.flatMap((b) => b.runs ?? [])
			.find((r) => r.type === "image");
		expect(image?.floating).toBe(true);
		expect(typeof image?.wrap).toBe("string");
		expect(typeof image?.align).toBe("string");
	});

	test("an image wider than the text column is flagged overflow=yes", async () => {
		// images.docx's ~1in images are in bounds → no overflow.
		expect(await readMarkdown(join(FIXTURES, "images.docx"))).not.toContain(
			"overflow=",
		);

		// Insert a 10in-wide image into a default-Letter doc (~6.5in usable).
		const workspace = tempWorkspace("img-overflow");
		const media = join(workspace, "media");
		await runCli(
			"images",
			"extract",
			join(FIXTURES, "images.docx"),
			"--to",
			media,
		);
		const png = readdirSync(media).find((f) => f.endsWith(".png")) as string;
		const mdPath = join(workspace, "s.md");
		await Bun.write(mdPath, "# D\n\nbody\n");
		const doc = join(workspace, "d.docx");
		await runCli("create", doc, "--from", mdPath);
		await runCli(
			"insert",
			doc,
			"--after",
			"p1",
			"--image",
			join(media, png),
			"--width",
			"10",
		);
		expect(await readMarkdown(doc)).toMatch(/docx:image[^>]*overflow="yes"/);
	});
});
