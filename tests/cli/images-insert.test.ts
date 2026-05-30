import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

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
