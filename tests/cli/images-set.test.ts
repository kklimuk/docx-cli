import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { runCli, spawnCli, tempWorkspace } from "./harness";
import { expectValidClean, readDocumentXml, readMarkdown } from "./helpers";

const FIXTURES = join(import.meta.dir, "..", "fixtures");
const PNG = join(FIXTURES, "assets", "sample.png"); // 96x64 px → 1x0.67in

type SetAck = {
	operation: string;
	imageId?: string;
	count?: number;
	images: Array<{
		imageId: string;
		size: string;
		widthEmu: number;
		heightEmu: number;
		alt: string;
	}>;
};

/** A Letter doc (6.5in text width) with `p0` text and the image at `p1`,
 *  sized `widthInches` wide (3:2 aspect, alt "Old"). */
async function docWithImage(widthInches: string): Promise<string> {
	const doc = join(tempWorkspace("images-set"), "doc.docx");
	await runCli("create", doc, "--text", "Hello");
	const added = await runCli(
		"images",
		"add",
		doc,
		"--after",
		"p0",
		"--image",
		PNG,
		"--alt",
		"Old",
		"--width",
		widthInches,
	);
	expect(added.exitCode).toBe(0);
	return doc;
}

function extents(xml: string): string[] {
	return [...xml.matchAll(/<(wp|a):ext(?:ent)? cx="(\d+)" cy="(\d+)"/g)].map(
		(match) => `${match[1]}:${match[2]}x${match[3]}`,
	);
}

describe("docx images set — resize", () => {
	let doc: string;
	beforeEach(async () => {
		doc = await docWithImage("3");
	});

	test("--width keeps the aspect ratio and updates wp:extent AND a:xfrm/a:ext", async () => {
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"1.5in",
		);
		expect(result.exitCode).toBe(0);
		const ack = result.parsed as SetAck;
		expect(ack.operation).toBe("images.set");
		expect(ack.images[0]?.size).toBe("1.5x1in");
		expect(extents(await readDocumentXml(doc))).toEqual([
			"wp:1371600x914400",
			"a:1371600x914400",
		]);
		// Write-read loop: the docx:image hint shows the new size.
		expect(await readMarkdown(doc)).toContain(
			'<!-- docx:image img0 size="1.5x1in" -->',
		);
		await expectValidClean(doc);
	});

	test("--height keeps the aspect ratio; units parse through the shared grammar", async () => {
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--height",
			"2.54cm",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as SetAck).images[0]?.size).toBe("1.5x1in");
	});

	test("--size accepts the docx:image size hint verbatim (both sides exact)", async () => {
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--size",
			"2x0.5in",
		);
		expect(result.exitCode).toBe(0);
		expect(await readMarkdown(doc)).toContain('docx:image img0 size="2x0.5in"');
	});

	test("the default ack names the image and its new size hint", async () => {
		const result = await spawnCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"2",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("images.set img0\n");
		expect(result.stdout).toContain('img0 size="2x1.33in" alt="Old"');
	});

	test("one size option at a time; bad lengths are usage errors", async () => {
		const both = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"2",
			"--fit",
		);
		expect(both.exitCode).toBe(2);
		const bad = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"big",
		);
		expect(bad.exitCode).toBe(2);
		const nothing = await runCli("images", "set", doc, "--at", "img0");
		expect(nothing.exitCode).toBe(2);
	});

	test("an unknown image id is IMAGE_NOT_FOUND and nothing is written", async () => {
		const before = await readDocumentXml(doc);
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img7",
			"--width",
			"1",
		);
		expect(result.exitCode).toBe(3);
		expect(result.stdout).toContain("IMAGE_NOT_FOUND");
		expect(await readDocumentXml(doc)).toBe(before);
	});
});

describe("docx images set --fit", () => {
	test("shrinks an overflowing image to the text width and clears overflow=yes", async () => {
		const doc = await docWithImage("9");
		expect(await readMarkdown(doc)).toContain('size="9x6in" overflow="yes"');
		const result = await runCli("images", "set", doc, "--at", "img0", "--fit");
		expect(result.exitCode).toBe(0);
		const md = await readMarkdown(doc);
		expect(md).toContain('<!-- docx:image img0 size="6.5x4.33in" -->');
		expect(md).not.toContain("overflow=");
	});

	test("never enlarges: --fit on an in-bounds image changes nothing and exits nonzero", async () => {
		const doc = await docWithImage("2");
		const before = await readDocumentXml(doc);
		const result = await runCli("images", "set", doc, "--at", "img0", "--fit");
		expect(result.exitCode).toBe(3);
		expect(result.stdout).toContain("MATCH_NOT_FOUND");
		expect(await readDocumentXml(doc)).toBe(before);
	});

	test("resizes a Word-authored floating image in place, keeping the anchor", async () => {
		const doc = join(tempWorkspace("images-set-float"), "doc.docx");
		await Bun.write(doc, Bun.file(join(FIXTURES, "large-mixed.docx")));
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img1",
			"--width",
			"3in",
		);
		expect(result.exitCode).toBe(0);
		const md = await readMarkdown(doc);
		expect(md).toMatch(
			/docx:image img1 size="3x[\d.]+in" float="yes" wrap="topAndBottom" align="left"/,
		);
	});
});

describe("docx images set --alt", () => {
	test("sets docPr/cNvPr descr, shows in read, and survives an edit that restates the image", async () => {
		const doc = await docWithImage("3");
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--alt",
			"Company logo",
			"--width",
			"2",
		);
		expect(result.exitCode).toBe(0);
		const xml = await readDocumentXml(doc);
		expect(xml).toMatch(/<wp:docPr [^>]*descr="Company logo"/);
		expect(xml).toMatch(/<pic:cNvPr [^>]*descr="Company logo"/);

		const line = (await readMarkdown(doc))
			.split("\n")
			.find((text) => text.includes("docx:image img0")) as string;
		expect(line).toContain("![Company logo](");
		expect(line).toContain('size="2x1.33in"');

		// A markdown edit that restates the image keeps its alt AND its size —
		// it must not snap back to the native pixel size.
		const restated = line.replace(/ <!--.*$/, "");
		const edit = await runCli("edit", doc, "--at", "p1", `Logo: ${restated}`);
		expect(edit.exitCode).toBe(0);
		const after = await readMarkdown(doc);
		expect(after).toContain("Logo: ![Company logo](");
		expect(after).toContain('docx:image img0 size="2x1.33in"');
	});

	test('--alt "" clears the description', async () => {
		const doc = await docWithImage("1");
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--alt",
			"",
		);
		expect(result.exitCode).toBe(0);
		expect(await readDocumentXml(doc)).not.toContain('descr="Old"');
	});

	test("re-applying the current alt is a no-op error", async () => {
		const doc = await docWithImage("1");
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--alt",
			"Old",
		);
		expect(result.exitCode).toBe(3);
	});
});

describe("docx images set --batch", () => {
	test("applies every entry from one read; ack counts the changed images", async () => {
		const doc = await docWithImage("9");
		await runCli(
			"images",
			"add",
			doc,
			"--after",
			"p1",
			"--image",
			PNG,
			"--width",
			"3",
		);
		const batch = join(tempWorkspace("images-set-batch"), "sets.jsonl");
		await Bun.write(
			batch,
			[
				JSON.stringify({ at: "img0", fit: true, alt: "Wide chart" }),
				JSON.stringify({ at: "img1", width: 1.5 }),
			].join("\n"),
		);
		const result = await runCli("images", "set", doc, "--batch", batch);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as SetAck).count).toBe(2);
		const md = await readMarkdown(doc);
		expect(md).toContain("![Wide chart](");
		expect(md).toContain('docx:image img0 size="6.5x4.33in"');
		expect(md).toContain('docx:image img1 size="1.5x1in"');
	});

	test("a bad entry anywhere writes nothing; CLI flags can't ride along", async () => {
		const doc = await docWithImage("3");
		const before = await readDocumentXml(doc);
		const bad = await runCli(
			"images",
			"set",
			doc,
			"--batch",
			`${JSON.stringify({ at: "img0", width: 1 })}\n${JSON.stringify({ at: "img9", width: 1 })}`,
		);
		expect(bad.exitCode).toBe(3);
		expect(bad.stdout).toContain("entry 1");
		expect(await readDocumentXml(doc)).toBe(before);

		const twice = await runCli(
			"images",
			"set",
			doc,
			"--batch",
			`${JSON.stringify({ at: "img0", width: 1 })}\n${JSON.stringify({ at: "img0", alt: "x" })}`,
		);
		expect(twice.exitCode).toBe(2);
		expect(await readDocumentXml(doc)).toBe(before);

		const mixed = await runCli(
			"images",
			"set",
			doc,
			"--batch",
			'{"at":"img0","width":1}',
			"--alt",
			"x",
		);
		expect(mixed.exitCode).toBe(2);
	});
});

describe("docx images set under track-changes", () => {
	test("applies untracked and leaves a [docx-cli] audit comment on the image", async () => {
		const doc = await docWithImage("3");
		await runCli("track-changes", "on", doc);
		const result = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"1.5",
			"--author",
			"Ed",
		);
		expect(result.exitCode).toBe(0);
		const comments = (await runCli("comments", "list", doc)).parsed as Array<{
			text: string;
			author: string;
		}>;
		expect(comments).toHaveLength(1);
		expect(comments[0]?.text).toContain(
			"[docx-cli] image resized: 3x2in → 1.5x1in",
		);
		expect(comments[0]?.author).toBe("Ed");
		expect((await runCli("track-changes", "list", doc)).parsed).toEqual([]);
	});
});

describe("docx images set --dry-run / -o", () => {
	test("--dry-run previews without writing; -o writes a parallel file", async () => {
		const doc = await docWithImage("3");
		const before = await readDocumentXml(doc);
		const dry = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"1",
			"--dry-run",
		);
		expect(dry.exitCode).toBe(0);
		expect((dry.parsed as { dryRun?: boolean }).dryRun).toBe(true);
		expect(await readDocumentXml(doc)).toBe(before);

		const out = join(tempWorkspace("images-set-out"), "out.docx");
		const parallel = await runCli(
			"images",
			"set",
			doc,
			"--at",
			"img0",
			"--width",
			"1",
			"-o",
			out,
		);
		expect(parallel.exitCode).toBe(0);
		expect(await readDocumentXml(doc)).toBe(before);
		expect(await readMarkdown(out)).toContain('size="1x0.67in"');
	});
});
