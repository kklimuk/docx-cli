import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Document, XmlNode } from "../../src/core";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

const ASSETS = join(import.meta.dir, "..", "fixtures", "assets");
const PNG_PATH = join(ASSETS, "sample.png");
const JPG_PATH = join(ASSETS, "sample.jpg");

type ImageRun = { type: string; trackedChange?: { kind: string } };

async function imageRuns(docPath: string): Promise<ImageRun[]> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as { blocks: Array<{ runs?: ImageRun[] }> };
	return doc.blocks
		.flatMap((block) => block.runs ?? [])
		.filter((run) => run.type === "image");
}

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
