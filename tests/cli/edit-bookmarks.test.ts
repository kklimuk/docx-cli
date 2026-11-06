import { describe, expect, test } from "bun:test";
import { XmlNode } from "@core/parser";
import { runCli } from "./harness";
import {
	buildRawDoc,
	expectValidClean,
	freshFixture,
	readDocumentXml,
} from "./helpers";

// Issue #40 table 2: a whole-paragraph `edit` used to drop the paragraph's
// bookmark (markdown), or its `<w:bookmarkEnd>` alone when that sat inside a
// `<w:hyperlink>` the `--text` rebuild removed — silently breaking every
// cross-reference to it ("Error! Bookmark not defined." in print).

const ACM = "tests/fixtures/multi-column.docx";
// p103 (reference [8]) holds `_Ref279752219` (w:id 13), cited from p37 + p93;
// its bookmarkEnd sits inside the paragraph's trailing `<w:hyperlink>`.
const CITED = "_Ref279752219";

/** Body-paragraph index (among `<w:body>`'s direct `<w:p>` children) holding
 *  each half of bookmark `name`, or -1 for a missing half. */
async function bookmarkSpan(
	path: string,
	name: string,
): Promise<{
	start: number;
	end: number;
	startCount: number;
	endCount: number;
}> {
	const root = XmlNode.parse(await readDocumentXml(path));
	const body = XmlNode.findRoot(root, "w:document")?.findChild("w:body");
	const paragraphs = (body?.children ?? []).filter(
		(child) => child.tag === "w:p",
	);
	const starts = paragraphs.flatMap((paragraph, index) =>
		descendants(paragraph, "w:bookmarkStart")
			.filter((marker) => marker.getAttribute("w:name") === name)
			.map((marker) => ({ index, id: marker.getAttribute("w:id") })),
	);
	const id = starts[0]?.id;
	const ends = paragraphs.flatMap((paragraph, index) =>
		descendants(paragraph, "w:bookmarkEnd")
			.filter((marker) => marker.getAttribute("w:id") === id)
			.map(() => index),
	);
	return {
		start: starts[0]?.index ?? -1,
		end: ends[0] ?? -1,
		startCount: starts.length,
		endCount: ends.length,
	};
}

function descendants(node: XmlNode, tag: string): XmlNode[] {
	return node.children.flatMap((child) => [
		...(child.tag === tag ? [child] : []),
		...descendants(child, tag),
	]);
}

describe("whole-paragraph edit keeps the paragraph's bookmarks", () => {
	test("--markdown rebuild keeps a cited bookmark paired on the paragraph", async () => {
		const path = await freshFixture("bm-markdown", ACM);
		const before = await bookmarkSpan(path, CITED);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p103",
			"Psy. 2013. *Gangnam Style*. Video.",
		);
		expect(result.exitCode).toBe(0);
		const after = await bookmarkSpan(path, CITED);
		expect(after).toEqual(before);
		expect(after.start).toBe(after.end);
		await expectValidClean(path);
	});

	test("--text keeps the bookmarkEnd that lived inside the rebuilt hyperlink", async () => {
		const path = await freshFixture("bm-text", ACM);
		const before = await bookmarkSpan(path, CITED);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p103",
			"--text",
			"Psy. 2013. Gangnam Style. Video.",
		);
		expect(result.exitCode).toBe(0);
		expect(await bookmarkSpan(path, CITED)).toEqual(before);
		await expectValidClean(path);
	});

	test("the bookmark brackets the new text, so the edit reads unchanged", async () => {
		const path = await buildRawDoc(
			'<w:p><w:bookmarkStart w:id="7" w:name="Methods"/><w:r><w:t>Old methods.</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
			"bm-bracket",
		);
		const result = await runCli("edit", path, "--at", "p0", "New **methods**.");
		expect(result.exitCode).toBe(0);
		const xml = await readDocumentXml(path);
		expect(xml).toMatch(
			/<w:bookmarkStart w:id="7" w:name="Methods"\/>.*New .*methods.*<w:bookmarkEnd w:id="7"\/><\/w:p>/,
		);
		const read = await runCli("read", path);
		expect(read.stdout).toContain("New **methods**.");
	});

	test("a cross-paragraph bookmark keeps its half on the edited paragraph", async () => {
		const path = await buildRawDoc(
			'<w:p><w:bookmarkStart w:id="3" w:name="Span"/><w:r><w:t>First.</w:t></w:r></w:p><w:p><w:r><w:t>Second.</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>',
			"bm-cross",
		);
		const result = await runCli(
			"edit",
			path,
			"--at",
			"p0",
			"*First*, revised.",
		);
		expect(result.exitCode).toBe(0);
		const span = await bookmarkSpan(path, "Span");
		expect(span).toEqual({ start: 0, end: 1, startCount: 1, endCount: 1 });
	});

	for (const verdict of ["accept", "reject"] as const) {
		test(`tracked --markdown edit keeps the bookmark after ${verdict}`, async () => {
			const path = await freshFixture(`bm-tracked-${verdict}`, ACM);
			const before = await bookmarkSpan(path, CITED);
			const edited = await runCli(
				"edit",
				path,
				"--track",
				"--at",
				"p103",
				"Psy. 2013. *Gangnam Style*. Video.",
			);
			expect(edited.exitCode).toBe(0);
			expect(await bookmarkSpan(path, CITED)).toEqual(before);
			const settled = await runCli("track-changes", verdict, path, "--all");
			expect(settled.exitCode).toBe(0);
			expect(await bookmarkSpan(path, CITED)).toEqual(before);
			await expectValidClean(path);
		});
	}
});
