import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/package";
import { runCli, tempWorkspace } from "./harness";

/**
 * End-to-end coverage for tracked moves (<w:moveFrom>/<w:moveTo>) against
 * tests/fixtures/tracked-moves.docx. Layout:
 *   p0: "Origin paragraph: " + moveFrom("the moved sentence") + "."
 *   p1: "Destination paragraph: " + moveTo("the moved sentence") + "."
 */
const FIXTURE = "tests/fixtures/tracked-moves.docx";

type Doc = {
	blocks: Array<{
		type: string;
		runs?: Array<{
			type: string;
			text?: string;
			trackedChange?: { id: string; kind: string; author: string };
		}>;
	}>;
};

async function freshCopy(label: string): Promise<string> {
	const workspace = tempWorkspace(label);
	const docPath = join(workspace, "tracked-moves.docx");
	await Bun.write(docPath, Bun.file(FIXTURE));
	return docPath;
}

function paragraphText(doc: Doc, index: number): string {
	const block = doc.blocks[index];
	if (!block || block.type !== "paragraph") return "";
	return (block.runs ?? [])
		.filter((run) => run.type === "text")
		.map((run) => run.text ?? "")
		.join("");
}

describe("tracked moves — AST surface", () => {
	test("read exposes moveFrom and moveTo as TrackedChange entries", async () => {
		const result = await runCli("read", FIXTURE);
		const doc = result.parsed as Doc;

		const movedRunFrom = (doc.blocks[0]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);
		const movedRunTo = (doc.blocks[1]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);

		expect(movedRunFrom?.trackedChange?.kind).toBe("moveFrom");
		expect(movedRunFrom?.trackedChange?.author).toBe("Reviewer");
		expect(movedRunTo?.trackedChange?.kind).toBe("moveTo");
		expect(movedRunTo?.trackedChange?.author).toBe("Reviewer");
	});

	test("track-changes list reports both halves of the move", async () => {
		const result = await runCli("track-changes", "list", FIXTURE);
		const records = result.parsed as Array<{
			id: string;
			kind: string;
			text: string;
		}>;
		expect(records).toHaveLength(2);
		const kinds = records.map((record) => record.kind).sort();
		expect(kinds).toEqual(["moveFrom", "moveTo"]);
		// Both halves carry the same text.
		expect(
			records.every((record) => record.text === "the moved sentence"),
		).toBe(true);
	});
});

describe("tracked moves — wc views", () => {
	test("--accepted skips the moveFrom origin", async () => {
		const result = await runCli("wc", FIXTURE, "--accepted");
		// Accepted view: "Origin paragraph: ." + "Destination paragraph: the moved sentence."
		// Words: "Origin", "paragraph:", "." → "Origin paragraph: ." has 3 word-like tokens.
		// Word counter splits on whitespace and counts non-whitespace runs:
		//   p0 accepted: "Origin paragraph: ." = 3 tokens
		//   p1 accepted: "Destination paragraph: the moved sentence." = 5 tokens
		// total = 8
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("--baseline skips the moveTo destination", async () => {
		const result = await runCli("wc", FIXTURE, "--baseline");
		// Baseline view: "Origin paragraph: the moved sentence." + "Destination paragraph: ."
		//   p0: "Origin", "paragraph:", "the", "moved", "sentence." = 5 tokens
		//   p1: "Destination", "paragraph:", "." = 3 tokens
		// total = 8
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("default is the accepted view (skips moveFrom origin)", async () => {
		const result = await runCli("wc", FIXTURE);
		// Same as --accepted: 8 words (default flipped from "current"
		// for consistency with `read --markdown` / `find` / `replace`).
		expect((result.parsed as { words: number }).words).toBe(8);
	});

	test("--current counts both halves (legacy default)", async () => {
		const result = await runCli("wc", FIXTURE, "--current");
		// Both halves visible: 5 + 5 = 10
		expect((result.parsed as { words: number }).words).toBe(10);
	});
});

describe("tracked moves — accept", () => {
	test("accept --all unwraps moveTo (text stays) and deletes moveFrom (text gone)", async () => {
		const docPath = await freshCopy("moves-accept-all");
		const result = await runCli("track-changes", "accept", docPath, "--all");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const doc = read.parsed as Doc;
		expect(paragraphText(doc, 0)).toBe("Origin paragraph: .");
		expect(paragraphText(doc, 1)).toBe(
			"Destination paragraph: the moved sentence.",
		);

		// And the underlying XML has no remaining move wrappers.
		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		expect(xml).not.toContain("<w:moveFrom");
		expect(xml).not.toContain("<w:moveTo");
	});

	test("accept --at one moveFrom alone leaves the moveTo intact", async () => {
		const docPath = await freshCopy("moves-accept-one");
		// tc0 is the moveFrom (it appears first in document order); tc1 is the moveTo.
		const result = await runCli(
			"track-changes",
			"accept",
			docPath,
			"--at",
			"tc0",
		);
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const doc = read.parsed as Doc;
		// p0: moveFrom accepted → text gone.
		expect(paragraphText(doc, 0)).toBe("Origin paragraph: .");
		// p1: moveTo still wrapped.
		const movedRunTo = (doc.blocks[1]?.runs ?? []).find(
			(run) => run.type === "text" && run.text === "the moved sentence",
		);
		expect(movedRunTo?.trackedChange?.kind).toBe("moveTo");
	});
});

describe("tracked moves — reject", () => {
	test("reject --all unwraps moveFrom (text stays) and deletes moveTo (text gone)", async () => {
		const docPath = await freshCopy("moves-reject-all");
		const result = await runCli("track-changes", "reject", docPath, "--all");
		expect(result.exitCode).toBe(0);

		const read = await runCli("read", docPath);
		const doc = read.parsed as Doc;
		expect(paragraphText(doc, 0)).toBe("Origin paragraph: the moved sentence.");
		expect(paragraphText(doc, 1)).toBe("Destination paragraph: .");

		const pkg = await Pkg.open(docPath);
		const xml = await pkg.readText("word/document.xml");
		expect(xml).not.toContain("<w:moveFrom");
		expect(xml).not.toContain("<w:moveTo");
	});
});
