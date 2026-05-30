import { beforeEach, describe, expect, test } from "bun:test";
import { copyFileSync } from "node:fs";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

const SECTIONS_FIXTURE = join(
	import.meta.dir,
	"..",
	"fixtures",
	"sections.docx",
);

type SectionBlock = {
	id: string;
	type: "sectionBreak";
	columns?: number;
	sectionType?: string;
};

type ParagraphBlock = {
	id: string;
	type: "paragraph";
	runs?: Array<{ type: string; text?: string }>;
};

type Block = SectionBlock | ParagraphBlock | { id: string; type: string };

describe("sections via insert / edit / delete", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("sections");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "First");
	});

	test("insert --section emits a sentinel paragraph + new sN block", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
			"--type",
			"continuous",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const sections = doc.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections).toHaveLength(2);
		expect(sections[0]).toMatchObject({
			id: "s0",
			columns: 2,
			sectionType: "continuous",
		});
		expect(sections[1]).toMatchObject({ id: "s1" });
		expect(sections[1]?.columns).toBeUndefined();
	});

	test("edit --at sN updates columns + type on the trailing section", async () => {
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		const read = await runCli("read", docPath, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const section = doc.blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(section).toMatchObject({
			id: "s0",
			columns: 3,
			sectionType: "continuous",
		});
	});

	test("edit --at sN supports updating only one property at a time", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "2");
		const firstRead = await runCli("read", docPath, "--ast");
		const after1 = (firstRead.parsed as { blocks: Block[] }).blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(after1?.columns).toBe(2);
		expect(after1?.sectionType).toBeUndefined();

		await runCli("edit", docPath, "--at", "s0", "--type", "nextPage");
		const secondRead = await runCli("read", docPath, "--ast");
		const after2 = (secondRead.parsed as { blocks: Block[] }).blocks.find(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(after2?.columns).toBe(2);
		expect(after2?.sectionType).toBe("nextPage");
	});

	test("edit --at sN rejects --text/--runs", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--text",
			"nope",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("edit --at pN rejects --columns/--type", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"p0",
			"--text",
			"x",
			"--columns",
			"2",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("edit --at sN rejects invalid --type", async () => {
		const result = await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--type",
			"bogus",
		);
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("delete --at sN removes an inline sectPr; the sentinel paragraph remains", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
		);
		const beforeRead = await runCli("read", docPath, "--ast");
		const before = (beforeRead.parsed as { blocks: Block[] }).blocks;
		const beforeSections = before.filter((b) => b.type === "sectionBreak");
		expect(beforeSections).toHaveLength(2);

		const deleteResult = await runCli("delete", docPath, "--at", "s0");
		expect(deleteResult.exitCode).toBe(0);

		const afterRead = await runCli("read", docPath, "--ast");
		const after = (afterRead.parsed as { blocks: Block[] }).blocks;
		const afterSections = after.filter((b) => b.type === "sectionBreak");
		expect(afterSections).toHaveLength(1);
		// sentinel paragraph survives — only the sectPr was stripped.
		const paragraphs = after.filter((b) => b.type === "paragraph");
		expect(paragraphs).toHaveLength(2);
	});

	test("delete --at sN rejects the trailing section break", async () => {
		const result = await runCli("delete", docPath, "--at", "s0");
		expect(result.exitCode).toBe(2);
		expect(result.parsed).toMatchObject({ ok: false, code: "USAGE" });
	});

	test("wc reports a sections count alongside words", async () => {
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
		);
		const wc = await runCli("wc", docPath);
		expect(wc.parsed).toMatchObject({
			ok: true,
			scope: "document",
			words: 1,
			sections: 2,
		});
	});
});

describe("sections under track-changes", () => {
	let docPath: string;

	beforeEach(async () => {
		const workspace = tempWorkspace("sections-tc");
		docPath = join(workspace, "out.docx");
		await runCli("create", docPath, "--text", "Hello");
	});

	test("insert --section under tracking + reject removes the sentinel paragraph entirely", async () => {
		await runCli("track-changes", docPath, "on");
		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--section",
			"--columns",
			"2",
			"--author",
			"Alice",
		);
		const beforeRead = await runCli("read", docPath, "--ast");
		const beforeBlocks = (beforeRead.parsed as { blocks: Block[] }).blocks;
		expect(beforeBlocks.filter((b) => b.type === "sectionBreak")).toHaveLength(
			2,
		);
		expect(beforeBlocks.filter((b) => b.type === "paragraph")).toHaveLength(2);

		await runCli("track-changes", "reject", docPath, "--all");

		const afterRead = await runCli("read", docPath, "--ast");
		const afterBlocks = (afterRead.parsed as { blocks: Block[] }).blocks;
		// Sentinel paragraph + its inline sectPr both gone; only the original
		// content paragraph and the trailing sectPr remain.
		expect(afterBlocks.filter((b) => b.type === "sectionBreak")).toHaveLength(
			1,
		);
		expect(afterBlocks.filter((b) => b.type === "paragraph")).toHaveLength(1);
	});

	test("delete --at pN under tracking + accept merges the deleted paragraph with the next", async () => {
		await runCli("insert", docPath, "--after", "p0", "--text", "Second");
		await runCli("track-changes", docPath, "on");
		await runCli("delete", docPath, "--at", "p0", "--author", "Alice");

		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{ id: string; kind: string }>;
		// Two trackings on p0: run-level <w:del> wrapping "Original body",
		// and paragraph-mark <w:del> on the paragraph mark itself.
		expect(changes).toHaveLength(2);
		expect(changes.map((c) => c.kind)).toEqual(["del", "del"]);

		await runCli("track-changes", "accept", docPath, "--all");

		const afterRead = await runCli("read", docPath, "--ast");
		const paragraphs = (afterRead.parsed as { blocks: Block[] }).blocks.filter(
			(b): b is ParagraphBlock => b.type === "paragraph",
		);
		// p0 is gone (its runs deleted, its paragraph mark merged with p1's
		// content). The single remaining paragraph carries "Second".
		expect(paragraphs).toHaveLength(1);
		const text = (paragraphs[0]?.runs ?? [])
			.map((run) => run.text ?? "")
			.join("");
		expect(text).toBe("Second");
	});

	test("edit --at sN under tracking emits a sectPrChange snapshot", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
			"--author",
			"Alice",
		);
		const list = await runCli("track-changes", "list", docPath);
		const changes = list.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			blockId: string;
			prior?: { columns?: number; sectionType?: string };
			current?: { columns?: number; sectionType?: string };
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "sectPrChange",
			author: "Alice",
			blockId: "s0",
			prior: { columns: 1 },
			current: { columns: 3, sectionType: "continuous" },
		});
	});

	test("track-changes accept removes the sectPrChange but keeps the new properties", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		await runCli("track-changes", "accept", docPath, "--all");
		const read = await runCli("read", docPath, "--ast");
		const section = (read.parsed as { blocks: Block[] }).blocks.find(
			(b): b is SectionBlock => b.type === "sectionBreak",
		);
		expect(section).toMatchObject({
			id: "s0",
			columns: 3,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("track-changes reject restores prior section properties from snapshot", async () => {
		await runCli("edit", docPath, "--at", "s0", "--columns", "1");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"3",
			"--type",
			"continuous",
		);
		await runCli("track-changes", "reject", docPath, "--all");
		const read = await runCli("read", docPath, "--ast");
		const section = (read.parsed as { blocks: Block[] }).blocks.find(
			(b): b is SectionBlock => b.type === "sectionBreak",
		);
		expect(section).toMatchObject({ id: "s0", columns: 1 });
		expect(section?.sectionType).toBeUndefined();
	});

	// Per ECMA-376 §17.6.18, <w:sectPrChange> must be the LAST child of
	// <w:sectPr>. When tracking is on and we add a NEW property to a sectPr
	// that started empty, applyColumns / applySectionType must insert ahead
	// of the freshly-pushed sectPrChange, not after it.
	test("edit --at sN under tracking keeps sectPrChange as the last child", async () => {
		await runCli("track-changes", docPath, "on");
		await runCli(
			"edit",
			docPath,
			"--at",
			"s0",
			"--columns",
			"2",
			"--type",
			"continuous",
		);
		const pkg = await Pkg.open(docPath);
		const documentXml = await pkg.readText("word/document.xml");
		const sectPrMatch = documentXml.match(
			/<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/,
		);
		expect(sectPrMatch).not.toBeNull();
		const inner = sectPrMatch?.[1] ?? "";
		const childTags = Array.from(inner.matchAll(/<w:(\w+)\b/g)).map(
			(match) => match[1],
		);
		// Only top-level children of the live sectPr — exclude tags nested inside
		// the sectPrChange snapshot (those appear AFTER the sectPrChange opener).
		const changeIndex = childTags.indexOf("sectPrChange");
		expect(changeIndex).toBeGreaterThanOrEqual(0);
		const liveChildren = childTags.slice(0, changeIndex + 1);
		// The invariant per ECMA-376 §17.6.18: sectPrChange must be the LAST
		// child of the live sectPr. Don't assert on the full list of siblings —
		// the create template legitimately seeds pgSz/pgMar/docGrid which would
		// make a strict equality brittle.
		expect(liveChildren[liveChildren.length - 1]).toBe("sectPrChange");
		expect(liveChildren).toContain("type");
		expect(liveChildren).toContain("cols");
	});
});

describe("sections.docx fixture", () => {
	test("AST surfaces every section type with the expected properties", async () => {
		const read = await runCli("read", SECTIONS_FIXTURE, "--ast");
		const doc = read.parsed as { blocks: Block[] };
		const sections = doc.blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections).toEqual([
			{ id: "s0", type: "sectionBreak", columns: 1, sectionType: "continuous" },
			{ id: "s1", type: "sectionBreak", columns: 2, sectionType: "continuous" },
			{ id: "s2", type: "sectionBreak", columns: 2, sectionType: "nextColumn" },
			{ id: "s3", type: "sectionBreak", columns: 2, sectionType: "continuous" },
			{ id: "s4", type: "sectionBreak", columns: 1, sectionType: "nextPage" },
			{ id: "s5", type: "sectionBreak", columns: 1, sectionType: "evenPage" },
			{ id: "s6", type: "sectionBreak", columns: 1, sectionType: "oddPage" },
			{ id: "s7", type: "sectionBreak", columns: 2, sectionType: "continuous" },
		]);
	});

	test("track-changes list surfaces the sectPrChange with prior+current diff", async () => {
		const list = await runCli("track-changes", "list", SECTIONS_FIXTURE);
		const changes = list.parsed as Array<{
			id: string;
			kind: string;
			author: string;
			date: string;
			blockId: string;
			prior?: { columns?: number; sectionType?: string };
			current?: { columns?: number; sectionType?: string };
		}>;
		expect(changes).toHaveLength(1);
		expect(changes[0]).toMatchObject({
			id: "tc0",
			kind: "sectPrChange",
			author: "Reviewer",
			date: "2026-05-06T10:00:00Z",
			blockId: "s7",
			prior: { columns: 1, sectionType: "continuous" },
			current: { columns: 2, sectionType: "continuous" },
		});
	});

	test("wc reports words and the section count from the fixture", async () => {
		const wc = await runCli("wc", SECTIONS_FIXTURE);
		expect(wc.parsed).toMatchObject({
			ok: true,
			scope: "document",
			sections: 8,
		});
		const words = (wc.parsed as { words: number }).words;
		expect(words).toBeGreaterThan(800);
	});

	test("track-changes accept on the fixture removes the snapshot, keeps current props", async () => {
		const workspace = tempWorkspace("sections-fixture-accept");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		await runCli("track-changes", "accept", docPath, "--all");

		const read = await runCli("read", docPath, "--ast");
		const sections = (read.parsed as { blocks: Block[] }).blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections.at(-1)).toMatchObject({
			id: "s7",
			columns: 2,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("track-changes reject on the fixture restores the prior 1-column state", async () => {
		const workspace = tempWorkspace("sections-fixture-reject");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		await runCli("track-changes", "reject", docPath, "--all");

		const read = await runCli("read", docPath, "--ast");
		const sections = (read.parsed as { blocks: Block[] }).blocks.filter(
			(block): block is SectionBlock => block.type === "sectionBreak",
		);
		expect(sections.at(-1)).toMatchObject({
			id: "s7",
			columns: 1,
			sectionType: "continuous",
		});
		const list = await runCli("track-changes", "list", docPath);
		expect(list.parsed).toEqual([]);
	});

	test("wc sN counts words in the section's content range", async () => {
		const wc = await runCli("wc", SECTIONS_FIXTURE, "s4");
		expect(wc.parsed).toMatchObject({
			ok: true,
			scope: "section",
			locator: "s4",
		});
		const words = (wc.parsed as { words: number }).words;
		expect(words).toBeGreaterThan(0);
	});

	test("per-section wc sums to the doc-level total", async () => {
		const docResult = await runCli("wc", SECTIONS_FIXTURE);
		const docWords = (docResult.parsed as { words: number; sections: number })
			.words;
		const sectionCount = (
			docResult.parsed as { words: number; sections: number }
		).sections;

		let perSectionSum = 0;
		for (let index = 0; index < sectionCount; index += 1) {
			const result = await runCli("wc", SECTIONS_FIXTURE, `s${index}`);
			perSectionSum += (result.parsed as { words: number }).words;
		}
		expect(perSectionSum).toBe(docWords);
	});

	test("wc returns BLOCK_NOT_FOUND for an out-of-range section locator", async () => {
		const result = await runCli("wc", SECTIONS_FIXTURE, "s99");
		expect(result.exitCode).toBe(3);
		expect(result.parsed).toMatchObject({
			ok: false,
			code: "BLOCK_NOT_FOUND",
		});
	});

	test("each sN locator resolves to a sectPr — no overlap with paragraph/table ids", async () => {
		const workspace = tempWorkspace("sections-fixture-locators");
		const docPath = join(workspace, "sections.docx");
		copyFileSync(SECTIONS_FIXTURE, docPath);

		// edit --at sN succeeds for every enumerated section; if any sN
		// resolved to a paragraph instead, this would error with USAGE.
		for (let index = 0; index < 8; index += 1) {
			const result = await runCli(
				"edit",
				docPath,
				"--at",
				`s${index}`,
				"--columns",
				"1",
			);
			expect(result.exitCode).toBe(0);
		}
	});
});
