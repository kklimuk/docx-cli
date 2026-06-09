import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "../../src/core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { trackedKinds } from "./helpers";

type FootnoteRefRun = {
	type: string;
	kind?: "footnote" | "endnote";
	id?: string;
	trackedChange?: { kind: string };
};

type Note = { id: string; text: string };

async function freshDoc(label: string, paragraphs: string[]): Promise<string> {
	const docPath = join(tempWorkspace(label), "out.docx");
	const [first, ...rest] = paragraphs;
	await runCli("create", docPath, "--text", first ?? "");
	for (let index = 0; index < rest.length; index++) {
		await runCli(
			"insert",
			docPath,
			"--after",
			`p${index}`,
			"--text",
			rest[index] ?? "",
		);
	}
	return docPath;
}

async function refRuns(docPath: string): Promise<FootnoteRefRun[]> {
	const result = await runCli("read", docPath, "--ast");
	const doc = result.parsed as { blocks: Array<{ runs?: FootnoteRefRun[] }> };
	return doc.blocks
		.flatMap((block) => block.runs ?? [])
		.filter((run) => run.type === "noteRef");
}

async function listFootnotes(docPath: string): Promise<Note[]> {
	const result = await runCli("footnotes", "list", docPath);
	return result.parsed as Note[];
}

async function listEndnotes(docPath: string): Promise<Note[]> {
	const result = await runCli("endnotes", "list", docPath);
	return result.parsed as Note[];
}

describe("docx footnotes add", () => {
	test("provisions footnotes.xml from scratch with reserved boilerplate", async () => {
		const docPath = await freshDoc("add-fresh", ["Body text."]);

		// Source doc has no footnotes part.
		const before = await Pkg.open(docPath);
		expect(before.hasPart("word/footnotes.xml")).toBe(false);

		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Inserted footnote.",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { id: string }).id).toBe("fn1");

		// Part is now created and references the user footnote alongside the two
		// reserved Word entries (id=-1 separator, id=0 continuationSeparator).
		const after = await Pkg.open(docPath);
		expect(after.hasPart("word/footnotes.xml")).toBe(true);
		const footnotesXml = await after.readText("word/footnotes.xml");
		expect(footnotesXml).toContain('w:type="separator"');
		expect(footnotesXml).toContain('w:type="continuationSeparator"');
		expect(footnotesXml).toContain('w:id="1"');
		expect(footnotesXml).toContain("Inserted footnote.");

		// AST reader filters the boilerplate; only the user footnote surfaces.
		const notes = await listFootnotes(docPath);
		expect(notes).toEqual([{ id: "fn1", text: "Inserted footnote." }]);

		// The reference run is in the body.
		const runs = await refRuns(docPath);
		expect(runs).toHaveLength(1);
		expect(runs[0]?.kind).toBe("footnote");
		expect(runs[0]?.id).toBe("fn1");
	});

	test("--at pN appends the reference at the end of the paragraph", async () => {
		const docPath = await freshDoc("add-end", ["Trailing anchor."]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"At end.",
		);
		// The reference must sit AFTER the period — verify via the rendered
		// markdown so we exercise the actual XML position, not just the AST run
		// order.
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("Trailing anchor.[^fn1]");
	});

	test("--at pN:0 inserts the reference at the start of the paragraph", async () => {
		const docPath = await freshDoc("add-start", ["Leading anchor."]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:0",
			"--text",
			"At start.",
		);
		const result = await runCli("read", docPath);
		expect(result.stdout).toContain("[^fn1]Leading anchor.");
	});

	test("--at pN:offset splits a run mid-text and inserts between halves", async () => {
		const docPath = await freshDoc("add-mid", ["Anchored mid-run."]);
		await runCli("footnotes", "add", docPath, "--at", "p0:8", "--text", "Mid.");
		const result = await runCli("read", docPath);
		// Offset 8 = right after "Anchored".
		expect(result.stdout).toContain("Anchored[^fn1] mid-run.");
	});

	test("ids allocate independently of the reserved -1/0 boilerplate", async () => {
		const docPath = await freshDoc("add-ids", ["One.", "Two."]);
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "First.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p1",
			"--text",
			"Second.",
		);
		const notes = await listFootnotes(docPath);
		expect(notes.map((note) => note.id)).toEqual(["fn1", "fn2"]);
	});

	test("offset out of paragraph length is INVALID_LOCATOR", async () => {
		const docPath = await freshDoc("add-bad", ["Short."]);
		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:50",
			"--text",
			"Too far.",
		);
		expect(result.exitCode).toBe(2);
		expect((result.parsed as { code: string }).code).toBe("INVALID_LOCATOR");
	});

	test("missing paragraph is BLOCK_NOT_FOUND", async () => {
		const docPath = await freshDoc("add-missing", ["Only one."]);
		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p9",
			"--text",
			"Anywhere.",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});

	test("dry-run mints an id without writing the file", async () => {
		const docPath = await freshDoc("add-dry", ["Body."]);
		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Skipped.",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await listFootnotes(docPath)).toEqual([]);
		expect(await refRuns(docPath)).toHaveLength(0);
	});
});

describe("docx footnotes edit", () => {
	test("replaces the body text and keeps the reference intact", async () => {
		const docPath = await freshDoc("edit-one", ["Body."]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Original.",
		);
		const result = await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"Updated.",
		);
		expect(result.exitCode).toBe(0);
		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Updated." },
		]);
		// The reference run stays — edit is body-only.
		expect(await refRuns(docPath)).toHaveLength(1);
	});

	test("unknown id is BLOCK_NOT_FOUND", async () => {
		const docPath = await freshDoc("edit-missing", ["Body."]);
		const result = await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn9",
			"--text",
			"Nope.",
		);
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});
});

describe("docx footnotes delete", () => {
	test("removes the body AND every reference run", async () => {
		const docPath = await freshDoc("del-one", ["First.", "Second."]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Note one.",
		);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p1",
			"--text",
			"Note two.",
		);
		expect(await listFootnotes(docPath)).toHaveLength(2);
		expect(await refRuns(docPath)).toHaveLength(2);

		const result = await runCli("footnotes", "delete", docPath, "--at", "fn1");
		expect(result.exitCode).toBe(0);

		const remaining = await listFootnotes(docPath);
		expect(remaining).toEqual([{ id: "fn2", text: "Note two." }]);
		const refs = await refRuns(docPath);
		expect(refs).toHaveLength(1);
		expect(refs[0]?.id).toBe("fn2");
	});

	test("unknown id is BLOCK_NOT_FOUND", async () => {
		const docPath = await freshDoc("del-missing", ["Body."]);
		const result = await runCli("footnotes", "delete", docPath, "--at", "fn9");
		expect(result.exitCode).toBe(3);
		expect((result.parsed as { code: string }).code).toBe("BLOCK_NOT_FOUND");
	});

	test("dry-run leaves the doc untouched", async () => {
		const docPath = await freshDoc("del-dry", ["Body."]);
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Stay.");
		const result = await runCli(
			"footnotes",
			"delete",
			docPath,
			"--at",
			"fn1",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { dryRun: boolean }).dryRun).toBe(true);
		expect(await listFootnotes(docPath)).toHaveLength(1);
		expect(await refRuns(docPath)).toHaveLength(1);
	});
});

describe("docx footnotes list", () => {
	test("returns [] for a doc with no footnotes part", async () => {
		const docPath = await freshDoc("list-empty", ["Body."]);
		expect(await listFootnotes(docPath)).toEqual([]);
	});
});

describe("docx endnotes (parity with footnotes)", () => {
	test("end-to-end add/list/edit/delete via the endnotes verb", async () => {
		const docPath = await freshDoc("end-flow", ["Body."]);

		await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"First endnote.",
		);
		expect(await listEndnotes(docPath)).toEqual([
			{ id: "en1", text: "First endnote." },
		]);
		const refs = await refRuns(docPath);
		expect(refs).toHaveLength(1);
		expect(refs[0]?.kind).toBe("endnote");
		expect(refs[0]?.id).toBe("en1");

		await runCli(
			"endnotes",
			"edit",
			docPath,
			"--at",
			"en1",
			"--text",
			"Updated endnote.",
		);
		expect(await listEndnotes(docPath)).toEqual([
			{ id: "en1", text: "Updated endnote." },
		]);

		await runCli("endnotes", "delete", docPath, "--at", "en1");
		expect(await listEndnotes(docPath)).toEqual([]);
		expect(await refRuns(docPath)).toHaveLength(0);
	});

	test("footnotes and endnotes ids allocate independently", async () => {
		const docPath = await freshDoc("end-indep", ["Body."]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Footnote.",
		);
		await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Endnote.",
		);
		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Footnote." },
		]);
		expect(await listEndnotes(docPath)).toEqual([
			{ id: "en1", text: "Endnote." },
		]);
	});
});

describe("docx footnotes fixture round-trip", () => {
	test("tests/fixtures/footnotes-mutations.docx reads cleanly via the AST", async () => {
		const fixture = join(
			import.meta.dir,
			"..",
			"fixtures",
			"footnotes-mutations.docx",
		);
		const footnotes = await listFootnotes(fixture);
		// fn1–fn3 are untracked. fn4 is added under tracking (still surfaces in
		// `list` — the AST reader ignores tracking wrappers in note bodies;
		// `track-changes accept --all` is what makes it permanent).
		expect(footnotes.map((note) => note.id)).toEqual([
			"fn1",
			"fn2",
			"fn3",
			"fn4",
		]);
		expect(footnotes.find((note) => note.id === "fn1")?.text).toBe(
			"Edited body fn1.",
		);

		const endnotes = await listEndnotes(fixture);
		expect(endnotes.map((note) => note.id)).toEqual(["en1"]);

		const refs = await refRuns(fixture);
		expect(refs.filter((run) => run.kind === "footnote")).toHaveLength(4);
		expect(refs.filter((run) => run.kind === "endnote")).toHaveLength(1);
	});
});

// Each tracked footnote test exercises a full lifecycle: emit the operation
// under tracking, sanity-check the XML shape (matches Word's empirical output
// per `/tmp/fn-probe/*` — see `scripts/word-redlines.sh` for the oracle path),
// then verify accept/reject reaches the expected end state. `--all` exercises
// both the doc-body reference revision AND the paired body-side revision; the
// post-pass GC in `cli/track-changes/apply.ts::applyNotePairing` is what makes
// orphan `<w:footnote>` wrappers disappear on reject-add and accept-delete.
describe("docx footnotes under track-changes", () => {
	async function readPart(docPath: string, part: string): Promise<string> {
		const pkg = await Pkg.open(docPath);
		return pkg.readText(part);
	}

	async function setupTracked(label: string): Promise<string> {
		const docPath = join(tempWorkspace(label), "out.docx");
		await runCli("create", docPath, "--text", "First paragraph.");
		await runCli("insert", docPath, "--after", "p0", "--text", "Second.");
		await runCli("track-changes", docPath, "on");
		return docPath;
	}

	test("add emits paired <w:ins> wrappers on reference and body", async () => {
		const docPath = await setupTracked("tracked-add");
		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Tracked.",
			"--author",
			"Probe",
		);
		expect(result.exitCode).toBe(0);

		const documentXml = await readPart(docPath, "word/document.xml");
		// Reference run is wrapped in <w:ins> with the chosen author.
		expect(documentXml).toContain('<w:ins w:id="0" w:author="Probe"');
		expect(documentXml).toContain('<w:footnoteReference w:id="1"/>');

		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		// Body content (footnoteRef + text run) is wrapped in <w:ins> too. Word
		// uses a distinct revision id for the body side (`/tmp/fn-probe/add.docx`
		// reads id=0 ref, id=1 body) — we match.
		expect(footnotesXml).toContain('<w:ins w:id="1" w:author="Probe"');
		expect(footnotesXml).toContain("<w:footnoteRef/>");
		expect(footnotesXml).toContain("Tracked.");
	});

	test("track-changes list surfaces ONE tcN per tracked footnote add", async () => {
		const docPath = await setupTracked("tracked-add-list");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"One.",
			"--author",
			"Probe",
		);
		const list = (await runCli("track-changes", "list", docPath))
			.parsed as Array<{
			id: string;
			kind: string;
		}>;
		// One tcN — the body-reference revision. The body-side <w:ins> is
		// paired and intentionally hidden from `list`/`apply`.
		expect(list).toHaveLength(1);
		expect(list[0]?.kind).toBe("ins");
	});

	test("accept --all unwraps both sides; reference + body persist", async () => {
		const docPath = await setupTracked("tracked-add-accept");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Stays.",
			"--author",
			"Probe",
		);
		await runCli("track-changes", "accept", docPath, "--all");

		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Stays." },
		]);
		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).not.toContain("<w:ins");
		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		expect(footnotesXml).not.toContain("<w:ins");
	});

	test("reject --all on tracked add GCs the orphan footnote body", async () => {
		const docPath = await setupTracked("tracked-add-reject");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Gone.",
			"--author",
			"Probe",
		);
		await runCli("track-changes", "reject", docPath, "--all");

		// Reference disappeared.
		expect(await refRuns(docPath)).toHaveLength(0);
		// And the body — Word GCs the orphan, we match. footnotes.xml keeps
		// only the two reserved boilerplate entries (separator + continuation).
		expect(await listFootnotes(docPath)).toEqual([]);
		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		// Reserved entries have w:type set; the regex looks for any unreserved
		// <w:footnote> element to confirm true GC.
		expect(footnotesXml).not.toMatch(/<w:footnote w:id="[1-9][0-9]*"/);
	});

	test("delete wraps reference run + body content in <w:del>", async () => {
		const docPath = join(tempWorkspace("tracked-del"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Pre-existing.",
		);
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"delete",
			docPath,
			"--at",
			"fn1",
			"--author",
			"Probe",
		);

		const documentXml = await readPart(docPath, "word/document.xml");
		expect(documentXml).toContain('<w:del w:id="0" w:author="Probe"');
		// Reference run still present, just wrapped — accept will remove,
		// reject will unwrap.
		expect(documentXml).toContain('<w:footnoteReference w:id="1"/>');

		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		// Body content wrapped in <w:del>, with <w:t> renamed to <w:delText>.
		expect(footnotesXml).toContain('<w:del w:id="1" w:author="Probe"');
		expect(footnotesXml).toContain("<w:delText");
		expect(footnotesXml).not.toMatch(/<w:t[^a-zA-Z]/);
		// AND a paragraph-mark deletion marker in <w:pPr><w:rPr>.
		expect(footnotesXml).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:del/);
	});

	test("accept --all on tracked delete removes reference and GCs the body", async () => {
		const docPath = join(tempWorkspace("tracked-del-accept"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Gone.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"delete",
			docPath,
			"--at",
			"fn1",
			"--author",
			"Probe",
		);
		await runCli("track-changes", "accept", docPath, "--all");

		expect(await refRuns(docPath)).toHaveLength(0);
		expect(await listFootnotes(docPath)).toEqual([]);
	});

	test("reject --all on tracked delete restores the footnote intact", async () => {
		const docPath = join(tempWorkspace("tracked-del-reject"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Stays.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"delete",
			docPath,
			"--at",
			"fn1",
			"--author",
			"Probe",
		);
		await runCli("track-changes", "reject", docPath, "--all");

		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Stays." },
		]);
		const refs = await refRuns(docPath);
		expect(refs).toHaveLength(1);
		expect(refs[0]?.trackedChange).toBeUndefined();
	});

	test("edit emits <w:ins>NEW</w:ins><w:del>OLD</w:del> with ins before del", async () => {
		const docPath = join(tempWorkspace("tracked-edit"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Old.");
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"New.",
			"--author",
			"Probe",
		);

		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		expect(footnotesXml).toContain('<w:ins w:id="0" w:author="Probe"');
		expect(footnotesXml).toContain('<w:del w:id="1" w:author="Probe"');
		expect(footnotesXml).toContain("New.");
		expect(footnotesXml).toContain("Old.");
		// Word's order: ins precedes del. Slice the xml between the two
		// wrappers to assert direction.
		const insIndex = footnotesXml.indexOf("<w:ins");
		const delIndex = footnotesXml.indexOf("<w:del");
		expect(insIndex).toBeGreaterThan(-1);
		expect(delIndex).toBeGreaterThan(insIndex);
	});

	test("body-only edit surfaces two tcN entries (ins + del) in list", async () => {
		const docPath = join(tempWorkspace("tracked-edit-list"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "Old.");
		await runCli("track-changes", docPath, "on");
		await runCli("footnotes", "edit", docPath, "--at", "fn1", "--text", "New.");
		const list = (await runCli("track-changes", "list", docPath))
			.parsed as Array<{
			kind: string;
		}>;
		expect(list).toHaveLength(2);
		expect(list.map((entry) => entry.kind).sort()).toEqual(["del", "ins"]);
	});

	test("accept --all on tracked edit keeps the new text", async () => {
		const docPath = join(tempWorkspace("tracked-edit-accept"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Original.",
		);
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"Replaced.",
		);
		await runCli("track-changes", "accept", docPath, "--all");

		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Replaced." },
		]);
	});

	test("reject --all on tracked edit restores the original text", async () => {
		const docPath = join(tempWorkspace("tracked-edit-reject"), "out.docx");
		await runCli("create", docPath, "--text", "Body.");
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Original.",
		);
		await runCli("track-changes", docPath, "on");
		await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"Replaced.",
		);
		await runCli("track-changes", "reject", docPath, "--all");

		expect(await listFootnotes(docPath)).toEqual([
			{ id: "fn1", text: "Original." },
		]);
	});

	test("endnotes go through the same tracked add/accept lifecycle", async () => {
		const docPath = await setupTracked("tracked-endnote");
		await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"Endnote.",
			"--author",
			"Probe",
		);
		const endnotesXml = await readPart(docPath, "word/endnotes.xml");
		expect(endnotesXml).toContain('<w:ins w:id="1" w:author="Probe"');
		expect(endnotesXml).toContain("Endnote.");

		await runCli("track-changes", "accept", docPath, "--all");
		expect(await listEndnotes(docPath)).toEqual([
			{ id: "en1", text: "Endnote." },
		]);
	});

	test("revision ids never collide with body-side ids when tracking", async () => {
		// Regression for the allocator-only-walks-documentTree bug: editing a
		// body that already carries a tracked-added <w:ins w:id="N"> would
		// previously mint a fresh ID starting from 0, overlapping the body-side
		// id. Now `computeMaxRevisionId` walks footnotes/endnotes too.
		const docPath = await setupTracked("tracked-no-collision");
		await runCli("footnotes", "add", docPath, "--at", "p0", "--text", "First.");
		await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"Edited.",
		);
		const footnotesXml = await readPart(docPath, "word/footnotes.xml");
		// Only look at revision wrapper ids — the footnote's own `w:id="1"`
		// lives in a different id namespace and would muddle this check.
		const ids = [...footnotesXml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map(
			(match) => match[1],
		);
		const unique = new Set(ids);
		expect(ids.length).toBeGreaterThan(0);
		expect(unique.size).toBe(ids.length);
	});
});

// The "under track-changes" block above drives tracking via the global toggle
// (`track-changes on`). The per-invocation `--track` flag (forcing one note
// op tracked while the toggle is OFF) is a separate code path — pinned here.
describe("footnotes / endnotes — --track forces tracking with the toggle off", () => {
	test("footnotes add --track records a tracked insertion", async () => {
		const docPath = await freshDoc("fn-track-add", [
			"Revenue grew sharply this year.",
		]);
		const result = await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:7",
			"--text",
			"FY24 close.",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(docPath)).toContain("ins");
	});

	test("footnotes edit --track records a tracked change to the body", async () => {
		const docPath = await freshDoc("fn-track-edit", [
			"Revenue grew sharply this year.",
		]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:7",
			"--text",
			"original",
		);
		const result = await runCli(
			"footnotes",
			"edit",
			docPath,
			"--at",
			"fn1",
			"--text",
			"revised",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect((await trackedKinds(docPath)).length).toBeGreaterThan(0);
	});

	test("footnotes delete --track records a tracked deletion", async () => {
		const docPath = await freshDoc("fn-track-del", [
			"Revenue grew sharply this year.",
		]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:7",
			"--text",
			"doomed",
		);
		const result = await runCli(
			"footnotes",
			"delete",
			docPath,
			"--at",
			"fn1",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(docPath)).toContain("del");
	});

	test("endnotes add --track records a tracked insertion", async () => {
		const docPath = await freshDoc("en-track-add", [
			"Endnote anchor text here now.",
		]);
		const result = await runCli(
			"endnotes",
			"add",
			docPath,
			"--at",
			"p0:8",
			"--text",
			"see ref.",
			"--track",
		);
		expect(result.exitCode).toBe(0);
		expect(await trackedKinds(docPath)).toContain("ins");
	});

	test("no --track on an untracked doc records nothing", async () => {
		const docPath = await freshDoc("fn-track-control", [
			"Revenue grew sharply this year.",
		]);
		await runCli(
			"footnotes",
			"add",
			docPath,
			"--at",
			"p0:7",
			"--text",
			"plain",
		);
		expect(await trackedKinds(docPath)).toHaveLength(0);
	});
});

describe("docx footnotes — -o parallel write", () => {
	test("footnotes add -o writes to the output and leaves the source byte-unchanged", async () => {
		const src = await freshDoc("fn-o-src", ["Revenue grew sharply here."]);
		const before = await Bun.file(src).bytes();
		const out = join(tempWorkspace("fn-o-out"), "out.docx");

		const result = await runCli(
			"footnotes",
			"add",
			src,
			"--at",
			"p0:7",
			"--text",
			"fn",
			"-o",
			out,
		);
		expect(result.exitCode).toBe(0);
		expect((result.parsed as { path: string }).path).toBe(out);
		expect(await Bun.file(src).bytes()).toEqual(before);

		expect((await listFootnotes(out)).length).toBe(1);
		expect(await listFootnotes(src)).toEqual([]);
	});
});
