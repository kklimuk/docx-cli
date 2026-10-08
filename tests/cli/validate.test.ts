import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";
import { freshFixture } from "./helpers";

let workspace: string;
let docPath: string;

beforeEach(async () => {
	workspace = tempWorkspace("validate");
	docPath = join(workspace, "doc.docx");
	await runCli("create", docPath, "--text", "clean document");
});

describe("docx validate", () => {
	test("a clean document validates with exit 0", async () => {
		const result = await runCli("validate", docPath);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("valid");
	});

	test("a Word-authored document (MCE-laden) validates clean", async () => {
		const wordDoc = await freshFixture(
			"validate-word",
			"tests/fixtures/academic-paper.docx",
		);
		const result = await runCli("validate", wordDoc);
		expect(result.exitCode).toBe(0);
	});

	test("schema errors list per part and exit 1", async () => {
		await runCli(
			"raw",
			"insert",
			docPath,
			"--after",
			"p0",
			"--no-validate",
			"--xml",
			"<w:p><w:r><w:bogusChild/><w:t>x</w:t></w:r></w:p>",
		);
		const result = await runCli("validate", docPath);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain("word/document.xml: 1 error");
		expect(result.stdout).toContain("bogusChild");
	});

	test("--json reports valid/errors/parts", async () => {
		const result = await runCli("validate", docPath, "--json");
		const parsed = result.parsed as {
			valid: boolean;
			errors: number;
			parts: { part: string; issues: unknown[] }[];
		};
		expect(parsed.valid).toBe(true);
		expect(parsed.errors).toBe(0);
		expect(parsed.parts.some((part) => part.part === "word/document.xml")).toBe(
			true,
		);
	});

	test("an ISO-strict document is skipped with a note, not flooded with noise", async () => {
		const strictDoc = await freshFixture(
			"validate-strict",
			"tests/fixtures/strict-profile.docx",
		);
		const result = await runCli("validate", strictDoc);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("strict-profile");
	});
});

// The emitter-conformance pin: every fixture the repo ships must be
// schema-valid OOXML. This is what keeps the bug classes the validator audit
// found (settings trackRevisions ordering, tblGridChange attributes, comments
// mc:Ignorable/textId, nested w:rPr in math runs, stale hand-built XML) from
// regressing — Word TOLERATES most of them, so renders and round-trips can't.
// strict-profile.docx is the deliberate exception (ISO-strict namespaces; the
// skip test above covers it).
describe("every fixture validates clean", () => {
	const glob = new Bun.Glob("*.docx");
	const fixtures = [...glob.scanSync("tests/fixtures")]
		.filter((name) => name !== "strict-profile.docx")
		.sort();

	test.each(fixtures)("%s", async (name) => {
		const result = await runCli("validate", `tests/fixtures/${name}`);
		expect(result.stdout.trim()).toMatch(/^valid/);
		expect(result.exitCode).toBe(0);
	});
});

describe("legacy schema noise self-heals on any save", () => {
	test("a misnamed <w:trackChanges> and an undeclared w14 comments part are repaired by an unrelated edit", async () => {
		// Documents produced by earlier docx-cli versions (and some Word files)
		// carry two defects our views can fix: the schema-invalid legacy
		// `<w:trackChanges/>` toggle, and comment paragraphs with `w14:paraId`
		// while the root omits `mc:Ignorable`. An agent that runs `validate` after
		// its own edit read those as "I broke the file." Every save now heals both,
		// so ANY mutation leaves the document clean.
		const legacy = join(workspace, "legacy.docx");
		await runCli("create", legacy, "--text", "Clause one. Clause two.");
		await runCli("track-changes", "on", legacy);
		await runCli("comments", "add", legacy, "--at", "p0", "--text", "note");
		const pkg = await Pkg.open(legacy);
		pkg.writeText(
			"word/settings.xml",
			(await pkg.readText("word/settings.xml"))
				.replace(/<w:trackRevisions\/>/, "")
				.replace(/(<w:settings[^>]*>)/, "$1<w:trackChanges/>"),
		);
		pkg.writeText(
			"word/comments.xml",
			(await pkg.readText("word/comments.xml")).replace(
				/ mc:Ignorable="[^"]*"/,
				"",
			),
		);
		await pkg.save();
		const before = await runCli("validate", legacy);
		expect(before.exitCode).toBe(1);
		expect(before.stdout).toContain("word/settings.xml");
		expect(before.stdout).toContain("word/comments.xml");
		// Tracking still reads as ON through the legacy element.
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);

		// An unrelated mutation — no toggle, no comment — heals both parts.
		await runCli(
			"edit",
			legacy,
			"--at",
			"p0",
			"--text",
			"Clause one, amended.",
		);
		const after = await runCli("validate", legacy);
		expect(after.exitCode).toBe(0);
		expect(after.stdout).toMatch(/^valid/);
		const healed = await Pkg.open(legacy);
		const settings = await healed.readText("word/settings.xml");
		expect(settings).toContain("<w:trackRevisions/>");
		expect(settings).not.toContain("w:trackChanges");
		expect(await healed.readText("word/comments.xml")).toMatch(
			/<w:comments[^>]*mc:Ignorable="[^"]*w14/,
		);
		// The tracked edit landed as a real revision (tracking stayed on).
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);
		const revisions = (await runCli("track-changes", "list", legacy, "--json"))
			.parsed as unknown[];
		expect(revisions.length).toBeGreaterThan(0);
	});

	test("a legacy ON toggle beside a present-but-off <w:trackRevisions> stays ON after an unrelated save", async () => {
		// An earlier docx-cli's `track-changes on` added the misnamed element next
		// to a producer's `w:val="false"` real one. `read` reports ON (any ON
		// toggle wins), so the save-time migration must keep it ON — not let the
		// stale OFF element win and silently flip tracking off.
		const legacy = join(workspace, "conflict.docx");
		await runCli("create", legacy, "--text", "Clause one.");
		const pkg = await Pkg.open(legacy);
		pkg.writeText(
			"word/settings.xml",
			(await pkg.readText("word/settings.xml"))
				.replace(/(<w:settings[^>]*>)/, "$1<w:trackChanges/>")
				// The real element at its CT_Settings slot (before defaultTabStop).
				.replace(/(<w:defaultTabStop)/, '<w:trackRevisions w:val="false"/>$1'),
		);
		await pkg.save();
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);

		await runCli(
			"edit",
			legacy,
			"--at",
			"p0",
			"--text",
			"Clause one, amended.",
		);
		expect((await runCli("read", legacy)).stdout).toContain(
			"docx:track-changes on",
		);
		const settings = await (await Pkg.open(legacy)).readText(
			"word/settings.xml",
		);
		expect(settings).toContain("<w:trackRevisions/>");
		expect(settings).not.toContain("w:trackChanges");
		expect((await runCli("validate", legacy)).exitCode).toBe(0);
	});
});

describe("justified alignment emits ST_Jc `both`", () => {
	test("every paragraph emitter writes a schema-valid justify that reads back", async () => {
		// ST_Jc has no `justify` — Word's spelling is `both`. The image and
		// equation paragraphs build their own <w:pPr>, so they must map it too.
		await runCli("edit", docPath, "--at", "p0", "--alignment", "justify");
		await runCli(
			"images",
			"add",
			docPath,
			"--after",
			"p0",
			"--image",
			join(import.meta.dir, "..", "fixtures", "assets", "sample.png"),
			"--alignment",
			"justify",
		);
		await runCli(
			"equations",
			"add",
			docPath,
			"--at-end",
			"--equation",
			"x^2",
			"--alignment",
			"justify",
		);
		const result = await runCli("validate", docPath);
		expect(result.stdout).toMatch(/^valid/);
		expect(result.exitCode).toBe(0);
		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown.match(/align="justify"/g)).toHaveLength(3);
	});
});
