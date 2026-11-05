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

describe("cross-part consistency (issue #34)", () => {
	// Schema-valid packages Word still repairs on open: the per-part XSD check
	// can't see a comment missing from its paraId-keyed side-part, or comment
	// markers that don't pair up. `validate` reports them beside schema errors.
	let commented: string;

	beforeEach(async () => {
		commented = join(workspace, "commented.docx");
		await runCli("create", commented, "--text", "hello world");
		await runCli("comments", "add", commented, "--at", "p0", "--text", "n");
	});

	async function patchPart(
		name: string,
		patch: (xml: string) => string,
	): Promise<void> {
		const pkg = await Pkg.open(commented);
		pkg.writeText(name, patch(await pkg.readText(name)));
		await pkg.save();
	}

	test("a consistent commented document validates clean", async () => {
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^valid/);
	});

	test("a comment missing from a present commentsIds.xml is an error", async () => {
		const pkg = await Pkg.open(commented);
		pkg.writeText(
			"word/commentsIds.xml",
			'<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>',
		);
		await pkg.save();
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain("word/commentsIds.xml: 1 error");
		expect(result.stdout).toMatch(
			/comment c0 \(paraId [0-9A-F]{8}\) has no entry in word\/commentsIds\.xml \(consistency\)/,
		);
	});

	test("a present commentsIds.xml that keys the comment passes", async () => {
		const comments = await (await Pkg.open(commented)).readText(
			"word/comments.xml",
		);
		const paraId = comments.match(/w14:paraId="([0-9A-F]{8})"/)?.[1];
		expect(paraId).toBeDefined();
		const pkg = await Pkg.open(commented);
		pkg.writeText(
			"word/commentsIds.xml",
			`<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="1A2B3C4D"/></w16cid:commentsIds>`,
		);
		await pkg.save();
		expect((await runCli("validate", commented)).exitCode).toBe(0);
	});

	test("a comment range whose reference run was deleted is an error", async () => {
		await patchPart("word/document.xml", (xml) =>
			xml.replace(
				/<w:r><w:rPr><w:rStyle w:val="CommentReference"\/><\/w:rPr><w:commentReference w:id="0"\/><\/w:r>/,
				"",
			),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain(
			"comment range 0 has no matching commentReference (consistency)",
		);
	});

	test("an unpaired commentRangeStart is an error", async () => {
		await patchPart("word/document.xml", (xml) =>
			xml.replace('<w:commentRangeEnd w:id="0"/>', ""),
		);
		const result = await runCli("validate", commented);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toContain(
			"commentRangeStart 0 has no matching commentRangeEnd",
		);
	});

	test("a commentReference to a missing comment is an error, filed under its story part", async () => {
		await patchPart("word/comments.xml", (xml) =>
			xml.replace(/<w:comment [\s\S]*<\/w:comment>/, ""),
		);
		const result = await runCli("validate", commented, "--json");
		expect(result.exitCode).toBe(1);
		const parsed = result.parsed as {
			valid: boolean;
			parts: { part: string; issues: { message: string; kind?: string }[] }[];
		};
		expect(parsed.valid).toBe(false);
		const documentPart = parsed.parts.find(
			(part) => part.part === "word/document.xml",
		);
		expect(documentPart?.issues).toContainEqual({
			message: "commentReference 0 points at no comment in word/comments.xml",
			kind: "consistency",
		});
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
