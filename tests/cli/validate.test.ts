import { beforeEach, describe, expect, test } from "bun:test";
import { join } from "node:path";
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
