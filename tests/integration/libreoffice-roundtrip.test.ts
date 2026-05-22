import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { runCli } from "../cli/harness";

const SOFFICE = await detectSoffice();

// Each soffice convert is ~1s, so by default we round-trip only a curated CORE
// subset — one fixture per distinct surface we emit or must preserve — rather
// than every fixture (many overlap: 3 comments fixtures, 4 tracked fixtures,
// several Word-authored docs). Set DOCX_LO_ALL=1 for the full sweep; CI does
// that so coverage isn't lost there.
const CORE_FIXTURES = [
	"minimal.docx", // canonical parts baseline (what `create` emits)
	"styles-injection.docx", // styles.xml provisioned from scratch
	"lists.docx", // numbering.xml
	"notes.docx", // footnotes / endnotes write-back
	"multi-tracked.docx", // run-level <w:ins>/<w:del>/move emit
	"multi-column.docx", // <w:sectPr> + columns
	"comments-with-replies.docx", // comments + commentsExtended
	"tables-mutations.docx", // tables + merges + tracked-table revisions (richest)
] as const;

const EXTRA_FIXTURES = [
	"comments-simple.docx",
	"comments-batch.docx",
	"academic-paper.docx",
	"large-mixed.docx",
	"resume-styling.docx",
	"tracked-changes.docx",
	"tracked-moves.docx",
	"chained-tracked-edits.docx",
	"transparent-wrappers.docx",
	"sections.docx",
	"normalize-query.docx",
	"word-formatted.docx",
	"tables.docx",
] as const;

const FIXTURES = Bun.env.DOCX_LO_ALL
	? [...CORE_FIXTURES, ...EXTRA_FIXTURES]
	: CORE_FIXTURES;

let workspace: string;
let sofficeProfile: string;

beforeAll(() => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-lo-"));
	// Isolated soffice user-profile dir, shared across this file's tests (which
	// run serially) — prevents contention with the system-default profile or
	// with concurrent `bun test` invocations sharing one machine.
	sofficeProfile = mkdtempSync(join(tmpdir(), "docx-cli-soffice-"));
});

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
	if (sofficeProfile) rmSync(sofficeProfile, { recursive: true, force: true });
});

describe("LibreOffice round-trip", () => {
	if (!SOFFICE) {
		test.skip("soffice not on PATH — skipping integration tests", () => {});
		return;
	}

	for (const fixture of FIXTURES) {
		test(`${fixture} re-saves cleanly through the create→insert→read pipeline`, async () => {
			const docPath = join(workspace, fixture);
			await Bun.write(docPath, Bun.file(`tests/fixtures/${fixture}`));

			// Mutate: insert a small paragraph at the end of the body to force
			// a write through our serializer.
			const read = await runCli("read", docPath, "--ast");
			const doc = read.parsed as {
				blocks: Array<{ id: string; type: string }>;
			};
			const lastParagraph = [...doc.blocks]
				.reverse()
				.find((block) => block.type === "paragraph");
			expect(lastParagraph).toBeDefined();

			await runCli(
				"insert",
				docPath,
				"--after",
				lastParagraph?.id ?? "p0",
				"--text",
				"docx-cli round-trip probe",
			);

			const outDir = join(workspace, `lo-${fixture}`);
			mkdirSync(outDir, { recursive: true });
			const { exitCode, stdout, stderr } = await runSoffice(docPath, outDir);

			expect(exitCode).toBe(0);
			// LibreOffice prints some macOS-only "Task policy" noise on stderr; ignore it.
			const meaningfulStderr = stderr
				.split("\n")
				.filter(
					(line) =>
						!line.includes("Task policy") &&
						!line.includes("Warning") &&
						line.trim().length > 0,
				)
				.join("\n");
			expect(meaningfulStderr).toBe("");
			expect(stdout).toContain("convert");
			const converted = Bun.file(join(outDir, fixture));
			expect(await converted.exists()).toBe(true);
			expect((await converted.arrayBuffer()).byteLength).toBeGreaterThan(0);
		}, 30_000);
	}

	test("tracked-on edits round-trip through LibreOffice with markers preserved", async () => {
		const docPath = join(workspace, "tracked-roundtrip.docx");
		await runCli("create", docPath, "--text", "Original body text here.");
		await runCli("track-changes", docPath, "on");

		await runCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"Inserted with tracking",
		);
		await runCli("replace", docPath, "Original", "Modified");

		const beforeRead = await runCli("read", docPath, "--ast");
		const beforeRuns = (
			beforeRead.parsed as {
				blocks: Array<{
					runs?: Array<{ trackedChange?: { kind: string } }>;
				}>;
			}
		).blocks.flatMap((b) => b.runs ?? []);
		const insBefore = beforeRuns.filter(
			(r) => r.trackedChange?.kind === "ins",
		).length;
		const delBefore = beforeRuns.filter(
			(r) => r.trackedChange?.kind === "del",
		).length;
		expect(insBefore).toBeGreaterThan(0);
		expect(delBefore).toBeGreaterThan(0);

		const outDir = join(workspace, "lo-tracked");
		mkdirSync(outDir, { recursive: true });
		const { exitCode } = await runSoffice(docPath, outDir);
		expect(exitCode).toBe(0);

		const converted = join(outDir, "tracked-roundtrip.docx");
		const afterRead = await runCli("read", converted, "--ast");
		const afterRuns = (
			afterRead.parsed as {
				blocks: Array<{
					runs?: Array<{ trackedChange?: { kind: string } }>;
				}>;
			}
		).blocks.flatMap((b) => b.runs ?? []);
		const insAfter = afterRuns.filter(
			(r) => r.trackedChange?.kind === "ins",
		).length;
		const delAfter = afterRuns.filter(
			(r) => r.trackedChange?.kind === "del",
		).length;
		expect(insAfter).toBeGreaterThan(0);
		expect(delAfter).toBeGreaterThan(0);
	}, 30_000);
});

async function detectSoffice(): Promise<string | null> {
	const candidates = [
		"soffice",
		"/Applications/LibreOffice.app/Contents/MacOS/soffice",
	];
	for (const candidate of candidates) {
		const proc = Bun.spawn([candidate, "--version"], {
			stdout: "ignore",
			stderr: "ignore",
		});
		const exitCode = await proc.exited;
		if (exitCode === 0) return candidate;
	}
	return null;
}

/** Spawn `soffice --headless --convert-to docx` against the per-file isolated
 * user-profile directory. Without -env:UserInstallation, every soffice process
 * competes for a lock on the default profile (~/Library/Application Support/
 * LibreOffice/4 on macOS). A stale or concurrent soffice causes the new one
 * to exit non-zero — a real flake source. Tests within this file run serially,
 * so they safely share the same isolated profile. */
async function runSoffice(
	docPath: string,
	outDir: string,
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
	if (!SOFFICE) throw new Error("soffice not on PATH");
	const proc = Bun.spawn(
		[
			SOFFICE,
			`-env:UserInstallation=file://${sofficeProfile}`,
			"--headless",
			"--convert-to",
			"docx",
			docPath,
			"--outdir",
			outDir,
		],
		{ stdout: "pipe", stderr: "pipe" },
	);
	const stdout = await new Response(proc.stdout).text();
	const stderr = await new Response(proc.stderr).text();
	const exitCode = await proc.exited;
	return { exitCode, stdout, stderr };
}
