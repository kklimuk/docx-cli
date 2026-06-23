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
	"comments-with-replies.docx", // comments + commentsExtended (Word-authored)
	"comments-threaded.docx", // CLI-authored thread: anchored reply markers + root-attached nesting
	"tables-mutations.docx", // tables + merges + tracked-table revisions (richest)
	"images.docx", // inserted media parts + inline <w:drawing> picture
	"footnotes-mutations.docx", // footnotes/endnotes parts authored from scratch
	"task-lists.docx", // SDT <w14:checkbox> inside list paragraphs (GFM tasks)
	"task-lists-web.docx", // Word-for-Web shape: Wingdings ☐ bullet + strike-for-done
	"task-lists-tracked.docx", // tracked checkbox toggles inside SDT content controls
	"equations.docx", // OMML coverage: <m:f> <m:rad> <m:nary> <m:lim> <m:acc> <m:d> <m:m> <m:eqArr> <m:bar> <m:box> + 70+ formulas (statistics, engineering, chemistry)
	"markdown-import.docx", // S8 markdown walker end-to-end (`create --from`) + CodeBlock/Code rStyle/lowlight colors (incl. 6E7781 comment token, formerly code-blocks.docx)
	"word-formatted.docx", // run-level rPr (color/themeColor, shd, u@color, vertAlign, smallCaps, caps; p4) + bold/italic edit-preservation layout (p0-p3); absorbed run-formatting.docx
	"letter.docx", // paragraph pPr: <w:spacing> (before/after/line) + <w:ind> (left/right/firstLine/hanging), authored via insert/edit
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
] as const;

const FIXTURES = Bun.env.DOCX_LO_ALL
	? [...CORE_FIXTURES, ...EXTRA_FIXTURES]
	: CORE_FIXTURES;

type RoundTrip = {
	exitCode: number;
	stdout: string;
	stderr: string;
	convertedBytes: number;
};

// soffice converts (~1s each) are the slow part and are independent, so we run
// them through a bounded worker pool. Each worker needs its OWN user-profile —
// the default profile is a global lock (see CLAUDE.md), so concurrent spawns
// sharing one would contend.
const CONCURRENCY = 4;
let workspace: string;
const profiles: string[] = [];
const results = new Map<string, RoundTrip>();

beforeAll(async () => {
	workspace = mkdtempSync(join(tmpdir(), "docx-cli-lo-"));
	if (!SOFFICE) return;

	// 1. Prep every fixture serially: the in-process runCli shares a global
	//    output sink so it isn't concurrency-safe — but it's fast (no spawn).
	const jobs: Array<{ fixture: string; docPath: string; outDir: string }> = [];
	for (const fixture of FIXTURES) {
		const docPath = join(workspace, fixture);
		await Bun.write(docPath, Bun.file(`tests/fixtures/${fixture}`));
		// Force a write through our serializer by inserting a probe paragraph.
		const read = await runCli("read", docPath, "--ast");
		const body = read.parsed as { blocks: Array<{ id: string; type: string }> };
		const lastParagraph = [...body.blocks]
			.reverse()
			.find((block) => block.type === "paragraph");
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
		jobs.push({ fixture, docPath, outDir });
	}

	// 2. Convert concurrently — one reusable profile per worker.
	const queue = [...jobs];
	const workers = Array.from(
		{ length: Math.min(CONCURRENCY, jobs.length) },
		async () => {
			const profile = newProfile();
			for (let job = queue.shift(); job; job = queue.shift()) {
				const { exitCode, stdout, stderr } = await runSoffice(
					job.docPath,
					job.outDir,
					profile,
				);
				const converted = Bun.file(join(job.outDir, job.fixture));
				results.set(job.fixture, {
					exitCode,
					stdout,
					stderr,
					convertedBytes: (await converted.exists())
						? (await converted.arrayBuffer()).byteLength
						: 0,
				});
			}
		},
	);
	await Promise.all(workers);
}, 120_000);

afterAll(() => {
	if (workspace) rmSync(workspace, { recursive: true, force: true });
	for (const profile of profiles) {
		rmSync(profile, { recursive: true, force: true });
	}
});

function newProfile(): string {
	const profile = mkdtempSync(join(tmpdir(), "docx-cli-soffice-"));
	profiles.push(profile);
	return profile;
}

describe("LibreOffice round-trip", () => {
	if (!SOFFICE) {
		test.skip("soffice not on PATH — skipping integration tests", () => {});
		return;
	}

	for (const fixture of FIXTURES) {
		test(`${fixture} re-saves cleanly through the create→insert→read pipeline`, () => {
			// Conversion ran in the beforeAll pool; assert its precomputed result.
			const result = results.get(fixture);
			expect(result).toBeDefined();
			if (!result) return;
			expect(result.exitCode).toBe(0);
			// LibreOffice emits known-harmless noise depending on the host:
			// - macOS: "Task policy" lines
			// - Linux/CI: "Fontconfig error: No writable cache directories" when
			//   parallel soffice workers race to populate ~/.cache/fontconfig.
			//   We pre-warm in CI (fc-cache -f), but a tighter race or a stale
			//   profile can still print one line — ignore it. The conversion
			//   succeeds either way (font cache only matters for glyph rendering;
			//   docx-to-docx is a format transform).
			const meaningfulStderr = result.stderr
				.split("\n")
				.filter(
					(line) =>
						!line.includes("Task policy") &&
						!line.includes("Warning") &&
						!line.includes("Fontconfig error") &&
						line.trim().length > 0,
				)
				.join("\n");
			expect(meaningfulStderr).toBe("");
			expect(result.stdout).toContain("convert");
			expect(result.convertedBytes).toBeGreaterThan(0);
		});
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
		const { exitCode } = await runSoffice(docPath, outDir, newProfile());
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

/** Spawn `soffice --headless --convert-to docx` against an isolated
 * user-profile directory. Without -env:UserInstallation, every soffice process
 * competes for a lock on the default profile (~/Library/Application Support/
 * LibreOffice/4 on macOS); a stale or concurrent soffice then exits non-zero —
 * a real flake source. The caller passes a `profile` it owns exclusively for
 * the duration of the call, so conversions can run concurrently. */
async function runSoffice(
	docPath: string,
	outDir: string,
	profile: string,
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
	if (!SOFFICE) throw new Error("soffice not on PATH");
	const proc = Bun.spawn(
		[
			SOFFICE,
			`-env:UserInstallation=file://${profile}`,
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
