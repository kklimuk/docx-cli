import { describe, expect, test } from "bun:test";
import { join } from "node:path";
import {
	spawnCli as rawCli,
	runCli,
	spawnCli,
	spawnCliStdin,
	tempWorkspace,
} from "./harness";

// Mutators print a concise text-first confirmation line on success by default;
// --verbose swaps that for the full JSON ack. The shared harness auto-injects
// --verbose for mutators (so the rest of the suite keeps asserting on JSON
// acks); these tests bypass it by spawning the REAL binary via spawnCli. They
// ARE the binary smoke layer (real process boundary + exit codes) now that the
// bulk of the suite runs in-process — the process-level cases that once lived
// in binary-smoke.test.ts are folded in here.

describe("docx mutators — confirm by default", () => {
	test("create prints a confirmation line when --verbose is omitted", async () => {
		const workspace = tempWorkspace("quiet-create");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli("create", docPath, "--text", "hello");
		expect(result.exitCode).toBe(0);
		// Not silent: a one-line, text-first confirmation (not JSON).
		expect(result.stdout.trim().length).toBeGreaterThan(0);
		expect(result.stdout).toContain("create");
		expect(result.stdout.trim().startsWith("{")).toBe(false);
	});

	test("create with --verbose prints the JSON ack", async () => {
		const workspace = tempWorkspace("verbose-create");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli(
			"create",
			docPath,
			"--text",
			"hello",
			"--verbose",
		);
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "create" });
	});

	test("-v shorthand also enables the ack", async () => {
		const workspace = tempWorkspace("verbose-short");
		const docPath = join(workspace, "out.docx");
		const result = await rawCli("create", docPath, "--text", "hello", "-v");
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ ok: true });
	});

	test("errors print regardless of --verbose", async () => {
		const result = await rawCli("create", "--text", "hello"); // missing FILE
		expect(result.exitCode).not.toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ code: "USAGE" });
	});

	test("--dry-run prints regardless of --verbose", async () => {
		const workspace = tempWorkspace("dry-run-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		const result = await rawCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"new para",
			"--dry-run",
		);
		expect(result.exitCode).toBe(0);
		const payload = JSON.parse(result.stdout.trim());
		expect(payload).toMatchObject({ dryRun: true });
	});

	test("a layout-affecting mutator appends a render-verify hint (quiet mode)", async () => {
		const workspace = tempWorkspace("layout-hint");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		// A multi-column section is layout-bearing — read can't show how it flows,
		// so the `sections` success ack nudges a render. Plain text inserts must NOT.
		const layout = await rawCli(
			"sections",
			docPath,
			"--at",
			"p0",
			"--columns",
			"2",
		);
		expect(layout.exitCode).toBe(0);
		expect(layout.stdout).toContain("docx render");
		const plain = await rawCli(
			"insert",
			docPath,
			"--after",
			"p0",
			"--text",
			"just text",
		);
		expect(plain.exitCode).toBe(0);
		expect(plain.stdout).not.toContain("docx render");
	});

	test("a default replace that leaves matches behind nudges --all (quiet mode)", async () => {
		const workspace = tempWorkspace("replace-partial");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "fox fox fox");
		const result = await rawCli("replace", docPath, "fox", "cat");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("--all");
		expect(result.stdout).toMatch(/1 of 3/);
		// A full sweep stays silent about remaining matches.
		await rawCli("create", docPath, "--text", "fox fox fox", "--force");
		const full = await rawCli("replace", docPath, "fox", "cat", "--all");
		expect(full.stdout).not.toContain("--all to replace");
	});

	test("read commands stay loud (no --verbose needed)", async () => {
		const workspace = tempWorkspace("read-loud");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		// Query verbs default to text output (no --verbose / --json needed) — the
		// bare match locator line proves they're loud-by-default, unlike mutators.
		const result = await rawCli("find", docPath, "hello");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.length).toBeGreaterThan(0);
		expect(result.stdout.trim()).toBe("p0:0-5");
	});

	test("track-changes on confirms by default; --verbose prints the toggle ack", async () => {
		const workspace = tempWorkspace("tc-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello");
		const quiet = await rawCli("track-changes", docPath, "on");
		expect(quiet.exitCode).toBe(0);
		// Text-first confirmation, not JSON.
		expect(quiet.stdout.trim()).toBe("track-changes tracking on");

		await rawCli("track-changes", docPath, "off");
		const verbose = await rawCli("track-changes", docPath, "on", "--verbose");
		expect(verbose.exitCode).toBe(0);
		const payload = JSON.parse(verbose.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "track-changes" });
	});

	test("comments add prints the minted cN by default; --verbose returns commentId", async () => {
		const workspace = tempWorkspace("comments-quiet");
		const docPath = join(workspace, "out.docx");
		await rawCli("create", docPath, "--text", "hello world");

		// Handle-minting mutators print the bare locator line by default.
		const quiet = await rawCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"check",
		);
		expect(quiet.exitCode).toBe(0);
		expect(quiet.stdout.trim()).toBe("c0");

		const verbose = await rawCli(
			"comments",
			"add",
			docPath,
			"--at",
			"p0",
			"--text",
			"check 2",
			"--verbose",
		);
		expect(verbose.exitCode).toBe(0);
		const payload = JSON.parse(verbose.stdout.trim());
		expect(payload).toMatchObject({ ok: true, operation: "comments.add" });
		expect(payload).toHaveProperty("commentId");
	});
});

// The bulk of the CLI suite runs in-process (see harness.ts runCli) for speed.
// These few tests spawn the REAL binary (`bun src/index.ts`) so the process
// boundary itself stays covered: argv handling, exit-code propagation via
// src/index.ts, and the 64 KB Bun.stdout truncation that respond.ts guards.
// quiet-default.test.ts also spawns (quiet/verbose acks); keep subprocess tests
// to this handful.

const FIXTURES = join(import.meta.dir, "..", "fixtures");

describe("binary smoke (real subprocess)", () => {
	test("--version prints the version and exits 0", async () => {
		const result = await spawnCli("--version");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toMatch(/^docx \d+\.\d+\.\d+/);
	});

	test("--help prints usage and exits 0", async () => {
		const result = await spawnCli("--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("Usage");
	});

	test("unknown command exits non-zero with an error payload", async () => {
		const result = await spawnCli("frobnicate");
		expect(result.exitCode).not.toBe(0);
		expect(result.parsed).toMatchObject({ code: "USAGE" });
	});

	test("not-found locator propagates exit code 3", async () => {
		const workspace = tempWorkspace("smoke-exit");
		const docPath = join(workspace, "out.docx");
		await spawnCli("create", docPath, "--text", "hi");
		const result = await spawnCli("edit", docPath, "--at", "p9", "--text", "x");
		expect(result.exitCode).toBe(3);
	});

	test("large stdout (>64 KB) is not truncated", async () => {
		// academic-paper.docx --ast is ~77 KB — past the 64 KB write boundary that
		// respond.ts/writeStdout exist to handle. The output must arrive whole.
		const result = await spawnCli(
			"read",
			join(FIXTURES, "academic-paper.docx"),
			"--ast",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout.length).toBeGreaterThan(64 * 1024);
		// Complete + valid JSON end-to-end (truncation would break the parse).
		expect(() => JSON.parse(result.stdout)).not.toThrow();
	});
});

// The full command tree. Every command and sub-verb must answer `--help` with a
// usable screen — this is the regression guard for the help-drift bug class
// (an implemented flag with no docs, or docs for a flag that doesn't exist).
const COMMANDS: string[][] = [
	["create"],
	["read"],
	["edit"],
	["insert"],
	["delete"],
	["find"],
	["replace"],
	["wc"],
	["outline"],
	["render"],
	["info", "schema"],
	["info", "locators"],
	["comments", "add"],
	["comments", "reply"],
	["comments", "resolve"],
	["comments", "delete"],
	["comments", "list"],
	["footnotes", "add"],
	["footnotes", "edit"],
	["footnotes", "delete"],
	["footnotes", "list"],
	["endnotes", "add"],
	["endnotes", "edit"],
	["endnotes", "delete"],
	["endnotes", "list"],
	["images", "list"],
	["images", "extract"],
	["images", "replace"],
	["images", "delete"],
	["hyperlinks", "add"],
	["hyperlinks", "list"],
	["hyperlinks", "replace"],
	["hyperlinks", "delete"],
	["tables", "insert-row"],
	["tables", "delete-row"],
	["tables", "insert-column"],
	["tables", "delete-column"],
	["tables", "set-widths"],
	["tables", "merge"],
	["tables", "unmerge"],
	["tables", "borders"],
	["track-changes", "list"],
	["track-changes", "accept"],
	["track-changes", "reject"],
	["track-changes", "apply"],
];

// Commands that take a locator advertise the unified `--at` (or the placement /
// slice variants) — none should still mention a removed addressing flag.
const REMOVED_ADDRESSING_FLAGS = ["--range ", "--id ", "--to cN", "--to ID"];

describe("help smoke", () => {
	for (const command of COMMANDS) {
		const label = command.join(" ");
		test(`docx ${label} --help`, async () => {
			const result = await runCli(...command, "--help");
			expect(result.exitCode).toBe(0);
			expect(result.stdout).toContain("Usage:");
			// The comprehensive-help pass gave every command an Output section
			// describing its success/error shape. The `info` reference printers
			// are the exception — they ARE the output, described inline.
			if (command[0] !== "info") {
				expect(result.stdout).toContain("Output:");
			}
			for (const removed of REMOVED_ADDRESSING_FLAGS) {
				expect(result.stdout).not.toContain(removed);
			}
		});
	}
});

// The `-` (stdin) ingress for --batch / --code-file / --markdown-file / --from
// can only be exercised through the real process boundary — the in-process
// runCli has no stdin to feed. These pin that wiring end-to-end.
describe("stdin '-' ingress (process boundary)", () => {
	test("edit --batch - applies a JSONL batch piped on stdin", async () => {
		const docPath = join(tempWorkspace("stdin-batch"), "doc.docx");
		await runCli("create", docPath, "--text", "one");
		await runCli("insert", docPath, "--after", "p0", "--text", "two");

		const batch = `${JSON.stringify({ at: "p0", text: "ONE" })}\n${JSON.stringify(
			{ at: "p1", text: "TWO" },
		)}\n`;
		const result = await spawnCliStdin(batch, "edit", docPath, "--batch", "-");
		expect(result.exitCode).toBe(0);

		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown).toContain("ONE");
		expect(markdown).toContain("TWO");
	});

	test("insert --code-file - reads the code body from stdin", async () => {
		const docPath = join(tempWorkspace("stdin-code"), "doc.docx");
		await runCli("create", docPath, "--text", "intro");

		const result = await spawnCliStdin(
			"print(1)\nprint(2)\n",
			"insert",
			docPath,
			"--after",
			"p0",
			"--code-file",
			"-",
			"--language",
			"python",
		);
		expect(result.exitCode).toBe(0);

		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown).toContain("print(1)");
	});
});
