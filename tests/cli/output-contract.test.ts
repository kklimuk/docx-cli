import { describe, expect, test } from "bun:test";
import { spawn } from "node:child_process";
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

// Use real OS pipes, bounded in both time and bytes: the old BunFile.write
// path could repeat a prefix forever after process.stdout was materialized.
const RESPOND_MODULE = join(import.meta.dir, "../../src/cli/respond.ts");
const PIPE_PAYLOAD = pipePayload();

function pipePayload(): string {
	return Array.from(
		{ length: 12000 },
		(_, index) => `${index}: héllo 漢 🦊\n`,
	).join("");
}

function pipeScript(body: string): string {
	return `
		import { writeStdout, writeStderr, respond, captureOutput } from ${JSON.stringify(RESPOND_MODULE)};
		const text = (${pipePayload.toString()})();
		${body}
	`;
}

function runPipeScript(
	body: string,
	stream: "stdout" | "stderr" = "stdout",
	consumer?: string,
) {
	return runPipeCommand(
		[process.execPath, "-e", pipeScript(body)],
		stream,
		consumer,
	);
}

async function runPipeCommand(
	command: string[],
	stream: "stdout" | "stderr" = "stdout",
	consumer = "{ dd bs=1 count=1 2>/dev/null; sleep 0.1; cat; }",
) {
	// Bun's subprocess capture can use sockets, which hide this OS-pipe bug.
	// bash supplies a real pipe on the chosen descriptor; pipefail preserves the
	// producer's exit status. A separate process group lets the bounds kill
	// the entire pipeline, even if a regressed producer spins forever. The
	// consumer pauses after its first byte so the producer meets backpressure.
	const producer = spawn(
		"bash",
		[
			"-o",
			"pipefail",
			"-c",
			`${stream === "stderr" ? '"$@" 3>&1 1>&2 2>&3' : '"$@"'} | ${consumer}`,
			"docx-output-test",
			...command,
		],
		{ detached: true, stdio: ["ignore", "pipe", "pipe"] },
	);
	const chunks: { stdout: Buffer[]; stderr: Buffer[] } = {
		stdout: [],
		stderr: [],
	};
	let size = 0;
	let error: Error | undefined;
	function stop(reason: Error) {
		error ??= reason;
		if (producer.pid) {
			try {
				process.kill(-producer.pid, "SIGKILL");
			} catch {
				/* already exited */
			}
		}
	}
	const timer = setTimeout(
		() => stop(new Error("Pipe producer timed out")),
		3000,
	);
	for (const stream of ["stdout", "stderr"] as const) {
		producer[stream]?.on("data", (chunk: Buffer) => {
			size += chunk.length;
			if (size > 2 * 1024 * 1024) {
				stop(new Error("Pipe producer exceeded output limit"));
				return;
			}
			chunks[stream].push(chunk);
		});
	}
	try {
		const status = await new Promise<number | null>((resolve, reject) => {
			producer.once("error", reject);
			producer.once("close", resolve);
		});
		const output = {
			stdout: Buffer.concat(chunks.stdout).toString("utf8"),
			stderr: Buffer.concat(chunks.stderr).toString("utf8"),
		};
		// Undo the shell's descriptor swap for a stderr-pipe probe.
		if (stream === "stderr")
			[output.stdout, output.stderr] = [output.stderr, output.stdout];
		return { status, error, ...output };
	} finally {
		clearTimeout(timer);
		stop(new Error("Pipe test cleanup"));
	}
}

// bash's OS-pipe setup is POSIX-only. The regular CLI smoke tests above also
// cover Windows; these specifically guard the macOS/Linux pipe regression.
describe.skipIf(process.platform === "win32")(
	"output sinks through OS pipes",
	() => {
		test("large CLI Markdown and AST match captured output through a slow pipe", async () => {
			const workspace = tempWorkspace("large-read-pipe");
			const docPath = join(workspace, "large.docx");
			const inputPath = join(workspace, "input.txt");
			await Bun.write(inputPath, PIPE_PAYLOAD);
			expect(
				(await rawCli("create", docPath, "--text-file", inputPath)).exitCode,
			).toBe(0);
			for (const flags of [[], ["--ast"]]) {
				const args = ["read", docPath, ...flags];
				const expected = await runCli(...args);
				expect(expected.exitCode).toBe(0);
				expect(Buffer.byteLength(expected.stdout)).toBeGreaterThan(65536);
				const result = await runPipeCommand([
					process.execPath,
					join(import.meta.dir, "../../src/index.ts"),
					...args,
				]);
				expect(result.error).toBeUndefined();
				expect(result.status).toBe(0);
				expect(result.stdout === expected.stdout).toBe(true);
				expect(result.stderr).toBe("");
			}
		});

		for (const materialize of [false, true]) {
			for (const stream of ["stdout", "stderr"] as const) {
				test(`${stream} delivers large UTF-8 writes exactly (process streams touched: ${materialize})`, async () => {
					const write = stream === "stdout" ? "writeStdout" : "writeStderr";
					const result = await runPipeScript(
						`
					${materialize ? "void process.stdout; void process.stderr;" : ""}
					await ${write}("start\\n");
					await ${write}(text);
					await ${write}("end\\n");
					process.exit(0);
				`,
						stream,
					);
					expect(result.error).toBeUndefined();
					expect(result.status, result.stderr.slice(0, 1000)).toBe(0);
					expect(result[stream].length).toBe(PIPE_PAYLOAD.length + 10);
					expect(result[stream] === `start\n${PIPE_PAYLOAD}end\n`).toBe(true);
					expect(result[stream === "stdout" ? "stderr" : "stdout"]).toBe("");
				});
			}
		}

		test("respond flushes large JSON before immediate exit", async () => {
			const result = await runPipeScript(`
			void process.stdout;
			await respond({ text });
			process.exit(0);
		`);
			expect(result.error).toBeUndefined();
			expect(result.status).toBe(0);
			expect(
				result.stdout === `${JSON.stringify({ text: PIPE_PAYLOAD })}\n`,
			).toBe(true);
		});

		test("captureOutput can restore real sinks and allow natural exit", async () => {
			const result = await runPipeScript(`
			let captured = "";
			captureOutput(async value => { captured += value; }, async value => { captured += value; });
			await writeStdout("captured-out");
			await writeStderr("captured-err");
			if (captured !== "captured-outcaptured-err") process.exit(9);
			captureOutput();
			await writeStdout(text);
			await writeStderr("restored");
		`);
			expect(result.error).toBeUndefined();
			expect(result.status).toBe(0);
			expect(result.stdout === PIPE_PAYLOAD).toBe(true);
			expect(result.stderr).toBe("restored");
		});

		test("a consumer closing early terminates the producer without hanging", async () => {
			const result = await runPipeScript(
				`
			void process.stdout;
			for (let index = 0; index < 100; index++) await writeStdout(text);
			process.exit(0);
		`,
				"stdout",
				"head -c 1 > /dev/null",
			);
			expect(result.error).toBeUndefined();
			expect(result.status).not.toBe(0);
		});
	},
);

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
	["code", "add"],
	["code", "edit"],
	["equations", "add"],
	["equations", "edit"],
	["tasks", "add"],
	["tasks", "check"],
	["tasks", "uncheck"],
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
	["images", "add"],
	["images", "list"],
	["images", "extract"],
	["images", "replace"],
	["images", "delete"],
	["hyperlinks", "add"],
	["hyperlinks", "list"],
	["hyperlinks", "replace"],
	["hyperlinks", "delete"],
	["tables", "create"],
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

	test("find --batch - reads JSONL queries from stdin", async () => {
		const docPath = join(tempWorkspace("stdin-find-batch"), "doc.docx");
		await runCli("create", docPath, "--text", "one two one");

		const batch = `${JSON.stringify({ query: "one", nth: 1 })}\n${JSON.stringify(
			{ query: "two" },
		)}\n`;
		const result = await spawnCliStdin(batch, "find", docPath, "--batch", "-");
		expect(result.exitCode).toBe(0);
		expect(result.stdout.trim().split("\n")).toEqual(["p0:8-11", "p0:4-7"]);
	});

	test("code add --code-file - reads the code body from stdin", async () => {
		const docPath = join(tempWorkspace("stdin-code"), "doc.docx");
		await runCli("create", docPath, "--text", "intro");

		const result = await spawnCliStdin(
			"print(1)\nprint(2)\n",
			"code",
			"add",
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

// Parser-surface ergonomics (weak-agent fixes): `--help` beats a parse error,
// and dash-led positionals (money/negatives) stop being read as options. Both go
// through the shared `tryParseArgs`, so a few representative commands suffice.
describe("parser ergonomics (tryParseArgs)", () => {
	test("--help wins even when a value-taking flag is left without a value", async () => {
		// `replace --batch --help` used to die "argument is ambiguous".
		const result = await runCli("replace", "missing.docx", "--batch", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("Usage:");
	});

	test("--help placed after other flags still shows help", async () => {
		const result = await runCli(
			"edit",
			"missing.docx",
			"--at",
			"p0",
			"--text",
			"--help",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stdout).toContain("Usage:");
	});

	test("a literal '--help' text still travels via the --flag=value form", async () => {
		const docPath = join(tempWorkspace("literal-help"), "doc.docx");
		await runCli("create", docPath, "--text", "seed");
		const result = await runCli("edit", docPath, "--at", "p0", "--text=--help");
		expect(result.exitCode).toBe(0);
		expect((await runCli("read", docPath)).stdout).toContain("--help");
	});

	test("a dash-led numeric positional is a query, not an option (find)", async () => {
		const docPath = join(tempWorkspace("dash-find"), "doc.docx");
		await runCli("create", docPath, "--text", "a -5.00 discount");
		const result = await runCli("find", docPath, "-5.00");
		expect(result.exitCode).toBe(0); // not USAGE(2)
		expect(
			(result.parsed as { totalMatches: number }).totalMatches,
		).toBeGreaterThan(0);
	});

	test("dash-led positionals fill replace's pattern AND replacement", async () => {
		const docPath = join(tempWorkspace("dash-replace"), "doc.docx");
		await runCli("create", docPath, "--text", "owed -5.00 total");
		const result = await runCli("replace", docPath, "-5.00", "-9.99");
		expect(result.exitCode).toBe(0);
		expect((await runCli("read", docPath)).stdout).toContain("-9.99");
	});

	test("a positional that looks like the internal sentinel is not clobbered", async () => {
		// A dash value is shielded behind a NUL-delimited sentinel; a literal
		// positional resembling the old guessable sentinel must survive untouched.
		const docPath = join(tempWorkspace("sentinel-collision"), "doc.docx");
		await runCli("create", docPath, "--text", "token docx-dash-pos-0 here");
		const result = await runCli("replace", docPath, "docx-dash-pos-0", "-5");
		expect(result.exitCode).toBe(0); // pattern matched; not clobbered to "-5"
		expect((await runCli("read", docPath)).stdout).toContain("token -5 here");
	});

	test("a real trailing flag after dash positionals is still a flag", async () => {
		const docPath = join(tempWorkspace("dash-flag"), "doc.docx");
		await runCli("create", docPath, "--text", "owed -5 then -5 again");
		// --all must stay a flag, not get swallowed into positionals.
		const result = await runCli("replace", docPath, "-5", "-9", "--all");
		expect(result.exitCode).toBe(0);
		const markdown = (await runCli("read", docPath)).stdout;
		expect(markdown).toContain("-9 then -9 again");
		expect(markdown).not.toContain("-5");
	});
});
