import { basename, dirname } from "node:path";
// Embedded at BUILD time, never fetched — see the installer invariant in the root
// CLAUDE.md for why that distinction is load-bearing.
import installScript from "../../../skills/docx-cli/scripts/install.sh" with {
	type: "text",
};
import { VERSION } from "../help";
import {
	EXIT,
	fail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStderr,
	writeStdout,
} from "../respond";

const HELP = `docx upgrade — replace the installed docx binary with a newer release

Usage:
  docx upgrade [options]

Options:
  --to TAG     Install a specific release tag (e.g. v0.23.0). Default: latest
  --dry-run    Report what would be installed; download and change nothing
  --verbose    Full JSON ack instead of the one-line confirmation
  -h, --help   Show this help

Only the STANDALONE binary upgrades itself, replacing it where it already lives
(PREFIX is ignored). An npm/bun install is owned by the package manager, so
upgrade reports that and exits nonzero rather than fighting it.

It runs the same installer published as a release asset — embedded in this
binary, not downloaded — so the new binary is pinned to a release tag and its
SHA-256 is verified against that release's SHA256SUMS before it replaces
anything.

Output:
  Success       one-line ack (--verbose → {ok, operation, from, to, path})
  --dry-run     {operation, dryRun, from, to, path}
  Not upgraded  {code, error, hint} + nonzero exit (npm install, Windows, or a
                failed install) — the binary is left untouched
`;

/** A release tag, and nothing that can steer the URL somewhere else. `--to` is
 *  interpolated into install.sh's release URL, and curl NORMALIZES `..` — so an
 *  unvalidated value ("../../../../other/repo/releases/download/v1") redirects both
 *  the binary AND the SHA256SUMS it is checked against to an attacker's repo, making
 *  the checksum gate verify the attacker's own manifest. Agents drive this CLI over
 *  untrusted document content, so the value must be constrained here, at the CLI
 *  ingress, per "the CLI surface is concerned with parsing user input". */
const RELEASE_TAG = /^v?\d+\.\d+\.\d+(?:[-.][0-9A-Za-z.-]+)?$/;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			to: { type: "string" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;
	setVerboseAck(Boolean(parsed.values.verbose));

	// `docx upgrade v0.23.0` is the shape an agent reaches for, and parseArgs would
	// silently drop it into positionals — upgrading to LATEST while exiting 0, i.e. a
	// confidently-wrong result at the one moment a weak agent stops reading.
	const [stray] = parsed.positionals;
	if (stray !== undefined) {
		return fail(
			"USAGE",
			`upgrade takes no positional arguments (got ${JSON.stringify(stray)})`,
			"Pin a release with the flag:  docx upgrade --to v0.23.0",
		);
	}

	const requested = parsed.values.to as string | undefined;
	if (requested !== undefined && !RELEASE_TAG.test(requested)) {
		return fail(
			"USAGE",
			`--to must be a release tag like v0.23.0 (got ${JSON.stringify(requested)})`,
			"Tags are listed at https://github.com/kklimuk/docx-cli/releases",
		);
	}
	// `docx --version` prints "0.23.0" but the release is tagged "v0.23.0"; accept what
	// the tool itself printed rather than 404ing on the most likely typo.
	const target =
		requested === undefined
			? "latest"
			: requested.startsWith("v")
				? requested
				: `v${requested}`;

	if (!isStandaloneBinary()) {
		// Self-replacing a package-manager-owned install would strand the manager's
		// metadata on bytes it never wrote and put a second `docx` on PATH. Nonzero:
		// nothing was upgraded, and weak agents read the exit code, not the prose.
		return fail(
			"UPGRADE_FAILED",
			`docx ${VERSION} was installed from the npm registry, which owns updates`,
			"bun add -g bun-docx      # or: npm install -g bun-docx",
		);
	}

	const binaryPath = process.execPath;
	const prefix = dirname(binaryPath);

	if (process.platform === "win32") {
		// POSIX can rename over a running binary's inode; Windows cannot.
		return fail(
			"UPGRADE_FAILED",
			`Windows cannot replace ${binaryPath} while it is running`,
			"Download the new binary and swap it while docx is not running: https://github.com/kklimuk/docx-cli/releases/latest",
		);
	}

	// install.sh always writes `${PREFIX}/docx` — it owns the platform→name rule. So a
	// binary the user RENAMED cannot be self-replaced: the install would drop a second
	// file beside it (clobbering any unrelated `docx` already there) and leave the
	// command the user actually runs on the old version, while we reported success.
	// Refuse instead of reporting a replacement that did not happen.
	const binaryName = basename(binaryPath);
	if (binaryName !== "docx") {
		return fail(
			"UPGRADE_FAILED",
			`the installer places the binary as "docx", but this one is named "${binaryName}" — upgrading would leave ${binaryPath} on ${VERSION} and write a separate ${prefix}/docx`,
			`Replace it by hand: download docx-<platform> from https://github.com/kklimuk/docx-cli/releases/latest, verify it against SHA256SUMS, and move it over ${binaryPath}.`,
		);
	}

	if (parsed.values["dry-run"]) {
		await respond({
			operation: "upgrade",
			dryRun: true,
			from: VERSION,
			to: target,
			path: binaryPath,
		});
		return EXIT.OK;
	}

	// Pipe the embedded script to `sh -s` rather than staging a temp file: install.sh
	// reads no stdin and uses no $0, so this leaves nothing behind in a PATH directory,
	// and it hands the PREFIX-writability decision back to install.sh, whose error for
	// that case is far better than an EACCES escaping from a temp write.
	const proc = Bun.spawn(["sh", "-s"], {
		stdin: new TextEncoder().encode(installScript),
		stdout: "pipe",
		stderr: "pipe",
		env: {
			...process.env,
			PREFIX: prefix,
			VERSION: target,
			REQUIRE_CHECKSUM: "1",
		},
	});

	// Forward as it arrives, not after it finishes: install.sh prints "→ Downloading …"
	// BEFORE the ~75 MB transfer precisely so the terminal isn't dead for a minute on a
	// slow link. Routing through respond.ts's sinks keeps the one stdout chokepoint.
	await Promise.all([
		forward(proc.stdout, writeStdout),
		forward(proc.stderr, writeStderr),
	]);

	if ((await proc.exited) !== 0) {
		// install.sh has already narrated the specific cause on stderr; don't paper over
		// it with a blanket "use npm instead" that misdirects on a bad tag or a blip.
		// The one thing worth correcting: its unwritable-PREFIX hint says to re-run with
		// PREFIX set, which `upgrade` overrides (it replaces the binary where it lives).
		return fail(
			"UPGRADE_FAILED",
			`upgrade failed — ${binaryPath} is unchanged`,
			`The installer printed the cause above. If it was a permissions problem, re-run with write access to ${prefix} — upgrade replaces the binary where it lives, so setting PREFIX has no effect here.`,
		);
	}

	await respondAck({
		operation: "upgrade",
		from: VERSION,
		to: target,
		path: binaryPath,
	});
	return EXIT.OK;
}

/** Stream one of the installer's pipes to a sink as it arrives. The decoder is
 *  per-stream and in `stream: true` mode: install.sh's output is UTF-8 ("→", "✓"),
 *  and a chunk boundary can land mid-sequence — a stateless decode would turn that
 *  character into U+FFFD. The final `decode()` flushes whatever is buffered. */
async function forward(
	stream: ReadableStream<Uint8Array>,
	write: (text: string) => Promise<void>,
): Promise<void> {
	const decoder = new TextDecoder();
	for await (const chunk of stream) {
		const text = decoder.decode(chunk, { stream: true });
		if (text) await write(text);
	}
	const tail = decoder.decode();
	if (tail) await write(tail);
}

/** Bun mounts a compiled single-file executable's own sources on a virtual FS —
 *  POSIX "/$bunfs/", Windows "B:\\~BUN". That's the definitive "I am a standalone
 *  binary" signal; an npm install runs dist/index.js from a real path under Bun. */
function isStandaloneBinary(): boolean {
	const path = import.meta.path;
	return path.startsWith("/$bunfs/") || path.includes("~BUN");
}
