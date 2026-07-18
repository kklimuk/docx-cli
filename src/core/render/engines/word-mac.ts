import {
	closeSync,
	existsSync,
	openSync,
	readFileSync,
	renameSync,
	rmSync,
	statSync,
	writeSync,
} from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { type RenderEngine, RenderEngineError } from "./types";

/** Word for Mac engine. Drives Microsoft Word via `osascript` to open the
 * docx and `save as` to PDF — the same automation oracle
 * `scripts/word-redlines.sh` uses for tracked-change probing.
 *
 * Word for Mac runs sandboxed (the app's container is
 * `~/Library/Containers/com.microsoft.Word`), so arbitrary file paths
 * trigger a "Grant File Access" prompt. We work around this by staging the
 * docx inside the container's own Documents dir, which Word can open
 * without a prompt. The first run on a fresh machine triggers a one-time
 * macOS Automation permission prompt for the controlling terminal; once
 * granted it sticks. */
export const wordMacEngine: RenderEngine = {
	name: "word-mac",
	async available(): Promise<boolean> {
		if (process.platform !== "darwin") return false;
		// `Bun.file().exists()` returns false for directories — they're not
		// "files". Use `existsSync` from node:fs for both the app bundle
		// (which is technically a directory) and the container Documents dir
		// (also a directory).
		if (!existsSync("/Applications/Microsoft Word.app")) return false;
		// Need a writable container Documents dir to stage into. Without it
		// we'd hit the "Grant File Access" prompt for every render.
		return existsSync(containerDocumentsDir());
	},
	async convertToPdf(inputDocx: string, outputPdf: string): Promise<void> {
		const stageDir = containerDocumentsDir();
		const tag = `${process.pid}-${stamp()}`;
		const stagedName = `.docx-cli-render-${tag}.docx`;
		const stagedDocx = join(stageDir, stagedName);
		const stagedPdf = join(stageDir, `.docx-cli-render-${tag}.pdf`);
		try {
			await Bun.write(stagedDocx, Bun.file(inputDocx));
			// Word's AppleScript `save as` with `format PDF` is the canonical
			// way to export a PDF from Word for Mac. `close ... saving no`
			// suppresses the "Save changes?" prompt (we already wrote the
			// PDF; the docx itself isn't being modified).
			//
			// `set display alerts to none` makes Word AUTO-DECLINE the modal
			// "Word found unreadable content — recover?" dialog a corrupt .docx
			// pops on open. Without it that dialog is MODAL and blocks the
			// AppleScript until the AppleEvent times out (-1712), wedging every
			// later render in a batch behind one bad file; and a human clicking
			// "Yes" to clear it makes Word REPAIR the file, so the exported PDF
			// then reflects recovered content, not the artifact under test
			// (silently inflating a render-fidelity comparison). With alerts off,
			// a corrupt file instead fails to open — no PDF — an honest
			// RENDER_FAILED, the correct outcome for "this .docx doesn't open in
			// Word." Verified empirically against a namespace-corrupted docx:
			// good files still render, the bad file fails without leaving a
			// modal, Word stays responsive for the next render. We apply it ONLY
			// to a Word we cold-launched ourselves (reaped after, so it can't
			// leak) and skip it when sharing the user's instance -- see the
			// `suppressAlerts` decision inside the lock below. `with timeout`
			// bounds any unexpected modal so a wedge can never outlast one render.
			//
			// The document is bound BY STAGED FILENAME, never `active document`.
			// This is a safety invariant, empirically earned: when the render
			// shares an instance with an already-open user document and the
			// staged open FAILS (corrupt file), `active document` falls through
			// to the USER'S document — `save as` then exports THEIR doc as our
			// "render" and `close saving no` closes THEIR work. Binding by name
			// can only ever resolve to our staged copy; a failed open makes the
			// `document "name"` lookup throw instead, surfacing RENDER_FAILED.
			// Timeouts are split: a corrupt file blocks the `open` AE on its
			// (auto-declined but still displayed) error dialog until the timeout
			// fires — measured at the full window, unattended — so open+bind get
			// a bounded leash to fail fast, while save-as/close keep 90s. The open
			// leash is 60s, not tighter: a corrupt file wedges for the WHOLE leash,
			// so a short one is tempting, but a legitimately large/complex .docx on
			// a cold Word launch can genuinely need most of a minute to open, and
			// false-failing a valid document (RENDER_FAILED) is worse than a slower
			// corrupt-file failure.
			//
			// Paths are escaped for the AppleScript string literals (backslash and
			// double-quote) — the staged filename is generated (safe), but the
			// container path descends from `homedir()`, so a home directory
			// containing either character would otherwise break the osascript
			// source or (worse) truncate the string onto a different path.
			// Serialize the Word automation across processes. Word for Mac drives a
			// SINGLE app instance and the script grabs `active document`, so two
			// concurrent renders race on which document is active — one silently
			// exports the OTHER doc's pages with a success exit code (verified). An
			// advisory lock makes concurrent `docx render` invocations QUEUE instead
			// of corrupt. Only the open→save→close window needs it; staging the docx
			// and moving the PDF out use per-pid-unique paths and are race-free.
			await withRenderLock(async () => {
				// Snapshot the Word PIDs that predate this render; the reap in
				// `finally` kills only what appeared during it. Why this is the one
				// safe isolation mechanism: see reapRenderWord.
				const preexistingWord = wordPids();
				// Suppress alerts ONLY when we cold-launch our own instance (no
				// pre-existing Word, snapshot known-empty). If we share the user's
				// instance, or the snapshot is uncertain (null), leave their alerts
				// alone (see the alerts note above) -- nothing here would reset them.
				const suppressAlerts =
					preexistingWord !== null && preexistingWord.size === 0;
				const script = [
					'tell application "Microsoft Word"',
					...(suppressAlerts ? ["\tset display alerts to none"] : []),
					"\twith timeout of 60 seconds",
					`\t\topen "${escapeAppleScriptString(stagedDocx)}"`,
					`\t\tset d to document "${escapeAppleScriptString(stagedName)}"`,
					"\tend timeout",
					"\twith timeout of 90 seconds",
					`\t\tsave as d file name "${escapeAppleScriptString(stagedPdf)}" file format format PDF`,
					"\t\tclose d saving no",
					"\tend timeout",
					"end tell",
				].join("\n");
				try {
					const proc = Bun.spawn(["osascript", "-e", script], {
						stdout: "pipe",
						stderr: "pipe",
					});
					const exit = await proc.exited;
					if (exit !== 0) {
						const stderr = await new Response(proc.stderr).text();
						throw new RenderEngineError(
							"RENDER_FAILED",
							`Word for Mac failed (exit ${exit}): ${stderr.trim() || "(no stderr)"}`,
							"If this is the first run, macOS may have prompted for Automation permission. Grant it under System Settings → Privacy & Security → Automation, then retry.",
						);
					}
					const stagedFile = Bun.file(stagedPdf);
					if (!(await stagedFile.exists())) {
						throw new RenderEngineError(
							"RENDER_FAILED",
							"Word reported success but produced no PDF at the staged path",
							`Expected: ${stagedPdf}`,
						);
					}
					await Bun.write(outputPdf, stagedFile);
				} finally {
					reapRenderWord(preexistingWord);
				}
			});
		} finally {
			rmSync(stagedDocx, { force: true });
			rmSync(stagedPdf, { force: true });
		}
	},
};

/** Escape a string for embedding inside an AppleScript double-quoted literal:
 * backslash first, then the double-quote. Applied to every path we interpolate
 * into the render script — the paths descend from `homedir()`, so a `"` or `\`
 * in the home directory would otherwise break or hijack the string. */
function escapeAppleScriptString(value: string): string {
	return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

/** All currently-running Microsoft Word process ids, or `null` if pgrep itself
 * ERRORED (exit ≥ 2 — a usage/internal failure, distinct from exit 1 "no match").
 * The distinction is load-bearing for the reap: a transient pgrep failure must
 * not read as "no Word running," which would leave `preexisting` empty and let
 * the reap SIGKILL the user's own instance. When we can't be sure, we return
 * null and the caller reaps nothing. */
function wordPids(): Set<number> | null {
	const out = Bun.spawnSync(["pgrep", "-x", "Microsoft Word"]);
	if (out.exitCode === 1) return new Set(); // pgrep: matched no process
	if (out.exitCode !== 0) return null; // pgrep errored — can't determine
	return new Set(
		out.stdout.toString().trim().split("\n").filter(Boolean).map(Number),
	);
}

/** Terminate ONLY the Word process(es) that appeared DURING this render — the
 * instance our own `tell` cold-launched, or a wedged one a corrupt file left —
 * leaving every pre-existing Word untouched. A Word the user already had open
 * predates `preexisting`, so it can never be in the kill set, and its documents
 * are safe. Renders are serialized by withRenderLock, so the only Word that
 * should appear between the snapshot and here is ours (a user launching Word in
 * that few-second window is the one accepted race — their just-launched, still
 * empty instance would be reaped once).
 *
 * We kill by PID, NOT `killall` (which would also take out the user's Word),
 * and we NEVER quit via AppleScript. Both halves are empirically earned:
 * spawning extra instances with `open -n` works, but with two same-bundle-id
 * instances running, Apple Event routing is NONDETERMINISTIC — AppleScript
 * `tell`, JXA `Application(pid)`, and frontmost-first all delivered events to
 * the WRONG instance in testing (a `quit saving no` could discard the user's
 * unsaved work; an `open` landed in the user's Word). So a dedicated
 * second-instance-per-render cannot be driven safely, renders stay serialized
 * on the single addressable instance, and kill(2) — the one interface macOS
 * makes truly PID-precise — is the only cleanup mechanism we allow ourselves.
 * SIGKILL is safe here: our instance's staged doc is already closed (or never
 * opened, for a corrupt file), so no recovery state is left behind. */
function reapRenderWord(preexisting: Set<number> | null): void {
	// Fail closed: if EITHER snapshot is unavailable (pgrep errored), we can't
	// tell our instance from the user's, so reap nothing rather than risk their
	// Word. Leaving our own instance running is harmless (the next render reaps
	// it, or it's an idle empty Word).
	if (!preexisting) return;
	const current = wordPids();
	if (!current) return;
	for (const pid of current) {
		if (preexisting.has(pid)) continue;
		try {
			process.kill(pid, "SIGKILL");
		} catch {
			// Already gone — nothing to do.
		}
	}
}

/** The sandboxed Documents dir Word for Mac can access without an explicit
 * file-access prompt. First-launch of Word on a clean install creates this;
 * if it doesn't exist, the user hasn't run Word yet. */
function containerDocumentsDir(): string {
	return join(
		homedir(),
		"Library/Containers/com.microsoft.Word/Data/Documents",
	);
}

/** A short non-random suffix for the staged filename. Goal is just
 * intra-process uniqueness (PID + monotonic counter); the cross-process
 * case is covered by the PID component. We don't use `Date.now()` /
 * `Math.random()` here because the bundled JSX runtime forbids them in
 * scripts that may run under our workflow harness — staying consistent
 * with that convention even outside Workflow context keeps the API clean. */
let monotonic = 0;
function stamp(): string {
	monotonic += 1;
	return String(monotonic);
}

/** A holder older than this is presumed crashed and its lock is stolen, so a dead
 * render can't deadlock the queue. A single Word render takes seconds, never
 * minutes, so this is generous enough to never steal a live holder's lock. */
const RENDER_LOCK_STALE_MS = 5 * 60 * 1000;

/** Serialize the Word automation across processes via an advisory lock file. Word
 * for Mac is a single app instance whose AppleScript `active document` is global,
 * so concurrent renders silently corrupt each other. We gate the open→save→close
 * window on an exclusive-create lock: exactly one process wins `open(…, "wx")`;
 * the rest wait and retry until it's released (or stolen if stale). */
async function withRenderLock<T>(critical: () => Promise<T>): Promise<T> {
	const lockPath = join(containerDocumentsDir(), ".docx-cli-render.lock");
	// A token unique to THIS acquisition (pid + monotonic). We only ever delete a
	// lock whose contents still match it, so a holder whose lock was stolen as
	// stale can't delete the thief's fresh lock on the way out.
	const token = `${process.pid}:${stamp()}`;
	await acquireRenderLock(lockPath, token);
	try {
		return await critical();
	} finally {
		releaseRenderLock(lockPath, token);
	}
}

async function acquireRenderLock(
	lockPath: string,
	token: string,
): Promise<void> {
	for (;;) {
		try {
			// "wx" = create exclusively; throws EEXIST if another holder has it.
			const fd = openSync(lockPath, "wx");
			writeSync(fd, token);
			closeSync(fd);
			return;
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
			if (renderLockIsStale(lockPath)) stealStaleLock(lockPath, token);
			else await Bun.sleep(150);
		}
	}
}

/** Atomically steal a presumed-dead lock. `rename` is atomic on POSIX, so if
 *  two waiters both see the same stale lock, exactly one wins the rename (moving
 *  the inode aside) and the other gets ENOENT — preventing the blind-`rmSync`
 *  race where both delete and both then win `openSync('wx')`, driving the single
 *  Word instance concurrently. The winner clears the stolen file; the lock is
 *  then free and the next `openSync('wx')` (by anyone) re-establishes a single
 *  owner. Loser/ENOENT just loops. */
function stealStaleLock(lockPath: string, token: string): void {
	const stealPath = `${lockPath}.steal.${token}`;
	try {
		renameSync(lockPath, stealPath);
	} catch {
		// Someone else stole/released it first — just retry the acquire loop.
		return;
	}
	rmSync(stealPath, { force: true });
}

/** Release only if the lock still holds OUR token. If it was stolen as stale and
 *  re-created by another process, the token differs and we leave it alone. */
function releaseRenderLock(lockPath: string, token: string): void {
	try {
		if (readFileSync(lockPath, "utf8") === token) {
			rmSync(lockPath, { force: true });
		}
	} catch {
		// Already gone (stolen/released) — nothing to do.
	}
}

/** True if the lock is older than the stale threshold (holder presumed dead) or
 * vanished between the failed acquire and this check (already free). Uses the wall
 * clock — this runs in the CLI process, not the workflow sandbox that bans
 * Date.now(); see the note on `stamp()` above. */
function renderLockIsStale(lockPath: string): boolean {
	try {
		return Date.now() - statSync(lockPath).mtimeMs > RENDER_LOCK_STALE_MS;
	} catch {
		return true;
	}
}
