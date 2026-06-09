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
		const stagedDocx = join(stageDir, `.docx-cli-render-${tag}.docx`);
		const stagedPdf = join(stageDir, `.docx-cli-render-${tag}.pdf`);
		try {
			await Bun.write(stagedDocx, Bun.file(inputDocx));
			// Word's AppleScript `save as` with `format PDF` is the canonical
			// way to export a PDF from Word for Mac. `close ... saving no`
			// suppresses the "Save changes?" prompt (we already wrote the
			// PDF; the docx itself isn't being modified).
			const script = [
				'tell application "Microsoft Word"',
				`	open "${stagedDocx}"`,
				"	set d to active document",
				`	save as d file name "${stagedPdf}" file format format PDF`,
				"	close d saving no",
				"end tell",
			].join("\n");
			// Serialize the Word automation across processes. Word for Mac drives a
			// SINGLE app instance and the script grabs `active document`, so two
			// concurrent renders race on which document is active — one silently
			// exports the OTHER doc's pages with a success exit code (verified). An
			// advisory lock makes concurrent `docx render` invocations QUEUE instead
			// of corrupt. Only the open→save→close window needs it; staging the docx
			// and moving the PDF out use per-pid-unique paths and are race-free.
			await withRenderLock(async () => {
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
			});
		} finally {
			rmSync(stagedDocx, { force: true });
			rmSync(stagedPdf, { force: true });
		}
	},
};

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
