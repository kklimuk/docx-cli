import { existsSync, rmSync } from "node:fs";
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
