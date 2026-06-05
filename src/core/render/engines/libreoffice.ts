import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { type RenderEngine, RenderEngineError } from "./types";

/** Cross-platform LibreOffice engine. Drives `soffice --headless
 * --convert-to pdf` against the input docx. Used as the fallback when Word
 * isn't installed; also usable explicitly on a server / CI machine where
 * Word isn't an option.
 *
 * Each invocation mints a fresh user profile (`-env:UserInstallation`)
 * because macOS LibreOffice locks its default profile while running —
 * concurrent or stale `soffice` processes make new spawns exit non-zero.
 * This mirrors the LibreOffice round-trip integration test's worker-pool
 * pattern; ALL spawners of soffice in this codebase must take their own
 * exclusive profile or they'll flake. */
export const libreofficeEngine: RenderEngine = {
	name: "libreoffice",
	async available(): Promise<boolean> {
		return (await findSoffice()) !== null;
	},
	async convertToPdf(inputDocx: string, outputPdf: string): Promise<void> {
		const soffice = await findSoffice();
		if (!soffice) {
			throw new RenderEngineError(
				"RENDER_ENGINE",
				"LibreOffice (`soffice`) not found on PATH",
				"Install via `brew install --cask libreoffice` (macOS), `apt install libreoffice` (Linux), or download from libreoffice.org (Windows).",
			);
		}
		const profile = mkdtempSync(join(tmpdir(), "docx-cli-soffice-"));
		const outDir = dirname(outputPdf);
		try {
			const proc = Bun.spawn(
				[
					soffice,
					"--headless",
					`-env:UserInstallation=file://${profile}`,
					"--convert-to",
					"pdf",
					"--outdir",
					outDir,
					inputDocx,
				],
				{ stdout: "pipe", stderr: "pipe" },
			);
			const exit = await proc.exited;
			if (exit !== 0) {
				const stderr = await new Response(proc.stderr).text();
				throw new RenderEngineError(
					"RENDER_FAILED",
					`soffice exited with code ${exit}: ${stderr.trim() || "(no stderr)"}`,
				);
			}
			// soffice names its output `<basename>.pdf` in --outdir. If that
			// path differs from the caller's requested output, rename — keeps
			// the engine interface honest (caller picks the output path).
			const sofficeOutput = join(
				outDir,
				`${basenameWithoutExt(inputDocx)}.pdf`,
			);
			if (sofficeOutput !== outputPdf) {
				await Bun.write(outputPdf, Bun.file(sofficeOutput));
				rmSync(sofficeOutput, { force: true });
			}
		} finally {
			rmSync(profile, { recursive: true, force: true });
		}
	},
};

async function findSoffice(): Promise<string | null> {
	// PATH search first (the brew install on macOS, apt on Linux, the LO
	// installer's program/ dir on Windows when added to PATH). Then check
	// the canonical install paths in case PATH wasn't set up.
	const onPath = await whichEverywhere("soffice");
	if (onPath) return onPath;
	const fallbacks =
		process.platform === "darwin"
			? ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
			: process.platform === "win32"
				? [
						"C:\\Program Files\\LibreOffice\\program\\soffice.com",
						"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
					]
				: ["/usr/bin/soffice", "/snap/bin/libreoffice"];
	for (const candidate of fallbacks) {
		const file = Bun.file(candidate);
		if (await file.exists()) return candidate;
	}
	return null;
}

/** `command -v NAME` (POSIX) / `where NAME` (Windows) wrapped to return the
 * first match path or null. We use this instead of just spawning the bare
 * command because we want to know if the lookup itself succeeded. */
async function whichEverywhere(name: string): Promise<string | null> {
	const args =
		process.platform === "win32" ? ["where", name] : ["command", "-v", name];
	// `command` is a shell builtin on POSIX; route through /bin/sh -c.
	const cmd =
		process.platform === "win32"
			? args
			: ["/bin/sh", "-c", `command -v ${name}`];
	const proc = Bun.spawn(cmd, { stdout: "pipe", stderr: "pipe" });
	const out = await new Response(proc.stdout).text();
	const exit = await proc.exited;
	if (exit !== 0) return null;
	const first = out.split(/\r?\n/)[0]?.trim();
	return first ? first : null;
}

function basenameWithoutExt(path: string): string {
	const last = path.split(/[\\/]/).pop() ?? path;
	const dot = last.lastIndexOf(".");
	return dot === -1 ? last : last.slice(0, dot);
}
