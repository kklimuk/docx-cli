import { rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { type RenderEngine, RenderEngineError } from "./types";

/** Word for Windows engine. Drives Microsoft Word via PowerShell's COM
 * automation. Word's `Document.ExportAsFixedFormat($pdfPath, 17)` (where
 * `17` is `wdExportFormatPDF`) writes a PDF at the requested location with
 * no UI prompts.
 *
 * No sandbox staging is needed here — Word for Windows opens arbitrary
 * paths directly. We do route the input through a temp dir to normalize
 * pathing (PowerShell COM gets unhappy about UNC paths and certain
 * trailing whitespace). */
export const wordWindowsEngine: RenderEngine = {
	name: "word-win",
	async available(): Promise<boolean> {
		if (process.platform !== "win32") return false;
		// Probe via PowerShell — if Word's COM class is registered, it'll
		// instantiate cheaply, then we release. This avoids guessing install
		// paths across Office versions / 32-vs-64-bit editions.
		const script =
			"try { $w = New-Object -ComObject Word.Application -ErrorAction Stop; $w.Quit(); $true } catch { $false }";
		const proc = Bun.spawn(["powershell", "-NoProfile", "-Command", script], {
			stdout: "pipe",
			stderr: "pipe",
		});
		const out = (await new Response(proc.stdout).text()).trim();
		const exit = await proc.exited;
		return exit === 0 && out.toLowerCase() === "true";
	},
	async convertToPdf(inputDocx: string, outputPdf: string): Promise<void> {
		// PowerShell single-quoting is the safest path: ' is escaped as '' inside
		// a single-quoted string. The input path itself is sanitized here.
		const quote = (path: string): string => path.replace(/'/g, "''");
		// Stage input + output in tmpdir so we don't surprise the user with a
		// COM-permission denial on their working directory.
		const stage = join(tmpdir(), `docx-cli-render-${process.pid}-${stamp()}`);
		const stagedDocx = `${stage}.docx`;
		const stagedPdf = `${stage}.pdf`;
		try {
			await Bun.write(stagedDocx, Bun.file(inputDocx));
			const script = [
				"$ErrorActionPreference = 'Stop'",
				"$word = New-Object -ComObject Word.Application",
				"$word.Visible = $false",
				"try {",
				`	$doc = $word.Documents.Open('${quote(stagedDocx)}', $false, $true)`,
				// 17 = wdExportFormatPDF (Word's enum constant).
				`	$doc.ExportAsFixedFormat('${quote(stagedPdf)}', 17)`,
				"	$doc.Close([ref]0)",
				"} finally {",
				"	$word.Quit()",
				"	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null",
				"}",
			].join("\n");
			const proc = Bun.spawn(["powershell", "-NoProfile", "-Command", script], {
				stdout: "pipe",
				stderr: "pipe",
			});
			const exit = await proc.exited;
			if (exit !== 0) {
				const stderr = await new Response(proc.stderr).text();
				throw new RenderEngineError(
					"RENDER_FAILED",
					`Word for Windows COM failed (exit ${exit}): ${stderr.trim() || "(no stderr)"}`,
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

let monotonic = 0;
function stamp(): string {
	monotonic += 1;
	return String(monotonic);
}
