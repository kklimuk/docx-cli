import { mkdirSync, rmSync } from "node:fs";
import { basename, extname, join } from "node:path";
import { detectEngine } from "./detect";
import {
	type RenderEngine,
	RenderEngineError,
	type RenderEngineName,
} from "./engines/types";
import { splitPdf } from "./split-pdf";

export { detectEngine, engineByName, listAvailable } from "./detect";
export { libreofficeEngine } from "./engines/libreoffice";
export {
	type RenderEngine,
	RenderEngineError,
	type RenderEngineErrorCode,
	type RenderEngineName,
} from "./engines/types";
export { wordMacEngine } from "./engines/word-mac";
export { wordWindowsEngine } from "./engines/word-windows";
export { type SplitOptions, splitPdf } from "./split-pdf";

/** Options accepted by `renderDocxPages`. Mirrors the CLI surface in
 * `src/cli/render/index.ts`: the CLI parses flags into this shape and
 * hands it off. */
export type RenderOptions = {
	/** Directory to write per-page images into. Created (recursively) if
	 *  missing. The intermediate PDF lives here too during the render and is
	 *  removed before this function returns. */
	outDir: string;
	/** Engine to use for docx → PDF. When omitted, `renderDocxPages`
	 *  auto-detects via `detectEngine()`. */
	engine?: RenderEngine;
	/** Pixels per inch for the per-page rasterization. */
	dpi: number;
	/** Output image format. */
	format: "png" | "jpg";
	/** 1-indexed page range, inclusive on both ends. Omit for all pages. */
	range?: { first: number; last: number };
};

export type RenderResult = {
	engine: RenderEngineName;
	outDir: string;
	pages: string[];
};

/** Two-stage docx → per-page-images pipeline:
 *
 * 1. **docx → PDF** via the supplied or auto-detected `RenderEngine`.
 *    Microsoft Word (macOS / Windows) and LibreOffice (cross-platform) are
 *    the two engine families. See `engines/` for the implementations.
 * 2. **PDF → PNG/JPG** via `splitPdf` (PDFium-as-WASM, bundled — no system
 *    tools required).
 *
 * Throws `RenderEngineError` for engine-domain failures (engine not
 * installed, conversion exit code non-zero, PDFium WASM init failure,
 * --pages range out of bounds). The intermediate PDF is removed even if
 * the split throws. */
export async function renderDocxPages(
	docxPath: string,
	options: RenderOptions,
): Promise<RenderResult> {
	const engine = options.engine ?? (await detectEngine());
	if (!engine) {
		throw new RenderEngineError(
			"RENDER_ENGINE",
			"No render engine detected on this machine",
			"Install Microsoft Word (macOS / Windows) or LibreOffice (cross-platform). PDF rasterization is bundled — no other tools needed.",
		);
	}

	try {
		mkdirSync(options.outDir, { recursive: true });
	} catch (error) {
		// Surface mkdir failures (typically ENOTDIR: --out pointed at an
		// existing file, not a directory) as a structured error rather than
		// a raw stack trace.
		throw new RenderEngineError(
			"RENDER_FAILED",
			`Could not create outDir ${options.outDir}: ${
				error instanceof Error ? error.message : String(error)
			}`,
		);
	}

	const pdfPath = join(
		options.outDir,
		`${basename(docxPath, extname(docxPath))}.pdf`,
	);

	let pages: string[];
	try {
		await engine.convertToPdf(docxPath, pdfPath);
		pages = await splitPdf(pdfPath, options.outDir, {
			dpi: options.dpi,
			format: options.format,
			range: options.range,
		});
	} finally {
		rmSync(pdfPath, { force: true });
	}

	return { engine: engine.name, outDir: options.outDir, pages };
}
