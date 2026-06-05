import { join, resolve } from "node:path";
import { PDFiumLibrary } from "@hyzyla/pdfium";
// Embed PDFium's WASM as a build-time asset via Bun's `with { type: "file" }`
// import attribute. Under `bun build --compile` (the build:binary path,
// which produces the standalone `dist/docx` binary we actually ship) the
// .wasm file gets baked into the executable and resolves at runtime to a
// `/$bunfs/root/...` virtual path — `Bun.file(path).arrayBuffer()` reads
// it transparently. Under plain `bun src/index.ts` / `bun dist/index.js`
// the import resolves to the on-disk node_modules path. Both modes pass
// the bytes into `PDFiumLibrary.init({ wasmBinary })` so we never hit the
// default entry's `fs.readFile("pdfium.wasm")` path that breaks under
// --compile. The /browser/base64 entry is the obvious alternative, but
// its Emscripten module is compiled with `ENVIRONMENT=web` and panics
// when imported from Node/Bun — verified empirically.
import pdfiumWasmPath from "@hyzyla/pdfium/dist/pdfium.wasm" with {
	type: "file",
};
import jpegJs from "jpeg-js";
import { PNG } from "pngjs";
import { RenderEngineError } from "./engines/types";

export type SplitOptions = {
	/** Pixels per inch. Word's default print render is 72 dpi; 150 is the
	 *  sweet spot for verification screenshots (legible at 100% zoom in any
	 *  modern image viewer without ballooning file size). */
	dpi: number;
	/** Image format. PNG by default — lossless, smaller for the text-heavy
	 *  pages docx render produces. JPG goes through `jpeg-js` at quality 85
	 *  (the standard look-indistinguishable point). */
	format: "png" | "jpg";
	/** 1-indexed page range. Inclusive on both ends. Omit for all pages. */
	range?: { first: number; last: number };
};

/** Split a PDF into per-page images via the bundled `@hyzyla/pdfium`
 * (PDFium-as-WASM, MIT wrapper + BSD-3-Clause / Apache-2.0 PDFium binary).
 * The .wasm file is embedded as a build-time asset via
 * `with { type: "file" }` — both the `bun build` bundle (sibling to
 * `dist/index.js`) and the `bun build --compile` standalone binary
 * (embedded inside `dist/docx`) ship the WASM, so no system tools are
 * required at runtime.
 *
 * Output files land at `<outDir>/page-NNN.<ext>` (3-digit zero-padded for
 * natural sort up to 999 pages). PDFium itself is the same engine driving
 * Chrome's PDF viewer — output quality matches what the user sees in any
 * modern PDF reader. */
export async function splitPdf(
	pdfPath: string,
	outDir: string,
	options: SplitOptions,
): Promise<string[]> {
	let lib: Awaited<ReturnType<typeof PDFiumLibrary.init>>;
	try {
		// `pdfiumWasmPath` can be:
		//   • absolute under dev (the node_modules path)
		//   • absolute under `bun build --compile` (a `/$bunfs/root/...`
		//     virtual path inside the embedded bundle)
		//   • RELATIVE under `bun build` (the asset filename next to the
		//     bundle, e.g. `./pdfium-jj0zq23w.wasm`) — that resolves
		//     against the CWD by default, which breaks any invocation from
		//     outside the dist dir. `path.resolve(scriptDir, p)` returns
		//     `p` unchanged when it's absolute and joins it with the
		//     bundle's directory when it's relative — the right answer in
		//     all three modes.
		const wasmFullPath = resolve(import.meta.dir, pdfiumWasmPath);
		const wasmBinary = await Bun.file(wasmFullPath).arrayBuffer();
		lib = await PDFiumLibrary.init({ wasmBinary });
	} catch (error) {
		throw asRenderError(error, "Failed to initialize PDFium WASM");
	}
	try {
		const bytes = await Bun.file(pdfPath).bytes();
		let document: Awaited<ReturnType<typeof lib.loadDocument>>;
		try {
			document = await lib.loadDocument(bytes);
		} catch (error) {
			throw asRenderError(error, `Failed to open PDF "${pdfPath}"`);
		}
		try {
			const totalPages = document.getPageCount();
			const first = options.range?.first ?? 1;
			const last = options.range?.last ?? totalPages;
			// Validate up front so the user gets a coherent error instead of
			// "requested 5-4 of 4" when they ask for `--pages 5` on a 4-page
			// doc, or a similarly confusing clamp on `--pages 100-200`.
			if (first > totalPages) {
				throw new RenderEngineError(
					"RENDER_FAILED",
					`--pages first (${first}) exceeds page count (${totalPages})`,
				);
			}
			if (last > totalPages) {
				throw new RenderEngineError(
					"RENDER_FAILED",
					`--pages last (${last}) exceeds page count (${totalPages})`,
				);
			}
			// PDFium's renderer takes a `scale` factor relative to the PDF's
			// native 72 dpi. `scale = dpi / 72` produces a bitmap at the
			// caller's chosen DPI.
			const scale = options.dpi / 72;
			const written: string[] = [];
			// Use `getPage(i)` rather than the `pages()` generator: the
			// generator opens every page sequentially before yielding, and
			// `_FPDF_ClosePage` only fires inside `page.render()`'s finally —
			// so a `--pages 100-101` on a 1000-page doc would leak 99 page
			// handles. `getPage` opens just the ones we need.
			for (let pageIndex = first; pageIndex <= last; pageIndex++) {
				const page = document.getPage(pageIndex - 1);
				let rendered: { data: Uint8Array };
				try {
					rendered = await page.render({
						scale,
						// PDFium with the default BGRA colorspace sets
						// REVERSE_BYTE_ORDER internally, so the callback
						// receives the buffer in **RGBA** order — confirmed
						// by inspecting node_modules/@hyzyla/pdfium/dist/
						// index.esm.js. pngjs + jpeg-js both expect RGBA;
						// no channel swap needed. PDFium also internally
						// deep-copies the WASM heap data (see the `slice()`
						// comment in the upstream source), so the
						// `data: Uint8Array` we receive is already owned —
						// pass it through to the encoders without an extra
						// `Buffer.from()` copy (each one allocates ~140 MB
						// at 600 DPI A4).
						render: async ({ data, width, height }) =>
							encode(options.format, data, width, height),
					});
				} catch (error) {
					throw asRenderError(
						error,
						`Failed to render page ${pageIndex} of ${totalPages}`,
					);
				}
				const padded = String(pageIndex).padStart(3, "0");
				const outPath = join(outDir, `page-${padded}.${options.format}`);
				await Bun.write(outPath, rendered.data);
				written.push(outPath);
			}
			return written;
		} finally {
			document.destroy();
		}
	} finally {
		lib.destroy();
	}
}

function encode(
	format: "png" | "jpg",
	data: Uint8Array,
	width: number,
	height: number,
): Uint8Array {
	if (format === "png") {
		const png = new PNG({ width, height });
		png.data = data as unknown as Buffer;
		return PNG.sync.write(png);
	}
	return jpegJs.encode({ data, width, height }, 85).data;
}

/** Wrap an unknown engine/encoder failure as a structured `RenderEngineError`
 * so the CLI top-level catch can produce a clean JSON ack instead of leaking
 * a raw stack trace to agents consuming the output. */
function asRenderError(error: unknown, prefix: string): RenderEngineError {
	if (error instanceof RenderEngineError) return error;
	const message = error instanceof Error ? error.message : String(error);
	return new RenderEngineError("RENDER_FAILED", `${prefix}: ${message}`);
}
