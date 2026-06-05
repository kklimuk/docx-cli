/** Shared interface every render engine implements. The CLI picks one via
 * `detect.ts` (or an explicit `--engine` flag) and asks it to convert a docx
 * to a PDF; the PDF is then split into per-page PNGs by `split-pdf.ts`.
 *
 * Engines are platform-specific (Word for Mac via AppleScript, Word for
 * Windows via PowerShell COM, LibreOffice via `soffice --headless`) but
 * agree on this shape so the dispatch in `index.ts` stays generic. */
export type RenderEngineName = "word-mac" | "word-win" | "libreoffice";

export type RenderEngine = {
	readonly name: RenderEngineName;
	/** Does this engine work on the current machine? Should be a cheap
	 * lookup (typically `command -v` or a registry probe — no actual doc
	 * conversion). Used both by `detect.ts` for auto-selection and by the
	 * CLI when the user passes `--engine NAME` explicitly. */
	available(): Promise<boolean>;
	/** Convert `inputDocx` to `outputPdf` (absolute paths). Throws
	 * `RenderEngineError` if the conversion fails. The engine is responsible
	 * for any platform-specific staging (e.g., Word for Mac's sandboxed
	 * Container Documents dir) and for cleaning up after itself. */
	convertToPdf(inputDocx: string, outputPdf: string): Promise<void>;
};

/** Domain error raised by a render engine. `code` maps to the CLI's
 * `ErrorCode` union in `src/cli/respond.ts` so callers can `return
 * fail(err.code, err.message, err.hint)` directly. */
export type RenderEngineErrorCode = "RENDER_ENGINE" | "RENDER_FAILED";

export class RenderEngineError extends Error {
	constructor(
		public code: RenderEngineErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "RenderEngineError";
	}
}
