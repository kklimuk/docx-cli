/** Domain error raised by `MarkdownImport`. The `code` literal is a strict
 * subset of the CLI's `ErrorCode` union (`src/cli/respond.ts`) so callers can
 * `return fail(err.code, err.message, err.hint)` without a cast. */
export type MarkdownImportErrorCode =
	| "USAGE"
	| "IMAGE_SOURCE"
	| "TRACKED_CHANGE_CONFLICT";

export class MarkdownImportError extends Error {
	constructor(
		public code: MarkdownImportErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "MarkdownImportError";
	}
}
