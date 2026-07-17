/** The one error type every raw gate throws. A leaf module so the gate files
 *  (parse, element-order, namespaces, references, parts, relationships) can
 *  share it without importing the `Raw` composition root back — index.ts
 *  re-exports it for external consumers. */
export type RawErrorCode = "INVALID_XML" | "USAGE" | "BLOCK_NOT_FOUND";

export class RawError extends Error {
	code: RawErrorCode;
	hint?: string;

	constructor(code: RawErrorCode, message: string, hint?: string) {
		super(message);
		this.code = code;
		this.hint = hint;
	}
}
