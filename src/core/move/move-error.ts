/** Why a move was refused. `code` is a subset of the CLI's error codes so the
 *  CLI can pass it straight through; `entry` names the batch line. */
export class MoveError extends Error {
	constructor(
		public code: MoveErrorCode,
		message: string,
		public hint?: string,
		public entry?: number,
	) {
		super(message);
		this.name = "MoveError";
	}
}

export type MoveErrorCode =
	| "USAGE"
	| "INVALID_LOCATOR"
	| "BLOCK_NOT_FOUND"
	| "TABLE_STRUCTURE"
	| "UNSUPPORTED";
