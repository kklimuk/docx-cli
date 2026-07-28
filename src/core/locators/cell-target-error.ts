/** Why a cell target was refused. Its own module because BOTH sides of the
 * cell-targeting stack raise it — `locators/resolve.ts` when a locator names a
 * cell we won't mutate bare, and `table/cell-content.ts` when a cell's insertion
 * boundary goes stale — and `resolve.ts` already value-imports `cell-content.ts`.
 * Homing the class in either one would close that import cycle. */
export class CellTargetError extends Error {
	constructor(
		public code: CellTargetErrorCode,
		message: string,
		public hint?: string,
	) {
		super(message);
		this.name = "CellTargetError";
	}
}

export type CellTargetErrorCode =
	| "INVALID_LOCATOR"
	| "BLOCK_NOT_FOUND"
	| "TABLE_STRUCTURE";
