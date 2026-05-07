import {
	type BlockReference,
	type DocView,
	LocatorResolveError,
	openDocView,
	PkgError,
	resolveBlock,
} from "@core";

export const EXIT = {
	OK: 0,
	GENERAL_ERROR: 1,
	USAGE_ERROR: 2,
	NOT_FOUND: 3,
} as const;

export type ErrorCode =
	| "USAGE"
	| "FILE_NOT_FOUND"
	| "PART_NOT_FOUND"
	| "NOT_A_ZIP"
	| "INVALID_LOCATOR"
	| "BLOCK_NOT_FOUND"
	| "COMMENT_NOT_FOUND"
	| "IMAGE_NOT_FOUND"
	| "HYPERLINK_NOT_FOUND"
	| "TRACKED_CHANGE_NOT_FOUND"
	| "MATCH_NOT_FOUND"
	| "TRACKED_CHANGE_CONFLICT"
	| "UNHANDLED";

export async function respond(payload: unknown): Promise<void> {
	await Bun.write(Bun.stdout, `${JSON.stringify(payload)}\n`);
}

let verboseAck = false;

/** Switch on full JSON acks for the current process. Mutating commands call
 *  this when they parse `--verbose`/`-v`. Errors always print regardless;
 *  dry-run payloads always print regardless. */
export function setVerboseAck(verbose: boolean): void {
	verboseAck = verbose;
}

/** Mutating-command success ack: prints the JSON payload only when
 *  `--verbose` is set. By default mutators are silent on success — agents
 *  rely on exit code 0 + the absence of an error payload. */
export async function respondAck(payload: unknown): Promise<void> {
	if (!verboseAck) return;
	await respond(payload);
}

export async function writeStdout(text: string): Promise<void> {
	await Bun.write(Bun.stdout, text);
}

export async function fail(
	code: ErrorCode,
	message: string,
	hint?: string,
): Promise<number> {
	const payload: { ok: false; code: ErrorCode; error: string; hint?: string } =
		{
			ok: false,
			code,
			error: message,
		};
	if (hint) payload.hint = hint;
	await respond(payload);
	return exitCodeFor(code);
}

function exitCodeFor(code: ErrorCode): number {
	switch (code) {
		case "USAGE":
		case "INVALID_LOCATOR":
			return EXIT.USAGE_ERROR;
		case "FILE_NOT_FOUND":
		case "PART_NOT_FOUND":
		case "BLOCK_NOT_FOUND":
		case "COMMENT_NOT_FOUND":
		case "IMAGE_NOT_FOUND":
		case "HYPERLINK_NOT_FOUND":
		case "TRACKED_CHANGE_NOT_FOUND":
		case "MATCH_NOT_FOUND":
			return EXIT.NOT_FOUND;
		case "NOT_A_ZIP":
		case "TRACKED_CHANGE_CONFLICT":
		case "UNHANDLED":
			return EXIT.GENERAL_ERROR;
	}
}

export async function openOrFail(path: string): Promise<DocView | number> {
	try {
		return await openDocView(path);
	} catch (err) {
		if (err instanceof PkgError) {
			if (err.code === "FILE_NOT_FOUND") {
				return await fail("FILE_NOT_FOUND", err.message);
			}
			if (err.code === "NOT_A_ZIP") return await fail("NOT_A_ZIP", err.message);
		}
		throw err;
	}
}

export async function resolveBlockOrFail(
	view: DocView,
	locator: string,
): Promise<BlockReference | number> {
	try {
		return resolveBlock(view, locator);
	} catch (err) {
		if (err instanceof LocatorResolveError) {
			return await fail("BLOCK_NOT_FOUND", err.message);
		}
		throw err;
	}
}
