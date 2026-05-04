export const EXIT = {
	OK: 0,
	GENERAL_ERROR: 1,
	USAGE_ERROR: 2,
	NOT_FOUND: 3,
	PERMISSION_DENIED: 4,
	ALREADY_APPLIED: 5,
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
	| "PERMISSION_DENIED"
	| "ALREADY_APPLIED"
	| "UNHANDLED";

export async function respond(payload: unknown): Promise<void> {
	await Bun.write(Bun.stdout, `${JSON.stringify(payload)}\n`);
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
			return EXIT.NOT_FOUND;
		case "PERMISSION_DENIED":
			return EXIT.PERMISSION_DENIED;
		case "ALREADY_APPLIED":
			return EXIT.ALREADY_APPLIED;
		case "NOT_A_ZIP":
		case "UNHANDLED":
			return EXIT.GENERAL_ERROR;
	}
}
