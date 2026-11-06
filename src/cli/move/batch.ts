import type { MoveRequest } from "@core";
import { readJsonlObjects } from "../parse-helpers";
import { fail } from "../respond";

/** `docx move --batch FILE.jsonl`: one `{ at, after|before, subtree? }` per
 *  line. Only the SHAPE is checked here; the Move lens resolves every locator
 *  against the document as read before anything moves. */
export async function readMoveBatch(
	source: string,
	values: Record<string, unknown>,
): Promise<MoveRequest[] | number> {
	const singleFlags = BATCH_KEYS.filter((flag) => values[flag] !== undefined);
	if (singleFlags.length > 0) {
		return fail(
			"USAGE",
			`--batch reads every move from the JSONL file; don't also pass --${singleFlags.join(", --")}`,
			'Put one {"at": LOCATOR, "after"|"before": LOCATOR} per line.',
		);
	}
	let entries: Record<string, unknown>[];
	try {
		entries = await readJsonlObjects(source);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (entries.length === 0) return fail("USAGE", "Batch file is empty");

	const requests: MoveRequest[] = [];
	for (const [index, entry] of entries.entries()) {
		const { at, after, before, subtree } = entry;
		const unknownKeys = Object.keys(entry).filter(
			(key) => !BATCH_KEYS.includes(key),
		);
		if (unknownKeys.length > 0) {
			return fail(
				"USAGE",
				`entry ${index}: unknown key ${unknownKeys.map((key) => `"${key}"`).join(", ")}`,
				'Each line is {"at": LOCATOR, "after"|"before": LOCATOR, "subtree"?: true}.',
			);
		}
		if (typeof at !== "string" || at.length === 0) {
			return fail("USAGE", `entry ${index}: "at" is required`);
		}
		if (
			(after !== undefined && typeof after !== "string") ||
			(before !== undefined && typeof before !== "string") ||
			(subtree !== undefined && typeof subtree !== "boolean")
		) {
			return fail(
				"USAGE",
				`entry ${index}: "after"/"before" must be locator strings and "subtree" a boolean`,
			);
		}
		const destination = pickDestination(after, before);
		if (typeof destination === "string") {
			return fail(
				"USAGE",
				`entry ${index}: ${destination.replace(/--(after|before)/g, '"$1"')}`,
			);
		}
		requests.push({ at, ...destination, subtree: subtree ?? false });
	}
	return requests;
}

/** The keys a batch line may carry — the single-shot flags, one to one. */
const BATCH_KEYS: readonly string[] = ["at", "after", "before", "subtree"];

/** Exactly one of `--after` / `--before`; returns the error message otherwise.
 *  Shared with the batch reader so both surfaces word it the same. */
export function pickDestination(
	after: string | undefined,
	before: string | undefined,
): Pick<MoveRequest, "to" | "side"> | string {
	if (after !== undefined && before !== undefined) {
		return "Pass exactly one of --after or --before, not both";
	}
	if (after) return { to: after, side: "after" };
	if (before) return { to: before, side: "before" };
	return "Missing destination: pass --after LOCATOR or --before LOCATOR";
}
