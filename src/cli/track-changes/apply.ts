import { TrackChanges, TrackedChangeNotFoundError } from "@core/track-changes";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import {
	expandRevisionTargets,
	revisionGroups,
	UnknownRevisionError,
} from "./groups";

export type ApplyVerb = "accept" | "reject";

export async function runApply(
	args: string[],
	verb: ApplyVerb,
	help: string,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string", multiple: true },
			all: { type: "boolean" },
			...SAVE_FLAGS,
		},
		help,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const atRaw = parsed.values.at as string[] | undefined;
	const all = Boolean(parsed.values.all);
	if (atRaw && atRaw.length > 0 && all) {
		return fail("USAGE", "--at and --all are mutually exclusive", help);
	}
	if ((!atRaw || atRaw.length === 0) && !all) {
		return fail("USAGE", "Specify --at tcN (repeatable) or --all", help);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const trackChanges = new TrackChanges(document);
	// `--at revN` addresses a del+ins replace pair (from `list`'s `group` field) and
	// expands to both member tcNs, so one logical change is one call — no re-list
	// between halves. tcN/--all are untouched.
	let target: "all" | string[];
	if (all) {
		target = "all";
	} else {
		try {
			target = expandRevisionTargets(
				atRaw ?? [],
				revisionGroups(trackChanges.list()),
			);
		} catch (error) {
			if (error instanceof UnknownRevisionError) {
				return fail(
					"TRACKED_CHANGE_NOT_FOUND",
					error.message,
					"Run 'docx track-changes list FILE' — revN handles appear as the `group` field on paired changes.",
				);
			}
			throw error;
		}
	}
	const outputPath = parsed.values.output as string | undefined;

	try {
		if (parsed.values["dry-run"]) {
			await respond({
				operation: `track-changes.${verb}`,
				dryRun: true,
				path,
				...(outputPath ? { output: outputPath } : {}),
				applied: trackChanges.preview(target, verb),
			});
			return EXIT.OK;
		}

		const applied =
			verb === "accept"
				? trackChanges.accept(target)
				: trackChanges.reject(target);

		await document.save(outputPath);

		await respondAck({
			ok: true,
			operation: `track-changes.${verb}`,
			path: outputPath ?? path,
			applied,
		});
		return EXIT.OK;
	} catch (error) {
		if (error instanceof TrackedChangeNotFoundError) {
			return fail(
				"TRACKED_CHANGE_NOT_FOUND",
				error.message,
				"Run 'docx track-changes list FILE' to see available ids.",
			);
		}
		throw error;
	}
}
