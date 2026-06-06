import { TrackChanges, TrackedChangeNotFoundError } from "@core/track-changes";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

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
			output: { type: "string", short: "o" },
			"dry-run": { type: "boolean" },
			verbose: { type: "boolean", short: "v" },
			help: { type: "boolean", short: "h" },
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

	const target = all ? "all" : (atRaw ?? []);
	const trackChanges = new TrackChanges(document);
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
