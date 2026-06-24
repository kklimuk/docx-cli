import { type MarginalKind, Marginals } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	renderVerifyHint,
	resolveTracked,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import {
	isTypeError,
	marginalNoun,
	resolveMarginalType,
	resolveTargetSectPrs,
} from "./shared";

const OPTION_SPEC = {
	at: { type: "string" },
	type: { type: "string" },
	"first-page": { type: "boolean" },
	even: { type: "boolean" },
	odd: { type: "boolean" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

function helpFor(kind: MarginalKind): string {
	const noun = marginalNoun(kind);
	return `docx ${noun} clear — remove a ${kind} reference

Usage:
  docx ${noun} clear FILE [--at sN] [--type T | --first-page | --even | --odd] [options]

Removes the ${kind} reference of the given type from one section (--at sN) or every
section (default). The ${kind} part itself is left in the file as a harmless orphan
(re-setting mints a fresh part). Idempotent — clearing an absent ${kind} is a no-op.

Options:
  --at sN            Target one section (default: the whole document)
  --type T           default | first | even (default: default)
  --first-page       ≡ --type first    --even ≡ --type even    --odd ≡ --type default
  --track            Record the removal as a tracked <w:sectPrChange>
  --author NAME      Tracked-change author
  -o, --output PATH  Write to PATH instead of editing FILE in place
  --dry-run          Preview; write nothing
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help

Examples:
  docx ${noun} clear doc.docx
  docx ${noun} clear doc.docx --at s0 --first-page
`;
}

export async function runClearMarginal(
	args: string[],
	kind: MarginalKind,
): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, helpFor(kind));
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(helpFor(kind));
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", helpFor(kind));

	const type = resolveMarginalType(parsed.values);
	if (isTypeError(type)) return fail("USAGE", type.error, type.hint);

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const atLocator = parsed.values.at as string | undefined;
	const targets = await resolveTargetSectPrs(document, atLocator);
	if (typeof targets === "number") return targets;

	const tracked = resolveTracked(document, parsed.values.track);
	const result = new Marginals(document).clear(targets, kind, type, {
		track: tracked,
		authorFlag: parsed.values.author as string | undefined,
	});

	const noun = marginalNoun(kind);
	const dryRun = Boolean(parsed.values["dry-run"]);
	const outputPath = parsed.values.output as string | undefined;
	if (dryRun) {
		await respond({
			operation: `${noun}.clear`,
			dryRun: true,
			path: filePath,
			kind,
			type,
			removed: result.removed,
			...(atLocator ? { locator: atLocator } : {}),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	await document.save(outputPath);
	const destination = outputPath ?? filePath;
	await respondAck(
		{
			ok: true,
			operation: `${noun}.clear`,
			path: destination,
			kind,
			type,
			removed: result.removed,
			...(atLocator ? { locator: atLocator } : { applied: result.removed }),
		},
		renderVerifyHint(destination),
	);
	return EXIT.OK;
}
