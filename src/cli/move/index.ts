import { Move, MoveError, type MoveRequest, type PlannedMove } from "@core";
import { batchExampleIntro } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondMinted,
	SAVE_FLAGS,
	STALE_LOCATOR_HINT,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { pickDestination, readMoveBatch } from "./batch";

const HELP = `docx move — relocate existing blocks, untouched

Moves the blocks themselves (no copy, no Markdown round-trip), so everything
in them comes along: formatting, comments, footnotes, bookmarks, fields,
images, text boxes. Prefer this over insert-a-copy + delete-the-original.

Usage:
  docx move FILE --at LOCATOR (--after | --before) LOCATOR [options]
  docx move FILE --at pN --subtree (--after | --before) LOCATOR [options]
  docx move FILE --batch FILE.jsonl [options]   # many moves, one read (- = stdin)

Examples:
  docx move doc.docx --at p7 --before p2              # one paragraph up
  docx move doc.docx --at p4-p9 --after p20           # a run of blocks
  docx move doc.docx --at p12 --subtree --before p3   # a heading + its section
  docx move doc.docx --at t1 --after p0               # a table
  docx move doc.docx --at p5 --after t0:r1c2          # into a table cell (at its end)
${batchExampleIntro("Reorder several things")}
  #   moves.jsonl:
  #     {"at":"p3","after":"p9"}
  #     {"at":"p12","subtree":true,"before":"p0"}
  docx move doc.docx --batch moves.jsonl

What moves (--at):
  pN            a paragraph          pN-pM   paragraphs pN..pM and all between
  tN            a table              tN:rRcC:pK / tbxN:pK   a cell / text-box paragraph
  --subtree     with a heading pN: the heading plus everything under it, up to
                the next heading of the same or higher level (see \`docx outline\`)

Where it goes (exactly one):
  --after LOCATOR   after a block, or at the END of a bare cell (tN:rRcC)
  --before LOCATOR  before a block, or at the START of a bare cell

Refused (nonzero exit, nothing written): a destination inside the moved
blocks; a move that lands where the blocks already are; a paragraph carrying a
section break; a comment or bookmark with one end inside the moved blocks and
one outside; a move between the body and a text box; any move while track
changes is on (tracked moves aren't supported yet).

Batch (--batch PATH | -): one {"at", "after"|"before", "subtree"?} per line.
  Every locator addresses the document AS READ; moves apply in file order, and
  several moved after the SAME block land in that order. Entries may not move
  overlapping blocks or anchor on a block another entry moves.

Options:
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Check the move(s); write nothing
  -v, --verbose     Full JSON ack
  -h, --help        Show this help

Output: the moved blocks' NEW locators, one per line (exit 0) — the old ones
are stale after a move. Errors print {code, error, hint?} with a nonzero exit.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	after: { type: "string" },
	before: { type: "string" },
	subtree: { type: "boolean" },
	batch: { type: "string" },
	...SAVE_FLAGS,
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}
	setVerboseAck(Boolean(parsed.values.verbose));

	const [filePath, extra] = parsed.positionals;
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);
	if (extra !== undefined) {
		return fail(
			"USAGE",
			`Unexpected argument: ${extra}`,
			"move takes no content — it relocates existing blocks. Pass --at LOCATOR and --after/--before LOCATOR.",
		);
	}

	const batchSource = parsed.values.batch as string | undefined;
	const requests =
		batchSource !== undefined
			? await readMoveBatch(batchSource, parsed.values)
			: await parseSingleRequest(parsed.values);
	if (typeof requests === "number") return requests;

	return commitMoves(filePath, requests, {
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	});
}

async function parseSingleRequest(
	values: Record<string, unknown>,
): Promise<MoveRequest[] | number> {
	const at = values.at as string | undefined;
	if (!at) return fail("USAGE", "Missing --at LOCATOR (or --batch FILE)", HELP);
	const destination = pickDestination(
		values.after as string | undefined,
		values.before as string | undefined,
	);
	if (typeof destination === "string") {
		return fail("USAGE", destination, HELP);
	}
	return [{ at, ...destination, subtree: Boolean(values.subtree) }];
}

async function commitMoves(
	filePath: string,
	requests: MoveRequest[],
	opts: { outputPath?: string; dryRun: boolean },
): Promise<number> {
	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	if (document.isTrackChangesEnabled()) {
		return fail(
			"UNSUPPORTED",
			"Track changes is on, and tracked moves aren't supported yet — nothing was moved",
			"To record it as a tracked change, insert the content at the destination and delete the original (both redline). To move it untracked, run `docx track-changes off`, move, then `docx track-changes on`.",
		);
	}

	const lens = new Move(document);
	let planned: PlannedMove[];
	try {
		planned = lens.plan(requests);
	} catch (error) {
		if (!(error instanceof MoveError)) throw error;
		const prefix = error.entry !== undefined ? `entry ${error.entry}: ` : "";
		const hint =
			error.hint ??
			(error.code === "BLOCK_NOT_FOUND" ? STALE_LOCATOR_HINT : undefined);
		return fail(error.code, `${prefix}${error.message}`, hint);
	}

	const moves = requests.map(({ at, side, to, subtree }) => ({
		at,
		[side]: to,
		...(subtree ? { subtree: true } : {}),
	}));
	if (opts.dryRun) {
		await respond({
			operation: "move",
			dryRun: true,
			path: filePath,
			moves,
			...(opts.outputPath ? { output: opts.outputPath } : {}),
		});
		return EXIT.OK;
	}

	const placed = lens.apply(planned);
	await document.save(opts.outputPath);
	document.reread();
	const locators = document.body.locatorsOf(new Set(placed));
	await respondMinted(locators, {
		ok: true,
		operation: "move",
		path: opts.outputPath ?? filePath,
		locators,
		moves,
	});
	return EXIT.OK;
}
