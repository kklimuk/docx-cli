import { Comments, CommentsError, resolveAuthor } from "@core";
import {
	batchExampleIntro,
	decodeInlineEscapes,
	readJsonlObjects,
} from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondMinted,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx comments reply — reply to an existing comment

Usage:
  docx comments reply FILE --at cN --text TEXT [options]
  docx comments reply FILE --batch FILE.jsonl [options]   # many replies, one call (- = stdin)

Examples:
${batchExampleIntro("Answer every comment")}
  #   replies.jsonl:
  #     {"at":"c0","text":"Agreed, narrowed to project work.","resolve":true}
  #     {"at":"c2","text":"Keeping as drafted; see section 4."}
  docx comments reply doc.docx --batch replies.jsonl
  # …or one at a time:
  docx comments reply doc.docx --at c0 --text "Good catch" --author "Reviewer"

Required:
  --at cN           Parent comment id (e.g., c0). The "c" prefix is optional.
                    Replying to a reply attaches to the thread root (Word
                    threads are single-level); the ack's parentId reports it.
  --text TEXT       Reply body
  --batch PATH      JSONL, one {"at":"cN","text":"…"} per line (optional
                    "author", and "resolve":true to also mark that thread
                    resolved; "resolved" is accepted too). Every id addresses
                    the comments AS READ; all entries are validated (unknown
                    fields rejected) before anything is written.

Optional:
  --author NAME     Author name (default: $DOCX_AUTHOR, else "Reviewer")
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Print what would be added; do not write the file
  -v, --verbose     Print the full success ack JSON
  -h, --help        Show this help

Output:
  Prints the new reply's comment id (e.g. c4) on success — one per line for
  --batch. --verbose prints the
  full ack {ok:true, operation, path, commentId, parentId}. Errors print
  {code, error, hint?} with a nonzero exit.
  Discover existing comment ids with \`docx comments list FILE\`.
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			text: { type: "string" },
			author: { type: "string" },
			batch: { type: "string" },
			...SAVE_FLAGS,
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const batchInput = parsed.values.batch as string | undefined;
	if (batchInput !== undefined) {
		if (parsed.values.at !== undefined || parsed.values.text !== undefined) {
			return fail(
				"USAGE",
				"--batch reads each entry's at/text from JSONL — do not pass --at/--text on the CLI",
				HELP,
			);
		}
		let rawEntries: Record<string, unknown>[];
		try {
			rawEntries = await readJsonlObjects(batchInput);
		} catch (error) {
			const message = error instanceof Error ? error.message : String(error);
			return fail("USAGE", `Failed to read batch: ${message}`);
		}
		if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");
		return runReplies(path, rawEntries, parsed.values, { batch: true });
	}

	// A single reply is a one-entry batch through the same pipeline.
	const parentInput = parsed.values.at as string | undefined;
	const text = decodeInlineEscapes(parsed.values.text as string | undefined);
	if (!parentInput) return fail("USAGE", "Missing --at cN", HELP);
	if (!text) return fail("USAGE", "Missing --text TEXT", HELP);
	return runReplies(path, [{ at: parentInput, text }], parsed.values, {
		batch: false,
	});
}

/** Validate every entry, check every parent against the comments AS READ (a bad
 *  id aborts before anything is written), then mint the replies in entry order
 *  and resolve the threads marked `resolve: true` — answering and closing a
 *  review comment is one entry, and a finalize pass is one call. A single
 *  `--at/--text` reply runs as a one-entry batch and keeps its own ack shape
 *  (`commentId` + the EFFECTIVE `parentId`: the thread root, which differs from
 *  `--at` when the target was itself a reply). */
async function runReplies(
	path: string,
	rawEntries: Record<string, unknown>[],
	values: Record<string, unknown>,
	{ batch }: { batch: boolean },
): Promise<number> {
	const label = (index: number) => (batch ? `entry ${index}: ` : "");
	const defaultAuthor = values.author as string | undefined;
	const entries: ReplyEntry[] = [];
	for (const [index, raw] of rawEntries.entries()) {
		const entry = readReplyEntry(raw, label(index), defaultAuthor);
		if (typeof entry === "string") return fail("USAGE", entry);
		entries.push(entry);
	}

	const document = await openOrFail(path);
	if (typeof document === "number") return document;
	for (const [index, entry] of entries.entries()) {
		if (!document.comments?.findById(entry.parentId)) {
			return fail(
				"COMMENT_NOT_FOUND",
				`${label(index)}comment not found: c${entry.parentId}`,
				"Discover comment ids with `docx comments list FILE`. Nothing was written.",
			);
		}
	}
	const outputPath = values.output as string | undefined;
	const threadRoot = (parentId: string) =>
		`c${document.comments?.threadRootId(parentId) ?? parentId}`;
	const toResolve = [
		...new Set(
			entries
				.filter((entry) => entry.resolve)
				.map((entry) => threadRoot(entry.parentId)),
		),
	];
	const ack = (minted: { commentId: string; parentId: string }[]) =>
		batch
			? { batch: minted, resolved: toResolve }
			: { ...(minted[0] as { commentId: string; parentId: string }) };

	if (values["dry-run"]) {
		// Reply ids append in entry order from the next free id, so the preview
		// names the same commentIds the real run will mint.
		const firstId = Number(document.comments?.nextId() ?? "0");
		await respond({
			operation: "comments.reply",
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			...ack(
				entries.map((entry, index) => ({
					commentId: `c${firstId + index}`,
					parentId: threadRoot(entry.parentId),
				})),
			),
		});
		return EXIT.OK;
	}

	const comments = new Comments(document);
	const minted: { commentId: string; parentId: string }[] = [];
	try {
		for (const entry of entries) {
			const numericId = comments.reply(entry.parentId, entry.text, {
				author: entry.author,
			});
			minted.push({
				commentId: `c${numericId}`,
				parentId: threadRoot(entry.parentId),
			});
		}
		if (toResolve.length > 0) comments.resolve(toResolve, true);
	} catch (error) {
		if (error instanceof CommentsError) {
			// Nothing is saved on a mid-batch failure; name the entry that broke.
			const prefix = minted.length < entries.length ? label(minted.length) : "";
			return fail(error.code, `${prefix}${error.message}`, error.hint);
		}
		throw error;
	}

	await document.save(outputPath);
	// A reply is a new comment with an id the agent can't reconstruct, so the
	// ids print by default; --verbose upgrades to the full ack.
	await respondMinted(
		minted.map((entry) => entry.commentId),
		{
			ok: true,
			operation: "comments.reply",
			path: outputPath ?? path,
			...ack(minted),
		},
	);
	return EXIT.OK;
}

/** Validate one JSONL entry into a `ReplyEntry`, or return the USAGE message.
 *  Unknown fields are refused rather than ignored: `{"resolved": true}` (the
 *  spelling `comments list` prints) used to be dropped silently, leaving the
 *  thread open at exit 0 — so that spelling is accepted as an alias, and
 *  anything else unrecognized errors. */
function readReplyEntry(
	raw: Record<string, unknown>,
	label: string,
	defaultAuthor: string | undefined,
): ReplyEntry | string {
	const unknown = Object.keys(raw).find((key) => !REPLY_ENTRY_KEYS.has(key));
	if (unknown !== undefined) {
		return `${label}unknown field "${unknown}" (fields: at, text, author, resolve)`;
	}
	const target = raw.at ?? raw.id;
	if (typeof target !== "string" || target.length === 0) {
		return `${label}missing "at" (the comment id, e.g. "c0")`;
	}
	if (typeof raw.text !== "string" || raw.text.length === 0) {
		return `${label}missing "text" (the reply body)`;
	}
	if (raw.resolve !== undefined && raw.resolved !== undefined) {
		return `${label}"resolve" and "resolved" are the same field — pass it once`;
	}
	const resolve = raw.resolve ?? raw.resolved ?? false;
	if (typeof resolve !== "boolean") {
		return `${label}"resolve" must be true or false, got ${JSON.stringify(resolve)}`;
	}
	const author = typeof raw.author === "string" ? raw.author : defaultAuthor;
	return {
		parentId: target.startsWith("c") ? target.slice(1) : target,
		text: raw.text,
		author: resolveAuthor(author),
		resolve,
	};
}

const REPLY_ENTRY_KEYS: ReadonlySet<string> = new Set([
	"at",
	"id",
	"text",
	"author",
	"resolve",
	"resolved",
]);

type ReplyEntry = {
	parentId: string;
	text: string;
	author: string;
	resolve: boolean;
};
