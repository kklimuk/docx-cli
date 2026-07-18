import { realpathSync } from "node:fs";
import { resolve } from "node:path";
import { Document } from "@core";
import { diffHunks, diffStats, renderUnified } from "@core/diff";
import { MarkdownLocatorError } from "../read/markdown";
import { renderReadMarkdown } from "../read/render";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { normalizeReadMarkers } from "./markers";

const HELP = `docx diff — show what changed between this document and another version

Usage:
  docx diff FILE --against SRC [options]

Examples:
  # There is no undo/history (edits overwrite FILE in place), so SNAPSHOT
  # before editing, then diff against the snapshot:
  cp doc.docx doc.orig.docx        # keep the original
  docx edit doc.docx ...           # make your changes
  docx diff doc.docx --against doc.orig.docx
  # …other baselines:
  git show main:report.docx | docx diff report.docx --against -
  docx read old.docx > old.md && docx diff new.docx --against old.md

Options:
  FILE               the current document (the + / "current" side)
  --against SRC      REQUIRED. The baseline to compare against (the - side):
                       • a .docx file        (rendered and compared)
                       • a saved \`docx read\` text file (compared as-is)
                       • -                   (stdin: a piped .docx or text)
  --from LOC         compare only a block slice (both sides); .docx --against only
  --to LOC           end of the block slice
  --comments         include comment footnotes in the compared markdown
  --json             emit structured hunks + stats as JSON
  -h, --help         show this help

This is a read-only REPORT of what changed (not a patch to apply): both sides
render to their \`docx read\` markdown and compare as a git-style unified diff.
Locators (\`<!-- p3 -->\`, cell/row ids) are normalized out so a structural
edit doesn't renumber every following line; formatting/structure changes
(shading, borders, tracked-changes state) still show. To get a locator to
edit at, run \`docx read\`.

Output:
  A unified diff: a summary line (hunk / +− counts + legend), then hunks —
  "-" lines are the baseline, "+" lines are FILE (current), and a small
  one-line change shows inline as [-removed-]{+added+}. Identical documents
  print "No differences …"; exit 0 either way. --json: structured hunks +
  stats (no envelope). Errors print {code, error, hint?} with a nonzero exit.
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			against: { type: "string" },
			from: { type: "string" },
			to: { type: "string" },
			comments: { type: "boolean" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);
	if (filePath === "-") {
		return fail(
			"USAGE",
			"FILE must be a path — only --against reads stdin.",
			HELP,
		);
	}
	if (parsed.positionals[1] !== undefined) {
		return fail(
			"USAGE",
			`Unexpected extra argument: ${parsed.positionals[1]} (diff takes one FILE and --against SRC)`,
			HELP,
		);
	}

	const against = parsed.values.against as string | undefined;
	if (!against) {
		return fail(
			"USAGE",
			"Missing --against SRC (the baseline to compare against). Snapshot before editing: cp doc.docx doc.orig.docx, then --against doc.orig.docx.",
			HELP,
		);
	}

	const from = parsed.values.from as string | undefined;
	const to = parsed.values.to as string | undefined;
	const showComments = Boolean(parsed.values.comments);
	const json = Boolean(parsed.values.json);

	// Current (NEW / + ) side.
	const currentDoc = await openOrFail(filePath);
	if (typeof currentDoc === "number") return currentDoc;

	// Baseline (OLD / - ) side — read bytes from a path or stdin, sniff the zip
	// magic to decide .docx-vs-text.
	const bytes = await readAgainstBytes(against);
	if (typeof bytes === "number") return bytes;
	const isZip = looksLikeZip(bytes);

	if (!isZip && (from || to)) {
		return fail(
			"USAGE",
			"--from/--to require a .docx --against (a raw-text baseline can't be sliced).",
			HELP,
		);
	}

	// Self-compare guard: diffing a file against itself yields nothing — say so
	// plainly rather than a bland "no differences".
	if (against !== "-" && samePath(filePath, against)) {
		await writeStdout(
			`Comparing ${filePath} to itself — pass --against a different version.\n`,
		);
		return EXIT.OK;
	}

	// Both sides render with the SAME options — each doc supplies its own
	// defaults inside renderReadMarkdown — so the diff reflects real content
	// changes, not option drift between the two renders.
	const renderOptions = { from, to, view: "accepted" as const, showComments };

	const oldMarkdown = await resolveBaselineMarkdown(
		bytes,
		isZip,
		against,
		renderOptions,
	);
	if (typeof oldMarkdown === "number") return oldMarkdown;

	let newMarkdown: string;
	try {
		newMarkdown = await renderReadMarkdown(currentDoc, renderOptions);
	} catch (err) {
		if (err instanceof MarkdownLocatorError) {
			return fail("INVALID_LOCATOR", err.message);
		}
		throw err;
	}

	const oldNorm = normalizeReadMarkers(oldMarkdown);
	const newNorm = normalizeReadMarkers(newMarkdown);
	const hunks = diffHunks(oldNorm, newNorm);
	const stats = diffStats(hunks);
	const oldLabel = `${againstLabel(against)} (baseline)`;
	const newLabel = `${filePath} (current)`;

	if (json) {
		await respond({ hunks, oldLabel, newLabel, stats });
		return EXIT.OK;
	}

	if (hunks.length === 0) {
		await writeStdout(
			`No differences between ${filePath} and ${againstLabel(against)}.\n`,
		);
		return EXIT.OK;
	}

	const preamble = `# ${plural(stats.hunks, "hunk")}, +${stats.added} -${stats.removed}  ·  legend: - baseline, + current, [-removed-]{+added+} = changed on one line\n`;
	const body = renderUnified(hunks, { oldLabel, newLabel, wordDiff: true });
	await writeStdout(preamble + body);
	return EXIT.OK;
}

/** Render the baseline (`--against`) side to its read markdown: decode an
 *  already-rendered read-text baseline verbatim, or open+render a `.docx` (piped
 *  or on disk). Returns the markdown, or a fail() exit code when the bytes are
 *  neither valid text nor a readable `.docx`, or a `--from`/`--to` slice is bad. */
async function resolveBaselineMarkdown(
	bytes: Uint8Array,
	isZip: boolean,
	against: string,
	options: Parameters<typeof renderReadMarkdown>[1],
): Promise<string | number> {
	const label = againstLabel(against);
	if (!isZip) {
		try {
			return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
		} catch {
			return fail(
				"NOT_A_ZIP",
				`--against ${label} is neither a .docx nor decodable UTF-8 text.`,
			);
		}
	}
	try {
		const doc = await Document.openFromBytes(bytes, label);
		return await renderReadMarkdown(doc, options);
	} catch (err) {
		if (err instanceof MarkdownLocatorError) {
			return fail("INVALID_LOCATOR", err.message);
		}
		return fail(
			"NOT_A_ZIP",
			`Could not read ${label} as a .docx: ${err instanceof Error ? err.message : String(err)}`,
		);
	}
}

/** Read the `--against` source into bytes: stdin for `-`, else a file (with an
 *  existence check that maps to FILE_NOT_FOUND like every other missing path). */
async function readAgainstBytes(src: string): Promise<Uint8Array | number> {
	if (src === "-") {
		return new Uint8Array(await new Response(Bun.stdin.stream()).arrayBuffer());
	}
	const file = Bun.file(src);
	if (!(await file.exists())) {
		return fail("FILE_NOT_FOUND", `--against file not found: ${src}`);
	}
	return new Uint8Array(await file.arrayBuffer());
}

/** ZIP local-file-header magic `PK\x03\x04` — a .docx is a zip, so this
 *  distinguishes a piped/binary .docx from saved read-markdown text. */
function looksLikeZip(bytes: Uint8Array): boolean {
	return (
		bytes.length >= 4 &&
		bytes[0] === 0x50 &&
		bytes[1] === 0x4b &&
		bytes[2] === 0x03 &&
		bytes[3] === 0x04
	);
}

function againstLabel(src: string): string {
	return src === "-" ? "<stdin>" : src;
}

/** True when two paths point at the same file on disk (resolves symlinks; falls
 *  back to absolute-path comparison if a path can't be realpath'd). */
function samePath(a: string, b: string): boolean {
	try {
		return realpathSync(a) === realpathSync(b);
	} catch {
		return resolve(a) === resolve(b);
	}
}

function plural(count: number, noun: string): string {
	return `${count} ${noun}${count === 1 ? "" : "s"}`;
}
