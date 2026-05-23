import { saveDocView } from "@core";
import { w } from "@core/jsx";
import {
	findNoteByNumericId,
	type NoteKind,
	noteConfig,
	wrapNoteBodyAsEdited,
} from "@core/notes";
import {
	createRevisionAllocator,
	isTrackChangesEnabled,
	resolveAuthor,
	resolveDate,
	type TrackedMeta,
} from "@core/track-changes";
import { parseArgs } from "util";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	respondAck,
	setVerboseAck,
	writeStdout,
} from "../respond";

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	return `docx ${verb} edit — replace a ${kind}'s body text

Usage:
  docx ${verb} edit FILE --id ${idPrefix}N --text TEXT [options]

Required:
  --id ID              ${capitalize(kind)} id (e.g., ${idPrefix}1). The
                       \`${idPrefix}\` prefix is optional.
  --text TEXT          New body text. Replaces the current paragraph(s).

Optional:
  -o, --output PATH    Write to PATH instead of overwriting FILE.
  --dry-run            Print what would change; do not write the file.
  -v, --verbose        Print the success ack JSON (default: silent on success).
  -h, --help           Show this help.

Examples:
  docx ${verb} edit doc.docx --id ${idPrefix}2 --text "Updated citation."
`;
}

function capitalize(word: string): string {
	return word.charAt(0).toUpperCase() + word.slice(1);
}

export async function runEditNote(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				id: { type: "string" },
				text: { type: "string" },
				author: { type: "string" },
				output: { type: "string", short: "o" },
				"dry-run": { type: "boolean" },
				verbose: { type: "boolean", short: "v" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, helpFor(kind));
	}

	const help = helpFor(kind);
	if (parsed.values.help) {
		await writeStdout(help);
		return EXIT.OK;
	}

	setVerboseAck(Boolean(parsed.values.verbose));

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", help);

	const idInput = parsed.values.id as string | undefined;
	const text = parsed.values.text as string | undefined;
	const config = noteConfig(kind);
	if (!idInput) return fail("USAGE", `Missing --id ${config.idPrefix}N`, help);
	if (text === undefined) return fail("USAGE", "Missing --text TEXT", help);

	const numericId = idInput.startsWith(config.idPrefix)
		? idInput.slice(config.idPrefix.length)
		: idInput;
	const idLabel = `${config.idPrefix}${numericId}`;

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const reference = findNoteByNumericId(view, kind, numericId);
	if (!reference) {
		return fail("BLOCK_NOT_FOUND", `${capitalize(kind)} not found: ${idLabel}`);
	}

	const outputPath = parsed.values.output as string | undefined;

	if (parsed.values["dry-run"]) {
		await respond({
			ok: true,
			operation: `${kind}s.edit`,
			dryRun: true,
			path,
			...(outputPath ? { output: outputPath } : {}),
			id: idLabel,
		});
		return EXIT.OK;
	}

	const tracked = isTrackChangesEnabled(view);
	if (tracked) {
		// Tracked replace: `<w:ins>NEW</w:ins><w:del>OLD</w:del>` inside the
		// existing body paragraph, with `<w:footnoteRef/>` and leading
		// whitespace bare. Matches Word's empirical shape from
		// `/tmp/fn-probe/edit.docx` (ins precedes del in document order).
		// Word allocates distinct revision ids for the two sides — we match
		// that for empirical parity even though the OOXML spec allows shared
		// ids.
		const allocator = createRevisionAllocator(view);
		const author = resolveAuthor(parsed.values.author as string | undefined);
		const date = resolveDate();
		const insMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		const delMeta: TrackedMeta = {
			author,
			date,
			revisionId: allocator.next(),
		};
		wrapNoteBodyAsEdited(reference.node, ` ${text}`, insMeta, delMeta);
	} else {
		// Untracked: replace the note's content with a fresh single-paragraph
		// body. The `<w:footnote>` / `<w:endnote>` wrapper itself stays —
		// anything that referenced the original (back-references, custom
		// marker attributes) keeps pointing here. Only the inner paragraph
		// content is rewritten.
		reference.node.children = [<NoteParagraph config={config} text={text} />];
	}

	await saveDocView(view, outputPath);

	await respondAck({
		ok: true,
		operation: `${kind}s.edit`,
		path: outputPath ?? path,
		id: idLabel,
	});
	return EXIT.OK;
}

function NoteParagraph({
	config,
	text,
}: {
	config: ReturnType<typeof noteConfig>;
	text: string;
}) {
	const BodyRef = config.kind === "footnote" ? w.footnoteRef : w.endnoteRef;
	return (
		<w.p>
			<w.pPr>
				<w.pStyle w-val={config.textStyle} />
			</w.pPr>
			<w.r>
				<w.rPr>
					<w.rStyle w-val={config.referenceStyle} />
				</w.rPr>
				<BodyRef />
			</w.r>
			<w.r>
				<w.t {...{ "xml:space": "preserve" }}>{` ${text}`}</w.t>
			</w.r>
		</w.p>
	);
}

export async function run(args: string[]): Promise<number> {
	return runEditNote(args, "footnote");
}
