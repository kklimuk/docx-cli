import type { XmlNode } from "@core";
import { findContainingParagraph } from "@core/comments";
import { ImageFrameError, Images } from "@core/image";
import { addAuditComment } from "../audit-comment";
import {
	batchExampleIntro,
	readJsonlObjects,
	rejectBatchOnlyFlags,
} from "../parse-helpers";
import { formatImageSize } from "../read/annotations";
import {
	EXIT,
	fail,
	openOrFail,
	RENDER_VERIFY_EXAMPLE,
	renderVerifyHint,
	respond,
	respondAck,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { parseImageInches, parseImageSizeInches } from "./length";

const HELP = `docx images set — resize an existing image or change its alt text

Usage:
  docx images set FILE --at imgN [--width L | --height L | --size WxH | --fit] [--alt TEXT]
  docx images set FILE --batch FILE.jsonl

Examples:
  docx images set doc.docx --at img0 --width 2in          # height follows (aspect kept)
  docx images set doc.docx --at img1 --fit                # shrink to the text width
  docx images set doc.docx --at img0 --alt "Company logo"
  docx images set doc.docx --at img0 --size 2x1.5in       # the docx:image size="…" hint, verbatim
${batchExampleIntro("Resize / re-describe many images")}
  printf '%s\\n' '{"at":"img0","width":"2in","alt":"Logo"}' '{"at":"img3","fit":true}' > sets.jsonl
  docx images set doc.docx --batch sets.jsonl

Target:
  --at imgN         The image to change (ids from \`docx read\`'s
                    <!-- docx:image imgN … --> hints or \`docx images list\`)
  --batch FILE      JSONL, one {"at":"imgN", …} per line (keys mirror the
                    flags: width, height, size, fit, alt). "-" reads stdin.

Size (at most one):
  --width LENGTH    New width; the height follows to keep the aspect ratio.
                    Bare inches (2) or a unit suffix — in, cm, mm, pt, px (96dpi)
  --height LENGTH   New height; the width follows to keep the aspect ratio.
  --size WxH        Both sides exactly, as the docx:image hint prints them
                    (2x1.5in) — can distort if the ratio differs.
  --fit             Shrink to the page's text width (aspect kept) when the
                    image is wider — the cure for overflow="yes". Never enlarges.

Description:
  --alt TEXT        Alt text (Word's "Edit Alt Text"); --alt "" clears it.

Options:
  --author NAME     Author for the audit comment when track-changes is on
                    (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

To CENTER an image, align its paragraph: docx edit FILE --at pN --alignment center.

Image size/alt changes can't be recorded as tracked changes, so when
track-changes is on they apply directly and an audit comment marks each image.
A call that changes nothing (e.g. --fit on an image that already fits) exits
nonzero. After resizing, render and look:
${RENDER_VERIFY_EXAMPLE}

Output:
  Prints "images.set imgN" plus the image's new size hint (exit 0). --verbose
  prints {ok:true, operation, path, images:[{imageId, size, alt, …}]}. Errors
  print {code, error, hint?} with a nonzero exit.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	width: { type: "string" },
	height: { type: "string" },
	size: { type: "string" },
	fit: { type: "boolean" },
	alt: { type: "string" },
	batch: { type: "string" },
	author: { type: "string" },
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

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const entries = await readEntries(parsed.values);
	if (typeof entries === "number") return entries;

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	// Every entry is applied in memory before anything is written, so a bad
	// entry anywhere leaves the file untouched. `imgN` ids never shift here —
	// `set` adds and removes no drawing — so batch locators address the document
	// as read without a pre-resolve pass.
	const images = new Images(document);
	const results: AppliedSet[] = [];
	const targeted = new Set<XmlNode>();
	for (const [index, entry] of entries.entries()) {
		const { at, ...change } = entry;
		// One entry per image: a second would see the first's result as its
		// "before" and stack a second audit comment on the same run.
		const run = images.find(at)?.run;
		if (run && targeted.has(run)) {
			return fail(
				"USAGE",
				`entry ${index}: ${at} is already targeted by an earlier entry`,
				"Put every change for one image (size + alt) in a single entry.",
			);
		}
		if (run) targeted.add(run);
		try {
			results.push({ imageId: at, ...images.set(at, change) });
		} catch (error) {
			if (!(error instanceof ImageFrameError)) throw error;
			const prefix = entries.length > 1 ? `entry ${index}: ` : "";
			return fail(error.code, `${prefix}${at}: ${error.message}`, error.hint);
		}
	}

	const changed = results.filter(
		(result) => result.resized || result.altChanged,
	);
	if (changed.length === 0) return failNoChange(results);

	const outputPath = parsed.values.output as string | undefined;
	const payload = {
		operation: "images.set",
		path: outputPath ?? path,
		...(results.length === 1
			? { imageId: results[0]?.imageId }
			: { count: changed.length }),
		images: changed.map(describeResult),
	};

	if (parsed.values["dry-run"]) {
		await respond({ ...payload, dryRun: true });
		return EXIT.OK;
	}

	if (document.isTrackChangesEnabled()) {
		const authorFlag = parsed.values.author as string | undefined;
		for (const result of changed) {
			const paragraph = findContainingParagraph(
				document.documentTree,
				result.hit.run,
			);
			if (!paragraph) continue;
			addAuditComment(
				document,
				{ kind: "run", paragraph, run: result.hit.run },
				auditMessage(result),
				authorFlag,
			);
		}
	}

	await document.save(outputPath);

	const summary = changed.map(
		(result) => `${result.imageId} ${formatState(result.after)}`,
	);
	if (changed.some((result) => result.resized)) {
		summary.push(renderVerifyHint(outputPath ?? path));
	}
	await respondAck({ ok: true, ...payload }, summary.join("\n"));
	return EXIT.OK;
}

/** The entries to apply: one from the flags, or one per `--batch` JSONL line. */
async function readEntries(
	values: Record<string, unknown>,
): Promise<SetEntry[] | number> {
	const batchSource = values.batch as string | undefined;
	if (batchSource === undefined) {
		const entry = parseEntry(values, "flags");
		if ("error" in entry) return fail("USAGE", entry.error, entry.hint);
		return [entry];
	}
	const conflict = await rejectBatchOnlyFlags(
		values,
		["at", "width", "height", "size", "fit", "alt"],
		"image change",
		'Put one {"at":"imgN", …} per line.',
	);
	if (conflict !== undefined) return conflict;
	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");
	const entries: SetEntry[] = [];
	for (const [index, raw] of rawEntries.entries()) {
		const entry = parseEntry(raw, "batch");
		if ("error" in entry) {
			return fail("USAGE", `entry ${index}: ${entry.error}`, entry.hint);
		}
		entries.push(entry);
	}
	return entries;
}

/** One change from CLI flags or a JSONL object (same keys). Batch values may
 *  be JSON numbers (bare inches) or strings; `fit` must be a real boolean. */
function parseEntry(
	raw: Record<string, unknown>,
	source: "flags" | "batch",
): SetEntry | { error: string; hint?: string } {
	const flag = (key: string) => (source === "flags" ? `--${key}` : `"${key}"`);
	const at = raw.at;
	if (typeof at !== "string" || at.length === 0) {
		return { error: `${flag("at")} imgN is required` };
	}
	const entry: SetEntry = { at };

	const lengths: Record<string, string> = {};
	for (const key of ["width", "height", "size"] as const) {
		const value = raw[key];
		if (value === undefined) continue;
		if (typeof value !== "string" && typeof value !== "number") {
			return { error: `${flag(key)} must be a length like "2in"` };
		}
		lengths[key] = String(value);
	}
	if (raw.fit !== undefined && typeof raw.fit !== "boolean") {
		return { error: `${flag("fit")} must be true or false` };
	}
	const sizeFlags = [
		...Object.keys(lengths),
		...(raw.fit === true ? ["fit"] : []),
	];
	if (sizeFlags.length > 1) {
		return {
			error: `Pass one size option, not ${sizeFlags.map(flag).join(" + ")}`,
			hint: "--width OR --height keeps the aspect ratio; --size WxH sets both; --fit shrinks to the text width.",
		};
	}

	if (lengths.size !== undefined) {
		const size = parseImageSizeInches(lengths.size);
		if ("error" in size) return size;
		entry.widthEmu = inchesToEmu(size.widthInches);
		entry.heightEmu = inchesToEmu(size.heightInches);
	}
	for (const key of ["width", "height"] as const) {
		const value = lengths[key];
		if (value === undefined) continue;
		const inches = parseImageInches(value, key);
		if (typeof inches !== "number") return inches;
		entry[key === "width" ? "widthEmu" : "heightEmu"] = inchesToEmu(inches);
	}
	if (raw.fit === true) entry.fit = true;

	if (raw.alt !== undefined) {
		if (typeof raw.alt !== "string") {
			return { error: `${flag("alt")} must be a string` };
		}
		entry.alt = raw.alt;
	}

	if (sizeFlags.length === 0 && entry.alt === undefined) {
		return {
			error: "Nothing to change",
			hint: "Pass --width/--height/--size/--fit to resize, and/or --alt TEXT.",
		};
	}
	return entry;
}

const inchesToEmu = (inches: number): number => Math.round(inches * 914400);

/** A zero-effect `set` exits nonzero: a weak agent keys off the exit code, so
 *  "--fit on an image that already fits" must not read as a success. */
function failNoChange(results: AppliedSet[]): Promise<number> {
	const states = results
		.map((result) => `${result.imageId} ${formatState(result.before)}`)
		.join("; ");
	return fail(
		"MATCH_NOT_FOUND",
		`Nothing changed — the image${results.length === 1 ? " is" : "s are"} already that way (${states}).`,
		"--fit only shrinks an image WIDER than the text width; pass --width/--height to resize to a specific size.",
	);
}

/** The new state in the `docx:image` hint's own vocabulary, so the agent sees
 *  the result without re-reading. */
function formatState(state: AppliedSet["after"]): string {
	const size = `size="${formatImageSize(state.widthEmu, state.heightEmu)}"`;
	return state.alt ? `${size} alt=${JSON.stringify(state.alt)}` : size;
}

function describeResult(result: AppliedSet) {
	return {
		imageId: result.imageId,
		size: formatImageSize(result.after.widthEmu, result.after.heightEmu),
		widthEmu: result.after.widthEmu,
		heightEmu: result.after.heightEmu,
		alt: result.after.alt,
		resized: result.resized,
		altChanged: result.altChanged,
	};
}

function auditMessage(result: AppliedSet): string {
	const parts: string[] = [];
	if (result.resized) {
		parts.push(
			`image resized: ${formatImageSize(result.before.widthEmu, result.before.heightEmu)} → ${formatImageSize(result.after.widthEmu, result.after.heightEmu)}`,
		);
	}
	if (result.altChanged) {
		parts.push(
			`alt text: ${JSON.stringify(result.before.alt)} → ${JSON.stringify(result.after.alt)}`,
		);
	}
	return parts.join("; ");
}

type SetEntry = {
	at: string;
	widthEmu?: number;
	heightEmu?: number;
	fit?: boolean;
	alt?: string;
};

type AppliedSet = { imageId: string } & ReturnType<Images["set"]>;
