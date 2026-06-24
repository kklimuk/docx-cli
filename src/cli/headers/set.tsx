import {
	type MarginalField,
	type MarginalKind,
	type MarginalSpec,
	Marginals,
} from "@core";
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
	text: { type: "string" },
	align: { type: "string" },
	"page-number": { type: "boolean" },
	"of-pages": { type: "boolean" },
	date: { type: "boolean" },
	"date-format": { type: "string" },
	"style-ref": { type: "string" },
	field: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
	...SAVE_FLAGS,
} as const;

function helpFor(kind: MarginalKind): string {
	const noun = marginalNoun(kind);
	const otherNoun = kind === "header" ? "footers" : "headers";
	return `docx ${noun} set — set (create or replace) a ${kind}

Usage:
  docx ${noun} set FILE [--at sN] [placement] (content) [options]

Placement (default = every page, all sections):
  --at sN            Target one section (default: the whole document — every
                     section, sharing one ${kind} part)
  --type T           default | first | even  (default: default)
  --first-page       ≡ --type first  (a different/blank first page)
  --even             ≡ --type even   (even pages; toggles evenAndOddHeaders)
  --odd              ≡ --type default (the default/odd-page ${kind})

Content (pick ONE primary source; --text + one field = two-zone):
  --text TEXT        Static text
  --align A          left | center | right (default: left for text, center for a
                     lone field; ignored in two-zone, where text is left + field right)
  --page-number      Insert the page number (PAGE field)
  --of-pages         Make it "Page X of Y" (PAGE of NUMPAGES); implies --page-number
  --date             Insert an auto-updating date (DATE field)
  --date-format FMT  Date format (e.g. "MMMM d, yyyy"); with --date
  --style-ref STYLE  Running head: the current STYLE heading's text (STYLEREF)
  --field F          A document field: filename | title | author

Tracking:
  --track            Record the reference change as a tracked <w:sectPrChange>
  --author NAME      Tracked-change author (NOT the --field author doc property)

${SAVE_FLAGS_HELP}

Examples:
  docx ${noun} set doc.docx --text "Q3 Report"
  docx ${noun} set doc.docx --page-number --of-pages          # "Page 3 of 12", centered
  docx ${noun} set doc.docx --text "Q3 Report" --page-number  # title left, page right
  docx ${noun} set doc.docx --first-page --text ""            # blank first page
  docx ${noun} set doc.docx --at s1 --style-ref "Heading 1"   # running head on one section
  docx ${noun} set doc.docx --date --date-format "MMMM yyyy"

See also: docx ${noun} list / clear, docx ${otherNoun} set.
`;
}

const SAVE_FLAGS_HELP = `Output:
  Prints a one-line confirmation (exit 0). --verbose prints the full ack JSON.
  -o, --output PATH  Write to PATH instead of editing FILE in place
  --dry-run          Preview; write nothing
  -v, --verbose      Print the success ack JSON
  -h, --help         Show this help`;

export async function runSetMarginal(
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

	const field = resolveField(parsed.values);
	if (isFieldError(field)) return fail("USAGE", field.error, field.hint);

	const align = parsed.values.align as string | undefined;
	if (align !== undefined && !isAlign(align)) {
		return fail(
			"USAGE",
			`Invalid --align: ${align}`,
			"Valid: left, center, right.",
		);
	}

	const text = parsed.values.text as string | undefined;
	if (text === undefined && field === undefined) {
		return fail(
			"USAGE",
			`Nothing to set — pass --text and/or a field (--page-number/--date/--style-ref/--field).`,
			helpFor(kind),
		);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const atLocator = parsed.values.at as string | undefined;
	const targets = await resolveTargetSectPrs(document, atLocator);
	if (typeof targets === "number") return targets;

	const spec: MarginalSpec = { text, align, field };
	const tracked = resolveTracked(document, parsed.values.track);
	const result = new Marginals(document).set(targets, kind, type, spec, {
		track: tracked,
		authorFlag: parsed.values.author as string | undefined,
	});

	const noun = marginalNoun(kind);
	const dryRun = Boolean(parsed.values["dry-run"]);
	const outputPath = parsed.values.output as string | undefined;
	if (dryRun) {
		await respond({
			operation: `${noun}.set`,
			dryRun: true,
			path: filePath,
			kind,
			type,
			sections: result.sections,
			...(atLocator ? { locator: atLocator } : {}),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	await document.save(outputPath);
	const destination = outputPath ?? filePath;
	// `--track` records a `<w:sectPrChange>` only when a section's REFERENCE
	// changes; a pure content replace of an existing reference rewrites the part
	// body directly (header/footer body edits aren't individually tracked in v1).
	// Warn so the agent isn't misled into thinking the edit is redlined.
	const untrackedContentReplace = tracked && result.referencesChanged === 0;
	const trackNote = untrackedContentReplace
		? `note: --track recorded no revision — the ${kind} reference was unchanged, and ${kind} body edits aren't individually tracked in v1 (the content was replaced directly). `
		: "";
	await respondAck(
		{
			ok: true,
			operation: `${noun}.set`,
			path: destination,
			kind,
			type,
			sections: result.sections,
			part: result.partName,
			...(untrackedContentReplace ? { trackedRevision: false } : {}),
			...(atLocator ? { locator: atLocator } : { applied: result.sections }),
		},
		trackNote + renderVerifyHint(destination),
	);
	return EXIT.OK;
}

type FieldError = { error: string; hint?: string };

/** Resolve the (at most one) field source from the field flags. */
function resolveField(values: {
	"page-number"?: unknown;
	"of-pages"?: unknown;
	date?: unknown;
	"date-format"?: unknown;
	"style-ref"?: unknown;
	field?: unknown;
}): MarginalField | undefined | FieldError {
	const fields: MarginalField[] = [];
	const ofPages = Boolean(values["of-pages"]);
	if (values["page-number"] || ofPages) {
		fields.push({ type: "page", ofPages });
	}
	if (values.date) {
		const format = values["date-format"] as string | undefined;
		fields.push({ type: "date", ...(format ? { format } : {}) });
	}
	const styleRef = values["style-ref"] as string | undefined;
	if (styleRef !== undefined)
		fields.push({ type: "styleRef", style: styleRef });
	const fieldRaw = values.field as string | undefined;
	if (fieldRaw !== undefined) {
		if (
			fieldRaw !== "filename" &&
			fieldRaw !== "title" &&
			fieldRaw !== "author"
		) {
			return {
				error: `Invalid --field: ${fieldRaw}`,
				hint: "Valid: filename, title, author.",
			};
		}
		fields.push({ type: fieldRaw });
	}
	if (fields.length > 1) {
		return {
			error:
				"Pass at most one field source (--page-number / --date / --style-ref / --field)",
			hint: "Combine ONE field with --text for a two-zone marginal (text left, field right).",
		};
	}
	return fields[0];
}

function isFieldError(value: unknown): value is FieldError {
	return typeof value === "object" && value !== null && "error" in value;
}

function isAlign(value: string): value is "left" | "center" | "right" {
	return value === "left" || value === "center" || value === "right";
}
