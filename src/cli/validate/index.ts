import {
	EXIT,
	fail,
	openPkgOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { runFix } from "./fix";
import { reportJson, reportText, validatePackage } from "./report";

const HELP = `docx validate — schema-check a document against ECMA-376, and fix what Word repairs

Usage:
  docx validate FILE [--json]
  docx validate FILE --fix [-o PATH] [--dry-run] [--json]

Validates every WordprocessingML part in the package — document, styles,
numbering, settings, footnotes, endnotes, comments, every header/footer,
fontTable, webSettings, … (any part whose root wml.xsd declares) — against
the bundled ECMA-376 5th-edition TRANSITIONAL schemas (the profile Word
itself writes) after markup-compatibility preprocessing (mc:Ignorable
extension markup is skipped, exactly as a consumer would).

Also checks the cross-part comment invariants a per-part schema can't see,
graded by what Word does on open:
  error    a commentReference names no comment — Word refuses to open the file
  warning  a comment missing from commentsExtended.xml / commentsIds.xml, or
           with no w14:paraId (Word re-adds it — as an OPEN, TOP-LEVEL comment,
           so a reply loses its thread and a resolved comment reopens)
  warning  a commentRangeStart without its commentRangeEnd, or the reverse
           (Word collapses the comment to its reference mark)
  warning  a comment range with no commentReference (Word drops the comment)
  warning  a comment with no marker in the document at all (Word deletes it)
  warning  a reply whose thread parent names no comment (Word drops the link)
These list under the part they concern, tagged "(consistency)"; one Word
handles them harmlessly (a stale commentsIds entry, a duplicate durableId) is
not reported, but --fix and every save clean it up anyway.

Output: a verdict line first — "valid — N parts checked[, W warnings]" or
"invalid — E errors …" — then each part's findings as "error: …" /
"warning: …". Exit 0 = no errors (warnings alone don't fail); exit 1 =
errors. --json: {path, valid, errors, warnings, parts:[{part, issues:[{message,
severity, kind?, line?, fixable?, hint?}]}]}.

--fix cures every finding marked "(fixable)" and saves:
  - removes a dangling commentReference (and its comment's other markers)
  - collapses a half-open comment range onto its reference mark, as Word does
  - closes a reference-less range and adds its reference mark
  - adds missing commentsExtended/commentsIds entries and w14:paraIds, prunes
    stale ones, de-duplicates durableIds (what every docx-cli save does)
It never deletes a comment: an orphan comment and a broken thread link have
no deterministic cure and stay reported. Prints what it fixed ("nothing to
fix" on a clean file — not an error), then the post-fix report; exit 1 only
if errors remain.
  -o, --output PATH   Write the fixed file to PATH instead of overwriting FILE
  --dry-run           Report what --fix would do; write nothing

Useful before/after hand-made XML changes (docx raw, or edits made outside
docx-cli) and as a "will Word open this?" first check — note Word TOLERATES
some schema violations, so a schema finding is "invalid OOXML," not always
"Word rejects it."
`;

const OPTION_SPEC = {
	json: { type: "boolean" },
	fix: { type: "boolean" },
	output: { type: "string", short: "o" },
	"dry-run": { type: "boolean" },
	help: { type: "boolean", short: "h" },
} as const;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, OPTION_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	const output = values.output as string | undefined;
	const dryRun = Boolean(values["dry-run"]);
	if (!values.fix && (output !== undefined || dryRun)) {
		return fail(
			"USAGE",
			"-o/--output and --dry-run apply only with --fix",
			`Did you mean: docx validate ${filePath} --fix${dryRun ? " --dry-run" : ""}${output ? ` -o ${output}` : ""}`,
		);
	}

	const pkg = await openPkgOrFail(filePath);
	if (typeof pkg === "number") return pkg;

	const report = await validatePackage(pkg);
	if (report.strict) {
		const note =
			"strict-profile document (ISO/IEC 29500 Strict namespaces) — the bundled schemas cover the TRANSITIONAL profile Word writes, so no verdict is possible";
		if (values.json) {
			await respond({ path: filePath, skipped: "strict-profile", note });
		} else {
			await writeStdout(`skipped — ${note}\n`);
		}
		return EXIT.OK;
	}

	if (values.fix) {
		return runFix(
			pkg,
			filePath,
			{ output, dryRun, json: Boolean(values.json) },
			report,
		);
	}

	if (values.json) {
		await respond({ path: filePath, ...reportJson(report, filePath) });
	} else {
		await writeStdout(reportText(report, filePath));
	}
	return report.errors === 0 ? EXIT.OK : EXIT.GENERAL_ERROR;
}
