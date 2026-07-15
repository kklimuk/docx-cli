import { Edit, EditError, type ParagraphContentSpec } from "@core";
import { parseParagraphOptions } from "../insert/index";
import {
	EXIT,
	fail,
	openOrFail,
	resolveBlockOrFail,
	resolveBlockRangeOrFail,
	resolveTracked,
	respondAck,
	respondEditDryRun,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { resolveCodeContent } from "./content";

const HELP = `docx code edit — replace a paragraph (or range) with a code block

Usage:
  docx code edit FILE --at LOCATOR --code TEXT [--language LANG] [options]
  docx code edit FILE --at LOCATOR --code-file PATH [--language LANG] [options]

Locator (required):
  --at LOCATOR      What to replace. One of:
                      pN       a single paragraph
                      pN-pM    a range of paragraphs (replaced as a unit)
                    Discover ids with \`docx read FILE\`.

Content (exactly one of):
  --code TEXT       Inline code. Newlines split into one CodeBlock-styled
                    paragraph per source line.
  --code-file PATH  Read the code body from PATH ("-" reads stdin, so
                    \`cat main.py | docx code edit … --code-file -\` works).
  --language LANG   Syntax-highlight via lowlight (37 common languages bundled).
                    Survives round-trip via a CodeBlock-LANG pStyle suffix.

Paragraph options (apply to every code paragraph):
  --style NAME       Paragraph style      --alignment ALIGN  left|center|right|justify
  --space-before PT  --space-after PT     --line-spacing N
  --indent-left IN   --indent-right IN    --first-line IN   --hanging IN

Options:
  --track           Record the edit as a tracked change even when the
                    document's track-changes toggle is off.
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: a one-line confirmation)
  -h, --help        Show this help

Examples:
  docx code edit doc.docx --at p4 --code-file snippet.go --language go
  docx code edit doc.docx --at p4-p8 --code "print('hi')" --language python
  cat main.py | docx code edit doc.docx --at p4 --code-file - --language python

Output:
  Prints a one-line confirmation on success (exit 0). --verbose prints
  {ok:true, operation:"edit", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	at: { type: "string" },
	code: { type: "string" },
	"code-file": { type: "string" },
	language: { type: "string" },
	style: { type: "string" },
	alignment: { type: "string" },
	"space-before": { type: "string" },
	"space-after": { type: "string" },
	"line-spacing": { type: "string" },
	"indent-left": { type: "string" },
	"indent-right": { type: "string" },
	"first-line": { type: "string" },
	hanging: { type: "string" },
	author: { type: "string" },
	track: { type: "boolean" },
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

	const filePath = parsed.positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE argument", HELP);

	const locator = parsed.values.at as string | undefined;
	if (!locator)
		return fail("USAGE", "Missing --at LOCATOR (pN or pN-pM)", HELP);

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	const code = await resolveCodeContent(parsed.values, HELP);
	if (typeof code === "number") return code;
	const spec: Extract<ParagraphContentSpec, { kind: "code" }> = {
		kind: "code",
		...code,
		paragraphOptions,
	};

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const outputPath = parsed.values.output as string | undefined;
	if (parsed.values["dry-run"]) {
		return respondEditDryRun(filePath, locator, outputPath);
	}

	const track = resolveTracked(document, Boolean(parsed.values.track));
	const authorFlag = parsed.values.author as string | undefined;

	try {
		if (/^p\d+-p\d+$/.test(locator)) {
			const rangeRef = await resolveBlockRangeOrFail(document, locator);
			if (typeof rangeRef === "number") return rangeRef;
			new Edit(document).range(rangeRef, spec, { authorFlag, track });
		} else {
			const blockRef = await resolveBlockOrFail(document, locator);
			if (typeof blockRef === "number") return blockRef;
			new Edit(document).paragraph(blockRef, spec, { authorFlag, track });
		}
	} catch (error) {
		if (error instanceof EditError) {
			return fail(error.code, error.message, error.hint);
		}
		throw error;
	}

	await document.save(outputPath);
	await respondAck({
		ok: true,
		operation: "edit",
		path: outputPath ?? filePath,
		locator,
	});
	return EXIT.OK;
}
