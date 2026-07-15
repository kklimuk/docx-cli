import type { InsertSpec } from "@core";
import { parseParagraphOptions } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import {
	EXIT,
	fail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import { resolveCodeContent } from "./content";

const HELP = `docx code add — insert a syntax-highlighted code block

Usage:
  docx code add FILE --after LOCATOR --code TEXT [--language LANG] [options]
  docx code add FILE --before LOCATOR --code-file PATH [--language LANG] [options]
  docx code add FILE (--at-start | --at-end) --code-file PATH [options]

Placement (exactly one required):
  --after LOCATOR   Insert after the block at LOCATOR (a pN / tN / cell paragraph)
  --before LOCATOR  Insert before the block at LOCATOR
  --at-start        Insert at the very top of the document (no locator needed)
  --at-end          Insert at the very end, before the trailing section properties

Content (exactly one of):
  --code TEXT       Inline code. Newlines split into one CodeBlock-styled
                    paragraph per source line.
  --code-file PATH  Read the code body from PATH ("-" reads stdin, so
                    \`cat main.py | docx code add … --code-file -\` works).
  --language LANG   Syntax-highlight via lowlight (37 common languages bundled).
                    Survives round-trip via a CodeBlock-LANG pStyle suffix.

Paragraph options (apply to every code paragraph):
  --style NAME       Paragraph style      --alignment ALIGN  left|center|right|justify
  --space-before PT  --space-after PT     --line-spacing N
  --indent-left IN   --indent-right IN    --first-line IN   --hanging IN

Options:
  --track           Record the insert as a tracked change even when the
                    document's track-changes toggle is off.
  --author NAME     Author for tracked changes (default: $DOCX_AUTHOR)
  -o, --output PATH Write to PATH instead of overwriting FILE
  --dry-run         Report what would change without writing the file
  -v, --verbose     Print the success ack JSON (default: the minted locator)
  -h, --help        Show this help

Examples:
  docx code add doc.docx --after p4 --code-file snippet.go --language go
  cat main.py | docx code add doc.docx --at-end --code-file - --language python
  docx code add doc.docx --before p0 --code "print('hi')" --language python

Output:
  Prints the inserted block's locator (pN), one per line. --verbose prints
  {ok:true, operation:"insert", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
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

	const placement = await parseTargetPlacement(parsed.values, HELP);
	if (typeof placement === "number") return placement;

	const code = await resolveCodeContent(parsed.values, HELP);
	if (typeof code === "number") return code;
	const spec: Extract<InsertSpec, { kind: "code" }> = { kind: "code", ...code };

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return placeSpec({
		filePath,
		placement,
		spec,
		paragraphOptions,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	});
}
