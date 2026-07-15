import { parseParagraphOptions, type RawValues } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import { decodeInlineEscapes } from "../parse-helpers";
import {
	EXIT,
	fail,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx images add — insert an image

Usage:
  docx images add FILE --after LOCATOR --image SRC [options]
  docx images add FILE --before LOCATOR --image SRC [options]
  docx images add FILE (--at-start | --at-end) --image SRC [options]

Placement (exactly one required):
  --after LOCATOR   Insert after the block at LOCATOR (a pN / tN / cell paragraph)
  --before LOCATOR  Insert before the block at LOCATOR
  --at-start        Insert at the very top of the document (no locator needed)
  --at-end          Insert at the very end, before the trailing section properties

Content (required):
  --image SRC       The image source: a local file path, a data: URI, or an
                    http(s) URL (fetched with an SSRF guard; sized to fit the
                    page width by default).
  --alt TEXT        Alt text / description for the image
  --width INCHES    Display width in inches (default: native pixel size at 96dpi)
  --height INCHES   Display height in inches (default: scales to preserve aspect)
  --caption TEXT    Add a caption paragraph below the figure in Word's built-in
                    "Caption" style (kept with the image; shows in a Table of
                    Figures). You supply the label, e.g. "Figure 1: Revenue".

Paragraph options (apply to the figure paragraph):
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

Agent tip: VERIFY LAYOUT VISUALLY. \`docx read\` shows the figure and its size
hint, but NOT how it lands on the page. After adding an image, render and look:
  docx render FILE --out pages/      # writes page-001.png, page-002.png, …
Check the figure is sized sensibly (no margin overflow) and re-render if needed.

Examples:
  docx images add report.docx --after p4 --image chart.png --alt "Figure 1"
  docx images add report.docx --before p0 --image logo.png --width 1.5
  docx images add report.docx --after p4 --image fig.png --caption "Figure 1: Revenue by quarter"
  docx images add report.docx --at-end --image https://example.com/logo.png --width 2

Output:
  Prints the inserted block's locator (pN), one per line. --verbose prints
  {ok:true, operation:"insert", …}. Errors print {code, error, hint?}.
`;

const OPTION_SPEC = {
	after: { type: "string" },
	before: { type: "string" },
	"at-start": { type: "boolean" },
	"at-end": { type: "boolean" },
	image: { type: "string" },
	alt: { type: "string" },
	width: { type: "string" },
	height: { type: "string" },
	caption: { type: "string" },
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

	const imageFlags = await parseImageFlags(parsed.values);
	if (typeof imageFlags === "number") return imageFlags;

	const paragraphOptions = await parseParagraphOptions(parsed.values);
	if (typeof paragraphOptions === "number") return paragraphOptions;

	return placeSpec({
		filePath,
		placement,
		spec: { kind: "image", ...imageFlags },
		paragraphOptions,
		authorFlag: parsed.values.author as string | undefined,
		trackFlag: Boolean(parsed.values.track),
		outputPath: parsed.values.output as string | undefined,
		dryRun: Boolean(parsed.values["dry-run"]),
	});
}

/** Resolve `--image SRC` (required) plus `--alt`/`--width`/`--height`/`--caption`
 *  into the fields of an `image` spec. `--width`/`--height` parse as positive
 *  inches. `--caption` is a body paragraph (rendered via <Paragraph>, which turns
 *  a `\n` into <w:br/>), so it decodes inline escapes like the sibling --text
 *  flag; --alt is single-line accessibility metadata where a newline is
 *  meaningless, so it stays literal. */
async function parseImageFlags(values: RawValues): Promise<
	| {
			src: string;
			alt?: string;
			widthInches?: number;
			heightInches?: number;
			caption?: string;
	  }
	| number
> {
	const src = values.image as string | undefined;
	if (!src) return fail("USAGE", "--image requires a SRC argument", HELP);

	const out: {
		src: string;
		alt?: string;
		widthInches?: number;
		heightInches?: number;
		caption?: string;
	} = { src };

	const alt = values.alt as string | undefined;
	if (alt !== undefined) out.alt = alt;

	const caption = decodeInlineEscapes(values.caption as string | undefined);
	if (caption !== undefined) out.caption = caption;

	const widthRaw = values.width as string | undefined;
	if (widthRaw !== undefined) {
		const width = Number.parseFloat(widthRaw);
		if (!Number.isFinite(width) || width <= 0) {
			return fail(
				"USAGE",
				`--width must be a positive number of inches, got "${widthRaw}"`,
			);
		}
		out.widthInches = width;
	}

	const heightRaw = values.height as string | undefined;
	if (heightRaw !== undefined) {
		const height = Number.parseFloat(heightRaw);
		if (!Number.isFinite(height) || height <= 0) {
			return fail(
				"USAGE",
				`--height must be a positive number of inches, got "${heightRaw}"`,
			);
		}
		out.heightInches = height;
	}

	return out;
}
