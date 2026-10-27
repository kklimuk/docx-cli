import { parseParagraphOptions, type RawValues } from "../insert/index";
import { parseTargetPlacement, placeSpec } from "../insert/place";
import { decodeInlineEscapes } from "../parse-helpers";
import { type LengthOptions, parseLength } from "../parse-length";
import {
	EXIT,
	fail,
	RENDER_VERIFY_EXAMPLE,
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

Examples:
  docx images add report.docx --after p4 --image chart.png --alt "Figure 1"
  docx images add report.docx --before p0 --image logo.png --width 1.5
  docx images add report.docx --after p4 --image fig.png --caption "Figure 1: Revenue by quarter"
  docx images add report.docx --at-end --image https://example.com/logo.png --width 2

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
  --width LENGTH    Display width: bare inches (2) or a unit suffix — in, cm,
                    mm, pt, px (96dpi), tw (e.g. 2in, 5cm, 300px). Default:
                    native pixel size at 96dpi.
  --height LENGTH   Display height, same units (default: scales to preserve
                    aspect)
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

AGENT VERIFICATION: \`docx read\` shows the figure and its size hint, but NOT
how it lands on the page. After adding an image, render and READ the images:
${RENDER_VERIFY_EXAMPLE}
Check the figure is sized sensibly (no margin overflow) and re-render if needed.

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
 *  into the fields of an `image` spec. `--width`/`--height` parse through the
 *  shared length grammar (`IMAGE_LENGTH`): bare inches as before, or a unit
 *  suffix — `3cm` used to be silently read as 3 INCHES. `--caption` is a body
 *  paragraph (rendered via <Paragraph>, which turns a `\n` into <w:br/>), so it
 *  decodes inline escapes like the sibling --text flag; --alt is single-line accessibility metadata where a newline is
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

	for (const flag of ["width", "height"] as const) {
		const raw = values[flag] as string | undefined;
		if (raw === undefined) continue;
		const twips = parseLength(raw, flag, IMAGE_LENGTH);
		if (typeof twips !== "number")
			return fail("USAGE", twips.error, twips.hint);
		out[flag === "width" ? "widthInches" : "heightInches"] =
			twips / TWIPS_PER_INCH;
	}

	return out;
}

/** Image sizes: bare inches (the long-standing meaning), positive, and `px` at
 *  96 dpi — an image is the one thing with a natural pixel size (the default
 *  `--width` already IS its pixel size at 96 dpi), so `300px` is what an agent
 *  holding pixel dimensions reaches for. */
const IMAGE_LENGTH: LengthOptions = {
	bare: "in",
	positive: true,
	pixels: true,
};

const TWIPS_PER_INCH = 1440;
