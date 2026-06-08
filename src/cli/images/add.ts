import { run as insertRun } from "../insert";
import { writeStdout } from "../respond";

const HELP = `docx images add — insert an image (an alias for \`docx insert --image\`)

Usage:
  docx images add FILE --image PATH --after pN [options]
  docx images add FILE --image PATH --before pN [options]

This is sugar for the canonical \`docx insert --image\`; every option that
\`insert\` accepts works here (--alt, --width, --height, --caption, --track, -o,
--dry-run, --verbose). Images are sized to fit the page width by default; pass
--width INCHES to override. Discover paragraph locators with \`docx read FILE\`.

Examples:
  docx images add report.docx --image chart.png --after p4 --alt "Figure 1"
  docx images add report.docx --image logo.png --before p0 --width 1.5
  docx images add report.docx --image fig.png --after p4 --caption "Figure 1: Revenue by quarter"
`;

export async function run(args: string[]): Promise<number> {
	if (args[0] === "-h" || args[0] === "--help") {
		await writeStdout(HELP);
		return 0;
	}
	// Thin alias: forward verbatim to `insert` (which owns image parsing, the
	// width clamp, and tracked-insert behavior). `--image` is required there.
	return insertRun(args);
}
