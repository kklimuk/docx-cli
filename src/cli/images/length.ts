import {
	type LengthError,
	type LengthOptions,
	parseLength,
	parseLengthList,
} from "../parse-length";

/** Parse one image `--width`/`--height` value into inches through the shared
 *  length grammar (`IMAGE_LENGTH`): bare inches, or a unit suffix — `3cm` used
 *  to be silently read as 3 INCHES. Shared by `images add` and `images set`. */
export function parseImageInches(
	raw: string,
	flag: string,
): number | LengthError {
	const twips = parseLength(raw, flag, IMAGE_LENGTH);
	if (typeof twips !== "number") return twips;
	return twips / TWIPS_PER_INCH;
}

/** Parse `images set --size WxH` — exactly the shape the `docx:image` hint
 *  prints (`size="2x1.5in"`: a unit on the height shared by both sides, the
 *  `docx:page size` convention), so a hint value re-applies verbatim. The
 *  height side must start with a number so a `px`'s `x` isn't the separator. */
export function parseImageSizeInches(
	raw: string,
): { widthInches: number; heightInches: number } | LengthError {
	const invalid: LengthError = {
		error: `Invalid --size: ${raw}`,
		hint: "Use WxH as the docx:image hint prints it, e.g. --size 2x1.5in (or one side alone: --width 2in keeps the aspect ratio).",
	};
	const sides = raw.trim().match(/^(.+?)\s*x\s*(\.?\d.*)$/i);
	if (!sides) return invalid;
	const twips = parseLengthList(
		`${sides[1]},${sides[2]}`,
		"size",
		IMAGE_LENGTH,
	);
	if (!Array.isArray(twips)) return twips;
	if (twips.length !== 2) return invalid;
	const [width, height] = twips as [number, number];
	return {
		widthInches: width / TWIPS_PER_INCH,
		heightInches: height / TWIPS_PER_INCH,
	};
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
