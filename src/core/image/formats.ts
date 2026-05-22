// The image formats docx-cli treats as first-class — the SAME set we extract,
// insert, and replace, so a file pulled out by `images extract` drops straight
// back in via `insert --image` or `images replace`. Extension spellings are
// Word-canonical (jpeg/tiff, the names Word writes into word/media), and the
// HEIC/HEIF transcode lives in image-source.ts (insert-only input, never an
// embedded format).

// SVG is supported but sanitized on the insert path (see `sanitizeSvg` in
// `source.ts`) — unlike every other format here, SVG carries arbitrary XML
// (`<script>`, `on*` handlers, external `xlink:href`, animation event attrs)
// and we strip those before embedding so the docx doesn't smuggle active
// content. `extract` returns the bytes verbatim; consumers should treat them
// the same as any untrusted SVG.

/** Canonical extension → MIME, for the formats we embed/extract/replace. */
export const IMAGE_MIME_BY_EXTENSION: Record<string, string> = {
	png: "image/png",
	jpeg: "image/jpeg",
	jpg: "image/jpeg",
	gif: "image/gif",
	webp: "image/webp",
	bmp: "image/bmp",
	tiff: "image/tiff",
	tif: "image/tiff",
	svg: "image/svg+xml",
	emf: "image/x-emf",
	wmf: "image/x-wmf",
	ico: "image/vnd.microsoft.icon",
};

/** MIME → canonical extension (the inverse, plus the `image/jpg` alias). */
export const EXTENSION_BY_IMAGE_MIME: Record<string, string> = {
	"image/png": "png",
	"image/jpeg": "jpeg",
	"image/jpg": "jpeg",
	"image/gif": "gif",
	"image/webp": "webp",
	"image/bmp": "bmp",
	"image/tiff": "tiff",
	"image/tif": "tiff",
	"image/svg+xml": "svg",
	"image/x-emf": "emf",
	"image/x-wmf": "wmf",
	"image/vnd.microsoft.icon": "ico",
};

/** Common extension aliases normalized to the canonical spelling. */
export const CANONICAL_IMAGE_EXTENSION: Record<string, string> = {
	jpg: "jpeg",
	tif: "tiff",
};

/** Human-readable supported-format list for error hints. */
export const SUPPORTED_IMAGE_FORMATS = Object.keys(
	IMAGE_MIME_BY_EXTENSION,
).join(", ");

export function extensionForImageMime(mimeType: string): string | undefined {
	return EXTENSION_BY_IMAGE_MIME[mimeType.toLowerCase()];
}

/** Resolve a file extension to its canonical extension + Word MIME. Keying off
 * the extension (not a sniffed MIME) keeps insert and replace on the same
 * vocabulary the docx uses, so emf/wmf/bmp/ico round-trip with extract instead
 * of tripping over Bun's divergent MIME table (`image/emf` vs `image/x-emf`). */
export function imageFormatForExtension(
	extension: string,
): { extension: string; mimeType: string } | undefined {
	const lower = extension.toLowerCase();
	const canonical = CANONICAL_IMAGE_EXTENSION[lower] ?? lower;
	const mimeType = IMAGE_MIME_BY_EXTENSION[canonical];
	return mimeType ? { extension: canonical, mimeType } : undefined;
}
