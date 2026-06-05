/** Parse a `--pages` spec into a `{first, last}` range. Supports:
 *
 *   "1"       → { first: 1, last: 1 }
 *   "1-3"     → { first: 1, last: 3 }
 *   "3-1"     → error (descending range)
 *   ""        → error (empty)
 *
 * Returns the parsed range, or a string error message. Discontinuous
 * specs (`"1,3,5"`) are deliberately out of scope: keeping the parser
 * single-range matches the CLI's "one render produces a contiguous block"
 * mental model. PDFium *could* render arbitrary pages, but supporting it
 * here would split the output paths and complicate the JSON ack's
 * `pages: []` array semantics. If you need discrete pages, run `docx
 * render` multiple times. */
export function parsePagesSpec(
	spec: string,
): { first: number; last: number } | string {
	const trimmed = spec.trim();
	if (trimmed.length === 0) return "--pages cannot be empty";
	if (trimmed.includes(",")) {
		return '--pages does not yet support discontinuous ranges (e.g., "1,3,5"); pass a single page or contiguous range like "1-3"';
	}
	if (trimmed.includes("-")) {
		const parts = trimmed.split("-");
		if (parts.length !== 2 || !parts[0] || !parts[1]) {
			return `--pages range must be FIRST-LAST, got "${spec}"`;
		}
		const first = Number(parts[0]);
		const last = Number(parts[1]);
		if (!Number.isInteger(first) || first < 1) {
			return `--pages first must be a positive integer, got "${parts[0]}"`;
		}
		if (!Number.isInteger(last) || last < 1) {
			return `--pages last must be a positive integer, got "${parts[1]}"`;
		}
		if (last < first) {
			return `--pages range is descending (${first}-${last}); pass FIRST <= LAST`;
		}
		return { first, last };
	}
	const single = Number(trimmed);
	if (!Number.isInteger(single) || single < 1) {
		return `--pages must be a positive integer or "FIRST-LAST" range, got "${spec}"`;
	}
	return { first: single, last: single };
}
