/** Normalize the read renderer's locator markers so `docx diff` shows the
 *  changes an agent made, not positional thrash. Applied to BOTH sides before
 *  diffing.
 *
 *  Read puts an inline locator on (almost) every line, and those renumber
 *  wholesale on a structural edit: a single row insert shifts `t2:r2`→`t2:r3`,
 *  `r3`→`r4`, … so a raw diff flags every downstream row as changed. The probe
 *  in the plan measured 17 "changed" lines for a 1-row insert; normalizing
 *  collapses that to the 1 line that actually changed.
 *
 *  Two moves:
 *  (a) strip **bare** locator comments entirely — `<!-- p3 -->`,
 *      `<!-- t2:r5c0:p0 -->`, `<!-- c0 -->`, `<!-- img0 -->`, … (an address the
 *      agent gets from `docx read`, not something it changed);
 *  (b) inside `docx:` annotations, keep the **stable coarse anchor** (table
 *      `tN`, section `sN` — they only renumber when a whole table/section is
 *      added) but drop the **thrashing fine part**: a leading paragraph `pN`
 *      token, and a cell locator's `:rXcY[:pZ]` suffix (`docx:cell t2:r5c0` →
 *      `docx:cell t2`). Formatting/structure attributes (shading, borders,
 *      vAlign, track-changes state) are untouched, so those deltas still diff. */
export function normalizeReadMarkers(md: string): string {
	return (
		md
			// (a) bare locator comment: a locator token (letters + digit, optional
			// `:`-separated cell chain), no `docx:` prefix. Eat the leading space too.
			.replace(/[ \t]*<!--\s*[a-z]+\d[a-z0-9:]*\s*-->/g, "")
			// (b) drop a leading paragraph token in a docx: annotation (pN thrashes
			// on any block insert above): `<!-- docx:p p5 style=… -->` → `… docx:p style=…`
			.replace(/(<!--\s*docx:[a-z-]+) p\d+/g, "$1")
			// … and reduce a leading cell locator to its table anchor:
			// `<!-- docx:cell t2:r5c0 shading=… -->` → `<!-- docx:cell t2 shading=… -->`
			.replace(/(<!--\s*docx:[a-z-]+ t\d+):[a-z0-9:]+/g, "$1")
	);
}
