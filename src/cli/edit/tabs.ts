import { type Document, getPageContentWidthEmu } from "@core";
import type { TabStop } from "@core/blocks";

/** A parsed `--tabs` value, resolved to concrete tab stops once the document is
 *  open (the `right` cure needs the section's content width). `right-margin` is
 *  the headline: it cures the fragile right-edge LEFT tab — which left-aligns the
 *  trailing content from a fixed point so a long value overflows the margin and
 *  wraps (the résumé `San` / `Francisco` split that `read` flags as `docx:layout
 *  warn`) — by replacing the paragraph's tab stops with a single RIGHT tab flush
 *  at the margin. `clear` removes them. Shared by single-shot `edit` and the
 *  `--batch` path (so a `"tabs"` key works in JSONL too). */
export type TabsDirective =
	| { kind: "right-margin" }
	| { kind: "clear" }
	| { kind: "explicit"; stops: TabStop[] };

const TAB_ALIGNS = new Set(["left", "right", "center"]);

const TABS_HINT =
	"Use `right` (a right tab at the margin — cures the wrapping LEFT-tab `docx:layout` warning from `read`), `clear`, or explicit stops like `right@7.5in` / `left@1in,right@7.5in`.";

/** Parse a `--tabs` value into a directive, or a `{ error, hint? }` the caller
 *  turns into its own failure (fail() single-shot, EntryError in --batch).
 *  Pure/sync so both call paths can use it. */
export function parseTabsValue(
	value: string,
): TabsDirective | { error: string; hint: string } {
	const normalized = value.trim().toLowerCase();
	if (normalized === "right" || normalized === "right-margin")
		return { kind: "right-margin" };
	if (normalized === "clear" || normalized === "none" || normalized === "off")
		return { kind: "clear" };

	const stops: TabStop[] = [];
	for (const token of value.split(",")) {
		const match = token.trim().match(/^(left|right|center)@([\d.]+)in$/i);
		const align = (match?.[1] ?? "").toLowerCase();
		const inches = Number(match?.[2] ?? Number.NaN);
		if (!match || !TAB_ALIGNS.has(align) || !Number.isFinite(inches)) {
			return {
				error: `Invalid --tabs value: "${token.trim()}"`,
				hint: TABS_HINT,
			};
		}
		stops.push({ align, pos: Math.round(inches * 1440) });
	}
	return { kind: "explicit", stops };
}

/** Resolve a directive to concrete tab stops (twips). `right-margin` reads the
 *  document's content width so the right tab lands flush at the text margin. */
export function resolveTabsDirective(
	directive: TabsDirective,
	document: Document,
): TabStop[] {
	if (directive.kind === "clear") return [];
	if (directive.kind === "explicit") return directive.stops;
	const marginTwips = Math.round(getPageContentWidthEmu(document) / 635);
	return [{ align: "right", pos: marginTwips }];
}
