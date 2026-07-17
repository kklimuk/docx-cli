import {
	type Document,
	findMarginalRef,
	isMarginalLocator,
	isMarginalType,
	type MarginalKind,
	type MarginalType,
	type XmlNode,
} from "@core";
import { fail, resolveBlockOrFail } from "../respond";

/** The CLI noun for a kind — `headers` / `footers`. */
export function marginalNoun(kind: MarginalKind): "headers" | "footers" {
	return kind === "header" ? "headers" : "footers";
}

/** Resolve WHICH sections + WHICH placement type a `set`/`clear` targets, from
 *  the `--at` locator and the placement flags together:
 *   - `--at hdrN`/`ftrN` — a specific marginal: its owning section, and its type
 *     is FIXED by the id, so an explicit `--type`/`--first-page`/… is rejected as
 *     redundant (and a kind mismatch, e.g. `headers … --at ftr0`, is rejected).
 *   - `--at sN` — that one section, with the placement from the flags.
 *   - omitted — EVERY section (document-wide), with the placement from the flags.
 *  This is what lets the id `list`/`read` report double as the `--at` handle. */
export async function resolveMarginalScope(
	document: Document,
	atLocator: string | undefined,
	kind: MarginalKind,
	flagType: MarginalType,
	hasExplicitType: boolean,
): Promise<{ sectPrs: XmlNode[]; type: MarginalType } | number> {
	if (atLocator !== undefined && isMarginalLocator(atLocator)) {
		const ref = findMarginalRef(document, atLocator);
		if (!ref) {
			return fail(
				"BLOCK_NOT_FOUND",
				`No ${atLocator} in the document`,
				`docx ${marginalNoun(kind)} list prints every ${kind} with its id.`,
			);
		}
		if (ref.kind !== kind) {
			return fail(
				"INVALID_LOCATOR",
				`${atLocator} is a ${ref.kind}, not a ${kind}`,
				`Use "docx ${marginalNoun(ref.kind)}" for that one, or pass a ${kind} id.`,
			);
		}
		if (hasExplicitType) {
			return fail(
				"USAGE",
				`${atLocator} already fixes the placement (${ref.type}) — drop --type / --first-page / --even / --odd`,
			);
		}
		return { sectPrs: [ref.sectPr], type: ref.type };
	}
	const sectPrs = await resolveTargetSectPrs(document, atLocator);
	if (typeof sectPrs === "number") return sectPrs;
	return { sectPrs, type: flagType };
}

/** Resolve the target section `<w:sectPr>` nodes for a header/footer operation:
 *  a single `--at sN` section, or EVERY section (document-wide) when `--at` is
 *  omitted. A document always has at least the trailing section, so the
 *  document-wide list is never empty. Returns a fail() exit code on a bad locator. */
async function resolveTargetSectPrs(
	document: Document,
	atLocator: string | undefined,
): Promise<XmlNode[] | number> {
	if (atLocator !== undefined) {
		const reference = await resolveBlockOrFail(document, atLocator);
		if (typeof reference === "number") return reference;
		if (reference.node.tag !== "w:sectPr") {
			return fail(
				"INVALID_LOCATOR",
				`--at must be a section locator (sN); ${atLocator} is not a section`,
				"Run `docx read` to see section ids (sN), or omit --at to apply to the whole document.",
			);
		}
		return [reference.node];
	}
	const sectPrs: XmlNode[] = [];
	for (const block of document.body.blocks) {
		if (block.type !== "sectionBreak") continue;
		const node = document.body.blockReferences.get(block.id)?.node;
		if (node) sectPrs.push(node);
	}
	return sectPrs;
}

type TypeError = { error: string; hint?: string };

/** Resolve the placement type from `--type` / `--first-page` / `--even` / `--odd`
 *  (at most one source; defaults to `default`). `--odd` is an alias for `default`
 *  — there's no separate odd part in OOXML. `explicit` reports whether ANY of
 *  those flags was passed, so a marginal-id `--at` can reject the redundancy —
 *  one scan of the flag set feeds both the type and the explicit bit. */
export function resolveMarginalType(values: {
	type?: unknown;
	"first-page"?: unknown;
	even?: unknown;
	odd?: unknown;
}): { type: MarginalType; explicit: boolean } | TypeError {
	const sources: MarginalType[] = [];
	if (values["first-page"]) sources.push("first");
	if (values.even) sources.push("even");
	if (values.odd) sources.push("default");
	const typeRaw = values.type as string | undefined;
	if (typeRaw !== undefined) {
		if (!isMarginalType(typeRaw)) {
			return {
				error: `Invalid --type: ${typeRaw}`,
				hint: "Valid: default, first, even (or use --first-page / --even / --odd).",
			};
		}
		sources.push(typeRaw);
	}
	if (sources.length > 1) {
		return {
			error: "Pass at most one of --type / --first-page / --even / --odd",
		};
	}
	return { type: sources[0] ?? "default", explicit: sources.length > 0 };
}

export function isTypeError(value: unknown): value is TypeError {
	return typeof value === "object" && value !== null && "error" in value;
}

/** The shared dispatcher help, parameterized by noun. */
export function dispatcherHelp(kind: MarginalKind): string {
	const noun = marginalNoun(kind);
	return `docx ${noun} — author page ${noun}

Usage:
  docx ${noun} <verb> FILE [options]

Verbs:
  set      Set (create or replace) a ${kind} — text, page numbers, date, fields
  list     Print existing ${noun} as JSON
  clear    Remove a ${kind} reference from a section (or the whole document)

Run "docx ${noun} <verb> --help" for verb-specific help.
`;
}
