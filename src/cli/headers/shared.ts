import {
	type Document,
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

/** Resolve the target section `<w:sectPr>` nodes for a header/footer operation:
 *  a single `--at sN` section, or EVERY section (document-wide) when `--at` is
 *  omitted. A document always has at least the trailing section, so the
 *  document-wide list is never empty. Returns a fail() exit code on a bad locator. */
export async function resolveTargetSectPrs(
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
 *  — there's no separate odd part in OOXML. */
export function resolveMarginalType(values: {
	type?: unknown;
	"first-page"?: unknown;
	even?: unknown;
	odd?: unknown;
}): MarginalType | TypeError {
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
	return sources[0] ?? "default";
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
