import type { Document } from "@core";
import { Raw } from "@core/raw";
import { fail, openOrFail } from "../respond";
import {
	finishMint,
	type RawMutationOperation,
	rawOrFail,
	respondDryRun,
} from "./commit";
import type { RawValues } from "./xml-source";

/** The relationships side of `raw insert`/`raw replace`. A fragment whose
 *  roots are `<Relationship>` routes here by shape (see `isRelationshipFragment`);
 *  a `replace --at rIdN` routes here by locator. Relationships are an
 *  unordered set keyed by Id, so there are no placement flags, no batch (rIds
 *  never shift, so nothing goes stale between calls), no `dcx:raw` marker
 *  (the rels part has no MCE), no tracking note (a relationship alone changes
 *  no content), and no full-document schema gate (the attribute gates ARE the
 *  rels schema — Id/Type/Target/TargetMode is all of it). */
export async function runRelationshipInsert(
	filePath: string,
	xml: string,
	values: RawValues,
): Promise<number> {
	const placementFlag = ["after", "before", "at-start", "at-end"].find(
		(flag) => values[flag] !== undefined,
	);
	if (placementFlag) {
		return fail(
			"USAGE",
			`--${placementFlag} does not apply to a <Relationship> fragment — the rels part is an unordered set`,
			"Drop the placement flag; the relationship is keyed by its Id, not a position.",
		);
	}

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const raw = new Raw(document);
	const prepared = await rawOrFail(() =>
		raw.prepareRelationships(xml, { kind: "insert" }),
	);
	if (typeof prepared === "number") return prepared;

	const warnings: string[] = [];
	for (const id of prepared.ids) {
		if (!document.relationships.isReferenced(id, document.documentTree)) {
			warnings.push(
				`${id} is not yet referenced — pair it with body XML (r:id="${id}" / r:embed="${id}") via raw insert; an unreferenced relationship is harmless`,
			);
		}
	}

	if (values["dry-run"]) {
		return respondDryRun(
			"raw.insert",
			filePath,
			{ relationships: prepared.ids },
			warnings,
			values.output as string | undefined,
		);
	}

	raw.applyRelationships(prepared);
	return saveRelationships(
		document,
		filePath,
		"raw.insert",
		prepared.ids,
		warnings,
		values,
	);
}

export async function runRelationshipReplace(
	filePath: string,
	rId: string,
	xml: string,
	values: RawValues,
): Promise<number> {
	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	return replaceRelationshipWithXml({
		document,
		filePath,
		rId,
		xml,
		values,
		operation: "raw.replace",
	});
}

/** The replace-a-relationship tail, shared by `raw replace --at rIdN` and
 *  `raw edit --at rIdN` (which feeds it a find/with-patched serialization). */
export async function replaceRelationshipWithXml(options: {
	document: Document;
	filePath: string;
	rId: string;
	xml: string;
	values: RawValues;
	operation: RawMutationOperation;
	extraWarnings?: string[];
}): Promise<number> {
	const { document, filePath, rId, xml, values } = options;
	if (!document.relationships.findByRid(rId)) {
		return failRelationshipNotFound(filePath, rId);
	}

	const raw = new Raw(document);
	const prepared = await rawOrFail(() =>
		raw.prepareRelationships(xml, { kind: "replace", rId }),
	);
	if (typeof prepared === "number") return prepared;
	const warnings = [...(options.extraWarnings ?? [])];

	if (values["dry-run"]) {
		return respondDryRun(
			options.operation,
			filePath,
			{ relationships: prepared.ids },
			warnings,
			values.output as string | undefined,
		);
	}

	raw.applyRelationships(prepared);
	return saveRelationships(
		document,
		filePath,
		options.operation,
		prepared.ids,
		warnings,
		values,
	);
}

/** The RELATIONSHIP_NOT_FOUND failure (message + rels-listing hint), one place
 *  for get/replace/edit so the wording can't drift across the three surfaces. */
export function failRelationshipNotFound(
	filePath: string,
	rId: string,
): Promise<number> {
	return fail(
		"RELATIONSHIP_NOT_FOUND",
		`No relationship "${rId}" in the document`,
		`docx raw get ${filePath} --at rels lists every relationship with its id.`,
	);
}

function saveRelationships(
	document: Document,
	filePath: string,
	operation: RawMutationOperation,
	ids: string[],
	warnings: string[],
	values: RawValues,
): Promise<number> {
	return finishMint({
		document,
		filePath,
		outputPath: values.output as string | undefined,
		operation,
		handles: ids,
		ackFields: { relationships: ids },
		warnings,
	});
}
