import {
	type BlockReference,
	isCellScopedLocator,
	isRelationshipLocator,
} from "@core";
import type { XmlNode } from "@core/parser";
import { isRelationshipFragment, type PreparedFragment, Raw } from "@core/raw";

import { parseTargetPlacement, resolvePlacement } from "../insert/place";
import { readJsonlObjects } from "../parse-helpers";
import { fail, openOrFail, resolveBlockOrFail } from "../respond";
import {
	commitRawMutation,
	noteUntrackedRawChange,
	rawOrFail,
	replaceContextFor,
	respondDryRun,
} from "./commit";
import { type RawValues, readXmlSource } from "./xml-source";

/** `docx raw insert --batch FILE.jsonl`: many raw inserts from one read. Each
 *  line is `{"after"|"before": LOCATOR, "xml"|"xmlFile": …}`. All locators
 *  address the document AS READ (the batch invariant): every anchor resolves
 *  to a live node ref and every fragment runs the full gate pipeline BEFORE
 *  any splice — one shared `Raw` so the reference audit walks the document
 *  once and minted ids can't collide across entries — then the splices land
 *  in entry order (stacked after-inserts chain off the previous entry's last
 *  node). One reread, one schema gate, one save. */
export async function runRawInsertBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const rejected = await rejectSingleShotFlags(values, [
		"xml",
		"xml-file",
		"after",
		"before",
		"at-start",
		"at-end",
	]);
	if (rejected !== undefined) return rejected;

	const entries = await readEntries(batchSource);
	if (typeof entries === "number") return entries;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const raw = new Raw(document);

	// Build phase: resolve every anchor + prepare every fragment, zero mutation.
	const planned: {
		anchor: {
			blockRef: BlockReference;
			mode: "after" | "before";
			locator: string;
		};
		prepared: PreparedFragment;
	}[] = [];
	const warnings: string[] = [];
	for (let index = 0; index < entries.length; index++) {
		const entry = entries[index];
		if (!entry) continue;
		const placement = await parseBatchPlacement(entry, index);
		if (typeof placement === "number") return placement;
		const xml = await readEntryXml(entry, index);
		if (typeof xml === "number") return xml;
		const anchor = await resolvePlacement(document, placement);
		if (typeof anchor === "number") return anchor;
		const prepared = await rawOrFail(
			() => raw.prepareFragment(xml, "blocks"),
			`entry ${index}`,
		);
		if (typeof prepared === "number") return prepared;
		warnings.push(
			...prepared.warnings.map((warning) => `entry ${index}: ${warning}`),
		);
		planned.push({ anchor, prepared });
	}

	if (values["dry-run"]) {
		return respondDryRun(
			"raw.insert",
			filePath,
			{ entries: planned.length },
			warnings,
			values.output as string | undefined,
		);
	}

	// Splice phase: stacked after-inserts chain off the previous entry's last
	// spliced node, so entry order is preserved without offset bookkeeping.
	const lastAfter = new Map<XmlNode, XmlNode>();
	const insertedNodes: XmlNode[] = [];
	for (let index = 0; index < planned.length; index++) {
		const plan = planned[index];
		if (!plan) continue;
		const { anchor, prepared } = plan;
		const target =
			anchor.mode === "after"
				? {
						node: lastAfter.get(anchor.blockRef.node) ?? anchor.blockRef.node,
						parent: anchor.blockRef.parent,
					}
				: anchor.blockRef;
		const spliced = await rawOrFail(
			() =>
				raw.spliceBlocks(target, anchor.mode, prepared, {
					cellScoped: isCellScopedLocator(anchor.locator),
				}),
			`entry ${index}`,
		);
		if (typeof spliced === "number") return spliced;
		warnings.push(...spliced.map((warning) => `entry ${index}: ${warning}`));
		if (anchor.mode === "after") {
			const last = prepared.nodes[prepared.nodes.length - 1];
			if (last) lastAfter.set(anchor.blockRef.node, last);
		}
		insertedNodes.push(...prepared.nodes);
	}

	noteUntrackedRawChange(
		document,
		insertedNodes,
		"raw.insert",
		values.author as string | undefined,
		warnings,
	);
	return commitRawMutation({
		document,
		filePath,
		operation: "raw.insert",
		insertedNodes,
		validate: !values["no-validate"],
		warnings,
		outputPath: values.output as string | undefined,
	});
}

/** `docx raw replace --batch FILE.jsonl`: each line is `{"at": LOCATOR,
 *  "xml"|"xmlFile": …}`. Every target resolves to a live ref before any
 *  splice (the delete-batch pattern), so sibling shifts never misfire. */
export async function runRawReplaceBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	const rejected = await rejectSingleShotFlags(values, [
		"xml",
		"xml-file",
		"at",
	]);
	if (rejected !== undefined) return rejected;

	const entries = await readEntries(batchSource);
	if (typeof entries === "number") return entries;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const raw = new Raw(document);

	const planned: {
		locator: string;
		reference: BlockReference;
		prepared: PreparedFragment;
	}[] = [];
	const warnings: string[] = [];
	const seenNodes = new Set<XmlNode>();
	for (let index = 0; index < entries.length; index++) {
		const entry = entries[index];
		if (!entry) continue;
		const locator = entry.at;
		if (typeof locator !== "string" || locator.length === 0) {
			return fail("USAGE", `entry ${index}: "at" is required`);
		}
		if (isRelationshipLocator(locator)) {
			return fail(
				"USAGE",
				`entry ${index}: relationship targets (${locator}) aren't batchable — rIds never shift, so run one raw replace per relationship`,
			);
		}
		const xml = await readEntryXml(entry, index);
		if (typeof xml === "number") return xml;
		const reference = await resolveBlockOrFail(document, locator);
		if (typeof reference === "number") return reference;
		if (seenNodes.has(reference.node)) {
			return fail(
				"USAGE",
				`entry ${index}: ${locator} resolves to a block already targeted by an earlier entry`,
				"Each --batch replace must address a distinct block.",
			);
		}
		seenNodes.add(reference.node);
		const prepared = await rawOrFail(
			() => raw.prepareFragment(xml, replaceContextFor(reference.node)),
			`entry ${index}`,
		);
		if (typeof prepared === "number") return prepared;
		warnings.push(
			...prepared.warnings.map((warning) => `entry ${index}: ${warning}`),
		);
		planned.push({ locator, reference, prepared });
	}

	if (values["dry-run"]) {
		return respondDryRun(
			"raw.replace",
			filePath,
			{ entries: planned.length },
			warnings,
			values.output as string | undefined,
		);
	}

	const insertedNodes: XmlNode[] = [];
	for (let index = 0; index < planned.length; index++) {
		const plan = planned[index];
		if (!plan) continue;
		const { locator, reference, prepared } = plan;
		const spliced = await rawOrFail(
			() =>
				raw.spliceBlocks(reference, "replace", prepared, {
					cellScoped: isCellScopedLocator(locator),
				}),
			`entry ${index}`,
		);
		if (typeof spliced === "number") return spliced;
		warnings.push(...spliced.map((warning) => `entry ${index}: ${warning}`));
		insertedNodes.push(...prepared.nodes);
	}

	noteUntrackedRawChange(
		document,
		insertedNodes,
		"raw.replace",
		values.author as string | undefined,
		warnings,
	);
	return commitRawMutation({
		document,
		filePath,
		operation: "raw.replace",
		insertedNodes,
		validate: !values["no-validate"],
		warnings,
		outputPath: values.output as string | undefined,
	});
}

async function rejectSingleShotFlags(
	values: RawValues,
	flags: string[],
): Promise<number | undefined> {
	const conflicting = flags.find((flag) => values[flag] !== undefined);
	if (conflicting === undefined) return undefined;
	return fail(
		"USAGE",
		`--batch reads everything from the JSONL entries; don't also pass --${conflicting}`,
	);
}

async function readEntries(
	batchSource: string,
): Promise<Record<string, unknown>[] | number> {
	let entries: Record<string, unknown>[];
	try {
		entries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (entries.length === 0) return fail("USAGE", "Batch file is empty");
	return entries;
}

/** Coerce a JSONL entry's placement keys onto the shared flag parser
 *  (`parseTargetPlacement`), then reject the boundary forms — the pattern
 *  `insert/batch.ts` established. */
async function parseBatchPlacement(
	entry: Record<string, unknown>,
	index: number,
): Promise<{ mode: "after" | "before"; locator: string } | number> {
	const placement = await parseTargetPlacement({
		after: asString(entry.after),
		before: asString(entry.before),
		"at-start": Boolean(entry["at-start"]),
		"at-end": Boolean(entry["at-end"]),
	});
	if (typeof placement === "number") {
		return fail(
			"USAGE",
			`entry ${index}: pass exactly one of "after"/"before"`,
		);
	}
	if ("boundary" in placement) {
		return fail(
			"USAGE",
			`entry ${index}: --at-start/--at-end aren't batchable — use "after"/"before" with a locator`,
		);
	}
	return placement;
}

/** Resolve an entry's `xml`/`xmlFile` through the shared source reader (no
 *  stdin — `--batch -` already owns it), rejecting `<Relationship>` fragments
 *  up front with the single-shot pointer. */
async function readEntryXml(
	entry: Record<string, unknown>,
	index: number,
): Promise<string | number> {
	const xml = await readXmlSource(
		{ xml: asString(entry.xml), "xml-file": asString(entry.xmlFile) },
		{ label: `entry ${index}`, stdin: false },
	);
	if (typeof xml === "number") return xml;
	if (isRelationshipFragment(xml)) {
		return fail(
			"USAGE",
			`entry ${index}: <Relationship> fragments aren't batchable — rIds never shift, so run one raw insert per relationship`,
		);
	}
	return xml;
}

function asString(value: unknown): string | undefined {
	return typeof value === "string" ? value : undefined;
}
