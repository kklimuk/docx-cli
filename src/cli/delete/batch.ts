import {
	type BlockReference,
	LocatorParseError,
	LocatorResolveError,
	TrackChanges,
	type XmlNode,
} from "@core";
import { readJsonlObjects } from "../parse-helpers";
import {
	EXIT,
	fail,
	openOrFail,
	resolveTracked,
	respond,
	respondAck,
} from "../respond";

type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** `docx delete --batch FILE.jsonl`: remove many blocks from one read. Each
 *  JSONL line is `{ "at": LOCATOR }` (a whole-block locator: pN, tN, or a cell
 *  paragraph tN:rRcC:pK). Like `edit --batch`, every locator is resolved to a
 *  LIVE node ref BEFORE anything is removed, so positional ids that shift as
 *  siblings disappear never misfire. Range (pN-pM), section (sN), equation (eqN),
 *  and span (pN:S-E) deletes are done one at a time. */
export async function runDeleteBatch(
	filePath: string,
	batchSource: string,
	values: RawValues,
): Promise<number> {
	if (values.at !== undefined) {
		return fail(
			"USAGE",
			"--batch reads each locator from the JSONL file; don't also pass --at",
			'Put one {"at": LOCATOR} per line.',
		);
	}
	const authorFlag = values.author as string | undefined;
	const trackFlag = Boolean(values.track);
	const outputPath = values.output as string | undefined;
	const dryRun = Boolean(values["dry-run"]);

	let rawEntries: Record<string, unknown>[];
	try {
		rawEntries = await readJsonlObjects(batchSource);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail("USAGE", `Failed to read batch: ${message}`);
	}
	if (rawEntries.length === 0) return fail("USAGE", "Batch file is empty");

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const tracked = resolveTracked(document, trackFlag);

	// Resolve every locator to a live node ref up front (the load-bearing batch
	// invariant) before removing anything.
	const refs: { locator: string; ref: BlockReference }[] = [];
	const seenNodes = new Set<XmlNode>();
	for (let index = 0; index < rawEntries.length; index++) {
		const raw = rawEntries[index];
		if (!raw) continue;
		const at = raw.at;
		if (typeof at !== "string" || at.length === 0) {
			return fail("USAGE", `entry ${index}: "at" is required`);
		}
		if (/^p\d+-p\d+$/.test(at)) {
			return fail(
				"USAGE",
				`entry ${index}: range deletes (${at}) aren't supported in --batch`,
				"Delete a range one at a time with `docx delete --at pN-pM`.",
			);
		}
		if (/^s\d+$/.test(at) || /^eq\d+$/.test(at)) {
			return fail(
				"USAGE",
				`entry ${index}: ${at} isn't supported in --batch`,
				"Delete sections / equations one at a time.",
			);
		}
		if (/:\d+-\d+$/.test(at)) {
			return fail(
				"INVALID_LOCATOR",
				`entry ${index}: delete works on whole blocks, not character spans (${at})`,
				'Use a paragraph/block locator (pN, tN, tN:rRcC:pK). To remove a span\'s text, use `edit --at <span> --text ""`.',
			);
		}
		let blockRef: BlockReference;
		try {
			blockRef = document.body.resolveBlock(at);
		} catch (error) {
			if (error instanceof LocatorResolveError) {
				return fail("BLOCK_NOT_FOUND", `entry ${index}: ${error.message}`);
			}
			if (error instanceof LocatorParseError) {
				return fail("INVALID_LOCATOR", `entry ${index}: ${error.message}`);
			}
			throw error;
		}
		if (tracked && blockRef.node.tag !== "w:p") {
			return fail(
				"TRACKED_CHANGE_CONFLICT",
				`entry ${index}: tracked deletion of a non-paragraph block (${at}) is not supported`,
				"Turn track-changes off, or delete it individually.",
			);
		}
		// Two locators can resolve to the SAME node (e.g. p3 and a cell locator for
		// the same paragraph, or a literal repeat). Deleting it twice is at best a
		// no-op and at worst — under tracking — pushes a second <w:del> into one
		// CT_ParaRPr (schema-invalid: max one) and double-merges on accept. Reject
		// up front, mirroring `edit --batch`'s same-node guard.
		if (seenNodes.has(blockRef.node)) {
			return fail(
				"USAGE",
				`entry ${index}: ${at} resolves to a block already targeted by an earlier entry`,
				"Each --batch delete must address a distinct block.",
			);
		}
		seenNodes.add(blockRef.node);
		refs.push({ locator: at, ref: blockRef });
	}

	if (dryRun) {
		await respond({
			operation: "delete",
			dryRun: true,
			path: filePath,
			batch: refs.map((entry) => ({ locator: entry.locator })),
			...(outputPath ? { output: outputPath } : {}),
		});
		return EXIT.OK;
	}

	for (const { ref } of refs) {
		if (tracked) {
			new TrackChanges(document).applyDeletion(ref.node, authorFlag);
		} else {
			// Find the CURRENT index each time — earlier removals shift siblings,
			// but the held node ref stays valid, so indexOf is always right.
			const index = ref.parent.indexOf(ref.node);
			if (index >= 0) ref.parent.splice(index, 1);
		}
	}

	await document.save(outputPath);
	await respondAck({
		ok: true,
		operation: "delete",
		path: outputPath ?? filePath,
		count: refs.length,
		batch: refs.map((entry) => ({ locator: entry.locator })),
	});
	return EXIT.OK;
}
