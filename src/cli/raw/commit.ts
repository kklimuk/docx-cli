import { type Document, isCellScopedLocator } from "@core";
import { Pkg } from "@core/ast/document/package";
import { XmlNode } from "@core/parser";
import { type FragmentRoots, Raw, RawError } from "@core/raw";
import { requireDocumentRoot } from "@core/raw/namespaces";
import { addAuditComment } from "../audit-comment";
import {
	EXIT,
	fail,
	renderVerifyHint,
	respond,
	respondMinted,
} from "../respond";
import type { RawValues } from "./xml-source";

export type RawMutationOperation = "raw.insert" | "raw.replace" | "raw.edit";

/** Translate a `RawError` thrown by any gate into the CLI error shape; rethrow
 *  anything else. The single home for the RawError→fail mapping every raw verb
 *  shares — batch callers pass a `label` like "entry 3" to prefix messages. */
export async function rawOrFail<T>(
	fn: () => T | Promise<T>,
	label?: string,
): Promise<T | number> {
	try {
		return await fn();
	} catch (error) {
		if (error instanceof RawError) {
			return fail(
				error.code,
				label ? `${label}: ${error.message}` : error.message,
				error.hint,
			);
		}
		throw error;
	}
}

/** The shared tail of every raw block mutation: prove the spliced roots became
 *  addressable, run the baseline-diff schema gate, and only then save. Order is
 *  the point — everything up to the save is in-memory, so ANY failure here
 *  leaves the file byte-identical.
 *
 *  The schema engine (libxml2-wasm + the bundled XSDs) is imported LAZILY: its
 *  module instantiates WASM at load, which `--dry-run`, `--no-validate`,
 *  `--help`, and gate-rejected runs should never pay for. */
export async function commitRawMutation(options: {
	document: Document;
	filePath: string;
	operation: RawMutationOperation;
	insertedNodes: XmlNode[];
	/** Run gate 6 (the full-document schema diff). False for `--no-validate`
	 *  and `--dry-run`. */
	validate: boolean;
	warnings: string[];
	outputPath?: string;
}): Promise<number> {
	const { document } = options;

	const locators = await rawOrFail(() =>
		new Raw(document).proveAddressable(options.insertedNodes),
	);
	if (typeof locators === "number") return locators;

	if (options.validate) {
		const { validationXmlFor } = await import("@core/raw/validate");
		const gate = await schemaDiffGateOrFail(
			validationXmlFor(requireDocumentRoot(document)),
			// The pre-mutation state is still on disk (nothing is written until
			// the gate passes), so rebuild the baseline from the file — and only
			// on the rare candidate-has-issues path, never on the clean happy one.
			() => documentBaselineXml(options.filePath),
			"change",
		);
		if (gate !== undefined) return gate;
	}

	return finishMint({
		document,
		filePath: options.filePath,
		outputPath: options.outputPath,
		operation: options.operation,
		handles: locators,
		ackFields: { locators },
		warnings: options.warnings,
		extraNotes: [renderVerifyHint(options.outputPath ?? options.filePath)],
	});
}

async function documentBaselineXml(filePath: string): Promise<string> {
	const { validationXmlFor } = await import("@core/raw/validate");
	const pkg = await Pkg.open(filePath);
	const tree = await pkg.readPart("word/document.xml");
	const root = tree ? XmlNode.findRoot(tree, "w:document") : undefined;
	if (!root) throw new Error("baseline: document.xml has no <w:document> root");
	return validationXmlFor(root);
}

/** Gate 6 for any WML serialization: validate the candidate, and only if it has
 *  issues at all, diff against the baseline (lazily loaded) so a raw edit is
 *  judged solely on errors it ADDS — a third of real documents carry
 *  pre-existing violations. Shared by the document commit path and the part
 *  gates; `subject` selects the noun the message uses. Returns a failure exit
 *  code, or undefined when the candidate is clean or adds nothing. */
export async function schemaDiffGateOrFail(
	candidateXml: string,
	loadBaselineXml: () => string | undefined | Promise<string | undefined>,
	subject: "change" | "part",
): Promise<number | undefined> {
	const { validateWmlXml, diffValidationIssues, describeIssues } = await import(
		"@core/raw/validate"
	);
	const candidate = validateWmlXml(candidateXml);
	if (candidate.length === 0) return undefined;
	const baselineXml = await loadBaselineXml();
	const baseline = baselineXml === undefined ? [] : validateWmlXml(baselineXml);
	const fresh = diffValidationIssues(baseline, candidate);
	if (fresh.length === 0) return undefined;

	const carrier =
		subject === "part" ? "part would carry" : "change would introduce";
	const fixTarget = subject === "part" ? "XML" : "fragment";
	return fail(
		"VALIDATION_FAILED",
		`The ${carrier} ${fresh.length} new schema error${fresh.length === 1 ? "" : "s"} — nothing was written:\n${describeIssues(fresh)}`,
		`Fix the ${fixTarget} (the errors name the offending elements). If you are certain it is right, --no-validate skips this gate.`,
	);
}

/** Save + mint: the shared success tail for every raw mutation. `handles`
 *  drives the bare-locator lines (one per line); `ackFields` carries the
 *  surface-specific verbose-ack shape (`{locators}` for blocks, `{relationships}`
 *  for the rels part, `{part}` singular for a part) so each keeps its identity. */
export async function finishMint(options: {
	document: Document;
	filePath: string;
	outputPath?: string;
	operation: string;
	handles: string[];
	ackFields: Record<string, unknown>;
	warnings: string[];
	extraNotes?: string[];
}): Promise<number> {
	await options.document.save(options.outputPath);
	const destination = options.outputPath ?? options.filePath;
	await respondMinted(
		options.handles,
		{
			ok: true,
			operation: options.operation,
			path: destination,
			...options.ackFields,
			...(options.warnings.length > 0 ? { warnings: options.warnings } : {}),
		},
		notesFor(options.warnings, options.extraNotes),
	);
	return EXIT.OK;
}

/** The dry-run ack shared by every raw verb: the fixed `{operation, dryRun,
 *  path}` head plus verb-specific `extra` keys, with `warnings`/`output`
 *  appended only when present. */
export async function respondDryRun(
	operation: string,
	path: string,
	extra: Record<string, unknown>,
	warnings: string[],
	outputPath: string | undefined,
): Promise<number> {
	await respond({
		operation,
		dryRun: true,
		path,
		...extra,
		...(warnings.length > 0 ? { warnings } : {}),
		...(outputPath ? { output: outputPath } : {}),
	});
	return EXIT.OK;
}

/** Format warnings (and any extra notes) for a `respondMinted` text tail, or
 *  undefined when there's nothing to say. */
export function notesFor(
	warnings: string[],
	extra: string[] = [],
): string | undefined {
	const lines = [...warnings.map((warning) => `note: ${warning}`), ...extra];
	return lines.length > 0 ? lines.join("\n") : undefined;
}

/** The `FragmentRoots` for replacing an existing block: an `sN` target takes
 *  exactly one `<w:sectPr>`, anything else takes body blocks. */
export function replaceContextFor(node: XmlNode): FragmentRoots {
	return node.tag === "w:sectPr" ? "sectPr" : "blocks";
}

/** The shared `--find`/`--with` validation for `raw edit` and `raw part edit`:
 *  `--find` is required and non-empty; `--with` is required but MAY be empty
 *  (that's how you delete the matched text). */
export async function requireFindWith(
	values: RawValues,
	help: string,
): Promise<{ find: string; replaceWith: string } | number> {
	const find = values.find as string | undefined;
	if (find === undefined || find === "") {
		return fail(
			"USAGE",
			"Missing --find STR (the literal text to patch)",
			help,
		);
	}
	const replaceWith = values.with as string | undefined;
	if (replaceWith === undefined) {
		return fail(
			"USAGE",
			"Missing --with STR (pass --with '' to delete the matched text)",
			help,
		);
	}
	return { find, replaceWith };
}

/** Literal find/replace over a target's serialization — the shared core of
 *  `raw edit` and `raw part edit`. All occurrences; zero is an error — per
 *  the no-silent-no-op invariant, a mutation that changes nothing must not
 *  exit 0. `readBack` is the exact command that prints the text being
 *  searched (the recovery path the hint teaches). */
export async function patchOrFail(
	current: string,
	find: string,
	replaceWith: string,
	readBack: string,
): Promise<{ xml: string; note: string } | number> {
	const occurrences = current.split(find).length - 1;
	if (occurrences === 0) {
		return fail(
			"MATCH_NOT_FOUND",
			"--find matched nothing in the target's XML — nothing was written",
			`${readBack} prints the exact XML being searched; match it byte for byte (attribute order and quoting as printed).`,
		);
	}
	return {
		xml: current.split(find).join(replaceWith),
		note: `patched ${occurrences} occurrence${occurrences === 1 ? "" : "s"} of --find`,
	};
}

export function failTrackRejected(): Promise<number> {
	return fail(
		"TRACKED_CHANGE_CONFLICT",
		"raw XML cannot be tracked — no honest tracked-change wrapper exists for arbitrary OOXML",
		"Drop --track: the change applies untracked (document tracking stays on for other edits, and an audit comment marks the spot). The modeled verbs (edit/insert/tables …) track natively.",
	);
}

/** Raw changes are never tracked — OOXML has no honest tracked-change wrapper
 *  for arbitrary XML. When the document toggle is ON we follow the settled
 *  audit-comment precedent (hyperlinks/images/table merges): apply untracked,
 *  anchor a `[docx-cli]` comment to the first inserted paragraph (silent when
 *  the fragment has none), and say so in the ack. */
export function noteUntrackedRawChange(
	document: Document,
	insertedNodes: XmlNode[],
	operation: RawMutationOperation,
	authorFlag: string | undefined,
	warnings: string[],
): void {
	if (!document.isTrackChangesEnabled()) return;
	warnings.push(
		"document tracking is ON, but raw changes are NOT tracked (no tracked-change wrapper exists for arbitrary XML)",
	);
	const paragraph = insertedNodes.find((node) => node.tag === "w:p");
	if (!paragraph) return;
	const described = {
		"raw.insert": "raw XML inserted",
		"raw.replace": "block replaced with raw XML",
		"raw.edit": "block patched with raw XML",
	}[operation];
	addAuditComment(
		document,
		{ kind: "span", paragraph, span: { start: 0, end: 0 } },
		`${described} — this change is not tracked`,
		authorFlag,
	);
}

/** The whole replace-a-block tail — gates, dry-run, splice, commit — shared by
 *  `raw replace` (verbatim replacement) and `raw edit` (the same flow fed a
 *  find/with-patched serialization). */
export async function replaceBlockWithXml(options: {
	document: Document;
	filePath: string;
	reference: { node: XmlNode; parent: XmlNode[] };
	locator: string;
	xml: string;
	values: RawValues;
	operation: RawMutationOperation;
	extraWarnings?: string[];
}): Promise<number> {
	const { document, filePath, reference, locator, xml, values } = options;
	const raw = new Raw(document);
	const prepared = await rawOrFail(() =>
		raw.prepareFragment(xml, replaceContextFor(reference.node)),
	);
	if (typeof prepared === "number") return prepared;

	const warnings = [...(options.extraWarnings ?? []), ...prepared.warnings];
	if (values["dry-run"]) {
		return respondDryRun(
			options.operation,
			filePath,
			{ locator, roots: prepared.nodes.map((node) => node.tag) },
			warnings,
			values.output as string | undefined,
		);
	}

	const spliced = await rawOrFail(() =>
		raw.spliceBlocks(reference, "replace", prepared, {
			cellScoped: isCellScopedLocator(locator),
		}),
	);
	if (typeof spliced === "number") return spliced;

	warnings.push(...spliced);
	noteUntrackedRawChange(
		document,
		prepared.nodes,
		options.operation,
		values.author as string | undefined,
		warnings,
	);

	return commitRawMutation({
		document,
		filePath,
		operation: options.operation,
		insertedNodes: prepared.nodes,
		validate: !values["no-validate"],
		warnings,
		outputPath: values.output as string | undefined,
	});
}
