import type { Document } from "@core";
import { ContentTypesView } from "@core/ast/document/content-types";
import type { XmlNode } from "@core/parser";
import {
	isXmlPartName,
	listPackageParts,
	normalizePartName,
	Raw,
} from "@core/raw";
import {
	EXIT,
	fail,
	openOrFail,
	openPkgOrFail,
	respond,
	SAVE_FLAGS,
	setVerboseAck,
	tryParseArgs,
	writeStdout,
} from "../respond";
import {
	finishMint,
	patchOrFail,
	type RawMutationOperation,
	rawOrFail,
	requireFindWith,
	respondDryRun,
	schemaDiffGateOrFail,
} from "./commit";
import { type RawValues, readXmlSource } from "./xml-source";

/** `docx raw part <verb>` — OPC parts addressed by NAME. Parts are the
 *  package's files, not document content, so they deliberately do NOT ride
 *  the `--at` locator system: locators address what a reader sees (blocks,
 *  sections, relationships an rId names); part names address the container. */
const HELP = `docx raw part — the package's parts, addressed by name

Usage:
  docx raw part list FILE [--json]
  docx raw part get FILE --name PARTNAME [--json | --to-file PATH]
  docx raw part add FILE --name PARTNAME (--xml STR | --xml-file PATH|- | --from PATH)
                    [--content-type CT]
  docx raw part replace FILE --name PARTNAME (--xml STR | --xml-file PATH|-)
  docx raw part edit FILE --name PARTNAME --find STR --with STR

Verbs:
  list      every part with its content type — the discovery path
  get       one part's exact content (binary parts need --to-file)
  add       create a NEW part: XML is gated (well-formed, one root, WML roots
            schema-checked), binary is stored verbatim; --content-type
            registers an Override (required unless the extension already has
            a Default — an OPC part without a content type breaks the package)
  replace   swap a whole XML part: unmodeled parts directly; styles/numbering/
            settings and header/footer parts land through their views; notes/
            comments/document.xml belong to their own verbs. The root tag must
            match (a part's root is its identity); WML roots are schema-gated.
  edit      patch a part in place: literal --find/--with over its exact text
            (all occurrences; zero matches errors), result gated like replace

Common flags: --no-validate (skip the schema gate), -o/--output PATH,
--dry-run, -v/--verbose.

The embedded-object loop:
  docx raw part add report.docx --name word/embeddings/object1.bin --from ole.bin \\
      --content-type application/vnd.openxmlformats-officedocument.oleObject
  docx raw insert report.docx --xml '<Relationship Type="…/oleObject" Target="embeddings/object1.bin"/>'
  docx raw insert report.docx --after p3 --xml '<w:p>…<o:OLEObject r:id="rId7"/>…</w:p>'
`;

const VERBS: Record<string, (args: string[]) => Promise<number>> = {
	list: runPartList,
	get: runPartGet,
	add: runPartAdd,
	replace: runPartReplace,
	edit: runPartEdit,
};

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const handler = VERBS[verb];
	if (!handler) {
		return fail(
			"USAGE",
			`Unknown raw part subcommand: ${verb}`,
			'Run "docx raw part --help" — the verbs are list, get, add, replace, edit.',
		);
	}
	return handler(args.slice(1));
}

// `list`/`get` are queries, but the whole `raw part` noun is registered as a
// mutator in the test harness (it keys off `args[1] === "part"`), so every
// subverb must tolerate the `--verbose` it injects — accepted, then ignored.
const LIST_SPEC = {
	json: { type: "boolean" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

async function runPartList(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, LIST_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);

	const pkg = await openPkgOrFail(filePath);
	if (typeof pkg === "number") return pkg;
	const contentTypes = await ContentTypesView.fromPackage(pkg);
	const parts = listPackageParts(pkg, contentTypes);
	if (values.json) {
		await respond(parts);
		return EXIT.OK;
	}
	await writeStdout(
		`${parts.map((part) => `${part.name}\t${part.contentType}`).join("\n")}\n`,
	);
	return EXIT.OK;
}

const GET_SPEC = {
	name: { type: "string" },
	json: { type: "boolean" },
	"to-file": { type: "string" },
	verbose: { type: "boolean", short: "v" },
	help: { type: "boolean", short: "h" },
} as const;

async function runPartGet(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, GET_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);

	const pkg = await openPkgOrFail(filePath);
	if (typeof pkg === "number") return pkg;
	const name = await resolvePartName(pkg, values, filePath);
	if (typeof name === "number") return name;

	// --to-file extracts raw bytes regardless of XML/binary; stdout/--json need
	// text, so those two forms are the ones the XML check gates.
	const toFile = values["to-file"] as string | undefined;
	if (toFile) {
		const bytes = await pkg.readBytes(name);
		await Bun.write(toFile, bytes);
		await writeStdout(`wrote ${bytes.byteLength} bytes to ${toFile}\n`);
		return EXIT.OK;
	}
	if (!isXmlPartName(name)) {
		return fail(
			"USAGE",
			`"${name}" is a binary part — pass --to-file PATH to extract its bytes`,
		);
	}
	const xml = await pkg.readText(name);
	if (values.json) {
		await respond({ name, xml });
		return EXIT.OK;
	}
	await writeStdout(`${xml}\n`);
	return EXIT.OK;
}

const ADD_SPEC = {
	name: { type: "string" },
	xml: { type: "string" },
	"xml-file": { type: "string" },
	from: { type: "string" },
	"content-type": { type: "string" },
	"no-validate": { type: "boolean" },
	...SAVE_FLAGS,
} as const;

async function runPartAdd(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, ADD_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	setVerboseAck(Boolean(values.verbose));

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	const name = values.name as string | undefined;
	if (!name) return fail("USAGE", "Missing --name PARTNAME", HELP);

	const source = await readPartSource(values);
	if (typeof source === "number") return source;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;

	const raw = new Raw(document);
	const prepared = await rawOrFail(() =>
		raw.preparePartAdd(
			name,
			source,
			values["content-type"] as string | undefined,
		),
	);
	if (typeof prepared === "number") return prepared;

	if (prepared.root) {
		const gate = await partSchemaGate(
			prepared.root,
			undefined,
			Boolean(values["no-validate"]),
		);
		if (gate !== undefined) return gate;
	}

	const warnings = [
		`${prepared.name} is inert until a relationship targets it — mint one with: docx raw insert ${filePath} --xml '<Relationship Type="…" Target="${relsTargetFor(prepared.name)}"/>'`,
	];

	if (values["dry-run"]) {
		return respondDryRun(
			"raw.part-add",
			filePath,
			{
				part: prepared.name,
				...(prepared.contentType ? { contentType: prepared.contentType } : {}),
				bytes:
					typeof prepared.content === "string"
						? Buffer.byteLength(prepared.content)
						: prepared.content.byteLength,
			},
			warnings,
			values.output as string | undefined,
		);
	}

	raw.applyPartAdd(prepared);
	return finishMint({
		document,
		filePath,
		outputPath: values.output as string | undefined,
		operation: "raw.part-add",
		handles: [prepared.name],
		ackFields: {
			part: prepared.name,
			...(prepared.contentType ? { contentType: prepared.contentType } : {}),
		},
		warnings,
	});
}

const REPLACE_SPEC = {
	name: { type: "string" },
	xml: { type: "string" },
	"xml-file": { type: "string" },
	"no-validate": { type: "boolean" },
	...SAVE_FLAGS,
} as const;

async function runPartReplace(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, REPLACE_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	setVerboseAck(Boolean(values.verbose));

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	const xml = await readXmlSource(values);
	if (typeof xml === "number") return xml;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const name = await resolvePartName(document.pkg, values, filePath);
	if (typeof name === "number") return name;

	return replacePartWithXml({
		document,
		filePath,
		name,
		xml,
		values,
		operation: "raw.replace",
	});
}

const EDIT_SPEC = {
	name: { type: "string" },
	find: { type: "string" },
	with: { type: "string" },
	"no-validate": { type: "boolean" },
	...SAVE_FLAGS,
} as const;

async function runPartEdit(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(args, EDIT_SPEC, HELP);
	if (typeof parsed === "number") return parsed;
	const { values, positionals } = parsed;
	setVerboseAck(Boolean(values.verbose));

	const filePath = positionals[0];
	if (!filePath) return fail("USAGE", "Missing FILE", HELP);
	const findWith = await requireFindWith(values, HELP);
	if (typeof findWith === "number") return findWith;

	const document = await openOrFail(filePath);
	if (typeof document === "number") return document;
	const name = await resolvePartName(document.pkg, values, filePath);
	if (typeof name === "number") return name;
	if (!isXmlPartName(name)) {
		return fail(
			"USAGE",
			`"${name}" is a binary part — only XML can be patched`,
		);
	}

	const currentXml = await document.pkg.readText(name);
	const patched = await patchOrFail(
		currentXml,
		findWith.find,
		findWith.replaceWith,
		`docx raw part get ${filePath} --name ${name}`,
	);
	if (typeof patched === "number") return patched;
	return replacePartWithXml({
		document,
		filePath,
		name,
		xml: patched.xml,
		values,
		operation: "raw.edit",
		extraWarnings: [patched.note],
		currentXml,
	});
}

/** The shared replace tail (replace + edit): gates, schema diff, apply, save. */
async function replacePartWithXml(options: {
	document: Document;
	filePath: string;
	name: string;
	xml: string;
	values: RawValues;
	operation: RawMutationOperation;
	extraWarnings?: string[];
	/** The part's current text if the caller already read it (`part edit`),
	 *  so `preparePartReplace` needn't inflate it a second time. */
	currentXml?: string;
}): Promise<number> {
	const { document, filePath, name, xml, values } = options;
	const raw = new Raw(document);
	const prepared = await rawOrFail(() =>
		raw.preparePartReplace(name, xml, options.currentXml),
	);
	if (typeof prepared === "number") return prepared;

	const gate = await partSchemaGate(
		prepared.newRoot,
		prepared.oldRoot,
		Boolean(values["no-validate"]),
	);
	if (gate !== undefined) return gate;

	const warnings = options.extraWarnings ?? [];
	if (values["dry-run"]) {
		return respondDryRun(
			options.operation,
			filePath,
			{ part: prepared.name },
			warnings,
			values.output as string | undefined,
		);
	}

	raw.applyPartReplace(prepared);
	return finishMint({
		document,
		filePath,
		outputPath: values.output as string | undefined,
		operation: options.operation,
		handles: [prepared.name],
		ackFields: { part: prepared.name },
		warnings,
	});
}

/** Validate + existence-check the `--name` argument against the package. */
async function resolvePartName(
	pkg: { hasPart(name: string): boolean },
	values: RawValues,
	filePath: string,
): Promise<string | number> {
	const rawName = values.name as string | undefined;
	if (!rawName) return fail("USAGE", "Missing --name PARTNAME", HELP);
	const name = await rawOrFail(() => normalizePartName(rawName));
	if (typeof name === "number") return name;
	if (!pkg.hasPart(name)) {
		return fail(
			"PART_NOT_FOUND",
			`No part "${name}" in the package`,
			`docx raw part list ${filePath} shows every part.`,
		);
	}
	return name;
}

/** The schema gate for a whole part: WML-rooted parts validate against the
 *  bundled schemas — baseline-diffed against the OLD part on replace (its
 *  pre-existing violations never block), all-errors-are-new on add. Non-WML
 *  roots (theme, customXml, …) pass — no schema covers them; the structural
 *  gates already ran. Validator dynamically imported, as everywhere. */
async function partSchemaGate(
	newRoot: XmlNode,
	oldRoot: XmlNode | undefined,
	noValidate: boolean,
): Promise<number | undefined> {
	if (noValidate) return undefined;
	const { validationXmlFor, WML_ROOT_TAGS } = await import(
		"@core/raw/validate"
	);
	if (!WML_ROOT_TAGS.has(newRoot.tag)) return undefined;
	return schemaDiffGateOrFail(
		validationXmlFor(newRoot),
		() => (oldRoot ? validationXmlFor(oldRoot) : undefined),
		"part",
	);
}

/** How the document rels part would spell a Target for this part name —
 *  relative to word/ when inside it, package-absolute otherwise. */
function relsTargetFor(name: string): string {
	return name.startsWith("word/") ? name.slice("word/".length) : `/${name}`;
}

async function readPartSource(
	values: RawValues,
): Promise<{ xml?: string; bytes?: Uint8Array } | number> {
	const from = values.from as string | undefined;
	const hasXmlFlags =
		values.xml !== undefined || values["xml-file"] !== undefined;
	if (from !== undefined && hasXmlFlags) {
		return fail(
			"USAGE",
			"Pass either --from (bytes) or --xml/--xml-file (XML), not both",
		);
	}
	if (from !== undefined) {
		const source = Bun.file(from);
		if (!(await source.exists())) {
			return fail("FILE_NOT_FOUND", `No such file: ${from}`);
		}
		return { bytes: new Uint8Array(await source.arrayBuffer()) };
	}
	if (!hasXmlFlags) {
		return fail(
			"USAGE",
			"Missing part content: pass --xml STR, --xml-file PATH (- = stdin), or --from PATH for bytes",
			HELP,
		);
	}
	const xml = await readXmlSource(values);
	if (typeof xml === "number") return xml;
	return { xml };
}
