import {
	XmlDocument,
	XmlValidateError,
	XsdValidator,
	xmlRegisterInputProvider,
} from "libxml2-wasm";
import { prefixOf } from "../mc";
import { XmlNode } from "../parser";
import dmlChart from "./schemas/dml-chart.xsd" with { type: "text" };
import dmlChartDrawing from "./schemas/dml-chartDrawing.xsd" with {
	type: "text",
};
import dmlDiagram from "./schemas/dml-diagram.xsd" with { type: "text" };
import dmlLockedCanvas from "./schemas/dml-lockedCanvas.xsd" with {
	type: "text",
};
import dmlMain from "./schemas/dml-main.xsd" with { type: "text" };
import dmlPicture from "./schemas/dml-picture.xsd" with { type: "text" };
import dmlWordprocessingDrawing from "./schemas/dml-wordprocessingDrawing.xsd" with {
	type: "text",
};
import sharedCommonSimpleTypes from "./schemas/shared-commonSimpleTypes.xsd" with {
	type: "text",
};
import sharedCustomXmlSchemaProperties from "./schemas/shared-customXmlSchemaProperties.xsd" with {
	type: "text",
};
import sharedMath from "./schemas/shared-math.xsd" with { type: "text" };
import sharedRelationshipReference from "./schemas/shared-relationshipReference.xsd" with {
	type: "text",
};
import wml from "./schemas/wml.xsd" with { type: "text" };
import xmlXsd from "./schemas/xml.xsd" with { type: "text" };

/** Full-document schema validation (gate 6) — the bundled tight-and-local
 *  answer to "did this raw edit corrupt the file?": libxml2 compiled to WASM
 *  (~1.3 MB) plus the ECMA-376 5th-edition TRANSITIONAL XSD closure for
 *  WordprocessingML (~0.5 MB, from the freely published Part 4 annex; the
 *  `xml.xsd` imports are patched to resolve locally). Word writes transitional
 *  markup, so this is the profile real documents validate against. Measured:
 *  ~55 ms one-time schema compile, single-digit ms per document after.
 *
 *  Real-world markup only validates after MCE preprocessing: Word documents
 *  carry `mc:Ignorable`-declared extension attributes (`w14:paraId` on every
 *  paragraph) that are VALID per OOXML's markup-compatibility rules but
 *  unknown to the base schema. {@link validationXmlFor} therefore strips
 *  ignorable-namespace content from a CLONE before serializing — which also
 *  strips our own `dcx:raw` markers, since `raw` registers `dcx` as ignorable. */
export function validateWmlXml(xml: string): ValidationIssue[] {
	const validator = wmlValidator();
	let parsed: XmlDocument;
	try {
		parsed = XmlDocument.fromString(xml);
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return [{ message: `not well-formed XML: ${message.trim()}` }];
	}
	try {
		validator.validate(parsed);
		return [];
	} catch (error) {
		if (error instanceof XmlValidateError) {
			return error.details.map((detail) => ({
				message: detail.message.trim(),
				line: detail.line,
			}));
		}
		throw error;
	} finally {
		parsed.dispose();
	}
}

export type ValidationIssue = { message: string; line?: number };

/** Root elements wml.xsd declares — a part validates against the bundled
 *  schemas iff its root is one of these. Shared by `docx validate` (which
 *  enumerates parts from the package by root tag) and the raw part gates. */
export const WML_ROOT_TAGS = new Set([
	"w:document",
	"w:styles",
	"w:numbering",
	"w:settings",
	"w:footnotes",
	"w:endnotes",
	"w:comments",
	"w:hdr",
	"w:ftr",
	"w:fonts",
	"w:webSettings",
	"w:glossaryDocument",
]);

/** Serialize a part root for validation: clone, apply MCE preprocessing, emit. */
export function validationXmlFor(root: XmlNode): string {
	const cleaned = root.clone();
	mceStrip(cleaned, new Set());
	return XmlNode.serialize([cleaned]);
}

/** The issues in `candidate` that are NOT in `baseline`, as a multiset diff on
 *  the message text. Real documents can carry pre-existing schema violations
 *  (a third of our own fixtures did), so a raw edit is judged only on the
 *  errors it ADDS. Messages carry element names but not positions, so an
 *  existing error whose line shifts after an insert still cancels out. */
export function diffValidationIssues(
	baseline: ValidationIssue[],
	candidate: ValidationIssue[],
): ValidationIssue[] {
	const seen = new Map<string, number>();
	for (const issue of baseline) {
		seen.set(issue.message, (seen.get(issue.message) ?? 0) + 1);
	}
	const fresh: ValidationIssue[] = [];
	for (const issue of candidate) {
		const remaining = seen.get(issue.message) ?? 0;
		if (remaining > 0) {
			seen.set(issue.message, remaining - 1);
			continue;
		}
		fresh.push(issue);
	}
	return fresh;
}

/** Render issues for an error payload: one per line, capped. */
export function describeIssues(issues: ValidationIssue[], limit = 10): string {
	const lines = issues
		.slice(0, limit)
		.map((issue) =>
			issue.line === undefined
				? issue.message
				: `${issue.message} (line ${issue.line})`,
		);
	const hidden = Math.max(0, issues.length - limit);
	if (hidden > 0) lines.push(`… and ${hidden} more`);
	return lines.join("\n");
}

/** OOXML Markup Compatibility preprocessing (ECMA-376 Part 3), the subset that
 *  matters for validation: drop attributes AND elements whose prefix is listed
 *  in an in-scope `mc:Ignorable`, drop the `mc:*` control attributes
 *  themselves, and resolve `mc:AlternateContent` to its `mc:Fallback` branch
 *  (the branch a consumer that understands none of the choices would take). */
function mceStrip(node: XmlNode, ignorableAbove: Set<string>): void {
	if (node.isText) return;
	let ignorable = ignorableAbove;
	const declared = node.getAttribute("mc:Ignorable");
	if (declared) {
		ignorable = new Set(ignorableAbove);
		for (const prefix of declared.split(/\s+/)) {
			if (prefix) ignorable.add(prefix);
		}
	}
	for (const key of Object.keys(node.attributes)) {
		if (key.startsWith("xmlns")) continue;
		const prefix = prefixOf(key);
		if (prefix === "mc" || (prefix !== undefined && ignorable.has(prefix))) {
			delete node.attributes[key];
		}
	}
	const kept: XmlNode[] = [];
	for (const child of node.children) {
		if (child.isText) {
			kept.push(child);
			continue;
		}
		if (child.tag === "mc:AlternateContent") {
			const fallback = child.findChild("mc:Fallback");
			for (const inner of fallback?.children ?? []) {
				if (!inner.isText) {
					const prefix = prefixOf(inner.tag);
					if (prefix !== undefined && ignorable.has(prefix)) continue;
					mceStrip(inner, ignorable);
				}
				kept.push(inner);
			}
			continue;
		}
		const prefix = prefixOf(child.tag);
		if (prefix !== undefined && ignorable.has(prefix)) continue;
		mceStrip(child, ignorable);
		kept.push(child);
	}
	node.children = kept;
}

/** Compile the schema once per process and keep it (libxml2 objects need
 *  explicit disposal, but a process-lifetime validator is the point here).
 *  Cross-file `xsd:import`s resolve through a virtual input provider over the
 *  bundled schema texts, keyed by basename. */
let compiledValidator: XsdValidator | undefined;

function wmlValidator(): XsdValidator {
	if (compiledValidator) return compiledValidator;
	registerSchemaProvider();
	const schemaDoc = XmlDocument.fromString(wml, { url: "wml.xsd" });
	compiledValidator = XsdValidator.fromDoc(schemaDoc);
	return compiledValidator;
}

const SCHEMA_TEXTS: Record<string, string> = {
	"wml.xsd": wml,
	"dml-chart.xsd": dmlChart,
	"dml-chartDrawing.xsd": dmlChartDrawing,
	"dml-diagram.xsd": dmlDiagram,
	"dml-lockedCanvas.xsd": dmlLockedCanvas,
	"dml-main.xsd": dmlMain,
	"dml-picture.xsd": dmlPicture,
	"dml-wordprocessingDrawing.xsd": dmlWordprocessingDrawing,
	"shared-commonSimpleTypes.xsd": sharedCommonSimpleTypes,
	"shared-customXmlSchemaProperties.xsd": sharedCustomXmlSchemaProperties,
	"shared-math.xsd": sharedMath,
	"shared-relationshipReference.xsd": sharedRelationshipReference,
	"xml.xsd": xmlXsd,
};

function registerSchemaProvider(): void {
	const encoder = new TextEncoder();
	const openReads = new Map<number, { bytes: Uint8Array; offset: number }>();
	let nextHandle = 1;
	xmlRegisterInputProvider({
		match: (filename) => basename(filename) in SCHEMA_TEXTS,
		open: (filename) => {
			const text = SCHEMA_TEXTS[basename(filename)];
			if (text === undefined) return undefined;
			const handle = nextHandle++;
			openReads.set(handle, { bytes: encoder.encode(text), offset: 0 });
			return handle;
		},
		read: (handle, buffer) => {
			const state = openReads.get(handle as number);
			if (!state) return -1;
			const chunk = state.bytes.subarray(
				state.offset,
				state.offset + buffer.byteLength,
			);
			buffer.set(chunk);
			state.offset += chunk.byteLength;
			return chunk.byteLength;
		},
		close: (handle) => openReads.delete(handle as number),
	});
}

function basename(filename: string): string {
	const cleaned = filename.replace(/^file:\/\//, "");
	const slash = cleaned.lastIndexOf("/");
	return slash === -1 ? cleaned : cleaned.slice(slash + 1);
}
