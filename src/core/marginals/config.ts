/** Headers and footers share their entire mechanics — separate OPC parts
 *  (`word/header1.xml` / `word/footer1.xml`), but identical XML shape
 *  (`<w:hdr>` / `<w:ftr>` wrapping body paragraphs), the same per-section
 *  `<w:headerReference>` / `<w:footerReference>` wiring inside `<w:sectPr>`, the
 *  same relationship + content-type registration, and the same three placement
 *  types. We call the shared abstraction a "marginal" (page-margin content — the
 *  header/footer analog of `Note` for footnote/endnote) and parameterize on
 *  `MarginalKind` everywhere instead of duplicating two near-identical modules. */
export function marginalConfig(kind: MarginalKind): MarginalConfig {
	return kind === "header" ? HEADER_CONFIG : FOOTER_CONFIG;
}

export type MarginalKind = "header" | "footer";

/** The three OOXML `w:type` placement values for a header/footer reference
 *  (CT_HdrFtrRef, ECMA-376 §17.6.10). **There is no `odd`** — when
 *  `<w:evenAndOddHeaders/>` is set, the `default` marginal IS the odd-page one
 *  (Word's UI relabels it "Odd Page Header" in that mode). So `--odd` is a CLI
 *  alias for `default`, not a fourth type. */
export type MarginalType = "default" | "first" | "even";

export const MARGINAL_TYPES: readonly MarginalType[] = [
	"default",
	"first",
	"even",
];

export function isMarginalType(value: string): value is MarginalType {
	return (MARGINAL_TYPES as readonly string[]).includes(value);
}

export type MarginalConfig = {
	kind: MarginalKind;
	/** AST / locator prefix: `hdr` / `ftr`. */
	locatorPrefix: "hdr" | "ftr";
	/** Part-name stem: `header` / `footer` (→ `word/header1.xml`). */
	partPrefix: "header" | "footer";
	/** Part root element: `w:hdr` / `w:ftr`. */
	rootTag: "w:hdr" | "w:ftr";
	/** The `<w:sectPr>` child that points at the part. */
	referenceTag: "w:headerReference" | "w:footerReference";
	relationshipType: string;
	contentType: string;
};

const HEADER_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const FOOTER_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
const HEADER_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

const HEADER_CONFIG: MarginalConfig = {
	kind: "header",
	locatorPrefix: "hdr",
	partPrefix: "header",
	rootTag: "w:hdr",
	referenceTag: "w:headerReference",
	relationshipType: HEADER_RELATIONSHIP_TYPE,
	contentType: HEADER_CONTENT_TYPE,
};

const FOOTER_CONFIG: MarginalConfig = {
	kind: "footer",
	locatorPrefix: "ftr",
	partPrefix: "footer",
	rootTag: "w:ftr",
	referenceTag: "w:footerReference",
	relationshipType: FOOTER_RELATIONSHIP_TYPE,
	contentType: FOOTER_CONTENT_TYPE,
};

/** True for a `word/header1.xml` / `word/footer2.xml` … part path. Used by
 *  `MarginalsView.fromPackage` to find every marginal part in the zip. */
export function isMarginalPartName(name: string): boolean {
	return /^word\/(?:header|footer)\d+\.xml$/.test(name);
}

/** Resolve a relationship `Target` (e.g. `header1.xml`, or `/word/header1.xml`)
 *  to its full part name (`word/header1.xml`). */
export function marginalPartNameFromTarget(target: string): string {
	return target.startsWith("/") ? target.slice(1) : `word/${target}`;
}
