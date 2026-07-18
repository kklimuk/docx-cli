# src/core/raw — the raw-OOXML escape hatch + schema validation

`Raw` is a cross-cutting lens (`new Raw(document).prepareFragment(xml, { allow })`)
backing `docx raw get/insert/replace`. Its job is to make arbitrary user XML AS
SAFE as a modeled verb: `prepareFragment` runs the gate pipeline (order in
[index.ts](index.ts)'s doc block) and throws `RawError` on any reject —
`INVALID_XML` maps to exit 2, and the CLI writes NOTHING on failure because all
mutation is in-memory until `cli/raw/commit.ts` saves after the last gate.

## The pieces

- [index.ts](index.ts) — `Raw` (gates 1–5 entry), plus the domain half of the
  mutation: `spliceBlocks` (live-ref splice + the OOXML post-splice cures — a
  cell must end with a `<w:p>`, adjacent tables merge, an inline-sectPr replace
  drops a section boundary) and `proveAddressable` (reread + write-read-loop
  assert, run before any save). One `Raw` instance per command: it lazily
  builds ONE reference-audit context (drawing/bookmark id sets + cursors)
  shared across every fragment it prepares, so a batch neither re-walks the
  document per entry nor double-mints a "fresh" id. `prepareFragment` takes a
  bare `FragmentRoots` union (`"blocks" | "sectPr" | "relationships"`), not a
  context object.
- [error.ts](error.ts) — `RawError` + `RawErrorCode`. A LEAF so every gate
  module shares it without importing the `Raw` composition root back (index.ts
  re-exports it); the barrel is the composition root, not a dependency of its
  own dependencies.
- [parse.ts](parse.ts) — gates 1–2.5 (`assertWellFormed`, `parseRoots`,
  `stripInterElementWhitespace`), the parsing pieces every raw surface (blocks,
  relationships, parts) runs before anything domain-specific — also a leaf, so
  `relationships.ts`/`parts.ts` reach them without the barrel cycle. Gate 1
  wraps the fragment in a synthetic `<end-of-fragment>` element because
  `XMLValidator` rejects multi-root documents AND our parser silently DROPS
  top-level text — parsing inside the wrapper is what makes stray text visible
  to gate 2.
- [element-order.ts](element-order.ts) — child-order checks. Reuses
  `PPR_CHILD_ORDER`/`RPR_CHILD_ORDER` (blocks.tsx) and `SECTPR_CHILD_ORDER`
  (sections.tsx) — the SAME tables the emitters splice by, exported rather
  than copied so they can't drift — plus the leading-props rules (`w:pPr`
  first in `w:p`, …). Containers without a table pass — the XSD layer catches
  those.
- [parts.ts](parts.ts) — OPC parts. Parts are NOT locators (locators address
  document content; part names address the container), so the CLI surface is
  its own noun — `docx raw part list|get|add|replace|edit FILE --name NAME`
  (`cli/raw/part.ts`; a `--at part:…` attempt on get/edit/replace gets a
  redirect via `rejectPartAsLocator`). `list`/`get` read anything (binary
  bytes via `--to-file`); `add` creates a new part (XML gated
  well-formed/one-root, WML roots schema-checked CLI-side; binary stored
  verbatim; the content type must come from `--content-type` or an existing
  extension Default — an OPC part without one makes Word reject the
  package). Whole-part replace/edit route by `partReplaceRoute`: unmodeled
  parts go straight to `Pkg` (stored VERBATIM, `<?xml ?>` declaration and
  all); styles/numbering/settings and header/footer parts take the VIEW
  route — the patched XML reparses into a fresh view so the view's
  serialization on save writes the patch instead of clobbering it (their id
  families are ones Word tolerates dangling); notes/comments/document.xml/
  rels/plumbing REJECT toward their own surfaces, because their ids pair
  with body markers and a whole-part swap could dangle references no gate
  can re-mint. Root tag must match on replace (a part's root is its
  identity). Inside a part there is deliberately NO positional addressing —
  keyed constructs belong to modeled verbs, everything else is whole-part
  get → modify → replace (or `part edit`: literal `--find`/`--with` over
  the part text, result re-gated, zero matches = MATCH_NOT_FOUND).
- [relationships.ts](relationships.ts) — rels addressing: `--at rIdN`/`rels`
  on get, `<Relationship>` fragments routed by shape on insert (no placement
  flags — the part is an unordered set keyed by Id), Id-preserving replace.
  Its gate set is proportionally small because Id/Type/Target/TargetMode IS
  the whole rels schema: no children, no unknown attributes, insert Ids must
  be free (omitted → minted, and the minted id is the printed handle),
  replace must keep the target's Id (a rename dangles every body reference),
  and an internal Target must resolve to an existing part — the dangling-rId
  invariant pointed the other way. No `dcx:raw` marker (the rels part has no
  MCE) and no schema gate (document.xml is untouched); rels never surface in
  `read` — they're plumbing, and `raw get --at rels` is the discovery path.
- [namespaces.ts](namespaces.ts) — prefix table + auto-declare on the doc root,
  and the `dcx:raw="1"` marker stamp. `dcx` is registered in `mc:Ignorable`, so
  Word AND the validator skip it (verified), while `read` surfaces it as the
  `raw` note token / `rawXml` AST field. A resave by Word may drop the marker —
  annotation degrades, content stays.
- [references.ts](references.ts) — reference/id audit. Dangling REFERENCES
  (rId, note/comment ids) reject — we can't invent the target; colliding
  DEFINITIONS (`wp:docPr`/`pic:cNvPr`, bookmark ids) re-mint inside the
  fragment (deterministic → fix, don't hint); unknown style/numId → warning
  (Word falls back cleanly).
- [validate.ts](validate.ts) — the bundled schema engine: libxml2-wasm
  (~1.3 MB) + the ECMA-376 5th-edition TRANSITIONAL XSD closure for WML in
  [schemas/](schemas/) (~0.5 MB, text-imported; cross-file `xsd:import`s
  resolve through an in-memory `xmlRegisterInputProvider`). One compiled
  validator per process (~55 ms compile, single-digit ms per document).
  `validationXmlFor` applies **MCE preprocessing to a CLONE** first — strip
  `mc:Ignorable`-declared attrs/elements, resolve `mc:AlternateContent` →
  `mc:Fallback` — because real Word markup (`w14:paraId` on every `w:p`) is
  valid OOXML that the base schema doesn't know. Gate 6 is a **baseline diff**
  (`diffValidationIssues`, multiset on message text): a third of real documents
  carry pre-existing violations, so a raw edit is judged only on errors it
  ADDS. `docx validate` reuses the engine standalone, per WML part (parts are
  enumerated from the PACKAGE by root tag, not a name list — headers/footers
  are a numbered family no static list covers). The module instantiates the
  WASM at import, so the CLI imports it LAZILY (`cli/raw/commit.ts`
  dynamic-imports it only when the gate will actually run) — keep it out of
  static import chains that non-validating paths load.

## Schema provenance

The XSDs come from the freely published ECMA-376 5th-edition Part 4
(Transitional) annex zip (ecma-international.org), trimmed to wml.xsd's import
closure, with two mechanical patches: the W3C `xml.xsd` is bundled and the
`xsd:import`s that referenced it by URL (or with no `schemaLocation` at all)
point at the local copy. Word writes transitional markup, so transitional —
not strict — is the profile real documents validate against. Don't edit the
schemas; re-derive from the ECMA zip if they ever need updating.

## Things that look wrong but aren't

- The transitional schema is MORE permissive than the .NET SDK's Microsoft365
  validation profile (e.g. it accepts `w:rPr` children in any order via a
  choice group). Gate 3's order tables cover the strict-order containers we
  know; the two layers are complementary, not redundant.
- `Raw.prepareFragment` mutates the DOC ROOT (namespace declares,
  `mc:Ignorable`) even when a later gate fails — harmless, because a failed
  command never saves, and the in-memory Document is discarded.
