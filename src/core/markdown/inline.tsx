import type {
	Definition,
	ImageReference,
	Link,
	LinkReference,
	Image as MdImage,
	PhrasingContent,
} from "mdast";
import type { InlineMath } from "mdast-util-math";
import type { Document } from "../ast/document";
import type { RelationshipsView } from "../ast/document/relationships";
import { latexToOmml } from "../equation";
import { Image } from "../image";
import { w } from "../jsx";
import type { NullableXmlNode, XmlNode } from "../parser";
import type { TrackedMeta } from "../track-changes";
import { Del, Ins } from "../track-changes/emit";
import type { CriticDelete, CriticInsert } from "./critic";
import { MarkdownImportError } from "./errors";

/** Pre-resolved image entry the lens populates by walking the mdast tree
 * once before the synchronous OOXML emit. Inline image runs read from this
 * cache keyed by mdast `image.url`. */
export type ResolvedImage = {
	relationshipId: string;
	drawingId: number;
	widthEmu: number;
	heightEmu: number;
	alt: string | undefined;
};

/** Context threaded through inline + block walks. The lens populates this
 * once and passes it down; the walker reads but never mutates `document`,
 * only `mintedNoteIds` (for footnote ref ↔ definition pairing) and the
 * lazily-built `revisionAllocator` (for tracked ins/del ids). */
export type WalkContext = {
	document: Document;
	tracked: boolean;
	authorFlag?: string;
	/** Markdown footnote identifier → the list of minted note ids for it, one
	 *  per reference in document order. Markdown lets the same `[^x]` be cited
	 *  repeatedly, but OOXML/Word require a *distinct* footnote definition per
	 *  reference (Word treats N references to one definition as corruption and
	 *  "repairs" it by cloning) — so the lens mints one clone per reference and
	 *  the walker consumes them in order via `footnoteRefCursor`. */
	mintedNoteIds: Map<string, string[]>;
	/** Per-identifier cursor into `mintedNoteIds`: how many references to this
	 *  footnote the walker has emitted so far. Each `[^x]` consumes the next id. */
	footnoteRefCursor: Map<string, number>;
	/** Pre-resolved image bytes, alt, and minted rel/drawing ids — keyed by
	 *  mdast `image.url`. Populated by the lens; the walker only reads. */
	imageCache: Map<string, ResolvedImage>;
	/** Link/image reference definitions (`[ref]: url "title"`) keyed by their
	 *  normalized identifier. Populated by the lens; consulted when the
	 *  inline walker hits a `linkReference` / `imageReference` so the
	 *  reference resolves to a real hyperlink / image rather than dropping. */
	definitions: Map<string, Definition>;
	/** Where hyperlink (and future media) relationships minted during this walk
	 *  are recorded. Defaults to the document's rels; when building a footnote
	 *  /endnote body the lens swaps in that part's own rels so a note-body
	 *  `<w:hyperlink r:id>` resolves against `word/_rels/footnotes.xml.rels`. */
	relationships: RelationshipsView;
	/** Allocator over `core/track-changes/computeMaxRevisionId`. Lazy-built
	 *  on first tracked emission so untracked walks don't pay the scan cost.
	 *  The lens preallocates an author/date pair so all tracked nodes from
	 *  one import share metadata. */
	mintTrackedMeta?: () => TrackedMeta;
};

/** Walk a sequence of phrasing nodes into `<w:p>`-child siblings: `<w:r>`,
 * `<w:hyperlink>`, `<m:oMath>`, `<w:ins>`, `<w:del>`. The shape is uniform
 * across paragraph / heading / table cell / list item, so every block-level
 * builder splats this result into its own `<w:p>` wrapper. */
export function walkInline(
	nodes: readonly PhrasingContent[],
	ctx: WalkContext,
): XmlNode[] {
	const out: XmlNode[] = [];
	for (const node of nodes) {
		for (const child of walkPhrasing(node, ctx, EMPTY_FORMAT)) out.push(child);
	}
	return out;
}

/** Inline run formatting. Accumulated through nested strong / em / delete /
 * inlineCode wrappers as the walker descends. Mirrors the field set on
 * `TextRun` that `<RunProperties>` in `core/blocks.tsx` consumes — we
 * duplicate the emit here so the walker is self-contained and the schema
 * order stays under one component. `hyperlinkId` is threaded but never
 * surfaces as a property (Word styles inherits from the `<w:hyperlink>`
 * wrapper); we keep it on the type so future refinements (e.g. a real
 * `Hyperlink` baseline style) have a single place to flip. */
type InlineFormat = {
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	code?: boolean;
	hyperlinkId?: string;
};

const EMPTY_FORMAT: InlineFormat = Object.freeze({});

function walkPhrasing(
	node: PhrasingContent,
	ctx: WalkContext,
	format: InlineFormat,
): XmlNode[] {
	switch (node.type) {
		case "text":
			return [textRunNode(node.value, format)];
		case "strong":
			return node.children.flatMap((child) =>
				walkPhrasing(child, ctx, { ...format, bold: true }),
			);
		case "emphasis":
			return node.children.flatMap((child) =>
				walkPhrasing(child, ctx, { ...format, italic: true }),
			);
		case "delete":
			return node.children.flatMap((child) =>
				walkPhrasing(child, ctx, { ...format, strike: true }),
			);
		case "inlineCode":
			return [textRunNode(node.value, { ...format, code: true })];
		case "break":
			return [
				<w.r>
					<RunProperties format={format} />
					<w.br />
				</w.r>,
			];
		case "link":
			return wrapHyperlink(node, ctx, format);
		case "image":
			return [inlineImageRun(node, ctx)];
		case "html":
			// Raw HTML (locator comments like `<!-- pN -->`, stray tags) has no
			// useful OOXML mapping — drop. Locator-comment fidelity is handled
			// at the read side; import is by design lossy here.
			return [];
		case "footnoteReference":
			return [footnoteReferenceRun(node.identifier, ctx)];
		case "inlineMath":
			return [inlineMathElement(node)];
		case "criticInsert":
			return wrapTracked("ins", node, ctx, format);
		case "criticDelete":
			return wrapTracked("del", node, ctx, format);
		case "linkReference":
			return resolveLinkReference(node, ctx, format);
		case "imageReference":
			return resolveImageReference(node, ctx);
		default:
			throw new MarkdownImportError(
				"USAGE",
				`Unsupported inline markdown node: ${(node as { type: string }).type}`,
			);
	}
}

/** Build a `<w:r>` for a text segment with the accumulated format. Emits
 * `<w:rPr>` only when at least one property is set. Schema child order
 * follows §17.3.2.28 — matches the equivalent block in `core/blocks.tsx`. */
function textRunNode(text: string, format: InlineFormat): XmlNode {
	return (
		<w.r>
			<RunProperties format={format} />
			<w.t {...{ "xml:space": "preserve" }}>{text}</w.t>
		</w.r>
	);
}

function RunProperties({ format }: { format: InlineFormat }): NullableXmlNode {
	if (
		!format.bold &&
		!format.italic &&
		!format.strike &&
		!format.code &&
		!format.hyperlinkId
	) {
		return null;
	}
	// `runStyle` schema order is `<w:rStyle>` first. When both Code and
	// Hyperlink would apply (e.g. ``[`code`](url)``), Code wins — it carries
	// the visually-stronger semantics; the surrounding `<w:hyperlink>`
	// wrapper still routes the click. Word/LibreOffice render this with a
	// monospace blue underlined run via the inheritance chain.
	return (
		<w.rPr>
			{format.code && <w.rStyle w-val="Code" />}
			{format.hyperlinkId && !format.code && <w.rStyle w-val="Hyperlink" />}
			{format.bold && <w.b />}
			{format.italic && <w.i />}
			{format.strike && <w.strike />}
		</w.rPr>
	);
}

function wrapHyperlink(
	node: Link,
	ctx: WalkContext,
	format: InlineFormat,
): XmlNode[] {
	const relationshipId = ctx.relationships.addHyperlink(node.url);
	ctx.document.ensureStyles().ensureStyle("Hyperlink");
	const inner = node.children.flatMap((child) =>
		walkPhrasing(child, ctx, { ...format, hyperlinkId: relationshipId }),
	);
	return [<w.hyperlink {...{ "r:id": relationshipId }}>{inner}</w.hyperlink>];
}

/** Resolve `[text][ref]` against the pre-collected `ctx.definitions`. If the
 * reference is unknown (typo / missing definition), fall back to rendering the
 * literal children as plain text — matches GFM's "if no definition is found,
 * the link reference becomes plain text" behavior. */
function resolveLinkReference(
	node: LinkReference,
	ctx: WalkContext,
	format: InlineFormat,
): XmlNode[] {
	const definition = ctx.definitions.get(node.identifier);
	if (!definition) {
		// GFM fallback: render the children as plain text. The children carry
		// the original visible label, which is what an author saw in source.
		return node.children.flatMap((child) => walkPhrasing(child, ctx, format));
	}
	const synthetic: Link = {
		type: "link",
		url: definition.url,
		title: definition.title ?? undefined,
		children: node.children,
	};
	return wrapHyperlink(synthetic, ctx, format);
}

/** Resolve `![alt][ref]` against `ctx.definitions`. Unknown references are
 * dropped silently — there's no plaintext fallback for an image (the `alt`
 * text is already lost in the OOXML mapping unless we drop in as a run, which
 * conflates the image with a text). */
function resolveImageReference(
	node: ImageReference,
	ctx: WalkContext,
): XmlNode[] {
	const definition = ctx.definitions.get(node.identifier);
	if (!definition) return [];
	const resolved = ctx.imageCache.get(definition.url);
	if (!resolved) {
		throw new MarkdownImportError(
			"USAGE",
			`Image reference [${node.identifier}] points at ${definition.url} but the lens didn't pre-resolve it`,
			"Internal walker invariant — please report.",
		);
	}
	return [
		<Image
			relationshipId={resolved.relationshipId}
			drawingId={resolved.drawingId}
			widthEmu={resolved.widthEmu}
			heightEmu={resolved.heightEmu}
			alt={node.alt ?? resolved.alt}
		/>,
	];
}

function inlineImageRun(image: MdImage, ctx: WalkContext): XmlNode {
	const resolved = ctx.imageCache.get(image.url);
	if (!resolved) {
		// The lens pre-walked + resolved every inline image before calling here;
		// a miss is an internal contract violation. Surface clearly rather than
		// silently dropping the image.
		throw new MarkdownImportError(
			"USAGE",
			`Inline image source not pre-resolved: ${image.url}`,
			"Internal walker invariant — please report.",
		);
	}
	return (
		<Image
			relationshipId={resolved.relationshipId}
			drawingId={resolved.drawingId}
			widthEmu={resolved.widthEmu}
			heightEmu={resolved.heightEmu}
			alt={resolved.alt}
		/>
	);
}

function footnoteReferenceRun(identifier: string, ctx: WalkContext): XmlNode {
	const ids = ctx.mintedNoteIds.get(identifier);
	if (!ids || ids.length === 0) {
		throw new MarkdownImportError(
			"USAGE",
			`Footnote reference [^${identifier}] has no matching definition`,
			"Provide `[^id]: body text` for every `[^id]` reference.",
		);
	}
	// Consume the next clone for this footnote. The lens minted one definition
	// per reference (OOXML requires a 1:1 reference↔definition mapping), so the
	// Nth `[^identifier]` in document order takes the Nth id. Clamp to the last
	// id defensively if reference/definition counts ever drift.
	const cursor = ctx.footnoteRefCursor.get(identifier) ?? 0;
	ctx.footnoteRefCursor.set(identifier, cursor + 1);
	const numericId = ids[Math.min(cursor, ids.length - 1)];
	// `ensureNoteStyles` is idempotent — registering it on every ref keeps the
	// dependency local. The lens calls `ensureFootnotes()` once when it sees
	// the first definition; this call here is a defensive no-op if there are
	// refs but no defs (the missing-def error above fires first).
	const footnotes = ctx.document.ensureFootnotes();
	footnotes.ensureNoteStyles(ctx.document.ensureStyles());
	return (
		<w.r>
			<w.rPr>
				<w.rStyle w-val="FootnoteReference" />
			</w.rPr>
			<w.footnoteReference w-id={numericId} />
		</w.r>
	);
}

function inlineMathElement(node: InlineMath): XmlNode {
	try {
		return latexToOmml(node.value, false);
	} catch (error) {
		throw new MarkdownImportError(
			"USAGE",
			`Could not parse inline math: ${error instanceof Error ? error.message : String(error)}`,
			"Check the LaTeX. We accept the temml dialect (KaTeX-compatible).",
		);
	}
}

function wrapTracked(
	kind: "ins" | "del",
	node: CriticInsert | CriticDelete,
	ctx: WalkContext,
	format: InlineFormat,
): XmlNode[] {
	const runs = [textRunNode(node.value, format)];
	if (!ctx.tracked) {
		// Tracking off: insertions splat as plain text (the "accepted" view);
		// deletions drop entirely (their visible content is removed).
		if (kind === "ins") return runs;
		return [];
	}
	if (!ctx.mintTrackedMeta) {
		throw new MarkdownImportError(
			"USAGE",
			"Internal: tracked-meta allocator was not provisioned on the walk context",
		);
	}
	const meta = ctx.mintTrackedMeta();
	if (kind === "ins") return [<Ins meta={meta}>{runs}</Ins>];
	// `<w:del>` requires `<w:t>` → `<w:delText>`. Inline the rename rather
	// than re-walking — runs here are always plain `<w:r><w:t/></w:r>`.
	for (const run of runs) {
		for (const child of run.children) {
			if (child.tag === "w:t") child.tag = "w:delText";
		}
	}
	return [<Del meta={meta}>{runs}</Del>];
}
