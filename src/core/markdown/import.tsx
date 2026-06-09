import type {
	Definition,
	FootnoteDefinition,
	Image,
	ImageReference,
	Nodes,
	PhrasingContent,
	Root,
} from "mdast";
import remarkGfm from "remark-gfm";
import remarkMath from "remark-math";
import remarkParse from "remark-parse";
import { unified } from "unified";
import type { Document } from "../ast/document";
import type { NotesView } from "../ast/document/notes";
import type { RelationshipsView } from "../ast/document/relationships";
import {
	computeExtentEmu,
	type ImageSource,
	ImageSourceError,
	Images,
	loadImageSource,
	nextDrawingId,
} from "../image";
import { w } from "../jsx";
import { NoteBody, noteConfig, TrackedNoteBody } from "../notes";
import { XmlNode } from "../parser";
import { getPageContentWidthEmu } from "../sections";
import { resolveAuthor, resolveDate, TrackChanges } from "../track-changes";
import { MarkdownImportError } from "./errors";
import { type ResolvedImage, type WalkContext, walkInline } from "./inline";
import { remarkInlineSurgery } from "./inline-surgery";
import { walkRoot } from "./walker";

/** Cross-cutting lens over "import this markdown source into the document."
 * Stateless — `new MarkdownImport(document).blocks(source)` parses the
 * markdown via remark + remark-gfm + remark-math + a CriticMarkup plugin,
 * pre-walks the mdast tree to resolve image bytes and register footnote
 * bodies, then walks once more synchronously to produce `<w:body>`-ready
 * `XmlNode[]`. The caller splices the result. Throws `MarkdownImportError`
 * for domain failures (bad LaTeX, broken image source, unsupported node). */
export class MarkdownImport {
	constructor(private document: Document) {}

	/** Parse `source` and return the block-level XmlNodes ready to splice
	 * into `<w:body>`. Footnote definitions and image rels/parts are written
	 * as a side effect on `this.document`; the caller's job is only to
	 * splice the returned blocks and persist.
	 *
	 * `options.relationships` overrides where hyperlink (and media) rels are
	 * minted — the note-body callers pass the notes part's OWN rels so a
	 * `<w:hyperlink r:id>` spliced into `footnotes.xml` resolves against
	 * `word/_rels/footnotes.xml.rels`, not the document's (a dangling rId
	 * otherwise — Word reports "unreadable content"). `options.stripImages`
	 * drops every `image`/`imageReference` before the walk so a note body
	 * (text + links only) can't mint a media rel in the wrong part. */
	async blocks(
		source: string,
		options: {
			authorFlag?: string;
			relationships?: RelationshipsView;
			stripImages?: boolean;
		} = {},
	): Promise<XmlNode[]> {
		// The leading `<!-- docx:base … -->` note (emitted by `read`) is a
		// VISIBILITY hint, not parse-back: it tells an agent the document's
		// dominant font/size so new content can match, but the importer does NOT
		// reconstruct it — it flows through as a block `html` node and `walkBlock`
		// drops it like every other comment (per "comments are never anything but
		// hints"). A full `read → create` rebuild therefore falls back to the
		// template docDefaults for the dominant font/size; `read --ast` stays
		// lossless and in-place `edit` preserves runs.
		const tree = parseToMdast(source);
		if (options.stripImages) stripImageNodes(tree);

		const tracked = this.document.isTrackChangesEnabled();
		const ctx: WalkContext = {
			document: this.document,
			tracked,
			authorFlag: options.authorFlag,
			mintedNoteIds: new Map(),
			footnoteRefCursor: new Map(),
			imageCache: new Map(),
			definitions: collectDefinitions(tree),
			relationships: options.relationships ?? this.document.relationships,
		};

		if (tracked) {
			const allocator = new TrackChanges(this.document).createAllocator();
			const author = resolveAuthor(options.authorFlag);
			const date = resolveDate();
			ctx.mintTrackedMeta = () => ({
				author,
				date,
				revisionId: allocator.next(),
			});
		}

		// Pre-walk: footnote definitions (mint ids + write bodies) and inline
		// images (fetch bytes, mint rels). Both must finish before the
		// synchronous block walk so the walker can read from `ctx.mintedNoteIds`
		// / `ctx.imageCache` without itself being async.
		const definitions = collectFootnoteDefinitions(tree);
		await registerFootnotes(definitions, countFootnoteReferences(tree), ctx);
		await preloadImages(tree, ctx);

		return walkRoot(tree, ctx);
	}
}

function parseToMdast(source: string): Root {
	const processor = unified()
		.use(remarkParse)
		.use(remarkGfm)
		.use(remarkMath)
		.use(remarkInlineSurgery);
	const tree = processor.parse(source);
	// `runSync` applies the transformer plugins (remarkInlineSurgery is the only
	// transformer in our pipeline; remarkParse/Gfm/Math register parser
	// extensions that already ran during `.parse`). It returns the same tree
	// mutated in place; we cast back to `Root` because unified's typing keeps
	// it generic.
	processor.runSync(tree);
	return tree as Root;
}

/** Walk the mdast tree once and pull out every `footnoteDefinition`. We
 * don't need to remove them — the block walker treats them as no-ops since
 * we've already registered the bodies into footnotes.xml. */
function collectFootnoteDefinitions(root: Root): FootnoteDefinition[] {
	const out: FootnoteDefinition[] = [];
	visitNodes(root, (node) => {
		if (node.type === "footnoteDefinition") out.push(node);
	});
	return out;
}

/** Count how many times each footnote identifier is referenced (`[^x]`) in the
 * body. Markdown allows reusing a footnote; OOXML/Word require a *distinct*
 * footnote definition per reference, so the registrar mints this many clones of
 * each definition. */
function countFootnoteReferences(root: Root): Map<string, number> {
	const counts = new Map<string, number>();
	visitNodes(root, (node) => {
		if (node.type === "footnoteReference") {
			counts.set(node.identifier, (counts.get(node.identifier) ?? 0) + 1);
		}
	});
	return counts;
}

/** Collect link/image reference definitions (`[ref]: url "title"`) keyed by
 *  their normalized identifier so the inline walker can resolve
 *  `[text][ref]` / `![alt][ref]` against them. mdast already attaches
 *  `node.identifier` (lowercased + whitespace-collapsed per the GFM spec) so
 *  no extra normalization is needed here. */
function collectDefinitions(root: Root): Map<string, Definition> {
	const out = new Map<string, Definition>();
	visitNodes(root, (node) => {
		if (node.type === "definition") out.set(node.identifier, node);
	});
	return out;
}

/** Recursively drop every `image` / `imageReference` node from the mdast tree,
 *  in place. Note bodies are text + links only; an image there would mint a
 *  media rel into the wrong part (the note-body markdown is spliced into
 *  `footnotes.xml`, whose rels are separate). Running before `preloadImages` +
 *  the walk means no image rel is minted and the walker never sees one. Mirrors
 *  the phrasing-level `stripImagePhrasing` used by `buildRichNoteBody`, but at
 *  the whole-tree level for the block-producing `blocks()` path. */
function stripImageNodes(node: { type: string; children?: unknown }): void {
	if (!Array.isArray(node.children)) return;
	const kept = (node.children as { type: string }[]).filter(
		(child) => child.type !== "image" && child.type !== "imageReference",
	);
	node.children = kept;
	for (const child of kept) stripImageNodes(child);
}

/** Walk every node in the tree and call `visitor`. Tiny in-house walker so
 * we don't reach for `unist-util-visit` at the lens level (the CriticMarkup
 * plugin uses it; this avoids the dependency creeping into the public lens). */
function visitNodes(
	node: { type: string; children?: unknown },
	visitor: (node: Nodes) => void,
): void {
	visitor(node as Nodes);
	if (Array.isArray(node.children)) {
		for (const child of node.children) {
			visitNodes(child as { type: string; children?: unknown }, visitor);
		}
	}
}

/** For each footnote definition: mint a numeric id, build the body XmlNode
 * (tracked-wrapped when `<w:trackChanges/>` is on), and append it to
 * `footnotes.xml`. The minted ids are stored in `ctx.mintedNoteIds` keyed by
 * the markdown identifier so inline `footnoteReference` walks find them. */
async function registerFootnotes(
	definitions: readonly FootnoteDefinition[],
	refCounts: ReadonlyMap<string, number>,
	ctx: WalkContext,
): Promise<void> {
	if (definitions.length === 0) return;
	const footnotes = ctx.document.ensureFootnotes();
	footnotes.ensureNoteStyles(ctx.document.ensureStyles());
	const config = noteConfig("footnote");
	const notesRoot = XmlNode.findRoot(footnotes.tree, config.rootTag);
	if (!notesRoot) {
		throw new MarkdownImportError(
			"USAGE",
			`Internal: <${config.rootTag}> root missing after ensureFootnotes`,
		);
	}

	for (const definition of definitions) {
		// OOXML/Word require a distinct footnote definition per reference (Word
		// treats N references to one definition as corruption and "repairs" it
		// by cloning). So mint one body per reference; an unreferenced
		// definition still gets a single orphan body.
		const copies = Math.max(refCounts.get(definition.identifier) ?? 0, 1);
		const ids: string[] = [];
		for (let copy = 0; copy < copies; copy++) {
			const numericId = footnotes.nextId();
			ids.push(numericId);
			if (ctx.tracked && ctx.mintTrackedMeta) {
				// Tracked footnote bodies carry a `<w:ins>` around their content
				// runs (mirrors Word's empirical shape from
				// `scripts/word-redlines.sh`). That verified shape is single-run
				// text only, so rich content (links/formatting) flattens to text
				// under tracking — the untracked path below keeps it.
				const bodyMeta = ctx.mintTrackedMeta();
				notesRoot.children.push(
					<TrackedNoteBody
						config={config}
						id={numericId}
						text={phrasingPlaintext(definition.children)}
						meta={bodyMeta}
					/>,
				);
			} else {
				// Untracked: walk the definition's inline content so hyperlinks and
				// bold/italic survive. Each clone mints its own rels into the
				// footnotes part (a `<w:hyperlink r:id>` inside footnotes.xml
				// resolves against word/_rels/footnotes.xml.rels, not the doc's).
				const body = buildRichNoteBody(definition, footnotes, ctx);
				notesRoot.children.push(
					<NoteBody
						config={config}
						id={numericId}
						runs={body.runs}
						paragraphs={body.paragraphs}
					/>,
				);
			}
		}
		ctx.mintedNoteIds.set(definition.identifier, ids);
	}
}

/** Build a footnote body that preserves inline formatting + hyperlinks. Walks
 *  each definition paragraph through the inline walker with a context whose
 *  relationships target the footnotes part's own rels, so note-body links are
 *  real, valid `<w:hyperlink>`s. The first paragraph's runs follow the back-ref
 *  numeral (with a leading space, matching Word); further paragraphs become
 *  sibling `<w:p>`. Images inside note bodies are dropped (rare, unsupported). */
function buildRichNoteBody(
	definition: FootnoteDefinition,
	notesView: NotesView,
	ctx: WalkContext,
): { runs: XmlNode[]; paragraphs: XmlNode[] } {
	const config = noteConfig("footnote");
	const noteCtx: WalkContext = {
		...ctx,
		relationships: notesView.ensureRelationships(),
	};
	const leadingSpace = (
		<w.r>
			<w.t {...{ "xml:space": "preserve" }}> </w.t>
		</w.r>
	);
	const paragraphs = definition.children.filter(
		(child) => child.type === "paragraph",
	);
	const [firstParagraph, ...restParagraphs] = paragraphs;
	if (!firstParagraph) return { runs: [leadingSpace], paragraphs: [] };

	const firstRuns = walkInline(
		stripImagePhrasing(firstParagraph.children),
		noteCtx,
	);
	const extraParagraphs = restParagraphs.map((paragraph) => (
		<w.p>
			<w.pPr>
				<w.pStyle w-val={config.textStyle} />
			</w.pPr>
			{walkInline(stripImagePhrasing(paragraph.children), noteCtx)}
		</w.p>
	));
	return { runs: [leadingSpace, ...firstRuns], paragraphs: extraParagraphs };
}

/** Recursively drop `image` / `imageReference` phrasing from a note body. Note
 *  bodies are text + links; an image there would mint a media rel in the wrong
 *  part. Returns the kept phrasing with the same nesting. */
function stripImagePhrasing(
	nodes: readonly PhrasingContent[],
): PhrasingContent[] {
	const out: PhrasingContent[] = [];
	for (const node of nodes) {
		if (node.type === "image" || node.type === "imageReference") continue;
		if ("children" in node && Array.isArray(node.children)) {
			out.push({
				...node,
				children: stripImagePhrasing(node.children as PhrasingContent[]),
			} as PhrasingContent);
		} else {
			out.push(node);
		}
	}
	return out;
}

/** Flatten a footnote definition's block children to a single plaintext
 * string. The `NoteBody` emitter takes plain text; we don't yet support
 * rich formatting inside footnote bodies (a follow-up could carry the full
 * block list and wrap each paragraph in its own `<w:p>`). `FootnoteDefinition`
 * children are typed `(BlockContent | DefinitionContent)[]`; we filter to
 * paragraphs and concatenate. */
function phrasingPlaintext(
	children: readonly FootnoteDefinition["children"][number][],
): string {
	const parts: string[] = [];
	for (const child of children) {
		if (child.type === "paragraph") {
			parts.push(phrasingContentToText(child.children));
		}
	}
	return parts.join(" ").trim();
}

function phrasingContentToText(nodes: readonly PhrasingContent[]): string {
	let out = "";
	for (const node of nodes) {
		if ("value" in node && typeof node.value === "string") {
			out += node.value;
			continue;
		}
		if ("children" in node && Array.isArray(node.children)) {
			out += phrasingContentToText(node.children as PhrasingContent[]);
		}
	}
	return out;
}

/** Pre-fetch every inline image referenced anywhere in the tree (paragraphs,
 * list items, table cells, blockquotes). Walks both direct `image` nodes and
 * `imageReference` nodes (resolved against `ctx.definitions`). The walker
 * then reads `ctx.imageCache.get(url)` synchronously. */
async function preloadImages(root: Root, ctx: WalkContext): Promise<void> {
	const images: Image[] = [];
	visitNodes(root, (node) => {
		if (node.type === "image") {
			images.push(node);
			return;
		}
		if (node.type === "imageReference") {
			const definition = ctx.definitions.get(
				(node as ImageReference).identifier,
			);
			if (definition) {
				images.push({
					type: "image",
					url: definition.url,
					title: definition.title ?? null,
					alt: (node as ImageReference).alt ?? null,
				});
			}
		}
	});
	// If any image URL is hash-shaped (`<sha256>.<ext>`), make sure every
	// ImageRun in the target document has its `hash` populated before we
	// look any up — the AST reader leaves them empty by default. Once-per
	// `MarkdownImport.blocks` call; cheap because `enrichHashes()` is
	// idempotent and dedupes by media part.
	if (images.some((image) => HASH_URL_PATTERN.test(image.url))) {
		await new Images(ctx.document).enrichHashes();
	}
	for (const image of images) {
		if (ctx.imageCache.has(image.url)) continue;
		ctx.imageCache.set(image.url, await loadAndRegisterImage(image, ctx));
	}
}

async function loadAndRegisterImage(
	image: Image,
	ctx: WalkContext,
): Promise<ResolvedImage> {
	// Hash-shaped URL (`<sha256>.<ext>`)? Try to reuse an existing image
	// from the target document first — that's the round-trip path: `docx
	// read --markdown` emits the same shape for every embedded image, and
	// piping the output back into `docx insert --markdown` should not
	// re-fetch / re-embed an image we already have. The hash is content-
	// addressed so identical bytes dedup naturally.
	const reused = await tryReuseImageByHash(image, ctx);
	if (reused) return reused;

	let source: ImageSource;
	try {
		source = await loadImageSource(image.url);
	} catch (error) {
		if (error instanceof ImageSourceError) {
			throw new MarkdownImportError("IMAGE_SOURCE", error.message);
		}
		throw error;
	}
	const extent = computeExtentEmu(
		source,
		{},
		getPageContentWidthEmu(ctx.document),
	);
	if (!extent) {
		throw new MarkdownImportError(
			"USAGE",
			`Could not read pixel dimensions from ${image.url}`,
			"The image source has no detectable intrinsic dimensions. Future flag `--image-width INCHES` will let you size it explicitly.",
		);
	}
	const { relationshipId } = new Images(ctx.document).add(source);
	return {
		relationshipId,
		drawingId: nextDrawingId(ctx.document.documentTree),
		widthEmu: extent.widthEmu,
		heightEmu: extent.heightEmu,
		alt: image.alt ?? undefined,
	};
}

/** A markdown image URL shaped like `<64-hex-char sha256>.<ext>` —
 * what `cli/read/markdown.ts` emits for every embedded image. The
 * extension is informational; the hash alone identifies the bytes.
 * Quick regex sniff keeps `loadImageSource` (file paths / data: URIs /
 * http(s) URLs) on the default path for everything else. */
const HASH_URL_PATTERN = /^([0-9a-f]{64})\.[a-z0-9]+$/i;

/** Look up the URL's content hash against the target document's media
 * inventory; if present, build a `ResolvedImage` that reuses the existing
 * relationship id so we don't mint a duplicate `word/media/imageN` part
 * (and don't issue any network/file IO). When the URL is hash-shaped but
 * the doc doesn't carry that image, raise a clear error — falling through
 * to `loadImageSource("abc123…ef0.png")` would just produce an ENOENT
 * with no hint about what went wrong. */
async function tryReuseImageByHash(
	image: Image,
	ctx: WalkContext,
): Promise<ResolvedImage | undefined> {
	const match = image.url.match(HASH_URL_PATTERN);
	if (!match?.[1]) return undefined;
	const hash = match[1].toLowerCase();
	const existing = ctx.document.body.findImageByHash(hash);
	if (!existing) {
		throw new MarkdownImportError(
			"IMAGE_SOURCE",
			`No image with hash ${hash} found in this document`,
			"This URL looks like a docx-cli content-addressed image reference (the shape `<sha256>.<ext>` that `docx read --markdown` emits). Either reference a file path, a data: URI, or an https:// URL — or use a hash that already exists in the target doc.",
		);
	}
	// Recompute extent from the stored media bytes — `<wp:extent>` is set
	// per-drawing in OOXML, not per-media-part, so we can't crib the
	// dimensions from the existing usage's drawing element (it may have
	// been resized).
	const bytes = await ctx.document.pkg.readBytes(existing.partName);
	if (!bytes) {
		throw new MarkdownImportError(
			"IMAGE_SOURCE",
			`Image hash ${hash} resolved to ${existing.partName} but the media part is missing from the package`,
		);
	}
	const source: ImageSource = await loadImageSource(
		`data:${existing.contentType};base64,${Buffer.from(bytes).toString("base64")}`,
	);
	const extent = computeExtentEmu(
		source,
		{},
		getPageContentWidthEmu(ctx.document),
	);
	if (!extent) {
		throw new MarkdownImportError(
			"IMAGE_SOURCE",
			`Could not read pixel dimensions from media part ${existing.partName} (hash ${hash})`,
		);
	}
	return {
		relationshipId: existing.relationshipId,
		drawingId: nextDrawingId(ctx.document.documentTree),
		widthEmu: extent.widthEmu,
		heightEmu: extent.heightEmu,
		alt: image.alt ?? undefined,
	};
}
