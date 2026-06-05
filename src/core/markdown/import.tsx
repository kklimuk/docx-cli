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
import {
	computeExtentEmu,
	type ImageSource,
	ImageSourceError,
	Images,
	loadImageSource,
	nextDrawingId,
} from "../image";
import { NoteBody, noteConfig, TrackedNoteBody } from "../notes";
import { XmlNode } from "../parser";
import { resolveAuthor, resolveDate, TrackChanges } from "../track-changes";
import { remarkCriticMarkup } from "./critic";
import { MarkdownImportError } from "./errors";
import type { ResolvedImage, WalkContext } from "./inline";
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
	 * splice the returned blocks and persist. */
	async blocks(
		source: string,
		options: { authorFlag?: string } = {},
	): Promise<XmlNode[]> {
		const tree = parseToMdast(source);

		const tracked = this.document.isTrackChangesEnabled();
		const ctx: WalkContext = {
			document: this.document,
			tracked,
			authorFlag: options.authorFlag,
			mintedNoteIds: new Map(),
			imageCache: new Map(),
			definitions: collectDefinitions(tree),
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
		await registerFootnotes(definitions, ctx);
		await preloadImages(tree, ctx);

		return walkRoot(tree, ctx);
	}
}

function parseToMdast(source: string): Root {
	const processor = unified()
		.use(remarkParse)
		.use(remarkGfm)
		.use(remarkMath)
		.use(remarkCriticMarkup);
	const tree = processor.parse(source);
	// `runSync` applies the transformer plugins (remarkCriticMarkup is the only
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
		const numericId = footnotes.nextId();
		ctx.mintedNoteIds.set(definition.identifier, numericId);
		const bodyText = phrasingPlaintext(definition.children);
		// Tracked-body wrapping: footnote bodies authored under tracking carry
		// a `<w:ins>` around their content runs (mirrors Word's empirical
		// shape from `scripts/word-redlines.sh`). The reference-side `<w:ins>`
		// is emitted at the inline walker; both get distinct revision ids from
		// the shared allocator.
		if (ctx.tracked && ctx.mintTrackedMeta) {
			const bodyMeta = ctx.mintTrackedMeta();
			notesRoot.children.push(
				<TrackedNoteBody
					config={config}
					id={numericId}
					text={bodyText}
					meta={bodyMeta}
				/>,
			);
		} else {
			notesRoot.children.push(
				<NoteBody config={config} id={numericId} text={bodyText} />,
			);
		}
	}
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
	const extent = computeExtentEmu(source, {});
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
	const extent = computeExtentEmu(source, {});
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
