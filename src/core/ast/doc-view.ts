import { Pkg } from "../package";
import { XmlNode } from "../parser";
import { buildDoc } from "./read";
import type { Block, Doc } from "./types";

export type BlockReference = { node: XmlNode; parent: XmlNode[] };
export type CommentReference = { node: XmlNode; parent: XmlNode[] };
export type ImageReference = {
	relationshipId: string;
	partName: string;
	contentType: string;
};
export type HyperlinkReference = {
	node: XmlNode;
	parent: XmlNode[];
	relationshipId?: string;
};

export type DocView = {
	pkg: Pkg;
	documentTree: XmlNode[];
	commentsTree?: XmlNode[];
	commentsExtTree?: XmlNode[];
	footnotesTree?: XmlNode[];
	endnotesTree?: XmlNode[];
	relationshipsTree: XmlNode[];
	contentTypesTree: XmlNode[];
	corePropertiesTree?: XmlNode[];
	settingsTree?: XmlNode[];
	doc: Doc;
	blockReferences: Map<string, BlockReference>;
	commentReferences: Map<string, CommentReference>;
	imagesByRelationshipId: Map<
		string,
		{ partName: string; contentType: string }
	>;
	imageById: Map<string, ImageReference>;
	hyperlinksByRelationshipId: Map<string, { url: string }>;
	hyperlinkById: Map<string, HyperlinkReference>;
};

export async function openDocView(path: string): Promise<DocView> {
	const pkg = await Pkg.open(path);

	const documentTree = XmlNode.parse(await pkg.readText("word/document.xml"));
	const relationshipsTree = XmlNode.parse(
		await pkg.readText("word/_rels/document.xml.rels"),
	);
	const contentTypesTree = XmlNode.parse(
		await pkg.readText("[Content_Types].xml"),
	);
	const commentsTree = pkg.hasPart("word/comments.xml")
		? XmlNode.parse(await pkg.readText("word/comments.xml"))
		: undefined;
	const commentsExtTree = pkg.hasPart("word/commentsExtended.xml")
		? XmlNode.parse(await pkg.readText("word/commentsExtended.xml"))
		: undefined;
	const corePropertiesTree = pkg.hasPart("docProps/core.xml")
		? XmlNode.parse(await pkg.readText("docProps/core.xml"))
		: undefined;
	const settingsTree = pkg.hasPart("word/settings.xml")
		? XmlNode.parse(await pkg.readText("word/settings.xml"))
		: undefined;
	const footnotesTree = pkg.hasPart("word/footnotes.xml")
		? XmlNode.parse(await pkg.readText("word/footnotes.xml"))
		: undefined;
	const endnotesTree = pkg.hasPart("word/endnotes.xml")
		? XmlNode.parse(await pkg.readText("word/endnotes.xml"))
		: undefined;

	const view: DocView = {
		pkg,
		documentTree,
		commentsTree,
		commentsExtTree,
		footnotesTree,
		endnotesTree,
		relationshipsTree,
		contentTypesTree,
		corePropertiesTree,
		settingsTree,
		doc: undefined as unknown as Doc,
		blockReferences: new Map(),
		commentReferences: new Map(),
		imagesByRelationshipId: new Map(),
		imageById: new Map(),
		hyperlinksByRelationshipId: new Map(),
		hyperlinkById: new Map(),
	};

	view.doc = buildDoc(view, pkg.path);
	return view;
}

export async function saveDocView(view: DocView, path?: string): Promise<void> {
	view.pkg.writeText("word/document.xml", XmlNode.serialize(view.documentTree));
	view.pkg.writeText(
		"word/_rels/document.xml.rels",
		XmlNode.serialize(view.relationshipsTree),
	);
	view.pkg.writeText(
		"[Content_Types].xml",
		XmlNode.serialize(view.contentTypesTree),
	);
	if (view.commentsTree) {
		view.pkg.writeText(
			"word/comments.xml",
			XmlNode.serialize(view.commentsTree),
		);
	}
	if (view.commentsExtTree) {
		view.pkg.writeText(
			"word/commentsExtended.xml",
			XmlNode.serialize(view.commentsExtTree),
		);
	}
	if (view.settingsTree) {
		view.pkg.writeText(
			"word/settings.xml",
			XmlNode.serialize(view.settingsTree),
		);
	}
	await view.pkg.save(path);
}

export async function enrichImageHashes(view: DocView): Promise<void> {
	const seen = new Set<string>();
	await walkBlocksForImages(view, view.doc.blocks, seen);
}

async function walkBlocksForImages(
	view: DocView,
	blocks: Block[],
	seen: Set<string>,
): Promise<void> {
	for (const block of blocks) {
		if (block.type === "paragraph") {
			for (const run of block.runs) {
				if (run.type !== "image" || run.hash) continue;
				if (seen.has(run.id)) continue;
				seen.add(run.id);
				const reference = view.imageById.get(run.id);
				if (!reference) continue;
				const bytes = await view.pkg.readBytes(reference.partName);
				run.hash = await sha256Hex(bytes);
			}
			continue;
		}
		if (block.type === "table") {
			for (const row of block.rows) {
				for (const cell of row.cells) {
					await walkBlocksForImages(view, cell.blocks, seen);
				}
			}
		}
	}
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
	const owned = new Uint8Array(bytes.byteLength);
	owned.set(bytes);
	const buffer = await crypto.subtle.digest("SHA-256", owned);
	const view = new Uint8Array(buffer);
	let hex = "";
	for (let index = 0; index < view.length; index++) {
		const byte = view[index];
		if (byte === undefined) continue;
		hex += byte.toString(16).padStart(2, "0");
	}
	return hex;
}
