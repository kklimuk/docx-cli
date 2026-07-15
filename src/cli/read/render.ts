import type { Document } from "@core";
import { Images } from "@core/image";
import { type MarkdownView, renderMarkdown } from "./markdown";

/** Render an already-open document to its `docx read` markdown — the
 *  `enrichHashes → renderMarkdown` tail of the read pipeline, with each
 *  document's OWN defaults (font/size/track-changes state). Shared by `read`
 *  and `diff` so both produce byte-identical read views (diff compares two of
 *  them). Throws `MarkdownLocatorError` on a bad `from`/`to` slice — callers
 *  translate it to a usage error. */
export async function renderReadMarkdown(
	document: Document,
	options: {
		from?: string;
		to?: string;
		view?: MarkdownView;
		showComments?: boolean;
	} = {},
): Promise<string> {
	// Content-addressed image hashes back the `![alt](<sha256>.<ext>)` URLs, so
	// both sides of a diff must enrich or an unchanged image reads as changed.
	await new Images(document).enrichHashes();
	return renderMarkdown(document.body, {
		from: options.from,
		to: options.to,
		view: options.view,
		showComments: options.showComments,
		defaultSizeHalfPoints: document.styles?.defaultSizeHalfPoints(),
		defaultFont: document.styles?.defaultFont(),
		trackChangesOn: document.isTrackChangesEnabled(),
	});
}
