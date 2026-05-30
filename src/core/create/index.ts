import { Pkg } from "../ast/document/package";
import { CANONICAL_PARTS } from "./canonical-parts";
import {
	CONTENT_TYPES,
	corePropertiesXml,
	DOCUMENT_RELS,
	documentXml,
	ROOT_RELS,
} from "./template";

export { CANONICAL_PARTS } from "./canonical-parts";

export type BlankDocxOptions = {
	/** Target path the returned `Pkg` saves to. */
	path: string;
	title?: string;
	author?: string;
	/** Seed the first paragraph with this text; omitted leaves it empty. */
	text?: string;
	/** ISO timestamp for core-properties created/modified; defaults to now. */
	now?: string;
};

/** Assemble a minimal, Word-canonical `.docx` as an unsaved `Pkg`: the
 * document body (one paragraph, optionally seeded with `text`), core
 * properties, the package + document relationships, content-types, and the
 * six canonical parts Word treats as required even though ECMA-376 marks them
 * optional (styles / settings / fontTable / webSettings / theme / app — see
 * [canonical-parts.ts](canonical-parts.ts)). Caller persists with
 * `pkg.save()`. Lives in core because it's pure OOXML construction; the CLI's
 * `create` verb is a thin wrapper over it. */
export function buildBlankPackage(options: BlankDocxOptions): Pkg {
	// Honor `DOCX_CLI_NOW` (same env var the track-changes `resolveDate`
	// reads) so fixture rebuilds can pin a deterministic timestamp without
	// every caller threading `options.now` through. Falls back to the
	// wall-clock for normal CLI usage.
	const now = options.now ?? Bun.env.DOCX_CLI_NOW ?? new Date().toISOString();
	const pkg = Pkg.empty(options.path);
	pkg.writeText("[Content_Types].xml", CONTENT_TYPES);
	pkg.writeText("_rels/.rels", ROOT_RELS);
	pkg.writeText("word/document.xml", documentXml(options.text));
	pkg.writeText("word/_rels/document.xml.rels", DOCUMENT_RELS);
	pkg.writeText(
		"docProps/core.xml",
		corePropertiesXml({
			title: options.title,
			author: options.author,
			now,
		}),
	);
	for (const part of Object.values(CANONICAL_PARTS)) {
		pkg.writeText(part.zipPath, part.body);
	}
	return pkg;
}
