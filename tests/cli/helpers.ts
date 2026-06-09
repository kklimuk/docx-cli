import { join } from "node:path";
import { Pkg } from "@core/ast/document/package";
import { runCli, tempWorkspace } from "./harness";

// Shared building blocks for the CLI tests. Verb-specific assertions stay in the
// per-verb file; the boilerplate every file kept re-deriving (read the markdown
// or raw document.xml, copy a fixture, list tracked-change kinds) lives here.

/** `docx read` (default markdown view) → stdout. */
export async function readMarkdown(path: string): Promise<string> {
	return (await runCli("read", path)).stdout;
}

/** Raw `word/document.xml` — for asserting on XML we don't model in the AST. */
export async function readDocumentXml(path: string): Promise<string> {
	const pkg = await Pkg.open(path);
	return await pkg.readText("word/document.xml");
}

/** The `kind`s reported by `track-changes list`, in order. */
export async function trackedKinds(path: string): Promise<string[]> {
	const result = await runCli("track-changes", "list", path);
	return (result.parsed as Array<{ kind: string }>).map(
		(change) => change.kind,
	);
}

/** A fresh, mutable temp copy of a committed fixture (so tests never write to
 *  `tests/fixtures/` in place). */
export async function freshFixture(
	label: string,
	fixturePath: string,
): Promise<string> {
	const docPath = join(tempWorkspace(label), "doc.docx");
	await Bun.write(docPath, Bun.file(fixturePath));
	return docPath;
}
