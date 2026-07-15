import type { RawValues } from "../insert/index";
import { fail } from "../respond";

/** Resolve `--code TEXT` (inline) or `--code-file PATH` (file / stdin) into the
 *  raw code body + optional language — the parser-free channel shared by `code
 *  add` and `code edit`. Every character lands verbatim; `--code-file -` reads
 *  stdin. Callers spread the result into their own spec shape (`InsertSpec` for
 *  add, `ParagraphContentSpec` + paragraph options for edit). `help` threads the
 *  command's own HELP into the mutex error. */
export async function resolveCodeContent(
	values: RawValues,
	help: string,
): Promise<{ content: string; language?: string } | number> {
	const inline = values.code as string | undefined;
	const file = values["code-file"] as string | undefined;
	if ((inline === undefined) === (file === undefined)) {
		return fail("USAGE", "Pass exactly one of --code or --code-file", help);
	}
	const language = values.language as string | undefined;
	if (inline !== undefined) {
		return { content: inline, ...(language ? { language } : {}) };
	}
	try {
		const content =
			file === "-"
				? await new Response(Bun.stdin.stream()).text()
				: await Bun.file(file as string).text();
		return { content, ...(language ? { language } : {}) };
	} catch (error) {
		const message = error instanceof Error ? error.message : String(error);
		return fail(
			"FILE_NOT_FOUND",
			`Failed to read --code-file ${file}: ${message}`,
		);
	}
}
