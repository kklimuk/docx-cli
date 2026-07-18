import { fail } from "../respond";

export type RawValues = Record<
	string,
	string | boolean | (string | boolean)[] | undefined
>;

/** Resolve the fragment source: inline `--xml` or `--xml-file PATH` (`-` =
 *  stdin). The inline value is VERBATIM — deliberately not routed through
 *  `decodeInlineEscapes`: XML is a machine format with its own escaping
 *  (`&#10;` for a literal newline), and a raw `\n` sequence inside an
 *  attribute value must survive untouched. This mirrors `--text-file` being
 *  the parser-free channel for prose.
 *
 *  The batch path reuses this with `{ label: "entry N", stdin: false }` —
 *  JSONL entries carry `xml`/`xmlFile` keys mapped onto the flag names, and
 *  stdin is already spoken for by `--batch -`. */
export async function readXmlSource(
	values: RawValues,
	options: { label?: string; stdin?: boolean } = {},
): Promise<string | number> {
	const prefix = options.label ? `${options.label}: ` : "";
	const inline = values.xml as string | undefined;
	const file = values["xml-file"] as string | undefined;
	if (inline !== undefined && file !== undefined) {
		return fail("USAGE", `${prefix}Pass either --xml or --xml-file, not both`);
	}
	if (inline !== undefined) {
		if (inline.trim() === "") return fail("USAGE", `${prefix}--xml is empty`);
		return inline;
	}
	if (file === undefined) {
		return fail(
			"USAGE",
			`${prefix}Missing fragment: pass --xml STR or --xml-file PATH (- = stdin)`,
		);
	}
	if (file === "-") {
		if (options.stdin === false) {
			return fail(
				"USAGE",
				`${prefix}stdin ("-") is not available here — it already feeds --batch`,
			);
		}
		return await new Response(Bun.stdin.stream()).text();
	}
	const source = Bun.file(file);
	if (!(await source.exists())) {
		return fail("FILE_NOT_FOUND", `${prefix}No such file: ${file}`);
	}
	return await source.text();
}
