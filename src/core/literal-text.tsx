import { Paragraph, type ParagraphOptions } from "./blocks";
import type { XmlNode } from "./parser";

/** Build one `<w:p>` per line of literal text — NO markdown parsing, so every
 *  character lands in the document verbatim. This is the "put this exact prose
 *  in, untouched" path behind `insert --text-file` and `create --text-file`:
 *  GFM parsing silently corrupts reviewer prose (`3. note` renumbers to `1.`,
 *  `*x*` italicizes, bare URLs autolink, `{++x++}` is eaten by CriticMarkup —
 *  and bare URLs have no escape sequence at all), so authoring from raw text
 *  needs a parser-free channel.
 *
 *  Splitting model: every newline starts a new paragraph (Word's "each Enter is
 *  a paragraph" model — what someone writing paragraphs into a text file
 *  expects). Interior blank lines therefore become empty paragraphs, exactly as
 *  authored. A `\t` inside a line still becomes a `<w:tab/>` via `Paragraph`'s
 *  `textToRuns`, which round-trips. */
export function literalParagraphs(
	text: string,
	options: ParagraphOptions = {},
): XmlNode[] {
	return splitLiteralLines(text).map((line) => (
		<Paragraph text={line} {...options} />
	));
}

/** Split literal text into per-line paragraph bodies. Normalizes CRLF/CR → LF
 *  and drops a SINGLE trailing newline (a conventional final newline shouldn't
 *  mint a stray empty trailing paragraph); every other newline — including the
 *  blank lines between paragraphs — yields its own (possibly empty) entry. An
 *  empty string yields one empty line so the caller always gets ≥1 paragraph. */
function splitLiteralLines(text: string): string[] {
	const normalized = text.replace(/\r\n?/g, "\n");
	const trimmed = normalized.endsWith("\n")
		? normalized.slice(0, -1)
		: normalized;
	return trimmed.split("\n");
}
