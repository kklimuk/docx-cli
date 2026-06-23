import type { TextRun } from "../ast/types";
import { Paragraph, type ParagraphOptions } from "../blocks";
import type { XmlNode } from "../parser";
import { codeBlockStyleIdFor } from "./style";
import { type CodeToken, highlightCode } from "./syntax-highlight";

/** Build one `<w:p>` per source line, each carrying `pStyle="CodeBlock"` (so
 *  the Code Block paragraph style provides monospace, indent, and adjacent-
 *  paragraph spacing collapse), and runs carrying `runStyle="Code"` so the
 *  Code character style picks up monospace as well — defensive: a few Word
 *  versions don't reliably cascade the paragraph-style font through to runs.
 *
 *  With `language`, each line's tokens get color via lowlight; without it,
 *  each line is a single uncolored run. An unknown language degrades to the
 *  uncolored path (better than failing — the agent's snippet still lands as
 *  a code block).
 *
 *  ECMA-376 §17.16.5 requires `xml:space="preserve"` on `<w:t>` whose content
 *  has leading/trailing whitespace; `@core/blocks::TextRunElement` already
 *  emits that on every text run, so we don't need to special-case indented
 *  lines here. */
export function buildCodeBlockParagraphs(
	content: string,
	language?: string,
	layout?: Pick<ParagraphOptions, "spacing" | "indent" | "tabs">,
): XmlNode[] {
	// Normalize line endings: a `--code-file` source from Windows (CRLF) or
	// classic Mac (CR) would otherwise leave each line with a trailing `\r`,
	// which serializes into `<w:t>line\r</w:t>` and renders as a stray glyph.
	const normalized = content.replace(/\r\n?/g, "\n");
	const tokenLines = collectTokenLines(normalized, language);
	const style = codeBlockStyleIdFor(language);
	// Layout-only paragraph options (spacing/indent/tabs from the ride-along flags)
	// compose with the CodeBlock pStyle — which still owns the monospace font and
	// contextualSpacing — and land on EVERY line so the block spaces uniformly.
	// Deliberately NOT style/list/alignment: those would fight the CodeBlock model.
	return tokenLines.map((tokens) => (
		<Paragraph
			style={style}
			runs={tokensToRuns(tokens)}
			spacing={layout?.spacing}
			indent={layout?.indent}
			tabs={layout?.tabs}
		/>
	));
}

function collectTokenLines(
	content: string,
	language: string | undefined,
): CodeToken[][] {
	if (language) {
		const tokens = highlightCode(language, content);
		if (tokens) return splitTokensByLine(tokens);
	}
	return content.split("\n").map((line) => [{ text: line }]);
}

function tokensToRuns(tokens: CodeToken[]): TextRun[] {
	// Empty line: emit a single empty run so `<w:p>` still has a run child —
	// otherwise the paragraph would be just `<w:p><w:pPr.../></w:p>` and Word
	// renders it as a blank line correctly anyway, but the rPr (font) doesn't
	// stick on the next paragraph break. One empty run keeps every paragraph
	// shaped consistently.
	if (tokens.length === 0) {
		return [{ type: "text", text: "", runStyle: "Code" }];
	}
	return tokens.map((token) => {
		const run: TextRun = {
			type: "text",
			text: token.text,
			runStyle: "Code",
		};
		if (token.color) run.color = token.color;
		return run;
	});
}

function splitTokensByLine(tokens: CodeToken[]): CodeToken[][] {
	// lowlight emits tokens whose text may contain `\n` (e.g. multi-line
	// string literals or comments). Split each token at `\n` boundaries so
	// every paragraph corresponds to one source line — that's the natural
	// OOXML representation of a code block (each line = one `<w:p>`).
	const lines: CodeToken[][] = [[]];
	for (const token of tokens) {
		const parts = token.text.split("\n");
		for (let index = 0; index < parts.length; index++) {
			if (index > 0) lines.push([]);
			const part = parts[index];
			if (part === undefined || part.length === 0) continue;
			const current = lines[lines.length - 1];
			if (!current) continue;
			const piece: CodeToken = { text: part };
			if (token.color) piece.color = token.color;
			current.push(piece);
		}
	}
	return lines;
}
