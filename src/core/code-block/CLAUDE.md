# src/core/code-block — code-block emit + syntax highlighting

Three files, reachable via `@core/code-block` ([index.ts](index.ts)):
- [syntax-highlight.ts](syntax-highlight.ts) — pure `(language, code) → CodeToken[]` tokenizer
- [paragraphs.tsx](paragraphs.tsx) — emitter that builds one `<w:p>` per source line
- [style.tsx](style.tsx) — `ensureCodeBlockStyles(document, language?)` + the language↔pStyle suffix mapping (`codeBlockStyleIdFor`, `codeBlockLanguageFromStyleId`, `isCodeBlockStyleId`)

The paragraph builder and tokenizer are pure (no `Document`, no package state); style provisioning takes a `Document` because it has to touch `styles.xml`. The lenses (`Insert.paragraph`, `Edit.paragraph`, `Edit.range`) call `ensureCodeBlockStyles` before building the paragraphs.

## Emit shape

`buildCodeBlockParagraphs(content, language?)` returns `XmlNode[]`. Each line of `content` becomes one paragraph with `pStyle="CodeBlock"`; each run inside carries `runStyle="Code"` — both styles are baseline-catalog entries (see `core/ast/document/styles.tsx`). `ensureCodeBlockStyles(document, language)` in [style.tsx](style.tsx) provisions `Code` (run) + `CodeBlock` (paragraph), and when a language is given the derived `CodeBlock-LANG` paragraph style so the language survives round-trip.

The `Code` character run-style is defensive: a few Word versions don't reliably cascade the paragraph-style font (Courier New) through to runs, so every run gets the inline rStyle too. Cheap and harmless.

Multi-line tokens (a string literal that spans `\n`, a multi-line comment) get split at line boundaries by `splitTokensByLine` so the "one paragraph per source line" invariant holds even with messy input.

## Why one paragraph per line, not `<w:br>` line breaks

OOXML supports both `<w:br>` within a single paragraph and N adjacent paragraphs. Word itself uses adjacent paragraphs when you paste code — each line is its own `<w:p w:pStyle="CodeBlock">`. We match that because (a) it's the Word-canonical shape, (b) `contextualSpacing` on the CodeBlock style collapses adjacent-paragraph spacing, so it visually matches a `<w:br>`-only block, and (c) it round-trips to GFM fenced blocks cleanly (the renderer in `cli/read/markdown.ts` groups consecutive CodeBlock paragraphs).

## Syntax highlighting

`highlightCode(language, code)` wraps `lowlight` (highlight.js with the bundled `common` grammar set — 37 languages). It walks the hast tree returned by lowlight and flattens it into a `{text, color?}` list: each leaf text node inherits the most specific highlight.js class color from its ancestor spans. Unknown classes fall through with no color (rendered in the doc body color). Unknown LANGUAGES return `null` so the caller can fall back to the uncolored path — never throws.

The color palette in [syntax-highlight.ts](syntax-highlight.ts) is GitHub-light inspired and intentionally narrow (≈12 colors covering keywords / strings / comments / numbers / functions / types). Distinguishing those reads code; finer gradations buy diminishing returns. **Adding a token class:** map the highlight.js class name (e.g. `hljs-built_in`) to a 6-char uppercase hex; the flattener picks it up automatically. Class precedence is order-of-classes inside the same span, so list more-specific classes first if a node carries multiple.

## What stays in cli/

Flag parsing (`--code`/`--code-file`/`--language`), stdin reading for `--code-file -`, and the markdown-render side (fenced-block collapse, inline backticks for `runStyle: "Code"`) are render/CLI concerns and stay in `cli/insert/` and `cli/read/markdown.ts`.
