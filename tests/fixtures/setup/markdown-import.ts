import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/markdown-import.docx — a doc whose entire body is
 * authored from one `docx create --from FILE.md` call exercising the full
 * GitHub-flavored Markdown spec plus our extensions (math, CriticMarkup).
 *
 * Sections are numbered so screenshot review can match Word's rendering to
 * each feature one-to-one. Covers:
 *
 *   1. Headings — ATX 1..6 + Setext
 *   2. Paragraphs + inline formatting (bold / italic / strike / inline code,
 *      nested decoration, backslash escapes, HTML entities, soft/hard breaks)
 *   3. Bullet lists with nesting
 *   4. Ordered lists (with `start` value)
 *   5. Task lists
 *   6. List items with multi-block content (paragraph + code + blockquote)
 *   7. Blockquotes (simple, multi-paragraph, nested list)
 *   8. Indented + fenced code blocks (typescript / python / sql / json); the
 *      typescript block carries a `// returns the sum` comment so lowlight
 *      emits the gray comment-token color (6E7781) — the round-trip surface
 *      formerly covered by the standalone code-blocks.docx fixture.
 *   9. GFM tables (basic + alignment markers)
 *  10. Links (inline / autolink / email / reference) + inline image
 *  11. Thematic breaks (---, ***, ___)
 *  12. Inline math
 *  13. Display math (Gaussian, matrix, fraction)
 *  14. CriticMarkup (insertion, deletion, combined)
 *  15. Footnotes (multiple)
 *
 * Not exercised: HTML blocks / inline (we drop), YAML frontmatter (we drop),
 * comment-style footnote refs (`[^cN]`), endnote refs (`[^enN]`), tracked-
 * change refs (`[^tcN]`) — deliberately out of scope per S8.
 *
 * The fixture joins CORE_FIXTURES so LibreOffice round-trips the composed
 * shape end-to-end — every emitter the walker composes is already
 * individually covered, but this catches integration failures where two
 * parts interact (e.g. a list paragraph immediately before a table, or
 * inline math inside a list item).
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/markdown-import.docx");
const cliEntry = resolve(root, "src/index.ts");
const mdPath = resolve(import.meta.dir, "markdown-import.source.md");
const imagePath = resolve(root, "tests/fixtures/assets/sample.png");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

const SOURCE = `# GFM Comprehensive Fixture

This document exercises the full GitHub-flavored Markdown spec plus our docx-cli extensions: inline math (\`$..$\`), display math (\`$$..$$\`), and CriticMarkup (\`{++ins++}\` / \`{--del--}\`).

## 1. Headings

### 1.1 ATX headings

# Heading level 1
## Heading level 2
### Heading level 3
#### Heading level 4
##### Heading level 5
###### Heading level 6

### 1.2 Setext headings

Setext heading 1
================

Setext heading 2
----------------

## 2. Paragraphs and inline formatting

### 2.1 Basic decoration

This paragraph has **bold**, *italic*, ~~strikethrough~~, and \`inline code\`.

### 2.2 Nested decoration

Bold containing italic: **bold with *italic inside***.

Italic containing bold: *italic with **bold inside***.

Strikethrough containing bold: ~~strike with **bold** inside~~.

### 2.3 Line breaks

Soft line break:
this stays on the same paragraph.

Hard line break:${"  "}
this is on a new line.

### 2.4 Escapes and entities

Backslash escapes: \\*not italic\\*, \\\`not code\\\`, and \\# not a heading.

HTML entities: &amp; for ampersand, &copy; for copyright, &hearts; for heart, &#42; for asterisk.

## 3. Bullet lists

- first
- second
  - nested second
  - another nested
    - third-level nested
- third
- fourth

## 4. Ordered lists

1. one
2. two
3. three

Starting at 10:

10. tenth item
11. eleventh
12. twelfth

## 5. Task lists

- [ ] unchecked task
- [x] checked task
- [ ] another unchecked
- [x] another done

## 6. List items with multi-block content

- A list item with multiple paragraphs.

  Second paragraph still in the same item.

  \`\`\`python
  print("code inside a list item")
  \`\`\`

  > A blockquote inside a list item.

- Another item with only a single line.

## 7. Blockquotes

### 7.1 Simple

> A simple single-paragraph blockquote.

### 7.2 Multi-paragraph

> A multi-paragraph blockquote.
>
> Second paragraph of the same quote.

### 7.3 With a nested list

> Blockquote containing a list:
>
> - one
> - two
>   - nested bullet still inside the quote
> - three

### 7.4 Nested blockquote

> outer quote line.
>
> > nested quote line.
> >
> > > deeply nested line.

### 7.5 Task list inside a blockquote

> Task list in quote:
>
> - [ ] todo unchecked
> - [x] todo done

### 7.6 Code block inside a blockquote (escapes — see CLAUDE.md)

The code fence below intentionally breaks out of the surrounding blockquote on
write. Adjacent quoted paragraphs before and after surface as separate
blockquotes on round-trip.

> Intro paragraph stays quoted.
>
> \`\`\`python
> print("this fence escapes the quote on import")
> \`\`\`
>
> Trailing paragraph re-enters the quote.

## 8. Code blocks

### 8.1 Indented (4-space)

    function indented() {
      return true;
    }

### 8.2 Fenced without language

\`\`\`
plain fenced code,
no highlighting.
\`\`\`

### 8.3 Fenced TypeScript

\`\`\`typescript
function add(a: number, b: number): number {
\t// returns the sum
\treturn a + b;
}
\`\`\`

### 8.4 Fenced Python

\`\`\`python
def hello(name: str) -> str:
\treturn f"Hello, {name}!"
\`\`\`

### 8.5 Fenced SQL

\`\`\`sql
SELECT users.id, users.name
FROM users
WHERE active = true
ORDER BY users.id DESC;
\`\`\`

### 8.6 Fenced JSON

\`\`\`json
{
\t"name": "docx-cli",
\t"version": "0.11.0"
}
\`\`\`

## 9. Tables

### 9.1 Basic

| Name | Type | Default |
| --- | --- | --- |
| color | string | "blue" |
| size | number | 10 |
| enabled | boolean | true |

### 9.2 With alignment markers

| Left | Center | Right |
| :--- | :---: | ---: |
| a | b | c |
| 1 | 2 | 3 |
| longer text | middle | end |

## 10. Links and images

### 10.1 Links

Inline link: [docx-cli on GitHub](https://github.com/kklimuk/docx-cli).

Autolink: <https://example.com>.

Email autolink: <hello@example.com>.

Reference-style link: [docx-cli][repo].

[repo]: https://github.com/kklimuk/docx-cli "docx-cli repo"

### 10.2 Inline image

![sample image](${imagePath})

## 11. Thematic breaks

Three dashes:

---

Three asterisks:

***

Three underscores:

___

## 12. Inline math

The Pythagorean theorem: $a^2 + b^2 = c^2$.

Euler's identity: $e^{i\\pi} + 1 = 0$.

A finite sum: $\\sum_{i=0}^{n} i = \\frac{n(n+1)}{2}$.

A square root: $\\sqrt{x^2 + y^2}$.

## 13. Display math

### 13.1 Gaussian integral

$$
\\int_{-\\infty}^{\\infty} e^{-x^2}\\, dx = \\sqrt{\\pi}
$$

### 13.2 Matrix

$$
A = \\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}
$$

### 13.3 Logistic function

$$
f(x) = \\frac{1}{1 + e^{-x}}
$$

## 14. CriticMarkup

### 14.1 Insertion

This sentence has {++an inserted phrase++} in the middle.

### 14.2 Deletion

This sentence has {--a deleted phrase--} that should disappear when tracking is off.

### 14.3 Combined

Before {++added text++} between {--removed text--} after.

## 15. Footnotes

A reference to the first footnote[^one].

[^one]: This is the body of the first footnote.

Another reference[^two] in a different sentence.

[^two]: Body of the second footnote, with **bold** that is currently flattened to plain text by the importer.

## 16. End

The end of the comprehensive fixture.
`;

mkdirSync(dirname(out), { recursive: true });
await Bun.write(mdPath, SOURCE);
await cli("create", out, "--force", "--from", mdPath);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
