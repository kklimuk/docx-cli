# GFM Comprehensive Fixture

This document exercises the full GitHub-flavored Markdown spec plus our docx-cli extensions: inline math (`$..$`), display math (`$$..$$`), and CriticMarkup (`{++ins++}` / `{--del--}`).

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

This paragraph has **bold**, *italic*, ~~strikethrough~~, and `inline code`.

### 2.2 Nested decoration

Bold containing italic: **bold with *italic inside***.

Italic containing bold: *italic with **bold inside***.

Strikethrough containing bold: ~~strike with **bold** inside~~.

### 2.3 Line breaks

Soft line break:
this stays on the same paragraph.

Hard line break:  
this is on a new line.

### 2.4 Escapes and entities

Backslash escapes: \*not italic\*, \`not code\`, and \# not a heading.

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

  ```python
  print("code inside a list item")
  ```

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
> ```python
> print("this fence escapes the quote on import")
> ```
>
> Trailing paragraph re-enters the quote.

## 8. Code blocks

### 8.1 Indented (4-space)

    function indented() {
      return true;
    }

### 8.2 Fenced without language

```
plain fenced code,
no highlighting.
```

### 8.3 Fenced TypeScript

```typescript
function add(a: number, b: number): number {
	return a + b;
}
```

### 8.4 Fenced Python

```python
def hello(name: str) -> str:
	return f"Hello, {name}!"
```

### 8.5 Fenced SQL

```sql
SELECT users.id, users.name
FROM users
WHERE active = true
ORDER BY users.id DESC;
```

### 8.6 Fenced JSON

```json
{
	"name": "docx-cli",
	"version": "0.11.0"
}
```

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

![sample image](/Users/kirill.klimuk/workspace/docx-cli/tests/fixtures/assets/sample.png)

## 11. Thematic breaks

Three dashes:

---

Three asterisks:

***

Three underscores:

___

## 12. Inline math

The Pythagorean theorem: $a^2 + b^2 = c^2$.

Euler's identity: $e^{i\pi} + 1 = 0$.

A finite sum: $\sum_{i=0}^{n} i = \frac{n(n+1)}{2}$.

A square root: $\sqrt{x^2 + y^2}$.

## 13. Display math

### 13.1 Gaussian integral

$$
\int_{-\infty}^{\infty} e^{-x^2}\, dx = \sqrt{\pi}
$$

### 13.2 Matrix

$$
A = \begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}
$$

### 13.3 Logistic function

$$
f(x) = \frac{1}{1 + e^{-x}}
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
