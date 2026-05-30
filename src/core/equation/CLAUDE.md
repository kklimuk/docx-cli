# src/core/equation — OMML ↔ LaTeX

Five files behind the `@core/equation` barrel ([index.ts](index.ts)):

**Read side** (OMML → LaTeX):
- [read.ts](read.ts) — entry (`ommlToLatex`) + dispatcher
- [handlers.ts](handlers.ts) — per-element handlers + the small chr-to-command
  tables tightly coupled to them (NARY_OP_BY_CHR, ACCENT_CMD_BY_CHR,
  GROUP_CHR_CMD_BY_CHR), plus the `<m:d>` matrix-shorthand /
  `\binom` detection
- [latex.ts](latex.ts) — LaTeX text/token shape helpers shared with the
  write side (TEXT_LATEX_MAP, escapeAndMap, convertGrouped,
  promoteOperatorName, KNOWN_OPERATORS)

**Write side** (LaTeX → OMML):
- [emit.tsx](emit.tsx) — entry (`latexToOmml`, `<Equation>`); shells the
  LaTeX through `temml` (LaTeX → MathML, MIT, ~300 KB), parses the MathML
  string into an `XmlNode` tree, hands it to the adapter
- [mathml-to-omml.tsx](mathml-to-omml.tsx) — own adapter (MathML → OMML).
  Pure JSX components for each OMML construct (`<Superscript>`,
  `<Fraction>`, `<NaryOperator>`, `<Delimited>`, `<Matrix>`,
  `<FunctionApply>`, `<Bar>`, `<Box>`, etc.) plus a dispatcher that walks
  the MathML tree and recognizes the patterns temml emits (e.g.,
  big-operator absorbed body, fenced expressions, function-name + arg)

The canonical read fixture is [tests/fixtures/equations.docx](
../../../tests/fixtures/equations.docx), built by
[tests/fixtures/setup/equations.ts](../../../tests/fixtures/setup/equations.ts).
It's authored by the CLI itself (`insert --equation`), dogfooding the
LaTeX→OMML pipeline. 72 equations cover atoms / Greek / accents / fractions
/ roots / big operators / decorations / spacing / labeled braces / norms /
matrices, plus famous formulas across statistics (mean, variance, Bayes,
Normal PDF), engineering / physics (Schrödinger, Maxwell-Ampère, Bernoulli,
Fourier), and chemistry (Henderson-Hasselbalch, Arrhenius, ideal gas,
Nernst).

## Reader (Phase 1, shipped)

`ommlToLatex(node)` walks an `<m:oMath>` or `<m:oMathPara>` subtree and
returns reconstructed LaTeX. Final-pass cleanup strips trailing whitespace
the disambiguation logic added (so `\xi)` doesn't read back as `\xi )`).
The walker has three key behaviors:

1. **Per-subtree degradation.** A handler that doesn't recognize an OMML
   element falls back to plaintext concatenation (`collectPlainText`) for
   just that subtree, not for the whole equation. A user encountering an
   unfamiliar inner construct gets `\frac{\partial f}{\partial g(x)}` where
   `g(x)` is the only degraded piece, not the whole fraction. The
   `text` field on `EquationRun` stays populated as a last-resort fallback
   when `latex` itself is empty.
2. **Tolerance for producer verbosity.** Word emits `<m:ctrlPr>`,
   `<m:*Pr>` wrappers carrying only `<m:ctrlPr>`, and
   `<w:rFonts ascii="Cambria Math">` on every math run. Pandoc emits none of
   that and renders identically. The walker's `PRESENTATION_TAGS` set in
   [read.ts](read.ts) catalogs every such element so both producers' OMML
   gives the same LaTeX.
3. **Matrix-shorthand recognition.** `<m:d>` wrapping a single `<m:m>` with
   `(…)`, `[…]`, `|…|`, `‖…‖` delimiters collapses to `\begin{pmatrix}` /
   `bmatrix` / `vmatrix` / `Vmatrix`; `<m:d>` wrapping a single
   `<m:f m:type="noBar"/>` with `(…)` collapses to `\binom{n}{k}`. Without
   the shorthand we'd emit `\left(\begin{matrix}…\end{matrix}\right)`,
   which works but adds visible noise.

The AST's `EquationRun` carries `id` (`eqN`, document order, parallels
imgN/linkN/tcN), `latex` (the new reconstructed form), `text` (legacy
plaintext fallback), and `display` (inline `<m:oMath>` vs block
`<m:oMathPara>`). The original `<m:oMath>` XmlNode lives in
`Document.equationReferences` for emit-back paths — same pattern as image
and tracked-change references. Reader registration happens in
[core/ast/read.ts](../ast/read.ts) inside the run walker.

### Operator name promotion

Pandoc encodes function operators (`\sin`, `\cos`, `\lim`, …) as math runs
with `<m:rPr><m:sty m:val="p"/>` (plain/upright style) holding the literal
ASCII name. We promote those to `\name` so LaTeX typesets them upright.
Word emits these via `<m:func>` (semantic function-application); the reader
recognizes both shapes. `KNOWN_OPERATORS` lives in [latex.ts](latex.ts)
(lim/log/sin/cos/exp/det/etc., plus their inverses). **Adding an operator:**
add the name to `KNOWN_OPERATORS` AND to `KNOWN_OPERATORS_FOR_WRAPPING` in
[mathml-to-omml.tsx](mathml-to-omml.tsx); both the leaf-text promotion
(`handleRun`) and the writer's function-name detection use it.

### Accent maps cover both Unicode encodings

Pandoc favors spacing/overscript codepoints (`‾` U+203E for `\bar`, `→`
U+2192 for `\vec`), while Word favors combining diacritics (U+0300–U+036F).
`ACCENT_CMD_BY_CHR` in [handlers.ts](handlers.ts) registers both forms so
each producer round-trips. The writer's `ACCENT_CHR_TO_COMBINING` table
([mathml-to-omml.tsx](mathml-to-omml.tsx)) translates temml's spacing forms
to OMML's expected combining forms so Word's renderer composes the accent
properly over the base letter. Same dual-encoding pattern applies to
`GROUP_CHR_CMD_BY_CHR` for overbrace/underbrace/etc.

### What the walker doesn't try to handle

- `<w:rPrChange>` and other property-revision wrappers nested inside math
  runs.
- Equation references (`\ref`/`\eqref`/`\label`) — these aren't in the OMML
  data model; they're a render-time concern handled by client-side JS in
  most LaTeX→HTML pipelines.
- Equation numbering (`\tag{N}`) — we emit the tag text inline after the
  equation; Word renders it as a trailing text run, not centered like
  display LaTeX. No native OMML construct.

## Emitter (Phase 2, shipped)

`latexToOmml(latex, display?)` and `<Equation latex display>` in
[emit.tsx](emit.tsx) drive the LaTeX → OMML pipeline used by
`insert --equation` and `edit --at eqN --equation`. Internally:

1. `temml.renderToString(latex, { displayMode })` parses LaTeX and emits a
   MathML string. temml is the well-maintained LaTeX-parsing library —
   broad coverage (KaTeX/MathJax-compatible), MIT-licensed, server-side
   ready.
2. `XmlNode.parse(mathmlString)` lifts the MathML into our XmlNode tree.
3. `mathmlToOmml(mathRoot)` walks the tree and emits OMML siblings via the
   PascalCase component builders in [mathml-to-omml.tsx](mathml-to-omml.tsx).

The component split (`<Superscript base sup>` not `handleMsup(node)`)
matches the rest of the codebase — pure `props → XmlNode` builders are
components per [src/core/CLAUDE.md](../CLAUDE.md). Adding a new construct:
add a component for the OMML shape, register a dispatcher entry in
`ELEMENT_HANDLERS` (or in `convertSiblings` for cross-sibling patterns)
that gathers the MathML children and instantiates the component.

### N-ary recognition pattern

`<msubsup><mo>∑</mo><sub><sup></msubsup>` followed by the next sibling
(which is the operator's body) is the shape temml emits for
`\sum_{i=1}^n a_i`. The dispatcher (`convertSiblings`) detects this and
calls `<NaryOperator>` consuming both — see `isNaryScript` /
`buildNary` in [mathml-to-omml.tsx](mathml-to-omml.tsx). Same recognition
applies to `\int`, `\prod`, `\oint`, etc. via `BIG_OP_CHARS`.

### Function-application pattern (`\sin x`, `\ln K`, `\log_{10} N`)

OMML's `<m:func>` is the semantic "function applied to argument" element
that tells Word to insert the conventional thin space between operator
name and arg. Without it, `<m:r upright>ln</m:r><m:r>K</m:r>` renders as
`lnK` (concatenated) in Word — wrong. The writer detects two shapes:

- **Inner-mrow function** (`\ln 2` shape): temml wraps the operator in an
  mrow with thin-space padding, `[mspace, mi=ln, mo⁡, mspace]`. The next
  outer sibling is the function argument. `wrappedOperatorName` returns
  the operator name; the dispatcher consumes the next sibling as `<m:e>`.
- **Outer-level FUNCTION APPLICATION** (`\log_{10} N` shape): the scripted
  operator (`<msub>`) sits at the outer level, and the `<mo>⁡</mo>`
  invisible operator marks the function-application boundary. The
  dispatcher pops the previously-emitted node (must look like a function
  name per `looksLikeFunctionName`) and wraps it.

Both routes build `<m:func>` with the fName carrying whatever structure
the operator had (plain run, sSub, sSup, …) and `<m:e>` carrying the arg.

### Fenced expressions

`\left(…\right)` and the matrix-wrapping `\begin{pmatrix}…\end{pmatrix}`
emit as `<mrow><mo fence>(</mo>…<mo fence>)</mo></mrow>` in MathML with
`stretchy="true"`. The dispatcher detects opening-fence `<mo>` and consumes
through the matching close into a `<Delimited>` component. Bracket-tracking
(`isFenceOpen`, `isFenceClose`, `findFenceClose`) prefers the explicit
`form="postfix"` attribute over the symmetric-char tables so symmetric
delimiters like `|` and `‖` (which are both opens AND closes in the char
tables) get tracked correctly. Non-stretchy atomic parens (`(x)`) skip the
`<m:d>` wrap so `(x)^2` reads back as `(x)^2`, not `\left(x\right)^2`.

### Boxed and accented constructs

- `\boxed{…}` → `<m:borderBox>` (NOT `<m:box>`; the latter is semantic
  grouping that doesn't draw a visible border)
- `\overline{…}` / `\underline{…}` → `<m:bar>` with `pos="top"`/`"bot"`
- `\vec`, `\hat`, `\tilde`, etc. → `<m:acc>` with the **combining**
  diacritic char (U+0300–U+036F range), translated from temml's spacing
  form by `ACCENT_CHR_TO_COMBINING`
- `\widehat`, `\widetilde`, `\overrightarrow` → `<m:groupChr>` (the
  stretchy variants — they extend to span their content)
- `\underbrace{X}_{Y}` / `\overbrace{X}^{Y}` with labels → `<m:limLow>` /
  `<m:limUpp>` wrapping `<m:groupChr>` so the label survives

## Round-trip stability

Most cases are exact fixed-points. Remaining cosmetic drift is documented
on the fixture: trailing-spaces-before-non-letter (e.g. `\delta = ...` vs
`\delta=...`) are inserted by the disambiguation logic in
`escapeAndMap` and not stripped. Both forms parse to identical OMML in
temml, so Word / LibreOffice / MathJax all render them the same. Tests
assert STRUCTURAL presence (`.toContain("\\sum_{i=1}^{n}")`) rather than
byte-exact round-trip strings.

## Tracked equations (Phase 3, shipped)

`<m:oMath>` and `<m:oMathPara>` are wrapped at the same nesting level as
`<w:r>` (direct children of `<w:p>`), so the tracked-change wrappers go
around them just like text runs. Three paths land:

**Insert under tracking** — `insert --after pN --equation X` extends
`applyTrackedInsertion` in [cli/insert/index.tsx](../../cli/insert/index.tsx)
to wrap `<m:oMath>` / `<m:oMathPara>` alongside `<w:r>`. The whole equation
shows up as a tracked insertion; accept keeps it, reject removes it.

**Delete under tracking** — `delete --at pN` (where the paragraph is an
equation) extends `applyTrackedDeletion` in [cli/delete/index.tsx](../../cli/delete/index.tsx)
similarly. Text runs in the paragraph still get `<w:t>` → `<w:delText>`
conversion; OMML siblings stay structurally intact inside the `<w:del>`
wrapper (no `<m:delText>` equivalent — Word treats the whole OMML as the
deleted unit).

**Edit under tracking** — `edit --at eqN --equation NEW` (or `--display`/
`--inline`) emits a paired `<w:del>OLD</w:del><w:ins>NEW</w:ins>` next to
each other in the same parent. `commitEquationEdit` in
[cli/edit/index.tsx](../../cli/edit/index.tsx) handles this. Our own
`track-changes accept --all` / `reject --all` resolve cleanly (accept keeps
NEW, reject restores OLD).

**Word's accept/reject** picks the semantically correct equation but leaves
a small structural skeleton (an empty `<m:sSup>` / `<m:f>`) next to the
kept equation — Word merges the two OMML siblings and strips the deleted
side's text content, but preserves the script/fraction container. Cosmetic;
the kept equation renders correctly.

What's still deferred:
- `<w:rPrChange>` inside math runs (font property revisions on individual
  symbols) — niche, would need a dedicated probe.

## Adding support for a new OMML element

1. Add a `handleXxx` function in [handlers.ts](handlers.ts) and register it
   in `ELEMENT_HANDLERS`.
2. Add its corresponding `<m:xxxPr>` wrapper to `PRESENTATION_TAGS` in
   [read.ts](read.ts) so the walker steps past it without recursing.
3. If it carries a character map (operator, accent, etc.), add the chars
   to the appropriate table in [handlers.ts](handlers.ts) — both spacing
   and combining variants for accents.
4. If the construct doesn't have a temml LaTeX → MathML mapping, plumb
   it into [mathml-to-omml.tsx](mathml-to-omml.tsx) so it round-trips.
5. Add a test case in [tests/cli/equations.test.ts](
   ../../../tests/cli/equations.test.ts) and extend the fixture builder
   if the construct isn't already covered by an existing entry.

The walker is bounded by what `<m:oMath>` schema allows. New OMML elements
don't appear out of nowhere — Microsoft's ECMA-376 §22 is the catalog.
