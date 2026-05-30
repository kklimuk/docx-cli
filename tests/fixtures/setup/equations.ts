import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps + tracked-change w:date to a fixed value so
// rebuilds are byte-deterministic. Honored by `core/create::buildBlankPackage`
// and by `track-changes::resolveDate`.
process.env.DOCX_CLI_NOW ??= "2026-05-22T00:00:00Z";

/**
 * Build tests/fixtures/equations.docx — the canonical equation read fixture.
 *
 * **CLI-authored**, dogfooding our own LaTeX→OMML emitter (`insert --equation`)
 * end-to-end. One paragraph per construct with a header above describing what
 * the equation exercises, so a failing case is easy to pin down by reading
 * `docx read` output. The body is laid out in 2 columns; a column break is
 * inserted every 5 equations so each rendered page has 10 equations (5 per
 * column), keeping the doc scannable across ~100 cases.
 *
 * Replaces three earlier fixtures: `equations-rich.docx` (Pandoc-built),
 * `equations-word.docx` (Word-emit verbosity probe — a producer we don't
 * control), and an older Wikipedia article export.
 *
 * Inputs are authored in **natural LaTeX** (what a human would type),
 * deliberately NOT pre-normalized to round-trip exactly. The remaining
 * round-trip whitespace differences (`\delta = ...` vs `\delta=...`,
 * `\sum_i` vs `\sum_{i}`) are LaTeX-source style only — both forms parse to
 * identical OMML, so Word/LibreOffice render them the same. Tests assert
 * structural presence (`.toContain("\\frac{FL")`) rather than exact strings.
 *
 * Construct coverage:
 *   - Atoms / Greek / accents / wide-accents
 *   - Fractions (simple, nested, partial, \binom)
 *   - Roots (square, nested, nth)
 *   - Big operators (sum / int / prod / lim with limits)
 *   - Decorations (\overline, \underline, \boxed)
 *   - Spacing (\,, \quad)
 *   - Labeled braces (\underbrace, \overbrace with labels)
 *   - Norms / delimiters
 *   - Matrices: pmatrix, aligned, cases
 *   - Tagged display (\tag)
 *   - Famous formulas: statistics, engineering/physics, chemistry
 *   - Run-level styling: \boldsymbol, \textcolor (named + hex), \cancel,
 *     font sizes (\Large/\small), math alphabet variants
 *     (\mathbf/\mathbb/\mathcal/\mathfrak/\mathsf/\mathtt)
 *   - Labelled relations: \overset, \underset, \stackrel
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/equations.docx");
const cliEntry = resolve(root, "src/index.ts");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

await cli("create", out, "--force", "--text", "Equation coverage fixture.");

// Two-column layout. The trailing `<w:sectPr>` on `<w:body>` (id `s0` after
// `create`) defines the columns for every paragraph in the body — editing
// it to `cols=2` makes the whole doc 2-column. Title + equations all flow
// through the 2-column layout; column breaks below give 5 equations per
// column = 10 per page.
await cli("edit", out, "--at", "s0", "--columns", "2");

// Each entry is [header, latex, display?]. The walker inserts the header at
// `pN`, then the equation at `pN+1`, so headers and equations interleave —
// makes the AST output readable. A page break is inserted after every 5
// equations so the rendered docx paginates cleanly (instead of one
// ~30-page wall of math).
const ENTRIES: Array<[string, string, boolean?]> = [
	// ---- Core constructs ----
	["Atom: superscript", "x^2"],
	["Atom: subscript", "a_i"],
	["Atom: both", "x_i^2"],
	["Atom: mass-energy", "E=mc^2"],
	["Greek: alpha+beta", "\\alpha + \\beta"],
	["Accent: \\hat", "\\hat{n}"],
	["Accent: \\bar", "\\bar{x}"],
	["Accent: \\vec", "\\vec{v}"],
	["Wide: \\widehat", "\\widehat{x}"],
	["Wide: \\widetilde", "\\widetilde{x}"],
	["Wide: \\overrightarrow", "\\overrightarrow{AB}"],
	["Fraction: simple", "\\frac{a}{b}"],
	["Fraction: nested", "\\frac{1}{1+\\frac{1}{x}}"],
	["Fraction: partial", "\\frac{\\partial f}{\\partial x}"],
	["Fraction: \\binom", "\\binom{n}{k}"],
	["Root: square", "\\sqrt{2}"],
	["Root: nested square", "\\sqrt{a^2+b^2}"],
	["Root: cube", "\\sqrt[3]{x+y}"],
	["Root: nth", "\\sqrt[n]{a_1 a_2 \\cdots a_n}"],
	["Big: \\sum", "\\sum_{i=1}^{n} a_i"],
	["Big: \\sum factorial", "\\sum_{k=0}^{\\infty} \\frac{1}{k!}"],
	["Big: \\int gaussian", "\\int_0^\\infty e^{-x^2} \\, dx"],
	["Big: \\iint", "\\iint_D f(x,y) \\, dx \\, dy"],
	["Big: \\oint", "\\oint_C \\vec{F} \\cdot d\\vec{r}"],
	["Big: \\prod", "\\prod_{i=1}^{n} i"],
	["Big: \\lim", "\\lim_{x \\to 0} \\frac{\\sin x}{x}"],
	["Bar: \\overline", "\\overline{x+y}"],
	["Bar: \\underline", "\\underline{x+y}"],
	["Box: \\boxed", "\\boxed{x+y}"],
	["Spacing: thin", "a \\, b"],
	["Spacing: quad", "a \\quad b"],
	["Brace: \\underbrace label", "\\underbrace{a+b+c}_{=n}"],
	["Brace: \\overbrace label", "\\overbrace{a+b+c}^{n}"],
	["Norm: \\|x\\|", "\\|x\\|"],
	["Delim: \\left(\\right)", "\\left(\\frac{a}{b}\\right)"],
	// ---- Statistics ----
	["Stat: mean", "\\bar{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_i"],
	["Stat: variance", "\\sigma^2 = \\frac{1}{n} \\sum_{i=1}^{n} (x_i - \\mu)^2"],
	[
		"Stat: normal PDF",
		"f(x) = \\frac{1}{\\sigma\\sqrt{2\\pi}} e^{-\\frac{(x-\\mu)^2}{2\\sigma^2}}",
	],
	["Stat: Bayes", "P(A|B) = \\frac{P(B|A) P(A)}{P(B)}"],
	["Stat: binomial PMF", "P(X=k) = \\binom{n}{k} p^k (1-p)^{n-k}"],
	["Stat: z-score", "z = \\frac{x - \\mu}{\\sigma}"],
	["Stat: chi-squared", "\\chi^2 = \\sum_{i=1}^{n} \\frac{(O_i - E_i)^2}{E_i}"],
	[
		"Stat: correlation",
		"\\rho_{XY} = \\frac{\\text{Cov}(X,Y)}{\\sigma_X \\sigma_Y}",
	],
	["Stat: expectation", "E[X] = \\sum_i x_i p_i"],
	// ---- Engineering / Physics ----
	["Eng: Newton gravity", "F = G \\frac{m_1 m_2}{r^2}"],
	["Eng: Coulomb", "F = k_e \\frac{q_1 q_2}{r^2}"],
	["Eng: energy-momentum", "E^2 = (mc^2)^2 + (pc)^2"],
	[
		"Eng: Schrödinger",
		"i\\hbar \\frac{\\partial \\Psi}{\\partial t} = \\hat{H} \\Psi",
	],
	[
		"Eng: wave equation",
		"\\frac{\\partial^2 u}{\\partial t^2} = c^2 \\nabla^2 u",
	],
	["Eng: Euler identity", "e^{i\\pi} + 1 = 0"],
	[
		"Eng: Maxwell-Ampère",
		"\\nabla \\times \\vec{B} = \\mu_0 \\vec{J} + \\mu_0 \\epsilon_0 \\frac{\\partial \\vec{E}}{\\partial t}",
	],
	[
		"Eng: Fourier transform",
		"\\hat{f}(\\xi) = \\int_{-\\infty}^{\\infty} f(x) e^{-2\\pi i x \\xi} \\, dx",
	],
	["Eng: Reynolds number", "Re = \\frac{\\rho v L}{\\mu}"],
	["Eng: Bernoulli", "P + \\frac{1}{2}\\rho v^2 + \\rho g h = \\text{const}"],
	[
		"Eng: heat equation",
		"\\frac{\\partial u}{\\partial t} = \\alpha \\nabla^2 u",
	],
	["Eng: beam deflection", "\\delta = \\frac{F L^3}{3 E I}"],
	// ---- Chemistry ----
	["Chem: pH", "\\text{pH} = -\\log_{10}[\\text{H}^+]"],
	[
		"Chem: Henderson-Hasselbalch",
		"\\text{pH} = pK_a + \\log_{10} \\frac{[A^-]}{[HA]}",
	],
	["Chem: Arrhenius", "k = A e^{-E_a / RT}"],
	["Chem: ideal gas", "PV = nRT"],
	["Chem: Gibbs free energy", "\\Delta G = -RT \\ln K"],
	["Chem: Nernst", "E = E^0 - \\frac{RT}{nF} \\ln Q"],
	["Chem: equilibrium constant", "K_{eq} = \\frac{[C]^c [D]^d}{[A]^a [B]^b}"],
	["Chem: rate law", "\\text{rate} = k [A]^m [B]^n"],
	["Chem: half-life (1st order)", "t_{1/2} = \\frac{\\ln 2}{k}"],
	["Chem: Boltzmann distribution", "\\frac{n_i}{n_0} = e^{-\\Delta E / k_B T}"],
	// ---- Math alphabet variants (Unicode mathematical alphanumerics) ----
	["Alphabet: \\mathbf", "\\mathbf{v} \\cdot \\mathbf{w}"],
	["Alphabet: \\mathit", "\\mathit{abc}"],
	["Alphabet: \\mathbb (R, N, Z, Q, C)", "\\mathbb{R} \\subset \\mathbb{C}"],
	["Alphabet: \\mathcal", "\\mathcal{L}(\\theta)"],
	["Alphabet: \\mathfrak (Lie algebra)", "\\mathfrak{g} \\oplus \\mathfrak{h}"],
	["Alphabet: \\mathsf", "\\mathsf{A} \\mathsf{B}"],
	["Alphabet: \\mathtt (monospace)", "\\mathtt{fix}(\\lambda)"],
	// ---- Bold symbols (\boldsymbol) ----
	[
		"Bold: \\boldsymbol{\\nabla}",
		"\\boldsymbol{\\nabla} \\cdot \\boldsymbol{E}",
	],
	["Bold: \\boldsymbol vector", "\\boldsymbol{F} = m \\boldsymbol{a}"],
	["Bold: \\boldsymbol stress tensor", "\\boldsymbol{\\sigma}_{ij}"],
	["Bold: \\boldsymbol Greek span", "\\boldsymbol{\\alpha + \\beta + \\gamma}"],
	// ---- Color (\textcolor) ----
	["Color: red text", "\\textcolor{red}{x + y = z}"],
	["Color: blue fraction", "\\textcolor{blue}{\\frac{a}{b}}"],
	["Color: hex (orange)", "\\textcolor[HTML]{FFA500}{f(x) = \\sin x}"],
	[
		"Color: mixed inline",
		"\\textcolor{red}{a} + \\textcolor{green}{b} = \\textcolor{blue}{c}",
	],
	// ---- Font sizes ----
	["Size: \\Large", "{\\Large E = mc^2}"],
	["Size: \\small inline", "x + {\\small \\epsilon}"],
	// ---- Cancel / strike-through ----
	["Cancel: simple", "\\frac{\\cancel{x}}{\\cancel{x} y} = \\frac{1}{y}"],
	["Cancel: in algebra step", "\\cancel{a} + b - \\cancel{a} = b"],
	// ---- Overset / underset / stackrel ----
	["Overset: question", "A \\overset{?}{=} B"],
	["Overset: chemistry arrow", "A \\overset{\\text{cat}}{\\longrightarrow} B"],
	[
		"Underset: convergence",
		"A_n \\underset{n \\to \\infty}{\\longrightarrow} A",
	],
	["Stackrel: definition", "f \\stackrel{\\text{def}}{=} g"],
	// ---- Showcase: styling features in real scientific use ----
	[
		"ML loss (annotated regularization)",
		// Machine learning — labeled, color-coded regularized loss.
		// Exercises: \mathcal, \boldsymbol on Greek + Latin, \underbrace+label,
		// \textcolor, \|⋅\|, summation, fractions.
		"\\mathcal{L}(\\boldsymbol{\\theta}) = \\underbrace{\\frac{1}{N}\\sum_{i=1}^{N}(y_i - \\hat{y}_i)^2}_{\\textcolor{blue}{\\text{data fit}}} + \\underbrace{\\lambda \\|\\boldsymbol{\\theta}\\|^2}_{\\textcolor{red}{\\text{regularizer}}}",
	],
	[
		"Polynomial simplification (cancel)",
		// Teaching algebra — show how (x-1) cancels from numerator and denominator.
		// Exercises: \cancel (twice), fractions.
		"\\frac{x^2 - 1}{x - 1} = \\frac{(x+1)\\cancel{(x-1)}}{\\cancel{(x-1)}} = x + 1",
	],
	[
		"Haber-Bosch synthesis (catalyzed)",
		// Industrial chemistry — ammonia synthesis with iron catalyst at 450°C
		// labeled above the equilibrium arrow.
		// Exercises: \stackrel, \rightleftharpoons, \text in math, subscripts.
		"\\text{N}_2 + 3\\,\\text{H}_2 \\stackrel{\\text{Fe, } 450°\\text{C}}{\\rightleftharpoons} 2\\,\\text{NH}_3",
	],
	[
		"Variance identity (color-coded sides)",
		// Probability — Var(X) = E[X²] - (E[X])² color-coded so the two halves
		// of the identity are visually distinguished.
		// Exercises: \textcolor, \mathbb (twice), \text in math.
		"\\text{Var}(X) = \\textcolor{red}{\\mathbb{E}[X^2]} - \\textcolor{blue}{(\\mathbb{E}[X])^2}",
	],
	// ---- Display equations ----
	[
		"Matrix: 2x2 pmatrix (display)",
		"\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
		true,
	],
	[
		"3x3 determinant (display)",
		"\\det \\begin{pmatrix} 1 & 2 & 3 \\\\ 4 & 5 & 6 \\\\ 7 & 8 & 9 \\end{pmatrix}",
		true,
	],
	[
		"Aligned (display)",
		"\\begin{aligned} x &= 1 \\\\ y &= 2 \\end{aligned}",
		true,
	],
	[
		"Maxwell aligned (display)",
		"\\begin{aligned} \\nabla \\cdot \\vec{E} &= \\frac{\\rho}{\\epsilon_0} \\\\ \\nabla \\cdot \\vec{B} &= 0 \\end{aligned}",
		true,
	],
	[
		"Cases (display)",
		"|x| = \\begin{cases} x & x \\geq 0 \\\\ -x & x < 0 \\end{cases}",
		true,
	],
	[
		"Quadratic formula (display)",
		"x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}",
		true,
	],
	["Tagged equation (display)", "x = 1 \\tag{1}", true],
];

// `create` produces a single paragraph (the title). The first equation pair
// (header + equation) starts after p0.
let position = 0;
let equationsInColumn = 0;
for (const [index, [header, latex, display]] of ENTRIES.entries()) {
	await cli("insert", out, "--after", `p${position}`, "--text", header);
	position += 1;
	const args = ["insert", out, "--after", `p${position}`, "--equation", latex];
	if (display) args.push("--display");
	await cli(...args);
	position += 1;
	equationsInColumn += 1;
	// Column break every 5 equations — in a 2-column section this fills col 1
	// then col 2, then naturally flows to col 1 of the next page. Skip after
	// the last equation so the doc doesn't end on a forced break.
	if (equationsInColumn === 5 && index < ENTRIES.length - 1) {
		await cli("insert", out, "--after", `p${position}`, "--column-break");
		position += 1;
		equationsInColumn = 0;
	}
}

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
