# docx-cli

**Let your AI agent edit Word documents—with tracked changes you can review.**

Fill templates, revise contracts, and leave comments in existing `.docx` files. Open the result in Word to accept or reject the edits. docx-cli edits the document’s XML in place to preserve surrounding formatting and content.

## Why choose docx-cli for Word?

In our October 10, 2026 Word-task benchmark, **docx-cli scored 18/18 successful attempts (100%) versus OfficeCLI’s 14/18 (77.8%), with 41% fewer tool calls**. The Anthropic docx skill also scored 18/18; docx-cli matched that result with **21% fewer tool calls**.

| docx-cli 0.27.0 reductions | vs. Anthropic docx skill | vs. OfficeCLI 1.0.156 |
| --- | ---: | ---: |
| Tool calls | **21% fewer** | **41% fewer** |
| Effective input tokens¹ | **22% fewer** | **53% fewer** |
| Output tokens² | **41% fewer** | **49% fewer** |
| Cumulative task time³ | **27% less** | **50% less** |

We used the same Haiku 5.5 model for six tasks, each repeated three times per tool: fill an NDA, edit an invoice, restyle a résumé, redline a contract, finalize a review, and author a journal. Reductions are calculated from unrounded averages per six-task run, then rounded to whole percentages.

¹ Cache-weighted input proxy, not dollar cost. ² Includes estimates where final token usage is unavailable. ³ Sum of task durations, not elapsed batch time.

All 54 outputs opened in Microsoft Word. All 18 docx-cli outputs received sound-layout and preserved-formatting judgments. This is our development corpus, and model judges knew which tool produced each output; success does not mean every output was defect-free. Results apply to the exact tested builds.

[Absolute measurements, ranges, per-task outcomes, and limitations](benchmarks/officecli/results/2026-10-10.md) · [Comparison protocol](benchmarks/officecli/README.md) · [Earlier benchmarks](guides/benchmarks.md)

[![Watch: Claude filling out and redlining an NDA](https://cdn.loom.com/sessions/thumbnails/da70269a970f42caa138fb3389b4b9cc-f477402396d4154d-full-play.gif#t=0.1)](https://www.loom.com/share/da70269a970f42caa138fb3389b4b9cc)

## Try it with your agent

Install the skill into your agent using [skills.sh](https://skills.sh):

```sh
npx skills add kklimuk/docx-cli
```

Choose your agent in the installer. The skill includes a bootstrap script that installs the standalone `docx` binary on first use. You need Node/npm for this skill-install command; the standalone binary does not require Bun. [Other installation options](guides/installation.md).

Download the [sample NDA](https://raw.githubusercontent.com/kklimuk/docx-cli/main/tests/fixtures/mnda.docx) into your working folder, then give your agent this prompt:

> Use the docx-cli skill to edit mnda.docx. Save a copy as mnda-reviewed.docx and keep the original unchanged. Turn on tracked changes. Replace “having a reasonable need to know” with “with a documented need to know”. Add a comment on the new wording: “Please confirm this narrower access requirement.” Keep the rest of the document unchanged and leave the revisions pending for review.

Open **mnda-reviewed.docx** in Word and choose **Review → All Markup**. You should see the wording change and an anchored comment, ready for review.

[Full walkthrough and troubleshooting](guides/first-use.md).

## Use it from the terminal

Install a standalone binary on macOS or Linux:

```sh
curl -fsSLO https://github.com/kklimuk/docx-cli/releases/latest/download/install.sh
sh install.sh
export PATH="$HOME/.local/bin:$PATH"
docx --version
```

For Windows, download `docx-windows-x64.exe` from [Releases](https://github.com/kklimuk/docx-cli/releases/latest). Already use Bun? Run `bun add -g bun-docx`. The project is **docx-cli**, its npm package is **bun-docx**, and its command is **docx**.

Try the same review yourself (macOS/Linux shell):

```sh
curl -fL https://raw.githubusercontent.com/kklimuk/docx-cli/main/tests/fixtures/mnda.docx -o mnda.docx
cp mnda.docx mnda-reviewed.docx
docx track-changes on mnda-reviewed.docx
docx replace mnda-reviewed.docx \
  "having a reasonable need to know" "with a documented need to know"
docx comments add mnda-reviewed.docx \
  --anchor "with a documented need to know" \
  --text "Please confirm this narrower access requirement."
docx read mnda-reviewed.docx --current --comments
```

Commands edit files in place unless you use `-o/--output`; keep an original copy. Reading and editing work without Microsoft Office. Optional page rendering requires Microsoft Word or LibreOffice.

## What you can do

| Task | Result |
| --- | --- |
| Review a contract | Tracked insertions/deletions and comments that can be reviewed in Word |
| Fill an existing template | Updated text and table cells using the document’s existing layout |
| Finalize a review | Accept/reject revisions, reply to comments, and resolve threads |
| Restyle a document | Changes to fonts, styles, paragraphs, tables, headers, and footers |
| Read or automate | Annotated Markdown, structured JSON, text search, diffs, and schema validation |
| Create a document | Word output from Markdown, including tables, equations, and images |

Agents discover commands through `docx --help` and address document content with locators from `docx read`. Locators can shift after structural edits; re-read before the next edit.

## Documentation

- [First document: sample, prompt, terminal steps, and troubleshooting](guides/first-use.md)
- [Installation, updates, and agent integrations](guides/installation.md)
- [Commands, locators, output formats, and architecture](guides/reference.md)
- [Agent skill](skills/docx-cli/SKILL.md)
- [Contributing](CONTRIBUTING.md) · [MIT license](LICENSE)

For exact flags and output, run `docx <command> --help`.
