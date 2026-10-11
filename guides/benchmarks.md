# Benchmark evidence

[Back to the quick start](../README.md)

The latest [October 10, 2026 comparison](../benchmarks/officecli/results/2026-10-10.md) covers docx-cli 0.27.0, the Anthropic docx skill, and OfficeCLI 1.0.156. docx-cli and the Anthropic skill each scored 18/18 task successes; OfficeCLI scored 14/18. The report includes effort measurements, per-task consistency, defects, and provenance.

## Historical Anthropic-skill experiments

The numbers below are earlier project-run measurements, retained as history rather than current performance claims. The corpus is small and has informed docx-cli development; treat it as a regression corpus, not an independent estimate of all Word editing tasks.

Agents can edit Word documents by unpacking the `.docx` and modifying its OOXML. docx-cli instead exposes commands and an annotated Markdown read view, reducing the XML work required for common edits. This benchmark tested whether that difference helped agents finish document tasks.

We measured it — a controlled A/B bake-off: **six real document tasks** (fill an NDA, fill an invoice, restyle a résumé, redline a contract, finalize a contract, author a journal), the same starting files, and one separate model judge grading every result from the **Word-rendered pages**. Three runs per arm at each of two model tiers. Haiku columns: July 2026 rerun on v0.20.4; Sonnet columns: June 2026 bake-off (**not** Sonnet 5):

|                            | Haiku (weak, cheap) |                       | Sonnet (strong, June run — not Sonnet 5) |                       |
| :------------------------- | ------------------: | --------------------: | ---------------------------------------: | --------------------: |
|                            |        **docx-cli** |         default skill |                             **docx-cli** |         default skill |
| Tasks solved (of 6)        |       **5.0** (5–5) |             0.7 (0–2) |                             **6.0** (6–6) |             4.0 (4–4) |
| Rendered correctly (of 6)  |             **6.0** |                   4.3 |                                      6.0 |                   4.7 |
| Outright-broken documents  |               **0** |    ~1.3/run (up to 2) |                                    **0** |                     0 |
| Effective input tokens     |            **3.2M** |           7.9M (2.5×) |                                 **1.6M** |           3.6M (2.2×) |
| Output tokens              |             **78k** |           155k (2.0×) |                                        — |                     — |
| Wall-clock                 |         **1,186 s** | 2,438 s (2.1× slower) |                              **1,175 s** | 2,029 s (1.7× slower) |

- **In this six-task corpus, the correctness gap was widest on Haiku (~7×)**. In the June Sonnet runs, the default skill solved 4/6 tasks, missing the contract redline and résumé tasks in each run. This does not establish performance on other documents or newer models.
- **Both tested model tiers used fewer tokens and finished faster with docx-cli** — ~2.2–2.5× more input tokens and ~1.7–2.1× slower at *both* tiers, with token/time ranges that never overlap. (Output tokens, measured in the July rerun: 2× more for the default skill. The June Sonnet run measured input only.)
- **Word couldn't reliably open the default skill's work** — in the July Haiku runs it failed to open 4 of the skill's 18 outputs (both legal-redline deliverables in 2 of 3 runs); all 18 of docx-cli's opened on the first try.

Full methodology, per-task rubrics, and the harness that produced these numbers: [`.claude/skills/weak-agent-test`](../.claude/skills/weak-agent-test).

The table preserves the previously reported results; its token totals have not been recalculated with the current collector. “Effective input tokens” is a cache-weighted proxy, not a dollar-cost measurement. New comparisons should retain fresh-input, cache-write, cache-read, and output tokens separately, together with exact model IDs and the collector version.

## October comparisons

The [October 10 results](../benchmarks/officecli/results/2026-10-10.md) supersede the [October 8 comparison](../benchmarks/officecli/results/2026-10-08.md) for the README. Both preserve their own per-run data and limitations. See the [comparison protocol](../benchmarks/officecli/README.md). These experiments use different builds and snapshots; do not compare their token totals directly.
