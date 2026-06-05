# src/cli/render — thin CLI wrapper for `docx render`

This is the CLI shell. All the real work — engine selection, docx → PDF conversion, PDF → image rasterization, output dir management, the PDFium WASM lifecycle — lives in [`@core/render`](../../core/render/CLAUDE.md). Read that file for the system playbook.

What's *here*:

- `index.ts` — `run(args)`: parses flags, validates `--engine` / `--dpi` / `--format` / `--pages`, calls `renderDocxPages` from `@core`, maps `RenderEngineError` → the CLI's structured `fail()` ack.
- `parse-pages.ts` — the `--pages N` / `--pages N-M` spec parser. Stays here because it's purely a CLI input format (the orchestrator takes a `{ first, last }` range; everything format-specific belongs to the CLI surface).

If you're adding an engine, splitter behavior, output options, or anything that the lens itself does — that's `@core/render`, not here. The CLI should keep getting *thinner*, not thicker.
