# src/core/render — `renderDocxPages`: docx → per-page PNGs / JPGs

`@core/render` is docx-cli's **only subsystem that depends on an external runtime** — it shells out to Microsoft Word or LibreOffice to produce a PDF, then rasterizes the PDF in-process via the bundled PDFium WASM. The lens exists because the CLI is consumed by AI agents that can read PNGs — having a ground-truth visual is qualitatively different from inspecting OOXML or the typed AST.

## Layout

```
src/core/render/
  index.ts             — barrel + `renderDocxPages` orchestrator (the one
                         high-level entry point the CLI calls)
  detect.ts            — engine registry + auto-select (`detectEngine`,
                         `engineByName`, `listAvailable`)
  engines/
    types.ts           — `RenderEngine` interface, `RenderEngineError`
    word-mac.ts        — Microsoft Word via `osascript`, staged through the
                         Word sandbox Container
    word-windows.ts    — Microsoft Word via PowerShell COM
                         (`ExportAsFixedFormat`)
    libreoffice.ts     — `soffice --headless --convert-to pdf`
                         (cross-platform)
  split-pdf.ts         — PDF → PNG/JPG via @hyzyla/pdfium (WASM) + pngjs /
                         jpeg-js
  assets.d.ts          — ambient `*.wasm` module declaration so the
                         `with { type: "file" }` import in split-pdf.ts
                         type-checks

src/cli/render/
  index.ts             — thin: arg-parse + engine-name validation, calls
                         `renderDocxPages` and maps `RenderEngineError` to
                         the CLI's `fail()` ack
  parse-pages.ts       — `--pages N` / `--pages N-M` spec parser (CLI-only)
```

## API surface

```ts
import { renderDocxPages, type RenderOptions } from "@core";

const result = await renderDocxPages(docxPath, {
  outDir,                       // required
  engine,                       // optional — falls back to detectEngine()
  dpi, format, range,           // mirror the CLI flags
});
// result: { engine: "word-mac" | "word-win" | "libreoffice",
//           outDir: string, pages: string[] }
```

Leaf primitives are exported too for callers that need finer control:

- `detectEngine()` — async; returns the highest-priority available engine, or `undefined`
- `engineByName(name)` — sync lookup by name (`"word"` resolves to the platform-appropriate Word variant)
- `listAvailable()` — async list of names that pass `available()` on this machine
- `splitPdf(pdfPath, outDir, options)` — PDF → image array, returns the page paths
- The engines themselves are named exports (`libreofficeEngine`, `wordMacEngine`, `wordWindowsEngine`)

## Two stages: docx → PDF (external) → PNG (in-process)

The pipeline is intentionally split:

1. **docx → PDF**: drive Word (macOS / Windows) or LibreOffice (cross-platform) via the `RenderEngine` abstraction. This **requires** Word or LibreOffice to be installed locally — there's no pure-JS docx renderer with high enough fidelity to drop in here.
2. **PDF → PNG/JPG**: rasterize via `@hyzyla/pdfium` — MIT wrapper + Apache-2.0 PDFium WASM binary, ~11 MB shipped inside the package's npm tarball. **No system tools required** for this half. Pixel data comes back from PDFium in RGBA order (BGRA colorspace + `REVERSE_BYTE_ORDER` flag set internally by @hyzyla — see comment in `splitPdf`) and gets piped through `pngjs` (PNG) or `jpeg-js` (JPG) for encoding.

The choice of "bundle PDFium + JS encoders, shell out only for docx → PDF" means an `npm install bun-docx` gives the user everything except the docx renderer — and they almost certainly have Word or LibreOffice already, since they're working with .docx files.

## Engine selection priority

`--engine auto` (default): Word for Mac → Word for Windows → LibreOffice, **in that order**. Word is preferred because it's the ground-truth renderer most users/agents care about, and our integration tests already use `scripts/word-redlines.sh` to confirm Word's behavior matches the spec. LibreOffice is the cross-platform fallback.

Each engine's `available()` check is **cheap** — `existsSync` on the app bundle / container dir for Word-mac, a tiny COM instantiate-then-quit probe for Word-win, `command -v soffice` plus canonical-path fallbacks for LibreOffice. Don't make `available()` actually render anything.

## Word-mac sandbox staging

Word for Mac is sandboxed: arbitrary file paths trigger a "Grant File Access" prompt. We copy the input docx into `~/Library/Containers/com.microsoft.Word/Data/Documents/.docx-cli-render-<pid>-<n>.docx`, drive Word to save-as-PDF inside the same container, then move the PDF out. Same staging pattern as [scripts/word-redlines.sh](../../../scripts/word-redlines.sh) — both first runs on a machine trigger a one-time macOS Automation TCC prompt that has to be granted to the controlling terminal.

## LibreOffice profile isolation

`soffice --headless` on macOS locks the default user profile (`~/Library/Application Support/LibreOffice/4`). Two concurrent invocations (or one stale soffice process) make subsequent spawns exit non-zero. We mint a temp profile per call via `-env:UserInstallation=file://<tmpdir>` and `rm -rf` it after. Mirrors [tests/integration/libreoffice-roundtrip.test.ts](../../../tests/integration/libreoffice-roundtrip.test.ts)'s worker-pool pattern; **any new soffice spawn in this codebase must do the same**.

## PDF rasterization details

`split-pdf.ts` imports `@hyzyla/pdfium`, `pngjs`, and `jpeg-js` at the top of the file. Two of these (`pngjs`, `jpeg-js`) are pure JS and inline normally into the bundle. The PDFium WASM binary needs special handling:

```ts
import pdfiumWasmPath from "@hyzyla/pdfium/dist/pdfium.wasm" with {
  type: "file",
};
// …
const wasmBinary = await Bun.file(resolve(import.meta.dir, pdfiumWasmPath)).arrayBuffer();
const lib = await PDFiumLibrary.init({ wasmBinary });
```

The `with { type: "file" }` attribute tells `bun build` to emit the asset and replace the import with a path string. Three runtime modes, all working:

- **`bun src/index.ts` (dev)** — `pdfiumWasmPath` is the absolute node_modules path; `Bun.file()` reads it directly.
- **`bun dist/index.js` (bundled)** — `bun build` emits `dist/pdfium-<hash>.wasm` next to `dist/index.js` and the import resolves to the relative filename `./pdfium-<hash>.wasm`. The `resolve(import.meta.dir, ...)` call joins that against the bundle's own directory so the path works from any cwd.
- **`./dist/docx` (compiled binary)** — `bun build --compile` embeds the WASM inside the binary; the import resolves to a virtual `/$bunfs/root/pdfium-<hash>.wasm` path that `Bun.file()` reads transparently.

This pattern (build-time asset + path import + runtime `Bun.file(arrayBuffer)`) is what lets us hand the bytes to `PDFiumLibrary.init({ wasmBinary })` — the default entry's own `fs.readFileSync("pdfium.wasm")` path fails under `--compile` because the relative resolution doesn't find the embedded file. The `/browser/base64` alternative entry is **not** usable here: it's compiled with Emscripten's `ENVIRONMENT=web` and panics under Node/Bun.

`--pages SPEC` accepts a single page (`5`) or contiguous range (`2-7`). Discontinuous specs (`1,3,5`) are out of scope — they'd split the output paths and complicate the JSON ack's `pages: []` array semantics. Run the command multiple times for discrete pages.

## Output is JSON ack via `respond()` (always-print)

The page list is essential output the agent can't reconstruct, so we call `respond()` directly (always prints) rather than `respondAck()` (gated on `--verbose`). Same pattern as `comments add --batch` printing minted ids unconditionally. `--verbose` switches the output from the bare page-path list to the full JSON ack (`{ok, operation, path, engine, output, pages}`).

## Adding a new engine

1. Add a file under `engines/` exporting a `RenderEngine`.
2. Append it to `ENGINES` in `detect.ts` in **priority order** (highest fidelity first; cross-platform last).
3. Extend the `RenderEngineName` union in `engines/types.ts`.
4. Update the help text in `index.ts` (`HELP` constant) and the engine-name validation in `resolveEngine`.
5. Add an availability probe — must be cheap, no actual conversion.
6. Cover via an integration test gated on the engine being present.

## Why only one rasterizer

Earlier iterations supported `pdftoppm` (poppler) + `magick` (ImageMagick) as PDF→image fallbacks. We dropped both. With PDFium bundled, every user has a working rasterizer out of the box; supporting two more code paths added complexity (file-naming normalization across tools, platform-specific tool discovery, a `DOCX_RENDER_SPLITTER` env var) for no incremental capability. If PDFium ever ships a regression that breaks our case, the right move is pinning the @hyzyla/pdfium version, not bringing the fallbacks back.

## License attribution

PDFium is dual-licensed BSD-3-Clause OR Apache-2.0 (we redistribute under BSD-3). The required notice lives in [`NOTICES`](../../../NOTICES) at the repo root, which the release workflow stages into each artifact bundle and the npm `files` list ships alongside `dist/index.js`. If the bundled WASM binary's source changes (we don't modify it — we consume @hyzyla/pdfium's published build), update NOTICES accordingly.

The same NOTICES file carries the BSD-3 attributions for `jpeg-js` and `highlight.js`, the LGPL-3.0 acknowledgement for `libheif-js` (statically linked into the compiled binary via `heic-convert`), and the MIT election for `jszip` (which is dual MIT-or-GPL).

## CI considerations

The end-to-end render tests in [tests/cli/render.test.ts](../../../tests/cli/render.test.ts) gate on `LIBREOFFICE_AVAILABLE` (just soffice; the rasterizer is bundled, so it's always present). On a clean CI runner without LibreOffice they auto-skip. Pure-logic tests (`--pages` parser, USAGE errors) always run since they don't touch the runtime.

The integration suite ([tests/integration/libreoffice-roundtrip.test.ts](../../../tests/integration/libreoffice-roundtrip.test.ts)) already requires soffice; render uses the same dep so CI infrastructure cost is zero on top.

If we ever want CI to render through Word for verification, we'd need a Windows runner with Office installed — significantly more expensive. Don't promise it as a CI feature; agents can render locally during a development loop and let CI stay LibreOffice-only.
