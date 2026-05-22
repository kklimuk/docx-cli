# src/core/image — image emit, source resolution, formats

Three files, all reachable via `@core/image` (the [index.ts](index.ts) barrel): [formats.ts](formats.ts) (the mime↔extension table), [source.ts](source.ts) (resolve an input reference to bytes), [drawing.tsx](drawing.tsx) (emit the `<w:drawing>` + manage its media part).

## Emit + package ops

`loadImageSource(src)` in [source.ts](source.ts) resolves a path / `data:` URI / `http(s)` URL to `{bytes, extension, mimeType, pixelWidth?, pixelHeight?}` (header-parsed dims for png/jpeg/gif; other formats need explicit sizing). `computeExtentEmu` turns pixels + optional inch overrides into the `<wp:extent>` EMU pair. `addImagePart(view, source)` in [drawing.tsx](drawing.tsx) is the *operation* (writes `word/media/imageN.ext`, mints the `image` relationship, adds a content-type `<Default>`); `<Image>` is the pure emitter for the inline `<w:drawing>` run; `collectImageRuns(view)` finds drawing runs in `imgN` order (shared by `images replace`/`delete`).

The `a:`/`pic:` namespaces are declared inline on `a:graphic`/`pic:pic` (matching Word) so the subtree validates regardless of the document root's declarations — `docx create`'s template declares `w/r/m/mc/wp/...` but **not** `a`/`pic`. Drawing object ids (`wp:docPr`/`pic:cNvPr` @id) must be unique per document, so `nextDrawingId` scans existing ids. `r:embed` is an *attribute*, so there's no `r` element namespace in `jsx/index.ts`. An image run inside `<w:ins>` surfaces `ImageRun.trackedChange` (populated in the `w:drawing` branch of `readRun` in `ast/read.ts`), same as a text run.

## One format table, keyed by extension

[formats.ts](formats.ts) is the single source of truth for which image formats we extract / insert / replace. Resolution is by **file extension → canonical Word MIME** (`imageFormatForExtension`), never by sniffed MIME: Bun's `Bun.file().type` reports `image/emf`/`image/x-ms-bmp` where the docx and `extract` use `image/x-emf`/`image/bmp`, so extension-keying is what lets emf/wmf/bmp/ico round-trip. `extract` reads the part's stored MIME instead (`extensionForImageMime`). HEIC/HEIF are insert-only *input* (transcoded to JPEG in [source.ts](source.ts)), never an embedded format. **Adding a format:** add the extension→MIME and MIME→extension entries to both maps; insert/extract/replace pick it up automatically (insert needs a `readPixelDimensions` case too, or it requires `--width`/`--height`).

## Threat model — `--image` is attacker-influenced input

An AI agent driving the CLI is steered by untrusted content (a docx's text, a web page, a user prompt), so the `SRC` value flowing into `loadImageSource` is effectively attacker-controllable. Three protections in [source.ts](source.ts) make that safe:

1. **SSRF gate on remote fetches.** `loadHttp` resolves the hostname via `node:dns/promises` and refuses any address in private (RFC1918, CGN), loopback (127/8, ::1), link-local (169.254/16, fe80::/10) — including the AWS/GCP/Azure metadata endpoint at 169.254.169.254 — or reserved ranges, for both IPv4 and IPv4-mapped IPv6. Redirects are followed *manually* (`redirect: "manual"`) and the same check runs at every hop, so a benign-looking host can't 302 into `http://169.254.169.254/...`. Residual: classic DNS rebinding is still possible (TOCTOU between lookup and connect) — would require bypassing `fetch` to pin the IP at the socket layer; documented and accepted.
2. **Streaming size cap.** `Content-Length` is rejected up front if it exceeds 25 MB, but it's a server-supplied hint — `readBodyWithCap` reads the body chunk-by-chunk and aborts as soon as the running total crosses the cap, so a server that omits or under-reports `Content-Length` still can't OOM the process.
3. **SVG sanitization.** SVG is the one format here whose bytes are executable in downstream renderers (web previews, mail clients, browser-based extractors). [svg-sanitize.ts](svg-sanitize.ts) parses with `fast-xml-parser` (which rejects external-entity DOCTYPEs, killing XXE for free), then walks the tree dropping `<script>`/`<style>`/`<foreignObject>`/animation elements, stripping `on*` handlers, and filtering `href`/`xlink:href` down to same-doc fragments and **raster** `data:image/*` (never `data:image/svg+xml`, which can recurse into more script).

`loadFile` is intentionally unconfined — a local CLI legitimately reads any local path the caller supplies. The threat model doesn't change that.
