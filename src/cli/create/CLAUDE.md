# src/cli/create

## docx create is a thin wrapper over `core/create`

The CLI handler in [index.tsx](index.tsx) parses args, then calls `buildBlankPackage({ path, title, author, text })` from [@core/create](../../core/create/index.ts) and saves. The OOXML templating + canonical-parts assembly lives in `src/core/create/`.

## Word-canonical parts

Word treats `word/styles.xml`, `word/settings.xml`, `word/fontTable.xml`, `word/webSettings.xml`, `word/theme/theme1.xml`, and `docProps/app.xml` as required even though ECMA-376 marks them optional. Without them Word shows an "unreadable content / recover?" prompt on open. LibreOffice is more permissive — the integration suite passed for years while Word users hit the warning.

All six are baked into [@core/create/canonical-parts.ts](../../core/create/canonical-parts.ts) (the theme is a static file under `core/create/canonical/theme1.xml`, the rest are inline string constants). `buildBlankPackage` writes them via the `CANONICAL_PARTS` loop; hand-rolled fixture scripts get them via `addCanonicalParts` in [tests/fixtures/setup/helpers/index.ts](../../../tests/fixtures/setup/helpers/index.ts). Both iterate `Object.values(CANONICAL_PARTS)`, so adding a part is usually a single edit there. New JSX tag names in a part body may need additions to the `W_TAGS` / namespace lists in [src/core/jsx/index.ts](../../core/jsx/index.ts).

The static templates in [@core/create/template.tsx](../../core/create/template.tsx) carry no user-supplied content, so they don't need escaping; `documentXml(text?)` and `corePropertiesXml({ title, author, now })` are the two functions that interpolate caller-provided strings, and they emit via JSX (which escapes).
