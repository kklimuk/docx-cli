# src/cli/create

## docx create must ship Word-canonical parts

Word treats `word/styles.xml`, `word/settings.xml`, `word/fontTable.xml`, `word/webSettings.xml`, `word/theme/theme1.xml`, and `docProps/app.xml` as required even though ECMA-376 marks them optional. Without them Word shows an "unreadable content / recover?" prompt on open. LibreOffice is more permissive — the integration suite passed for years while Word users hit the warning.

All six are baked into [canonical-parts.ts](canonical-parts.ts) (the theme is a static file under `canonical/theme1.xml`, the rest inline string constants). The create flow writes them via the `CANONICAL_PARTS` loop; hand-rolled fixture scripts get them via `addCanonicalParts` in [tests/fixtures/setup/helpers/index.ts](../../../tests/fixtures/setup/helpers/index.ts). Both iterate `Object.values(CANONICAL_PARTS)` for content-type overrides and relationships, so adding a part is usually a single edit to that map (set the right `scope`; `buildContentTypes`/`buildRootRels`/`buildDocumentRels` pick it up). New JSX tag names in a part body may need additions to the `W_TAGS` / namespace lists in `src/core/jsx/index.ts`.

The static templates in [template.tsx](template.tsx) carry no user-supplied content, so they don't need escaping.
