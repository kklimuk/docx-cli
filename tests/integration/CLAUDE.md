# tests/integration

## Integration tests must isolate the LibreOffice user profile

macOS `soffice --headless` locks the default user profile (`~/Library/Application Support/LibreOffice/4`); any concurrent or stale soffice causes new spawns to exit non-zero. [libreoffice-roundtrip.test.ts](libreoffice-roundtrip.test.ts) creates one `mkdtemp` profile in `beforeAll` and passes `-env:UserInstallation=file://<dir>` on every spawn. Any new test that calls `soffice` directly must follow the same pattern, or it'll flake the moment two test runs (or two test files) overlap.
