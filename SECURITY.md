# Security Policy

## Reporting a vulnerability

If you discover a security vulnerability in docx-cli, please report it privately:

- Open a [GitHub security advisory](https://github.com/kklimuk/docx-cli/security/advisories/new), or
- Email **kklimuk@gmail.com**.

Please don't open a public issue for security-sensitive reports. We aim to acknowledge within 72 hours and to ship a fix or mitigation promptly.

## Supported versions

Security fixes target the **latest published release**. Older versions are not maintained — upgrade with:

```sh
bun add -g bun-docx
# or: npm install -g bun-docx
```

On a standalone-binary install, re-run the same installer you used, or fetch the newest release asset from the [latest release](https://github.com/kklimuk/docx-cli/releases/latest). The skill's `scripts/bootstrap.sh` self-updates a stale binary.

## Install integrity

Every release publishes a `SHA256SUMS` manifest alongside the prebuilt binaries. There is one installer — `install.sh` — published two ways: as a release asset, and inside the skill folder where `scripts/bootstrap.sh` delegates to it. It downloads only release **assets** (never a script), tries `sha256sum`, `shasum`, then `openssl`, and **verifies the binary against the manifest**, aborting on a mismatch or a missing manifest.

The one behavioral difference is a parameter, not a second code path — what to do on a machine with no checksum tool at all:

| | Invoked by | No checksum tool available |
| --- | --- | --- |
| `install.sh` run by hand | you | warns and installs **unverified**, so a minimal box still has a path |
| `scripts/bootstrap.sh` | the agent skill, at session start | sets `REQUIRE_CHECKSUM=1` → **aborts**, never installs unverified |

An agent can't weigh that trade-off, so it doesn't get the lenient default. Set `REQUIRE_CHECKSUM=1` yourself to get the strict behavior by hand.

**Neither script fetches or executes remote code.** `bootstrap.sh` reads the GitHub releases API to resolve the latest tag, then runs its local sibling `install.sh`, which downloads that release's binary and `SHA256SUMS` assets directly.

Because `install.sh` is a release asset rather than a file served from the moving `main` branch, it is immutable per release and listed in that release's `SHA256SUMS` — so you can verify the installer itself before running it, which is why the README fetches it to a file instead of piping it into a shell. To skip the script entirely, use the npm registry (the package runs no install scripts) or download the binary and verify it by hand — pick the one asset for your platform, since `docx-*` would pull all five:

```sh
gh release download --repo kklimuk/docx-cli --pattern docx-darwin-arm64 --pattern SHA256SUMS
shasum -a 256 -c SHA256SUMS --ignore-missing
```

## Scope and data handling

docx-cli runs entirely **locally** against `.docx` files on disk and transmits no document content anywhere. The only network activity is:

- **`docx render`** — shells out to a locally installed Word (macOS/Windows) or LibreOffice to produce a PDF; no data leaves the machine.
- **`skills/docx-cli/scripts/install.sh`** (also published as the release asset) — fetches the prebuilt `docx` binary and its `SHA256SUMS` manifest from this repo's GitHub Releases over HTTPS. **`scripts/bootstrap.sh`** additionally reads the releases API to learn the latest tag.

Mutating commands overwrite the target file in place (git is the history); there is no telemetry and no external API.
