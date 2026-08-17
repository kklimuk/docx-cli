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

On a standalone-binary install, run **`docx upgrade`** — it replaces the binary in place, pinned and checksum-verified. (Or re-run the installer by hand from the [latest release](https://github.com/kklimuk/docx-cli/releases/latest).) The skill's `scripts/bootstrap.sh` self-updates a stale binary at session start.

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

docx-cli runs entirely **locally** against `.docx` files on disk. **No command sends document content anywhere** — there is no telemetry and no external API.

**`docx render` does no network I/O.** It shells out to a locally installed Word (macOS/Windows) or LibreOffice to produce a PDF, then rasterizes in-process via a bundled WASM package. It leaves the docx-cli process, but not the machine.

Network access happens in three places, and none of them uploads anything:

- **`skills/docx-cli/scripts/install.sh`** (also published as the release asset) — fetches the prebuilt `docx` binary and its `SHA256SUMS` manifest from this repo's GitHub Releases over HTTPS. **`scripts/bootstrap.sh`** additionally reads the releases page to learn the latest tag.
- **`docx upgrade`** — downloads a release binary, and only when you run it. It carries `install.sh` **embedded at build time** (never fetched) and executes that copy with `REQUIRE_CHECKSUM=1`, so the new binary is pinned to a tag and SHA-256-verified before it replaces anything. It refuses to touch a package-manager-owned install. Note the installer it runs is the one embedded in the release you are *currently on*, not the target's — so an installer bug shipped in vN can't be fixed by upgrading from vN; fetch that release's `install.sh` by hand instead (see above).
- **`docx images add --image https://…`** — the one verb that fetches *content* you asked for, when and only when you pass an `http(s)` source (a local path or `--markdown` with a local ref never does). It is guarded against SSRF: private/link-local addresses are blocked, redirects are walked manually and re-checked, capped at 5 hops.

Your document content is never sent anywhere by any of them.

Mutating commands overwrite the target file in place (git is the history).
