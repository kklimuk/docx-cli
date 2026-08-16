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

Every release publishes a `SHA256SUMS` manifest alongside the prebuilt binaries. Whenever an installer can compute a digest, it **verifies the binary against that manifest** and aborts on a mismatch or a missing manifest. The two differ in how they choose a version, and — importantly — in what they do on a machine with no checksum tool at all:

| | `scripts/bootstrap.sh` (the skill) | `install.sh` (standalone) |
| --- | --- | --- |
| Version | resolves the latest release **tag** and pins to it | `releases/latest` by default; set `VERSION=v1.2.3` to pin a tag |
| Checksum tool | `sha256sum`, `shasum`, or `openssl` — whichever is present | `sha256sum` or `shasum` |
| None of those on the box | **aborts** — never installs unverified | warns and installs **unverified** |

The skill's **`scripts/bootstrap.sh` fetches no scripts and executes no remote code**: it reads the GitHub releases API to resolve the latest tag, then downloads that release's binary and `SHA256SUMS` assets directly.

`install.sh` is the standalone convenience installer for humans. It ships as a **release asset** covered by that release's `SHA256SUMS`, not as a file served from the moving `main` branch, so it is immutable per release and can be verified before it runs — which is why the README fetches it to a file rather than piping it into a shell. If you'd rather skip the script entirely, use the npm registry (the package runs no install scripts) or download the binary and verify it by hand — pick the one asset for your platform, since `docx-*` would pull all five:

```sh
gh release download --repo kklimuk/docx-cli --pattern docx-darwin-arm64 --pattern SHA256SUMS
shasum -a 256 -c SHA256SUMS --ignore-missing
```

## Scope and data handling

docx-cli runs entirely **locally** against `.docx` files on disk and transmits no document content anywhere. The only network activity is:

- **`docx render`** — shells out to a locally installed Word (macOS/Windows) or LibreOffice to produce a PDF; no data leaves the machine.
- **`skills/docx-cli/scripts/bootstrap.sh`** and **`install.sh`** — fetch the prebuilt `docx` binary and its `SHA256SUMS` manifest from this repo's GitHub Releases over HTTPS.

Mutating commands overwrite the target file in place (git is the history); there is no telemetry and no external API.
