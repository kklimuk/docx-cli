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

On a macOS/Linux standalone-binary install, run **`docx upgrade`** — it replaces the binary in place, pinned and checksum-verified. (Or re-run the installer by hand from the [latest release](https://github.com/kklimuk/docx-cli/releases/latest).) On Windows, download and verify the replacement binary from Releases. The skill's `scripts/bootstrap.sh` also updates a stale binary, but only when an agent runs it: the skill tells the agent to run it when `docx` isn't installed, and nothing runs it automatically.

## Install integrity

Every release publishes a `SHA256SUMS` manifest alongside the prebuilt binaries. There is one installer — `install.sh` — delivered three ways: as a release asset, inside the skill folder where `scripts/bootstrap.sh` delegates to it, and embedded in the binary at build time for `docx upgrade`. It downloads only release **assets** (never a script), tries `sha256sum`, `shasum`, then `openssl`, and **verifies the binary against the manifest**, aborting on a mismatch or a missing manifest. (On a machine with no checksum tool, see the table below.)

macOS binaries built by the current release workflow are ad-hoc signed and
strictly signature-verified before upload; `SHA256SUMS` covers those final signed
bytes. Ad-hoc signing provides executable integrity, not Developer ID identity
or Apple notarization.

The one behavioral difference is a parameter, not a second code path — what to do on a machine with no checksum tool at all:

|                          | Invoked by                   | No checksum tool available                                           |
| ------------------------ | ---------------------------- | -------------------------------------------------------------------- |
| `install.sh` run by hand | you                          | warns and installs **unverified**, so a minimal box still has a path |
| `scripts/bootstrap.sh`   | an agent following the skill | sets `REQUIRE_CHECKSUM=1` → **aborts**, never installs unverified    |

An agent can't weigh that trade-off, so it doesn't get the lenient default. Set `REQUIRE_CHECKSUM=1` yourself to get the strict behavior by hand. The README's agent setup verifies `install.sh` against `SHA256SUMS` before running it, and runs it with this flag set.

**Neither script fetches or executes remote code.** `bootstrap.sh` resolves the latest tag from the `releases/latest` redirect (or the GitHub releases API when only `wget` is available), then runs its local sibling `install.sh`, which downloads that release's binary and `SHA256SUMS` assets directly.

Because `install.sh` is a release asset rather than a file served from the moving `main` branch, it is immutable per release and listed in that release's `SHA256SUMS` — so you can verify the installer itself before running it, which is why the README fetches it to a file instead of piping it into a shell. To skip the script entirely, use the npm registry (the package runs no install scripts) or download the binary and verify it by hand — pick the one asset for your platform, since `docx-*` would pull all five:

```sh
gh release download --repo kklimuk/docx-cli --pattern docx-darwin-arm64 --pattern SHA256SUMS
shasum -a 256 -c SHA256SUMS --ignore-missing
```

## Scope and data handling

docx-cli runs entirely **locally** against `.docx` files on disk. **No command sends document content anywhere** — there is no telemetry and no external API.

**`docx render` does no network I/O.** It shells out to a locally installed Word (macOS/Windows) or LibreOffice to produce a PDF, then rasterizes in-process via a bundled WASM package. It leaves the docx-cli process, but not the machine.

Network access happens in three places in docx-cli itself (plus your package manager, if you install through npm), and none of them uploads anything:

- **`skills/docx-cli/scripts/install.sh`** (also published as the release asset) — fetches the prebuilt `docx` binary and its `SHA256SUMS` manifest from this repo's GitHub Releases over HTTPS. **`scripts/bootstrap.sh`** additionally reads the releases page to learn the latest tag.
- **`docx upgrade`** — downloads a release binary, and only when you run it. It carries `install.sh` **embedded at build time** (never fetched) and executes that copy with `REQUIRE_CHECKSUM=1`, so the new binary is pinned to a tag and SHA-256-verified before it replaces anything. It refuses to touch a package-manager-owned install. Note the installer it runs is the one embedded in the release you are _currently on_, not the target's — so an installer bug shipped in vN can't be fixed by upgrading from vN; fetch that release's `install.sh` by hand instead (see above).
- **`docx images add --image https://…`** — the one verb that fetches _content_ you asked for, when and only when you pass an `http(s)` source (a local path or `--markdown` with a local ref never does). It is guarded against SSRF: private/link-local addresses are blocked, redirects are walked manually and re-checked, capped at 5 hops.

Your document content is never sent anywhere by any of them.

Mutating commands overwrite the target file in place (git is the history).

## Deploying in an organization

For IT and security teams rolling docx-cli out to many machines.

**What it touches.** One file: the `docx` binary (default `~/.local/bin/docx`; set `PREFIX` to choose). It keeps no config, cache, or state of its own, and it writes only the documents it's asked to edit. The one exception is `docx render`, covered below.

**Network endpoints.** Allow these only where you want people to install or upgrade:

| Endpoint                                                                                                  | Used by                                           | When                                                                                                              |
| --------------------------------------------------------------------------------------------------------- | ------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `github.com` and `release-assets.githubusercontent.com` (older releases: `objects.githubusercontent.com`) | `install.sh`, `bootstrap.sh`, `docx upgrade`      | installing or upgrading                                                                                           |
| `api.github.com`                                                                                          | `bootstrap.sh`                                    | only when `curl` is missing and it falls back to `wget`                                                           |
| `registry.npmjs.org`                                                                                      | `bun add -g bun-docx` / `npm install -g bun-docx` | installing via npm (dependencies resolve at install time, so this route is not hash-pinned the way the binary is) |
| any URL a user passes                                                                                     | `docx images add --image https://…`               | inserting a remote image                                                                                          |

Block all of them and every command except remote image insertion still works.

**Nothing updates in the background.** The binary changes only when someone runs `docx upgrade`, the installer, or the skill's `bootstrap.sh`. The skill tells agents to run `bootstrap.sh` when `docx` is missing, and bootstrap also upgrades a binary that is behind the latest release. To keep agents from moving a pinned fleet, set one of these in the agent's environment (for example, through managed settings):

- `DOCX_CLI_VERSION=v0.27.0`: bootstrap installs exactly that release when `docx` is missing and never updates a binary that's already present.
- `DOCX_CLI_NO_UPDATE=1`: bootstrap never updates a binary that's already present. (A missing one still installs, from the latest release; set `DOCX_CLI_VERSION` too to control which.)

These variables govern `bootstrap.sh`, the path the skill sends agents down when either one is set. They don't stop someone from running `docx upgrade` or `npm install -g bun-docx` directly; restrict those through your normal software policy if you need to.

Offline, bootstrap keeps the working binary and exits 0.

**Pinning a version.**

- Install a specific release with `VERSION=v0.27.0 sh install.sh`, or `docx upgrade --to v0.27.0`, and set `DOCX_CLI_VERSION` (above) so the skill's bootstrap respects the pin.
- Use release-tag URLs (`…/releases/download/v0.27.0/…`) rather than `latest`.
- The Claude Code and Codex plugin installs follow this repository's default branch, so the skill text can change when the plugin updates. To pin the skill too, copy `skills/docx-cli/` from a release tag into your managed skills directory instead of installing the plugin.

**Air-gapped or mirrored installs.**

1. On a connected machine, download `docx-<platform>` and `SHA256SUMS` from a tagged release and verify them with `shasum -a 256 -c SHA256SUMS --ignore-missing`.
2. Copy the binary, plus `NOTICES` and `LGPL-3.0.txt` from the same release, to each machine (or your software-distribution tool), mark the binary executable, and put it on `PATH`.

Nothing in docx-cli phones home after that.

**Windows.**

- `docx-windows-x64.exe` is not Authenticode-signed, so SmartScreen, WDAC, or AppLocker may block it. Allowlist it by hash.
- Verify it with `Get-FileHash docx-windows-x64.exe -Algorithm SHA256` against `SHA256SUMS`.
- `install.sh` also works under Git Bash/MSYS.
- `docx upgrade` works from Git Bash. A running `.exe` can't be overwritten, so the installer writes to a staging folder beside it, and only the verified binary is swapped in; the old one is left as `docx.exe.old` until the next upgrade.

**macOS.** Release binaries are ad-hoc signed and signature-verified before checksumming, but not Developer ID-signed or notarized. Downloading with `curl` (as the installer does) doesn't set the quarantine flag. A browser download does, so Gatekeeper will ask.

**`docx render` (optional).** It's the only command that launches another application: Word (macOS/Windows) or LibreOffice, to turn the document into page images. On macOS it drives Word through AppleScript, which triggers a one-time Automation consent prompt (pre-approve it with a PPPC profile, or install only LibreOffice). It stages a temporary copy of the document, in Word's sandbox `Documents` folder on macOS or the system temp directory otherwise, and deletes it afterwards. Every other command works without it.

**What the checksum does and doesn't prove.** `SHA256SUMS` is published in the same release as the binaries. It protects against a corrupted or swapped download, not against a compromised release. For higher assurance, build from source (`bun run build:binary`) at a reviewed tag.

**License.** MIT. The standalone binaries statically link libheif (LGPL-3.0) for HEIC support; [NOTICES](NOTICES) lists the attributions and your LGPL rights, and redistributing the binary internally is fine with those files alongside it.
