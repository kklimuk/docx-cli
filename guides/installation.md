# Installation and agent setup

[Back to the quick start](../README.md)

## Install

**Bun package (npm registry)** — the simplest path with Bun >= 1.3:

```sh
bun add -g bun-docx
# or run without installing:
bunx bun-docx read doc.docx
```

**Standalone binary** (macOS/Linux, no Bun required). Every release publishes prebuilt binaries, the installer, and a `SHA256SUMS` manifest covering all of them. Download and verify the installer before running it, using the same release tag throughout:

```sh
DOCX_VERSION=v0.27.0
curl -fsSLO "https://github.com/kklimuk/docx-cli/releases/download/$DOCX_VERSION/install.sh" &&
curl -fsSLO "https://github.com/kklimuk/docx-cli/releases/download/$DOCX_VERSION/SHA256SUMS" &&
shasum -a 256 -c SHA256SUMS --ignore-missing &&
VERSION="$DOCX_VERSION" REQUIRE_CHECKSUM=1 sh install.sh
```

Use `sha256sum -c SHA256SUMS --ignore-missing` if your system has `sha256sum` instead of `shasum`. Continue only when verification reports `install.sh: OK`; the `&&` chain stops on a failed download or checksum. The installer then downloads the prebuilt binary and **verifies its SHA-256 before installing**. `REQUIRE_CHECKSUM=1` makes it refuse installation if no checksum tool is available.

Choose a tag from [Releases](https://github.com/kklimuk/docx-cli/releases) to change `DOCX_VERSION`. `install.sh` has been published as a release asset since v0.23.0. The installer honors `PREFIX` (default `$HOME/.local/bin`) and `VERSION` (default `latest`); the example pins both the installer and binary to the same tag. Add the destination to your `PATH`:

```sh
export PATH="$HOME/.local/bin:$PATH"
docx --version
```

macOS artifacts built by the current release workflow are ad-hoc signed and signature-verified before checksumming; they are not Developer ID-signed or notarized.

Prebuilt binaries are available for linux/x64, linux/arm64, darwin/x64, darwin/arm64, and windows/x64.

On macOS/Linux, **`docx upgrade`** updates a standalone binary — it replaces the binary wherever it already lives (`PREFIX` does not apply), running the same installer embedded in the binary at build time rather than fetched, so the download stays pinned to a release tag and SHA-256-verified. `--to v0.27.0` pins a version, `--dry-run` reports what would change. On an npm/bun install it tells you to use the package manager instead.

To install manually, download the binary for your platform and `SHA256SUMS` from [Releases](https://github.com/kklimuk/docx-cli/releases/latest), verify the checksum, and put the binary on `PATH`. On macOS/Linux, rename it to `docx` and run `chmod +x docx`; on Windows, use `docx-windows-x64.exe` (rename it to `docx.exe`). See [install integrity](../SECURITY.md#install-integrity) for the verification policy.

## Use as an agent skill

docx-cli ships as an [Agent Skill](https://agentskills.io) — one `SKILL.md` that works across Claude Code, Codex, Pi, and the other harnesses that read the open skill format. The skill teaches the locator model and the redline / comment / fill workflows, then uses `docx <command> --help` for the installed version’s flags and output.

### Install

**Any agent ([skills.sh](https://skills.sh))** — one cross-harness command, installs into whichever agent you're using:

```sh
npx skills add kklimuk/docx-cli
```

**Claude Code plugin:**

```
/plugin marketplace add kklimuk/docx-cli
/plugin install docx-cli@docx-cli
```

**Codex** — add the marketplace (the plugin's skills auto-discover):

```
codex plugin marketplace add kklimuk/docx-cli
```

**Pi** — one-command install (the `pi` manifest in `package.json` pulls in the skill), then invoke `/skill:docx-cli`:

```
pi install git:github.com/kklimuk/docx-cli      # global; add -l for a project (team-shared) install
# manual alternative: pi --skill /path/to/docx-cli/skills/docx-cli
```

**Any harness / manual** — drop [`skills/docx-cli/`](../skills/docx-cli) into your agent's skills directory (e.g. `~/.claude/skills/` or the cross-tool `~/.agents/skills/`). On first activation the skill's [`scripts/bootstrap.sh`](../skills/docx-cli/scripts/bootstrap.sh) installs the `docx` binary (and self-updates a stale one).

### Keeping the skill current

The binary is the source of truth: `docx info skill` prints the canonical `SKILL.md` for the installed version, and a CI test fails if the committed copy drifts. Regenerate after any change with:

```
docx info skill > skills/docx-cli/SKILL.md
```
