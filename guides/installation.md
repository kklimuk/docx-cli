# Installation and agent setup

[Back to the quick start](../README.md)

## Install

**Bun package (npm registry)** — requires Bun >= 1.3:

```sh
bun add -g bun-docx
# or run without installing:
bunx bun-docx read doc.docx
```

**Standalone binary** (no Bun required). Every release publishes prebuilt binaries, the installer, and a `SHA256SUMS` manifest covering all of them; the installer **verifies the binary's SHA-256** before installing:

```sh
curl -fsSLO https://github.com/kklimuk/docx-cli/releases/latest/download/install.sh
sh install.sh
```

`install.sh` is a release asset, not a file served from a moving branch — so it is immutable per release and listed in that release's `SHA256SUMS`. (Published from v0.23.0 onward; on earlier releases use the npm path above.) To verify it before running it:

```sh
curl -fsSLO https://github.com/kklimuk/docx-cli/releases/latest/download/SHA256SUMS
shasum -a 256 -c SHA256SUMS --ignore-missing   # or: sha256sum -c …
sh install.sh
```

macOS artifacts built by the current release workflow are ad-hoc signed and signature-verified before checksumming; they are not Developer ID-signed or notarized.

Honors `PREFIX` (default `$HOME/.local/bin`) and `VERSION` (default `latest`). Pre-built for linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64.

Once installed, **`docx upgrade`** updates a standalone binary — it replaces the binary wherever it already lives (`PREFIX` does not apply), running the same installer embedded in the binary at build time rather than fetched, so the download stays pinned to a release tag and SHA-256-verified. `--to v0.23.0` pins a version, `--dry-run` reports what would change. On an npm/bun install it tells you to use the package manager instead.

Prefer to skip the script entirely? Download `docx-<platform>` + `SHA256SUMS` from the [latest release](https://github.com/kklimuk/docx-cli/releases/latest), verify, `chmod +x`, and put it on `PATH`.

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
