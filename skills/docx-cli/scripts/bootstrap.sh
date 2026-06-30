#!/bin/sh
# docx-cli skill bootstrap.
#
# Ensures the `docx` binary is installed and reasonably current, so an agent that
# activated this skill can rely on it. Run it once at the start of a session:
#   bash scripts/bootstrap.sh
#
# Supply-chain posture (why this is NOT a `curl | sh`):
#   - It resolves the latest RELEASE TAG (not the moving `main` branch).
#   - It downloads that tag's install.sh to a FILE (never pipes a remote script into a
#     shell), then runs it pinned to that exact version.
#   - install.sh downloads the prebuilt binary and VERIFIES its SHA-256 against the
#     release's published SHA256SUMS before installing.
#
# Behavior:
#   - not installed -> resolve latest tag, install it (pinned + checksum-verified)
#   - installed     -> compare against the latest release; update only if BEHIND
#   - offline / can't resolve latest -> keep the working binary and exit 0 (every verb
#     except `render` works offline anyway)
#
# POSIX sh, no bashisms. The binary is the source of truth — this script only keeps it
# present and fresh; it never edits the skill.

set -eu

REPO="kklimuk/docx-cli"
API_LATEST="https://api.github.com/repos/${REPO}/releases/latest"

# ─── Pick a downloader (fetch a URL to stdout) ───
if command -v curl >/dev/null 2>&1; then
  fetch() { curl -fsSL "$1"; }
elif command -v wget >/dev/null 2>&1; then
  fetch() { wget -qO- "$1"; }
else
  echo "docx-cli bootstrap: need curl or wget." >&2
  exit 1
fi

# Resolve the latest release tag (e.g. "v0.19.1"), best-effort; empty on failure.
resolve_latest_tag() {
  fetch "$API_LATEST" 2>/dev/null \
    | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1
}

# Install the given release tag: download THAT TAG's install.sh to a file (no
# pipe-to-shell), then run it pinned to the tag so it fetches and SHA-256-verifies the
# matching binary. Returns the installer's exit status.
run_installer() {
  tag="$1"
  installer_url="https://raw.githubusercontent.com/${REPO}/${tag}/install.sh"
  echo "→ Installing docx-cli ${tag} (pinned, checksum-verified) ..."
  installer="$(mktemp 2>/dev/null || mktemp -t docx-cli-install)"
  if ! fetch "$installer_url" > "$installer" || [ ! -s "$installer" ]; then
    rm -f "$installer"
    echo "docx-cli bootstrap: could not download the installer from $installer_url (offline or rate-limited)." >&2
    return 1
  fi
  if VERSION="$tag" sh "$installer"; then rc=0; else rc=$?; fi
  rm -f "$installer"
  return "$rc"
}

# install.sh drops the binary in ${PREFIX:-$HOME/.local/bin} and only PRINTS a PATH hint
# — it can't edit the caller's PATH. So after installing, confirm `docx` is resolvable
# and fail LOUDLY if not, rather than reporting a false success the agent trips over on
# its first `docx` call (and re-downloading every session).
ensure_reachable() {
  command -v docx >/dev/null 2>&1 && return 0
  bindir="${PREFIX:-$HOME/.local/bin}"
  echo "docx-cli bootstrap: installed to $bindir, but it is NOT on your PATH." >&2
  echo "  Add it:  export PATH=\"$bindir:\$PATH\"   (then re-run)" >&2
  echo "  Or invoke the binary directly: $bindir/docx" >&2
  return 1
}

# ─── Not installed: resolve tag, install pinned + verified, finish ───
if ! command -v docx >/dev/null 2>&1; then
  echo "docx not found on PATH."
  tag="$(resolve_latest_tag)"
  if [ -z "$tag" ]; then
    echo "docx-cli bootstrap: could not resolve the latest release (offline or rate-limited) — cannot install safely." >&2
    exit 1
  fi
  run_installer "$tag" || exit 1
  ensure_reachable || exit 1
  exit 0
fi

installed="$(docx --version 2>/dev/null | awk '{print $NF}')"
echo "docx-cli present: ${installed:-unknown}"

# Couldn't read a version? Don't enter a reinstall loop — leave the present binary.
if [ -z "$installed" ]; then
  echo "Could not read the installed version — leaving the present binary in place."
  exit 0
fi

tag="$(resolve_latest_tag)"
latest="$(printf '%s' "$tag" | sed 's/^v//')"
if [ -z "$latest" ]; then
  echo "Could not determine the latest release (offline or rate-limited) — keeping the installed binary."
  exit 0
fi

if [ "$installed" = "$latest" ]; then
  echo "✓ Up to date (${installed})."
  exit 0
fi

# Versions differ. Only UPDATE when installed is OLDER than latest — never downgrade a
# locally-built/pre-release binary ahead of the published release. Use version sort when
# the platform's `sort` supports -V; otherwise fall back to updating.
if printf '%s\n' "0.0" "0.1" | sort -V >/dev/null 2>&1; then
  oldest="$(printf '%s\n%s\n' "$installed" "$latest" | sort -V | head -n 1)"
  if [ "$oldest" != "$installed" ]; then
    echo "Installed (${installed}) is newer than the latest release (${latest}) — keeping it."
    exit 0
  fi
fi

echo "A newer release is available: ${installed} -> ${latest}."
run_installer "$tag" || { echo "docx-cli bootstrap: update failed — keeping ${installed}." >&2; exit 1; }
ensure_reachable || exit 1
