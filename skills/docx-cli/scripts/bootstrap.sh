#!/bin/sh
# docx-cli skill bootstrap.
#
# Ensures the `docx` binary is installed and reasonably current, so an agent that
# activated this skill can rely on it. Run it once at the start of a session:
#   bash scripts/bootstrap.sh
#
# Behavior:
#   - not installed  -> install the latest release via the canonical installer
#   - installed       -> compare against the latest GitHub release tag; if behind,
#                        self-update (re-running the installer, which fetches latest)
#   - no network / can't determine latest -> leave the working binary in place and
#                        exit 0 (every verb except `render` works offline anyway)
#
# POSIX sh, no bashisms. The binary is the source of truth — this script only keeps
# it present and fresh; it never edits the skill.

set -eu

REPO="kklimuk/docx-cli"
INSTALL_URL="https://raw.githubusercontent.com/${REPO}/main/install.sh"

# ─── Pick a downloader (shared by the version check and the installer fetch) ───
if command -v curl >/dev/null 2>&1; then
  fetch() { curl -fsSL "$1"; }
elif command -v wget >/dev/null 2>&1; then
  fetch() { wget -qO- "$1"; }
else
  echo "docx-cli bootstrap: need curl or wget." >&2
  exit 1
fi

# Download install.sh to a temp file FIRST, then run it — so a failed/empty download
# is caught here instead of being silently swallowed (`fetch | sh` would let `sh` read
# empty stdin and exit 0, masking the failure). Returns the installer's exit status.
run_installer() {
  echo "→ Installing docx-cli (latest) ..."
  installer="$(mktemp 2>/dev/null || mktemp -t docx-cli-install)"
  if ! fetch "$INSTALL_URL" > "$installer" || [ ! -s "$installer" ]; then
    rm -f "$installer"
    echo "docx-cli bootstrap: could not download the installer from $INSTALL_URL (offline or rate-limited)." >&2
    return 1
  fi
  if sh "$installer"; then rc=0; else rc=$?; fi
  rm -f "$installer"
  return "$rc"
}

# install.sh drops the binary in ${PREFIX:-$HOME/.local/bin} and only PRINTS a PATH hint
# — it can't edit the caller's PATH. So after installing, confirm `docx` is actually
# resolvable and fail LOUDLY if not, rather than reporting a false success the agent
# trips over on its first `docx` call (and re-downloading from scratch every session).
ensure_reachable() {
  command -v docx >/dev/null 2>&1 && return 0
  bindir="${PREFIX:-$HOME/.local/bin}"
  echo "docx-cli bootstrap: installed to $bindir, but it is NOT on your PATH." >&2
  echo "  Add it:  export PATH=\"$bindir:\$PATH\"   (then re-run)" >&2
  echo "  Or invoke the binary directly: $bindir/docx" >&2
  return 1
}

# ─── Not installed: install, confirm reachable, finish ───
if ! command -v docx >/dev/null 2>&1; then
  echo "docx not found on PATH."
  run_installer || exit 1
  ensure_reachable || exit 1
  exit 0
fi

installed="$(docx --version 2>/dev/null | awk '{print $NF}')"
echo "docx-cli present: ${installed:-unknown}"

# Couldn't read a version (broken or changed --version output)? Do NOT enter a reinstall
# loop — the binary is present and resolvable; leave it and exit clean.
if [ -z "$installed" ]; then
  echo "Could not read the installed version — leaving the present binary in place."
  exit 0
fi

# ─── Determine the latest released version (best-effort) ───
latest_json="$(fetch "https://api.github.com/repos/${REPO}/releases/latest" 2>/dev/null || true)"
latest="$(printf '%s' "$latest_json" | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"v\{0,1\}\([^"]*\)".*/\1/p' | head -n 1)"

if [ -z "$latest" ]; then
  echo "Could not determine the latest release (offline or rate-limited) — keeping the installed binary."
  exit 0
fi

if [ "$installed" = "$latest" ]; then
  echo "✓ Up to date (${installed})."
  exit 0
fi

# Versions differ. Only UPDATE when installed is OLDER than latest — never downgrade a
# locally-built/pre-release binary that's ahead of the published release. Use version
# sort when the platform's `sort` supports -V; otherwise fall back to updating.
if printf '%s\n' "0.0" "0.1" | sort -V >/dev/null 2>&1; then
  oldest="$(printf '%s\n%s\n' "$installed" "$latest" | sort -V | head -n 1)"
  if [ "$oldest" != "$installed" ]; then
    echo "Installed (${installed}) is newer than the latest release (${latest}) — keeping it."
    exit 0
  fi
fi

echo "A newer release is available: ${installed} -> ${latest}."
run_installer || { echo "docx-cli bootstrap: update failed — keeping ${installed}." >&2; exit 1; }
ensure_reachable || exit 1
