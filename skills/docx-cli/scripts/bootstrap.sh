#!/bin/sh
# docx-cli skill bootstrap.
#
# Ensures the `docx` binary is installed and reasonably current, so an agent that
# activated this skill can rely on it. Run it once at the start of a session:
#   sh scripts/bootstrap.sh
#
# This script does NOT install anything itself — it decides WHETHER to install, and
# delegates the how to its sibling `install.sh` (a local, shipped file; nothing is
# fetched and run). Division of labour:
#   install.sh    "put docx here"          — detect platform, download, verify, place
#   bootstrap.sh  "keep docx current"      — everything below
#
# Behavior:
#   - not installed -> resolve latest tag, install it (pinned + checksum-verified)
#   - installed     -> compare against the latest release; update only if BEHIND
#   - offline / can't resolve latest -> keep the working binary and exit 0 (every verb
#     except `render` works offline anyway)
#
# It is STRICTER than a by-hand `install.sh` run: it sets REQUIRE_CHECKSUM=1, so a box
# with no sha256sum/shasum/openssl gets a refusal rather than an unverified binary. An
# agent cannot weigh that trade-off, so it does not get the lenient default.
#
# POSIX sh, no bashisms. The binary is the source of truth — this script only keeps it
# present and fresh; it never edits the skill.

set -eu

REPO="kklimuk/docx-cli"
API_LATEST="https://api.github.com/repos/${REPO}/releases/latest"
# ${HOME-} (not $HOME) so `set -u` doesn't abort a HOME-less run (containers, cron)
# before the up-to-date fast path — which never touches $PREFIX — can exit cleanly.
PREFIX="${PREFIX:-${HOME-}/.local/bin}"
INSTALLER="$(dirname "$0")/install.sh"
case "$(uname -s)" in
  MINGW*|MSYS*|CYGWIN*) binary_name="docx.exe" ;;
  *) binary_name="docx" ;;
esac

npm_hint() { echo "  Install from the npm registry instead:  bun add -g bun-docx" >&2; }

# ─── Pick a downloader (only needed to ask the API for the latest tag) ───
if command -v curl >/dev/null 2>&1; then
  fetch() { curl -fsSL "$1"; }
elif command -v wget >/dev/null 2>&1; then
  fetch() { wget -qO- "$1"; }
else
  echo "docx-cli bootstrap: need curl or wget." >&2
  exit 1
fi

docx_version() { docx --version 2>/dev/null | awk '{print $NF}'; }

# Resolve the latest release tag (e.g. "v0.19.1"), best-effort; empty on failure.
resolve_latest_tag() {
  fetch "$API_LATEST" 2>/dev/null \
    | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1
}

# Hand the actual install to the sibling script, pinned to the resolved tag. REQUIRE_
# CHECKSUM=1 makes it refuse rather than fall back to an unverified install.
install_release() {
  tag="$1"
  if [ ! -f "$INSTALLER" ]; then
    echo "docx-cli bootstrap: $INSTALLER is missing — run this from the skill folder." >&2
    npm_hint
    return 1
  fi
  REQUIRE_CHECKSUM=1 PREFIX="$PREFIX" VERSION="$tag" sh "$INSTALLER"
}

# install.sh drops the binary in $PREFIX and neither script can edit the caller's PATH.
# So after installing, confirm the `docx` an agent will actually invoke IS the one just
# installed, and fail LOUDLY otherwise — rather than reporting a false success the agent
# trips over on its first `docx` call (and re-downloading every session).
#
# RUNNING `docx --version` is the check, not `command -v`: one probe catches all three
# ways a "successful" install still isn't usable — not on PATH at all, on PATH but
# shadowed by an older docx elsewhere (bun global, /usr/local/bin), or resolvable but
# not executable on this system (e.g. a glibc build on musl).
ensure_reachable() {
  want="${1#v}"

  if ! command -v docx >/dev/null 2>&1; then
    echo "docx-cli bootstrap: installed to $PREFIX, but it is NOT on your PATH." >&2
    echo "  Add it:  export PATH=\"$PREFIX:\$PATH\"   (then re-run)" >&2
    echo "  Or invoke the binary directly: $PREFIX/$binary_name" >&2
    return 1
  fi

  got="$(docx_version)"
  if [ -z "$got" ]; then
    echo "docx-cli bootstrap: installed $PREFIX/$binary_name, but \`docx --version\` does not run." >&2
    echo "  The release binary may not be compatible with this system." >&2
    npm_hint
    return 1
  fi

  if [ "$got" != "$want" ]; then
    echo "docx-cli bootstrap: installed ${want} to $PREFIX/$binary_name, but \`docx\` on PATH" >&2
    echo "  still resolves to ${got} ($(command -v docx)) — an older install shadows it." >&2
    echo "  Put $PREFIX first:  export PATH=\"$PREFIX:\$PATH\"   (then re-run)" >&2
    echo "  Or remove the shadowing copy." >&2
    return 1
  fi

  return 0
}

# ─── Not installed: resolve tag, install pinned + verified, finish ───
if ! command -v docx >/dev/null 2>&1; then
  echo "docx not found on PATH."
  tag="$(resolve_latest_tag)"
  if [ -z "$tag" ]; then
    echo "docx-cli bootstrap: could not resolve the latest release (offline or rate-limited) — cannot install safely." >&2
    npm_hint
    exit 1
  fi
  install_release "$tag" || exit 1
  ensure_reachable "$tag" || exit 1
  exit 0
fi

installed="$(docx_version)"
echo "docx-cli present: ${installed:-unknown}"

# Couldn't read a version? Don't enter a reinstall loop — leave the present binary.
if [ -z "$installed" ]; then
  echo "Could not read the installed version — leaving the present binary in place."
  exit 0
fi

tag="$(resolve_latest_tag)"
latest="${tag#v}"
if [ -z "$latest" ]; then
  echo "Could not determine the latest release (offline or rate-limited) — keeping the installed binary."
  exit 0
fi

if [ "$installed" = "$latest" ]; then
  echo "✓ Up to date (${installed})."
  exit 0
fi

# Versions differ. Only UPDATE when installed is OLDER than latest — never downgrade a
# locally-built/pre-release binary ahead of the published release. A `sort` that lacks
# -V prints nothing, which falls through to updating (the documented default).
if [ "$(printf '%s\n%s\n' "$installed" "$latest" | sort -V 2>/dev/null | head -n 1)" = "$latest" ]; then
  echo "Installed (${installed}) is newer than the latest release (${latest}) — keeping it."
  exit 0
fi

echo "A newer release is available: ${installed} -> ${latest}."
install_release "$tag" || { echo "docx-cli bootstrap: update failed — keeping ${installed}." >&2; exit 1; }
ensure_reachable "$tag" || exit 1
