#!/bin/sh
# docx-cli skill bootstrap.
#
# Ensures the `docx` binary is installed and reasonably current, so an agent that
# activated this skill can rely on it. Run it once at the start of a session:
#   sh scripts/bootstrap.sh
#
# Supply-chain posture: no remote CODE, ever. It asks the GitHub releases API for the
# latest tag, downloads that tag's prebuilt binary plus its SHA256SUMS manifest, and
# refuses to install unless the digest matches (no checksum tool -> no install).
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
# ${HOME-} (not $HOME) so `set -u` doesn't abort a HOME-less run (containers, cron)
# before the up-to-date fast path — which never touches $PREFIX — can exit cleanly.
PREFIX="${PREFIX:-${HOME-}/.local/bin}"
binary_name="docx"

npm_hint() { echo "  Install from the npm registry instead:  bun add -g bun-docx" >&2; }
prefix_hint() { echo "  Or pick a writable location:  PREFIX=/usr/local/bin sh scripts/bootstrap.sh" >&2; }

# ─── Platform → release asset name. Nonzero when we publish no binary for it. ───
detect_target() {
  case "$(uname -s)" in
    Linux)
      case "$(uname -m)" in
        x86_64|amd64) echo "docx-linux-x64" ;;
        aarch64|arm64) echo "docx-linux-arm64" ;;
        *) return 1 ;;
      esac
      ;;
    Darwin)
      case "$(uname -m)" in
        x86_64) echo "docx-darwin-x64" ;;
        arm64) echo "docx-darwin-arm64" ;;
        *) return 1 ;;
      esac
      ;;
    MINGW*|MSYS*|CYGWIN*) echo "docx-windows-x64.exe" ;;
    *) return 1 ;;
  esac
}

# ─── Pick a downloader ───
if command -v curl >/dev/null 2>&1; then
  fetch() { curl -fsSL "$1"; }
  download() { curl -fsSL "$1" -o "$2"; }
elif command -v wget >/dev/null 2>&1; then
  fetch() { wget -qO- "$1"; }
  download() { wget -qO "$2" "$1"; }
else
  echo "docx-cli bootstrap: need curl or wget." >&2
  exit 1
fi

# ─── Pick a checksum tool ───
# sha256_of() prints ONLY the lowercase hex digest, so callers need not know each tool's
# output shape: the *sum tools put the digest FIRST ("<hash>  <file>"), while `openssl
# dgst` puts it LAST ("SHA256(file)= <hash>" on LibreSSL, "SHA2-256(file)= <hash>" on
# OpenSSL 3.x). openssl is the third fallback because it is a native binary: on macOS,
# /usr/bin/shasum is a PERL script, and Apple has deprecated the bundled scripting
# runtimes, so shasum disappears whenever Perl does. Left UNDEFINED when no tool is
# present — install_release tests for it and refuses rather than install unverified.
if command -v sha256sum >/dev/null 2>&1; then
  sha256_of() { sha256sum "$1" | awk '{print $1}'; }
elif command -v shasum >/dev/null 2>&1; then
  sha256_of() { shasum -a 256 "$1" | awk '{print $1}'; }
elif command -v openssl >/dev/null 2>&1; then
  sha256_of() { openssl dgst -sha256 "$1" | awk '{print $NF}'; }
fi

docx_version() { docx --version 2>/dev/null | awk '{print $NF}'; }

# Resolve the latest release tag (e.g. "v0.19.1"), best-effort; empty on failure.
resolve_latest_tag() {
  fetch "$API_LATEST" 2>/dev/null \
    | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1
}

# Install the given release tag. Downloads only release ASSETS — never code — and
# verifies the binary's digest before it lands. Checks run cheapest-first, so a doomed
# install costs a few hundred bytes of manifest rather than a ~100 MB binary transfer.
#
# NOTE: this runs with `set -e` SUSPENDED — every caller invokes it as
# `install_release "$tag" || …`, which disables errexit for the whole function body.
# So every fallible step below MUST be checked explicitly; an unchecked failure would
# fall through to the "✓ Installed" line and report a false success.
install_release() {
  tag="$1"
  base="https://github.com/${REPO}/releases/download/${tag}"

  target="$(detect_target)" || {
    echo "docx-cli bootstrap: unsupported platform: $(uname -s) $(uname -m)" >&2
    echo "  Supported: linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64." >&2
    npm_hint
    return 1
  }
  case "$target" in *.exe) binary_name="docx.exe" ;; esac

  command -v sha256_of >/dev/null 2>&1 || {
    echo "docx-cli bootstrap: no sha256sum, shasum, or openssl on this system — refusing" >&2
    echo "  to install a binary whose integrity cannot be verified." >&2
    npm_hint
    return 1
  }

  # Stage both downloads INSIDE $PREFIX: the final step is then a same-directory rename
  # rather than a cross-device copy of ~100 MB, and an unwritable destination fails here
  # — before the transfer — instead of after it.
  mkdir -p "$PREFIX" 2>/dev/null || {
    echo "docx-cli bootstrap: could not create ${PREFIX}." >&2
    prefix_hint
    return 1
  }
  bin_tmp="${PREFIX}/.docx-download.$$"
  sums_tmp="${PREFIX}/.docx-sums.$$"
  trap 'rm -f "$bin_tmp" "$sums_tmp"' EXIT INT TERM
  touch "$bin_tmp" 2>/dev/null || {
    echo "docx-cli bootstrap: ${PREFIX} is not writable." >&2
    prefix_hint
    return 1
  }

  echo "→ Installing docx-cli ${tag} (${target}, pinned + checksum-verified) ..."

  # Manifest first: if this release carries no checksum for our asset, we find out
  # before spending the binary download on it.
  if ! download "${base}/SHA256SUMS" "$sums_tmp" || [ ! -s "$sums_tmp" ]; then
    echo "docx-cli bootstrap: could not download ${base}/SHA256SUMS — refusing to install" >&2
    echo "  an unverified binary." >&2
    return 1
  fi
  expected="$(awk -v f="$target" '$2 == f || $2 == "*"f { print $1; exit }' "$sums_tmp")"
  if [ -z "$expected" ]; then
    echo "docx-cli bootstrap: no checksum for ${target} in SHA256SUMS — refusing to install." >&2
    return 1
  fi

  if ! download "${base}/${target}" "$bin_tmp" || [ ! -s "$bin_tmp" ]; then
    echo "docx-cli bootstrap: could not download ${base}/${target} (offline or rate-limited)." >&2
    return 1
  fi
  actual="$(sha256_of "$bin_tmp")"
  if [ "$actual" != "$expected" ]; then
    echo "docx-cli bootstrap: SHA-256 mismatch for ${target} — refusing to install." >&2
    echo "  expected: $expected" >&2
    echo "  actual:   $actual" >&2
    return 1
  fi
  echo "✓ Verified SHA-256 (${target})"

  chmod +x "$bin_tmp" || {
    echo "docx-cli bootstrap: could not make the downloaded binary executable." >&2
    return 1
  }
  mv "$bin_tmp" "${PREFIX}/${binary_name}" || {
    echo "docx-cli bootstrap: could not install into ${PREFIX}/${binary_name}." >&2
    prefix_hint
    return 1
  }
  echo "✓ Installed: ${PREFIX}/${binary_name}"
  return 0
}

# The binary lands in $PREFIX and this script cannot edit the caller's PATH. So after
# installing, confirm the `docx` an agent will actually invoke IS the one we just
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
