#!/bin/sh
# docx-cli installer.
#
# Downloads the pre-built binary for your platform and drops it into $PREFIX
# (default: $HOME/.local/bin) as `docx`.
#
# This script is published as a RELEASE ASSET and listed in that release's SHA256SUMS,
# so fetch it from a release (never from a branch) and run it from a file — that way it
# is immutable per release and you can verify it before running it.
#
# Usage:
#   curl -fsSLO https://github.com/kklimuk/docx-cli/releases/latest/download/install.sh
#   sh install.sh
#
#   PREFIX=/usr/local/bin VERSION=v1.2.3 sh install.sh
#
# Environment:
#   PREFIX            where to install (default: $HOME/.local/bin)
#   VERSION           release tag to install, or "latest" (default: latest)
#   REQUIRE_CHECKSUM  1 = refuse to install when no checksum tool is available, instead
#                     of warning and installing unverified. The skill's bootstrap.sh
#                     sets this; a human running the script by hand gets the lenient
#                     default, so a minimal box still has a working install path.
#
# It downloads only release ASSETS — the binary and its SHA256SUMS manifest — and never
# fetches or executes remote code.
#
# POSIX sh, no bashisms. Tested under bash, dash, busybox sh.

set -eu

REPO="kklimuk/docx-cli"
# ${HOME-} (not $HOME) so `set -u` doesn't abort before an error path can explain itself
# in a HOME-less environment (containers, cron). An unusable default surfaces at mkdir.
PREFIX="${PREFIX:-${HOME-}/.local/bin}"
# A trailing slash would break the PATH-membership test at the end (":$PREFIX:" never
# matches a PATH entry written without one).
PREFIX="${PREFIX%/}"
VERSION="${VERSION:-latest}"
REQUIRE_CHECKSUM="${REQUIRE_CHECKSUM:-0}"

npm_hint() { echo "  Install from the npm registry instead:  bun add -g bun-docx" >&2; }
prefix_hint() { echo "  Pick a writable location:  PREFIX=/usr/local/bin sh install.sh" >&2; }

os="$(uname -s)"
arch="$(uname -m)"

# ─── Platform → release asset name. Nonzero when we publish no binary for it. ───
# These names must match the `name:` column of the build matrix in
# .github/workflows/release.yml; tests/cli/installer.test.ts asserts that they do.
detect_target() {
  case "$os" in
    Linux)
      case "$arch" in
        x86_64|amd64) echo "docx-linux-x64" ;;
        aarch64|arm64) echo "docx-linux-arm64" ;;
        *) return 1 ;;
      esac
      ;;
    Darwin)
      case "$arch" in
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
  download() { curl -fsSL "$1" -o "$2"; }
elif command -v wget >/dev/null 2>&1; then
  download() { wget -qO "$2" "$1"; }
else
  echo "Error: need curl or wget to install." >&2
  npm_hint
  exit 1
fi

# ─── Pick a checksum tool ───
# sha256_of() prints ONLY the lowercase hex digest, so callers need not know each tool's
# output shape: the *sum tools put the digest FIRST ("<hash>  <file>"), while `openssl
# dgst` puts it LAST ("SHA256(file)= <hash>" on LibreSSL, "SHA2-256(file)= <hash>" on
# OpenSSL 3.x). openssl is the third fallback because it is a native binary: on macOS,
# /usr/bin/shasum is a PERL script, and Apple has deprecated the bundled scripting
# runtimes, so shasum disappears whenever Perl does.
verify=1
if command -v sha256sum >/dev/null 2>&1; then
  sha256_of() { sha256sum "$1" | awk '{print $1}'; }
elif command -v shasum >/dev/null 2>&1; then
  sha256_of() { shasum -a 256 "$1" | awk '{print $1}'; }
elif command -v openssl >/dev/null 2>&1; then
  sha256_of() { openssl dgst -sha256 "$1" | awk '{print $NF}'; }
else
  verify=0
fi

target="$(detect_target)" || {
  echo "Error: unsupported platform: ${os} ${arch}" >&2
  echo "  Supported: linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64." >&2
  npm_hint
  exit 1
}
binary_name="docx"
case "$target" in *.exe) binary_name="docx.exe" ;; esac

if [ "$verify" = 0 ]; then
  if [ "$REQUIRE_CHECKSUM" = "1" ]; then
    echo "Error: no sha256sum, shasum, or openssl on this system — refusing to install a" >&2
    echo "  binary whose integrity cannot be verified." >&2
    npm_hint
    exit 1
  fi
  echo "Warning: no sha256sum/shasum/openssl found — installing WITHOUT integrity verification." >&2
fi

if [ "$VERSION" = "latest" ]; then
  base="https://github.com/${REPO}/releases/latest/download"
else
  base="https://github.com/${REPO}/releases/download/${VERSION}"
fi

# Get a release asset or die. $3, when given, is an extra line of context.
fetch_asset() {
  download "${base}/$1" "$2" && [ -s "$2" ] || {
    echo "Error: could not download ${base}/$1." >&2
    [ -z "${3-}" ] || echo "  $3" >&2
    exit 1
  }
}

# Stage both downloads INSIDE $PREFIX: the final step is then a same-directory rename
# rather than a cross-device copy of ~100 MB, and an unwritable destination fails here —
# before the transfer — instead of after it.
bin_tmp="${PREFIX}/.docx-download.$$"
sums_tmp="${PREFIX}/.docx-sums.$$"
mkdir -p "$PREFIX" || { echo "Error: could not create ${PREFIX}." >&2; prefix_hint; exit 1; }
# Sweep leftovers from a run killed by a signal the trap below cannot catch (SIGKILL,
# power loss). Each stray file is a ~100 MB binary sitting in a bin directory.
rm -f "$PREFIX"/.docx-download.* "$PREFIX"/.docx-sums.* 2>/dev/null || true
trap 'rm -f "$bin_tmp" "$sums_tmp"' EXIT INT TERM
touch "$bin_tmp" 2>/dev/null || { echo "Error: ${PREFIX} is not writable." >&2; prefix_hint; exit 1; }

# Manifest first: if this release carries no checksum for our asset, we find out before
# spending the binary download on it.
expected=""
if [ "$verify" = 1 ]; then
  fetch_asset SHA256SUMS "$sums_tmp" "Refusing to install an unverified binary."
  expected="$(awk -v f="$target" '$2 == f || $2 == "*"f { print $1; exit }' "$sums_tmp")"
  [ -n "$expected" ] || {
    echo "Error: no checksum for ${target} in SHA256SUMS — refusing to install." >&2
    exit 1
  }
fi

echo "→ Downloading ${target} from ${base}/${target}"
fetch_asset "$target" "$bin_tmp"

if [ -n "$expected" ]; then
  actual="$(sha256_of "$bin_tmp")"
  [ "$actual" = "$expected" ] || {
    echo "Error: SHA-256 mismatch for ${target} — refusing to install." >&2
    echo "  expected: $expected" >&2
    echo "  actual:   $actual" >&2
    exit 1
  }
  echo "✓ Verified SHA-256 (${target})"
fi

chmod +x "$bin_tmp"
mv "$bin_tmp" "${PREFIX}/${binary_name}"
echo "✓ Installed: ${PREFIX}/${binary_name}"

# ─── PATH hint ───
case ":${PATH}:" in
  *":$PREFIX:"*) ;;
  *)
    echo
    echo "  Note: $PREFIX is not on your PATH."
    echo "  Add this to your shell profile (~/.zshrc, ~/.bashrc, etc):"
    echo "    export PATH=\"$PREFIX:\$PATH\""
    ;;
esac

# ─── Confirm it actually runs here (a checksum proves bytes, not compatibility) ───
echo
"${PREFIX}/${binary_name}" --version || true
