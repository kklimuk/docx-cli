#!/bin/sh
# docx-cli installer.
#
# Downloads the latest pre-built binary for your platform and drops it into
# $PREFIX (default: $HOME/.local/bin) as `docx`.
#
# Usage:
#   curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh | sh
#   PREFIX=/usr/local sh -c "$(curl -fsSL https://raw.githubusercontent.com/kklimuk/docx-cli/main/install.sh)"
#
# POSIX sh, no bashisms. Tested under bash, dash, busybox sh.

set -eu

REPO="kklimuk/docx-cli"
PREFIX="${PREFIX:-$HOME/.local/bin}"
VERSION="${VERSION:-latest}"

# ─── Detect platform ───
detect_target() {
  os="$(uname -s)"
  arch="$(uname -m)"

  case "$os" in
    Linux)
      case "$arch" in
        x86_64|amd64) echo "docx-linux-x64" ;;
        aarch64|arm64) echo "docx-linux-arm64" ;;
        *) echo "unsupported-arch" ;;
      esac
      ;;
    Darwin)
      case "$arch" in
        x86_64) echo "docx-darwin-x64" ;;
        arm64) echo "docx-darwin-arm64" ;;
        *) echo "unsupported-arch" ;;
      esac
      ;;
    MINGW*|MSYS*|CYGWIN*) echo "docx-windows-x64.exe" ;;
    *) echo "unsupported-os" ;;
  esac
}

target="$(detect_target)"
case "$target" in
  unsupported-*)
    echo "Unsupported platform: $(uname -s) $(uname -m)" >&2
    echo "Supported: linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64." >&2
    exit 1
    ;;
esac

# ─── Pick downloader ───
if command -v curl >/dev/null 2>&1; then
  download() { curl -fsSL "$1" -o "$2"; }
elif command -v wget >/dev/null 2>&1; then
  download() { wget -qO "$2" "$1"; }
else
  echo "Need curl or wget to install." >&2
  exit 1
fi

# ─── Compose URLs (binary + the release's checksum manifest) ───
if [ "$VERSION" = "latest" ]; then
  url="https://github.com/${REPO}/releases/latest/download/${target}"
  sums_url="https://github.com/${REPO}/releases/latest/download/SHA256SUMS"
else
  url="https://github.com/${REPO}/releases/download/${VERSION}/${target}"
  sums_url="https://github.com/${REPO}/releases/download/${VERSION}/SHA256SUMS"
fi

# ─── Download and install ───
mkdir -p "$PREFIX"
binary_name="docx"
case "$target" in
  *.exe) binary_name="docx.exe" ;;
esac
target_path="$PREFIX/$binary_name"
tmp_path="$(mktemp 2>/dev/null || mktemp -t docx-cli)"
trap 'rm -f "$tmp_path"' EXIT INT TERM

echo "→ Downloading $target from $url"
download "$url" "$tmp_path"

# ─── Verify the download against the release's published SHA256SUMS ───
if command -v sha256sum >/dev/null 2>&1; then
  sha_cmd="sha256sum"
elif command -v shasum >/dev/null 2>&1; then
  sha_cmd="shasum -a 256"
else
  sha_cmd=""
fi
if [ -n "$sha_cmd" ]; then
  sums_path="$(mktemp 2>/dev/null || mktemp -t docx-cli-sums)"
  if ! download "$sums_url" "$sums_path" || [ ! -s "$sums_path" ]; then
    rm -f "$sums_path"
    echo "Error: could not download SHA256SUMS from $sums_url — refusing to install an unverified binary." >&2
    exit 1
  fi
  expected="$(awk -v f="$target" '$2 == f || $2 == "*"f { print $1; exit }' "$sums_path")"
  rm -f "$sums_path"
  if [ -z "$expected" ]; then
    echo "Error: no checksum for $target in SHA256SUMS — refusing to install." >&2
    exit 1
  fi
  actual="$($sha_cmd "$tmp_path" | awk '{print $1}')"
  if [ "$actual" != "$expected" ]; then
    echo "Error: SHA-256 mismatch for $target — refusing to install." >&2
    echo "  expected: $expected" >&2
    echo "  actual:   $actual" >&2
    exit 1
  fi
  echo "✓ Verified SHA-256 ($target)"
else
  echo "Warning: no sha256sum/shasum found — installing WITHOUT integrity verification." >&2
fi

chmod +x "$tmp_path"
mv "$tmp_path" "$target_path"

echo "✓ Installed: $target_path"

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

# ─── Verify ───
if [ -x "$target_path" ]; then
  echo
  "$target_path" --version || true
fi
