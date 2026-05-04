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

# ─── Compose URL ───
if [ "$VERSION" = "latest" ]; then
  url="https://github.com/${REPO}/releases/latest/download/${target}"
else
  url="https://github.com/${REPO}/releases/download/${VERSION}/${target}"
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
