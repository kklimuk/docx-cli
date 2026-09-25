#!/bin/sh
# Run after all binary modifications; upload/hash only the resulting signed bytes.
set -eu

if [ "$#" -ne 1 ]; then
  echo "Usage: sh scripts/sign-macos-binary.sh BINARY" >&2
  exit 2
fi

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
# Explicit entitlements retain Bun's JavaScriptCore JIT permissions, even when
# the compiler's original signature is absent or invalid. No signing key needed.
codesign --force --sign - --entitlements "$script_dir/macos-entitlements.plist" "$1"
codesign --verify --strict --verbose=2 "$1"
