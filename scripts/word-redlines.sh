#!/usr/bin/env bash
# Accept or reject ALL tracked changes in a .docx using Microsoft Word for Mac
# (AppleScript automation) — the authoritative oracle for cross-checking
# docx-cli's own track-changes accept/reject. Unlike LibreOffice, Word actually
# applies tracked table row/column/merge revisions.
#
#   scripts/word-redlines.sh accept|reject INPUT.docx OUTPUT.docx
#
# Word for Mac is sandboxed and prompts ("Grant File Access") for arbitrary
# paths, so the file is staged inside Word's own container Documents dir, which
# it can open without a prompt. macOS-only; requires Microsoft Word and that
# automation permission has been granted to the controlling terminal (first run
# triggers a one-time TCC prompt).
set -euo pipefail

mode="${1:?usage: word-redlines.sh accept|reject INPUT OUTPUT}"
input="${2:?missing INPUT}"
output="${3:?missing OUTPUT}"

case "$mode" in
	accept) cmd="accept all revisions" ;;
	reject) cmd="reject all revisions" ;;
	*) echo "mode must be accept or reject" >&2; exit 2 ;;
esac

stage="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
[ -d "$stage" ] || { echo "Word container not found ($stage)" >&2; exit 1; }
staged="$stage/.word-redlines-$$.docx"
trap 'rm -f "$staged"' EXIT
cp "$input" "$staged"

osascript \
	-e 'tell application "Microsoft Word"' \
	-e "open \"$staged\"" \
	-e 'set d to active document' \
	-e "$cmd d" \
	-e 'save d' \
	-e 'close d saving no' \
	-e 'end tell' >/dev/null 2>&1

cp "$staged" "$output"
