#!/usr/bin/env bash
# Prep ONE scenario folder for a Gemma exercise run. No model involved — this is the
# "pre-staged scenario folder" step: copy the pristine scenario in, strip the judge's
# answer key, pre-authorize the docx binary, and render the static prompt template.
#
# After this, the RUN is a bare harness call (no wrapper). Pass --cwd so the harness
# resolves .inkling (permissions + ledger) and the bash tool's working dir INTO the
# scenario folder:
#   INKLING_MODEL_PATH=<model.gguf> INKLING_MMPROJ_PATH=<mmproj.gguf> \
#     bun <harness>/src/index.ts --cwd <runDir>/<key> \
#       --prompt-file <runDir>/<key>/exercise-prompt.md --auto --context-size 32768
# and the PARSE is: gemma-parse-ledger.ts <runDir>/<key> <binary> <doc>
#
# Usage: stage-gemma.sh <key> <scenariosDir> <runDir> <binary> [doc] [kind]
#   [kind] = edit (default) | author — swaps the working-document line in the prompt.
set -euo pipefail

key="${1:?key}"; scenariosDir="${2:?scenariosDir}"; runDir="${3:?runDir}"; binary="${4:?binary}"
doc="${5:-$key.docx}"
kind="${6:-edit}"
here="$(cd "$(dirname "$0")" && pwd)"

if [ "$kind" = "author" ]; then
	workline="You are authoring from scratch. Create your output at EXACTLY this path:"
else
	workline="Your working document (already a private copy — edit it IN PLACE, do NOT pass -o/--output):"
fi
src="$scenariosDir/$key"
dst="$runDir/$key"

mkdir -p "$dst/.inkling"
cp -R "$src/." "$dst/"
rm -f "$dst/criteria.md"   # withhold the judge-only answer key from the agent's workspace

# Pre-authorize the docx binary so Gemma drives it unimpeded (clean tool-economy
# signal, not the --auto permission judge); deny web/agent so it can't wander.
# NOTE: this JSON mirrors the Inkling harness's `permissions` rule format
# (allow/deny/ask arrays of `Tool(matcher)` rule strings; see the harness repo's
# src/permissions). If that grammar ever changes, the harness silently IGNORES
# rules it can't parse — update this block to match rather than debugging a
# "why is Gemma getting prompted" mystery.
cat > "$dst/.inkling/settings.json" <<JSON
{
  "permissions": {
    "allow": ["Bash($binary:*)", "Bash(ls:*)", "Bash(cat:*)", "Bash(pwd)", "Bash(mkdir:*)"],
    "deny": ["Bash(rm:*)", "Bash(sudo:*)", "WebSearch", "WebFetch", "Agent", "AgentPool"],
    "ask": []
  }
}
JSON

# Render the static prompt template with this scenario's binary path + working doc.
# {{WORKLINE}} first (kind-dependent, no {{DOC}} inside it), then binary + doc.
sed -e "s#{{WORKLINE}}#$workline#g" -e "s#{{BINARY}}#$binary#g" -e "s#{{DOC}}#$doc#g" \
  "$here/gemma-exercise-prompt.md" > "$dst/exercise-prompt.md"

echo "staged $key -> $dst (doc=$doc, kind=$kind)"
