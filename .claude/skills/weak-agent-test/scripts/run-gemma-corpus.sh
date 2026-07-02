#!/usr/bin/env bash
# Run the weak-agent EXERCISE phase for the whole corpus using the LOCAL Gemma
# harness (E4B) instead of Claude Haiku. Serial by design — the GPU is a single
# resource, so concurrent Gemma runs would contend badly. Each scenario:
#   stage-gemma.sh  ->  bare `bun index.ts --cwd … --prompt-file …`  ->  gemma-parse-ledger.ts
# producing <runDir>/<key>/{<doc>, exercise.json, gemma-run.log}. A failed or
# timed-out run does NOT abort the corpus — it degrades to whatever the ledger holds
# and the loop moves on.
#
# Usage:
#   run-gemma-corpus.sh <scenariosDir> <runDir> <binary> <harness> [ctx] [capSec] [key...]
#     ctx     chat context size (default 32768)
#     capSec  per-scenario watchdog kill (default 1500s = 25m)
#     key...  optional subset of scenario keys (default: all six)
set -uo pipefail

scenariosDir="${1:?scenariosDir}"; runDir="${2:?runDir}"; binary="${3:?binary}"; harness="${4:?harness}"
ctx="${5:-32768}"; capSec="${6:-1500}"
shift $(( $# < 6 ? $# : 6 ))
want=("$@")   # remaining args = scenario key subset (empty = all)

here="$(cd "$(dirname "$0")" && pwd)"
# GEMMA_MODEL_PATH / GEMMA_MMPROJ_PATH override the model (default E4B) — e.g. the 12B control run.
model="${GEMMA_MODEL_PATH:-$harness/models/gemma-4-E4B-it-UD-Q4_K_XL.gguf}"
mmproj="${GEMMA_MMPROJ_PATH:-$harness/models/gemma-4-e4b-mmproj-BF16.gguf}"
test -f "$model" && test -f "$mmproj" || { echo "MISSING model files: $model / $mmproj" >&2; exit 1; }

# key|doc|kind — mirrors the workflow SCENARIOS manifest.
manifest=(
	"mnda|mnda.docx|edit"
	"invoice|invoice.docx|edit"
	"resume|resume.docx|edit"
	"contract-markup|contract.docx|edit"
	"contract-finalize|contract-redlined.docx|edit"
	"eliot-journal|journal.docx|author"
)

wanted() {
	[ "${#want[@]}" -eq 0 ] && return 0
	for k in "${want[@]}"; do [ "$k" = "$1" ] && return 0; done
	return 1
}

mkdir -p "$runDir"
for entry in "${manifest[@]}"; do
	IFS='|' read -r key doc kind <<< "$entry"
	wanted "$key" || continue
	dir="$runDir/$key"
	echo "=================== $key ($kind, $doc) ==================="

	bash "$here/stage-gemma.sh" "$key" "$scenariosDir" "$runDir" "$binary" "$doc" "$kind" \
		|| { echo "[$key] STAGE FAILED — skipping" >&2; continue; }

	start=$(date +%s)
	INKLING_MODEL_PATH="$model" INKLING_MMPROJ_PATH="$mmproj" \
		bun "$harness/src/index.ts" --cwd "$dir" --prompt-file "$dir/exercise-prompt.md" \
		--auto --context-size "$ctx" > "$dir/gemma-run.log" 2>&1 &
	pid=$!
	( sleep "$capSec"; kill -TERM "$pid" 2>/dev/null ) & wd=$!
	wait "$pid"; rc=$?
	kill "$wd" 2>/dev/null; wait "$wd" 2>/dev/null   # stop AND reap the watchdog subshell
	elapsed=$(( $(date +%s) - start ))
	# rc 143 = 128+SIGTERM: the watchdog fired, i.e. the run hit the ${capSec}s ceiling.
	# The parse below still runs against whatever partial ledger exists, but flag it so a
	# force-killed run isn't silently read as a clean completion.
	[ "$rc" -eq 143 ] && echo "[$key] TIMEOUT at ${capSec}s — watchdog killed the run (result is partial)" >&2

	bun "$here/gemma-parse-ledger.ts" "$dir" "$binary" "$doc" >/dev/null 2>&1 \
		|| echo "[$key] PARSE FAILED" >&2
	completed=$(bun -e 'try{console.log(JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")).completed)}catch{console.log("?")}' "$dir/exercise.json" 2>/dev/null)
	echo "[$key] rc=$rc elapsed=${elapsed}s completed=$completed"
done
echo "=================== corpus done ==================="
