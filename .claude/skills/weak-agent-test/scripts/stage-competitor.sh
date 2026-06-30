#!/usr/bin/env bash
#
# Stage the COMPETITOR for the A/B bake-off: Anthropic's official, bundled "docx"
# Agent Skill (python + raw-OOXML), provisioned with its FULL intended toolset.
#
# Fairness is the whole ballgame for the bake-off — an under-equipped competitor
# would void the comparison. So this script (a) fetches the real skill from
# anthropics/skills, and (b) installs every dependency the skill's SKILL.md relies
# on (python-docx, lxml, the Node `docx` library, pandoc), then VERIFIES each one is
# actually usable. It exits non-zero if any REQUIRED piece is missing, so a human can
# confirm a green competitor setup before spending a bake-off run.
#
# Usage:
#   stage-competitor.sh <SKILL_DEST> [RUN_DIR]
#     SKILL_DEST  where to place the skill (will contain SKILL.md + scripts/). Pass
#                 this same path to the workflow as args.competitorDir.
#     RUN_DIR     optional: the bake-off run dir. The Node `docx` library is installed
#                 here so every scenario subfolder (<RUN_DIR>/<key>/) resolves it via
#                 Node's upward module resolution.
#
# Idempotent: re-running skips work that's already done.

set -uo pipefail

SKILL_DEST="${1:-}"
RUN_DIR="${2:-}"
if [ -z "$SKILL_DEST" ]; then
	echo "usage: stage-competitor.sh <SKILL_DEST> [RUN_DIR]" >&2
	exit 2
fi

REPO_URL="https://github.com/anthropics/skills.git"
SKILL_SUBPATH="skills/docx"
FAILED=0

# The skill's required artifacts — used for BOTH the "already staged?" skip guard and
# the post-copy verify, so a partial copy can't pass the guard yet fail the verify (and
# then never re-fetch to self-heal).
REQUIRED="SKILL.md scripts/office/unpack.py scripts/office/pack.py scripts/comment.py"
have_all_required() {
	for required in $REQUIRED; do
		[ -f "$SKILL_DEST/$required" ] || return 1
	done
	return 0
}

note() { printf '[ok]   %s\n' "$1"; }
warn() { printf '[warn] %s\n' "$1"; }
fail() { printf '[FAIL] %s\n' "$1"; FAILED=1; }

echo "=== Staging competitor: Anthropic docx skill ==="
echo "    skill dest: $SKILL_DEST"
[ -n "$RUN_DIR" ] && echo "    run dir:    $RUN_DIR"
echo

# ---------------------------------------------------------------------------
# 1. Fetch the real skill (SKILL.md + scripts/) into SKILL_DEST.
# ---------------------------------------------------------------------------
if have_all_required; then
	note "skill already present at $SKILL_DEST (skipping clone)"
else
	command -v git >/dev/null 2>&1 || { fail "git not found — cannot fetch the skill"; exit 1; }
	TMP_CLONE="$(mktemp -d)"
	trap 'rm -rf "$TMP_CLONE"' EXIT
	echo "    cloning $REPO_URL (shallow) ..."
	# Try a sparse partial clone first (fast); fall back to a full clone if ANY part of
	# the sparse path fails — the clone, the sparse-checkout, OR an empty resulting tree
	# (old git, cone-mode quirks, or a server refusing the blob filter). The original
	# only fell back when the FIRST clone failed, so a sparse-checkout failure hard-exited.
	if ! { git clone --depth 1 --filter=blob:none --sparse "$REPO_URL" "$TMP_CLONE" >/dev/null 2>&1 \
		&& ( cd "$TMP_CLONE" && git sparse-checkout set "$SKILL_SUBPATH" >/dev/null 2>&1 ) \
		&& [ -d "$TMP_CLONE/$SKILL_SUBPATH" ]; }; then
		rm -rf "$TMP_CLONE"
		git clone --depth 1 "$REPO_URL" "$TMP_CLONE" >/dev/null 2>&1 \
			|| { fail "git clone of $REPO_URL failed (network? auth?)"; exit 1; }
	fi
	if [ ! -d "$TMP_CLONE/$SKILL_SUBPATH" ]; then
		fail "$SKILL_SUBPATH not found in the clone — repo layout may have changed"
		exit 1
	fi
	mkdir -p "$SKILL_DEST"
	cp -R "$TMP_CLONE/$SKILL_SUBPATH/." "$SKILL_DEST/"
	note "copied skill into $SKILL_DEST"
fi

# Verify the skill's required artifacts landed.
for required in $REQUIRED; do
	if [ -f "$SKILL_DEST/$required" ]; then
		note "present: $required"
	else
		fail "missing: $SKILL_DEST/$required"
	fi
done

# ---------------------------------------------------------------------------
# 2. Python deps (python-docx, lxml, defusedxml) for EVERY python the agents might
#    invoke. macOS commonly has MULTIPLE Homebrew pythons (e.g. `python`=3.13 and
#    `python3`=3.14) with SEPARATE site-packages, and the skill's docs say
#    `python scripts/...` — so the deps must be importable by BOTH `python` and
#    `python3`, or an agent hits ModuleNotFoundError and a PEP-668-blocked
#    `pip install` it can't recover from. Install with --break-system-packages (the
#    reliable option on externally-managed Homebrew python) and VERIFY each.
# ---------------------------------------------------------------------------
PY_FOUND=0
for PY in python python3; do
	command -v "$PY" >/dev/null 2>&1 || continue
	PY_FOUND=$((PY_FOUND + 1))
	if "$PY" -c "import docx, lxml, defusedxml.minidom, PIL, numpy" >/dev/null 2>&1; then
		note "$PY: python deps already importable"
		continue
	fi
	echo "    installing python-docx + lxml + defusedxml + pillow + numpy for $PY ..."
	"$PY" -m pip install --quiet --break-system-packages python-docx lxml defusedxml pillow numpy >/dev/null 2>&1 \
		|| "$PY" -m pip install --quiet --user --break-system-packages python-docx lxml defusedxml pillow numpy >/dev/null 2>&1 \
		|| "$PY" -m pip install --quiet --user python-docx lxml defusedxml pillow numpy >/dev/null 2>&1
	if "$PY" -c "import docx, lxml, defusedxml.minidom, PIL, numpy" >/dev/null 2>&1; then
		note "$PY: python deps installed and importable"
	else
		fail "$PY: python-docx/lxml/defusedxml NOT importable — install them manually, then re-run"
	fi
done
[ "$PY_FOUND" -eq 0 ] && fail "no python/python3 on PATH — the skill's scripts cannot run"

# ---------------------------------------------------------------------------
# 3. Node `docx` library (the skill creates documents with it). Install where the
#    weak agents can resolve it: RUN_DIR (so every scenario subfolder finds it) and,
#    as a fallback, SKILL_DEST.
# ---------------------------------------------------------------------------
if command -v npm >/dev/null 2>&1; then
	install_docx_lib() {
		local target="$1"
		[ -d "$target/node_modules/docx" ] && return 0
		mkdir -p "$target"
		( cd "$target" && { [ -f package.json ] || npm init -y >/dev/null 2>&1; } )
		( cd "$target" && npm install --silent --no-audit --no-fund docx >/dev/null 2>&1 )
	}
	for target in "$RUN_DIR" "$SKILL_DEST"; do
		[ -z "$target" ] && continue
		if install_docx_lib "$target" && [ -d "$target/node_modules/docx" ]; then
			note "Node \`docx\` available at $target/node_modules"
		else
			warn "could not install Node \`docx\` at $target (create-from-scratch flow may be limited)"
		fi
	done
else
	warn "npm not found — the skill's Node \`docx\` create flow will be unavailable"
fi

# ---------------------------------------------------------------------------
# 4. pandoc (the skill's read/text-extraction path). Strongly recommended.
# ---------------------------------------------------------------------------
if command -v pandoc >/dev/null 2>&1; then
	note "pandoc present ($(pandoc --version 2>/dev/null | head -1))"
elif command -v brew >/dev/null 2>&1; then
	echo "    installing pandoc via brew ..."
	if brew install pandoc >/dev/null 2>&1 && command -v pandoc >/dev/null 2>&1; then
		note "pandoc installed via brew"
	else
		warn "pandoc install failed — the skill's read path will be degraded (install manually for full fairness)"
	fi
else
	warn "pandoc NOT found and no brew to install it — the skill's read path will be degraded"
fi

# ---------------------------------------------------------------------------
# 5. Image tooling for SVG figure scenarios. The skill documents LibreOffice for
#    PDF, but weak agents reach for ImageMagick `convert` + librsvg to rasterize
#    SVG figures (e.g. a frontispiece or an SVG logo). Provide them so an image
#    task isn't blocked by a missing system tool (the docx-cli arm handles SVG
#    natively, so withholding this would handicap only the competitor).
# ---------------------------------------------------------------------------
if command -v convert >/dev/null 2>&1 && command -v rsvg-convert >/dev/null 2>&1; then
	note "image tooling present (convert + rsvg-convert)"
elif command -v brew >/dev/null 2>&1; then
	echo "    installing imagemagick + librsvg via brew (may take a few min) ..."
	brew install imagemagick librsvg >/dev/null 2>&1
	if command -v convert >/dev/null 2>&1; then
		note "imagemagick + librsvg installed"
	else
		warn "image tooling install failed — SVG figure scenarios may be limited for the competitor"
	fi
else
	warn "imagemagick/librsvg NOT found and no brew — SVG figure scenarios may be limited for the competitor"
fi

echo
if [ "$FAILED" -ne 0 ]; then
	echo "=== COMPETITOR STAGING INCOMPLETE — fix the [FAIL]s above before running the competitor arm. ==="
	echo "    (A handicapped competitor would void the bake-off.)"
	exit 1
fi
echo "=== Competitor staged. Pass this to the workflow: ==="
echo "    arm: \"anthropic-docx-skill\", competitorDir: \"$SKILL_DEST\""
[ -n "$RUN_DIR" ] && echo "    (Node \`docx\` installed under $RUN_DIR for scenario-folder resolution.)"
