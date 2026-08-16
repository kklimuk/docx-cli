#!/bin/sh
# docx-cli skill bootstrap.
#
# Ensures the `docx` binary is installed and reasonably current, so an agent that
# activated this skill can rely on it. Run it once at the start of a session:
#   sh scripts/bootstrap.sh
#
# This script does NOT install anything itself — it decides WHETHER to install and
# delegates the how to its sibling `install.sh` (a local, shipped file; nothing is
# fetched and run). Division of labour:
#   install.sh    "put docx here"      — detect platform, download, verify, place
#   bootstrap.sh  "keep docx current"  — everything below
#
# Behavior:
#   - not installed -> install the latest release (pinned + checksum-verified)
#   - installed     -> compare against the latest release; update only if BEHIND
#   - offline / can't resolve latest -> keep the working binary and exit 0 (every verb
#     except `render` works offline anyway)
#
# It passes REQUIRE_CHECKSUM=1, so a box with no sha256sum/shasum/openssl gets a refusal
# rather than an unverified binary — an agent cannot weigh that trade-off, so it does not
# get install.sh's lenient default.
#
# POSIX sh, no bashisms. The binary is the source of truth — this script only keeps it
# present and fresh; it never edits the skill.

set -eu

REPO="kklimuk/docx-cli"
API_LATEST="https://api.github.com/repos/${REPO}/releases/latest"
RELEASES_LATEST="https://github.com/${REPO}/releases/latest"
# ${HOME-} (not $HOME) so `set -u` doesn't abort a HOME-less run (containers, cron)
# before the up-to-date fast path — which never touches $PREFIX — can exit cleanly.
PREFIX="${PREFIX:-${HOME-}/.local/bin}"
PREFIX="${PREFIX%/}"
# `dirname` is an external binary in sh/dash; this is the same answer with no fork.
case "$0" in
  */*) INSTALLER="${0%/*}/install.sh" ;;
  *) INSTALLER="./install.sh" ;;
esac

npm_hint() { echo "  Install from the npm registry instead:  bun add -g bun-docx" >&2; }

# ─── Pick a downloader, and with it how to resolve the latest tag ───
if command -v curl >/dev/null 2>&1; then
  # Ask the releases page where it redirects instead of reading the JSON API: same
  # answer, one round trip, no body to parse — and crucially OFF the API's 60/hr/IP
  # budget, which a NAT'd office or CI egress IP shares across every agent session.
  resolve_latest_tag() {
    url="$(curl -fsS -o /dev/null -w '%{redirect_url}' "$RELEASES_LATEST" 2>/dev/null || true)"
    echo "${url##*/}"
  }
elif command -v wget >/dev/null 2>&1; then
  resolve_latest_tag() {
    wget -qO- "$API_LATEST" 2>/dev/null \
      | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1
  }
else
  echo "docx-cli bootstrap: need curl or wget." >&2
  npm_hint
  exit 1
fi

# Version of whatever binary $1 names, or empty if it won't run.
version_of() {
  out="$("$1" --version 2>/dev/null)" || return 0
  echo "${out##* }"
}

# Path of the binary install.sh writes. PROBED, not re-derived: install.sh owns the
# platform→`.exe` rule, and a second copy of it here would silently disagree the day
# another Windows asset appears — surfacing to the user as a bogus "shadowed" report.
staged_path() {
  if [ -e "$PREFIX/docx.exe" ]; then echo "$PREFIX/docx.exe"; else echo "$PREFIX/docx"; fi
}

# Hand the install to the sibling script, pinned to $1 ("latest" is also accepted).
ensure_release_staged() {
  if [ ! -f "$INSTALLER" ]; then
    echo "docx-cli bootstrap: $INSTALLER is missing — run this from the skill folder." >&2
    npm_hint
    return 1
  fi
  # Already the target version? Then what failed last time was REACHABILITY, and
  # re-pulling ~100 MB every session cannot fix a PATH problem.
  staged="$(staged_path)"
  if [ "$(version_of "$staged")" = "${1#v}" ]; then
    echo "docx-cli ${1#v} already staged at ${staged} — skipping the download."
    return 0
  fi
  REQUIRE_CHECKSUM=1 PREFIX="$PREFIX" VERSION="$1" sh "$INSTALLER"
}

# install.sh drops the binary in $PREFIX and neither script can edit the caller's PATH.
# So after installing, confirm the `docx` an agent will actually invoke IS the one just
# installed, and fail LOUDLY otherwise — rather than reporting a false success the agent
# trips over on its first `docx` call (and re-downloading every session).
#
# Resolve the PATH entry FIRST and split on identity, not on version: a broken FOREIGN
# `docx` earlier on PATH is a PATH problem, and blaming it on our binary's compatibility
# would send the user to npm instead of to their PATH.
ensure_reachable() {
  want="${1#v}"
  staged="$(staged_path)"
  resolved="$(command -v docx 2>/dev/null || true)"

  if [ -z "$resolved" ]; then
    echo "docx-cli bootstrap: installed to $PREFIX, but it is NOT on your PATH." >&2
    echo "  Add it:  export PATH=\"$PREFIX:\$PATH\"   (then re-run)" >&2
    echo "  Or invoke the binary directly: $staged" >&2
    return 1
  fi

  if [ "$resolved" != "$staged" ]; then
    echo "docx-cli bootstrap: installed ${want} to ${staged}, but \`docx\` on PATH resolves" >&2
    echo "  to ${resolved} — a different install shadows it." >&2
    echo "  Put $PREFIX first:  export PATH=\"$PREFIX:\$PATH\"   (then re-run)" >&2
    echo "  Or remove the shadowing copy." >&2
    return 1
  fi

  got="$(version_of docx)"
  if [ -z "$got" ]; then
    echo "docx-cli bootstrap: installed ${staged}, but \`docx --version\` does not run." >&2
    echo "  The release binary may not be compatible with this system." >&2
    npm_hint
    return 1
  fi

  if [ "$got" != "$want" ]; then
    echo "docx-cli bootstrap: installed release ${want}, but ${staged} reports ${got}." >&2
    echo "  The release assets for ${want} do not match the tag." >&2
    npm_hint
    return 1
  fi

  return 0
}

# ─── Not installed: install the latest release, then confirm it is reachable ───
if ! command -v docx >/dev/null 2>&1; then
  echo "docx not found on PATH."
  tag="$(resolve_latest_tag)"
  if [ -z "$tag" ]; then
    # Couldn't learn the tag — but the ASSET host resolves `latest` server-side and is
    # not rate-limited, so a first install must not be blocked by a tag lookup. Install
    # "latest" and read the version back off the binary we just placed.
    echo "Could not resolve the latest tag (offline or rate-limited) — installing 'latest'."
    ensure_release_staged latest || exit 1
    tag="$(version_of "$(staged_path)")"
    if [ -z "$tag" ]; then
      echo "docx-cli bootstrap: install produced no runnable binary." >&2
      npm_hint
      exit 1
    fi
  else
    ensure_release_staged "$tag" || exit 1
  fi
  ensure_reachable "$tag" || exit 1
  exit 0
fi

installed="$(version_of docx)"
echo "docx-cli present: ${installed:-unknown}"

# Couldn't read a version? Don't enter a reinstall loop — leave the present binary.
if [ -z "$installed" ]; then
  echo "Could not read the installed version — leaving the present binary in place."
  exit 0
fi

# Keep BOTH forms: the tag ("v0.22.0") addresses the release URL, the stripped version
# ("0.22.0") compares against `docx --version`. Passing the stripped form to the
# installer builds releases/download/0.22.0/… and 404s.
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
ensure_release_staged "$tag" || { echo "docx-cli bootstrap: update failed — keeping ${installed}." >&2; exit 1; }
ensure_reachable "$tag" || exit 1
