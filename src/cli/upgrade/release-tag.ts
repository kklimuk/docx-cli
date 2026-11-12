/** A release tag, and nothing that can steer the URL somewhere else. `--to` (and the
 *  admin's `DOCX_CLI_VERSION` pin) are interpolated into install.sh's release URL, and
 *  curl NORMALIZES `..` — so an unvalidated value
 *  ("../../../../other/repo/releases/download/v1") redirects both the binary AND the
 *  SHA256SUMS it is checked against to an attacker's repo, making the checksum gate
 *  verify the attacker's own manifest. Agents drive this CLI over untrusted document
 *  content, so the value must be constrained at the CLI ingress, per "the CLI surface
 *  is concerned with parsing user input". bootstrap.sh applies the same pattern. */
const RELEASE_TAG = /^v?\d+\.\d+\.\d+(?:[-.][0-9A-Za-z.-]+)?$/;

/** The canonical `vX.Y.Z` tag for a validated value, or null when it isn't one.
 *  `docx --version` prints "0.23.0" but the release is tagged "v0.23.0"; accept what
 *  the tool itself printed rather than 404ing on the most likely typo. */
export function toReleaseTag(value: string): string | null {
	if (!RELEASE_TAG.test(value)) return null;
	return value.startsWith("v") ? value : `v${value}`;
}
