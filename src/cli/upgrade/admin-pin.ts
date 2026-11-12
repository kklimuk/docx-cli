import { toReleaseTag } from "./release-tag";

/** Why an administrator's environment forbids this upgrade, or null when it doesn't.
 *  `DOCX_CLI_NO_UPDATE=1` and `DOCX_CLI_VERSION` are the fleet controls SECURITY.md
 *  documents and bootstrap.sh honors; the skill tells agents they mean "an admin has
 *  pinned the version". If `upgrade` ignored them, the pin would hold only until an
 *  agent (or a human) reached for the other door. So, mirroring bootstrap:
 *  - `DOCX_CLI_NO_UPDATE=1` refuses every upgrade;
 *  - `DOCX_CLI_VERSION` refuses any target other than the pinned tag — `--to` that
 *    tag is how the admin's pin gets installed, so that one stays allowed;
 *  - a malformed `DOCX_CLI_VERSION` is a usage error, as in bootstrap (exit 2) — and,
 *    as there, it is validated FIRST, so `DOCX_CLI_NO_UPDATE=1` doesn't mask it.
 *  Checked before the install-method and `--dry-run` branches: a dry run previews what
 *  would happen, and what would happen is this refusal. */
export function adminPinRefusal(
	env: Record<string, string | undefined>,
	target: string,
	installed: string,
): AdminPinRefusal | null {
	const raw = env.DOCX_CLI_VERSION;
	const pin = raw ? toReleaseTag(raw) : null;
	if (raw && pin === null) {
		return {
			code: "USAGE",
			error: `DOCX_CLI_VERSION=${JSON.stringify(raw)} is not a release tag (e.g. v0.27.0)`,
			hint: "Fix or unset DOCX_CLI_VERSION in the environment; nothing was changed.",
		};
	}
	if (env.DOCX_CLI_NO_UPDATE === "1") {
		return {
			code: "UPGRADE_FAILED",
			error: `upgrades are disabled by DOCX_CLI_NO_UPDATE=1 — docx ${installed} was left in place`,
			hint: "An administrator has pinned this install. Keep using it, and ask them to update it.",
		};
	}
	if (pin === null || target === pin) return null;
	return {
		code: "UPGRADE_FAILED",
		error: `DOCX_CLI_VERSION=${raw} pins docx to ${pin}, so upgrading to ${target} is refused — docx ${installed} was left in place`,
		hint:
			`v${installed}` === pin
				? `docx ${installed} already matches the pin; nothing to upgrade.`
				: `Install the pinned release with:  docx upgrade --to ${pin}`,
	};
}

type AdminPinRefusal = {
	code: "USAGE" | "UPGRADE_FAILED";
	error: string;
	hint: string;
};
