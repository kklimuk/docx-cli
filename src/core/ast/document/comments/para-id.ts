/** Mint a fresh `w14:paraId` value. Word writes 8-char uppercase hex; we
 * match. Per MS-DOCX, a paraId MUST be greater than 0x00000000 and less
 * than 0x80000000 — Word silently re-keys out-of-range ids on open, which
 * orphans every sidecar entry (`w15:commentEx` resolved flags, reply
 * `w15:paraIdParent` links) keyed to the old value. Used by
 * `CommentsView.ensureParaId`, the sidecar reconcile, and the comments lens
 * for replies + audit comments.
 *
 * `DOCX_CLI_PARA_ID_SEED` (8-char hex) switches to sequential minting from
 * that base — the byte-determinism seam for fixture rebuilds, mirroring
 * `DOCX_CLI_NOW`. The counter is per-process: a builder that invokes the CLI
 * multiple times must set a DISTINCT seed per invocation or ids collide. */
export function generateParaId(): string {
	const seed = Bun.env.DOCX_CLI_PARA_ID_SEED;
	if (seed) {
		const minted = (Number.parseInt(seed, 16) + seededParaIds++) % 0x80000000;
		return (minted || 1).toString(16).padStart(8, "0").toUpperCase();
	}
	const bytes = new Uint8Array(4);
	crypto.getRandomValues(bytes);
	bytes[0] = (bytes[0] ?? 0) & 0x7f;
	const hex = [...bytes]
		.map((byte) => byte.toString(16).padStart(2, "0"))
		.join("");
	if (hex === "00000000") return "00000001";
	return hex.toUpperCase();
}

let seededParaIds = 0;

/** Mint a paraId not in `taken` (uppercase hex). The random minter's
 * collision odds are negligible, but the seeded one counts up from a fixed
 * base — a file built under one seed and reconciled under another would
 * collide without the check. */
export function mintUniqueParaId(taken: Set<string>): string {
	for (;;) {
		const candidate = generateParaId();
		if (!taken.has(candidate)) return candidate;
	}
}

/** Mint a `w16cid:durableId` not in `taken` (uppercase hex). MS-DOCX bounds
 * it to (0, 0x7FFFFFFF) — the paraId range minus its top value — so reuse
 * the paraId minter (and its `DOCX_CLI_PARA_ID_SEED` determinism seam). */
export function mintDurableId(taken: Set<string>): string {
	for (;;) {
		const candidate = generateParaId();
		if (isValidDurableId(candidate) && !taken.has(candidate)) return candidate;
	}
}

/** Whether `value` is a durableId Word accepts: hex, in (0, 0x7FFFFFFF). */
export function isValidDurableId(value: string | undefined): boolean {
	if (!value || !/^[0-9A-Fa-f]{1,8}$/.test(value)) return false;
	const numeric = Number.parseInt(value, 16);
	return numeric > 0 && numeric < 0x7fffffff;
}
