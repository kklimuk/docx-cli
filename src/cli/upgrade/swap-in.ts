import { rename, rm } from "node:fs/promises";

/** Windows locks a running `.exe` against overwrite and delete, but not against
 *  rename. So a verified binary staged beside it is swapped in with two renames:
 *  the running one steps aside to `<path>.old` (it can't delete itself, so the
 *  next upgrade sweeps it) and the staged one takes its name. If that second
 *  rename fails (antivirus often holds a freshly written `.exe` open for a
 *  moment), the running binary is put back before the error propagates, so the
 *  user is never left with only `.old`. */
export async function swapIn(
	staged: string,
	binaryPath: string,
): Promise<void> {
	const aside = `${binaryPath}.old`;
	await rm(aside, { force: true });
	await rename(binaryPath, aside);
	try {
		await rename(staged, binaryPath);
	} catch (error) {
		await rename(aside, binaryPath);
		throw error;
	}
}
