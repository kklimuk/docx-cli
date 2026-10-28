import { rename, rm } from "node:fs/promises";

/** Windows locks a running `.exe` against overwrite and delete, but not against
 *  rename. So a verified binary staged beside it is swapped in with two renames:
 *  the running one steps aside to `<path>.old` (it can't delete itself, so the
 *  next upgrade sweeps it) and the staged one takes its name. */
export async function swapIn(
	staged: string,
	binaryPath: string,
): Promise<void> {
	const aside = `${binaryPath}.old`;
	await rm(aside, { force: true });
	await rename(binaryPath, aside);
	await rename(staged, binaryPath);
}
