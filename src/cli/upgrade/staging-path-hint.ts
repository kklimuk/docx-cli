/** On Windows, install.sh runs with PREFIX set to the staging folder, so its
 *  closing "Note: <PREFIX> is not on your PATH … export PATH=…" fires on every
 *  upgrade and tells the user to put a folder on PATH that is deleted a moment
 *  later. install.sh is the one shared installer, so the hint is dropped here,
 *  from the forwarded stdout, rather than patched there. Output is held only
 *  until its line ends; every other line passes through unchanged. */
export function withoutStagingPathHint(
	write: (text: string) => Promise<void>,
	stagingPrefix: string,
): { write: (text: string) => Promise<void>; flush: () => Promise<void> } {
	let pending = "";
	const isHint = (line: string): boolean =>
		line.includes(`${stagingPrefix} is not on your PATH`) ||
		line.includes(`export PATH="${stagingPrefix}:`) ||
		line.includes("Add this to your shell profile");
	return {
		async write(text) {
			pending += text;
			const lines = pending.split("\n");
			pending = lines.pop() ?? "";
			const kept = lines.filter((line) => !isHint(line));
			if (kept.length > 0) await write(`${kept.join("\n")}\n`);
		},
		async flush() {
			const rest = pending;
			pending = "";
			if (rest && !isHint(rest)) await write(rest);
		},
	};
}
