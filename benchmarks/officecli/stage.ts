#!/usr/bin/env bun
import { chmod, mkdir } from "node:fs/promises";
import { resolve } from "node:path";
import { writeStdout } from "../../src/cli/respond";

// An explicit release and a fresh destination make accidental upgrades/resumes fail.
const [tag, destination] = Bun.argv.slice(2);
if (!tag || !/^v\d+\.\d+\.\d+$/.test(tag) || !destination) {
	throw new Error("Usage: bun benchmarks/officecli/stage.ts vX.Y.Z EMPTY_DEST");
}
const root = resolve(destination);
await mkdir(root, { recursive: false });
const repo = "https://api.github.com/repos/iOfficeAI/OfficeCLI";
const release = await getJson(`${repo}/releases/tags/${tag}`);
const platforms: Record<string, string> = { darwin: "mac", linux: "linux", win32: "win" };
const platform = platforms[process.platform];
if (!platform || !["arm64", "x64"].includes(process.arch)) {
	throw new Error(`Unsupported platform: ${process.platform}/${process.arch}`);
}
const assetName = `officecli-${platform}-${process.arch}${process.platform === "win32" ? ".exe" : ""}`;
const asset = release.assets.find((entry: { name: string }) => entry.name === assetName);
if (!asset || !/^sha256:[a-f0-9]{64}$/.test(asset.digest ?? "")) {
	throw new Error(`No SHA-256-verified asset available for ${assetName}`);
}
const binary = resolve(root, process.platform === "win32" ? "officecli.exe" : "officecli");
const bytes = await download(asset.browser_download_url);
const digest = new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
if (`sha256:${digest}` !== asset.digest) throw new Error("OfficeCLI checksum mismatch");
await Bun.write(binary, bytes);
await chmod(binary, 0o755);

// Fetch the entire skill tree from the release commit, including specialized skills.
const commit = await getJson(`${repo}/commits/${tag}`);
if (!/^[a-f0-9]{40}$/.test(commit.sha ?? "")) throw new Error("Missing release commit");
const source = resolve(root, "source");
const clone = Bun.spawn(["git", "clone", "--depth", "1", "--branch", tag,
	"https://github.com/iOfficeAI/OfficeCLI.git", source], { stdout: "pipe", stderr: "pipe" });
const [cloneError, cloneStatus] = await Promise.all([new Response(clone.stderr).text(), clone.exited]);
if (cloneStatus !== 0) throw new Error(`Source clone failed: ${cloneError}`);
const head = Bun.spawn(["git", "-C", source, "rev-parse", "HEAD"], { stdout: "pipe" });
const [headText, headStatus] = await Promise.all([new Response(head.stdout).text(), head.exited]);
if (headStatus !== 0 || headText.trim() !== commit.sha) throw new Error("Release source commit mismatch");
if (!(await Bun.file(resolve(source, "SKILL.md")).exists())) throw new Error("Missing SKILL.md");
const versionProcess = Bun.spawn([binary, "--version"], {
	env: { ...Bun.env, OFFICECLI_SKIP_UPDATE: "1" }, stdout: "pipe", stderr: "pipe",
});
const [version, versionError, status] = await Promise.all([
	new Response(versionProcess.stdout).text(), new Response(versionProcess.stderr).text(), versionProcess.exited,
]);
if (status !== 0 || !version.includes(tag.slice(1))) {
	throw new Error(`Version preflight failed: ${version} ${versionError}`);
}
const manifest = {
	tag, commit: commit.sha, asset: assetName, sha256: digest,
	url: asset.browser_download_url, version: version.trim(), binary, source,
	skillSha256: new Bun.CryptoHasher("sha256").update(await Bun.file(resolve(source, "SKILL.md")).arrayBuffer()).digest("hex"),
	stagedAt: new Date().toISOString(), platform: process.platform, arch: process.arch,
};
await Bun.write(resolve(root, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
await writeStdout(`${JSON.stringify(manifest, null, 2)}\n`);

async function getJson(url: string) {
	const response = await fetch(url, { headers: { Accept: "application/vnd.github+json" } });
	if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
	return response.json();
}

async function download(url: string) {
	const response = await fetch(url);
	if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
	return response.arrayBuffer();
}
