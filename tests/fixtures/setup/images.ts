import { mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { $ } from "bun";

/**
 * Build tests/fixtures/images.docx — a doc that starts with no images and gets
 * three inserted through the public `insert --image` surface (a parity check
 * that every image source kind the resolver supports is reachable from the CLI):
 *   1. PNG via local file path (native pixel size)
 *   2. JPEG via a data: URI (native pixel size)
 *   3. PNG via local file path, scaled to an explicit 1.5" width (aspect kept)
 *
 * The raw images live under tests/fixtures/assets/ as committed bytes (real
 * encodings so LibreOffice renders them on round-trip); this script consumes
 * them rather than regenerating, so it runs anywhere without an image encoder.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/images.docx");
const cliEntry = resolve(root, "src/index.ts");
const assets = resolve(root, "tests/fixtures/assets");

async function cli(...args: string[]): Promise<void> {
	await $`bun ${cliEntry} ${args}`.quiet();
}

mkdirSync(dirname(out), { recursive: true });

const pngPath = resolve(assets, "sample.png");
const jpgBytes = await Bun.file(resolve(assets, "sample.jpg")).bytes();
const jpgDataUri = `data:image/jpeg;base64,${Buffer.from(jpgBytes).toString("base64")}`;

await cli("create", out, "--force", "--text", "Images fixture");

await cli("insert", out, "--after", "p0", "--text", "PNG from a file path");
await cli(
	"insert",
	out,
	"--after",
	"p1",
	"--image",
	pngPath,
	"--alt",
	"Sample PNG",
);

await cli("insert", out, "--after", "p2", "--text", "JPEG from a data: URI");
await cli(
	"insert",
	out,
	"--after",
	"p3",
	"--image",
	jpgDataUri,
	"--alt",
	"Sample JPEG",
);

await cli("insert", out, "--after", "p4", "--text", "PNG scaled to 1.5 inches");
await cli(
	"insert",
	out,
	"--after",
	"p5",
	"--image",
	pngPath,
	"--alt",
	"Scaled PNG",
	"--width",
	"1.5",
);

const bytes = (await Bun.file(out).bytes()).length;
console.log(`Wrote ${out} (${bytes} bytes)`);
