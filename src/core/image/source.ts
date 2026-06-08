// Resolves an image reference (file path, data: URI, or http(s) URL) to raw
// bytes plus the metadata the emitter needs: the file extension and MIME type
// for the package part / content-type registration, and the pixel dimensions
// (parsed from the file header) used to size the drawing's <wp:extent>.

import { lookup } from "node:dns/promises";
import convert from "heic-convert";
import {
	CANONICAL_IMAGE_EXTENSION,
	EXTENSION_BY_IMAGE_MIME,
	IMAGE_MIME_BY_EXTENSION,
} from "./formats";
import { sanitizeSvg } from "./svg-sanitize";

export async function loadImageSource(src: string): Promise<ImageSource> {
	if (src.startsWith("data:")) return loadDataURI(src);
	if (src.startsWith("http://") || src.startsWith("https://")) {
		return loadHttp(src);
	}
	return loadFile(src);
}

export type ImageSource = {
	bytes: Uint8Array;
	extension: string;
	mimeType: string;
	pixelWidth?: number;
	pixelHeight?: number;
};

/** Thrown for any source-resolution failure (unreadable path, non-image
 * content-type, fetch timeout, oversize download). The CLI maps it to a
 * single error code. */
export class ImageSourceError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "ImageSourceError";
	}
}

async function loadDataURI(uri: string): Promise<ImageSource> {
	// data:[<mediatype>][;base64],<data>
	const comma = uri.indexOf(",");
	if (comma === -1)
		throw new ImageSourceError("Malformed data: URI (no comma)");
	const header = uri.slice("data:".length, comma);
	const payload = uri.slice(comma + 1);
	const isBase64 = header.endsWith(";base64");
	const mimeType =
		(isBase64 ? header.slice(0, -";base64".length) : header) || "";
	const extension = EXTENSION_BY_MIME[mimeType.toLowerCase()];
	if (!extension) {
		throw new ImageSourceError(
			`Unsupported or missing image MIME type in data: URI: "${mimeType}"`,
		);
	}
	const bytes = isBase64
		? new Uint8Array(Buffer.from(payload, "base64"))
		: new Uint8Array(Buffer.from(decodeURIComponent(payload), "utf-8"));
	return finalizeSource(bytes, extension);
}

async function loadHttp(url: string): Promise<ImageSource> {
	const controller = new AbortController();
	const timeout = setTimeout(
		() => controller.abort(),
		DEFAULT_FETCH_TIMEOUT_MS,
	);
	try {
		const { response, finalUrl } = await fetchWithSafeRedirects(
			url,
			controller.signal,
		);
		if (!response.ok) {
			throw new ImageSourceError(
				`Failed to fetch image: HTTP ${response.status} ${response.statusText}`,
			);
		}

		const contentType = (response.headers.get("content-type") ?? "")
			.split(";")[0]
			?.trim()
			.toLowerCase();
		const extension =
			(contentType ? EXTENSION_BY_MIME[contentType] : undefined) ??
			extensionFromPath(new URL(finalUrl).pathname);
		if (!extension) {
			throw new ImageSourceError(
				`Could not determine image type from response (content-type "${contentType}") or URL`,
			);
		}

		// Reject an oversize download from the declared length before reading any
		// body bytes — Content-Length can be absent or lie, so the streaming
		// per-chunk tally below is the authoritative cap.
		const declaredLength = Number(response.headers.get("content-length"));
		if (Number.isFinite(declaredLength) && declaredLength > MAX_FETCH_BYTES) {
			throw new ImageSourceError(
				`Image exceeds ${MAX_FETCH_BYTES} byte limit (Content-Length ${declaredLength})`,
			);
		}

		const bytes = await readBodyWithCap(response, MAX_FETCH_BYTES);
		return finalizeSource(bytes, extension);
	} catch (fetchError) {
		if (fetchError instanceof ImageSourceError) throw fetchError;
		const reason =
			fetchError instanceof Error ? fetchError.message : String(fetchError);
		throw new ImageSourceError(`Failed to fetch image: ${reason}`);
	} finally {
		clearTimeout(timeout);
	}
}

const MAX_REDIRECT_HOPS = 5;

/** Walk a redirect chain manually so each hop's target is re-validated against
 * the private-IP blocklist — `fetch`'s default redirect handling would let a
 * benign-looking host 302 into `http://169.254.169.254/...` unchallenged. */
async function fetchWithSafeRedirects(
	initialUrl: string,
	signal: AbortSignal,
): Promise<{ response: Response; finalUrl: string }> {
	let current = initialUrl;
	for (let hop = 0; hop <= MAX_REDIRECT_HOPS; hop++) {
		const parsed = new URL(current);
		if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
			throw new ImageSourceError(
				`Refusing non-http(s) URL: ${parsed.protocol}//`,
			);
		}
		await ensurePublicHost(parsed.hostname);
		const response = await fetch(current, { redirect: "manual", signal });
		if (response.status < 300 || response.status >= 400) {
			return { response, finalUrl: current };
		}
		const location = response.headers.get("location");
		if (!location) return { response, finalUrl: current };
		current = new URL(location, current).href;
	}
	throw new ImageSourceError(
		`Too many redirects following ${initialUrl} (>${MAX_REDIRECT_HOPS} hops)`,
	);
}

/** Reject hostnames that resolve to private, loopback, link-local, or reserved
 * ranges (incl. cloud metadata at 169.254.169.254). We resolve to ALL addresses
 * and reject if any is non-public — defense in depth against DNS records that
 * mix public and private IPs (still TOCTOU-vulnerable to DNS rebinding at
 * connect time, which we can't address without bypassing fetch). */
async function ensurePublicHost(hostname: string): Promise<void> {
	let records: Array<{ address: string; family: number }>;
	try {
		records = await lookup(hostname, { all: true });
	} catch (error) {
		const reason = error instanceof Error ? error.message : String(error);
		throw new ImageSourceError(
			`Could not resolve host "${hostname}": ${reason}`,
		);
	}
	for (const { address, family } of records) {
		if (isBlockedAddress(address, family)) {
			throw new ImageSourceError(
				`Refused to fetch from non-public address ${address} (host "${hostname}")`,
			);
		}
	}
}

function isBlockedAddress(address: string, family: number): boolean {
	if (family === 4) return isBlockedIPv4(address);
	return isBlockedIPv6(address);
}

function isBlockedIPv4(addr: string): boolean {
	const parts = addr.split(".").map(Number);
	if (parts.length !== 4 || parts.some((part) => !Number.isFinite(part))) {
		return true;
	}
	const [a, b] = parts as [number, number, number, number];
	if (a === 0) return true; // "this network"
	if (a === 10) return true; // RFC1918
	if (a === 127) return true; // loopback
	if (a === 169 && b === 254) return true; // link-local incl. cloud metadata
	if (a === 172 && b >= 16 && b <= 31) return true; // RFC1918
	if (a === 192 && b === 168) return true; // RFC1918
	if (a === 100 && b >= 64 && b <= 127) return true; // CGN
	if (a >= 224) return true; // multicast (224/4) + reserved (240/4)
	return false;
}

function isBlockedIPv6(addr: string): boolean {
	const lower = addr.toLowerCase();
	if (lower === "::1" || lower === "::") return true;
	if (lower.startsWith("fe80:")) return true; // link-local
	if (lower.startsWith("fc") || lower.startsWith("fd")) return true; // ULA (fc00::/7)
	if (lower.startsWith("::ffff:")) {
		const v4 = lower.slice("::ffff:".length);
		if (v4.includes(".")) return isBlockedIPv4(v4);
	}
	return false;
}

/** Read the body in chunks, throwing as soon as the running total exceeds the
 * cap — bounds peak memory regardless of what the server actually streams. */
async function readBodyWithCap(
	response: Response,
	maxBytes: number,
): Promise<Uint8Array> {
	if (!response.body) {
		throw new ImageSourceError("Response has no body");
	}
	const reader = response.body.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	while (true) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.byteLength;
		if (total > maxBytes) {
			await reader.cancel();
			throw new ImageSourceError(
				`Fetched image exceeds ${maxBytes} byte limit (streamed ${total}+ bytes)`,
			);
		}
		chunks.push(value);
	}
	const out = new Uint8Array(total);
	let offset = 0;
	for (const chunk of chunks) {
		out.set(chunk, offset);
		offset += chunk.byteLength;
	}
	return out;
}

async function loadFile(path: string): Promise<ImageSource> {
	const cleaned = path.startsWith("file://")
		? path.slice("file://".length)
		: path;
	const file = Bun.file(cleaned);
	if (!(await file.exists())) {
		throw new ImageSourceError(`Image file not found: ${cleaned}`);
	}
	const bytes = new Uint8Array(await file.arrayBuffer());
	const extension = extensionFromPath(cleaned);
	if (!extension) {
		throw new ImageSourceError(
			`Cannot determine image type from path: ${cleaned}`,
		);
	}
	return finalizeSource(bytes, extension);
}

/** Normalize the extension, transcode HEIC→JPEG if needed, and read dimensions
 * — the shared tail every loader funnels through. */
async function finalizeSource(
	bytes: Uint8Array,
	extension: string,
): Promise<ImageSource> {
	const lower = extension.toLowerCase();
	let normalized = CANONICAL_EXTENSION[lower] ?? lower;
	let finalBytes = bytes;

	// Word can't render HEIC/HEIF, so transcode to JPEG before embedding. Detect
	// by header too (not just the label) so a `.jpg` that's really HEIC, or an
	// untyped fetch, is still caught.
	if (HEIC_EXTENSIONS.has(normalized) || isHeic(bytes)) {
		finalBytes = await transcodeHEICToJPEG(bytes);
		normalized = "jpeg";
	}

	// SVG is the one format here that's executable in downstream renderers, so
	// scrub it before embedding — fast-xml-parser already rejects XXE, this
	// strips scripts, event handlers, foreign-HTML islands, and SSRF-on-render
	// URLs at the element/attribute level (see svg-sanitize.ts).
	if (normalized === "svg") {
		try {
			finalBytes = sanitizeSvg(finalBytes);
		} catch (error) {
			const reason = error instanceof Error ? error.message : String(error);
			throw new ImageSourceError(`Could not sanitize SVG: ${reason}`);
		}
	}

	const mimeType = MIME_BY_EXTENSION[normalized];
	if (!mimeType) {
		throw new ImageSourceError(`Unsupported image extension: .${normalized}`);
	}
	const dimensions = readPixelDimensions(finalBytes, normalized);
	return {
		bytes: finalBytes,
		extension: normalized,
		mimeType,
		pixelWidth: dimensions?.width,
		pixelHeight: dimensions?.height,
	};
}

function isHeic(bytes: Uint8Array): boolean {
	if (bytes.length < 12) return false;
	const boxType = new TextDecoder().decode(bytes.slice(4, 8));
	if (boxType !== "ftyp") return false;
	const majorBrand = new TextDecoder().decode(bytes.slice(8, 12));
	return HEIF_BRANDS.has(majorBrand);
}

async function transcodeHEICToJPEG(bytes: Uint8Array): Promise<Uint8Array> {
	try {
		// heic-decode reads the buffer as a typed array (slice + char spread), so
		// pass the Uint8Array directly; @types says ArrayBufferLike but an actual
		// ArrayBuffer throws inside the lib.
		const jpeg = await convert({
			buffer: bytes as unknown as ArrayBufferLike,
			format: "JPEG",
			quality: 0.92,
		});
		return new Uint8Array(jpeg);
	} catch (error) {
		const reason = error instanceof Error ? error.message : String(error);
		throw new ImageSourceError(`Failed to transcode HEIC to JPEG: ${reason}`);
	}
}

function extensionFromPath(path: string): string | undefined {
	const lastDot = path.lastIndexOf(".");
	const lastSlash = path.lastIndexOf("/");
	if (lastDot <= lastSlash || lastDot === path.length - 1) return undefined;
	const ext = path.slice(lastDot + 1).toLowerCase();
	return MIME_BY_EXTENSION[ext] ? ext : undefined;
}

function readPixelDimensions(
	bytes: Uint8Array,
	extension: string,
): { width: number; height: number } | undefined {
	if (extension === "png") return readPNGDimensions(bytes);
	if (extension === "jpeg" || extension === "jpg") {
		return readJPEGDimensions(bytes);
	}
	if (extension === "gif") return readGIFDimensions(bytes);
	// webp/bmp/tiff/svg: dimensions not parsed — caller falls back to overrides.
	return undefined;
}

function readPNGDimensions(
	bytes: Uint8Array,
): { width: number; height: number } | undefined {
	// 8-byte signature, 4-byte length, "IHDR", then width (4) + height (4) BE.
	if (bytes.length < 24) return undefined;
	if (bytes[0] !== 0x89 || bytes[1] !== 0x50) return undefined; // \x89 P
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	return { width: view.getUint32(16), height: view.getUint32(20) };
}

function readGIFDimensions(
	bytes: Uint8Array,
): { width: number; height: number } | undefined {
	if (bytes.length < 10) return undefined;
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	// Logical screen width/height at offsets 6 and 8, little-endian.
	return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
}

function readJPEGDimensions(
	bytes: Uint8Array,
): { width: number; height: number } | undefined {
	if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) {
		return undefined;
	}
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let offset = 2;
	while (offset + 9 < bytes.length) {
		if (view.getUint8(offset) !== 0xff) {
			offset++;
			continue;
		}
		const marker = view.getUint8(offset + 1);
		// SOF markers carry the frame dimensions; skip DHT(C4)/JPG(C8)/DAC(CC).
		if (
			marker >= 0xc0 &&
			marker <= 0xcf &&
			marker !== 0xc4 &&
			marker !== 0xc8 &&
			marker !== 0xcc
		) {
			// [marker:2][length:2][precision:1][height:2][width:2]
			const height = view.getUint16(offset + 5);
			const width = view.getUint16(offset + 7);
			return { width, height };
		}
		// Standalone markers (no length payload) shouldn't appear here; segments
		// carry a 2-byte length we use to jump to the next marker.
		const segmentLength = view.getUint16(offset + 2);
		if (segmentLength < 2) return undefined;
		offset += 2 + segmentLength;
	}
	return undefined;
}

// --- Lookup tables & tuning constants (leaf data the functions above read) ---

const DEFAULT_FETCH_TIMEOUT_MS = 10_000;
const MAX_FETCH_BYTES = 25 * 1024 * 1024;

// Accepted *input* = the shared embeddable formats (image-formats.ts) plus the
// `jpg`/`tif` aliases and HEIC/HEIF, which finalizeSource normalizes/transcodes
// before embedding. Everything we extract or replace is therefore insertable.
const MIME_BY_EXTENSION: Record<string, string> = {
	...IMAGE_MIME_BY_EXTENSION,
	heic: "image/heic",
	heif: "image/heif",
};

const EXTENSION_BY_MIME: Record<string, string> = {
	...EXTENSION_BY_IMAGE_MIME,
	"image/heic": "heic",
	"image/heif": "heif",
};

const CANONICAL_EXTENSION = CANONICAL_IMAGE_EXTENSION;

const HEIC_EXTENSIONS = new Set(["heic", "heif"]);

// HEIF-family `ftyp` brands (the major brand at bytes 8–12 of the ISO-BMFF box).
const HEIF_BRANDS = new Set([
	"heic",
	"heix",
	"heim",
	"heis",
	"hevc",
	"hevx",
	"mif1",
	"msf1",
	"heif",
]);

// --- Display sizing (the other public entry, used by the insert CLI) ---

/** A 96-dpi pixel is 9525 EMU; an inch is 914400 EMU. The drawing layer speaks
 * EMU, so both pixel-derived and inch-override sizes convert through these. */
const EMU_PER_PIXEL = 9525;
const EMU_PER_INCH = 914400;

/** Pick a final EMU width/height for `<wp:extent>` from pixel dimensions and
 * optional inch overrides. When only one dimension is supplied (override or
 * pixels), the other is derived from the pixel aspect ratio. Returns null when
 * no dimension can be determined — the caller surfaces a usage error asking
 * for `--width`/`--height`. */
export function computeExtentEmu(
	source: { pixelWidth?: number; pixelHeight?: number },
	overrides: { widthInches?: number; heightInches?: number },
	/** Clamp the result so its width never exceeds the page's content width
	 *  (page width − margins), preserving aspect ratio. Applied ONLY when no
	 *  explicit `widthInches` is given — an explicit --width is always honored. */
	maxWidthEmu?: number,
): { widthEmu: number; heightEmu: number } | null {
	const { pixelWidth, pixelHeight } = source;
	const aspect =
		pixelWidth && pixelHeight ? pixelHeight / pixelWidth : undefined;

	const widthOverride =
		overrides.widthInches !== undefined
			? Math.round(overrides.widthInches * EMU_PER_INCH)
			: undefined;
	const heightOverride =
		overrides.heightInches !== undefined
			? Math.round(overrides.heightInches * EMU_PER_INCH)
			: undefined;

	let widthEmu =
		widthOverride ?? (pixelWidth ? pixelWidth * EMU_PER_PIXEL : undefined);
	let heightEmu =
		heightOverride ?? (pixelHeight ? pixelHeight * EMU_PER_PIXEL : undefined);

	// A single-axis override rescales the *other* axis to preserve the aspect
	// ratio — taking precedence over the native pixel size, so `--width 1.5`
	// alone produces a proportional height rather than the original height.
	if (
		widthOverride !== undefined &&
		heightOverride === undefined &&
		aspect !== undefined
	) {
		heightEmu = Math.round(widthOverride * aspect);
	}
	if (
		heightOverride !== undefined &&
		widthOverride === undefined &&
		aspect !== undefined
	) {
		widthEmu = Math.round(heightOverride / aspect);
	}

	if (widthEmu === undefined || heightEmu === undefined) return null;

	// Default-sized image wider than the page content area → scale down to fit,
	// preserving aspect. Skipped when the caller set an explicit width.
	if (
		widthOverride === undefined &&
		maxWidthEmu !== undefined &&
		maxWidthEmu > 0 &&
		widthEmu > maxWidthEmu
	) {
		const scale = maxWidthEmu / widthEmu;
		widthEmu = maxWidthEmu;
		heightEmu = Math.round(heightEmu * scale);
	}

	return { widthEmu, heightEmu };
}
