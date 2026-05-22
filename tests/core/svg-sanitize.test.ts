import { describe, expect, test } from "bun:test";
import { sanitizeSvg } from "../../src/core/image/svg-sanitize";

function sanitize(svg: string): string {
	return new TextDecoder().decode(sanitizeSvg(new TextEncoder().encode(svg)));
}

describe("sanitizeSvg", () => {
	test("drops <script>, <style>, <foreignObject>, and animation elements", () => {
		const out = sanitize(`
			<svg xmlns="http://www.w3.org/2000/svg">
				<script>alert(1)</script>
				<style>@import url(http://evil/css);</style>
				<foreignObject><div xmlns="http://www.w3.org/1999/xhtml"><script>x</script></div></foreignObject>
				<animate attributeName="x" onbegin="evil()"/>
				<rect width="10" height="10"/>
			</svg>`);
		expect(out).not.toContain("<script");
		expect(out).not.toContain("<style");
		expect(out).not.toContain("foreignObject");
		expect(out).not.toContain("<animate");
		expect(out).toContain("<rect");
	});

	test("strips on* event handler attributes", () => {
		const out = sanitize(
			`<svg xmlns="http://www.w3.org/2000/svg"><rect onclick="x()" onload="y()" onmouseover="z()" width="10"/></svg>`,
		);
		expect(out).not.toContain("onclick");
		expect(out).not.toContain("onload");
		expect(out).not.toContain("onmouseover");
		expect(out).toContain('width="10"');
	});

	test("rejects javascript:, http(s):, file: URLs in href/xlink:href", () => {
		const out = sanitize(`
			<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
				<a href="javascript:alert(1)"><rect/></a>
				<image xlink:href="http://attacker/tracker"/>
				<image xlink:href="file:///etc/passwd"/>
			</svg>`);
		expect(out).not.toContain("javascript:");
		expect(out).not.toContain("attacker");
		expect(out).not.toContain("file:///");
	});

	test("rejects data:image/svg+xml (recursive-SVG bypass)", () => {
		const out = sanitize(
			`<svg xmlns="http://www.w3.org/2000/svg"><use href="data:image/svg+xml,&lt;svg/onload=alert(1)/&gt;"/></svg>`,
		);
		expect(out).not.toContain("data:image/svg");
		expect(out).not.toContain("onload");
	});

	test("allows same-doc fragments and raster data: URIs", () => {
		const out = sanitize(`
			<svg xmlns="http://www.w3.org/2000/svg">
				<defs><pattern id="p"><rect/></pattern></defs>
				<use href="#p"/>
				<image href="data:image/png;base64,iVBOR"/>
			</svg>`);
		expect(out).toContain('href="#p"');
		expect(out).toContain("data:image/png;base64");
	});

	test("XXE is rejected at parse time (fast-xml-parser)", () => {
		expect(() =>
			sanitize(
				`<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><svg/>`,
			),
		).toThrow();
	});
});
