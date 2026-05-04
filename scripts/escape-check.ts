import { XmlNode } from "@core/parser";

console.log("Bun.escapeHTML samples:");
console.log(
	"  <:",
	JSON.stringify(Bun.escapeHTML("<script>alert(1)</script>")),
);
console.log("  &:", JSON.stringify(Bun.escapeHTML("a & b")));
console.log('  ":', JSON.stringify(Bun.escapeHTML('say "hi"')));
console.log("  ':", JSON.stringify(Bun.escapeHTML("don't")));

const node = XmlNode.element(
	"w:t",
	{ "xml:space": "preserve", title: 'has "quotes" & <stuff>' },
	[XmlNode.textNode("Hostile: <script>alert(1)</script> & <evil>")],
);

console.log("\nXmlNode.serialize output:");
console.log(XmlNode.serialize([node]));
