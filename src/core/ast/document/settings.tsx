import { w } from "../../jsx";
import { XmlNode } from "../../parser";
import type { ContentTypesView } from "./content-types";
import type { Pkg } from "./package";
import type { RelationshipsView } from "./relationships";

const SETTINGS_PART_NAME = "word/settings.xml";
const SETTINGS_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const SETTINGS_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

const W_NAMESPACE =
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export class SettingsView {
	tree: XmlNode[];

	constructor(tree: XmlNode[]) {
		this.tree = tree;
	}

	/** Load this view from a package; returns undefined if the part is absent. */
	static async fromPackage(pkg: Pkg): Promise<SettingsView | undefined> {
		const tree = await pkg.readPart(SETTINGS_PART_NAME);
		return tree ? new SettingsView(tree) : undefined;
	}

	/** Parse a view from raw XML; returns undefined if the input is absent. */
	static fromXml(xml: string | undefined): SettingsView | undefined {
		return xml ? new SettingsView(XmlNode.parse(xml)) : undefined;
	}

	/** Serialize this view's tree into the package's `word/settings.xml`. */
	writeTo(pkg: Pkg): void {
		pkg.writeText(SETTINGS_PART_NAME, XmlNode.serialize(this.tree));
	}

	/** Mint the settings relationship + content-type override on the
	 * containing package and return a fresh empty view. Idempotent on the
	 * relationship target. Called by `Document.ensureSettings()`. */
	static register(deps: {
		relationships: RelationshipsView;
		contentTypes: ContentTypesView;
	}): SettingsView {
		if (!deps.relationships.hasTarget("settings.xml")) {
			deps.relationships.add(SETTINGS_RELATIONSHIP_TYPE, "settings.xml");
		}
		deps.contentTypes.registerPart(SETTINGS_PART_NAME, SETTINGS_CONTENT_TYPE);
		return new SettingsView([]);
	}

	isTrackChangesEnabled(): boolean {
		const root = XmlNode.findRoot(this.tree, "w:settings");
		if (!root) return false;
		return root.children.some((child) => child.tag === "w:trackChanges");
	}

	setTrackChangesEnabled(on: boolean): void {
		const root = this.ensureSettingsRoot();
		const hasTrackChanges = root.children.some(
			(child) => child.tag === "w:trackChanges",
		);
		if (on && !hasTrackChanges) {
			root.children.unshift(<w.trackChanges />);
		} else if (!on && hasTrackChanges) {
			root.children = root.children.filter(
				(child) => child.tag !== "w:trackChanges",
			);
		}
	}

	private ensureSettingsRoot(): XmlNode {
		const existing = XmlNode.findRoot(this.tree, "w:settings");
		if (existing) return existing;
		const fresh = <w.settings {...{ "xmlns:w": W_NAMESPACE }} />;
		this.tree.push(fresh);
		return fresh;
	}
}
