import { describe, expect, test } from "bun:test";
import { SettingsView } from "@core/ast/document/settings";
import { XmlNode } from "@core/parser";

const SETTINGS = (inner: string): string =>
	`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:settings>`;

// `<w:trackChanges>` is a CT_OnOff toggle. Reading mere PRESENCE as on (ignoring
// `w:val="false"`) silently tracked every edit on a doc whose author turned
// tracking off — a high-cost wrong guess (see XmlNode.isToggleOn).
describe("SettingsView.isTrackChangesEnabled respects the CT_OnOff w:val", () => {
	test("a bare <w:trackChanges/> is ON", () => {
		const view = SettingsView.fromXml(SETTINGS("<w:trackChanges/>"));
		expect(view?.isTrackChangesEnabled()).toBe(true);
	});

	test('<w:trackChanges w:val="false"/> is OFF (not merely present)', () => {
		const view = SettingsView.fromXml(
			SETTINGS('<w:trackChanges w:val="false"/>'),
		);
		expect(view?.isTrackChangesEnabled()).toBe(false);
	});

	test('<w:trackChanges w:val="0"/> and "off" are OFF', () => {
		expect(
			SettingsView.fromXml(
				SETTINGS('<w:trackChanges w:val="0"/>'),
			)?.isTrackChangesEnabled(),
		).toBe(false);
		expect(
			SettingsView.fromXml(
				SETTINGS('<w:trackChanges w:val="off"/>'),
			)?.isTrackChangesEnabled(),
		).toBe(false);
	});

	test("absent <w:trackChanges> is OFF", () => {
		const view = SettingsView.fromXml(SETTINGS(""));
		expect(view?.isTrackChangesEnabled()).toBe(false);
	});

	test('turning ON a doc with an explicit w:val="false" toggle flips it (not a no-op)', () => {
		const view = SettingsView.fromXml(
			SETTINGS('<w:trackChanges w:val="false"/>'),
		);
		if (!view) throw new Error("view");
		view.setTrackChangesEnabled(true);
		expect(view.isTrackChangesEnabled()).toBe(true);
	});

	test("turning OFF removes the toggle regardless of its prior w:val", () => {
		const view = SettingsView.fromXml(SETTINGS("<w:trackChanges/>"));
		if (!view) throw new Error("view");
		view.setTrackChangesEnabled(false);
		expect(view.isTrackChangesEnabled()).toBe(false);
	});
});

describe("XmlNode.isToggleOn", () => {
	const toggle = (attrs: string): XmlNode => {
		const node = XmlNode.parse(`<w:b ${attrs}/>`)[0];
		if (!node) throw new Error("parse failed");
		return node;
	};

	test("bare element is on; true/1/on are on", () => {
		expect(toggle("").isToggleOn()).toBe(true);
		expect(toggle('w:val="true"').isToggleOn()).toBe(true);
		expect(toggle('w:val="1"').isToggleOn()).toBe(true);
		expect(toggle('w:val="on"').isToggleOn()).toBe(true);
	});

	test("false/0/off are off", () => {
		expect(toggle('w:val="false"').isToggleOn()).toBe(false);
		expect(toggle('w:val="0"').isToggleOn()).toBe(false);
		expect(toggle('w:val="off"').isToggleOn()).toBe(false);
	});
});
