import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { describeGuides, snapToGuides } from "../../src/userForm/guides";

const form = { width: 300, height: 200 };
const r = (left: number, top: number, width = 40, height = 20) => ({ left, top, width, height });

describe("smart guides", () => {
	it("snap to the edges and centres of a sibling", () => {
		const other = r(100, 100);
		// Right edge at 143 is near the sibling's right edge (140).
		assert.deepEqual(snapToGuides(r(103, 30), [other], form, 4), { left: 100, top: 30, snappedX: true, snappedY: false });
		// Centre near the sibling's centre.
		const centred = snapToGuides(r(110, 30, 20), [other], form, 4);
		assert.equal(centred.left, 110);
		const lines = describeGuides(r(100, 30), [other], form).filter((g) => g.kind === "line");
		assert.ok(lines.some((g) => g.axis === "x" && g.at === 100 && g.from === 30 && g.to === 120));
		assert.ok(lines.some((g) => g.axis === "x" && g.at === 140));
	});

	it("snap 8 / 16 / 24 from a neighbour and from the container's edge", () => {
		const other = r(20, 50);
		// 62 is 2 from 60 (right edge of the neighbour) + ... 60 + 8 = 68 is the nearest.
		assert.equal(snapToGuides(r(70, 52), [other], form, 3).left, 68);
		assert.equal(snapToGuides(r(85, 52), [other], form, 3).left, 84);
		// From the top edge.
		assert.equal(snapToGuides(r(200, 15), [], form, 3).top, 16);
		const gaps = describeGuides(r(68, 52), [other], form).filter((g) => g.kind === "gap");
		assert.ok(gaps.some((g) => g.axis === "x" && g.from === 60 && g.to === 68 && g.label === "8"));
	});

	it("snap to the gap already between siblings in the row", () => {
		const a = r(10, 50);
		const b = r(80, 50); // gap 30
		const snapped = snapToGuides(r(152, 50), [a, b], form, 4);
		assert.equal(snapped.left, 150);
		const gaps = describeGuides(r(150, 50), [a, b], form).filter((g) => g.kind === "gap" && g.axis === "x");
		assert.deepEqual(
			gaps.map((g) => [g.from, g.to]).sort((p, q) => p[0] - q[0]),
			[
				[50, 80],
				[120, 150],
			],
		);
	});

	it("snap midway between two neighbours", () => {
		const a = r(0, 50);
		const b = r(200, 50);
		assert.equal(snapToGuides(r(98, 50), [a, b], form, 4).left, 100);
	});

	it("do not snap beyond the threshold or outside the container", () => {
		assert.deepEqual(snapToGuides(r(133, 77), [r(0, 0)], form, 2), { left: 133, top: 77, snappedX: false, snappedY: false });
		assert.equal(snapToGuides(r(261, 0), [], form, 4).left, 260);
	});
});
