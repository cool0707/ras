// A stand-in for Excel: one worksheet whose cells are objects with a `Value`.
// Loaded by the debug run time through RAS_MOCK_HOST.
"use strict";
const cells = new Map();
function cell(r, c) {
    const k = `${r},${c}`;
    if (!cells.has(k)) cells.set(k, { Row: r, Column: c, Value: null, Offset: (dr, dc) => cell(r + (dr || 0), c + (dc || 0)) });
    return cells.get(k);
}
function at(address) {
    const m = /^([A-Z])(\d+)$/.exec(address);
    return cell(Number(m[2]), m[1].charCodeAt(0) - 64);
}
const Sheet1 = {
    $typeName: "Worksheet",
    Cells: (r, c) => cell(r, c),
    Range(a, b) {
        if (typeof a === "string") return at(a);
        // The cells from `a` to `b`: a value written to them goes to each.
        return {
            set Value(v) {
                for (let r = a.Row; r <= b.Row; r++) for (let c = a.Column; c <= b.Column; c++) cell(r, c).Value = v;
            },
        };
    },
};
module.exports = { globals: { Sheet1 } };
