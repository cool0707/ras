// A stand-in for Excel: one worksheet whose cells live in a Map.
// Loaded by the debug run time through RAS_MOCK_HOST.
"use strict";
const cells = new Map();
const rows = [
    ["顧客", "", "単価", "数量"],
    ["A社", "", 100, 2],
    ["B社", "", 250.5, 4],
    ["", "", null, 1],
    ["A社", "", "30", 3],
];
const key = (r, c) => `${r},${c}`;
rows.forEach((row, r) => row.forEach((v, c) => cells.set(key(r + 1, c + 1), { Value: v, Formula: "" })));
function cell(r, c) {
    if (!cells.has(key(r, c))) cells.set(key(r, c), { Value: null, Formula: "" });
    return Object.assign(cells.get(key(r, c)), { Row: r, End: () => ({ Row: rows.length }) });
}
const state = { ScreenUpdating: true, EnableEvents: true, Calculation: -4105, DisplayAlerts: true, StatusBar: false, Cursor: -4143, Interactive: true };
const log = [];
const Application = new Proxy(state, {
    set(target, name, value) {
        log.push(`${String(name)}=${value}`);
        target[name] = value;
        return true;
    },
});
const Sheet1 = {
    $typeName: "Worksheet",
    Name: "Data",
    Cells: (r, c) => cell(r, c),
    Rows: { Count: 1048576 },
    Range: () => ({ Columns: () => ({ Value: Array.from({ length: 6 }, (_, i) => [cell(i + 2, 1).Value]) }) }),
};
module.exports = { globals: { Application, Sheet1 } };
process.on("exit", () => {
    console.log(`E2=${cell(2, 5).Formula}`);
    console.log(`E5=${cell(5, 5).Formula}`);
    console.log(log.filter((l) => /^(StatusBar|ScreenUpdating)=/.test(l)).join(" "));
});
