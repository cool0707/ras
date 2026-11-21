// Lists what the stand-in for Excel of the debug run time has (spec §14.10.8): the members of
// its objects, those that need Excel itself (error 445), and the VBA functions the run time
// implements. The compiler warns about the rest (`ras check --target stand-in`).
//
//   RAS_BOOK_ADDON=<ras_book.node> node tools/standin/members.js <build/debug of any project>
//
// The output is crates/ras-compiler/stdlib/standin.json; a test keeps the two the same
// (`cargo test -p ras --test offline stand_in`; RAS_UPDATE_STANDIN=1 rewrites the file).
"use strict";
const path = require("path");

const dir = path.resolve(process.argv[2] || "build/debug");
const $rt = require(path.join(dir, "ras-runtime.js"));
$rt.configure({ manifest: {}, root: dir, addon: null, book: process.env.RAS_BOOK_ADDON, saveOnExit: false });
$rt.hooks.note = () => {};
const host = $rt.standIn();

const out = {};
for (const [name, cls] of Object.entries(host.classes)) {
    const members = new Set();
    const excel = new Set();
    for (let p = cls.prototype; p && p !== Object.prototype; p = Object.getPrototypeOf(p)) {
        for (const key of Object.getOwnPropertyNames(p)) {
            if (key === "constructor" || key.startsWith("$") || !/^[A-Z]/.test(key)) continue;
            members.add(key);
            const d = Object.getOwnPropertyDescriptor(p, key);
            const source = String(d.value || d.get || "");
            // A member that ends by saying it needs Excel has no other way out.
            if (/needsExcel\([^;]*\);\s*}$/.test(source)) excel.add(key);
        }
    }
    out[name] = { members: [...members].sort(), excel: [...excel].sort() };
}
out.VBA = { members: Object.keys($rt.vba).filter((k) => /^[A-Za-z]/.test(k)).sort(), excel: [] };
process.stdout.write(`${JSON.stringify(out, null, 1)}\n`);
