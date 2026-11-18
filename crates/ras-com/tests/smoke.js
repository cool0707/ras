// Smoke test of the ras_com addon on Windows, without Excel: drives COM objects that
// every Windows installation has (Scripting.Dictionary, WScript.Shell).
//
//   node crates/ras-com/tests/smoke.js <path to ras_com.node>
"use strict";
const assert = require("node:assert/strict");
const path = require("node:path");

const addonPath = path.resolve(process.argv[2]);
const com = require(addonPath);

// --- the addon itself
const dict = com.createObject("Scripting.Dictionary");
assert.ok(dict instanceof com.ComObject);
dict.call("Add", ["a", 1]);
dict.call("Add", ["b", 2.5]);
dict.call("Add", ["when", { $t: "date", v: 45000.5 }]);
dict.call("Add", ["flag", true]);
dict.call("Add", ["nothing", null]);
assert.equal(dict.get("Count", []), 5);
assert.equal(dict.call("Item", ["a"]), 1);
assert.equal(dict.call("Item", ["b"]), 2.5);
assert.deepEqual(dict.call("Item", ["when"]), { $t: "date", v: 45000.5 });
assert.equal(dict.call("Item", ["flag"]), true);
assert.equal(dict.call("Item", ["nothing"]), null);
assert.equal(dict.call("Exists", ["zz"]), false);
// Property put with an index argument, then the default member.
dict.put("Item", ["a"], "あいう");
assert.equal(dict.call("", ["a"]), "あいう");
// A method that returns an array of Variants, and For Each.
assert.deepEqual(dict.call("Keys", []), ["a", "b", "when", "flag", "nothing"]);
assert.deepEqual(dict.enumerate(), ["a", "b", "when", "flag", "nothing"]);
// Arrays go in as SAFEARRAYs and come back 0-based.
dict.put("Item", ["grid"], [[1, "x"], [2, "y"]]);
assert.deepEqual(dict.call("Item", ["grid"]), [[1, "x"], [2, "y"]]);
dict.put("Item", ["list"], [1, 2, 3]);
assert.deepEqual(dict.call("Item", ["list"]), [1, 2, 3]);
// Object references round-trip and compare by identity.
const inner = com.createObject("Scripting.Dictionary");
dict.put("Item", ["obj"], inner);
const back = dict.call("Item", ["obj"]);
assert.ok(back instanceof com.ComObject);
assert.ok(back.isSame(inner));
assert.ok(!back.isSame(dict));
const typeName = dict.typeName();
assert.equal(typeof typeName, "string");
console.log("TypeName(dictionary) =", typeName);
// Errors carry the VBA error number.
assert.throws(() => dict.call("Add", ["a", 1]), (e) => /RASCOM\{"number":457,/.test(e.message));
assert.throws(() => dict.call("NoSuchMember", []), (e) => /RASCOM\{"number":438,/.test(e.message));
assert.throws(() => com.createObject("No.Such.Class"), (e) => /RASCOM\{"number":429,/.test(e.message));
assert.equal(com.findWorkbook("C:\\no\\such\\book.xlsm"), null);

// --- through the JavaScript run time, as generated code uses it
const $rt = require("../../ras-compiler/src/js/runtime.js");
$rt.configure({ manifest: {}, root: ".", addon: addonPath });
const d = $rt.vba.CreateObject("Scripting.Dictionary");
$rt.call(d, "Add", ["k", $rt.date(45000)]);
assert.equal($rt.get(d, "Count"), 1);
assert.ok($rt.item(d, ["k"]) instanceof $rt.RDate);
assert.equal(String($rt.item(d, ["k"])), "2023/03/15");
assert.deepEqual([...$rt.iter(d)], ["k"]);
$rt.setItem(d, ["k"], 7);
assert.equal($rt.cnum($rt.item(d, ["k"])), 7);
let caught = null;
try {
    $rt.call(d, "Add", ["k", 1]);
} catch (e) {
    caught = $rt.caught(e);
}
assert.equal(caught && caught.code, 457);
assert.ok(caught.description.length > 0);
const shell = $rt.vba.CreateObject("WScript.Shell");
assert.equal($rt.call(shell, "ExpandEnvironmentStrings", ["%OS%"]), "Windows_NT");
assert.equal($rt.typeIs(shell, $rt.typeName(shell)), true);

console.log("ras_com smoke test passed");
