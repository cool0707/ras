// Smoke test of the ras_com addon on Windows, without Excel: drives COM objects that
// every Windows installation has (Scripting.Dictionary, WScript.Shell).
//
//   node crates/ras-com/tests/smoke.js <path to ras_com.node>
"use strict";
const assert = require("node:assert/strict");
const path = require("node:path");

// Progress goes out synchronously so that it survives a crash in native code.
const step = (name) => require("node:fs").writeSync(2, `[smoke] ${name}\n`);

const addonPath = path.resolve(process.argv[2]);
const com = require(addonPath);

step("create");
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
step("put with index");
// Property put with an index argument, then the default member.
dict.put("Item", ["a"], "あいう");
assert.equal(dict.call("", ["a"]), "あいう");
step("arrays out, For Each");
// A method that returns an array of Variants, and For Each.
assert.deepEqual(dict.call("Keys", []), ["a", "b", "when", "flag", "nothing"]);
assert.deepEqual(dict.enumerate(), ["a", "b", "when", "flag", "nothing"]);
step("arrays in");
// Arrays go in as SAFEARRAYs and come back 0-based.
dict.put("Item", ["list"], [1, 2, 3]);
step("1-D array stored");
assert.deepEqual(dict.call("Item", ["list"]), [1, 2, 3]);
step("1-D array read back");
dict.put("Item", ["grid"], [[1, "x"], [2, "y"]]);
step("2-D array stored");
assert.deepEqual(dict.call("Item", ["grid"]), [[1, "x"], [2, "y"]]);
step("2-D array read back");
step("object identity");
// Object references round-trip and compare by identity.
const inner = com.createObject("Scripting.Dictionary");
dict.put("Item", ["obj"], inner);
const back = dict.call("Item", ["obj"]);
assert.ok(back instanceof com.ComObject);
assert.ok(back.isSame(inner));
assert.ok(!back.isSame(dict));
step("type name");
const typeName = dict.typeName();
assert.equal(typeof typeName, "string");
console.log("TypeName(dictionary) =", typeName);
step("errors");
// Errors carry the VBA error number.
assert.throws(() => dict.call("Add", ["a", 1]), (e) => /RASCOM\{"number":457,/.test(e.message));
assert.throws(() => dict.call("NoSuchMember", []), (e) => /RASCOM\{"number":438,/.test(e.message));
assert.throws(() => com.createObject("No.Such.Class"), (e) => /RASCOM\{"number":429,/.test(e.message));
step("running object table");
assert.equal(com.findWorkbook("C:\\no\\such\\book.xlsm"), null);

step("run time");
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
step("WScript.Shell");
const shell = $rt.vba.CreateObject("WScript.Shell");
assert.equal($rt.call(shell, "ExpandEnvironmentStrings", ["%OS%"]), "Windows_NT");
assert.equal($rt.typeIs(shell, $rt.typeName(shell)), true);

step("dialog watcher");
// A call that does not return because a dialog is waiting for an answer: WScript.Shell is an
// in-process object, so its message box belongs to this very process. The watcher, on its own
// thread, has to find the box while the call is stuck and press the cancelling button.
// (The 20 s are the box's own timeout: without the watcher `Popup` returns -1 then.)
assert.equal(com.processOfWindow(0), 0);
assert.deepEqual(com.openDialogs(process.pid), []);
assert.throws(() => com.watchDialogs(process.pid, "sometimes"));
com.watchDialogs(process.pid, "dismiss");
const VB_OK_CANCEL = 1;
const ID_OK = 1;
const ID_CANCEL = 2;
assert.equal(shell.call("Popup", ["ras_com smoke test: OK / Cancel", 20, "ras_com", VB_OK_CANCEL]), ID_CANCEL);
step("dialog with a single answer");
assert.equal(shell.call("Popup", ["ras_com smoke test: OK only", 20, "ras_com", 0]), ID_OK);
step("dialog that is a real choice");
// Yes / No has no cancelling answer: it is reported but left alone, and times out by itself.
const VB_YES_NO = 4;
assert.equal(shell.call("Popup", ["ras_com smoke test: Yes / No", 2, "ras_com", VB_YES_NO]), -1);
com.watchDialogs(0, "off");
assert.equal(shell.call("Popup", ["ras_com smoke test: not watched", 1, "ras_com", VB_OK_CANCEL]), -1);

step("done");
console.log("ras_com smoke test passed");
