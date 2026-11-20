// A host for the forms of the debug run (spec §14.10.7) that plays a user: it answers each
// `show` with the next scripted session and prints what it is told to draw.
"use strict";
const sessions = [
    [
        { t: "focus", control: "txtName" },
        { t: "input", control: "txtName", text: "ng", seq: 1 },
        { t: "click", control: "btnOK" }, // refused: Exit cancels while the text is NG
        { t: "input", control: "txtName", text: "yamada", seq: 2 },
        { t: "select", control: "cboKind", index: 1 },
        { t: "click", control: "chkAgree" },
        // TripleState: off → Null → on → off.
        { t: "click", control: "chkMaybe" },
        { t: "click", control: "chkMaybe" },
        { t: "click", control: "chkMaybe" },
        { t: "click", control: "optM" },
        { t: "spin", control: "spnQty", delta: 1 },
        { t: "spin", control: "spnQty", delta: 1 },
        { t: "spin", control: "spnQty", delta: 1 }, // stays at Max
        { t: "click", control: "btnBoom" },
        { t: "close" }, // QueryClose keeps the form
        { t: "click", control: "btnOK", key: true },
    ],
    [{ t: "close" }],
];
let id = 0;
let queue = [];
let shown = 0;
module.exports = {
    globals: {},
    formHost: {
        send(msg) {
            if (msg.t === "show") {
                id = msg.id;
                queue = (sessions[shown++] || []).slice();
                const names = msg.controls.map((c) => `${c.parent ? `${c.parent}/` : ""}${c.name}:${c.type}`).join(" ");
                console.log(`show ${msg.name} modal=${msg.modal} "${msg.form.Caption}" ${names}`);
            } else if (msg.t === "update") {
                for (const c of msg.controls) {
                    if (c.name === "txtName") console.log(`draw txtName=${c.state.Text} seq=${c.inputSeq}`);
                }
                if (msg.focus) console.log(`focus ${msg.focus}`);
            } else if (msg.t === "error") {
                console.log(`error ${msg.message.replace(/\\/g, "/")}`);
            } else {
                console.log(msg.t);
            }
        },
        receive() {
            const next = queue.shift();
            return next ? { ...next, id } : { t: "disconnected" };
        },
    },
};
