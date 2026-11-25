// A host that plays a user filling in a form validated by `FormValidator` (spec §14.11.12): it
// prints the error labels it is told to draw.
"use strict";
const steps = [
    { t: "focus", control: "txtName" },
    { t: "focus", control: "txtAge" },
    { t: "input", control: "txtAge", text: "abc", seq: 1 },
    { t: "focus", control: "cboKind" },
    { t: "focus", control: "txtAge" },
    { t: "input", control: "txtAge", text: "30", seq: 2 },
    { t: "select", control: "cboKind", index: 1 },
    { t: "click", control: "btnOK" },
    { t: "input", control: "txtName", text: "山田", seq: 3 },
    { t: "click", control: "chkAgree" },
    { t: "input", control: "txtCorp", text: "ABC", seq: 4 },
    { t: "click", control: "btnOK" },
];
let id = 0;
module.exports = {
    formHost: {
        send(msg) {
            if (msg.t === "show") { id = msg.id; console.log("show", msg.controls.map((c) => c.name).join(",")); return; }
            if (msg.t === "update") {
                for (const c of msg.controls) if (/Error$/.test(c.name)) console.log(`  ${c.parent ? c.parent + "/" : ""}${c.name} "${c.state.Caption}" visible=${c.state.Visible} back=${c.state.BackStyle} fore=${c.state.ForeColor} at ${c.state.Left},${c.state.Top} ${c.state.Width}x${c.state.Height}`);
                if (msg.removed) console.log("  removed", msg.removed);
                if (msg.focus) console.log("  focus", msg.focus);
                return;
            }
            console.log(msg.t, msg.message || "");
        },
        receive() {
            const s = steps.shift();
            if (s) console.log(">", JSON.stringify(s));
            return s ? { ...s, id } : { t: "disconnected" };
        },
    },
};
