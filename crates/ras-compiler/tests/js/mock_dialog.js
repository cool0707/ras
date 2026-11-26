// A host that plays a user of the dialogs built in code (spec §14.11.15): it answers each
// `show` with the next scripted session, prints the controls the dialog laid out and what it is
// told to draw, and presses the cancel button of a progress bar when DoEvents asks.
"use strict";
const sessions = [
    // choose(): the second button.
    [{ t: "click", control: "btn1" }],
    // The form: OK is refused while a rule fails, then the fields are filled in.
    [
        { t: "input", control: "f3", text: "toolong", seq: 1 },
        { t: "click", control: "btn0" },
        { t: "input", control: "f1_Text", text: "C:\\out.csv", seq: 2 },
        { t: "input", control: "f3", text: "abc", seq: 3 },
        { t: "click", control: "f4_1" },
        { t: "click", control: "f5_0" },
        { t: "click", control: "f5_2" },
        { t: "select", control: "f0", index: 1 },
        { t: "click", control: "f2" },
        { t: "click", control: "btn0" },
    ],
    // The same form again: the close button.
    [{ t: "close" }],
    // The progress bar: the run asks while it works.
    [],
];
let id = 0;
let queue = [];
let shown = 0;
let polls = 0;
const show = (c) => `  ${c.parent ? `${c.parent}/` : ""}${c.name}:${c.type} ${c.state.Left},${c.state.Top} ${c.state.Width}x${c.state.Height} "${c.state.Caption ?? c.state.Text ?? ""}"`;
module.exports = {
    formHost: {
        send(msg) {
            if (msg.t === "show") {
                id = msg.id;
                queue = (sessions[shown++] || []).slice();
                console.log(`show ${msg.name} modal=${msg.modal} "${msg.form.Caption}" ${msg.form.Width}x${msg.form.Height}`);
                for (const c of msg.controls) console.log(show(c));
            } else if (msg.t === "update") {
                for (const c of msg.controls) {
                    if (/_Error$|^f1_Text$|^btnCancel$|^pb_Value$/.test(c.name)) console.log(`  draw ${c.name} "${c.state.Caption ?? c.state.Text}" visible=${c.state.Visible} enabled=${c.state.Enabled}`);
                }
            } else {
                console.log(msg.t);
            }
        },
        receive() {
            const s = queue.shift();
            if (s) console.log(">", JSON.stringify(s));
            return s ? { ...s, id } : { t: "disconnected" };
        },
        poll() {
            polls++;
            if (polls !== 2) return null;
            console.log("> (DoEvents) click btnCancel");
            return { t: "click", control: "btnCancel", id };
        },
    },
};
