// A host for `new frmAsk(...)` / `showModal()` (spec §14.11.9): each `show` gets the next session.
"use strict";
const sessions = [
    [{ t: "input", control: "txtAnswer", text: "赤" }, { t: "click", control: "btnOK" }],
    [{ t: "click", control: "btnCancel" }],
    [{ t: "close" }],
    [{ t: "click", control: "btnOK" }],
];
let queue = [];
let id = 0;
let shown = 0;
module.exports = {
    formHost: {
        send(msg) {
            if (msg.t === "show") {
                id = msg.id;
                queue = (sessions[shown++] || []).slice();
                console.log(`show ${msg.name}#${msg.id} "${msg.form.Caption}" modal=${msg.modal}`);
            } else if (msg.t === "hide" || msg.t === "unload") {
                console.log(`${msg.t} #${msg.id}`);
            }
        },
        receive() {
            const next = queue.shift();
            return next ? { ...next, id } : { t: "disconnected" };
        },
    },
};
