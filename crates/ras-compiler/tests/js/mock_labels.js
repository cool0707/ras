// A host that moves the focus out of txtA and clicks btnGo (spec §14.11.12, `errorLabel`).
"use strict";
const steps = [{ t: "focus", control: "txtA" }, { t: "focus", control: "btnGo" }, { t: "click", control: "btnGo" }];
let id = 0;
module.exports = { formHost: {
    send(msg) {
        if (msg.t === "show") { id = msg.id; return; }
        if (msg.t === "update") for (const c of msg.controls) if (c.type === "Label") console.log(`${c.name} "${c.state.Caption}" visible=${c.state.Visible}`);
    },
    receive() { const s = steps.shift(); return s ? { ...s, id } : { t: "disconnected" }; },
} };
