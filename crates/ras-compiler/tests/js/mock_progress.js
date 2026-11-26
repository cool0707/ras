// A host that clicks the button of a form with a progress bar (spec §14.11.13) and prints the
// labels of the bar it is told to draw.
"use strict";
const steps = [{ t: "click", control: "btnStart" }];
let id = 0;
const show = (c) => `  ${c.name} "${c.state.Caption}" visible=${c.state.Visible} back=${c.state.BackColor} at ${c.state.Left},${c.state.Top} ${c.state.Width}x${c.state.Height}`;
module.exports = {
    formHost: {
        send(msg) {
            if (msg.t === "show") {
                id = msg.id;
                console.log("show", msg.controls.map((c) => c.name).join(","));
                for (const c of msg.controls) if (c.name.startsWith("pbUpload")) console.log(show(c));
                return;
            }
            if (msg.t === "update") {
                for (const c of msg.controls) if (c.name.startsWith("pbUpload")) console.log(show(c));
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
