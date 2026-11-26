// A host that presses the cancel button of a progress bar when DoEvents asks 10 ms or more
// after the bar is shown (not while it is being shown): the forms shown hear the user while
// `sleep` / `waitUntil` wait (spec §10.4.1, §14.10.7).
"use strict";
let id = 0;
let pressed = false;
let shownAt = 0;
module.exports = {
    formHost: {
        send(msg) {
            if (msg.t === "show") {
                id = msg.id;
                pressed = false;
                shownAt = Date.now();
                console.log(`show modal=${msg.modal}`);
            }
        },
        receive() {
            return { t: "disconnected" };
        },
        poll() {
            if (pressed || Date.now() - shownAt < 10) return null;
            pressed = true;
            console.log("> (DoEvents) click btnCancel");
            return { t: "click", control: "btnCancel", id };
        },
    },
};
