// A host that pastes a quoted path into a file picker (spec §14.11.14), leaves the box, clicks
// its button (the dialog chooses two files) and then OK. `Application.FileDialog` prints how it
// was set up.
"use strict";
const steps = [
    { t: "focus", control: "fpInput_Text" },
    { t: "input", control: "fpInput_Text", text: "\"C:\\My Docs\\a.xlsx\"", seq: 1 },
    { t: "focus", control: "btnOK" },
    { t: "click", control: "fpInput_Button" },
    { t: "click", control: "btnOK" },
];
let id = 0;
const dialog = (type) => {
    const filters = [];
    const d = {
        Title: "", ButtonName: "", AllowMultiSelect: false, InitialFileName: "", InitialView: 1, FilterIndex: 1,
        Filters: { Clear() { filters.length = 0; }, Add(a, b) { filters.push(`${a}=${b}`); }, get Count() { return filters.length; } },
        SelectedItems: { get Count() { return 2; }, Item(i) { return ["C:\\x.xlsx", "C:\\y.xlsm"][i - 1]; } },
        Show() {
            console.log(`dialog ${type} "${d.Title}" multi=${d.AllowMultiSelect} filters=${filters.join(",")} initial=${d.InitialFileName}`);
            return -1;
        },
    };
    return d;
};
module.exports = {
    globals: { Application: { FileDialog: dialog } },
    formHost: {
        send(msg) {
            if (msg.t === "show") { id = msg.id; console.log("show", msg.controls.map((c) => `${c.name}:${c.type || ""}`).join(",")); return; }
            if (msg.t === "update") {
                for (const c of msg.controls) if (c.name.startsWith("fpInput")) console.log(`  ${c.name} "${c.state.Text ?? c.state.Caption}" at ${c.state.Left},${c.state.Top} ${c.state.Width}x${c.state.Height}`);
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
