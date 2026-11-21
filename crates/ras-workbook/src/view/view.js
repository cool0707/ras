// The grid of `ras book view` (and of the editor's sheet panel): draws the sheets of DATA.
// DATA = { file, sheets: [sheet_view...], diff: { before, structure: [text], changes: { sheet: [{ a, before, after }] } } | null, active: name | null }
"use strict";
(function () {
    const data = window.DATA;
    const $ = (id) => document.getElementById(id);
    const el = (tag, cls, text) => {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (text !== undefined) e.textContent = text;
        return e;
    };
    const colName = (c) => {
        let s = "";
        for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
        return s;
    };
    const address = (r, c) => colName(c) + (r + 1);
    const parse = (a) => {
        const m = /^([A-Z]+)(\d+)$/.exec(a);
        if (!m) return null;
        let c = 0;
        for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
        return [+m[2] - 1, c - 1];
    };
    const changes = (data.diff && data.diff.changes) || {};

    // What is kept when the page is drawn again (the editor redraws it at every stop).
    const vscode = typeof acquireVsCodeApi === "function" ? acquireVsCodeApi() : null;
    const saved = (vscode && vscode.getState()) || {};
    const remember = (patch) => {
        Object.assign(saved, patch);
        if (vscode) vscode.setState(saved);
    };

    let current = null;
    let selected = null;

    function select(sheet, td, r, c) {
        if (selected) selected.classList.remove("sel");
        selected = td;
        td.classList.add("sel");
        const cell = td.cell || {};
        $("name").textContent = address(r, c);
        const fx = $("fx");
        fx.textContent = cell.f || cell.v || cell.t || "";
        const was = (changes[sheet.name] || []).find((x) => x.a === address(r, c));
        if (was) fx.appendChild(el("span", "was", `変更前: ${was.before}`));
    }

    function draw(sheet) {
        current = sheet;
        selected = null;
        remember({ sheet: sheet.name });
        const grid = $("grid");
        grid.textContent = "";
        $("name").textContent = "";
        $("fx").textContent = "";
        $("note").hidden = !sheet.truncated;
        $("note").textContent = sheet.truncated ? `使用範囲 ${sheet.usedRange} のうち、${sheet.rows} 行 × ${sheet.cols} 列までを表示しています（--all ですべて表示します）` : "";
        for (const b of $("tabs").children) b.classList.toggle("on", b.sheet === sheet);
        if (!sheet.cells.length) {
            grid.appendChild(el("div", "empty", "このシートは空です"));
            return;
        }
        const cells = new Map();
        for (const c of sheet.cells) cells.set(c.r * 16384 + c.c, c);
        // Cells covered by a merged cell are not drawn; its first cell spans them.
        const span = new Map();
        const covered = new Set();
        for (const [r1, c1, r2, c2] of sheet.merges) {
            span.set(r1 * 16384 + c1, [r2, c2]);
            for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.add(r * 16384 + c);
        }
        const changed = new Map((changes[sheet.name] || []).map((x) => [x.a, x]));
        const filter = sheet.filter ? sheet.filter.split(":").map(parse) : null;
        const cols = [];
        for (let c = 0; c < sheet.cols; c++) if (sheet.colWidths[c] > 0) cols.push(c);
        const table = el("table");
        const group = el("colgroup");
        group.appendChild(el("col")).style.width = "40px";
        let width = 40;
        for (const c of cols) {
            group.appendChild(el("col")).style.width = `${sheet.colWidths[c]}px`;
            width += sheet.colWidths[c];
        }
        table.style.width = `${width}px`;
        table.appendChild(group);
        const head = table.appendChild(el("thead")).appendChild(el("tr"));
        head.appendChild(el("th"));
        cols.forEach((c, i) => {
            const th = head.appendChild(el("th", null, colName(c)));
            // A hidden column follows.
            if (i + 1 < cols.length && cols[i + 1] !== c + 1) th.classList.add("gap");
        });
        const body = table.appendChild(el("tbody"));
        const byAddress = new Map();
        for (let r = 0; r < sheet.rows; r++) {
            if (!(sheet.rowHeights[r] > 0)) continue;
            const tr = body.appendChild(el("tr"));
            tr.style.height = `${sheet.rowHeights[r]}px`;
            if (r + 1 < sheet.rows && !(sheet.rowHeights[r + 1] > 0)) tr.classList.add("gap");
            tr.appendChild(el("th", null, String(r + 1)));
            for (let i = 0; i < cols.length; i++) {
                const c = cols[i];
                const key = r * 16384 + c;
                if (covered.has(key)) continue;
                const cell = cells.get(key);
                const td = tr.appendChild(el("td"));
                td.cell = cell;
                if (cell) {
                    if (cell.s) td.style.cssText = sheet.styles[cell.s];
                    if (cell.t !== undefined) td.textContent = cell.t;
                    // Text runs over the empty cells to its right, as in Excel.
                    const next = cols[i + 1];
                    if (cell.k === "string" && !/pre-wrap/.test(td.style.cssText) && !/text-align/.test(td.style.cssText) && next !== undefined && !cells.has(r * 16384 + next)) td.classList.add("ov");
                }
                const merged = span.get(key);
                if (merged) {
                    td.rowSpan = merged[0] - r + 1;
                    td.colSpan = cols.filter((x) => x >= c && x <= merged[1]).length;
                }
                const a = address(r, c);
                const was = changed.get(a);
                if (was) {
                    td.classList.add("chg");
                    td.title = `変更前: ${was.before}`;
                }
                if (filter && filter[0] && filter[1] && r === filter[0][0] && c >= filter[0][1] && c <= filter[1][1]) td.classList.add("filter");
                td.addEventListener("click", () => select(sheet, td, r, c));
                byAddress.set(a, td);
            }
        }
        grid.appendChild(table);
        draw.find = (a) => byAddress.get(a);
    }

    function go(sheetName, a) {
        const sheet = data.sheets.find((s) => s.name === sheetName);
        if (!sheet) return;
        if (sheet !== current) draw(sheet);
        const td = draw.find && draw.find(a);
        const at = parse(a);
        if (td && at) {
            td.scrollIntoView({ block: "center", inline: "center" });
            select(sheet, td, at[0], at[1]);
        }
    }

    $("file").textContent = data.file;
    const total = Object.values(changes).reduce((n, list) => n + list.length, 0);
    if (data.diff) {
        $("summary").textContent = total || data.diff.structure.length ? `${data.diff.before} からの変更: セル ${total} 個` : `${data.diff.before} からの変更はありません`;
    }
    for (const sheet of data.sheets) {
        const b = el("button", sheet.visible === "visible" ? null : "hidden", sheet.name + (sheet.visible === "visible" ? "" : "（非表示）"));
        b.sheet = sheet;
        if (sheet.tabColor) b.style.borderBottomColor = sheet.tabColor;
        const n = (changes[sheet.name] || []).length;
        if (n) b.appendChild(el("span", "badge", String(n)));
        b.addEventListener("click", () => draw(sheet));
        $("tabs").appendChild(b);
    }
    const side = $("side");
    if (data.diff && (total || data.diff.structure.length)) {
        side.appendChild(el("h2", null, "変更されたセル"));
        for (const line of data.diff.structure) side.appendChild(el("div", "structure", line));
        for (const sheet of data.sheets) {
            const list = changes[sheet.name] || [];
            if (!list.length) continue;
            side.appendChild(el("div", "sheet", sheet.name));
            for (const x of list.slice(0, 500)) {
                const row = el("div", "change");
                row.appendChild(el("b", null, x.a));
                row.appendChild(document.createTextNode(`${x.before} → ${x.after}`));
                row.title = `${x.a}: ${x.before} → ${x.after}`;
                row.addEventListener("click", () => go(sheet.name, x.a));
                side.appendChild(row);
            }
            if (list.length > 500) side.appendChild(el("div", "structure", `…ほか ${list.length - 500} 個`));
        }
    } else {
        side.hidden = true;
    }
    // The sheet shown last time, else the active one, else the first with changes.
    const first =
        data.sheets.find((s) => s.name === saved.sheet) ||
        data.sheets.find((s) => (changes[s.name] || []).length) ||
        data.sheets.find((s) => s.name === data.active) ||
        data.sheets.find((s) => s.visible === "visible") ||
        data.sheets[0];
    if (first) {
        draw(first);
        const list = changes[first.name] || [];
        if (list.length && saved.sheet !== first.name) go(first.name, list[0].a);
    } else {
        $("grid").appendChild(el("div", "empty", "シートがありません"));
    }
})();
