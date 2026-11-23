// The grid of `ras book view` (and of the editor's sheet panel): draws the sheets of DATA.
// (A sheet: `filter` is its filter's range, `filterColumns` { offset: { values, blanks } | { other } } the columns filtered.)
// DATA = { editable?: true, file, sheets: [sheet_view...], images: [data URL], diff: { before, structure: [text], changes: { sheet: [{ a, before, after }] } } | null, active: name | null, focus?: true }
// (`focus`: `active` was asked for (`ras book view --sheet`) and comes before the sheet shown last time.)
//
// The grid goes on past the cells as Excel's does: empty rows and columns are added while the
// sheet is scrolled to its end. The objects of the sheet's drawing (shapes, pictures, controls)
// lie over the cells, placed by their anchors on the grid as it is drawn.
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
    const KEY = 16384;
    // The height of the column headers, until it is measured (px).
    const HEAD_H = 24;
    // How far past the visible part the grid is kept drawn (px).
    const AHEAD = 400;

    // What is kept when the page is drawn again (the editor redraws it at every stop).
    const vscode = typeof acquireVsCodeApi === "function" ? acquireVsCodeApi() : null;
    const saved = (vscode && vscode.getState()) || {};
    const remember = (patch) => {
        Object.assign(saved, patch);
        if (vscode) vscode.setState(saved);
    };

    let current = null;
    // The object (a shape, a control) selected, and the cell of the active cell.
    let selected = null;
    let selectedObj = null;
    let activeTd = null;
    // The sheet as drawn: its table, the rows and columns so far, where they are.
    let view = null;

    // Cells are edited only in the editor's panel, on a page made for it: the editor writes the
    // workbook file and draws the page again.
    const editable = !!(data.editable && vscode);
    const MAXR = 1048576;
    const MAXC = 16384;
    // The selection: { kind: "cells" | "rows" | "cols" | "all", nr, nc (where it was started), fr, fc
    // (where it reaches), ar, ac (the active cell), r1, c1, r2, c2 }. Kept by sheet.
    let sel = null;
    // The other areas of a selection made with Ctrl: [{ kind, r1, c1, r2, c2 }] (sel is the last).
    let areas = [];
    const sels = saved.sels || {};
    // What is being typed: { mode: "enter" | "edit", where: "cell" | "bar", original }.
    let entry = null;
    // An edit was posted: nothing more until the page is drawn again, or the edit is refused.
    let busy = false;
    let drag = null;
    // The cells copied or cut: { sheet, kind, r1, c1, r2, c2, cut, text }.
    let clip = saved.clip || null;

    function clearSelection() {
        if (selected) {
            selected.classList.remove("on");
            selected.classList.remove("framed");
            for (const f of selected.querySelectorAll(".objframe")) f.remove();
        }
        selected = null;
        selectedObj = null;
        refreshEventButton();
        refreshShapeTools();
        if (activeTd) activeTd.classList.remove("sel");
        activeTd = null;
        if (view) {
            view.refs.textContent = "";
            view.selbox.hidden = true;
            view.fillh.hidden = true;
            for (const th of view.hl) th.classList.remove("hs");
            view.hl = [];
        }
    }

    // ---- formulas: the references in them, colored as Excel colors them while one is edited

    const REF_COLORS = ["#2f6fdb", "#d64545", "#7b3fb8", "#2e9e4f", "#c46a00", "#0097a7", "#c0398a"];
    const REF = /(?<![\p{L}\p{N}_.$\]])((?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[0-9]+:\$?[0-9]+)(?![\p{L}\p{N}_(!])/gu;
    const colNumber = (letters) => {
        let c = 0;
        for (const ch of letters) c = c * 26 + ch.charCodeAt(0) - 64;
        return c - 1;
    };
    // A reference as rows and columns: [r1, c1, r2, c2] (whole columns and rows to the ends).
    function area(text) {
        const ends = text.replace(/\$/g, "").split(":");
        const one = (t) => {
            const m = /^([A-Z]*)([0-9]*)$/.exec(t);
            return m ? [m[2] ? +m[2] - 1 : null, m[1] ? colNumber(m[1]) : null] : [null, null];
        };
        const [a, b] = [one(ends[0]), one(ends[1] || ends[0])];
        const rows = [a[0], b[0]].map((x, i) => (x === null ? (i ? 1048575 : 0) : x));
        const cols = [a[1], b[1]].map((x, i) => (x === null ? (i ? 16383 : 0) : x));
        return [Math.min(...rows), Math.min(...cols), Math.max(...rows), Math.max(...cols)];
    }
    // The references of a formula outside its strings: [{ start, end, sheet, area }].
    function references(formula) {
        const out = [];
        // Strings ("..." with "" inside) are left out.
        const masked = formula.replace(/"(?:[^"]|"")*"/g, (m) => " ".repeat(m.length));
        for (const m of masked.matchAll(REF)) {
            const sheet = m[1] ? m[1].slice(0, -1).replace(/^'(.*)'$/, "$1").replace(/''/g, "'") : null;
            out.push({ start: m.index, end: m.index + m[0].length, sheet, area: area(m[2]) });
        }
        return out;
    }

    // The colors of the references of a formula, and their boxes on the grid (those of this sheet).
    function colorRefs(sheet, formula, refs) {
        const colors = new Map();
        for (const ref of refs) {
            const key = `${ref.sheet || ""}!${ref.area.join(",")}`;
            if (!colors.has(key)) colors.set(key, REF_COLORS[colors.size % REF_COLORS.length]);
            ref.color = colors.get(key);
            if (!ref.sheet || ref.sheet === sheet.name) ref.box = outline(ref.area, ref.color);
        }
        return refs;
    }

    // The name box and the formula bar for the active cell.
    function showCell(sheet, cell, r, c) {
        cell = cell || {};
        view.refs.textContent = "";
        showName(sheet, r, c);
        const fx = $("fx");
        fx.textContent = "";
        if (cell.f) {
            // The references in colors, the same on the grid.
            const refs = colorRefs(sheet, cell.f, references(cell.f));
            let at = 0;
            for (const ref of refs) {
                fx.appendChild(document.createTextNode(cell.f.slice(at, ref.start)));
                fx.appendChild(el("span", "ref", cell.f.slice(ref.start, ref.end))).style.color = ref.color;
                at = ref.end;
            }
            fx.appendChild(document.createTextNode(cell.f.slice(at)));
        } else {
            fx.textContent = cell.v || cell.t || "";
        }
        const was = (changes[sheet.name] || []).find((x) => x.a === address(r, c));
        if (was) fx.appendChild(el("span", "was", `変更前: ${was.before}`));
    }

    // A box around the cells of a reference on the grid (as far as it is drawn).
    function outline([r1, c1, r2, c2], color) {
        const v = view;
        if (!v || r1 >= v.nrows || c1 >= v.ncols) return;
        r2 = Math.min(r2, v.nrows - 1);
        c2 = Math.min(c2, v.ncols - 1);
        const x = columnLeft(v, c1),
            y = rowTop(v, r1);
        const x2 = columnLeft(v, c2 + 1),
            y2 = rowTop(v, r2 + 1);
        const box = v.refs.appendChild(el("div", "refbox"));
        Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${Math.max(2, x2 - x + 1)}px`, height: `${Math.max(2, y2 - y + 1)}px`, borderColor: color, background: `${color}14` });
        pinTo(box, pinOf(v, { x, y, w: x2 - x, h: y2 - y }));
        return box;
    }

    // ---- the formula view (Excel's 数式の表示): formulas in the cells, columns twice as wide

    let formulas = !!saved.formulas;
    function showFormulas(on) {
        formulas = on;
        remember({ formulas: on });
        $("formulas").classList.toggle("on", on);
        $("formulas").setAttribute("aria-pressed", String(on));
        if (current) {
            const grid = $("grid");
            const [x, y] = [grid.scrollLeft, grid.scrollTop];
            draw(current);
            scrollTo(x, y);
        }
    }

    function selectObject(box, obj) {
        if (entry) cancelEntry();
        clearSelection();
        selected = box;
        selectedObj = obj;
        box.classList.add("on");
        addHandles(box, obj);
        refreshEventButton();
        refreshShapeTools();
        $("name").value = obj.name || "";
        $("fx").textContent = obj.macro ? `マクロ: ${obj.macro}` : "";
    }

    // The width of a column and the height of a row (px), drawn or not.
    function width(sheet, c) {
        const k = formulas ? 2 : 1;
        if (c < sheet.cols) return sheet.colWidths[c] * k;
        const w = sheet.moreWidths && sheet.moreWidths[c];
        return (w !== undefined ? w : sheet.defaultWidth) * k;
    }
    function height(sheet, r) {
        if (r < sheet.rows) return sheet.rowHeights[r];
        const h = sheet.moreHeights && sheet.moreHeights[r];
        return h !== undefined ? h : sheet.defaultHeight;
    }

    // Numbers that do not fit their column, as Excel shows them: the General format with fewer
    // digits, the others (and General when nothing fits) as #### across the cell.
    const measurer = document.createElement("canvas").getContext("2d");
    const fonts = new Map(); // sheet and style -> the CSS font of its cells
    function textWidth(td, cell, text) {
        const k = `${current.name}\u0000${cell.s || 0}`;
        let font = fonts.get(k);
        if (font === undefined) {
            font = getComputedStyle(td).font;
            fonts.set(k, font);
        }
        measurer.font = font;
        return measurer.measureText(text).width;
    }
    // The room for text in the cell at column c (merged: to its last column): a column n
    // characters wide has n digits' room, the rest is the cell's padding.
    function roomOf(c, merged) {
        let px = 0;
        for (let x = c; x <= (merged ? merged[1] : c); x++) px += width(current, x);
        return px - 5;
    }
    // "Shrink to fit": a smaller font, until the text fits the cell.
    function shrinkToFit(td, cell, text, c, merged) {
        const w = textWidth(td, cell, text);
        const room = roomOf(c, merged);
        if (w > room && w > 0) td.style.fontSize = `${(parseFloat(getComputedStyle(td).fontSize) * room) / w}px`;
    }
    function fitNumber(td, cell, c, merged) {
        const room = roomOf(c, merged);
        const fits = (s) => textWidth(td, cell, s) <= room;
        if (fits(cell.t)) return;
        const shown = typeof cell.x === "number" && !cell.n ? generalFit(cell.x, fits) : null;
        if (shown !== null) {
            td.textContent = shown;
            return;
        }
        const n = Math.max(0, Math.floor(room / textWidth(td, cell, "#")));
        td.textContent = "#".repeat(n);
        td.style.textAlign = "left";
    }
    // What the General format shows of n where `fits` says what fits: fewer decimals, then
    // scientific notation, then 0 (as Excel narrows it); null when nothing fits. At most 11
    // characters besides the sign; under 1E-4 fixed notation only with all the digits.
    function generalFit(n, fits) {
        if (n === 0) return fits("0") ? "0" : null;
        const sign = n < 0 ? "-" : "";
        const a = Number(Math.abs(n).toPrecision(15));
        const trim = (s) => (s.includes(".") ? s.replace(/0+$/, "").replace(/\.$/, "") : s);
        const ok = (s) => s.length - sign.length <= 11 && fits(s);
        const [mant, ex] = a.toExponential(14).split("e");
        const e = Number(ex);
        if (e >= -4) {
            for (let d = Math.max(0, 10 - Math.max(e, 0)); d >= 0; d--) {
                const s = sign + trim(a.toFixed(d));
                if (/[1-9]/.test(s) && ok(s)) return s;
            }
        } else {
            const s = `${sign}0.${"0".repeat(-e - 1)}${mant.replace(".", "").replace(/0+$/, "")}`;
            if (ok(s)) return s;
        }
        for (let d = 5; d >= 0; d--) {
            const [m, x] = a.toExponential(d).split("e");
            const s = `${sign}${trim(m)}E${Number(x) < 0 ? "-" : "+"}${String(Math.abs(Number(x))).padStart(2, "0")}`;
            if (ok(s)) return s;
        }
        if (a < 1 && fits(`${sign}0`)) return `${sign}0`;
        return null;
    }

    function draw(sheet) {
        // What was being typed on the sheet shown before is dropped.
        entry = null;
        ed.value = "";
        bar.remove();
        current = sheet;
        selected = null;
        activeTd = null;
        remember({ sheet: sheet.name });
        const grid = $("grid");
        grid.textContent = "";
        $("name").value = "";
        fillNames(sheet);
        $("fx").textContent = "";
        $("note").hidden = !sheet.truncated;
        $("note").textContent = sheet.truncated ? `使用範囲 ${sheet.usedRange} のうち、${sheet.rows} 行 × ${sheet.cols} 列までを表示しています（--all ですべて表示します）` : "";
        for (const b of $("tabs").children) b.classList.toggle("on", b.sheet === sheet);
        const cells = new Map();
        const filled = new Set();
        for (const c of sheet.cells) {
            cells.set(c.r * KEY + c.c, c);
            if (c.t !== undefined && c.t !== "") filled.add(c.r * KEY + c.c);
        }
        // Cells covered by a merged cell are not drawn; its first cell spans them.
        const span = new Map();
        const covered = new Set();
        // Every cell of a merged cell -> the merged cell.
        const owner = new Map();
        for (const m of sheet.merges) {
            const [r1, c1, r2, c2] = m;
            span.set(r1 * KEY + c1, [r2, c2]);
            for (let r = r1; r <= r2; r++)
                for (let c = c1; c <= c2; c++) {
                    owner.set(r * KEY + c, m);
                    if (r !== r1 || c !== c1) covered.add(r * KEY + c);
                }
        }
        const changed = new Map((changes[sheet.name] || []).map((x) => [x.a, x]));
        const filter = sheet.filter ? sheet.filter.split(":").map(parse) : null;
        const notes = new Map((sheet.notes || []).map((n) => [n.r * KEY + n.c, n]));
        const links = new Map((sheet.links || []).map((l) => [l.r * KEY + l.c, l]));
        // The buttons of the tables' headings (drawn; the grid filters only the sheet's filter).
        const tableButtons = new Map((sheet.tableFilters || []).map(([r, c, f]) => [r * KEY + c, f]));
        // The panes frozen: the rows above fr and the columns left of fc stay on the screen.
        const [fr, fc] = sheet.frozen || [0, 0];

        const wrap = el("div", "sheetwrap");
        const table = el("table");
        if (sheet.gridLines === false) table.classList.add("nogrid");
        const group = el("colgroup");
        const corner = group.appendChild(el("col"));
        table.appendChild(group);
        const head = table.appendChild(el("thead")).appendChild(el("tr"));
        head.appendChild(el("th", "corner"));
        const body = table.appendChild(el("tbody"));
        wrap.appendChild(table);
        const layer = wrap.appendChild(el("div", "objs"));
        const refs = wrap.appendChild(el("div", "refs"));
        const marks = wrap.appendChild(el("div", "refs"));
        const selbox = marks.appendChild(el("div", "selbox"));
        const clipbox = marks.appendChild(el("div", "clipbox"));
        const fillbox = marks.appendChild(el("div", "fillbox"));
        const fillh = marks.appendChild(el("div", "fillhandle"));
        // The button of a list (an input rule) and its message, at the active cell.
        const dropbtn = marks.appendChild(el("div", "dropbtn", "▾"));
        const promptbox = marks.appendChild(el("div", "promptbox"));
        // The lines under the rows and right of the columns frozen.
        const frzh = marks.appendChild(el("div", "frzline h"));
        const frzv = marks.appendChild(el("div", "frzline v"));
        selbox.hidden = clipbox.hidden = fillbox.hidden = fillh.hidden = dropbtn.hidden = promptbox.hidden = true;
        frzh.hidden = !fr;
        frzv.hidden = !fc;
        wrap.appendChild(ed);
        grid.appendChild(wrap);

        const v = (view = {
            sheet,
            wrap,
            table,
            body,
            cells,
            filled,
            owner,
            selbox,
            clipbox,
            fillbox,
            fillh,
            dropbtn,
            promptbox,
            frzh,
            frzv,
            notes,
            fr,
            fc,
            filter, // the filter's corners: [[r, c], [r, c]]
            hl: [], // the headers of the selection

            layer,
            cols: [], // the visible columns drawn, in order
            rows: [], // the visible rows drawn: [row, tr]
            trs: new Map(), // row -> tr
            ths: new Map(), // column -> its header
            ncols: 0, // columns drawn (visible or not)
            nrows: 0,
            width: 0, // of the cells (without the row headers)
            refs, // the boxes of the references of the selected formula
            head: 28, // the width of the row headers
            // Where the frozen rows start (px from the table's top, under the column headers),
            // and the frozen columns (px from the row headers).
            tops: [],
            lefts: [],
            byAddress: new Map(),
            // Right gridlines hidden under text that runs over empty cells: key -> true.
            noRight: new Set(),
        });

        function makeCell(tr, r, c, i) {
            const key = r * KEY + c;
            if (covered.has(key)) return;
            const cell = cells.get(key);
            const td = tr.appendChild(el("td"));
            td.cell = cell;
            td.r = r;
            td.c = c;
            if (r < fr || c < fc) {
                td.classList.add("frz");
                if (r < fr) td.classList.add("fzr");
                if (c < fc) td.classList.add("fzc");
            }
            if (notes.has(key)) {
                td.classList.add("hasnote");
                td.note = notes.get(key);
            }
            td.link = links.get(key) || formulaLink(cell);
            if (td.link) td.classList.add("haslink");
            if (cell) {
                if (cell.s) td.style.cssText = sheet.styles[cell.s];
                // The formula view: formulas (left, as text) and the values as they are typed.
                const shown = formulas ? cell.f || cell.v || cell.t : cell.hv ? "" : cell.t;
                if (formulas && cell.f) {
                    td.style.textAlign = "left";
                    td.classList.add("fml");
                }
                if (shown !== undefined) {
                    // Text runs over the empty cells to its right, as in Excel.
                    const next = v.cols[i + 1] !== undefined ? v.cols[i + 1] : c + 1;
                    const text = cell.k === "string" || (formulas && cell.f);
                    // (A heading of the filter keeps its text by its button.)
                    const head = (filter && filter[0] && filter[1] && r === filter[0][0] && c >= filter[0][1] && c <= filter[1][1]) || tableButtons.has(key);
                    if (text && !head && !/pre-wrap|--shrink|--rot/.test(td.style.cssText) && (!/text-align/.test(td.style.cssText) || (formulas && cell.f)) && !cells.has(r * KEY + next) && !span.has(key)) {
                        td.classList.add("ov");
                        td.appendChild(el("div", "ovt")).appendChild(el("span", null, shown));
                    } else if (/pre-wrap/.test(td.style.cssText)) {
                        // Wrapped text is cut at the row's height (a table row would grow to it),
                        // from its first line, as Excel draws it.
                        let h = 0;
                        for (let x = r; x <= (span.has(key) ? span.get(key)[0] : r); x++) h += height(sheet, x);
                        const box = td.appendChild(el("div", "wr", shown));
                        box.style.maxHeight = `${Math.max(0, h - 1)}px`;
                    } else {
                        td.textContent = shown;
                        const rot = td.style.getPropertyValue("--rot");
                        if (rot) {
                            td.textContent = "";
                            td.appendChild(el("span", "rot", shown)).style.transform = `rotate(${-Number(rot)}deg)`;
                        } else if (td.style.getPropertyValue("--shrink")) {
                            shrinkToFit(td, cell, shown, c, span.get(key));
                        } else if (!formulas && (cell.k === "number" || cell.k === "date")) {
                            fitNumber(td, cell, c, span.get(key));
                        }
                    }
                }
                if (cell.bar && !formulas) {
                    // A data bar under the value, 2px in from the cell's edges.
                    td.classList.add("cfb");
                    const bar = td.appendChild(el("div", "cfbar"));
                    bar.style.left = `calc(2px + (100% - 4px) * ${cell.bar.x})`;
                    bar.style.width = `calc((100% - 4px) * ${cell.bar.w})`;
                    bar.style.background = cell.bar.bg;
                    if (cell.bar.bd) bar.style.border = `1px solid ${cell.bar.bd}`;
                    if (typeof cell.bar.axis === "number") {
                        const axis = td.appendChild(el("div", "cfaxis"));
                        axis.style.left = `calc(2px + (100% - 4px) * ${cell.bar.axis})`;
                    }
                }
                if (td.link) td.title = `${td.link.tip || td.link.url || td.link.to} - クリックするとリンク先に移動します。\nクリックしたままにすると、セルを選択します。`;
                if (cell.ic && !formulas && sheet.icons && sheet.icons[cell.ic]) {
                    // An icon of an icon set, at the left of the cell, as high as the text.
                    td.classList.add("cfi");
                    const icon = td.appendChild(el("span", "cficon"));
                    const valign = td.style.verticalAlign;
                    if (valign === "middle" || valign === "top") icon.classList.add(valign === "top" ? "top" : "mid");
                    icon.innerHTML = sheet.icons[cell.ic];
                }
                if (cell.f && !formulas && !td.title) td.title = cell.f;
            }
            if (v.noRight.has(key)) td.classList.add("nogr");
            const merged = span.get(key);
            if (merged) {
                td.rowSpan = merged[0] - r + 1;
                td.colSpan = v.cols.filter((x) => x >= c && x <= merged[1]).length || 1;
            }
            const a = address(r, c);
            const was = changed.get(a);
            if (was) {
                td.classList.add("chg");
                td.title = `変更前: ${was.before}`;
            }
            if (filter && filter[0] && filter[1] && r === filter[0][0] && c >= filter[0][1] && c <= filter[1][1]) {
                td.classList.add("filter");
                if ((sheet.filterColumns || {})[c - filter[0][1]]) td.classList.add("filtered");
            } else if (tableButtons.has(key)) {
                td.classList.add("tfilter");
                if (tableButtons.get(key)) td.classList.add("filtered");
            }
            if (td.note) td.appendChild(el("span", "notemark"));
            // Excel's error check: a formula giving an error value is marked at its corner.
            if (cell && cell.f && cell.k === "error" && !formulas) {
                td.classList.add("haserr");
                td.appendChild(el("span", "errmark"));
            }
            if (r < fr || c < fc) pinCell(td);
            if (cell || was || merged || td.note) v.byAddress.set(a, td);
        }

        // A cell of the panes frozen: where it stays (known once the cells before it are drawn).
        function pinCell(td) {
            if (td.r < fr) td.style.top = `${v.tops[td.r] || 0}px`;
            if (td.c < fc) td.style.left = `${(v.lefts[td.c] || 0) + v.head}px`;
        }

        function addRow(r) {
            v.nrows = r + 1;
            const h = height(sheet, r);
            if (!(h > 0)) {
                // A hidden row: the row before it is marked.
                const last = v.rows[v.rows.length - 1];
                if (last) last[1].classList.add("gap");
                return;
            }
            const tr = body.appendChild(el("tr"));
            tr.style.height = `${h}px`;
            const th = tr.appendChild(el("th", null, String(r + 1)));
            th.r = r;
            if (r < fr) {
                th.classList.add("fzr");
                th.style.top = `${v.tops[r] || 0}px`;
            }
            v.cols.forEach((c, i) => makeCell(tr, r, c, i));
            v.rows.push([r, tr]);
            v.trs.set(r, tr);
        }

        function addCol(c) {
            v.ncols = c + 1;
            const w = width(sheet, c);
            if (!(w > 0)) {
                const last = head.lastChild;
                if (last && last !== head.firstChild) last.classList.add("gap");
                return;
            }
            group.appendChild(el("col")).style.width = `${w}px`;
            const th = head.appendChild(el("th", null, colName(c)));
            th.c = c;
            if (c < fc) {
                th.classList.add("fzc");
                th.style.left = `${(v.lefts[c] || 0) + v.head}px`;
            }
            v.ths.set(c, th);
            const i = v.cols.length;
            v.cols.push(c);
            v.width += w;
            for (const [r, tr] of v.rows) makeCell(tr, r, c, i);
        }

        // The row headers as wide as Excel draws them for the largest row number: 29 px with their
        // line up to 3 digits, 8 px more for each digit after that (Excel 2021 at 100%).
        function sizeHead() {
            const digits = String(v.nrows).length;
            const was = v.head;
            v.head = 28 + 8 * Math.max(0, digits - 3);
            corner.style.width = `${v.head}px`;
            table.style.width = `${v.head + v.width}px`;
            // The frozen columns stay right of the row headers.
            if (fc && v.head !== was) for (const x of table.querySelectorAll("td.fzc, th.fzc")) x.style.left = `${(v.lefts[x.c] || 0) + v.head}px`;
        }

        // The frozen rows and columns: where each starts (hidden ones take no room).
        for (let r = 0, y = HEAD_H; r < fr; r++) {
            v.tops[r] = y;
            y += Math.max(0, height(sheet, r));
        }
        for (let c = 0, x = 0; c < fc; c++) {
            v.lefts[c] = x;
            x += Math.max(0, width(sheet, c));
        }
        // The cells of the sheet, then as many empty ones as the window shows.
        for (let c = 0; c < sheet.cols; c++) addCol(c);
        for (let r = 0; r < sheet.rows; r++) addRow(r);
        sizeHead();
        Object.assign(v, { addRow, addCol, sizeHead });
        v.fill = function fill() {
            let more = false;
            const z = v.zoom || 1;
            const right = (grid.scrollLeft + grid.clientWidth + AHEAD) / z;
            while (v.ncols < sheet.maxCols && v.head + v.width < right) {
                addCol(v.ncols);
                more = true;
            }
            const bottom = (grid.scrollTop + grid.clientHeight + AHEAD) / z;
            let tall = table.offsetHeight;
            while (v.nrows < sheet.maxRows && tall < bottom) {
                const before = v.rows.length;
                addRow(v.nrows);
                if (v.rows.length > before) tall += height(sheet, v.nrows - 1);
                more = true;
            }
            if (more) sizeHead();
        };
        v.fill();
        // The frozen rows start under the column headers, as tall as they are drawn.
        const headH = table.tHead.offsetHeight;
        if (fr && headH !== HEAD_H) {
            for (let r = 0, y = headH; r < fr; r++) {
                v.tops[r] = y;
                y += Math.max(0, height(sheet, r));
            }
            for (const x of table.querySelectorAll("td.fzr, th.fzr")) x.style.top = `${v.tops[x.r] || 0}px`;
        }
        overflow(v);
        placeObjects(v);
        placeFrozenLines(v);
        // The zoom the sheet had here, else the file's.
        const zoom = zooms[sheet.name] || sheet.zoom || 100;
        if (zoom !== 100) setZoom(zoom, false);
        else showZoom();
        draw.find = (a) => v.byAddress.get(a);
        // The selection the sheet had (at first, the one saved in the file, else A1).
        let was = sels[sheet.name];
        if (!was && sheet.selection) {
            const [ar, ac, r1, c1, r2, c2] = sheet.selection;
            was = { kind: "cells", nr: ar, nc: ac, fr: ar === r1 ? r2 : r1, fc: ac === c1 ? c2 : c1 };
        }
        setSel(was || { kind: "cells", nr: 0, nc: 0, fr: 0, fc: 0 }, false);
    }

    // Text that runs over empty cells hides their gridlines, as far as it reaches; it stops at the
    // first cell that is not empty.
    function overflow(v) {
        const sheet = v.sheet;
        const cells = new Set(sheet.cells.filter((c) => c.t !== undefined && c.t !== "").map((c) => c.r * KEY + c.c));
        const runs = [];
        for (const [r, tr] of v.rows) {
            for (const td of tr.querySelectorAll("td.ov")) {
                const c = td.c;
                let room = width(sheet, c) - 6;
                const over = [];
                for (let n = c + 1; n < Math.min(sheet.maxCols, c + 60) && !cells.has(r * KEY + n); n++) {
                    room += width(sheet, n);
                    over.push(n);
                }
                const div = td.firstChild;
                div.style.width = `${Math.max(0, room)}px`;
                runs.push({ r, c, td, span: div.firstChild, over });
            }
        }
        // Read every width, then mark the cells.
        for (const run of runs) run.text = run.span.offsetWidth;
        for (const { r, c, td, text, over } of runs) {
            let reach = width(sheet, c) - 3;
            if (text <= reach) continue;
            td.classList.add("nogr");
            v.noRight.add(r * KEY + c);
            for (const n of over) {
                reach += width(sheet, n);
                if (text <= reach) break;
                v.noRight.add(r * KEY + n);
                const next = v.trs.get(r) && Array.from(v.trs.get(r).children).find((x) => x.c === n);
                if (next) next.classList.add("nogr");
            }
        }
    }

    // ---- zoom: the sheet drawn larger or smaller (CSS zoom on it). Inside it, sizes are those at
    // 100%; the window's scrolling is zoomed: z turns one into the other.

    const ZOOMS = [10, 25, 50, 75, 100, 125, 150, 200, 300, 400];
    const zooms = saved.zooms || {};
    const Z = () => (view && view.zoom) || 1;
    function setZoom(pct, keep) {
        const v = view;
        if (!v) return;
        pct = Math.max(10, Math.min(400, Math.round(pct)));
        const g = $("grid");
        const was = Z();
        const z = pct / 100;
        // The cell at the top left stays there.
        const [x, y] = [g.scrollLeft / was, g.scrollTop / was];
        v.zoom = z;
        v.wrap.style.zoom = z === 1 ? "" : String(z);
        zooms[v.sheet.name] = pct;
        remember({ zooms });
        showZoom();
        if (keep !== false) scrollTo(x * z, y * z);
    }
    function showZoom() {
        const pct = Math.round(Z() * 100);
        $("zoomlabel").textContent = `${pct}%`;
        $("zoomrange").value = String(pct);
    }
    function zoomStep(up) {
        const pct = Math.round(Z() * 100);
        const next = up ? ZOOMS.find((x) => x > pct) : [...ZOOMS].reverse().find((x) => x < pct);
        if (next) setZoom(next);
    }
    (function makeZoom() {
        const box = $("zoom");
        const minus = box.appendChild(el("button", null, "−"));
        minus.type = "button";
        minus.title = "縮小";
        minus.addEventListener("click", () => zoomStep(false));
        const range = box.appendChild(el("input"));
        range.id = "zoomrange";
        range.type = "range";
        range.min = "10";
        range.max = "400";
        range.step = "5";
        range.title = "ズーム";
        range.addEventListener("input", () => setZoom(+range.value));
        const plus = box.appendChild(el("button", null, "+"));
        plus.type = "button";
        plus.title = "拡大";
        plus.addEventListener("click", () => zoomStep(true));
        const label = box.appendChild(el("button", "pct", "100%"));
        label.id = "zoomlabel";
        label.type = "button";
        label.title = "100% に戻す";
        label.addEventListener("click", () => setZoom(100));
        for (const b of box.querySelectorAll("button, input")) b.addEventListener("mouseup", () => setTimeout(focusGrid));
    })();
    // Ctrl+wheel, as in Excel.
    $("grid").addEventListener(
        "wheel",
        (e) => {
            if (!(e.ctrlKey || e.metaKey) || !view) return;
            e.preventDefault();
            setZoom(Z() * 100 + (e.deltaY < 0 ? 10 : -10));
        },
        { passive: false },
    );

    // ---- the panes frozen: their cells are sticky; the boxes over them (objects, the selection,
    // references) move with the scrolling the other way, marked by `data-pin` ("y": in the
    // frozen rows, "x": in the frozen columns).

    function pinOf(v, b) {
        if (!b) return "";
        const fy = v.fr > 0 && b.y + b.h <= rowTop(v, v.fr) + 1;
        const fx = v.fc > 0 && b.x + b.w <= columnLeft(v, v.fc) + 1;
        return (fy ? "y" : "") + (fx ? "x" : "");
    }
    function pinTo(div, pin) {
        if (pin) div.dataset.pin = pin;
        else delete div.dataset.pin;
        applyPin(div);
    }
    function applyPin(div) {
        const g = $("grid");
        const p = div.dataset.pin || "";
        const z = Z();
        div.style.transform = p ? `translate(${p.includes("x") ? g.scrollLeft / z : 0}px,${p.includes("y") ? g.scrollTop / z : 0}px)` : "";
    }
    function repin() {
        if (view && (view.fr || view.fc)) for (const d of view.wrap.querySelectorAll("[data-pin]")) applyPin(d);
    }
    function placeFrozenLines(v) {
        if (v.fr) {
            Object.assign(v.frzh.style, { top: `${rowTop(v, v.fr) - 1}px`, width: `${v.head + v.width}px` });
            pinTo(v.frzh, "y");
        }
        if (v.fc) {
            Object.assign(v.frzv.style, { left: `${columnLeft(v, v.fc) - 1}px`, height: `${v.table.offsetHeight}px` });
            pinTo(v.frzv, "x");
        }
    }

    // Where a column starts and a row starts on the grid (px from its top left corner).
    function columnLeft(v, c) {
        const th = v.ths.get(c);
        if (th) return th.offsetLeft;
        // Hidden or not drawn: after the column drawn before it.
        let k = c - 1;
        while (k >= 0 && !v.ths.has(k)) k--;
        let x = k >= 0 ? v.ths.get(k).offsetLeft + v.ths.get(k).offsetWidth : v.head;
        for (let i = k + 1; i < c; i++) x += width(v.sheet, i);
        return x;
    }
    function rowTop(v, r) {
        const tr = v.trs.get(r);
        if (tr) return tr.offsetTop;
        // Hidden or not drawn: after the row drawn before it.
        let k = r - 1;
        while (k >= 0 && !v.trs.has(k)) k--;
        let y = k >= 0 ? v.trs.get(k).offsetTop + v.trs.get(k).offsetHeight : v.table.tHead.offsetHeight;
        for (let i = k + 1; i < r; i++) y += height(v.sheet, i);
        return y;
    }

    function placeObjects(v) {
        const objects = v.sheet.objects || [];
        if (!objects.length) return;
        const layer = v.layer;
        layer.textContent = "";
        const at = (m) => [columnLeft(v, m[0]) + m[1], rowTop(v, m[2]) + m[3]];
        for (const obj of objects) {
            let x, y;
            if (obj.from) [x, y] = at(obj.from);
            else if (obj.pos) [x, y] = [v.head + obj.pos[0], v.table.tHead.offsetHeight + obj.pos[1]];
            else continue;
            let x2, y2;
            if (obj.to) [x2, y2] = at(obj.to);
            else if (obj.ext) [x2, y2] = [x + obj.ext[0], y + obj.ext[1]];
            else continue;
            if (!(x2 >= x && y2 >= y)) continue;
            // Turned sideways (45° to 135°, 225° to 315°): Excel anchors its box turned a quarter.
            const r = (((obj.rot || 0) % 360) + 360) % 360;
            if ((r >= 45 && r < 135) || (r >= 225 && r < 315)) {
                const [cx, cy, hw, hh] = [(x + x2) / 2, (y + y2) / 2, (x2 - x) / 2, (y2 - y) / 2];
                [x, y, x2, y2] = [cx - hh, cy - hw, cx + hh, cy + hw];
            }
            const box = layer.appendChild(el("div", `obj ${obj.kind}`));
            Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${x2 - x}px`, height: `${y2 - y}px` });
            pinTo(box, pinOf(v, { x, y, w: x2 - x, h: y2 - y }));
            box.innerHTML = obj.html;
            if (obj.chart && window.rasChart) {
                try {
                    window.rasChart.draw(box.querySelector(".chartbox") || box, obj.chart);
                } catch (e) {
                    // A chart that cannot be drawn leaves its box empty.
                }
            }
            if (obj.name) box.title = obj.macro ? `${obj.name}（マクロ: ${obj.macro}）` : obj.name;
            obj.box = box;
            if (movable(obj)) box.classList.add("movable");
            box.addEventListener("click", (e) => {
                e.stopPropagation();
                selectObject(box, obj);
            });
            // Dragged: moved (in the panel that edits the workbook).
            box.addEventListener("mousedown", (e) => {
                if (e.button !== 0 || drawingShape) return;
                // (What cannot be moved is selected all the same: Delete and the arrows are not
                // for the object selected before it.)
                if (!movable(obj)) {
                    if (selectedObj !== obj) selectObject(box, obj);
                    return;
                }
                e.preventDefault();
                if (selectedObj !== obj) selectObject(box, obj);
                focusGrid();
                beginObjectDrag(e, box, obj, "move");
            });
            // As the form designer: a double click opens the default event, the menu the others.
            // Without events, a shape's text is edited.
            box.addEventListener("dblclick", (e) => {
                e.stopPropagation();
                if (!hasEvents(obj)) {
                    if (movable(obj) && obj.text !== undefined) {
                        e.preventDefault();
                        selectObject(box, obj);
                        editObjectText(obj);
                    }
                    return;
                }
                e.preventDefault();
                selectObject(box, obj);
                openEvent(obj, false);
            });
            box.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                selectObject(box, obj);
                const items = objectItems(obj);
                if (items.length) showMenu(e.clientX, e.clientY, items);
            });
            // Drawn again (the rows grew, the fonts came): still selected.
            if (selectedObj === obj) {
                selected = box;
                box.classList.add("on");
                addHandles(box, obj);
            }
        }
        for (const img of layer.querySelectorAll("img[data-img]")) {
            const src = (data.images || [])[+img.dataset.img];
            if (src) img.src = src;
        }
        warpText(layer);
    }

    function go(sheetName, a) {
        const sheet = data.sheets.find((s) => s.name === sheetName);
        if (!sheet) return;
        if (sheet !== current) draw(sheet);
        const td = draw.find && draw.find(a);
        const at = parse(a);
        if (td && at) {
            td.scrollIntoView({ block: "center", inline: "center" });
            setSel({ kind: "cells", nr: at[0], nc: at[1], fr: at[0], fc: at[1] }, false);
        }
    }
    // ---- the selection: the active cell and a range, whole rows or columns, as in Excel

    const ownerOf = (r, c) => (view && view.owner.get(r * KEY + c)) || null;

    // Draws the grid as far as the row r and the column c.
    function grow(r, c) {
        const v = view;
        const before = [v.nrows, v.ncols, v.head];
        while (v.ncols <= c && v.ncols < v.sheet.maxCols) v.addCol(v.ncols);
        while (v.nrows <= r && v.nrows < v.sheet.maxRows) v.addRow(v.nrows);
        if (v.nrows === before[0] && v.ncols === before[1]) return;
        v.sizeHead();
        if (v.ncols !== before[1]) refreshOverflow(v, before[1]);
        if (v.head !== before[2]) placeObjects(v);
    }

    function tdAt(r, c) {
        const o = ownerOf(r, c);
        if (o) [r, c] = o;
        const tr = view.trs.get(r);
        if (!tr) return null;
        for (const td of tr.children) if (td.c === c) return td;
        return null;
    }

    // Where cells are on the grid (as far as it is drawn): { x, y, w, h } or null.
    function boxOf(r1, c1, r2, c2) {
        const v = view;
        if (r1 >= v.nrows || c1 >= v.ncols) return null;
        r2 = Math.min(r2, v.nrows - 1);
        c2 = Math.min(c2, v.ncols - 1);
        const x = columnLeft(v, c1),
            y = rowTop(v, r1);
        return { x, y, w: columnLeft(v, c2 + 1) - x, h: rowTop(v, r2 + 1) - y };
    }

    // A range as it is written: B2, B2:C5, 3:5, B:D.
    function rangeText(s) {
        if (s.kind === "all") return `1:${MAXR}`;
        if (s.kind === "rows") return `${s.r1 + 1}:${s.r2 + 1}`;
        if (s.kind === "cols") return `${colName(s.c1)}:${colName(s.c2)}`;
        const a = address(s.r1, s.c1);
        return single(s) ? a : `${a}:${address(s.r2, s.c2)}`;
    }
    // One cell (or one merged cell).
    function single(s) {
        if (s.kind !== "cells") return false;
        const o = ownerOf(s.r1, s.c1) || [s.r1, s.c1, s.r1, s.c1];
        return o[0] === s.r1 && o[1] === s.c1 && o[2] === s.r2 && o[3] === s.c2;
    }

    function setSel(s, scroll) {
        if (!view) return;
        const sheet = view.sheet;
        const lim = (x, max) => Math.max(0, Math.min(max - 1, x));
        const kind = s.kind || "cells";
        const [nr, nc, fr, fc] = [lim(s.nr, sheet.maxRows), lim(s.nc, sheet.maxCols), lim(s.fr, sheet.maxRows), lim(s.fc, sheet.maxCols)];
        let [r1, c1, r2, c2] = [Math.min(nr, fr), Math.min(nc, fc), Math.max(nr, fr), Math.max(nc, fc)];
        let [ar, ac] = [nr, nc];
        if (kind === "all") [r1, c1, r2, c2, ar, ac] = [0, 0, MAXR - 1, MAXC - 1, 0, 0];
        else if (kind === "rows") [c1, c2, ac] = [0, MAXC - 1, 0];
        else if (kind === "cols") [r1, r2, ar] = [0, MAXR - 1, 0];
        else {
            // Merged cells are taken whole.
            for (let more = true; more; ) {
                more = false;
                for (const [a, b, c, d] of sheet.merges) {
                    if (a > r2 || c < r1 || b > c2 || d < c1 || (a >= r1 && c <= r2 && b >= c1 && d <= c2)) continue;
                    [r1, c1, r2, c2] = [Math.min(r1, a), Math.min(c1, b), Math.max(r2, c), Math.max(c2, d)];
                    more = true;
                }
            }
            const o = ownerOf(ar, ac);
            if (o) [ar, ac] = o;
        }
        sel = { kind, nr, nc, fr, fc, ar, ac, r1, c1, r2, c2 };
        // The other areas stay while they are added to (`keep`), or as they were kept (`areas`).
        if (s.areas) areas = s.areas;
        else if (!s.keep) areas = [];
        sels[sheet.name] = { kind, nr, nc, fr, fc, areas };
        remember({ sels });
        paint(scroll);
    }
    // Every area of the selection, the last one (that of the active cell) last.
    const allAreas = () => (sel ? areas.concat([{ kind: sel.kind, r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 }]) : []);
    // Excel refuses most edits of several areas.
    function oneArea() {
        if (!areas.length) return true;
        showError("この操作は複数の選択範囲に対して実行できません。");
        return false;
    }

    function paint(scroll) {
        const v = view;
        if (!v || !sel) return;
        clearSelection();
        grow(sel.ar, sel.ac);
        if (scroll) grow(sel.kind === "cols" ? 0 : sel.fr, sel.kind === "rows" ? 0 : sel.fc);
        activeTd = tdAt(sel.ar, sel.ac);
        if (activeTd) activeTd.classList.add("sel");
        showCell(v.sheet, activeTd && activeTd.cell, sel.ar, sel.ac);
        drawBoxes();
        if (scroll) reveal(sel.kind === "cols" ? null : sel.fr, sel.kind === "rows" ? null : sel.fc);
        placeEditor();
        showTools();
        showStats();
        showMode();
    }

    // The box of the range, its headers, and the dashed box of the cells copied.
    function drawBoxes() {
        const v = view;
        if (!v) return;
        const put = (div, b) => {
            div.hidden = !b;
            if (b) Object.assign(div.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w + 1}px`, height: `${b.h + 1}px` });
            pinTo(div, pinOf(v, b));
        };
        for (const th of v.hl) th.classList.remove("hs");
        v.hl = [];
        for (const box of v.wrap.querySelectorAll(".selbox.more")) box.remove();
        if (sel && !selected) {
            const all = allAreas();
            for (const [c, th] of v.ths) if (all.some((a) => c >= a.c1 && c <= a.c2)) v.hl.push(th);
            for (const [r, tr] of v.rows) if (all.some((a) => r >= a.r1 && r <= a.r2)) v.hl.push(tr.firstChild);
            for (const th of v.hl) th.classList.add("hs");
            put(v.selbox, single(sel) && !areas.length ? null : boxOf(sel.r1, sel.c1, sel.r2, sel.c2));
            v.selbox.classList.toggle("multi", areas.length > 0);
            for (const a of areas) {
                const box = v.selbox.parentNode.insertBefore(el("div", "selbox multi more"), v.selbox);
                put(box, boxOf(a.r1, a.c1, a.r2, a.c2));
            }
        }
        put(v.clipbox, clip && clip.sheet === v.sheet.name ? boxOf(clip.r1, clip.c1, clip.r2, clip.c2) : null);
        // The fill handle: at the bottom right corner of the cells selected.
        const b = editable && sel && !selected && !entry && sel.kind === "cells" && !areas.length ? boxOf(sel.r1, sel.c1, sel.r2, sel.c2) : null;
        v.fillh.hidden = !b || sel.r2 >= v.nrows || sel.c2 >= v.ncols;
        if (b) Object.assign(v.fillh.style, { left: `${b.x + b.w - 3}px`, top: `${b.y + b.h - 3}px` });
        pinTo(v.fillh, pinOf(v, b));
        placeList();
    }

    // Scrolls a cell out from under the headers and into the window (null: that way stays).
    function reveal(r, c) {
        const v = view;
        const g = $("grid");
        const o = (r !== null && c !== null && ownerOf(r, c)) || [r || 0, c || 0, r || 0, c || 0];
        const inner = boxOf(o[0], o[1], o[2], o[3]);
        if (!inner) return;
        // (In the window's units.)
        const z = Z();
        const b = { x: inner.x * z, y: inner.y * z, w: inner.w * z, h: inner.h * z };
        // (Under the column headers and the rows frozen; right of the columns frozen. A cell
        // frozen is always in the window that way.)
        const top = rowTop(v, v.fr) * z;
        const left = columnLeft(v, v.fc) * z;
        if (c !== null && o[1] >= v.fc) {
            if (b.x - left < g.scrollLeft) g.scrollLeft = b.x - left;
            else if (b.x + b.w + 1 > g.scrollLeft + g.clientWidth) g.scrollLeft = Math.min(b.x - left, b.x + b.w + 1 - g.clientWidth);
        }
        if (r !== null && o[0] >= v.fr) {
            if (b.y - top < g.scrollTop) g.scrollTop = b.y - top;
            else if (b.y + b.h + 1 > g.scrollTop + g.clientHeight) g.scrollTop = Math.min(b.y - top, b.y + b.h + 1 - g.clientHeight);
        }
    }

    // Scrolls to where the grid was, drawing it that far.
    function scrollTo(x, y) {
        const v = view;
        const g = $("grid");
        const before = [v.nrows, v.ncols, v.head];
        const z = Z();
        while (v.ncols < v.sheet.maxCols && v.head + v.width < (x + g.clientWidth) / z) v.addCol(v.ncols);
        let tall = v.table.offsetHeight;
        while (v.nrows < v.sheet.maxRows && tall < (y + g.clientHeight) / z) {
            v.addRow(v.nrows);
            tall += Math.max(0, height(v.sheet, v.nrows - 1));
        }
        v.sizeHead();
        g.scrollLeft = x;
        g.scrollTop = y;
        v.fill();
        if (v.ncols !== before[1]) refreshOverflow(v, before[1]);
        if (v.head !== before[2]) placeObjects(v);
        drawBoxes();
        placeEditor();
    }

    // The next cell that is shown, from a cell (or from the edge of its merged cell).
    function step(r, c, dr, dc) {
        const s = view.sheet;
        const o = ownerOf(r, c) || [r, c, r, c];
        let nr = dr > 0 ? o[2] + 1 : dr < 0 ? o[0] - 1 : r;
        let nc = dc > 0 ? o[3] + 1 : dc < 0 ? o[1] - 1 : c;
        while (dr && nr >= 0 && nr < s.maxRows && !(height(s, nr) > 0)) nr += dr;
        while (dc && nc >= 0 && nc < s.maxCols && !(width(s, nc) > 0)) nc += dc;
        return nr < 0 || nc < 0 || nr >= s.maxRows || nc >= s.maxCols ? [r, c] : [nr, nc];
    }

    // Ctrl+arrow: to the edge of the cells with content (past them: as far as the grid is drawn).
    function jump(r, c, dr, dc) {
        const v = view;
        const has = (a, b) => v.filled.has(a * KEY + b);
        const inside = (a, b) => a >= 0 && b >= 0 && a < v.nrows && b < v.ncols;
        let [nr, nc] = [r + dr, c + dc];
        if (!inside(nr, nc)) return [r, c];
        if (has(r, c) && has(nr, nc)) {
            while (inside(nr + dr, nc + dc) && has(nr + dr, nc + dc)) [nr, nc] = [nr + dr, nc + dc];
            return [nr, nc];
        }
        while (inside(nr, nc) && !has(nr, nc)) [nr, nc] = [nr + dr, nc + dc];
        return [Math.max(0, Math.min(v.nrows - 1, nr)), Math.max(0, Math.min(v.ncols - 1, nc))];
    }

    function move(dr, dc, extend, far) {
        if (!sel) return;
        const next = far ? jump : step;
        if (extend) {
            const [r, c] = sel.kind === "cells" ? next(sel.fr, sel.fc, dr, dc) : next(sel.ar, sel.ac, dr, dc);
            const [nr, nc] = sel.kind === "cells" ? [sel.nr, sel.nc] : [sel.ar, sel.ac];
            setSel({ kind: "cells", nr, nc, fr: r, fc: c, keep: true }, true);
        } else {
            const [r, c] = next(sel.ar, sel.ac, dr, dc);
            setSel({ kind: "cells", nr: r, nc: c, fr: r, fc: c }, true);
        }
    }
    const cellSel = (r, c) => ({ kind: "cells", nr: r, nc: c, fr: r, fc: c });
    const areaSel = (r1, c1, r2, c2) => ({ kind: "cells", nr: r1, nc: c1, fr: r2, fc: c2 });

    // What the mouse is on: a cell, a header.
    function hitOf(t) {
        if (!t || !t.closest || !view || !view.table.contains(t)) return null;
        const td = t.closest("td");
        if (td && td.r !== undefined) return { kind: "cells", r: td.r, c: td.c };
        const th = t.closest("th");
        if (!th) return null;
        if (th.r !== undefined) return { kind: "rows", r: th.r, c: 0 };
        if (th.c !== undefined) return { kind: "cols", r: 0, c: th.c };
        return { kind: "all", r: 0, c: 0 };
    }
    const inSel = (hit) =>
        sel &&
        (hit.kind === "cells" ? hit.r >= sel.r1 && hit.r <= sel.r2 && hit.c >= sel.c1 && hit.c <= sel.c2 : hit.kind === sel.kind && (hit.kind === "all" || (hit.kind === "rows" ? hit.r >= sel.r1 && hit.r <= sel.r2 : hit.c >= sel.c1 && hit.c <= sel.c2)));

    $("grid").addEventListener("mousedown", (e) => {
        if (e.button !== 0 && e.button !== 2) return;
        if (e.target === ed) return;
        const hit = hitOf(e.target);
        if (!hit) return;
        e.preventDefault();
        const edge = e.button === 0 ? edgeOf(e) : null;
        const button = e.button === 0 && filterButton(e);
        closeMenu();
        // A formula being typed: the cell (the rows, the columns) clicked goes into it.
        if (entry && !edge && !button && hit.kind !== "all" && e.button === 0 && pointable()) {
            const at = [hit.r, hit.c];
            const p = entry.point;
            pointTo(e.shiftKey && p ? p.anchor : at, at, hit.kind);
            drag = `point:${hit.kind}`;
            return;
        }
        if (entry) commit(null);
        focusGrid();
        if (edge) {
            beginResize(edge, e);
            return;
        }
        if (e.button === 0 && hit.kind === "cells" && onBorder(e)) {
            beginMove(e, hit);
            return;
        }
        if (button) {
            setSel(cellSel(button.r, button.c), false);
            // (After this click, which closes what is open.)
            setTimeout(() => openFilter(button));
            return;
        }
        // A right click inside the selection keeps it for the menu.
        if (e.button === 2 && inSel(hit)) return;
        const td = e.target.closest && e.target.closest("td");
        if (e.button === 0 && td && td.link && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
            const press = (linkPress = { td, held: false });
            press.timer = setTimeout(() => {
                press.held = true;
                setSel(cellSel(td.r, td.c), false);
            }, 600);
            return;
        }
        if (e.shiftKey && sel && sel.kind === hit.kind && hit.kind !== "all") setSel({ kind: hit.kind, nr: sel.nr, nc: sel.nc, fr: hit.r, fc: hit.c, keep: true }, false);
        else if ((e.ctrlKey || e.metaKey) && e.button === 0 && sel && hit.kind !== "all" && sel.kind !== "all") {
            // Ctrl: one more area (Excel: up to the limit of its formulas; here 64).
            if (areas.length < 63) areas = allAreas();
            setSel({ kind: hit.kind, nr: hit.r, nc: hit.c, fr: hit.r, fc: hit.c, keep: true }, false);
        } else setSel({ kind: hit.kind, nr: hit.r, nc: hit.c, fr: hit.r, fc: hit.c }, false);
        if (e.button === 0 && hit.kind !== "all") drag = hit.kind;
    });
    document.addEventListener("mousemove", (e) => {
        if (resizing) {
            moveResize(e);
            return;
        }
        if (filling) {
            moveFill(e);
            return;
        }
        if (moving) {
            moveTo(e);
            return;
        }
        if (!drag) {
            hoverEdge(e);
            if (view) view.table.classList.toggle("mv", onBorder(e));
        }
        if (!drag || !sel) return;
        if (!(e.buttons & 1)) {
            drag = null;
            return;
        }
        if (drag.startsWith("point:")) {
            const kind = drag.slice(6);
            const hit = hitOf(document.elementFromPoint(e.clientX, e.clientY));
            const p = entry && entry.point;
            if (!p || !hit || hit.kind === "all" || (kind !== "cells" && hit.kind !== kind && hit.kind !== "cells")) return;
            if (kind === "cells" && hit.kind !== "cells") return;
            if (p.cur && p.cur[0] === hit.r && p.cur[1] === hit.c) return;
            pointTo(p.anchor, [hit.r, hit.c], kind);
            return;
        }
        const hit = hitOf(document.elementFromPoint(e.clientX, e.clientY));
        if (!hit || hit.kind === "all") return;
        // Along the headers, the cells under the mouse count by their row or column.
        if (drag !== "cells" && hit.kind !== drag && hit.kind !== "cells") return;
        if (drag === "cells" && hit.kind !== "cells") return;
        if (hit.r === sel.fr && hit.c === sel.fc) return;
        setSel({ kind: drag, nr: sel.nr, nc: sel.nc, fr: hit.r, fc: hit.c, keep: true }, false);
    });
    document.addEventListener("mouseup", (e) => {
        if (linkPress) {
            const press = linkPress;
            linkPress = null;
            clearTimeout(press.timer);
            const td = e.target.closest && e.target.closest("td");
            if (!press.held && td === press.td && e.button === 0) follow(td.link);
        }
        drag = null;
        if (resizing) endResize(e);
        if (filling) endFill();
        if (moving) endMove();
    });
    $("grid").addEventListener("dblclick", (e) => {
        const edge = edgeOf(e);
        if (edge) {
            autofit(edge);
            return;
        }
        if (filterButton(e)) return;
        const hit = hitOf(e.target);
        if (hit && hit.kind === "cells" && !entry) beginEntry("edit", null, "cell");
    });

    // ---- typing into cells
    //
    // The box typed into is always there, on the active cell: it has the focus while the grid is
    // worked on, so that what is typed (through an input method too), copied and pasted comes to it.

    const ed = el("textarea", "editor");
    const bar = el("textarea", "barinput");
    for (const t of [ed, bar]) {
        t.rows = 1;
        t.wrap = "off";
        t.spellcheck = false;
        t.setAttribute("autocomplete", "off");
    }
    ed.setAttribute("aria-label", "セル");
    const focusGrid = () => {
        if (!entry || entry.where === "cell") ed.focus({ preventScroll: true });
    };
    const onGrid = () => {
        const a = document.activeElement;
        return !a || a === document.body || a === ed || a === bar || a === $("grid");
    };

    function placeEditor() {
        if (!view || !sel) return;
        if (ed.parentNode !== view.wrap) view.wrap.appendChild(ed);
        const o = ownerOf(sel.ar, sel.ac) || [sel.ar, sel.ac, sel.ar, sel.ac];
        const b = boxOf(o[0], o[1], o[2], o[3]);
        if (!b) return;
        const on = !!entry && entry.where === "cell";
        ed.classList.toggle("on", on);
        Object.assign(ed.style, { left: `${b.x}px`, top: `${b.y}px`, width: "", height: "", font: "" });
        pinTo(ed, pinOf(view, b));
        if (!on) return;
        if (activeTd) ed.style.font = getComputedStyle(activeTd).font;
        // As wide and as tall as what is typed, the cell at least.
        ed.style.width = `${b.w + 1}px`;
        ed.style.height = `${b.h + 1}px`;
        ed.style.width = `${Math.max(b.w + 1, ed.scrollWidth + 10)}px`;
        ed.style.height = `${Math.max(b.h + 1, ed.scrollHeight + 2)}px`;
    }

    function beginEntry(mode, text, where) {
        if (!editable || busy || !sel || !view) return false;
        closeMenu();
        const cell = (activeTd && activeTd.cell) || {};
        const original = cell.f || cell.v || cell.t || "";
        // `point`: the reference being pointed at (clicked, or moved to by the arrows) while a
        // formula is typed: { start, end } in the text, `anchor` and `cur` its corners [r, c].
        entry = { mode, where, original, point: null };
        if (text === null) text = original;
        const fx = $("fx");
        fx.textContent = "";
        if (where === "bar") {
            fx.appendChild(bar);
            bar.value = text;
            bar.focus();
            bar.setSelectionRange(text.length, text.length);
        } else {
            // (What an input method is composing in the box stays as it is.)
            if (ed.value !== text) ed.value = text;
            fx.textContent = text;
            reveal(sel.ar, sel.ac);
            placeEditor();
            ed.focus({ preventScroll: true });
            ed.setSelectionRange(text.length, text.length);
        }
        document.body.classList.add("typing");
        entryRefs();
        placeList();
        showMode();
        return true;
    }

    function endEntry() {
        entry = null;
        closeAssist();
        showMode();
        ed.value = "";
        bar.remove();
        document.body.classList.remove("typing");
        if (view && sel) paint(false);
        focusGrid();
    }
    const cancelEntry = endEntry;

    // Enters what was typed: `how` is { move: [rows, columns] } (Enter, Tab), { fill: true }
    // (Ctrl+Enter: into every cell of the selection) or null (the cell is left by the mouse).
    function commit(how) {
        const e = entry;
        if (!e || !sel) return;
        const text = (e.where === "bar" ? bar : ed).value;
        // An entry the cell's list does not allow: refused, and typed on (as Excel's 再試行).
        const no = refused(text);
        if (no) {
            showError(no);
            (e.where === "bar" ? bar : ed).focus();
            return;
        }
        const fill = !!(how && how.fill) && (!single(sel) || areas.length > 0);
        const ranges = fill ? allAreas().map(rangeText) : [address(sel.ar, sel.ac)];
        const sheet = view.sheet.name;
        endEntry();
        if (how && how.move) move(how.move[0], how.move[1], false);
        if (text === e.original && !fill) return;
        postEdit(ranges.map((range) => (text === "" ? { op: "clear", sheet, range } : { op: "set", sheet, range, text })));
    }

    ed.addEventListener("input", () => {
        if (!entry) {
            // The first character typed starts an entry that replaces the cell's content.
            if (!beginEntry("enter", ed.value, "cell")) ed.value = "";
            return;
        }
        // Typing ends the pointing.
        entry.point = null;
        $("fx").textContent = ed.value;
        placeEditor();
        entryRefs();
        showMode();
        updateAssist();
    });
    bar.addEventListener("input", () => {
        if (!entry) return;
        entry.point = null;
        entryRefs();
        showMode();
        updateAssist();
    });
    // The caret moved: the arguments it is in.
    for (const t of [ed, bar]) {
        t.addEventListener("click", () => entry && updateAssist());
        t.addEventListener("keyup", (e) => {
            if (entry && /^(ArrowLeft|ArrowRight|Home|End)$/.test(e.key)) updateAssist();
        });
    }
    ed.addEventListener("compositionstart", () => {
        if (!entry && !beginEntry("enter", ed.value, "cell")) ed.blur();
    });
    $("fx").addEventListener("mousedown", (e) => {
        if (!editable || busy || !sel || e.target === bar) return;
        e.preventDefault();
        if (entry && entry.where === "cell") {
            // Goes on in the bar.
            const text = ed.value;
            entry.where = "bar";
            entry.mode = "edit";
            ed.value = "";
            placeEditor();
            $("fx").textContent = "";
            $("fx").appendChild(bar);
            bar.value = text;
            bar.focus();
        } else if (!entry) {
            beginEntry("edit", null, "bar");
        }
    });

    // ---- formulas typed: pointing at cells (a click, a drag, the arrows), F4, the references

    const entryInput = () => (entry && entry.where === "bar" ? bar : ed);
    // A reference may be pointed at: in a formula, after an operator, a parenthesis or a comma.
    function pointable() {
        if (!entry || !view) return false;
        const input = entryInput();
        const text = input.value;
        if (!/^[=+-]/.test(text)) return false;
        if (entry.point) return true;
        if (input.selectionStart !== input.selectionEnd) return false;
        return /[=(,+\-*/^&<>:;\s]$/.test(text.slice(0, input.selectionStart));
    }
    // The text of a reference from the corner a to the corner b ([r, c]); `kind`: cells, rows, cols.
    function refText(a, b, kind) {
        const [r1, r2] = [Math.min(a[0], b[0]), Math.max(a[0], b[0])];
        const [c1, c2] = [Math.min(a[1], b[1]), Math.max(a[1], b[1])];
        if (kind === "rows") return `${r1 + 1}:${r2 + 1}`;
        if (kind === "cols") return `${colName(c1)}:${colName(c2)}`;
        // A merged cell whole.
        const o1 = ownerOf(r1, c1) || [r1, c1, r1, c1];
        const o2 = ownerOf(r2, c2) || [r2, c2, r2, c2];
        const [x1, y1, x2, y2] = [Math.min(o1[0], o2[0]), Math.min(o1[1], o2[1]), Math.max(o1[2], o2[2]), Math.max(o1[3], o2[3])];
        return x1 === x2 && y1 === y2 ? address(x1, y1) : `${address(x1, y1)}:${address(x2, y2)}`;
    }
    function pointTo(anchor, cur, kind) {
        const input = entryInput();
        let p = entry.point;
        if (!p) {
            const at = input.selectionStart;
            p = entry.point = { start: at, end: at };
        }
        const text = refText(anchor, cur, kind || "cells");
        input.setRangeText(text, p.start, p.end, "end");
        Object.assign(p, { end: p.start + text.length, anchor, cur, kind: kind || "cells" });
        if (entry.where === "cell") $("fx").textContent = input.value;
        placeEditor();
        entryRefs();
        showMode();
        updateAssist();
    }
    // The references of what is typed, as boxes in their colors; the one pointed at dashed.
    function entryRefs() {
        if (!view || !entry) return;
        view.refs.textContent = "";
        const text = entryInput().value;
        if (!/^[=+-]/.test(text)) return;
        const refs = colorRefs(view.sheet, text, references(text));
        const p = entry.point;
        const ref = p && refs.find((r) => r.start === p.start && r.end === p.end);
        if (ref && ref.box) ref.box.classList.add("pointing");
    }
    // F4: the reference at the caret, $ by $ (A1, $A$1, A$1, $A1).
    function toggleAbsolute() {
        const input = entryInput();
        const text = input.value;
        if (!/^[=+-]/.test(text)) return;
        const at = input.selectionStart;
        const refs = references(text);
        const ref = refs.find((r) => at >= r.start && at <= r.end) || refs.filter((r) => r.end <= at).pop();
        if (!ref) return;
        const part = text.slice(ref.start, ref.end);
        const bang = part.lastIndexOf("!");
        const head = part.slice(0, bang + 1);
        const ends = part.slice(bang + 1).split(":").map((t) => /^(\$?)([A-Z]*)(\$?)([0-9]*)$/.exec(t));
        if (ends.some((m) => !m)) return;
        const [m] = ends;
        let state;
        if (m[2] && m[4]) {
            // Both: column and row; next as Excel goes round.
            const now = `${m[1] ? 1 : 0}${m[3] ? 1 : 0}`;
            state = { "00": [1, 1], "11": [0, 1], "01": [1, 0], "10": [0, 0] }[now];
        } else {
            const abs = m[2] ? !!m[1] : !!m[3];
            state = [abs ? 0 : 1, abs ? 0 : 1];
        }
        const body = ends.map((e) => `${e[2] && state[0] ? "$" : ""}${e[2]}${e[4] && state[1] ? "$" : ""}${e[4]}`).join(":");
        const next = head + body;
        input.setRangeText(next, ref.start, ref.end, "end");
        if (entry.point && entry.point.start === ref.start) entry.point.end = ref.start + next.length;
        if (entry.where === "cell") $("fx").textContent = input.value;
        placeEditor();
        entryRefs();
    }

    // ---- formulas typed: the functions (and names) that begin with what is typed, and the
    // arguments of the function the caret is in, as Excel shows them

    const FUNCTIONS = (data.functions || []).map(([name, sig, desc]) => ({ name, sig, desc }));
    const byName = new Map(FUNCTIONS.map((f) => [f.name, f]));
    const fnlist = el("div", "fnlist");
    const fnhint = el("div", "fnhint");
    fnlist.hidden = fnhint.hidden = true;
    document.body.append(fnlist, fnhint);
    // The list shown: { items: [{ name, fn, sig, desc }], at (the one chosen), start, end (of the word) }.
    let assist = null;
    function closeAssist() {
        assist = null;
        fnlist.hidden = fnhint.hidden = true;
    }
    // What the text before the caret is: outside strings, the word being typed and the calls
    // still open ([{ name, arg }], the innermost last).
    function scanFormula(before) {
        const calls = [];
        let inString = false;
        let word = "";
        let wordStart = 0;
        for (let i = 0; i < before.length; i++) {
            const ch = before[i];
            if (inString) {
                if (ch === '"') inString = false;
                continue;
            }
            if (ch === '"') {
                inString = true;
                word = "";
                continue;
            }
            if (/[\p{L}\p{N}_.]/u.test(ch)) {
                if (!word) wordStart = i;
                word += ch;
                continue;
            }
            if (ch === "(") calls.push({ name: word.toUpperCase(), arg: 0 });
            else if (ch === ")") calls.pop();
            else if (ch === "," && calls.length) calls[calls.length - 1].arg++;
            // ("{" "}" of array constants: commas there are not arguments.)
            else if (ch === "{") calls.push({ name: "{", arg: 0 });
            else if (ch === "}" && calls.length && calls[calls.length - 1].name === "{") calls.pop();
            word = "";
        }
        return { inString, word, wordStart, calls };
    }
    function updateAssist() {
        if (!entry || !view || !FUNCTIONS.length) return closeAssist();
        const input = entryInput();
        const text = input.value;
        if (!/^[=+-]/.test(text) || input.selectionStart !== input.selectionEnd) return closeAssist();
        const caret = input.selectionStart;
        const scan = scanFormula(text.slice(0, caret));
        const r = (entry.where === "bar" ? bar : ed).getBoundingClientRect();
        // The functions (then the names) that begin with the word, unless it is a reference
        // or a number, or follows a sheet's name.
        const w = scan.word;
        const prev = text[scan.wordStart - 1] || "";
        const items = [];
        if (!scan.inString && w && /^[\p{L}_]/u.test(w) && prev !== "!" && prev !== "$" && !/^[A-Z]{1,3}\d+$/i.test(w) && caret === scan.wordStart + w.length && !/^[\p{L}\p{N}_.(]/u.test(text[caret] || "")) {
            const up = w.toUpperCase();
            for (const f of FUNCTIONS) if (f.name.startsWith(up)) items.push({ name: f.name, fn: true, desc: f.desc });
            for (const n of namesOf(view.sheet)) if (n.name.toUpperCase().startsWith(up)) items.push({ name: n.name, fn: false, desc: n.ref });
        }
        if (items.length) {
            const was = assist && assist.items[assist.at] && assist.items[assist.at].name;
            const at = Math.max(0, items.findIndex((x) => x.name === was));
            assist = { items, at, start: scan.wordStart, end: caret };
            drawList(r);
        } else {
            assist = null;
            fnlist.hidden = true;
        }
        // The arguments of the innermost call to a function.
        const call = [...scan.calls].reverse().find((c) => c.name !== "{");
        const f = call && byName.get(call.name.replace(/^_XLFN\./, ""));
        if (f && !scan.inString) {
            fnhint.textContent = "";
            const m = /^([^(]*)\((.*)\)$/.exec(f.sig);
            const args = m ? splitArgs(m[2]) : [];
            fnhint.appendChild(document.createTextNode(`${m ? m[1] : f.name}(`));
            // Past the last named argument of a function that takes more: the last named one.
            const named = args.filter((a) => a !== "...");
            const current = call.arg < named.length ? call.arg : args.includes("...") ? named.length - 1 : -1;
            args.forEach((a, i) => {
                if (i) fnhint.appendChild(document.createTextNode(", "));
                const span = fnhint.appendChild(el("span", i === current ? "on" : null, a));
                span.title = a;
            });
            fnhint.appendChild(document.createTextNode(")"));
            fnhint.hidden = false;
            Object.assign(fnhint.style, { left: `${Math.max(0, r.left)}px`, top: `${r.bottom + 2}px` });
            if (!fnlist.hidden) fnlist.style.top = `${r.bottom + 4 + fnhint.offsetHeight}px`;
        } else {
            fnhint.hidden = true;
        }
    }
    // The arguments of a signature, split at the commas outside brackets.
    function splitArgs(list) {
        const out = [];
        let depth = 0;
        let cur = "";
        for (const ch of list) {
            if (ch === "[" || ch === "(") depth++;
            if (ch === "]" || ch === ")") depth--;
            if (ch === "," && depth === 0) {
                out.push(cur.trim());
                cur = "";
            } else cur += ch;
        }
        if (cur.trim()) out.push(cur.trim());
        return out;
    }
    function drawList(r) {
        fnlist.textContent = "";
        const shown = assist.items.slice(0, 200);
        shown.forEach((item, i) => {
            const row = fnlist.appendChild(el("div", i === assist.at ? "item on" : "item"));
            row.appendChild(el("span", item.fn ? "fx" : "nm", item.fn ? "fx" : "名"));
            row.appendChild(el("span", "name", item.name));
            row.addEventListener("mousedown", (e) => {
                e.preventDefault();
                assist.at = i;
                complete();
            });
        });
        const on = assist.items[assist.at];
        fnlist.appendChild(el("div", "desc", on ? on.desc : ""));
        fnlist.hidden = false;
        Object.assign(fnlist.style, { left: `${Math.max(0, r.left)}px`, top: `${r.bottom + (fnhint.hidden ? 2 : 4 + fnhint.offsetHeight)}px` });
        const row = fnlist.querySelector(".item.on");
        if (row) row.scrollIntoView({ block: "nearest" });
    }
    // Tab (or a click): the function chosen, with its parenthesis, in place of the word.
    function complete() {
        const a = assist;
        if (!a || !entry) return;
        const item = a.items[a.at];
        const input = entryInput();
        const text = item.fn ? `${item.name}(` : item.name;
        input.setRangeText(text, a.start, a.end, "end");
        input.dispatchEvent(new Event("input"));
    }
    // The keys of the list while it is open: true when one was taken.
    function assistKey(e) {
        if (!assist) return false;
        if (e.key === "ArrowDown" || e.key === "ArrowUp") {
            assist.at = Math.max(0, Math.min(assist.items.length - 1, assist.at + (e.key === "ArrowDown" ? 1 : -1)));
            drawList((entry.where === "bar" ? bar : ed).getBoundingClientRect());
            return true;
        }
        if (e.key === "Tab") {
            complete();
            return true;
        }
        if (e.key === "Escape") {
            assist = null;
            fnlist.hidden = true;
            return true;
        }
        return false;
    }

    // ---- edits: posted to the editor, which writes the workbook and draws the page again

    function postEdit(ops, keepClip) {
        if (!editable || busy) return false;
        if (!keepClip) setClip(null);
        setBusy(true);
        vscode.postMessage({ type: "edit", ops });
        return true;
    }
    function postHistory(type) {
        if (!editable || busy) return;
        setClip(null);
        setBusy(true);
        vscode.postMessage({ type });
    }
    function setBusy(on) {
        busy = on;
        document.body.classList.toggle("busy", on);
    }
    function setClip(c) {
        clip = c;
        remember({ clip });
        drawBoxes();
    }

    // What the editor could not do (Excel would show a dialog).
    function showError(message) {
        let box = $("error");
        if (!box) {
            box = el("div");
            box.id = "error";
            document.body.insertBefore(box, $("tabs"));
        }
        box.textContent = "";
        box.classList.remove("hint");
        box.appendChild(el("span", null, message));
        const close = box.appendChild(el("button", null, "×"));
        close.type = "button";
        close.title = "閉じる";
        close.addEventListener("click", () => {
            box.hidden = true;
            focusGrid();
        });
        box.hidden = false;
    }
    window.addEventListener("message", (e) => {
        const m = e.data;
        if (m && m.type === "events") {
            if (m.on) enableEvents();
        } else if (m && m.type === "error") {
            setBusy(false);
            remember({ pickObj: null });
            showError(String(m.message || "変更できませんでした。"));
            focusGrid();
        }
    });

    // ---- copy, cut and paste

    // The cells of the selection as the clipboard takes them: tab-separated, as they are shown.
    function tsvOf(s) {
        const v = view;
        const quote = (t) => (/["\t\n\r]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
        const r2 = Math.min(s.r2, Math.max(v.sheet.rows - 1, s.r1)),
            c2 = Math.min(s.c2, Math.max(v.sheet.cols - 1, s.c1));
        const lines = [];
        for (let r = s.r1; r <= r2; r++) {
            const line = [];
            for (let c = s.c1; c <= c2; c++) {
                const o = ownerOf(r, c);
                const cell = !o || (o[0] === r && o[1] === c) ? v.cells.get(r * KEY + c) : null;
                line.push(quote((cell && cell.t) || ""));
            }
            lines.push(line.join("\t"));
        }
        return lines.join("\r\n") + "\r\n";
    }
    // Clipboard text as rows of cells (quoted cells hold line breaks and tabs, as Excel writes them).
    function parseTsv(text) {
        const rows = [[]];
        let cell = "";
        let i = 0;
        const n = text.length;
        while (i < n) {
            if (cell === "" && text[i] === '"') {
                // A quoted cell, when its closing quote ends the cell.
                let j = i + 1;
                let s = "";
                let closed = false;
                while (j < n) {
                    if (text[j] === '"') {
                        if (text[j + 1] === '"') {
                            s += '"';
                            j += 2;
                            continue;
                        }
                        closed = true;
                        j++;
                        break;
                    }
                    s += text[j++];
                }
                if (closed && (j >= n || text[j] === "\t" || text[j] === "\n" || text[j] === "\r")) {
                    cell = s.replace(/\r\n/g, "\n");
                    i = j;
                    if (i >= n) break;
                }
            }
            const ch = text[i];
            if (ch === "\t") {
                rows[rows.length - 1].push(cell);
                cell = "";
            } else if (ch === "\n" || ch === "\r") {
                if (ch === "\r" && text[i + 1] === "\n") i++;
                rows[rows.length - 1].push(cell);
                cell = "";
                rows.push([]);
            } else {
                cell += ch;
            }
            i++;
        }
        if (cell !== "" || rows[rows.length - 1].length) rows[rows.length - 1].push(cell);
        else rows.pop();
        const wide = Math.max(0, ...rows.map((r) => r.length));
        for (const r of rows) while (r.length < wide) r.push("");
        return rows;
    }
    const sameText = (a, b) => a.replace(/\r\n?/g, "\n").replace(/\n$/, "") === b.replace(/\r\n?/g, "\n").replace(/\n$/, "");

    function copySel(cut) {
        if (!sel || !view) return "";
        const text = tsvOf(sel);
        setClip({ sheet: view.sheet.name, kind: sel.kind, r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, cut: !!cut && editable, text });
        return text;
    }

    function pasteText(text) {
        if (!editable || busy || !sel || !view || !oneArea()) return;
        const sheet = view.sheet.name;
        const one = single(sel);
        if (clip && sameText(text, clip.text)) {
            // The cells copied in this grid: formulas and formats go with them.
            const [h, w] = [clip.r2 - clip.r1 + 1, clip.c2 - clip.c1 + 1];
            const from = rangeText(clip);
            const fromSheet = clip.sheet;
            const [r, c] = [sel.r1, sel.c1];
            if (clip.kind === "all") return;
            if (clip.cut) {
                const to = clip.kind === "rows" ? address(r, 0) : clip.kind === "cols" ? address(0, c) : address(r, c);
                setSel(clip.kind === "rows" ? { kind: "rows", nr: r, nc: 0, fr: r + h - 1, fc: 0 } : clip.kind === "cols" ? { kind: "cols", nr: 0, nc: c, fr: 0, fc: c + w - 1 } : areaSel(r, c, r + h - 1, c + w - 1), false);
                postEdit([{ op: "move", sheet, from, fromSheet, to }]);
                return;
            }
            let to = rangeText(sel);
            if (clip.kind === "rows" && (one || (sel.kind === "rows" && sel.r1 === sel.r2))) {
                to = address(r, 0);
                setSel({ kind: "rows", nr: r, nc: 0, fr: r + h - 1, fc: 0 }, false);
            } else if (clip.kind === "cols" && (one || (sel.kind === "cols" && sel.c1 === sel.c2))) {
                to = address(0, c);
                setSel({ kind: "cols", nr: 0, nc: c, fr: 0, fc: c + w - 1 }, false);
            } else if (one) {
                to = address(r, c);
                setSel(areaSel(r, c, r + h - 1, c + w - 1), false);
            }
            postEdit([{ op: "copy", sheet, from, fromSheet, to }], true);
            return;
        }
        const rows = parseTsv(text);
        if (!rows.length || !rows[0].length) return;
        if (rows.length === 1 && rows[0].length === 1) {
            const range = rangeText(sel);
            postEdit([rows[0][0] === "" ? { op: "clear", sheet, range } : { op: "set", sheet, range, text: rows[0][0] }]);
            return;
        }
        const [r, c] = [sel.r1, sel.c1];
        setSel(areaSel(r, c, r + rows.length - 1, c + rows[0].length - 1), false);
        postEdit([{ op: "values", sheet, at: address(r, c), rows }]);
    }

    // "形式を選択して貼り付け": what of the cells copied in this grid goes where the selection is.
    const PASTES = [
        ["all", "すべて"],
        ["formulas", "数式"],
        ["values", "値"],
        ["formats", "書式"],
        ["allExceptBorders", "罫線を除くすべて"],
        ["columnWidths", "列幅"],
        ["formulasAndNumberFormats", "数式と数値の書式"],
        ["valuesAndNumberFormats", "値と数値の書式"],
    ];
    const OPERATIONS = [
        ["none", "しない"],
        ["add", "加算"],
        ["subtract", "減算"],
        ["multiply", "乗算"],
        ["divide", "除算"],
    ];
    function pasteSpecial(how) {
        if (!editable || busy || !sel || !view || !clip || !oneArea()) return;
        if (clip.cut) {
            showError("切り取ったセルは、形式を選択して貼り付けできません。");
            return;
        }
        if (clip.kind === "all") return;
        const sheet = view.sheet.name;
        const [r, c] = [sel.r1, sel.c1];
        const [h, w] = how.transpose ? [clip.c2 - clip.c1 + 1, clip.r2 - clip.r1 + 1] : [clip.r2 - clip.r1 + 1, clip.c2 - clip.c1 + 1];
        let to = rangeText(sel);
        if (single(sel) || how.transpose) {
            to = address(r, c);
            if (clip.kind === "cells") setSel(areaSel(r, c, r + h - 1, c + w - 1), false);
        }
        postEdit([Object.assign({ op: "pasteSpecial", sheet, from: rangeText(clip), fromSheet: clip.sheet, to }, how)], true);
    }
    function askPasteSpecial() {
        if (!editable || busy || !sel || !view || !oneArea()) return;
        if (!clip) {
            showError("形式を選択して貼り付けは、このグリッドでコピーしたセルに対して行えます。");
            return;
        }
        if (clip.cut) {
            showError("切り取ったセルは、形式を選択して貼り付けできません。");
            return;
        }
        if (entry) commit(null);
        closeMenu();
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg pastespecial"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "形式を選択して貼り付け"));
        const group = (title, name, options) => {
            const box = form.appendChild(el("fieldset"));
            box.appendChild(el("legend", null, title));
            options.forEach(([value, label], i) => {
                const line = box.appendChild(el("label"));
                const radio = line.appendChild(el("input"));
                radio.type = "radio";
                radio.name = name;
                radio.value = value;
                radio.checked = i === 0;
                line.appendChild(document.createTextNode(label));
            });
        };
        group("貼り付け", "paste", PASTES);
        group("演算", "operation", OPERATIONS);
        const checks = form.appendChild(el("div", "checks"));
        const check = (name, label) => {
            const line = checks.appendChild(el("label"));
            const box = line.appendChild(el("input"));
            box.type = "checkbox";
            box.name = name;
            line.appendChild(document.createTextNode(label));
        };
        check("skipBlanks", "空白セルを無視する");
        check("transpose", "行列を入れ替える");
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) closeDialog();
        });
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const f = new FormData(form);
            const how = { paste: f.get("paste"), operation: f.get("operation"), skipBlanks: f.has("skipBlanks"), transpose: f.has("transpose") };
            closeDialog();
            pasteSpecial(how);
        });
        form.querySelector("input:checked").focus();
    }

    document.addEventListener("copy", (e) => {
        if (entry || !onGrid() || !sel || selected) return;
        if (!oneArea()) {
            e.preventDefault();
            return;
        }
        e.clipboardData.setData("text/plain", copySel(false));
        e.preventDefault();
    });
    document.addEventListener("cut", (e) => {
        if (entry || !onGrid() || !sel || selected) return;
        if (!oneArea()) {
            e.preventDefault();
            return;
        }
        e.clipboardData.setData("text/plain", copySel(true));
        e.preventDefault();
    });
    document.addEventListener("paste", (e) => {
        if (entry || !onGrid()) return;
        e.preventDefault();
        if (pastedPicture(e)) return;
        pasteText(e.clipboardData.getData("text/plain"));
    });
    // From the menu: through the clipboard where the page may use it.
    function menuCopy(cut) {
        if (!oneArea()) return;
        const text = copySel(cut);
        const write = navigator.clipboard && navigator.clipboard.writeText ? navigator.clipboard.writeText(text) : Promise.reject();
        write.catch(() => {
            // The copy event of the page takes it.
            ed.value = " ";
            ed.select();
            try {
                document.execCommand(cut ? "cut" : "copy");
            } catch (x) {
                // Pasting in this grid still works.
            }
            ed.value = "";
        });
    }
    function menuPaste() {
        const read = navigator.clipboard && navigator.clipboard.readText ? navigator.clipboard.readText() : Promise.reject();
        read.then(pasteText, () => {
            if (clip) pasteText(clip.text);
            else showError("クリップボードを読み取れませんでした。Ctrl+V で貼り付けてください。");
        });
    }

    // ---- insert, delete, clear, fill

    function clearSel() {
        if (sel && view) postEdit(allAreas().map((a) => ({ op: "clear", sheet: view.sheet.name, range: rangeText(a) })));
    }

    // `how`: "shift" (the cells move right or down / left or up), "rows", "cols".
    function structure(insert, how, across) {
        if (!sel || !view || sel.kind === "all" || !oneArea()) return;
        const sheet = view.sheet.name;
        const rows = `${sel.r1 + 1}:${sel.r2 + 1}`;
        const cols = `${colName(sel.c1)}:${colName(sel.c2)}`;
        const op = insert ? "insert" : "delete";
        const range = how === "rows" ? rows : how === "cols" ? cols : rangeText(sel);
        const shift = how === "rows" ? (insert ? "down" : "up") : how === "cols" ? (insert ? "right" : "left") : insert ? (across ? "right" : "down") : across ? "left" : "up";
        // The rows and columns stay selected, as Excel leaves them.
        if (how === "rows" && sel.kind === "cells") setSel({ kind: "rows", nr: sel.r1, nc: 0, fr: sel.r2, fc: 0 }, false);
        if (how === "cols" && sel.kind === "cells") setSel({ kind: "cols", nr: 0, nc: sel.c1, fr: 0, fc: sel.c2 }, false);
        postEdit([{ op, sheet, range, shift }]);
    }
    function insertOrDelete(insert) {
        if (!editable || busy || !sel || sel.kind === "all" || !oneArea()) return;
        if (sel.kind === "rows") structure(insert, "rows");
        else if (sel.kind === "cols") structure(insert, "cols");
        else askShift(insert);
    }

    // Excel's dialog: which way the cells move.
    let dialog = null;
    function closeDialog() {
        if (dialog) dialog.remove();
        dialog = null;
        focusGrid();
    }
    function askShift(insert) {
        closeMenu();
        closeDialog();
        const options = insert
            ? [["right", "右方向にシフト"], ["down", "下方向にシフト"], ["rows", "行全体"], ["cols", "列全体"]]
            : [["left", "左方向にシフト"], ["up", "上方向にシフト"], ["rows", "行全体"], ["cols", "列全体"]];
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, insert ? "挿入" : "削除"));
        options.forEach(([value, label], i) => {
            const line = form.appendChild(el("label"));
            const radio = line.appendChild(el("input"));
            radio.type = "radio";
            radio.name = "shift";
            radio.value = value;
            radio.checked = i === 1;
            line.appendChild(document.createTextNode(label));
        });
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) closeDialog();
        });
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const value = new FormData(form).get("shift");
            closeDialog();
            if (value === "rows" || value === "cols") structure(insert, value);
            else structure(insert, "shift", value === "right" || value === "left");
        });
        form.querySelector("input:checked").focus();
    }

    // Ctrl+D / Ctrl+R: the first row (column) of the selection into the rest; of one row
    // (column), the cells above (to the left) into it.
    function fill(dir) {
        if (!sel || !view || sel.kind === "all" || !oneArea()) return;
        let { r1, c1, r2, c2 } = sel;
        const s = { kind: sel.kind, r1, c1, r2, c2 };
        if (dir === "down" && r1 === r2) {
            if (r1 === 0 || sel.kind === "cols") return;
            s.r1 = r1 - 1;
        }
        if (dir === "right" && c1 === c2) {
            if (c1 === 0 || sel.kind === "rows") return;
            s.c1 = c1 - 1;
        }
        const range = s.kind === "cells" ? `${address(s.r1, s.c1)}:${address(s.r2, s.c2)}` : rangeText(s);
        postEdit([{ op: "fill", sheet: view.sheet.name, range, dir }]);
    }

    // ---- 挿入: tables, hyperlinks, notes and symbols, as Excel puts them on a sheet

    const noteAt = (r, c) => (view ? (view.sheet.notes || []).find((n) => n.r === r && n.c === c) : null);
    const linkAt = (r, c) => (view ? (view.sheet.links || []).find((l) => l.r === r && l.c === c) : null);
    const cellOf = (r, c) => (view ? view.cells.get(r * KEY + c) : null);
    const filledAt = (r, c) => {
        const x = cellOf(r, c);
        return !!x && ((x.t !== undefined && x.t !== "") || !!x.f);
    };
    // The data around a cell (Excel's current region): reached through cells with something in
    // them, across the diagonals too.
    function regionAround(r, c) {
        let [r1, c1, r2, c2] = [r, c, r, c];
        const any = (a, b, x, y) => {
            for (let i = Math.max(0, a); i <= x; i++) for (let j = Math.max(0, b); j <= y; j++) if (filledAt(i, j)) return true;
            return false;
        };
        for (let more = true; more; ) {
            more = false;
            if (r1 > 0 && any(r1 - 1, c1 - 1, r1 - 1, c2 + 1)) (r1--, (more = true));
            if (r2 < MAXR - 1 && any(r2 + 1, c1 - 1, r2 + 1, c2 + 1)) (r2++, (more = true));
            if (c1 > 0 && any(r1 - 1, c1 - 1, r2 + 1, c1 - 1)) (c1--, (more = true));
            if (c2 < MAXC - 1 && any(r1 - 1, c2 + 1, r2 + 1, c2 + 1)) (c2++, (more = true));
        }
        return { r1, c1, r2, c2 };
    }
    // A text typed in as it is (a number, a date or a formula would be taken for one).
    const asText = (text) => (/^[=+\-@']/.test(text) || (text.trim() !== "" && !isNaN(+text.replace(/[,¥$%\s]/g, ""))) || /^\d+[/-]\d+/.test(text) ? `'${text}` : text);

    // A dialog of the page: its form, its buttons (OK and キャンセル), closed by a click beside it.
    function openDialog(title, cls) {
        closeMenu();
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", `dlg${cls ? ` ${cls}` : ""}`));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, title));
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) closeDialog();
        });
        return form;
    }
    function dialogButtons(form, okText) {
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", okText || "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        return buttons;
    }
    function field(form, label, type) {
        const line = form.appendChild(el("label"));
        line.appendChild(el("span", "lbl", label));
        const input = line.appendChild(el(type === "textarea" ? "textarea" : "input"));
        if (type !== "textarea") input.type = type || "text";
        input.spellcheck = false;
        return input;
    }
    function check(form, label, on) {
        const line = form.appendChild(el("label"));
        const box = line.appendChild(el("input"));
        box.type = "checkbox";
        box.checked = on;
        line.appendChild(document.createTextNode(label));
        return box;
    }
    const absText = (s) => (s.r1 === s.r2 && s.c1 === s.c2 ? `=$${colName(s.c1)}$${s.r1 + 1}` : `=$${colName(s.c1)}$${s.r1 + 1}:$${colName(s.c2)}$${s.r2 + 1}`);

    // Ctrl+T: "テーブルの作成" — the data (the selection, or the cells around the active cell),
    // and whether its first row is the headings.
    function askTable() {
        if (!editable || busy || !sel || !view || sel.kind === "all" || !oneArea()) return;
        if (entry) commit(null);
        // One cell stands for the sheet's filter it is in, else for the data around it.
        const filter = view.sheet.filter ? parseRef(view.sheet.filter, view.sheet) : null;
        const inFilter = filter && sel.ar >= filter.r1 && sel.ar <= filter.r2 && sel.ac >= filter.c1 && sel.ac <= filter.c2;
        const s = sel.kind !== "cells" || !single(sel) ? sel : inFilter ? filter : regionAround(sel.ar, sel.ac);
        const form = openDialog("テーブルの作成", "tabledlg");
        const input = field(form, "テーブルに変換するデータ範囲を指定してください:");
        input.value = sel.kind === "cells" ? absText(s) : `=${rangeText(sel).replace(/([A-Z]+|\d+)/g, "$$$1")}`;
        // Headings guessed as Excel guesses them: texts over a row that is not all texts.
        let guess = false;
        if (sel.kind === "cells" && s.r2 > s.r1) {
            guess = true;
            for (let c = s.c1; c <= s.c2; c++) {
                const x = cellOf(s.r1, c);
                if (!x || x.k !== "string" || x.f) guess = false;
            }
        }
        const head = check(form, "先頭行をテーブルの見出しとして使用する", guess);
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const ref = parseRef(input.value, view.sheet);
            if (!ref || ref.sheet !== view.sheet) {
                input.setCustomValidity("参照が正しくありません");
                input.reportValidity();
                return;
            }
            closeDialog();
            postEdit([{ op: "table", sheet: view.sheet.name, range: rangeText(ref), header: head.checked }]);
        });
        input.addEventListener("input", () => input.setCustomValidity(""));
        input.focus();
        input.select();
    }

    // Ctrl+K: "ハイパーリンクの挿入" (of a cell that has one, "ハイパーリンクの編集"): a web
    // page or a file, or a place in the workbook; what the cell shows, and the tip.
    function askLink() {
        if (!editable || busy || !sel || !view || sel.kind !== "cells" || !oneArea()) return;
        if (entry) commit(null);
        const sheet = view.sheet;
        const now = linkAt(sel.ar, sel.ac);
        const cell = cellOf(sel.ar, sel.ac) || {};
        const one = single(sel);
        const form = openDialog(now ? "ハイパーリンクの編集" : "ハイパーリンクの挿入", "linkdlg");
        const kinds = form.appendChild(el("div", "kinds"));
        const kind = (value, label) => {
            const line = kinds.appendChild(el("label"));
            const radio = line.appendChild(el("input"));
            radio.type = "radio";
            radio.name = "kind";
            radio.value = value;
            line.appendChild(document.createTextNode(label));
            return radio;
        };
        const web = kind("url", "ファイル、Web ページ");
        const place = kind("to", "このドキュメント内");
        const shown = field(form, "表示文字列:");
        shown.value = one ? (cell.f ? "" : cell.t || "") : "<<選択範囲内の文字列>>";
        shown.disabled = !one || !!cell.f;
        const addr = field(form, "アドレス:");
        addr.placeholder = "https://";
        const where = form.appendChild(el("div", "where"));
        const sheets = where.appendChild(el("label"));
        sheets.appendChild(el("span", "lbl", "場所:"));
        const list = sheets.appendChild(el("select"));
        for (const s of data.sheets) list.appendChild(el("option", null, s.name)).value = `s:${s.name}`;
        for (const n of namesOf(sheet)) list.appendChild(el("option", null, `名前: ${n.name}`)).value = `n:${n.name}`;
        const refLine = where.appendChild(el("label"));
        refLine.appendChild(el("span", "lbl", "セル参照:"));
        const ref = refLine.appendChild(el("input"));
        ref.type = "text";
        ref.spellcheck = false;
        ref.value = "A1";
        const tip = field(form, "ヒント:");
        if (now) {
            tip.value = now.tip || "";
            if (now.to !== undefined) {
                place.checked = true;
                const t = now.to;
                const bang = t.lastIndexOf("!");
                const named = namesOf(sheet).find((n) => n.name.toLowerCase() === t.toLowerCase());
                if (named) list.value = `n:${named.name}`;
                else if (bang >= 0) {
                    list.value = `s:${t.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'")}`;
                    ref.value = t.slice(bang + 1);
                }
            } else {
                web.checked = true;
                addr.value = now.url || "";
            }
        } else {
            web.checked = true;
            list.value = `s:${sheet.name}`;
        }
        const showKind = () => {
            addr.closest("label").hidden = !web.checked;
            where.hidden = web.checked;
            ref.disabled = list.value.startsWith("n:");
        };
        web.addEventListener("change", showKind);
        place.addEventListener("change", showKind);
        list.addEventListener("change", showKind);
        showKind();
        const buttons = dialogButtons(form);
        if (now) {
            const off = buttons.insertBefore(el("button", null, "リンクの解除"), buttons.firstChild);
            off.type = "button";
            off.addEventListener("click", () => {
                closeDialog();
                removeLinks();
            });
        }
        const range = rangeText(sel);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const op = { op: "link", sheet: sheet.name, range };
            let target;
            if (web.checked) {
                target = addr.value.trim();
                if (!target) {
                    addr.setCustomValidity("アドレスを入力してください");
                    addr.reportValidity();
                    return;
                }
                // A web address typed without its scheme, as Excel takes it.
                if (/^www\./i.test(target)) target = `http://${target}`;
                op.url = target;
            } else if (list.value.startsWith("n:")) {
                target = list.value.slice(2);
                op.to = target;
            } else {
                const name = list.value.slice(2);
                const at = ref.value.trim().replace(/\$/g, "").toUpperCase() || "A1";
                const on = sheetByName(name);
                const r = on && parseRef(at, on);
                if (!r) {
                    ref.setCustomValidity("セル参照が正しくありません");
                    ref.reportValidity();
                    return;
                }
                target = `${/^[A-Za-z_　-鿿][\w.　-鿿]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`}!${at}`;
                op.to = target;
            }
            if (tip.value.trim()) op.tip = tip.value.trim();
            closeDialog();
            const ops = [];
            if (one && !shown.disabled) {
                const text = shown.value || (cell.t ? "" : target);
                if (text && text !== (cell.t || "")) ops.push({ op: "set", sheet: sheet.name, range: address(sel.ar, sel.ac), text: asText(text) });
            }
            // A new link is shown as Excel's Hyperlink style shows it: blue, underlined.
            if (!now) ops.push({ op: "format", sheet: sheet.name, range, set: { color: "#0563C1", underline: true } });
            ops.push(op);
            postEdit(ops);
        });
        addr.addEventListener("input", () => addr.setCustomValidity(""));
        ref.addEventListener("input", () => ref.setCustomValidity(""));
        (now && place.checked ? ref : web.checked ? addr : shown).focus();
    }
    // "ハイパーリンクの削除": the links of the selection go, and their blue underline.
    function removeLinks() {
        if (!editable || busy || !sel || !view) return;
        const sheet = view.sheet.name;
        postEdit(allAreas().flatMap((a) => [{ op: "link", sheet, range: rangeText(a), off: true }, { op: "format", sheet, range: areaText(a), set: { color: null, underline: false } }]));
    }
    const linksIn = () => allAreas().some((a) => (view.sheet.links || []).some((l) => l.r >= a.r1 && l.r <= a.r2 && l.c >= a.c1 && l.c <= a.c2));

    // Shift+F2: "新しいメモ" (of a cell that has one, "メモの編集").
    function askNote() {
        if (!editable || busy || !sel || !view || sel.kind !== "cells") return;
        if (entry) commit(null);
        const [r, c] = [sel.ar, sel.ac];
        const now = noteAt(r, c);
        const form = openDialog(`${now ? "メモの編集" : "新しいメモ"}（${address(r, c)}）`, "notedlg");
        const text = form.appendChild(el("textarea"));
        text.rows = 6;
        text.value = now ? now.text : "";
        const buttons = dialogButtons(form);
        if (now) {
            const off = buttons.insertBefore(el("button", null, "削除"), buttons.firstChild);
            off.type = "button";
            off.addEventListener("click", () => {
                closeDialog();
                removeNote();
            });
        }
        const sheet = view.sheet.name;
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
            if (now && text.value === now.text) return;
            postEdit([{ op: "note", sheet, at: address(r, c), text: text.value }]);
        });
        // Ctrl+Enter enters the note (Enter starts a new line in it).
        text.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                form.requestSubmit();
            }
        });
        text.focus();
        text.setSelectionRange(text.value.length, text.value.length);
    }
    // "メモの削除": the notes of the cells selected.
    function removeNote() {
        if (!editable || busy || !sel || !view) return;
        const sheet = view.sheet.name;
        const ops = (view.sheet.notes || []).filter((n) => allAreas().some((a) => n.r >= a.r1 && n.r <= a.r2 && n.c >= a.c1 && n.c <= a.c2)).map((n) => ({ op: "note", sheet, at: address(n.r, n.c), off: true }));
        if (ops.length) postEdit(ops);
    }
    const notesIn = () => allAreas().some((a) => (view.sheet.notes || []).some((n) => n.r >= a.r1 && n.r <= a.r2 && n.c >= a.c1 && n.c <= a.c2));

    // ---- shapes and pictures: put on the sheet (挿入 > 図), moved, sized, given a text and
    // deleted with the mouse and the keys as in Excel (`shape`, `picture` and `object` edits)

    // The shapes of 挿入 > 図形: [kind, name, icon (16 × 16)], by group.
    const SHAPE_GROUPS = [
        ["線", [
            ["line", "直線", '<path d="M2.5 13.5 13.5 2.5"/>'],
            ["arrow", "線矢印", '<path d="M2.5 13.5 13 3"/><path d="M8.5 3H13v4.5"/>'],
        ]],
        ["四角形", [
            ["rect", "正方形/長方形", '<rect x="2" y="3.5" width="12" height="9"/>'],
            ["roundRect", "四角形: 角を丸くする", '<rect x="2" y="3.5" width="12" height="9" rx="2.5"/>'],
        ]],
        ["基本図形", [
            ["textBox", "テキスト ボックス", '<rect x="2" y="2.5" width="12" height="11"/><path d="M5 5.5h6M8 5.5v6"/>'],
            ["ellipse", "楕円", '<ellipse cx="8" cy="8" rx="6" ry="5"/>'],
            ["triangle", "二等辺三角形", '<path d="M8 2.5 14 13.5H2z"/>'],
            ["rtTriangle", "直角三角形", '<path d="M2.5 2.5v11h11z"/>'],
            ["diamond", "ひし形", '<path d="M8 1.5 14.5 8 8 14.5 1.5 8z"/>'],
            ["parallelogram", "平行四辺形", '<path d="M5 3.5h9l-3 9H2z"/>'],
            ["trapezoid", "台形", '<path d="M5 3.5h6l3 9H2z"/>'],
            ["pentagon", "五角形", '<path d="M8 2 14 6.5 11.7 13.5H4.3L2 6.5z"/>'],
            ["hexagon", "六角形", '<path d="M5 3h6l3 5-3 5H5L2 8z"/>'],
            ["octagon", "八角形", '<path d="M5.5 2h5L14 5.5v5L10.5 14h-5L2 10.5v-5z"/>'],
            ["heart", "ハート", '<path d="M8 13.5S2 9.8 2 5.8A3 3 0 0 1 8 4.5a3 3 0 0 1 6 1.3c0 4-6 7.7-6 7.7z"/>'],
            ["smileyFace", "スマイル", '<circle cx="8" cy="8" r="6"/><path d="M5.5 9.5c1.4 1.6 3.6 1.6 5 0"/><path d="M6 6h0M10 6h0" stroke-width="1.8"/>'],
        ]],
        ["ブロック矢印", [
            ["rightArrow", "矢印: 右", '<path d="M2 5.5h7V2.5l5 5.5-5 5.5v-3H2z"/>'],
            ["leftArrow", "矢印: 左", '<path d="M14 5.5H7V2.5L2 8l5 5.5v-3h7z"/>'],
            ["upArrow", "矢印: 上", '<path d="M5.5 14V7H2.5L8 2l5.5 5h-3v7z"/>'],
            ["downArrow", "矢印: 下", '<path d="M5.5 2v7H2.5L8 14l5.5-5h-3V2z"/>'],
        ]],
        ["星", [["star5", "星: 5 pt", '<path d="m8 1.5 1.9 4.2 4.6.4-3.5 3 1 4.5L8 11.3l-4 2.3 1-4.5-3.5-3 4.6-.4z"/>']]],
        ["吹き出し", [
            ["wedgeRectCallout", "吹き出し: 四角形", '<path d="M2 2.5h12v8H8l-3.5 3.5.5-3.5H2z"/>'],
            ["wedgeRoundRectCallout", "吹き出し: 角を丸めた四角形", '<path d="M4 2.5h8a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2H8l-3.5 3.5.5-3.5H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2z"/>'],
            ["wedgeEllipseCallout", "吹き出し: 円形", '<path d="M5.2 11.4A6 4.5 0 1 1 7 11.9L3.5 14z"/>'],
        ]],
    ];
    const LINES = new Set(["line", "arrow"]);
    // The size of a shape put with a click (px): a line goes right.
    const clickSize = (kind) => (LINES.has(kind) ? [96, 0] : kind === "textBox" ? [144, 36] : [96, 96]);
    // The shape being drawn: { kind, start: [x, y], ghost } (with the mouse down).
    let drawingShape = null;
    // An object being moved or sized: { obj, how: "move" | "n" | "ne" | ..., x, y (where the
    // mouse went down), box (the object's: x, y, w, h), ghost, now }.
    let objDrag = null;
    // A click that ends drawing or dragging is not one on the object or the cell under it.
    let swallowClick = false;
    // The objects that can be moved, sized and deleted here: not the controls (they are drawn
    // from the sheet's controls and their VML shapes too).
    // A form control too, by its shape (not an ActiveX control, nor an embedded object, which
    // have no macro).
    const isForm = (obj) => !!obj && obj.kind === "control" && obj.shapeId !== undefined && !!obj.onAction && obj.activex === undefined;
    const movable = (obj) => editable && !!obj && (isForm(obj) || (!!obj.name && obj.kind !== "control"));
    // How an edit names an object: a form control by its shape.
    const target = (obj) => (isForm(obj) ? { shape: obj.shapeId } : { name: obj.name });
    // What tells an object from the others after the page is drawn again.
    const keyOf = (obj) => (isForm(obj) ? `#${obj.shapeId}` : obj.name || "");

    // The form controls of 開発 > 挿入: [kind, name, icon (16 × 16)].
    const FORM_CONTROLS = [
        ["button", "ボタン", '<rect x="1.5" y="4" width="13" height="8" rx="1"/><path d="M4 10.5h8" stroke-opacity=".4"/>'],
        ["comboBox", "コンボ ボックス", '<rect x="1.5" y="5" width="13" height="6"/><path d="M10.5 5v6M11.8 7.3l.7.9.7-.9"/>'],
        ["checkBox", "チェック ボックス", '<rect x="2.5" y="4.5" width="7" height="7"/><path d="m3.8 8 1.7 1.7 2.7-3.4"/><path d="M11 8h3.5" stroke-opacity=".5"/>'],
        ["spinner", "スピン ボタン", '<rect x="5" y="1.5" width="6" height="13"/><path d="M5 8h6M6.8 5.3 8 4l1.2 1.3M6.8 10.7 8 12l1.2-1.3"/>'],
        ["listBox", "リスト ボックス", '<rect x="1.5" y="2.5" width="13" height="11"/><path d="M3.5 5.5h6M3.5 8h6M3.5 10.5h6M11 2.5v11"/>'],
        ["optionButton", "オプション ボタン", '<circle cx="6" cy="8" r="3.5"/><circle cx="6" cy="8" r="1.3" fill="currentColor"/><path d="M11 8h3.5" stroke-opacity=".5"/>'],
        ["groupBox", "グループ ボックス", '<path d="M5 3.5H1.5v10h13v-10H11"/><path d="M6 3.5h4" stroke-opacity=".5"/>'],
        ["label", "ラベル", '<text x="2" y="12" font-size="10" stroke="none" fill="currentColor" font-family="sans-serif">Aa</text>'],
        ["scrollBar", "スクロール バー", '<rect x="1.5" y="5" width="13" height="6"/><path d="M4 5v6M12 5v6M6.5 7h3v2h-3z"/>'],
    ];
    const isControlKind = (kind) => kind.startsWith("control:");

    // A point of the sheet (px at 100% from the top left of the grid, as the objects are placed).
    function sheetPoint(e) {
        const r = view.wrap.getBoundingClientRect();
        const z = Z();
        return [(e.clientX - r.left) / z, (e.clientY - r.top) / z];
    }
    // The drawn column (row) a point is in: the last whose start is not after it.
    function lineAt(list, start, x) {
        let [lo, hi, best] = [0, list.length - 1, list.length ? list[0] : 0];
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (start(list[mid]) <= x) {
                best = list[mid];
                lo = mid + 1;
            } else hi = mid - 1;
        }
        return best;
    }
    // Where a point of the sheet is for an edit: the cell, and the pixels right of and below its
    // top left corner (the columns are drawn twice as wide while formulas are shown).
    function anchorOf(x, y) {
        const v = view;
        const c = lineAt(v.cols, (i) => columnLeft(v, i), x);
        const r = lineAt(v.rows.map((row) => row[0]), (i) => rowTop(v, i), y);
        const round = (n) => Math.round(n * 100) / 100;
        return { at: address(r, c), dx: round(Math.max(0, x - columnLeft(v, c)) / (formulas ? 2 : 1)), dy: round(Math.max(0, y - rowTop(v, r))) };
    }
    // After the page is drawn again: the object the edit made or changed is selected.
    // `key`: the object's (`keyOf`), or null for the one the edit adds.
    function pickAfter(key) {
        if (!view) return;
        remember({ pickObj: { sheet: view.sheet.name, key: key || null, before: key ? null : (view.sheet.objects || []).map(keyOf) } });
    }
    function pickObject(v) {
        const p = saved.pickObj;
        if (!p) return;
        remember({ pickObj: null });
        if (p.sheet !== v.sheet.name) return;
        const objects = v.sheet.objects || [];
        const obj = p.key ? objects.find((o) => keyOf(o) === p.key) : objects.find((o) => keyOf(o) && !(p.before || []).includes(keyOf(o)));
        if (obj && obj.box) selectObject(obj.box, obj);
    }

    function startDrawing(kind) {
        if (!editable || busy || !view) return;
        if (entry) commit(null);
        drawingShape = { kind };
        document.body.classList.add("drawing");
        showError(`${shapeName(kind)}: シートをドラッグして描きます（クリックすると既定の大きさ。Esc で取り消し）`);
        $("error").classList.add("hint");
    }
    function stopDrawing() {
        if (drawingShape && drawingShape.ghost) drawingShape.ghost.remove();
        drawingShape = null;
        document.body.classList.remove("drawing");
        const hint = $("error");
        if (hint && hint.classList.contains("hint")) {
            hint.hidden = true;
            hint.classList.remove("hint");
        }
    }
    function shapeName(kind) {
        if (isControlKind(kind)) return (FORM_CONTROLS.find((c) => `control:${c[0]}` === kind) || [, kind])[1];
        for (const [, list] of SHAPE_GROUPS) for (const [k, name] of list) if (k === kind) return name;
        return kind;
    }
    // A dashed box (with a line across, for a line) where the object will be.
    function ghostAt(g, x0, y0, x1, y1, line) {
        const [x, y, w, h] = [Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0), Math.abs(y1 - y0)];
        Object.assign(g.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${h}px` });
        if (line) {
            const down = (x1 >= x0) === (y1 >= y0);
            g.innerHTML = `<svg width="${w + 2}" height="${h + 2}"><line x1="1" y1="${down ? 1 : h + 1}" x2="${w + 1}" y2="${down ? h + 1 : 1}"/></svg>`;
        }
    }
    function makeGhost(line) {
        const g = view.layer.appendChild(el("div", `objghost${line ? " line" : ""}`));
        return g;
    }

    // The mouse on the grid while a shape is drawn, or on a selected object's handles.
    $("grid").addEventListener(
        "mousedown",
        (e) => {
            if (!drawingShape || !view) return;
            e.preventDefault();
            e.stopPropagation();
            if (e.button !== 0) {
                stopDrawing();
                return;
            }
            const [x, y] = sheetPoint(e);
            drawingShape.start = [x, y];
            drawingShape.now = [x, y];
            drawingShape.ghost = makeGhost(LINES.has(drawingShape.kind));
            ghostAt(drawingShape.ghost, x, y, x, y, LINES.has(drawingShape.kind));
        },
        true,
    );
    $("grid").addEventListener(
        "click",
        (e) => {
            if (!swallowClick) return;
            swallowClick = false;
            e.preventDefault();
            e.stopPropagation();
        },
        true,
    );
    document.addEventListener("mousemove", (e) => {
        if (drawingShape && drawingShape.start) {
            let [x, y] = sheetPoint(e);
            const [x0, y0] = drawingShape.start;
            // Shift: as wide as it is high (a line: across, down, or at 45°).
            if (e.shiftKey) {
                const [w, h] = [Math.abs(x - x0), Math.abs(y - y0)];
                if (LINES.has(drawingShape.kind) && (w < h / 2 || h < w / 2)) {
                    if (w < h) x = x0;
                    else y = y0;
                } else {
                    const d = Math.max(w, h);
                    x = x0 + Math.sign(x - x0 || 1) * d;
                    y = y0 + Math.sign(y - y0 || 1) * d;
                }
            }
            drawingShape.now = [x, y];
            ghostAt(drawingShape.ghost, x0, y0, x, y, LINES.has(drawingShape.kind));
            return;
        }
        if (objDrag) dragObject(e);
    });
    document.addEventListener("mouseup", (e) => {
        if (drawingShape && drawingShape.start) {
            const d = drawingShape;
            stopDrawing();
            swallowClick = true;
            setTimeout(() => (swallowClick = false));
            let [x0, y0] = d.start;
            let [x1, y1] = d.now;
            // A click: the shape's size from there.
            if (Math.abs(x1 - x0) < 3 && Math.abs(y1 - y0) < 3) {
                // A control put with a click: the size Excel gives it.
                if (isControlKind(d.kind)) {
                    putControl(d.kind.slice(8), x0, y0, null);
                    return;
                }
                [x1, y1] = [x0 + clickSize(d.kind)[0], y0 + clickSize(d.kind)[1]];
            }
            putShape(d.kind, x0, y0, x1, y1);
            return;
        }
        if (objDrag) endObjectDrag();
    });

    // A shape put where it was drawn: a line keeps its direction (flipped as Excel writes it); a
    // text box asks for its text first.
    function putShape(kind, x0, y0, x1, y1) {
        if (isControlKind(kind)) {
            putControl(kind.slice(8), Math.min(x0, x1), Math.min(y0, y1), { width: Math.round(Math.abs(x1 - x0)), height: Math.round(Math.abs(y1 - y0)) });
            return;
        }
        const line = LINES.has(kind);
        const op = { op: "shape", sheet: view.sheet.name, shape: kind, ...anchorOf(Math.min(x0, x1), Math.min(y0, y1)), width: Math.round(Math.abs(x1 - x0)), height: Math.round(Math.abs(y1 - y0)) };
        if (line && x1 < x0) op.flipH = true;
        if (line && y1 < y0) op.flipV = true;
        if (kind !== "textBox") {
            pickAfter(null);
            postEdit([op]);
            return;
        }
        askText("テキスト ボックス", "", (text) => {
            op.text = text;
            pickAfter(null);
            postEdit([op]);
        });
    }
    // A form control put where it was drawn (`size` null: Excel's size for it).
    function putControl(kind, x, y, size) {
        const op = { op: "control", sheet: view.sheet.name, control: kind, ...anchorOf(x, y), ...(size || {}) };
        // A scroll bar drawn wider than high goes across.
        if (kind === "scrollBar" && size && size.width > size.height) op.horizontal = true;
        pickAfter(null);
        postEdit([op]);
    }
    // The text of a shape or a text box (Enter starts a new line; Ctrl+Enter enters it).
    function askText(title, now, done) {
        const form = openDialog(title, "notedlg");
        const text = form.appendChild(el("textarea"));
        text.rows = 5;
        text.value = now;
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
            done(text.value);
        });
        text.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                form.requestSubmit();
            }
        });
        text.focus();
        text.setSelectionRange(text.value.length, text.value.length);
    }
    function editObjectText(obj) {
        if (!movable(obj) || obj.text === undefined || busy || !view) return;
        const sheet = view.sheet.name;
        askText(`テキストの編集（${obj.name}）`, obj.text, (text) => {
            if (text === obj.text) return;
            pickAfter(keyOf(obj));
            postEdit([{ op: "object", sheet, ...target(obj), text }]);
        });
    }
    // コントロールの書式設定: the text, the linked cell, the items, the state, the values of a form
    // control (what its kind has).
    function askControlFormat(obj) {
        const f = obj.form;
        if (!isForm(obj) || !f || busy || !view) return;
        const t = f.type;
        const form = openDialog(`コントロールの書式設定（${obj.name || `_x0000_s${obj.shapeId}`}）`, "ctldlg");
        const change = {};
        const inputs = [];
        if (obj.text !== undefined) {
            const line = form.appendChild(el("label"));
            line.appendChild(el("span", "lbl", "文字:"));
            const text = line.appendChild(el("textarea"));
            text.rows = 2;
            text.value = obj.text;
            inputs.push(() => {
                if (text.value !== obj.text) change.text = text.value;
            });
        }
        if (t === "CheckBox" || t === "Radio") {
            const on = check(form, "オン（チェックする）", !!f.checked);
            inputs.push(() => {
                if (on.checked !== !!f.checked) change.checked = on.checked;
            });
        }
        const ref = (label, key) => {
            const input = field(form, label);
            input.value = f[key] || "";
            input.placeholder = key === "items" ? "$A$1:$A$5" : "$B$2";
            inputs.push(() => {
                const v = input.value.trim().replace(/^=/, "");
                if (v !== (f[key] || "")) change[key] = v || null;
            });
        };
        if (t === "List" || t === "Drop") ref("入力範囲:", "items");
        if (["CheckBox", "Radio", "List", "Drop", "Spin", "Scroll"].includes(t)) ref("リンクするセル:", "link");
        if (t === "Spin" || t === "Scroll") {
            const numbers = [["現在値:", "val", "value"], ["最小値:", "min", "min"], ["最大値:", "max", "max"], ["変化の増分:", "inc", "step"]];
            if (t === "Scroll") numbers.push(["ページ単位の変化:", "page", "page"]);
            for (const [label, key, op] of numbers) {
                const input = field(form, label, "number");
                Object.assign(input, { min: "0", max: "30000", step: "1", value: String(f[key]) });
                inputs.push(() => {
                    const v = Math.round(+input.value);
                    if (isFinite(v) && v !== f[key]) change[op] = v;
                });
            }
        }
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            for (const take of inputs) take();
            closeDialog();
            if (Object.keys(change).length) objectEdit(obj, change);
        });
        const firstInput = form.querySelector("textarea, input");
        if (firstInput) firstInput.focus();
    }
    function deleteObject(obj) {
        if (!movable(obj) || busy || !view) return;
        postEdit([{ op: "object", sheet: view.sheet.name, ...target(obj), delete: true }]);
    }
    // The box of an object as it is placed (px at 100%).
    const boxOfObject = (box) => ({ x: box.offsetLeft, y: box.offsetTop, w: box.offsetWidth, h: box.offsetHeight });
    // An object moved (its box) to where it now is, or sized.
    function moveObject(obj, b) {
        if (!view) return;
        pickAfter(keyOf(obj));
        postEdit([{ op: "object", sheet: view.sheet.name, ...target(obj), ...anchorOf(Math.max(0, b.x), Math.max(0, b.y)), width: Math.max(0, Math.round(b.w)), height: Math.max(0, Math.round(b.h)) }]);
    }
    // The handles of the selected object: its sides and corners size it.
    const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
    // (In a frame turned as the object is; above it, the handle that turns it.)
    function addHandles(box, obj) {
        if (!movable(obj)) return;
        box.classList.add("framed");
        const frame = box.appendChild(el("div", "objframe"));
        if (obj.rot) frame.style.transform = `rotate(${obj.rot}deg)`;
        const handles = canTurn(obj) ? HANDLES.concat(["rot"]) : HANDLES;
        for (const h of handles) {
            const d = frame.appendChild(el("div", `objhandle ${h}`));
            if (h === "rot") d.title = "回転（Shift で 15° ずつ）";
            d.addEventListener("mousedown", (e) => {
                if (e.button !== 0) return;
                e.preventDefault();
                e.stopPropagation();
                beginObjectDrag(e, box, obj, h);
            });
        }
    }
    function beginObjectDrag(e, box, obj, how) {
        if (busy || !view) return;
        if (entry) commit(null);
        const r = box.getBoundingClientRect();
        objDrag = { obj, box, how, x: e.clientX, y: e.clientY, from: boxOfObject(box), center: [r.left + r.width / 2, r.top + r.height / 2], ghost: null, now: null };
    }
    function dragObject(e) {
        const d = objDrag;
        if (!(e.buttons & 1)) {
            endObjectDrag();
            return;
        }
        const z = Z();
        let [dx, dy] = [(e.clientX - d.x) / z, (e.clientY - d.y) / z];
        if (!d.ghost && Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
        const f = d.from;
        let { x, y, w, h } = f;
        const turn = ((d.obj.rot || 0) * Math.PI) / 180;
        if (d.how === "rot") {
            // The angle of the mouse around the object's center (0: above it).
            let deg = (Math.atan2(e.clientY - d.center[1], e.clientX - d.center[0]) * 180) / Math.PI + 90;
            if (e.shiftKey) deg = Math.round(deg / 15) * 15;
            deg = ((Math.round(deg) % 360) + 360) % 360;
            if (!d.ghost) d.ghost = makeGhost(false);
            d.now = { rot: deg };
            ghostAt(d.ghost, f.x, f.y, f.x + f.w, f.y + f.h, false);
            d.ghost.style.transform = `rotate(${deg}deg)`;
            return;
        }
        if (d.how === "move") {
            // Shift: across or down only.
            if (e.shiftKey) {
                if (Math.abs(dx) > Math.abs(dy)) dy = 0;
                else dx = 0;
            }
            [x, y] = [f.x + dx, f.y + dy];
        } else {
            // Along the object's own sides when it is turned: the mouse's move in its frame.
            const [cos, sin] = [Math.cos(turn), Math.sin(turn)];
            const [lx, ly] = [dx * cos + dy * sin, -dx * sin + dy * cos];
            let [x0, y0] = [0, 0];
            if (d.how.includes("w")) [x0, w] = [Math.min(lx, f.w), f.w - Math.min(lx, f.w)];
            if (d.how.includes("e")) w = Math.max(0, f.w + lx);
            if (d.how.includes("n")) [y0, h] = [Math.min(ly, f.h), f.h - Math.min(ly, f.h)];
            if (d.how.includes("s")) h = Math.max(0, f.h + ly);
            // A picture keeps its proportion by its corners (Shift: a shape too).
            if (d.how.length === 2 && (d.obj.kind === "picture" || e.shiftKey) && f.w > 0 && f.h > 0) {
                const k = Math.max(w / f.w, h / f.h);
                const [nw, nh] = [f.w * k, f.h * k];
                if (d.how.includes("w")) x0 = f.w - nw;
                if (d.how.includes("n")) y0 = f.h - nh;
                [w, h] = [nw, nh];
            }
            // The side across stays where it is: the center moves by the change turned back.
            const [sx, sy] = [x0 + w / 2 - f.w / 2, y0 + h / 2 - f.h / 2];
            const [cx, cy] = [f.x + f.w / 2 + sx * cos - sy * sin, f.y + f.h / 2 + sx * sin + sy * cos];
            [x, y] = [cx - w / 2, cy - h / 2];
        }
        if (!d.ghost) d.ghost = makeGhost(false);
        d.now = { x, y, w, h };
        ghostAt(d.ghost, x, y, x + w, y + h, false);
        if (turn) d.ghost.style.transform = `rotate(${d.obj.rot}deg)`;
    }
    function endObjectDrag() {
        const d = objDrag;
        objDrag = null;
        if (!d || !d.ghost) return;
        d.ghost.remove();
        swallowClick = true;
        setTimeout(() => (swallowClick = false));
        if (d.how === "rot") objectEdit(d.obj, { rotation: d.now.rot });
        else moveObject(d.obj, d.now);
    }
    // The arrow keys move the selected object by a pixel (Ctrl: by 10).
    function nudgeObject(obj, box, dr, dc, far) {
        if (!movable(obj) || busy) return;
        const b = boxOfObject(box);
        const step = far ? 10 : 1;
        moveObject(obj, { x: b.x + dc * step, y: b.y + dr * step, w: b.w, h: b.h });
    }

    // ---- 図形の書式: the fill, the line, the rotation, the order of the object selected

    // What can be given a fill (a shape, a picture; not a line, a group or a chart), a line, a
    // rotation (not a chart).
    const canFill = (obj) => movable(obj) && (obj.el === "sp" || obj.el === "pic");
    const canLine = (obj) => movable(obj) && (obj.el === "sp" || obj.el === "pic" || obj.el === "cxnSp");
    const canTurn = (obj) => movable(obj) && !!obj.el && obj.el !== "graphicFrame";
    function objectEdit(obj, change) {
        if (!movable(obj) || busy || !view) return;
        pickAfter(keyOf(obj));
        postEdit([{ op: "object", sheet: view.sheet.name, ...target(obj), ...change }]);
    }
    // Turned by `by` degrees (clockwise), as 右へ 90 度回転 turns it.
    function turnObject(obj, by) {
        if (canTurn(obj)) objectEdit(obj, { rotation: ((((obj.rot || 0) + by) % 360) + 360) % 360 });
    }
    // 上下反転 / 左右反転: flipped as it is seen, a turned object turned the other way too.
    function flipObject(obj, vertical) {
        if (!canTurn(obj)) return;
        const change = vertical ? { flipV: !obj.flipV } : { flipH: !obj.flipH };
        if (obj.rot) change.rotation = (360 - obj.rot) % 360;
        objectEdit(obj, change);
    }
    // その他の回転オプション: the angle typed.
    function askRotation(obj) {
        if (!canTurn(obj) || busy) return;
        const form = openDialog(`回転（${obj.name}）`, "rotdlg");
        const input = field(form, "回転角度（°）:", "number");
        input.min = "-3600";
        input.max = "3600";
        input.step = "1";
        input.value = String(Math.round(obj.rot || 0));
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const deg = +input.value;
            if (!isFinite(deg)) return;
            closeDialog();
            objectEdit(obj, { rotation: (((deg % 360) + 360) % 360) });
        });
        input.focus();
        input.select();
    }
    const ORDERS = [
        ["front", "最前面へ移動"],
        ["forward", "前面へ移動"],
        ["backward", "背面へ移動"],
        ["back", "最背面へ移動"],
    ];
    const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
    // The contextual tab of the ribbon: shown while an object that can be changed is selected.
    function refreshShapeTools() {
        const ribbon = $("ribbon");
        if (!ribbon || !ribbon.showShapeTab) return;
        const obj = selected && movable(selectedObj) && !isForm(selectedObj) ? selectedObj : null;
        ribbon.showShapeTab(!!obj);
        if (!obj) return;
        tools.shapeFill.disabled = tools.shapeFillMore.disabled = !canFill(obj);
        tools.shapeLine.disabled = tools.shapeLineMore.disabled = !canLine(obj);
        tools.rotate.disabled = !canTurn(obj);
    }

    // 挿入 > 画像: a picture from a file, at the active cell (shrunk to the window when larger).
    function askPicture() {
        if (!editable || busy || !view || !sel) return;
        if (entry) commit(null);
        const input = document.body.appendChild(el("input"));
        input.type = "file";
        input.accept = "image/png,image/jpeg,image/gif,image/bmp,.png,.jpg,.jpeg,.gif,.bmp";
        input.hidden = true;
        input.addEventListener("change", () => {
            const file = input.files && input.files[0];
            input.remove();
            if (file) putPicture(file);
        });
        input.click();
    }
    function putPicture(file) {
        if (!view || !sel) return;
        if (file.size > 64 * 1024 * 1024) {
            showError("画像が大きすぎます（64 MB まで）。");
            return;
        }
        const sheet = view.sheet.name;
        const at = { at: address(sel.ar, sel.ac) };
        const reader = new FileReader();
        reader.addEventListener("load", () => {
            const url = String(reader.result);
            const img = new Image();
            const post = (size) => {
                pickAfter(null);
                postEdit([{ op: "picture", sheet, ...at, ...size, data: url.slice(url.indexOf(",") + 1) }]);
            };
            img.addEventListener("load", () => {
                const g = $("grid");
                const z = Z();
                const [maxW, maxH] = [Math.max(100, g.clientWidth / z - 60), Math.max(100, g.clientHeight / z - 40)];
                const k = Math.min(1, maxW / img.naturalWidth, maxH / img.naturalHeight);
                post(k < 1 ? { width: Math.round(img.naturalWidth * k), height: Math.round(img.naturalHeight * k) } : {});
            });
            // (A format the page cannot show, BMP in some browsers: its own size.)
            img.addEventListener("error", () => post({}));
            img.src = url;
        });
        reader.readAsDataURL(file);
    }
    // The pictures of the clipboard (a screenshot, an image copied), when it has no text.
    function pastedPicture(e) {
        const d = e.clipboardData;
        if (!editable || !d || d.getData("text/plain")) return false;
        const file = Array.from(d.files || []).find((f) => /^image\/(png|jpeg|gif|bmp)$/.test(f.type));
        if (!file) return false;
        putPicture(file);
        return true;
    }

    // "記号と特殊文字": a symbol typed where the cell is being typed in (or after what the active
    // cell has). The dialog stays open, as Excel's.
    const SYMBOLS = [
        ["記号", "、。・：；？！゛゜´｀¨＾￣＿ヽヾゝゞ〃仝々〆〇ー―‐／＼～∥｜…‥‘’“”（）〔〕［］｛｝〈〉《》「」『』【】＋－±×÷＝≠＜＞≦≧∞∴♂♀°′″℃￥＄￠￡％＃＆＊＠§☆★○●◎◇◆□■△▲▽▼※〒→←↑↓〓"],
        ["数学", "∈∋⊆⊇⊂⊃∪∩∧∨￢⇒⇔∀∃∠⊥⌒∂∇≡≒≪≫√∽∝∵∫∬∑∮√⊿∟"],
        ["ギリシャ", "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω"],
        ["単位", "㍉㌔㌢㍍㌘㌧㌃㌶㍑㍗㌍㌦㌣㌫㍊㌻㎜㎝㎞㎎㎏㏄㎡㍻№㏍℡㊤㊥㊦㊧㊨㈱㈲㈹㍾㍽㍼"],
        ["囲み数字", "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅰⅱⅲⅳⅴⅵⅶⅷⅸⅹ"],
        ["罫線", "─│┌┐┘└├┬┤┴┼━┃┏┓┛┗┣┳┫┻╋┠┯┨┷┿┝┰┥┸╂"],
    ];
    let lastSymbols = Array.isArray(saved.symbols) ? saved.symbols : [];
    function typeSymbol(sym) {
        if (!entry && !beginEntry("edit", null, "cell")) return;
        const box = entry.where === "bar" ? bar : ed;
        const [a, b] = [box.selectionStart ?? box.value.length, box.selectionEnd ?? box.value.length];
        box.setRangeText(sym, a, b, "end");
        box.dispatchEvent(new Event("input"));
        lastSymbols = [sym].concat(lastSymbols.filter((s) => s !== sym)).slice(0, 16);
        remember({ symbols: lastSymbols });
    }
    function askSymbol() {
        if (!editable || busy || !sel || !view || sel.kind !== "cells") return;
        const form = openDialog("記号と特殊文字", "symboldlg");
        const tabs = form.appendChild(el("div", "symtabs"));
        const grid = form.appendChild(el("div", "symgrid"));
        const recent = form.appendChild(el("div", "symrecent"));
        const showRecent = () => {
            recent.textContent = "";
            recent.appendChild(el("span", "lbl", "最近使用した記号:"));
            for (const s of lastSymbols) symbolButton(recent, s);
        };
        const symbolButton = (into, s) => {
            const b = into.appendChild(el("button", "sym", s));
            b.type = "button";
            b.title = `U+${s.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`;
            b.addEventListener("click", () => {
                typeSymbol(s);
                showRecent();
            });
        };
        const show = (i) => {
            grid.textContent = "";
            [...tabs.children].forEach((t, j) => t.classList.toggle("on", i === j));
            for (const s of SYMBOLS[i][1]) symbolButton(grid, s);
        };
        SYMBOLS.forEach(([name], i) => {
            const t = tabs.appendChild(el("button", "symtab", name));
            t.type = "button";
            t.addEventListener("click", () => show(i));
        });
        show(0);
        showRecent();
        const buttons = form.appendChild(el("div", "buttons"));
        const close = buttons.appendChild(el("button", "ok", "閉じる"));
        close.type = "submit";
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
        });
        close.focus();
    }

    // ---- the menu of the cells and the headers

    let menu = null;
    function closeMenu() {
        if (menu) menu.remove();
        menu = null;
    }
    function openMenu(x, y, kind) {
        closeMenu();
        const items = [["コピー", () => menuCopy(false), "Ctrl+C"]];
        if (editable) {
            const header = kind === "rows" || kind === "cols";
            items.unshift(["切り取り", () => menuCopy(true), "Ctrl+X"]);
            items.push(["貼り付け", menuPaste, "Ctrl+V"]);
            if (clip && !clip.cut && clip.kind !== "all") {
                items.push(["値の貼り付け", () => pasteSpecial({ paste: "values", operation: "none", skipBlanks: false, transpose: false }), ""]);
                items.push(["形式を選択して貼り付け…", askPasteSpecial, "Ctrl+Alt+V"]);
            }
            if (kind !== "all") {
                items.push(null);
                items.push([header ? "挿入" : "挿入…", () => insertOrDelete(true), "Ctrl++"]);
                items.push([header ? "削除" : "削除…", () => insertOrDelete(false), "Ctrl+-"]);
            }
            items.push(["数式と値のクリア", clearSel, "Delete"]);
            items.push(["書式のクリア", clearFormats, ""]);
            if (kind === "cells") {
                items.push(null);
                const note = noteAt(sel.ar, sel.ac);
                items.push([note ? "メモの編集" : "新しいメモ", askNote, "Shift+F2"]);
                if (notesIn()) items.push(["メモの削除", removeNote, ""]);
                items.push([linkAt(sel.ar, sel.ac) ? "リンクの編集" : "リンク", askLink, "Ctrl+K"]);
                if (linksIn()) items.push(["リンクの削除", removeLinks, ""]);
                items.push(null);
                if (mergedIn(sel)) items.push(["結合の解除", () => merge("unmerge"), ""]);
                if (!single(sel)) items.push(["セルの結合", () => merge("merge"), ""]);
                items.push(["昇順で並べ替え", () => sortBy(false), ""]);
                items.push(["降順で並べ替え", () => sortBy(true), ""]);
                items.push(["ユーザー設定の並べ替え…", askSort, ""]);
            } else {
                items.push(null);
                if (kind === "cols") items.push(["列の幅…", () => askSize(true), ""]);
                if (kind === "rows") items.push(["行の高さ…", () => askSize(false), ""]);
                if (kind !== "all") items.push(["非表示", () => hide(true), ""]);
                items.push(["再表示", () => hide(false), ""]);
            }
        }
        showMenu(x, y, items);
    }
    // `items`: [label, act, key] or null (a line between them).
    function showMenu(x, y, items) {
        closeMenu();
        menu = document.body.appendChild(el("div", "menu"));
        menu.setAttribute("role", "menu");
        menu.dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: true });
        for (const item of items) {
            if (!item) {
                menu.appendChild(el("hr"));
                continue;
            }
            const b = menu.appendChild(el("button"));
            b.type = "button";
            b.setAttribute("role", "menuitem");
            b.appendChild(el("span", null, item[0]));
            b.appendChild(el("kbd", null, item[2]));
            b.disabled = busy && item[0] !== "コピー";
            b.addEventListener("click", () => {
                closeMenu();
                focusGrid();
                item[1]();
            });
        }
        menu.style.left = `${Math.max(0, Math.min(x, window.innerWidth - menu.offsetWidth - 4))}px`;
        menu.style.top = `${Math.max(0, Math.min(y, window.innerHeight - menu.offsetHeight - 4))}px`;
    }
    $("grid").dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: true });
    $("grid").addEventListener("contextmenu", (e) => {
        const hit = hitOf(e.target);
        if (!hit || e.target === ed) return;
        e.preventDefault();
        e.stopPropagation();
        openMenu(e.clientX, e.clientY, sel ? sel.kind : hit.kind);
    });
    document.addEventListener("mousedown", (e) => {
        if (menu && !menu.contains(e.target)) closeMenu();
    });
    window.addEventListener("blur", closeMenu);

    // ---- the events of the objects: their handlers in the sheet's file, written by the editor
    // An ActiveX control's events are `@on("CommandButton1.Click")` functions of the sheet's
    // file; a shape or a form control runs the macro assigned to it (`OnAction`). The editor says
    // whether the workbook shown is the project's (`events`), whose sheets have files.

    let events = false;
    const hasEvents = (obj) => events && !!obj && (obj.activex !== undefined || !!obj.onAction);
    function openEvent(obj, choose) {
        if (!hasEvents(obj) || busy || !view) return;
        closeMenu();
        vscode.postMessage({
            type: "event",
            sheet: view.sheet.name,
            codeName: view.sheet.codeName || "",
            object: obj.name || "",
            shape: obj.shapeId,
            activex: obj.activex,
            macro: obj.macro || "",
            choose,
        });
    }
    function eventItems(obj) {
        if (!hasEvents(obj)) return [];
        if (obj.activex !== undefined) {
            return [
                ["既定のイベントを開く", () => openEvent(obj, false), ""],
                ["イベントを開く…", () => openEvent(obj, true), ""],
            ];
        }
        return obj.macro
            ? [
                  [`マクロを開く（${obj.macro}）`, () => openEvent(obj, false), ""],
                  ["マクロの登録…", () => openEvent(obj, true), ""],
              ]
            : [
                  ["新しいマクロを登録", () => openEvent(obj, false), ""],
                  ["マクロの登録…", () => openEvent(obj, true), ""],
              ];
    }
    // The menu of an object: its text, deleting it, its events.
    function objectItems(obj) {
        const items = [];
        if (movable(obj)) {
            if (obj.text !== undefined) items.push(["テキストの編集…", () => editObjectText(obj), "F2"]);
            if (isForm(obj) && obj.form) items.push(["コントロールの書式設定…", () => askControlFormat(obj), ""]);
            items.push(["削除", () => deleteObject(obj), "Delete"]);
            if (!isForm(obj)) {
                items.push(null);
                for (const [how, label] of ORDERS) items.push([label, () => objectEdit(obj, { order: how }), ""]);
            }
            if (canTurn(obj)) items.push(["回転…", () => askRotation(obj), ""]);
        }
        const more = eventItems(obj);
        if (items.length && more.length) items.push(null);
        return items.concat(more);
    }
    // The button of the bar: the events of the object selected (a menu).
    function refreshEventButton() {
        const b = $("event");
        if (!b) return;
        const on = hasEvents(selectedObj);
        b.disabled = !on;
        b.querySelector(".lbl").textContent = on && selectedObj.activex === undefined ? "マクロ" : "コードの表示";
    }
    function enableEvents() {
        events = true;
        $("ribbon").showDev();
        refreshEventButton();
    }
    if (editable) vscode.postMessage({ type: "ready" });

    // ---- formats, merged cells, sorting and the filter: the toolbar, the menus, the keys

    const rangeOrCell = () => (single(sel) ? address(sel.ar, sel.ac) : rangeText(sel));
    // The cells selected, a merged cell with all its cells (its format is that of all of them).
    const areaText = (s) => (s.kind !== "cells" ? rangeText(s) : s.r1 === s.r2 && s.c1 === s.c2 ? address(s.r1, s.c1) : `${address(s.r1, s.c1)}:${address(s.r2, s.c2)}`);
    // The CSS the active cell is drawn with (its format, as far as the page knows it).
    const activeCss = () => (view && activeTd && activeTd.cell && activeTd.cell.s && view.sheet.styles[activeTd.cell.s]) || "";
    const STATE = {
        bold: (css) => /font-weight:bold/.test(css),
        italic: (css) => /font-style:italic/.test(css),
        underline: (css) => /text-decoration:[^;]*underline/.test(css),
        strike: (css) => /text-decoration:[^;]*line-through/.test(css),
        wrap: (css) => /white-space:pre-wrap/.test(css),
    };
    function format(set) {
        if (sel && view) postEdit(allAreas().map((a) => ({ op: "format", sheet: view.sheet.name, range: areaText(a), set })));
    }
    function toggleFormat(name) {
        if (editable) format({ [name]: !STATE[name](activeCss()) });
    }
    function clearFormats() {
        if (sel && view) postEdit(allAreas().map((a) => ({ op: "clearFormats", sheet: view.sheet.name, range: areaText(a) })));
    }
    // A merged cell among the cells selected.
    function mergedIn(s) {
        return !!s && view.sheet.merges.some(([a, b, c, d]) => a <= s.r2 && c >= s.r1 && b <= s.c2 && d >= s.c1);
    }
    // `how`: "center" (merged and centered), "across" (each row on its own), "merge", "unmerge".
    function merge(how) {
        if (!sel || !view || sel.kind === "all" || !oneArea()) return;
        const sheet = view.sheet.name;
        const range = areaText(sel);
        if (how === "unmerge") postEdit([{ op: "unmerge", sheet, range }]);
        else if (single(sel)) return;
        else if (how === "across") postEdit([{ op: "merge", sheet, range, across: true }]);
        else if (how === "center") postEdit([{ op: "merge", sheet, range }, { op: "format", sheet, range, set: { align: "center" } }]);
        else postEdit([{ op: "merge", sheet, range }]);
    }
    // By the column of the active cell: the cells selected, or the data around the cell.
    function sortBy(desc) {
        if (sel && view && oneArea()) postEdit([{ op: "sort", sheet: view.sheet.name, range: rangeOrCell(), by: colName(sel.ac), desc }]);
    }
    function toggleFilter() {
        if (!editable || !sel || !view || !oneArea()) return;
        const sheet = view.sheet.name;
        postEdit([view.sheet.filter ? { op: "filter", sheet, off: true } : { op: "filter", sheet, range: rangeOrCell() }]);
    }
    function hide(hidden) {
        if (!sel || !view) return;
        const sheet = view.sheet.name;
        if (sel.kind === "all") postEdit([{ op: "hide", sheet, range: `1:${MAXR}`, hidden }, { op: "hide", sheet, range: `A:${colName(MAXC - 1)}`, hidden }]);
        else if (allAreas().every((a) => a.kind === sel.kind && a.kind !== "cells")) postEdit(allAreas().map((a) => ({ op: "hide", sheet, range: rangeText(a), hidden })));
    }

    // A box under a button of the toolbar (closed as the menu is).
    function openPop(anchor, cls) {
        closeMenu();
        menu = document.body.appendChild(el("div", `menu pop ${cls || ""}`));
        menu.dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: true });
        const b = anchor.getBoundingClientRect();
        const pop = menu;
        // Placed when it is filled.
        requestAnimationFrame(() => {
            if (menu !== pop) return;
            pop.style.left = `${Math.max(0, Math.min(b.left, window.innerWidth - pop.offsetWidth - 4))}px`;
            pop.style.top = `${Math.max(0, Math.min(b.bottom + 1, window.innerHeight - pop.offsetHeight - 4))}px`;
        });
        Object.assign(pop.style, { left: `${b.left}px`, top: `${b.bottom + 1}px` });
        return pop;
    }
    function popItem(pop, label, act) {
        const b = pop.appendChild(el("button"));
        b.type = "button";
        b.appendChild(el("span", null, label));
        b.addEventListener("click", () => {
            closeMenu();
            focusGrid();
            act();
        });
        return b;
    }
    const PALETTE = ["#000000", "#ffffff", "#7f7f7f", "#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0", "#f2f2f2", "#d9d9d9", "#fce4d6", "#fff2cc", "#e2efda", "#ddebf7", "#d9e1f2"];
    // A color chosen: given to the cells, and kept for the button (as Excel's last color).
    function useColor(name, color) {
        lastColor[name] = color;
        const bar = tools[name] && tools[name].querySelector(".bar");
        if (bar) bar.style.background = color;
        format({ [name]: color });
    }
    // The colors to choose from; `name` is the format ("color", "fill"), `none` what no color is called.
    // `pick`: what the color chosen is for (else the cells' format).
    function openColors(anchor, name, none, pick) {
        const pop = openPop(anchor, "colors");
        popItem(pop, none, () => (pick ? pick(null) : format({ [name]: null })));
        const grid = pop.appendChild(el("div", "swatches"));
        for (const color of PALETTE) {
            const b = grid.appendChild(el("button", "swatch"));
            b.type = "button";
            b.title = color;
            b.dataset.color = color;
            b.style.background = color;
            b.addEventListener("click", () => {
                closeMenu();
                focusGrid();
                if (pick) pick(color.toUpperCase());
                else useColor(name, color.toUpperCase());
            });
        }
        // Any other color, as #RRGGBB.
        const form = pop.appendChild(el("form", "hex"));
        const input = form.appendChild(el("input"));
        input.type = "text";
        input.placeholder = "#RRGGBB";
        input.maxLength = 7;
        form.appendChild(el("button", null, "OK")).type = "submit";
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const m = /^#?([0-9a-f]{6})$/i.exec(input.value.trim());
            if (!m) return;
            closeMenu();
            focusGrid();
            if (pick) pick(`#${m[1].toUpperCase()}`);
            else useColor(name, `#${m[1].toUpperCase()}`);
        });
    }
    const BORDERS = [["bottom", "下罫線"], ["top", "上罫線"], ["left", "左罫線"], ["right", "右罫線"], ["none", "枠なし"], ["all", "格子"], ["outline", "外枠"], ["thick", "太い外枠"]];
    const NUMBER_FORMATS = [["General", "標準"], ["0.00", "数値"], ["#,##0", "桁区切り"], ["¥#,##0", "通貨"], ["0%", "パーセント"], ["yyyy/m/d", "日付"], ["h:mm", "時刻"], ["@", "文字列"]];
    const FONT_SIZES = [6, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
    const FONT_NAMES = ["游ゴシック", "游ゴシック Light", "游明朝", "ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ 明朝", "ＭＳ Ｐ明朝", "メイリオ", "Meiryo UI", "BIZ UDゴシック", "BIZ UDPゴシック", "Arial", "Calibri", "Century", "Consolas", "Segoe UI", "Times New Roman", "Verdana"];

    // The number format with one decimal more (fewer): Excel's 小数点以下の表示桁数を増やす (減らす).
    function stepDecimals(code, up) {
        if (code === "@") return null;
        if (code === "General") return up ? "0.0" : "0";
        return code
            .split(";")
            .map((part) => {
                const m = /\.(0+)/.exec(part);
                if (m) {
                    const n = m[1].length + (up ? 1 : -1);
                    return part.slice(0, m.index) + (n > 0 ? `.${"0".repeat(n)}` : "") + part.slice(m.index + m[0].length);
                }
                if (!up) return part;
                const d = /[0#](?![\s\S]*[0#])/.exec(part);
                return d ? `${part.slice(0, d.index + 1)}.0${part.slice(d.index + 1)}` : part;
            })
            .join(";");
    }
    // The font size of the active cell one step up (down): Excel's フォント サイズの拡大 (縮小).
    function stepFontSize(up) {
        const now = +((/font-size:([\d.]+)pt/.exec(activeCss()) || [])[1] || 11);
        const next = up ? FONT_SIZES.find((n) => n > now) : [...FONT_SIZES].reverse().find((n) => n < now);
        if (next) format({ fontSize: next });
    }
    // The whole columns (rows) of the selection, as the ranges `size` and `hide` take.
    const wholeCols = () => (sel.kind === "all" ? `A:${colName(MAXC - 1)}` : `${colName(sel.c1)}:${colName(sel.c2)}`);
    const wholeRows = () => `${sel.r1 + 1}:${sel.r2 + 1}`;
    function hideWhole(cols, hidden) {
        if (!sel || !view || !oneArea()) return;
        postEdit([{ op: "hide", sheet: view.sheet.name, range: cols ? wholeCols() : wholeRows(), hidden }]);
    }
    // The columns (rows) of the selection as wide (tall) as their contents, each on its own.
    function autofitSel(cols) {
        if (!sel || !view || !oneArea()) return;
        const v = view;
        const indexes = cols ? v.cols.filter((c) => c >= sel.c1 && c <= sel.c2) : v.rows.map(([r]) => r).filter((r) => r >= sel.r1 && r <= sel.r2);
        const ops = indexes.map((i) => ({ op: "size", sheet: v.sheet.name, range: cols ? `${colName(i)}:${colName(i)}` : `${i + 1}:${i + 1}`, px: Math.round(fitPx(cols, i)) }));
        if (ops.length) postEdit(ops);
    }
    // Every column of the filter shown again (Excel's データ > クリア).
    function clearFilters() {
        if (!view || !view.sheet.filter || !view.filter) return;
        const sheet = view.sheet;
        const c1 = view.filter[0][1];
        const ops = Object.keys(sheet.filterColumns || {}).map((k) => ({ op: "filterBy", sheet: sheet.name, column: colName(c1 + +k), clear: true }));
        if (ops.length) postEdit(ops);
    }
    // The zoom that shows the cells selected in the window (Excel's 選択範囲に合わせて拡大/縮小).
    function zoomToSelection() {
        if (!sel || !view || sel.kind !== "cells") return;
        const b = boxOf(sel.r1, sel.c1, sel.r2, sel.c2);
        if (!b || !b.w || !b.h) return;
        const g = $("grid");
        const head = rowTop(view, 0);
        const left = columnLeft(view, 0);
        setZoom(Math.min(((g.clientWidth - left * Z()) / b.w) * 100, ((g.clientHeight - head * Z()) / b.h) * 100, 400), false);
        reveal(sel.r1, sel.c1);
    }
    function openZoomBox(anchor) {
        const pop = openPop(anchor, "zooms");
        pop.appendChild(el("div", "phead", "倍率"));
        for (const pct of [200, 100, 75, 50, 25]) popItem(pop, `${pct}%`, () => setZoom(pct)).classList.toggle("on", Math.round(Z() * 100) === pct);
        popItem(pop, "選択範囲に合わせる", zoomToSelection);
    }

    // The icons of the ribbon (16 × 16, drawn larger on the large buttons).
    const ICON_PATHS = {
        undo: '<path d="M5 3 2 6l3 3"/><path d="M2 6h8a4 4 0 0 1 0 8H7"/>',
        redo: '<path d="m11 3 3 3-3 3"/><path d="M14 6H6a4 4 0 0 0 0 8h3"/>',
        paste: '<rect x="2.5" y="2.5" width="9" height="12" rx="1"/><path d="M5 2.5V1.5h4v1"/><rect x="7.5" y="7.5" width="7" height="7.5" fill="#fff"/><path d="M9 10h4M9 12.5h4"/>',
        cut: '<circle cx="4.5" cy="12" r="2"/><circle cx="11.5" cy="12" r="2"/><path d="M5.8 10.5 11 2M10.2 10.5 5 2"/>',
        copy: '<rect x="2.5" y="2.5" width="8" height="9" rx="1"/><rect x="5.5" y="5.5" width="8" height="9" rx="1" fill="#fff"/>',
        grow: '<path d="M2 13 5.5 4 9 13M3.3 10h4.4"/><path d="m11 6 2-2.5L15 6"/>',
        shrink: '<path d="M2 13 5.5 4 9 13M3.3 10h4.4"/><path d="m11 4 2 2.5L15 4"/>',
        border: '<rect x="2.5" y="2.5" width="11" height="11" stroke-dasharray="1.5 1.5"/><path d="M2 13.5h12" stroke-width="2"/>',
        fill: '<path d="m3 8 5-5 5 5-5 5z"/><path d="M13 9.5c.8 1.2 1.2 2 1.2 2.5a1.2 1.2 0 0 1-2.4 0c0-.5.4-1.3 1.2-2.5z"/>',
        fontColor: '<path d="M4 12 8 2l4 10M5.5 8.5h5"/>',
        top: '<path d="M2 2.5h12"/><path d="M8 13V5M5 8l3-3 3 3"/>',
        middle: '<path d="M2 8h12"/><path d="M8 1v4M6 3.5 8 5l2-1.5M8 15v-4M6 12.5 8 11l2 1.5"/>',
        bottom: '<path d="M2 13.5h12"/><path d="M8 3v8M5 8l3 3 3-3"/>',
        left: '<path d="M2 3h12M2 6.3h8M2 9.6h12M2 13h8"/>',
        center: '<path d="M2 3h12M4 6.3h8M2 9.6h12M4 13h8"/>',
        right: '<path d="M2 3h12M6 6.3h8M2 9.6h12M6 13h8"/>',
        wrap: '<path d="M2 3h12M2 7h10a2 2 0 0 1 0 4H8"/><path d="m9.5 9.5-1.5 1.5 1.5 1.5"/><path d="M2 11h3"/>',
        merge: '<rect x="1.5" y="3.5" width="13" height="9"/><path d="M4 8h8M5.5 6.5 4 8l1.5 1.5M10.5 6.5 12 8l-1.5 1.5"/>',
        currency: '<path d="M4 2.5 8 8l4-5.5M8 8v6M5 9h6M5 11.5h6"/>',
        percent: '<circle cx="4.5" cy="4.5" r="2"/><circle cx="11.5" cy="11.5" r="2"/><path d="M13 3 3 13"/>',
        comma: '<path d="M6 12.5c0 1.5-.8 2.5-2 3"/><circle cx="5" cy="11.5" r="1" fill="currentColor"/><path d="M9 4v8M12 4v8"/>',
        decUp: '<path d="M1.5 13h1"/><path d="M5 5.5h3v7.5H5zM10.5 5.5h3v7.5h-3z"/><path d="M1.5 3h4M3.5 1v4"/>',
        decDown: '<path d="M1.5 13h1"/><path d="M5 5.5h3v7.5H5zM10.5 5.5h3v7.5h-3z"/><path d="M1.5 3h4"/>',
        insert: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M2.5 8h11M8 2.5v11"/><path d="M11.5 1v4M9.5 3h4" stroke-width="1.6"/>',
        delete: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M2.5 8h11M8 2.5v11"/><path d="m10 1 3.5 3.5M13.5 1 10 4.5" stroke-width="1.6"/>',
        cellFormat: '<rect x="2.5" y="2.5" width="11" height="11"/><path d="M2.5 6h11M6 2.5v11"/><path d="M8.5 12.5 13 8" stroke-width="1.6"/>',
        fillDown: '<rect x="3.5" y="1.5" width="9" height="4"/><path d="M8 7v7M5 11l3 3 3-3"/>',
        clear: '<path d="m6 13.5-3-3L10 3.5l4 4-6 6zM6.5 6.5l4 4M3 14.5h11"/>',
        sortFilter: '<path d="M3.5 2v11M1.5 11l2 2 2-2"/><path d="M7 3h7l-2.8 3.5V11l-1.4 1V6.5z"/>',
        find: '<circle cx="6.5" cy="6.5" r="4"/><path d="m9.5 9.5 4.5 4.5"/>',
        sortAsc: '<path d="M12 2v11M10 11l2 2 2-2"/><text x="1" y="7" font-size="6.5" stroke="none" fill="currentColor" font-family="sans-serif">A</text><text x="1" y="14.5" font-size="6.5" stroke="none" fill="currentColor" font-family="sans-serif">Z</text>',
        sortDesc: '<path d="M12 2v11M10 11l2 2 2-2"/><text x="1" y="7" font-size="6.5" stroke="none" fill="currentColor" font-family="sans-serif">Z</text><text x="1" y="14.5" font-size="6.5" stroke="none" fill="currentColor" font-family="sans-serif">A</text>',
        sort: '<rect x="1.5" y="2.5" width="7" height="11"/><path d="M1.5 6.2h7M1.5 9.8h7"/><path d="M12 2v11M10 11l2 2 2-2"/>',
        filter: '<path d="M2 2.5h12L9.5 8v5.5l-3-1.5V8z"/>',
        filterClear: '<path d="M2 2.5h9L7.5 7v5l-2-1V7z"/><path d="m10 9.5 4 4M14 9.5l-4 4"/>',
        formulas: '<text x="1" y="12" font-size="10" font-style="italic" stroke="none" fill="currentColor" font-family="serif">fx</text><path d="M11 5h4M11 9h4M11 13h4"/>',
        zoom: '<circle cx="6.5" cy="6.5" r="4.5"/><path d="m10 10 4.5 4.5M4.5 6.5h4M6.5 4.5v4"/>',
        zoom100: '<rect x="1.5" y="3.5" width="13" height="9"/><text x="3" y="10.5" font-size="5.5" stroke="none" fill="currentColor" font-family="sans-serif">100</text>',
        zoomSel: '<rect x="1.5" y="1.5" width="13" height="13" stroke-dasharray="2 1.5"/><rect x="4.5" y="4.5" width="7" height="7"/>',
        freeze: '<rect x="1.5" y="1.5" width="13" height="13"/><path d="M1.5 5.5h13M5.5 1.5v13" stroke-width="2"/>',
        code: '<path d="M5 4 1.5 8 5 12M11 4l3.5 4L11 12M9.5 2.5l-3 11"/>',
        gotoCell: '<path d="M2 8h9M8 4.5 11.5 8 8 11.5"/><path d="M13.5 2v12"/>',
        table: '<rect x="1.5" y="2.5" width="13" height="11"/><path d="M1.5 6h13" stroke-width="2.2"/><path d="M1.5 9.5h13M6 6v7.5M10.5 6v7.5"/>',
        link: '<path d="M6.5 9.5 9.5 6.5"/><path d="M7.5 4.5 9 3a2.5 2.5 0 0 1 3.5 3.5L11 8M8.5 11.5 7 13a2.5 2.5 0 0 1-3.5-3.5L5 8"/>',
        note: '<path d="M2.5 2.5h11v8l-3 3h-8z"/><path d="M10.5 13.5v-3h3"/><path d="M5 6h6M5 8.5h4"/>',
        symbol: '<path d="M3 13.5h3.2v-1.6A5 5 0 1 1 9.8 11.9v1.6H13"/>',
        picture: '<rect x="1.5" y="2.5" width="13" height="11" rx="1"/><circle cx="5.5" cy="6" r="1.3"/><path d="m1.5 12 4-4 3 3 2-2 4 4"/>',
        shapes: '<circle cx="5.5" cy="5.5" r="4"/><rect x="7.5" y="7.5" width="7" height="7" fill="#fff"/>',
        textBox: '<rect x="1.5" y="2.5" width="13" height="11"/><path d="M4.5 5.5h7M8 5.5v6"/>',
        shapeLine: '<path d="M12.5 1.5 5 9l-1 3 3-1 7.5-7.5z"/>',
        forward: '<rect x="1.5" y="1.5" width="8" height="8" stroke-dasharray="1.5 1.5"/><rect x="6.5" y="6.5" width="8" height="8" fill="#dbe5f1"/>',
        backward: '<rect x="6.5" y="6.5" width="8" height="8" stroke-dasharray="1.5 1.5"/><rect x="1.5" y="1.5" width="8" height="8" fill="#dbe5f1"/>',
        rotate: '<path d="M13 8a5 5 0 1 1-1.5-3.6"/><path d="M12 1.5v3.2H8.8"/>',
        controls: '<rect x="1.5" y="2.5" width="6" height="4" rx=".8"/><rect x="9.5" y="2.5" width="5" height="5"/><path d="m10.5 5 1.2 1.2 1.8-2.2"/><circle cx="4.5" cy="11.5" r="2.5"/><rect x="9" y="9.5" width="5.5" height="4"/>',
    };
    function icon(name, large) {
        const s = large ? 30 : 16;
        const span = el("span", "ico");
        span.innerHTML = `<svg viewBox="0 0 16 16" width="${s}" height="${s}" fill="none" stroke="currentColor" stroke-width="${large ? 1 : 1.2}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON_PATHS[name] || ""}</svg>`;
        return span;
    }

    // The ribbon, as Excel's: its tabs, their groups, and the buttons (by what they do). The
    // buttons that edit are off where the workbook cannot be edited.
    const tools = {};
    const lastColor = { color: "#FF0000", fill: "#FFFF00" };
    let lastBorder = "bottom";
    function makeRibbon() {
        const ribbon = $("ribbon");
        const tabs = ribbon.appendChild(el("div", "rtabs"));
        tabs.setAttribute("role", "tablist");
        const body = ribbon.appendChild(el("div", "rbody"));
        // The grid keeps the keys: a click on a button does not take them.
        ribbon.addEventListener("mousedown", (e) => {
            if (!e.target.closest("select, input")) e.preventDefault();
        });
        const panels = {};
        let shown = saved.ribbonTab || "home";
        let collapsed = saved.ribbonCollapsed === undefined ? !editable : saved.ribbonCollapsed;
        // `keep`: the tab chosen stays remembered (a contextual tab hidden for a while).
        const show = (key, keep) => {
            shown = key;
            for (const [k, p] of Object.entries(panels)) {
                p.panel.hidden = k !== key;
                p.tab.classList.toggle("on", k === key);
                p.tab.setAttribute("aria-selected", String(k === key));
            }
            if (!keep) remember({ ribbonTab: key });
        };
        const setCollapsed = (on) => {
            collapsed = on;
            ribbon.classList.toggle("collapsed", on);
            ribbon.classList.remove("open");
            remember({ ribbonCollapsed: on });
        };
        // A collapsed ribbon opens over the grid while a tab is clicked, and closes after a command.
        const closeOpen = () => ribbon.classList.remove("open");
        document.addEventListener("mousedown", (e) => {
            if (ribbon.classList.contains("open") && !ribbon.contains(e.target) && !(menu && menu.contains(e.target))) closeOpen();
        });
        const tab = (key, label) => {
            const t = tabs.appendChild(el("button", "rtab", label));
            t.type = "button";
            t.setAttribute("role", "tab");
            t.dataset.tab = key;
            t.title = "ダブルクリックでリボンを折りたたむ／展開する";
            t.addEventListener("click", () => {
                if (collapsed) ribbon.classList.toggle("open", !(ribbon.classList.contains("open") && shown === key));
                show(key);
            });
            t.addEventListener("dblclick", () => {
                setCollapsed(!collapsed);
                show(key);
                focusGrid();
            });
            const panel = body.appendChild(el("div", "rpanel"));
            panel.setAttribute("role", "tabpanel");
            panel.dataset.tab = key;
            panels[key] = { tab: t, panel };
            return panel;
        };
        let into = null;
        const group = (panel, label) => {
            const g = panel.appendChild(el("div", "rgroup"));
            into = g.appendChild(el("div", "rcontent"));
            g.appendChild(el("div", "rlabel", label));
            return into;
        };
        // A column of small buttons (Excel stacks up to three), a row of them.
        const stack = () => (into = into.closest(".rgroup").querySelector(".rcontent").appendChild(el("div", "rstack")));
        const row = (where) => (into = (where || into).appendChild(el("div", "rrow")));
        const run = (b, act, edits, keep) => {
            if (busy && edits) return;
            if (edits && !sel) return;
            if (entry && !keep) commit(null);
            act(b);
            if (!menu) closeOpen();
            focusGrid();
        };
        // `o`: { label, title, icon, large, edits, menu: (anchor) => fill it, split, cls, text,
        // keepEntry (what is being typed in the cell stays) }.
        const button = (key, o) => {
            const wrap = o.split ? into.appendChild(el("div", `rsplit${o.large ? " large" : ""}`)) : into;
            const b = wrap.appendChild(el("button", `rbtn${o.large ? " large" : ""}${o.cls ? ` ${o.cls}` : ""}`));
            b.type = "button";
            b.title = o.title || o.label || "";
            b.dataset.tool = key;
            if (o.icon) b.appendChild(icon(o.icon, o.large));
            if (o.text) b.appendChild(el("span", "glyph", o.text));
            if (o.label && (o.large || o.showLabel)) b.appendChild(el("span", "lbl", o.label));
            if (o.menu && !o.split) b.appendChild(el("span", "caret", "▾"));
            if (o.edits && !editable) b.disabled = true;
            const openMenuOf = (anchor) => {
                if (o.edits && busy) return;
                if (entry) commit(null);
                o.menu(anchor);
            };
            if (o.menu && !o.split) b.addEventListener("click", () => openMenuOf(wrap === into ? b : wrap));
            else b.addEventListener("click", () => run(b, o.act, o.edits, o.keepEntry));
            if (o.split) {
                const more = wrap.appendChild(el("button", "rbtn more", "▾"));
                more.type = "button";
                more.title = o.moreTitle || o.title || o.label;
                more.dataset.tool = `${key}More`;
                if (o.edits && !editable) more.disabled = true;
                more.addEventListener("click", () => openMenuOf(wrap));
                tools[`${key}More`] = more;
            }
            tools[key] = b;
            return b;
        };
        const select = (key, title, options, cls) => {
            const s = into.appendChild(el("select", cls || null));
            s.title = title;
            s.dataset.tool = key;
            for (const [value, label] of options) s.appendChild(el("option", null, label)).value = value;
            if (!editable) s.disabled = true;
            tools[key] = s;
            return s;
        };
        const sep = () => into.appendChild(el("span", "rsep"));
        const item = (pop, label, act, kbd) => {
            const b = popItem(pop, label, () => {
                act();
                closeOpen();
            });
            if (kbd) b.appendChild(el("kbd", null, kbd));
            return b;
        };
        const head = (pop, text) => pop.appendChild(el("div", "phead", text));

        // ---- ホーム
        const home = tab("home", "ホーム");
        group(home, "元に戻す");
        stack();
        button("undo", { icon: "undo", label: "元に戻す", title: "元に戻す（Ctrl+Z）", showLabel: true, edits: true, act: () => postHistory("undo") });
        button("redo", { icon: "redo", label: "やり直し", title: "やり直し（Ctrl+Y）", showLabel: true, edits: true, act: () => postHistory("redo") });

        group(home, "クリップボード");
        button("paste", {
            icon: "paste",
            label: "貼り付け",
            title: "貼り付け（Ctrl+V）",
            moreTitle: "貼り付けのオプション",
            large: true,
            split: true,
            edits: true,
            act: menuPaste,
            menu: (anchor) => {
                const pop = openPop(anchor, "pastes");
                item(pop, "貼り付け", menuPaste, "Ctrl+V");
                const special = clip && !clip.cut && clip.kind !== "all";
                item(pop, "値の貼り付け", () => pasteSpecial({ paste: "values", operation: "none", skipBlanks: false, transpose: false })).disabled = !special;
                item(pop, "書式の貼り付け", () => pasteSpecial({ paste: "formats", operation: "none", skipBlanks: false, transpose: false })).disabled = !special;
                item(pop, "行列を入れ替える", () => pasteSpecial({ paste: "all", operation: "none", skipBlanks: false, transpose: true })).disabled = !special;
                pop.appendChild(el("hr"));
                item(pop, "形式を選択して貼り付け…", askPasteSpecial, "Ctrl+Alt+V").disabled = !special;
            },
        });
        stack();
        button("cut", { icon: "cut", label: "切り取り", title: "切り取り（Ctrl+X）", showLabel: true, edits: true, act: () => menuCopy(true) });
        button("copy", { icon: "copy", label: "コピー", title: "コピー（Ctrl+C）", showLabel: true, act: () => menuCopy(false) });

        group(home, "フォント");
        const fontRows = into.appendChild(el("div", "rstack wide"));
        row(fontRows);
        const fontName = into.appendChild(el("input", "fontname"));
        fontName.type = "text";
        fontName.title = "フォント";
        fontName.dataset.tool = "fontName";
        fontName.setAttribute("list", "fontnames");
        fontName.spellcheck = false;
        fontName.autocomplete = "off";
        const names = into.appendChild(el("datalist"));
        names.id = "fontnames";
        for (const n of FONT_NAMES) names.appendChild(el("option")).value = n;
        if (!editable) fontName.disabled = true;
        tools.fontName = fontName;
        const applyFont = () => {
            const name = fontName.value.trim();
            if (name && editable && sel) format({ fontName: name });
            focusGrid();
        };
        fontName.addEventListener("change", applyFont);
        fontName.addEventListener("keydown", (e) => {
            if (e.key === "Escape") {
                e.preventDefault();
                showTools();
                focusGrid();
            }
        });
        const size = select("fontSize", "フォント サイズ", FONT_SIZES.map((n) => [String(n), String(n)]), "fontsize");
        size.addEventListener("change", () => {
            format({ fontSize: +size.value });
            focusGrid();
        });
        button("fontGrow", { icon: "grow", title: "フォント サイズの拡大", edits: true, act: () => stepFontSize(true) });
        button("fontShrink", { icon: "shrink", title: "フォント サイズの縮小", edits: true, act: () => stepFontSize(false) });
        row(fontRows);
        button("bold", { text: "B", title: "太字（Ctrl+B）", cls: "b", edits: true, act: () => toggleFormat("bold") });
        button("italic", { text: "I", title: "斜体（Ctrl+I）", cls: "i", edits: true, act: () => toggleFormat("italic") });
        button("underline", { text: "U", title: "下線（Ctrl+U）", cls: "u", edits: true, act: () => toggleFormat("underline") });
        sep();
        button("border", {
            icon: "border",
            title: "罫線",
            moreTitle: "罫線の種類",
            split: true,
            edits: true,
            act: () => format({ border: lastBorder }),
            menu: (anchor) => {
                const pop = openPop(anchor, "borders");
                head(pop, "罫線");
                for (const [value, label] of BORDERS) {
                    item(pop, label, () => {
                        lastBorder = value;
                        tools.border.title = `罫線（${label}）`;
                        format({ border: value });
                    }).dataset.border = value;
                }
            },
        });
        sep();
        const colorButton = (key, iconName, title, none) => {
            const b = button(key, {
                icon: iconName,
                title,
                moreTitle: `${title}の選択`,
                split: true,
                edits: true,
                cls: "color",
                act: () => format({ [key]: lastColor[key] }),
                menu: (anchor) => openColors(anchor, key, none),
            });
            b.querySelector(".ico").appendChild(el("span", "bar")).style.background = lastColor[key];
        };
        colorButton("fill", "fill", "塗りつぶしの色", "塗りつぶしなし");
        colorButton("color", "fontColor", "フォントの色", "自動");

        group(home, "配置");
        const alignRows = into.appendChild(el("div", "rstack wide"));
        row(alignRows);
        for (const [value, ic, title] of [["top", "top", "上揃え"], ["middle", "middle", "上下中央揃え"], ["bottom", "bottom", "下揃え"]]) button(`valign-${value}`, { icon: ic, title, edits: true, act: () => format({ valign: value }) });
        sep();
        button("wrap", { icon: "wrap", label: "折り返して全体を表示する", showLabel: true, title: "折り返して全体を表示する", edits: true, act: () => toggleFormat("wrap") });
        row(alignRows);
        for (const [value, ic, title] of [["left", "left", "左揃え"], ["center", "center", "中央揃え"], ["right", "right", "右揃え"]]) button(`align-${value}`, { icon: ic, title, edits: true, act: () => format({ align: value }) });
        sep();
        button("merge", {
            icon: "merge",
            label: "セルを結合して中央揃え",
            showLabel: true,
            title: "セルを結合して中央揃え（結合したセルでは結合の解除）",
            moreTitle: "セルの結合",
            split: true,
            edits: true,
            act: () => merge(sel && single(sel) && mergedIn(sel) ? "unmerge" : "center"),
            menu: (anchor) => {
                const pop = openPop(anchor, "merges");
                for (const [how, label] of [["center", "セルを結合して中央揃え"], ["across", "横方向に結合"], ["merge", "セルの結合"], ["unmerge", "セル結合の解除"]]) item(pop, label, () => merge(how)).dataset.merge = how;
            },
        });

        group(home, "数値");
        const numberRows = into.appendChild(el("div", "rstack wide"));
        row(numberRows);
        const number = select("numberFormat", "表示形式", [["", "ユーザー定義"]].concat(NUMBER_FORMATS), "numberformat");
        number.addEventListener("change", () => {
            if (number.value) format({ numberFormat: number.value });
            focusGrid();
        });
        row(numberRows);
        button("currency", { icon: "currency", title: "通貨表示形式", edits: true, act: () => format({ numberFormat: "¥#,##0" }) });
        button("percent", { icon: "percent", title: "パーセント スタイル", edits: true, act: () => format({ numberFormat: "0%" }) });
        button("comma", { icon: "comma", title: "桁区切りスタイル", edits: true, act: () => format({ numberFormat: "#,##0" }) });
        sep();
        const decimals = (up) => {
            const cell = (activeTd && activeTd.cell) || {};
            const code = stepDecimals(cell.n || "General", up);
            if (code) format({ numberFormat: code });
        };
        button("decUp", { icon: "decUp", title: "小数点以下の表示桁数を増やす", edits: true, act: () => decimals(true) });
        button("decDown", { icon: "decDown", title: "小数点以下の表示桁数を減らす", edits: true, act: () => decimals(false) });

        group(home, "セル");
        button("insert", {
            icon: "insert",
            label: "挿入",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "inserts");
                item(pop, "セルの挿入…", () => insertOrDelete(true), "Ctrl++");
                item(pop, "シートの行を挿入", () => structure(true, "rows"));
                item(pop, "シートの列を挿入", () => structure(true, "cols"));
            },
        });
        button("delete", {
            icon: "delete",
            label: "削除",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "deletes");
                item(pop, "セルの削除…", () => insertOrDelete(false), "Ctrl+-");
                item(pop, "シートの行を削除", () => structure(false, "rows"));
                item(pop, "シートの列を削除", () => structure(false, "cols"));
            },
        });
        button("cellFormat", {
            icon: "cellFormat",
            label: "書式",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "cellformats");
                head(pop, "セルのサイズ");
                item(pop, "行の高さ…", () => askSize(false));
                item(pop, "行の高さの自動調整", () => autofitSel(false));
                item(pop, "列の幅…", () => askSize(true));
                item(pop, "列の幅の自動調整", () => autofitSel(true));
                head(pop, "表示設定");
                item(pop, "行を表示しない", () => hideWhole(false, true));
                item(pop, "列を表示しない", () => hideWhole(true, true));
                item(pop, "行の再表示", () => hideWhole(false, false));
                item(pop, "列の再表示", () => hideWhole(true, false));
            },
        });

        group(home, "編集");
        stack();
        button("fillDown", {
            icon: "fillDown",
            label: "フィル",
            showLabel: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "fills");
                item(pop, "下方向へコピー", () => fill("down"), "Ctrl+D");
                item(pop, "右方向へコピー", () => fill("right"), "Ctrl+R");
            },
        });
        button("clear", {
            icon: "clear",
            label: "クリア",
            showLabel: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "clears");
                item(pop, "すべてクリア", () => {
                    if (sel && view) postEdit(allAreas().flatMap((a) => [{ op: "clear", sheet: view.sheet.name, range: rangeText(a) }, { op: "clearFormats", sheet: view.sheet.name, range: areaText(a) }]));
                });
                item(pop, "書式のクリア", clearFormats);
                item(pop, "数式と値のクリア", clearSel, "Delete");
            },
        });
        into = into.parentNode;
        button("sortFilter", {
            icon: "sortFilter",
            label: "並べ替えとフィルター",
            large: true,
            edits: true,
            cls: "two",
            menu: (anchor) => {
                const pop = openPop(anchor, "sorts");
                item(pop, "昇順", () => sortBy(false));
                item(pop, "降順", () => sortBy(true));
                item(pop, "ユーザー設定の並べ替え…", askSort);
                pop.appendChild(el("hr"));
                item(pop, "フィルター", toggleFilter, "Ctrl+Shift+L").classList.toggle("on", !!(view && view.sheet.filter));
                item(pop, "クリア", clearFilters).disabled = !(view && view.sheet.filter && Object.keys(view.sheet.filterColumns || {}).length);
            },
        });
        button("findSelect", {
            icon: "find",
            label: "検索と選択",
            large: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "finds");
                item(pop, "検索…", () => openFind(false), "Ctrl+F");
                if (editable) item(pop, "置換…", () => openFind(true), "Ctrl+H");
                item(pop, "ジャンプ…", askGoTo, "Ctrl+G");
            },
        });

        // ---- 挿入
        const insertTab = tab("insert", "挿入");
        group(insertTab, "テーブル");
        button("table", { icon: "table", label: "テーブル", large: true, title: "テーブル（Ctrl+T）: データの範囲をテーブルにします", edits: true, act: askTable });
        group(insertTab, "図");
        button("picture", { icon: "picture", label: "画像", large: true, title: "画像: ファイルの画像（PNG・JPEG・GIF・BMP）をアクティブなセルの位置に挿入します（画像の貼り付けでも挿入できます）", edits: true, act: askPicture });
        button("shapes", {
            icon: "shapes",
            label: "図形",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "shapes");
                for (const [title, list] of SHAPE_GROUPS) {
                    head(pop, title);
                    const row = pop.appendChild(el("div", "shaperow"));
                    for (const [kind, name, path] of list) {
                        const b = row.appendChild(el("button", "shapebtn"));
                        b.type = "button";
                        b.title = name;
                        b.dataset.shape = kind;
                        b.innerHTML = `<svg viewBox="0 0 16 16" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round" aria-hidden="true">${path}</svg>`;
                        b.addEventListener("click", () => {
                            closeMenu();
                            closeOpen();
                            startDrawing(kind);
                        });
                    }
                }
            },
        });
        group(insertTab, "テキスト");
        button("textBox", { icon: "textBox", label: "テキスト ボックス", large: true, cls: "two", title: "テキスト ボックス: シートをドラッグして描きます", edits: true, act: () => startDrawing("textBox") });
        group(insertTab, "リンク");
        button("link", { icon: "link", label: "リンク", large: true, title: "リンク（Ctrl+K）: Web ページ・ファイル・ブック内の場所へのハイパーリンクを挿入します", edits: true, act: askLink });
        group(insertTab, "メモ");
        button("note", {
            icon: "note",
            label: "メモ",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "notes");
                const has = !!(sel && noteAt(sel.ar, sel.ac));
                item(pop, has ? "メモの編集" : "新しいメモ", askNote, "Shift+F2");
                item(pop, "メモの削除", removeNote).disabled = !(sel && view && notesIn());
            },
        });
        group(insertTab, "記号と特殊文字");
        button("symbol", { icon: "symbol", label: "記号と特殊文字", large: true, cls: "two", title: "記号と特殊文字: 入力中のセル（またはアクティブなセル）に記号を入力します", edits: true, keepEntry: true, act: askSymbol });

        // ---- 数式
        const formulasTab = tab("formulas", "数式");
        group(formulasTab, "ワークシート分析");
        const f = button("formulas", { icon: "formulas", label: "数式の表示", large: true, title: "セルに値ではなく数式を表示します（Ctrl+` / Ctrl+Shift+@）", act: () => showFormulas(!formulas) });
        f.id = "formulas";
        f.setAttribute("aria-pressed", "false");

        // ---- データ
        const dataTab = tab("data", "データ");
        group(dataTab, "並べ替えとフィルター");
        stack();
        button("sortAsc", { icon: "sortAsc", title: "昇順で並べ替え", edits: true, act: () => sortBy(false) });
        button("sortDesc", { icon: "sortDesc", title: "降順で並べ替え", edits: true, act: () => sortBy(true) });
        into = into.parentNode;
        button("sortMore", { icon: "sort", label: "並べ替え", large: true, title: "ユーザー設定の並べ替え（複数の列で並べ替える）", edits: true, act: askSort });
        button("filter", { icon: "filter", label: "フィルター", large: true, title: "フィルター（Ctrl+Shift+L）", edits: true, act: toggleFilter });
        stack();
        button("filterClear", { icon: "filterClear", label: "クリア", showLabel: true, title: "フィルターの条件をすべてクリア", edits: true, act: clearFilters });

        // ---- 表示
        const viewTab = tab("view", "表示");
        group(viewTab, "ズーム");
        button("zoomBox", { icon: "zoom", label: "ズーム", large: true, menu: (anchor) => openZoomBox(anchor) });
        button("zoom100", { icon: "zoom100", label: "100%", large: true, title: "100% に戻す", act: () => setZoom(100) });
        button("zoomSel", { icon: "zoomSel", label: "選択範囲に合わせて拡大/縮小", large: true, cls: "two", act: zoomToSelection });
        group(viewTab, "ウィンドウ");
        button("freeze", { icon: "freeze", label: "ウィンドウ枠の固定", large: true, cls: "two", edits: true, menu: (anchor) => openFreeze(anchor) });

        // ---- 図形の書式 (shown while a shape, a picture or a chart is selected)
        const shapeTab = tab("shapeFormat", "図形の書式");
        panels.shapeFormat.tab.hidden = true;
        panels.shapeFormat.tab.classList.add("context");
        const onObj = (act) => () => selectedObj && act(selectedObj);
        group(shapeTab, "図形のスタイル");
        const lastShape = { fill: "#4472C4", line: "#000000" };
        const shapeColor = (key, iconName, label, none, more) => {
            const pick = (color) => {
                if (color) {
                    lastShape[key] = color;
                    const bar = tools[key === "fill" ? "shapeFill" : "shapeLine"].querySelector(".bar");
                    if (bar) bar.style.background = color;
                }
                if (selectedObj) objectEdit(selectedObj, { [key]: color });
            };
            const b = button(key === "fill" ? "shapeFill" : "shapeLine", {
                icon: iconName,
                label,
                showLabel: true,
                title: label,
                moreTitle: `${label}の選択`,
                split: true,
                cls: "color",
                act: () => pick(lastShape[key]),
                menu: (anchor) => {
                    openColors(anchor, key, none, pick);
                    if (more) more(menu);
                },
            });
            b.querySelector(".ico").appendChild(el("span", "bar")).style.background = lastShape[key];
        };
        stack();
        shapeColor("fill", "fill", "図形の塗りつぶし", "塗りつぶしなし");
        shapeColor("line", "shapeLine", "図形の枠線", "枠線なし", (pop) => {
            head(pop, "太さ");
            for (const w of LINE_WIDTHS) item(pop, `${w} pt`, onObj((obj) => objectEdit(obj, { lineWidth: w })));
        });
        group(shapeTab, "配置");
        const orderButton = (key, iconName, label, one, all) => {
            button(key, {
                icon: iconName,
                label,
                large: true,
                split: true,
                cls: "two",
                title: label,
                act: onObj((obj) => objectEdit(obj, { order: one })),
                menu: (anchor) => {
                    const pop = openPop(anchor, "orders");
                    for (const [how, text] of ORDERS.filter(([how]) => how === one || how === all)) item(pop, text, onObj((obj) => objectEdit(obj, { order: how })));
                },
            });
        };
        orderButton("forward", "forward", "前面へ移動", "forward", "front");
        orderButton("backward", "backward", "背面へ移動", "backward", "back");
        button("rotate", {
            icon: "rotate",
            label: "回転",
            large: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "rotates");
                item(pop, "右へ 90 度回転", onObj((obj) => turnObject(obj, 90)));
                item(pop, "左へ 90 度回転", onObj((obj) => turnObject(obj, -90)));
                item(pop, "上下反転", onObj((obj) => flipObject(obj, true)));
                item(pop, "左右反転", onObj((obj) => flipObject(obj, false)));
                pop.appendChild(el("hr"));
                item(pop, "その他の回転オプション…", onObj(askRotation));
            },
        });

        // ---- 開発 (where the objects' events open in the sheet's file)
        const dev = tab("dev", "開発");
        // (Where the workbook is edited, for its controls; elsewhere for the events.)
        panels.dev.tab.hidden = !editable;
        group(dev, "コントロール");
        button("controlInsert", {
            icon: "controls",
            label: "挿入",
            large: true,
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "shapes");
                head(pop, "フォーム コントロール");
                const row = pop.appendChild(el("div", "shaperow"));
                for (const [kind, name, path] of FORM_CONTROLS) {
                    const b = row.appendChild(el("button", "shapebtn"));
                    b.type = "button";
                    b.title = `${name}（フォーム コントロール）`;
                    b.dataset.control = kind;
                    b.innerHTML = `<svg viewBox="0 0 16 16" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round" aria-hidden="true">${path}</svg>`;
                    b.addEventListener("click", () => {
                        closeMenu();
                        closeOpen();
                        startDrawing(`control:${kind}`);
                    });
                }
                head(pop, "ActiveX コントロール");
                pop.appendChild(el("div", "pnote", "Excel で挿入してください"));
            },
        });
        const ev = button("event", {
            icon: "code",
            label: "コードの表示",
            large: true,
            title: "選択した図形・コントロールのイベント（マクロ）を開きます。シートのファイルにハンドラーを追加します（図形をダブルクリックすると既定のイベント）",
            menu: (anchor) => {
                const items = eventItems(selectedObj);
                if (!items.length) return;
                const r = anchor.getBoundingClientRect();
                showMenu(r.left, r.bottom + 1, items);
            },
        });
        ev.id = "event";
        ev.disabled = true;

        if (!panels[shown] || panels[shown].tab.hidden) shown = "home";
        show(shown, true);
        ribbon.classList.toggle("collapsed", collapsed);
        ribbon.showShapeTab = (on) => {
            panels.shapeFormat.tab.hidden = !on;
            if (on && saved.ribbonTab === "shapeFormat" && shown !== "shapeFormat") show("shapeFormat", true);
            if (!on && shown === "shapeFormat") show("home", true);
        };
        ribbon.showDev = () => {
            panels.dev.tab.hidden = false;
            if (saved.ribbonTab === "dev") show("dev");
        };
    }
    // The buttons show the format of the active cell.
    function showTools() {
        if (!tools.bold || !view) return;
        const css = activeCss();
        const cell = (activeTd && activeTd.cell) || {};
        for (const name of ["bold", "italic", "underline", "wrap"]) tools[name].classList.toggle("on", STATE[name](css));
        const align = (/text-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["left", "center", "right"]) tools[`align-${a}`].classList.toggle("on", align === a);
        const valign = (/vertical-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["top", "middle", "bottom"]) tools[`valign-${a}`].classList.toggle("on", valign === a);
        const size = (/font-size:([\d.]+)pt/.exec(css) || [])[1] || "11";
        if (![...tools.fontSize.options].some((o) => o.value === size)) tools.fontSize.appendChild(el("option", null, size)).value = size;
        tools.fontSize.value = size;
        if (document.activeElement !== tools.fontName) tools.fontName.value = (/font-family:'([^']*)'/.exec(css) || [])[1] || "游ゴシック";
        const code = cell.n || "General";
        tools.numberFormat.value = NUMBER_FORMATS.some(([c]) => c === code) ? code : "";
        tools.merge.classList.toggle("on", !!sel && single(sel) && mergedIn(sel));
        tools.filter.classList.toggle("on", !!view.sheet.filter);
        tools.filterClear.disabled = !editable || !(view.sheet.filter && Object.keys(view.sheet.filterColumns || {}).length);
        tools.freeze.classList.toggle("on", !!(view.fr || view.fc));
    }
    makeRibbon();

    // ---- the widths of columns and the heights of rows: the edges of the headers

    // How near an edge of a header the mouse takes it (px).
    const EDGE = 4;
    let resizing = null;
    const guide = el("div", "guide");
    // The edge of a header under the mouse: { cols, index } (the column or row before the edge).
    function edgeOf(e) {
        if (!editable || !view || busy) return null;
        const th = e.target.closest && e.target.closest("th");
        if (!th || !view.table.contains(th) || (th.r === undefined && th.c === undefined)) return null;
        const b = th.getBoundingClientRect();
        if (th.c !== undefined) {
            if (b.right - e.clientX <= EDGE) return { cols: true, index: th.c, th };
            const i = view.cols.indexOf(th.c);
            if (e.clientX - b.left <= EDGE && i > 0) return { cols: true, index: view.cols[i - 1], th: view.ths.get(view.cols[i - 1]) };
            return null;
        }
        if (b.bottom - e.clientY <= EDGE) return { cols: false, index: th.r, th };
        const i = view.rows.findIndex((x) => x[0] === th.r);
        if (e.clientY - b.top <= EDGE && i > 0) return { cols: false, index: view.rows[i - 1][0], th: view.rows[i - 1][1].firstChild };
        return null;
    }
    function hoverEdge(e) {
        if (!editable || !view) return;
        const edge = edgeOf(e);
        view.table.classList.toggle("ew", !!edge && edge.cols);
        view.table.classList.toggle("ns", !!edge && !edge.cols);
    }
    // The columns (rows) an edge sizes: those selected whole when it is one of them, else its own.
    function sized(edge) {
        const kind = edge.cols ? "cols" : "rows";
        const [lo, hi] = edge.cols ? [sel && sel.c1, sel && sel.c2] : [sel && sel.r1, sel && sel.r2];
        if (sel && sel.kind === kind && edge.index >= lo && edge.index <= hi) return rangeText(sel);
        return edge.cols ? `${colName(edge.index)}:${colName(edge.index)}` : `${edge.index + 1}:${edge.index + 1}`;
    }
    function beginResize(edge, e) {
        const b = edge.th.getBoundingClientRect();
        resizing = Object.assign(edge, { start: edge.cols ? b.left : b.top, at: edge.cols ? e.clientX : e.clientY, moved: false });
        document.body.appendChild(guide);
        guide.className = edge.cols ? "guide v" : "guide h";
        moveResize(e);
    }
    function moveResize(e) {
        const r = resizing;
        const at = Math.max(r.start, r.cols ? e.clientX : e.clientY);
        if (Math.abs(at - r.at) > 1) r.moved = true;
        r.now = at;
        guide.style.left = r.cols ? `${at}px` : "0";
        guide.style.top = r.cols ? "0" : `${at}px`;
    }
    function endResize() {
        const r = resizing;
        resizing = null;
        guide.remove();
        if (!r.moved || !view) return;
        // (Columns are drawn twice as wide while formulas are shown.)
        const px = Math.round((r.now - r.start) / (r.cols && formulas ? 2 : 1) / Z());
        const sheet = view.sheet.name;
        const range = sized(r);
        postEdit([px <= 0 ? { op: "hide", sheet, range, hidden: true } : { op: "size", sheet, range, px }]);
    }
    // As wide as the widest text of the column's cells, as tall as the tallest content of the row's.
    function autofit(edge) {
        postEdit([{ op: "size", sheet: view.sheet.name, range: sized(edge), px: Math.round(fitPx(edge.cols, edge.index)) }]);
    }
    function fitPx(cols, index) {
        const v = view;
        const sheet = v.sheet;
        const edge = { cols, index };
        let px = 0;
        if (edge.cols) {
            const ctx = document.createElement("canvas").getContext("2d");
            for (const [, tr] of v.rows) {
                for (const td of tr.children) {
                    if (td.c !== edge.index || td.colSpan > 1 || !td.cell || !td.textContent) continue;
                    ctx.font = getComputedStyle(td).font;
                    for (const line of td.textContent.split("\n")) px = Math.max(px, ctx.measureText(line).width);
                }
            }
            // The padding of a cell and its gridline.
            px = px > 0 ? Math.ceil(px) + 8 : sheet.defaultWidth;
        } else {
            const tr = v.trs.get(edge.index);
            const probe = v.wrap.appendChild(el("div", "probe"));
            for (const td of tr ? tr.children : []) {
                if (td.r === undefined || td.rowSpan > 1 || !td.cell || !td.textContent) continue;
                const cs = getComputedStyle(td);
                Object.assign(probe.style, { font: cs.font, whiteSpace: cs.whiteSpace === "pre-wrap" ? "pre-wrap" : "pre", width: `${td.clientWidth}px` });
                probe.textContent = td.textContent;
                px = Math.max(px, probe.offsetHeight + 1);
            }
            probe.remove();
            px = px > 0 ? px : sheet.defaultHeight;
        }
        return px;
    }
    // The width (height) of the columns (rows) selected, asked for in px.
    function askSize(cols) {
        if (!sel || !view) return;
        closeMenu();
        closeDialog();
        const sheet = view.sheet;
        const now = cols ? width(sheet, sel.c1) / (formulas ? 2 : 1) : height(sheet, sel.r1);
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, cols ? "列の幅" : "行の高さ"));
        const line = form.appendChild(el("label"));
        line.appendChild(document.createTextNode(cols ? "幅（ピクセル）" : "高さ（ピクセル）"));
        const input = line.appendChild(el("input"));
        input.type = "number";
        input.min = "0";
        input.max = cols ? "1790" : "545";
        input.step = "1";
        input.value = String(Math.round(now));
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) closeDialog();
        });
        // The whole columns (rows) of the cells selected.
        const range = sel.kind === (cols ? "cols" : "rows") ? rangeText(sel) : cols ? wholeCols() : wholeRows();
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const px = Math.round(+input.value);
            closeDialog();
            if (!(px >= 0)) return;
            postEdit([px === 0 ? { op: "hide", sheet: sheet.name, range, hidden: true } : { op: "size", sheet: sheet.name, range, px }]);
        });
        input.focus();
        input.select();
    }

    // ---- the filter: the buttons of its headings

    const BLANK = "(空白セル)";
    // The button (▾) of a heading of the filter under the mouse: { r, c }.
    function filterButton(e) {
        if (!editable || busy || !view) return null;
        const td = e.target.closest && e.target.closest("td.filter");
        if (!td || !view.table.contains(td)) return null;
        const b = td.getBoundingClientRect();
        return b.right - e.clientX <= 20 ? { r: td.r, c: td.c, td } : null;
    }
    function openFilter(at) {
        const v = view;
        const sheet = v.sheet;
        const [[r1, c1], [r2]] = v.filter;
        const column = colName(at.c);
        const now = (sheet.filterColumns || {})[at.c - c1] || null;
        // The texts the column shows, each once; empty cells last.
        const texts = new Set();
        let blanks = false;
        for (let r = r1 + 1; r <= r2; r++) {
            const cell = v.cells.get(r * KEY + at.c);
            if (cell && cell.t) texts.add(cell.t);
            else blanks = true;
        }
        const list = [...texts].sort((a, b) => a.localeCompare(b, "ja", { numeric: true }));
        const pop = openPop(at.td, "filterpop");
        const sort = (desc) => postEdit([{ op: "sort", sheet: sheet.name, range: sheet.filter, by: column, desc, header: true }]);
        popItem(pop, "昇順", () => sort(false)).dataset.sort = "asc";
        popItem(pop, "降順", () => sort(true)).dataset.sort = "desc";
        if (now) popItem(pop, "フィルターをクリア", () => postEdit([{ op: "filterBy", sheet: sheet.name, column, clear: true }])).dataset.clear = "1";
        // Numbers (dates) for the most part: the number filters, else the text filters.
        let numbers = 0;
        for (let r = r1 + 1; r <= r2; r++) {
            const cell = v.cells.get(r * KEY + at.c);
            if (cell && (cell.k === "number" || cell.k === "date")) numbers++;
        }
        const numeric = numbers * 2 > texts.size && numbers > 0;
        popItem(pop, numeric ? "数値フィルター…" : "テキスト フィルター…", () => askCustomFilter(column, now, list)).dataset.custom = "1";
        if (numeric) popItem(pop, "トップテン…", () => askTop10(column, now)).dataset.top = "1";
        pop.appendChild(el("hr"));
        const form = pop.appendChild(el("form"));
        const checks = [];
        const check = (label, on, value) => {
            const line = form.appendChild(el("label"));
            const box = line.appendChild(el("input"));
            box.type = "checkbox";
            box.checked = on;
            box.value = value === undefined ? "" : value;
            line.appendChild(document.createTextNode(label));
            return box;
        };
        const all = check("(すべて選択)", false);
        all.className = "all";
        const items = form.appendChild(el("div", "values"));
        const shown = (text) => !now || !now.values || now.values.some((x) => x.toLowerCase() === text.toLowerCase());
        const add = (label, on, value) => {
            const line = items.appendChild(el("label"));
            const box = line.appendChild(el("input"));
            box.type = "checkbox";
            box.checked = on;
            line.appendChild(document.createTextNode(label));
            checks.push([box, value]);
        };
        for (const text of list) add(text, shown(text), text);
        if (blanks) add(BLANK, !now || !now.values || !!now.blanks, null);
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", () => {
            closeMenu();
            focusGrid();
        });
        const sync = () => {
            const n = checks.filter(([b]) => b.checked).length;
            all.checked = n === checks.length;
            all.indeterminate = n > 0 && n < checks.length;
            ok.disabled = n === 0;
        };
        all.addEventListener("change", () => {
            for (const [b] of checks) b.checked = all.checked;
            sync();
        });
        items.addEventListener("change", sync);
        sync();
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const on = checks.filter(([b]) => b.checked);
            closeMenu();
            focusGrid();
            if (!on.length) return;
            if (on.length === checks.length) {
                if (now) postEdit([{ op: "filterBy", sheet: sheet.name, column, clear: true }]);
                return;
            }
            const op = { op: "filterBy", sheet: sheet.name, column, values: on.filter(([, value]) => value !== null).map(([, value]) => value) };
            if (on.some(([, value]) => value === null)) op.blanks = true;
            postEdit([op]);
        });
    }

    // Excel's "オートフィルター オプション": one or two comparisons, by what they are called.
    const COMPARISONS = [
        ["eq", "と等しい"],
        ["ne", "と等しくない"],
        ["gt", "より大きい"],
        ["ge", "以上"],
        ["lt", "より小さい"],
        ["le", "以下"],
        ["begins", "で始まる"],
        ["notBegins", "で始まらない"],
        ["ends", "で終わる"],
        ["notEnds", "で終わらない"],
        ["contains", "を含む"],
        ["notContains", "を含まない"],
    ];
    // A comparison as the filter keeps it ([op, value]), and back.
    function comparisonOf(kind, value) {
        const ops = { eq: "=", ne: "<>", gt: ">", ge: ">=", lt: "<", le: "<=" };
        if (ops[kind]) return [ops[kind], value];
        const not = kind.startsWith("not");
        const k = not ? kind.slice(3).toLowerCase() : kind;
        const v = k === "begins" ? `${value}*` : k === "ends" ? `*${value}` : `*${value}*`;
        return [not ? "<>" : "=", v];
    }
    function kindOf([op, value]) {
        const wild = op === "=" || op === "<>";
        const not = op === "<>" ? "not" : "";
        const cap = (s) => (not ? s[0].toUpperCase() + s.slice(1) : s);
        if (wild && value.length > 1 && value.startsWith("*") && value.endsWith("*")) return [not + cap("contains"), value.slice(1, -1)];
        if (wild && value.length > 0 && value.endsWith("*") && !value.endsWith("~*")) return [not + cap("begins"), value.slice(0, -1)];
        if (wild && value.length > 0 && value.startsWith("*")) return [not + cap("ends"), value.slice(1)];
        return [{ "=": "eq", "<>": "ne", ">": "gt", ">=": "ge", "<": "lt", "<=": "le" }[op] || "eq", value];
    }
    function askCustomFilter(column, now, values) {
        const sheet = view.sheet.name;
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg customfilter"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "オートフィルター オプション"));
        form.appendChild(el("div", "lbl", `抽出条件の指定: 列 ${column}`));
        const list = form.appendChild(el("datalist"));
        list.id = "filtervalues";
        for (const v of values.slice(0, 500)) list.appendChild(el("option")).value = v;
        const had = now && now.custom ? [now.custom.a, now.custom.b].filter(Boolean).map(kindOf) : [];
        const line = (i) => {
            const row = form.appendChild(el("div", "cond"));
            const input = row.appendChild(el("input"));
            input.type = "text";
            input.name = `v${i}`;
            input.setAttribute("list", "filtervalues");
            input.value = had[i] ? had[i][1] : "";
            const kind = row.appendChild(el("select"));
            kind.name = `k${i}`;
            if (i) kind.appendChild(el("option", null, "")).value = "";
            for (const [value, label] of COMPARISONS) kind.appendChild(el("option", null, label)).value = value;
            kind.value = had[i] ? had[i][0] : i ? "" : "eq";
            return input;
        };
        const first = line(0);
        const join = form.appendChild(el("div", "join"));
        for (const [value, label] of [["and", "AND"], ["or", "OR"]]) {
            const l = join.appendChild(el("label"));
            const radio = l.appendChild(el("input"));
            radio.type = "radio";
            radio.name = "join";
            radio.value = value;
            radio.checked = value === (now && now.custom && now.custom.and === false ? "or" : "and");
            l.appendChild(document.createTextNode(label));
        }
        line(1);
        form.appendChild(el("div", "hint", "? を任意の 1 文字、* を任意の文字列の代わりに使えます。"));
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const f = new FormData(form);
            const criteria = [];
            for (const i of [0, 1]) {
                const kind = f.get(`k${i}`);
                if (kind) criteria.push(comparisonOf(kind, String(f.get(`v${i}`) || "")));
            }
            closeDialog();
            if (criteria.length) postEdit([{ op: "filterBy", sheet, column, criteria, and: f.get("join") !== "or" }]);
        });
        first.focus();
    }
    function askTop10(column, now) {
        const sheet = view.sheet.name;
        const had = (now && now.top) || { bottom: false, percent: false, count: 10 };
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg top10"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "トップテン オートフィルター"));
        const row = form.appendChild(el("div", "cond"));
        const which = row.appendChild(el("select"));
        which.name = "which";
        for (const [value, label] of [["top", "上位"], ["bottom", "下位"]]) which.appendChild(el("option", null, label)).value = value;
        which.value = had.bottom ? "bottom" : "top";
        const count = row.appendChild(el("input"));
        count.type = "number";
        count.name = "count";
        count.min = "1";
        count.max = "500";
        count.value = String(had.count);
        const unit = row.appendChild(el("select"));
        unit.name = "unit";
        for (const [value, label] of [["items", "項目"], ["percent", "パーセント"]]) unit.appendChild(el("option", null, label)).value = value;
        unit.value = had.percent ? "percent" : "items";
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const n = Math.round(+count.value);
            const percent = unit.value === "percent";
            if (!(n >= 1 && n <= (percent ? 100 : 500))) {
                count.setCustomValidity(percent ? "1〜100 を指定してください" : "1〜500 を指定してください");
                count.reportValidity();
                return;
            }
            closeDialog();
            postEdit([{ op: "filterBy", sheet, column, [which.value]: n, percent }]);
        });
        count.addEventListener("input", () => count.setCustomValidity(""));
        count.focus();
        count.select();
    }

    // ---- "並べ替え": several keys, as Excel's dialog

    // The data around a cell, as Excel's "アクティブ セル領域": grown while cells next to it
    // (corners too) have content.
    function currentRegion(r, c) {
        const has = (a, b) => a >= 0 && b >= 0 && view.filled.has(a * KEY + b);
        const rowHas = (row, a, b) => {
            for (let x = Math.max(0, a); x <= b; x++) if (has(row, x)) return true;
            return false;
        };
        const colHas = (col, a, b) => {
            for (let y = Math.max(0, a); y <= b; y++) if (has(y, col)) return true;
            return false;
        };
        let [r1, c1, r2, c2] = [r, c, r, c];
        for (let grew = true; grew; ) {
            grew = false;
            if (r1 > 0 && rowHas(r1 - 1, c1 - 1, c2 + 1)) [r1, grew] = [r1 - 1, true];
            if (rowHas(r2 + 1, c1 - 1, c2 + 1)) [r2, grew] = [r2 + 1, true];
            if (c1 > 0 && colHas(c1 - 1, r1 - 1, r2 + 1)) [c1, grew] = [c1 - 1, true];
            if (colHas(c2 + 1, r1 - 1, r2 + 1)) [c2, grew] = [c2 + 1, true];
        }
        return [r1, c1, r2, c2];
    }
    function askSort() {
        if (!editable || busy || !sel || !view || sel.kind === "all" || !oneArea()) return;
        if (entry) commit(null);
        const v = view;
        const sheetName = v.sheet.name;
        // What is sorted: the cells selected, or the filter's range, or the data around the cell.
        let area;
        if (!single(sel) && sel.kind === "cells") area = [sel.r1, sel.c1, sel.r2, sel.c2];
        else if (v.filter && sel.ar >= v.filter[0][0] && sel.ar <= v.filter[1][0] && sel.ac >= v.filter[0][1] && sel.ac <= v.filter[1][1]) area = [v.filter[0][0], v.filter[0][1], v.filter[1][0], v.filter[1][1]];
        else area = currentRegion(sel.ar, sel.ac);
        const [r1, c1, r2, c2] = area;
        if (r2 - r1 < 1 && !v.filled.has(r1 * KEY + c1)) return;
        setSel(areaSel(r1, c1, r2, c2), false);
        // Headings: a first row of texts (as Excel guesses it).
        const firstTexts = [];
        let texts = 0;
        let other = false;
        for (let c = c1; c <= c2; c++) {
            const cell = v.cells.get(r1 * KEY + c);
            firstTexts.push((cell && cell.t) || "");
            if (cell && cell.t) {
                if (cell.k === "string") texts++;
                else other = true;
            }
        }
        const guess = texts > 0 && !other;
        closeMenu();
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg sortdlg"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "並べ替え"));
        const tools = form.appendChild(el("div", "sorttools"));
        const add = tools.appendChild(el("button", null, "レベルの追加"));
        add.type = "button";
        const del = tools.appendChild(el("button", null, "レベルの削除"));
        del.type = "button";
        const headLine = tools.appendChild(el("label"));
        const head = headLine.appendChild(el("input"));
        head.type = "checkbox";
        head.checked = guess && r2 > r1;
        headLine.appendChild(document.createTextNode("先頭行をデータの見出しとして使用する"));
        const levels = form.appendChild(el("div", "levels"));
        const columnLabel = (c) => (head.checked && firstTexts[c - c1] ? firstTexts[c - c1] : `列 ${colName(c)}`);
        let chosen = null;
        const level = (c, desc) => {
            const row = levels.appendChild(el("div", "level"));
            row.appendChild(el("span", "lbl", levels.children.length === 1 ? "最優先されるキー" : "次に優先されるキー"));
            const col = row.appendChild(el("select"));
            col.className = "col";
            for (let x = c1; x <= c2; x++) col.appendChild(el("option", null, columnLabel(x))).value = colName(x);
            col.value = colName(c);
            const order = row.appendChild(el("select"));
            order.className = "order";
            for (const [value, label] of [["asc", "昇順"], ["desc", "降順"]]) order.appendChild(el("option", null, label)).value = value;
            order.value = desc ? "desc" : "asc";
            row.addEventListener("mousedown", () => {
                if (chosen) chosen.classList.remove("on");
                chosen = row;
                row.classList.add("on");
            });
            return row;
        };
        level(sel.ac >= c1 && sel.ac <= c2 ? sel.ac : c1, false);
        head.addEventListener("change", () => {
            for (const s of levels.querySelectorAll("select.col")) for (const o of s.options) o.textContent = columnLabel(colNumber(o.value));
        });
        add.addEventListener("click", () => {
            if (levels.children.length < 64) level(c1, false);
        });
        del.addEventListener("click", () => {
            const row = chosen || levels.lastChild;
            if (row && levels.children.length > 1) row.remove();
            chosen = null;
            levels.querySelectorAll(".level .lbl").forEach((l, i) => (l.textContent = i ? "次に優先されるキー" : "最優先されるキー"));
        });
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const keys = [...levels.children].map((row) => {
                const k = { by: row.querySelector("select.col").value };
                if (row.querySelector("select.order").value === "desc") k.desc = true;
                return k;
            });
            closeDialog();
            const range = r1 === r2 && c1 === c2 ? address(r1, c1) : `${address(r1, c1)}:${address(r2, c2)}`;
            postEdit([{ op: "sort", sheet: sheetName, range, keys, header: head.checked }]);
        });
        levels.querySelector("select").focus();
    }

    // ---- the selection's border dragged: the cells moved there (Ctrl: copied), as in Excel

    let moving = null;
    // The pointer on the border of the cells selected (not on the fill handle).
    function onBorder(e) {
        if (!editable || busy || !view || !sel || areas.length || selected || entry || sel.kind !== "cells") return false;
        const tl = tdAt(sel.r1, sel.c1);
        const br = tdAt(sel.r2, sel.c2);
        if (!tl || !br) return false;
        const [a, b] = [tl.getBoundingClientRect(), br.getBoundingClientRect()];
        const [x, y] = [e.clientX, e.clientY];
        const near = 3;
        if (x < a.left - near || x > b.right + near || y < a.top - near || y > b.bottom + near) return false;
        // The fill handle's corner is the fill handle's.
        if (x > b.right - 8 && y > b.bottom - 8) return false;
        return x <= a.left + near || x >= b.right - near || y <= a.top + near || y >= b.bottom - near;
    }
    function beginMove(e, hit) {
        moving = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, dr: hit.r - sel.r1, dc: hit.c - sel.c1, to: [sel.r1, sel.c1], copy: false };
        moveTo(e);
    }
    function moveTo(e) {
        const m = moving;
        const hit = hitOf(document.elementFromPoint(e.clientX, e.clientY));
        if (hit && hit.kind === "cells") {
            const s = view.sheet;
            m.to = [Math.max(0, Math.min(s.maxRows - 1 - (m.r2 - m.r1), hit.r - m.dr)), Math.max(0, Math.min(s.maxCols - 1 - (m.c2 - m.c1), hit.c - m.dc))];
        }
        m.copy = e.ctrlKey || e.metaKey;
        const b = boxOf(m.to[0], m.to[1], m.to[0] + m.r2 - m.r1, m.to[1] + m.c2 - m.c1);
        const box = view.fillbox;
        box.hidden = !b;
        box.classList.add("moving");
        if (b) Object.assign(box.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w + 1}px`, height: `${b.h + 1}px` });
        pinTo(box, pinOf(view, b));
        document.body.classList.toggle("copying", m.copy);
    }
    function endMove() {
        const m = moving;
        moving = null;
        document.body.classList.remove("copying");
        if (!view) return;
        view.fillbox.hidden = true;
        view.fillbox.classList.remove("moving");
        const [r, c] = m.to;
        if (r === m.r1 && c === m.c1) return;
        const sheet = view.sheet.name;
        const from = m.r1 === m.r2 && m.c1 === m.c2 ? address(m.r1, m.c1) : `${address(m.r1, m.c1)}:${address(m.r2, m.c2)}`;
        const [h, w] = [m.r2 - m.r1, m.c2 - m.c1];
        const go = () => {
            setSel(areaSel(r, c, r + h, c + w), false);
            postEdit([{ op: m.copy ? "copy" : "move", sheet, from, fromSheet: sheet, to: address(r, c) }]);
        };
        // Moved over cells with content (outside the cells moved): asked first, as Excel asks.
        let filled = false;
        for (let y = r; y <= r + h && !filled; y++)
            for (let x = c; x <= c + w; x++) {
                const own = y >= m.r1 && y <= m.r2 && x >= m.c1 && x <= m.c2;
                if (!own && view.filled.has(y * KEY + x)) {
                    filled = true;
                    break;
                }
            }
        if (filled && !m.copy) askReplace(go);
        else go();
    }
    function askReplace(then) {
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("div", "msg", "ここには既にデータがあります。置き換えますか?"));
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
            then();
        });
        ok.focus();
    }

    // ---- the fill handle: the cells selected, drawn out over the cells next to them

    let filling = null;
    // The cells the handle is drawn out to: [r1, c1, r2, c2], along the way the mouse went further.
    function fillTarget(e) {
        const f = filling;
        const hit = hitOf(document.elementFromPoint(e.clientX, e.clientY));
        if (!hit || hit.kind !== "cells") return f.to;
        const down = Math.abs(e.clientY - f.y) >= Math.abs(e.clientX - f.x);
        if (down) return [Math.min(f.r1, hit.r), f.c1, Math.max(f.r2, hit.r), f.c2];
        return [f.r1, Math.min(f.c1, hit.c), f.r2, Math.max(f.c2, hit.c)];
    }
    function moveFill(e) {
        const f = filling;
        f.to = fillTarget(e);
        const b = boxOf(f.to[0], f.to[1], f.to[2], f.to[3]);
        view.fillbox.hidden = !b;
        if (b) Object.assign(view.fillbox.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w + 1}px`, height: `${b.h + 1}px` });
    }
    function autofill(f, to) {
        if (to[0] === f.r1 && to[1] === f.c1 && to[2] === f.r2 && to[3] === f.c2) return;
        const text = (r1, c1, r2, c2) => (r1 === r2 && c1 === c2 ? address(r1, c1) : `${address(r1, c1)}:${address(r2, c2)}`);
        setSel(areaSel(to[0], to[1], to[2], to[3]), false);
        postEdit([{ op: "autofill", sheet: view.sheet.name, from: text(f.r1, f.c1, f.r2, f.c2), to: text(to[0], to[1], to[2], to[3]) }]);
    }
    function endFill() {
        const f = filling;
        filling = null;
        if (!view) return;
        view.fillbox.hidden = true;
        autofill(f, f.to);
    }
    $("grid").addEventListener("mousedown", (e) => {
        if (!view || e.target !== view.dropbtn || e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        if (menu && menu.classList.contains("listpop")) closeMenu();
        else openList();
    });
    $("grid").addEventListener("mousedown", (e) => {
        if (!view || e.target !== view.fillh || e.button !== 0 || busy || !sel) return;
        e.preventDefault();
        e.stopPropagation();
        closeMenu();
        filling = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2, x: e.clientX, y: e.clientY, to: [sel.r1, sel.c1, sel.r2, sel.c2] };
    });
    // Double-click: down as far as the column next to the cells has content.
    $("grid").addEventListener("dblclick", (e) => {
        if (!view || e.target !== view.fillh || busy || !sel) return;
        e.stopPropagation();
        const v = view;
        const has = (r, c) => c >= 0 && v.filled.has(r * KEY + c);
        const beside = [sel.c1 - 1, sel.c2 + 1].find((c) => has(sel.r2 + 1, c));
        if (beside === undefined) return;
        let r = sel.r2;
        while (has(r + 1, beside)) r++;
        autofill(sel, [sel.r1, sel.c1, r, sel.c2]);
    });

    // ---- the name box and "ジャンプ" (F5, Ctrl+G): a cell, a range or a name to go to

    const sheetByName = (name) => data.sheets.find((s) => s.name === name) || data.sheets.find((s) => s.name.toLowerCase() === String(name).toLowerCase());
    // The names a sheet sees: its own, then the workbook's.
    function namesOf(sheet) {
        const all = data.names || [];
        const own = all.filter((n) => n.sheet === sheet.name);
        return own.concat(all.filter((n) => !n.sheet && !own.some((o) => o.name.toLowerCase() === n.name.toLowerCase())));
    }
    function fillNames(sheet) {
        const list = $("names");
        list.textContent = "";
        for (const n of namesOf(sheet)) list.appendChild(el("option")).value = n.name;
    }
    // A reference as it is typed: { sheet, kind, r1, c1, r2, c2 } or null.
    function parseRef(text, sheet) {
        let t = String(text).trim().replace(/^=/, "");
        if (!t) return null;
        const named = namesOf(sheet).find((n) => n.name.toLowerCase() === t.toLowerCase());
        // (A name with several areas: its first.)
        if (named) t = named.ref.split(",")[0];
        let on = sheet;
        const bang = t.lastIndexOf("!");
        if (bang >= 0) {
            on = sheetByName(t.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"));
            if (!on) return null;
            t = t.slice(bang + 1);
        }
        t = t.replace(/\$/g, "").toUpperCase();
        const ok = (r, c) => r >= 0 && r < MAXR && c >= 0 && c < MAXC;
        let m;
        if ((m = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(t))) {
            const [r1, c1] = [+m[2] - 1, colNumber(m[1])];
            const [r2, c2] = m[3] ? [+m[4] - 1, colNumber(m[3])] : [r1, c1];
            if (!ok(r1, c1) || !ok(r2, c2)) return null;
            return { sheet: on, kind: "cells", r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
        }
        if ((m = /^(\d{1,7}):(\d{1,7})$/.exec(t))) {
            const [a, b] = [+m[1] - 1, +m[2] - 1];
            if (!ok(a, 0) || !ok(b, 0)) return null;
            return { sheet: on, kind: "rows", r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAXC - 1 };
        }
        if ((m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(t))) {
            const [a, b] = [colNumber(m[1]), colNumber(m[2])];
            if (!ok(0, a) || !ok(0, b)) return null;
            return { sheet: on, kind: "cols", r1: 0, c1: Math.min(a, b), r2: MAXR - 1, c2: Math.max(a, b) };
        }
        return null;
    }
    // Selects what a reference points at (on its sheet): false when it points at nothing.
    function goTo(text) {
        if (!view) return false;
        const ref = parseRef(text, view.sheet);
        if (!ref) return false;
        if (entry) commit(null);
        if (ref.sheet !== view.sheet) draw(ref.sheet);
        if (ref.kind === "rows") setSel({ kind: "rows", nr: ref.r1, nc: 0, fr: ref.r2, fc: 0 }, true);
        else if (ref.kind === "cols") setSel({ kind: "cols", nr: 0, nc: ref.c1, fr: 0, fc: ref.c2 }, true);
        else setSel(areaSel(ref.r1, ref.c1, ref.r2, ref.c2), true);
        return true;
    }
    // The name box: the active cell, or the name of the cells selected when they have one.
    function showName(sheet, r, c) {
        const box = $("name");
        if (document.activeElement === box) return;
        let shown = address(r, c);
        if (sel && !single(sel)) {
            const here = rangeText(sel);
            const named = namesOf(sheet).find((n) => {
                const ref = parseRef(n.ref, sheet);
                return ref && ref.sheet === sheet && ref.kind === sel.kind && rangeText(ref) === here;
            });
            if (named) shown = named.name;
        }
        box.value = shown;
    }
    const nameBox = $("name");
    nameBox.addEventListener("focus", () => nameBox.select());
    nameBox.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            e.stopPropagation();
            const text = nameBox.value;
            if (!text.trim()) return;
            nameBox.blur();
            if (!goTo(text)) showError(`参照 \`${text}\` が正しくありません（A1、B2:D5、3:5、B:D、Sheet2!A1、定義された名前で指定してください）。`);
            focusGrid();
        } else if (e.key === "Escape") {
            e.preventDefault();
            e.stopPropagation();
            nameBox.blur();
            focusGrid();
        }
    });
    nameBox.addEventListener("blur", () => {
        if (view && sel) showName(view.sheet, sel.ar, sel.ac);
    });

    // Excel's "ジャンプ": the names to go to, or a reference.
    function askGoTo() {
        if (!view) return;
        if (entry) commit(null);
        closeMenu();
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg goto"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "ジャンプ"));
        form.appendChild(el("div", "lbl", "移動先:"));
        const list = form.appendChild(el("select"));
        list.size = 8;
        for (const n of namesOf(view.sheet)) list.appendChild(el("option", null, n.name)).value = n.name;
        const line = form.appendChild(el("label"));
        line.appendChild(document.createTextNode("参照先:"));
        const input = line.appendChild(el("input"));
        input.type = "text";
        input.spellcheck = false;
        list.addEventListener("change", () => (input.value = list.value));
        list.addEventListener("dblclick", () => form.requestSubmit());
        const buttons = form.appendChild(el("div", "buttons"));
        buttons.appendChild(el("button", "ok", "OK")).type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", closeDialog);
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) closeDialog();
        });
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const text = input.value.trim();
            if (!text) return;
            if (!parseRef(text, view.sheet)) {
                input.setCustomValidity("参照が正しくありません");
                input.reportValidity();
                return;
            }
            closeDialog();
            goTo(text);
        });
        input.addEventListener("input", () => input.setCustomValidity(""));
        input.focus();
    }

    // ---- Ctrl+; and Ctrl+:: today's date and the time, as Japanese Excel types them

    function nowText(time) {
        const d = new Date();
        return time ? `${d.getHours()}:${String(d.getMinutes()).padStart(2, "0")}` : `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
    }
    function typeNow(time) {
        if (!editable || !sel || sel.kind !== "cells") return;
        beginEntry("enter", nowText(time), "cell");
    }

    // ---- the status bar: what is being done, and the sum, the average and the count of the
    // cells selected (those shown), as Excel's

    function showMode() {
        $("mode").textContent = !entry ? "準備完了" : entry.point ? "参照" : entry.mode === "edit" ? "編集" : "入力";
    }
    // A number as the status bar shows it: in the number format of the active cell, as Excel
    // 2021 shows it (General in at most 11 characters, without separators).
    function statText(x, code) {
        const shown = code ? formatNumber(x, code) : null;
        return shown !== null ? shown : generalFit(x, () => true) || String(x);
    }
    function showStats() {
        const box = $("stats");
        box.textContent = "";
        if (!view || !sel || selected || (single(sel) && !areas.length)) return;
        const s = view.sheet;
        const all = allAreas();
        let count = 0,
            nums = 0,
            sum = 0,
            errors = false;
        for (const c of s.cells) {
            if (!all.some((a) => c.r >= a.r1 && c.r <= a.r2 && c.c >= a.c1 && c.c <= a.c2)) continue;
            if (c.t === undefined || (c.t === "" && !c.f)) continue;
            // Hidden rows and columns (filtered out) are not counted.
            if (!(height(s, c.r) > 0) || !(width(s, c.c) > 0)) continue;
            count++;
            if (c.k === "error") errors = true;
            if (typeof c.x === "number" && (c.k === "number" || c.k === "date")) {
                nums++;
                sum += c.x;
            }
        }
        if (!count) return;
        const active = view.cells.get(sel.ar * KEY + sel.ac);
        const code = active && (active.sn || active.n);
        const add = (label, value) => box.appendChild(el("span", null, `${label}: ${value}`));
        // With an error value among the cells, Excel shows only how many there are.
        const stats = nums && !errors;
        if (stats) add("平均", statText(sum / nums, code));
        add("データの個数", String(count));
        if (stats) add("合計", statText(sum, code));
    }

    // ---- number formats (the status bar; the cells come already shown): the codes for numbers,
    // percentages, scientific notation, dates and times. null for what is not read here
    // (fractions, conditions, Japanese eras), which is then shown as General.

    const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    function formatNumber(x, code) {
        if (/^general$/i.test(code)) return null;
        // The sections, split at `;` outside quotes, escapes and brackets.
        const secs = [];
        let part = "";
        for (let i = 0; i < code.length; i++) {
            const ch = code[i];
            if (ch === '"') {
                const j = code.indexOf('"', i + 1);
                const end = j < 0 ? code.length : j;
                part += code.slice(i, end + 1);
                i = end;
            } else if (ch === "\\" || ch === "_" || ch === "*") {
                part += code.slice(i, i + 2);
                i++;
            } else if (ch === "[") {
                const j = code.indexOf("]", i);
                const end = j < 0 ? code.length : j;
                part += code.slice(i, end + 1);
                i = end;
            } else if (ch === ";") {
                secs.push(part);
                part = "";
            } else part += ch;
        }
        secs.push(part);
        if (secs.some((s) => /\[[<>=]/.test(s))) return null;
        let s = secs[0],
            sign = x < 0 ? "-" : "";
        if (x < 0 && secs.length > 1) {
            s = secs[1];
            sign = "";
        } else if (x === 0 && secs.length > 2) return null; // (Excel shows 0, not the zero's section.)
        const toks = tokensOf(s);
        if (!toks) return null;
        if (toks.some((t) => t.k === "date" || t.k === "ampm")) {
            // (Not fractions of a second, nor dates before 1900.)
            if (x < 0 || toks.some((t) => t.k === "ph")) return null;
            // Times (without a day) are shown as hours on end: 32:24:00.
            const day = toks.some((t) => t.k === "date" && (/^[yda]/.test(t.v) || /^m{3,}$/.test(t.v)));
            const time = toks.some((t) => t.k === "ampm" || (t.k === "date" && /^[hs]/.test(t.v)));
            return formatDate(x, time && !day ? tokensOf("[h]:mm:ss") : toks);
        }
        const out = formatDigits(Math.abs(x), toks);
        return out === null ? null : sign + out;
    }
    // A section's parts: literal text, digit places (0 # ?), the point, commas, %, E+, date parts.
    function tokensOf(s) {
        const toks = [];
        const lit = (v) => toks.push({ k: "lit", v });
        for (let i = 0; i < s.length; ) {
            const ch = s[i];
            const rest = s.slice(i);
            let m;
            if (ch === '"') {
                const j = s.indexOf('"', i + 1);
                const end = j < 0 ? s.length : j;
                lit(s.slice(i + 1, end));
                i = end + 1;
            } else if (ch === "\\") {
                lit(s[i + 1] || "");
                i += 2;
            } else if (ch === "_") {
                lit(" ");
                i += 2;
            } else if (ch === "*") i += 2;
            else if (ch === "[") {
                const j = s.indexOf("]", i);
                const inner = s.slice(i + 1, j < 0 ? s.length : j);
                i = j < 0 ? s.length : j + 1;
                if ((m = /^\$([^-]*)(?:-(.*))?$/.exec(inner))) {
                    // A currency symbol, and a language (`[$-411]`); other calendars and the
                    // system's date formats (`[$-F800]`) are not read here.
                    if (m[2] !== undefined && !/^[0-9a-f]{1,4}$/i.test(m[2])) return null;
                    if (/^f[0-9a-f]00$/i.test(m[2] || "")) return null;
                    if (m[1]) lit(m[1]);
                } else if (/^(h+|m+|s+)$/i.test(inner)) toks.push({ k: "date", v: inner.toLowerCase(), elapsed: true });
                else if (/^dbnum/i.test(inner)) return null;
                // Colors are not shown in the status bar.
            } else if ("0#?".includes(ch)) {
                toks.push({ k: "ph", v: ch });
                i++;
            } else if (ch === ".") {
                toks.push({ k: "dot" });
                i++;
            } else if (ch === ",") {
                toks.push({ k: "comma" });
                i++;
            } else if (ch === "%") {
                toks.push({ k: "pct" });
                i++;
            } else if ((ch === "E" || ch === "e") && (s[i + 1] === "+" || s[i + 1] === "-")) {
                toks.push({ k: "exp", v: ch + s[i + 1] });
                i += 2;
            } else if ((m = /^(am\/pm|a\/p)/i.exec(rest))) {
                toks.push({ k: "ampm", v: m[1] });
                i += m[1].length;
            } else if ((m = /^(y+|m+|d+|h+|s+|a{3,4})/i.exec(rest))) {
                toks.push({ k: "date", v: m[1].toLowerCase() });
                i += m[1].length;
            } else if (/[gGeEbB]/.test(ch)) return null; // Japanese eras, Buddhist years
            else if (ch === "/" && toks.some((t) => t.k === "ph")) return null; // fractions
            else {
                lit(ch === "@" ? "" : ch);
                i++;
            }
        }
        return toks;
    }
    // x (not negative) with d decimals, rounded half up as written in decimals (16.15 → 16.2).
    function decimal(x, d) {
        const [m, e] = Number(x.toPrecision(15)).toExponential().split("e");
        return (Math.round(Number(`${m}e${Number(e) + d}`)) / 10 ** d).toFixed(d);
    }
    function formatDigits(x, toks) {
        x *= 100 ** toks.filter((t) => t.k === "pct").length;
        const expAt = toks.findIndex((t) => t.k === "exp");
        const end = expAt < 0 ? toks.length : expAt;
        const dotAt = toks.findIndex((t, i) => t.k === "dot" && i < end);
        const intEnd = dotAt < 0 ? end : dotAt;
        const ints = [],
            decs = [];
        toks.forEach((t, i) => {
            if (t.k === "ph" && i < intEnd) ints.push(i);
            else if (t.k === "ph" && i > dotAt && dotAt >= 0 && i < end) decs.push(i);
        });
        // Commas right after the last digit place divide by 1000; between digit places they group.
        let scale = 0;
        const last = Math.max(-1, ...ints, ...decs);
        if (last >= 0) for (let i = last + 1; i < end && toks[i].k === "comma"; i++) scale++;
        const group = ints.length > 1 && toks.some((t, i) => t.k === "comma" && i > ints[0] && i < ints[ints.length - 1]);
        // Text between the integer's digit places (000-0000) is not read here.
        if (ints.length && toks.some((t, i) => t.k === "lit" && t.v && i > ints[0] && i < ints[ints.length - 1])) return null;
        x /= 1000 ** scale;
        let exp = "";
        if (expAt >= 0) {
            if (ints.length > 1) return null;
            const digits = toks.slice(expAt + 1).filter((t) => t.k === "ph");
            let n = x === 0 ? 0 : Math.floor(Math.log10(Number(x.toPrecision(15))));
            if (Number(decimal(x / 10 ** n, decs.length)) >= 10) n++;
            x /= 10 ** n;
            exp = toks[expAt].v[0] + (n < 0 ? "-" : toks[expAt].v[1] === "+" ? "+" : "") + String(Math.abs(n)).padStart(digits.filter((t) => t.v === "0").length, "0");
        }
        let [ip, fp = ""] = decimal(x, decs.length).split(".");
        // Decimal places # and ? show no trailing zeros (? as spaces).
        const frac = fp.split("");
        for (let i = decs.length - 1; i >= 0 && frac[i] === "0" && toks[decs[i]].v !== "0"; i--) frac[i] = toks[decs[i]].v === "?" ? " " : "";
        if (ip === "0") ip = "";
        ip = ip.padStart(ints.filter((i) => toks[i].v === "0").length, "0");
        if (group) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        let out = "";
        toks.forEach((t, i) => {
            if (t.k === "lit") out += t.v;
            else if (t.k === "pct") out += "%";
            // (Without integer places, the integer goes before the point: `.00`.)
            else if (t.k === "dot" && i === dotAt) out += (ints.length ? "" : ip) + ".";
            else if (t.k === "exp") out += exp;
            else if (t.k === "ph" && i === ints[0]) out += ip;
            else if (t.k === "ph" && i === decs[0]) out += frac.join("");
        });
        return out;
    }
    function formatDate(x, toks) {
        let days = Math.floor(x);
        let secs = Math.round((x - days) * 86400);
        if (secs >= 86400) {
            days++;
            secs -= 86400;
        }
        const d = new Date(Date.UTC(1899, 11, 30) + days * 864e5);
        const ampm = toks.some((t) => t.k === "ampm");
        const h = Math.floor(secs / 3600),
            mi = Math.floor(secs / 60) % 60,
            sec = secs % 60;
        const two = (n, w) => String(n).padStart(w, "0");
        let out = "";
        toks.forEach((t, i) => {
            if (t.k === "lit") out += t.v;
            else if (t.k === "ampm") out += t.v.toLowerCase().startsWith("am") ? (h < 12 ? "AM" : "PM") : h < 12 ? "A" : "P";
            else if (t.k === "dot") out += ".";
            else if (t.k === "date") {
                const v = t.v;
                // m and mm are minutes after hours or before seconds.
                const prev = toks.slice(0, i).reverse().find((x) => x.k === "date") || {};
                const next = toks.slice(i + 1).find((x) => x.k === "date") || {};
                const minutes = v[0] === "m" && v.length <= 2 && (/^h/.test(prev.v || "") || /^s/.test(next.v || ""));
                if (t.elapsed) {
                    const total = v[0] === "h" ? Math.floor(x * 24) : v[0] === "m" ? Math.floor(x * 1440) : Math.round(x * 86400);
                    out += two(total, v.length);
                } else if (v[0] === "y") out += v.length > 2 ? String(d.getUTCFullYear()) : two(d.getUTCFullYear() % 100, 2);
                else if (minutes) out += two(mi, v.length);
                else if (v[0] === "m") {
                    const m = d.getUTCMonth();
                    out += v.length === 1 ? String(m + 1) : v.length === 2 ? two(m + 1, 2) : v.length === 3 ? MONTHS[m].slice(0, 3) : v.length === 5 ? MONTHS[m][0] : MONTHS[m];
                } else if (v[0] === "d") {
                    const w = d.getUTCDay();
                    out += v.length <= 2 ? two(d.getUTCDate(), v.length) : v.length === 3 ? DAYS[w].slice(0, 3) : DAYS[w];
                } else if (v[0] === "a") out += "日月火水木金土"[d.getUTCDay()] + (v.length === 4 ? "曜日" : "");
                else if (v[0] === "h") out += two(ampm ? h % 12 || 12 : h, Math.min(v.length, 2));
                else if (v[0] === "s") out += two(sec, Math.min(v.length, 2));
            }
        });
        return out;
    }

    // ---- notes (comments): a mark at the corner of the cell, the note while the mouse is on it

    const notebox = el("div", "notebox");
    notebox.hidden = true;
    document.body.appendChild(notebox);
    $("grid").addEventListener("mouseover", (e) => {
        const td = e.target.closest && e.target.closest("td");
        if (!td || !td.note || drag || filling || resizing) {
            if (!(e.target.closest && e.target.closest(".notebox"))) notebox.hidden = true;
            return;
        }
        notebox.textContent = "";
        if (td.note.author) notebox.appendChild(el("b", null, `${td.note.author}:`));
        notebox.appendChild(el("div", null, td.note.text));
        notebox.hidden = false;
        const b = td.getBoundingClientRect();
        const x = Math.min(b.right + 12, window.innerWidth - notebox.offsetWidth - 4);
        const y = Math.max(0, Math.min(b.top - 4, window.innerHeight - notebox.offsetHeight - 4));
        Object.assign(notebox.style, { left: `${Math.max(0, x)}px`, top: `${y}px` });
    });
    $("grid").addEventListener("mouseleave", () => (notebox.hidden = true));
    $("grid").addEventListener("scroll", () => (notebox.hidden = true), { passive: true });

    // ---- hyperlinks: a click follows the link, a click held selects the cell (as in Excel)

    // =HYPERLINK("...", ...) with its address as text.
    function formulaLink(cell) {
        const m = cell && cell.f && /^=\s*HYPERLINK\(\s*"((?:[^"]|"")*)"/i.exec(cell.f);
        if (!m) return null;
        const to = m[1].replace(/""/g, '"');
        return to.startsWith("#") ? { to: to.slice(1) } : /^[a-z][a-z0-9+.-]*:/i.test(to) ? { url: to } : { to };
    }
    function follow(link) {
        if (link.to !== undefined) {
            if (!goTo(link.to.replace(/^#/, ""))) showError(`参照が正しくありません（リンク先: ${link.to}）。`);
            return;
        }
        const url = link.url;
        if (!/^(https?:\/\/|mailto:)/i.test(url)) {
            showError(`リンク先 ${url} は開けません（http・https・mailto のリンクだけを開きます）。`);
            return;
        }
        if (vscode) vscode.postMessage({ type: "link", url });
        else window.open(url, "_blank", "noopener");
    }
    // The link pressed: { td, timer } until the button is let go.
    let linkPress = null;

    // ---- lists (input rules): the button of the active cell, its list, and what may be entered

    function listAt(r, c) {
        if (!view) return null;
        for (const rule of view.sheet.lists || []) {
            if (rule.ranges.some(([r1, c1, r2, c2]) => r >= r1 && r <= r2 && c >= c1 && c <= c2)) return rule;
        }
        return null;
    }
    function placeList() {
        const v = view;
        if (!v) return;
        const rule = sel && !selected && sel.kind === "cells" ? listAt(sel.ar, sel.ac) : null;
        const o = sel ? ownerOf(sel.ar, sel.ac) || [sel.ar, sel.ac, sel.ar, sel.ac] : null;
        const b = rule && o ? boxOf(o[0], o[1], o[2], o[3]) : null;
        const button = !!b && editable && !!rule.values && !entry;
        v.dropbtn.hidden = !button;
        if (button) {
            Object.assign(v.dropbtn.style, { left: `${b.x + b.w + 1}px`, top: `${b.y + Math.max(0, b.h - 17)}px` });
            pinTo(v.dropbtn, pinOf(v, b));
        }
        const message = !!b && !!(rule.title || rule.prompt);
        v.promptbox.hidden = !message;
        if (message) {
            v.promptbox.textContent = "";
            if (rule.title) v.promptbox.appendChild(el("b", null, rule.title));
            if (rule.prompt) v.promptbox.appendChild(el("div", null, rule.prompt));
            Object.assign(v.promptbox.style, { left: `${b.x + Math.min(b.w, 40)}px`, top: `${b.y + b.h + 6}px` });
            pinTo(v.promptbox, pinOf(v, b));
        }
    }
    function openList() {
        if (!editable || busy || !view || !sel) return;
        const rule = listAt(sel.ar, sel.ac);
        if (!rule || !rule.values) return;
        const pop = openPop(view.dropbtn, "listpop");
        const now = ((activeTd && activeTd.cell) || {}).t || "";
        const sheet = view.sheet.name;
        const at = address(sel.ar, sel.ac);
        let first = null;
        for (const value of rule.values) {
            const b = popItem(pop, value, () => postEdit([{ op: "set", sheet, range: at, text: value }]));
            if (value === now) b.classList.add("on");
            first = first || b;
        }
        if (!rule.values.length) pop.appendChild(el("div", "empty", "（リストは空です）"));
        const on = pop.querySelector("button.on") || first;
        if (on) requestAnimationFrame(() => on.focus());
        pop.addEventListener("keydown", (e) => {
            const items = [...pop.querySelectorAll("button")];
            const i = items.indexOf(document.activeElement);
            if (e.key === "ArrowDown" || e.key === "ArrowUp") {
                e.preventDefault();
                e.stopPropagation();
                const n = items[Math.max(0, Math.min(items.length - 1, i + (e.key === "ArrowDown" ? 1 : -1)))];
                if (n) n.focus();
            } else if (e.key === "Escape") {
                e.preventDefault();
                e.stopPropagation();
                closeMenu();
                focusGrid();
            }
        });
    }
    // What was typed into a cell with a list that stops other entries: the message, or null.
    function refused(text) {
        if (!sel || !view || text === "" || /^[=+-]./.test(text)) return null;
        const rule = listAt(sel.ar, sel.ac);
        if (!rule || rule.stop === undefined || !rule.all) return null;
        if (rule.all.some((x) => x.toLowerCase() === text.toLowerCase())) return null;
        return rule.stop || "この値は、このセルに定義されているデータ入力規則の制限を満たしていません。";
    }

    // ---- find and replace (Ctrl+F, Ctrl+H)

    // The state of the box, kept while the page is drawn again (after a replacement).
    const finder = Object.assign({ open: false, replace: false, what: "", with: "", matchCase: false, whole: false, lookIn: "formulas", within: "sheet" }, saved.finder || {});
    let findBox = null;
    // `what` as Excel reads it: * any characters, ? one, ~ the next as it is.
    function finderRegex(f, global) {
        let src = "";
        for (let i = 0; i < f.what.length; i++) {
            const ch = f.what[i];
            if (ch === "~" && i + 1 < f.what.length) src += f.what[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
            else if (ch === "*") src += "[\\s\\S]*";
            else if (ch === "?") src += "[\\s\\S]";
            else src += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        }
        if (f.whole) src = `^(?:${src})$`;
        return new RegExp(src, (f.matchCase ? "" : "i") + (global ? "g" : ""));
    }
    // What a cell is searched by: its formula (or what was typed), its value as shown, its note.
    function searchedText(sheet, c, f) {
        if (f.lookIn === "values") return c.t || "";
        return c.f || c.v || c.t || "";
    }
    // The cells that match, in Excel's order (by rows), from the sheet shown when `within` is the workbook.
    function findAll(f) {
        if (!f.what) return [];
        const re = finderRegex(f, false);
        const sheets = f.within === "book" ? data.sheets.filter((s) => s.visible === "visible") : [view.sheet];
        const out = [];
        for (const sheet of sheets) {
            const items = [];
            if (f.lookIn === "notes") {
                for (const n of sheet.notes || []) if (re.test(n.text)) items.push({ sheet, r: n.r, c: n.c, text: n.text });
            } else {
                for (const c of sheet.cells) {
                    const text = searchedText(sheet, c, f);
                    if (text !== "" && re.test(text)) items.push({ sheet, r: c.r, c: c.c, text, cell: c });
                }
            }
            items.sort((a, b) => a.r - b.r || a.c - b.c);
            out.push(...items);
        }
        return out;
    }
    function findNext(back) {
        const found = findAll(finder);
        if (!found.length) {
            findStatus(`「${finder.what}」が見つかりません。`);
            return null;
        }
        // After (before) the active cell, on to the next sheet and round again.
        const order = data.sheets.indexOf(view.sheet);
        const key = (x) => [data.sheets.indexOf(x.sheet), x.r, x.c];
        const here = [order, sel ? sel.ar : 0, sel ? sel.ac : 0];
        const cmp = (a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
        let hit = back ? [...found].reverse().find((x) => cmp(key(x), here) < 0) : found.find((x) => cmp(key(x), here) > 0);
        hit = hit || (back ? found[found.length - 1] : found[0]);
        showFound(hit);
        findStatus(`${found.length} 件見つかりました。`);
        return hit;
    }
    function showFound(hit) {
        if (hit.sheet !== view.sheet) draw(hit.sheet);
        setSel(cellSel(hit.r, hit.c), true);
    }
    function findStatus(text) {
        if (findBox) findBox.querySelector(".status").textContent = text;
    }
    // The cell's content with what matches replaced: the text to enter, or null (no match, or
    // nothing to replace in a value a formula gives).
    function replaced(cell) {
        const text = cell.f || cell.v || cell.t || "";
        if (!text) return null;
        const f = Object.assign({}, finder, { lookIn: "formulas" });
        if (!finderRegex(f, false).test(text)) return null;
        // (The text replacing: what it says, `$` included.)
        const next = text.replace(finderRegex(f, true), () => finder.with);
        return next === text ? null : next;
    }
    function replaceOne() {
        if (!editable || busy || !view || !sel) return;
        const cell = view.cells.get(sel.ar * KEY + sel.ac);
        const next = cell && finderRegex(Object.assign({}, finder, { lookIn: "formulas" }), false).test(cell.f || cell.v || cell.t || "") ? replaced(cell) : null;
        if (next === null) {
            findNext(false);
            return;
        }
        remember({ finder, findAfter: true });
        postEdit([{ op: "set", sheet: view.sheet.name, range: address(sel.ar, sel.ac), text: next }]);
    }
    function replaceAll() {
        if (!editable || busy || !view) return;
        const f = Object.assign({}, finder, { lookIn: "formulas" });
        const ops = [];
        for (const hit of findAll(f)) {
            const next = hit.cell ? replaced(hit.cell) : null;
            if (next !== null) ops.push({ op: "set", sheet: hit.sheet.name, range: address(hit.r, hit.c), text: next });
        }
        if (!ops.length) {
            findStatus(`「${finder.what}」が見つかりません。`);
            return;
        }
        remember({ finder, findReplaced: ops.length });
        postEdit(ops);
    }
    function closeFind() {
        if (findBox) findBox.remove();
        findBox = null;
        finder.open = false;
        remember({ finder });
        focusGrid();
    }
    function openFind(replace) {
        if (!view) return;
        if (entry) commit(null);
        closeMenu();
        finder.open = true;
        finder.replace = replace && editable;
        remember({ finder });
        if (findBox) findBox.remove();
        const box = (findBox = document.body.appendChild(el("form", "findbox")));
        box.setAttribute("role", "dialog");
        box.dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: false });
        const tabs = box.appendChild(el("div", "findtabs"));
        const tab = (label, on, act) => {
            const b = tabs.appendChild(el("button", on ? "on" : null, label));
            b.type = "button";
            b.addEventListener("click", act);
        };
        tab("検索", !finder.replace, () => openFind(false));
        if (editable) tab("置換", finder.replace, () => openFind(true));
        const close = tabs.appendChild(el("button", "x", "×"));
        close.type = "button";
        close.title = "閉じる（Esc）";
        close.addEventListener("click", closeFind);
        const field = (label, key) => {
            const line = box.appendChild(el("label", "field"));
            line.appendChild(el("span", null, label));
            const input = line.appendChild(el("input"));
            input.type = "text";
            input.spellcheck = false;
            input.value = finder[key];
            input.addEventListener("input", () => {
                finder[key] = input.value;
                remember({ finder });
            });
            return input;
        };
        const what = field("検索する文字列:", "what");
        if (finder.replace) field("置換後の文字列:", "with");
        const opts = box.appendChild(el("div", "opts"));
        const check = (label, key) => {
            const line = opts.appendChild(el("label"));
            const input = line.appendChild(el("input"));
            input.type = "checkbox";
            input.checked = !!finder[key];
            input.addEventListener("change", () => {
                finder[key] = input.checked;
                remember({ finder });
            });
            line.appendChild(document.createTextNode(label));
        };
        const choose = (label, key, options) => {
            const line = opts.appendChild(el("label"));
            line.appendChild(document.createTextNode(label));
            const s = line.appendChild(el("select"));
            for (const [value, text] of options) s.appendChild(el("option", null, text)).value = value;
            s.value = finder[key];
            s.addEventListener("change", () => {
                finder[key] = s.value;
                remember({ finder });
            });
        };
        choose("検索場所:", "within", [["sheet", "シート"], ["book", "ブック"]]);
        if (!finder.replace) choose("検索対象:", "lookIn", [["formulas", "数式"], ["values", "値"], ["notes", "メモ"]]);
        check("大文字と小文字を区別する", "matchCase");
        check("セル内容が完全に同一であるものを検索する", "whole");
        const buttons = box.appendChild(el("div", "buttons"));
        const button = (label, act, cls) => {
            const b = buttons.appendChild(el("button", cls || null, label));
            b.type = "button";
            b.addEventListener("click", () => {
                if (!finder.what) {
                    what.focus();
                    return;
                }
                act();
            });
            return b;
        };
        if (finder.replace) {
            button("すべて置換", replaceAll);
            button("置換", replaceOne);
        }
        button("すべて検索", () => listFound(findAll(finder)));
        button("次を検索", () => findNext(false), "ok");
        box.appendChild(el("div", "status"));
        box.appendChild(el("div", "found"));
        box.addEventListener("submit", (e) => e.preventDefault());
        box.addEventListener("keydown", (e) => {
            if (e.key === "Escape") {
                e.preventDefault();
                e.stopPropagation();
                closeFind();
            } else if (e.key === "Enter" && e.target.tagName === "INPUT") {
                e.preventDefault();
                e.stopPropagation();
                if (finder.what) findNext(e.shiftKey);
            }
        });
        what.focus();
        what.select();
    }
    function listFound(found) {
        if (!findBox) return;
        const list = findBox.querySelector(".found");
        list.textContent = "";
        findStatus(found.length ? `${found.length} 件見つかりました。` : `「${finder.what}」が見つかりません。`);
        for (const hit of found.slice(0, 1000)) {
            const row = list.appendChild(el("div", "hit"));
            row.appendChild(el("span", "where", `${hit.sheet.name}!${address(hit.r, hit.c)}`));
            row.appendChild(el("span", "what", hit.text));
            row.title = hit.text;
            row.addEventListener("click", () => showFound(hit));
        }
    }

    // ---- the panes frozen: the toolbar's menu

    function freeze(at) {
        if (!editable || !view) return;
        const sheet = view.sheet.name;
        postEdit([at ? { op: "freeze", sheet, at } : { op: "freeze", sheet, off: true }]);
    }
    function openFreeze(anchor) {
        const pop = openPop(anchor, "freezes");
        const v = view;
        if (v.fr || v.fc) popItem(pop, "ウィンドウ枠固定の解除", () => freeze(null)).dataset.freeze = "off";
        else if (sel && (sel.ar > 0 || sel.ac > 0)) popItem(pop, "ウィンドウ枠の固定", () => freeze(address(sel.ar, sel.ac))).dataset.freeze = "here";
        popItem(pop, "先頭行の固定", () => freeze("A2")).dataset.freeze = "row";
        popItem(pop, "先頭列の固定", () => freeze("B1")).dataset.freeze = "col";
    }

    // ---- the keys

    // The rows a page key moves by.
    function pageRows() {
        const g = $("grid");
        const room = g.clientHeight / Z() - rowTop(view, view.fr);
        return Math.max(1, Math.floor(room / Math.max(1, view.sheet.defaultHeight)) - 1);
    }

    function entryKey(e) {
        const input = entry.where === "bar" ? bar : ed;
        const ctrl = e.ctrlKey || e.metaKey;
        // (An input method takes its own keys.)
        if (e.isComposing || e.keyCode === 229) return;
        if (assistKey(e)) {
            e.preventDefault();
            return;
        }
        if (e.key === "Enter") {
            e.preventDefault();
            if (e.altKey) {
                input.setRangeText("\n", input.selectionStart, input.selectionEnd, "end");
                input.dispatchEvent(new Event("input"));
            } else if (ctrl) commit({ fill: true });
            else commit({ move: [e.shiftKey ? -1 : 1, 0] });
        } else if (e.key === "Tab") {
            e.preventDefault();
            commit({ move: [0, e.shiftKey ? -1 : 1] });
        } else if (e.key === "Escape") {
            e.preventDefault();
            cancelEntry();
        } else if (e.key === "F2") {
            e.preventDefault();
            entry.mode = entry.mode === "edit" ? "enter" : "edit";
            entry.point = null;
            showMode();
        } else if (e.key === "F4") {
            e.preventDefault();
            toggleAbsolute();
        } else if (ctrl && !e.altKey && ((e.key === ";" && !e.shiftKey) || e.key === ":")) {
            // Today's date, the time, at the caret.
            e.preventDefault();
            input.setRangeText(nowText(e.key === ":"), input.selectionStart, input.selectionEnd, "end");
            input.dispatchEvent(new Event("input"));
        } else if (entry.mode === "enter" && /^Arrow/.test(e.key) && !e.altKey && pointable()) {
            // Pointing with the arrows: from the cell pointed at (the active cell at first).
            e.preventDefault();
            const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
            const p = entry.point;
            const from = p && p.kind === "cells" ? p.cur : [sel.ar, sel.ac];
            const next = (ctrl ? jump : step)(from[0], from[1], d[0], d[1]);
            grow(next[0], next[1]);
            pointTo(e.shiftKey ? (p && p.kind === "cells" ? p.anchor : from) : next, next, "cells");
            reveal(next[0], next[1]);
        } else if (entry.mode === "enter" && entry.where === "cell" && /^Arrow/.test(e.key) && !ctrl && !e.shiftKey && !e.altKey) {
            // An entry that was typed over the cell is entered by the arrows, as in Excel.
            e.preventDefault();
            const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
            commit({ move: d });
        } else if (entry.point && !["Shift", "Control", "Alt", "Meta"].includes(e.key)) {
            // Any other key (the caret moved, a character typed) ends the pointing.
            entry.point = null;
            showMode();
        }
    }

    document.addEventListener("keydown", (e) => {
        if (dialog) {
            if (e.key === "Escape") {
                e.preventDefault();
                closeDialog();
            }
            return;
        }
        if (menu && e.key === "Escape") {
            e.preventDefault();
            closeMenu();
            return;
        }
        if (entry) {
            entryKey(e);
            return;
        }
        const ctrl = e.ctrlKey || e.metaKey;
        const key = e.key;
        // Find, replace, go to: from the grid and from the boxes of the page.
        if (view && ctrl && !e.altKey && !e.shiftKey && (key === "f" || key === "F" || key === "h" || key === "H")) {
            e.preventDefault();
            openFind((key === "h" || key === "H") && editable);
            return;
        }
        // (F5 only where cells are edited: elsewhere it reloads the page, or goes on debugging.)
        if (view && ((key === "F5" && editable && !e.altKey && !ctrl && !e.shiftKey) || (ctrl && !e.altKey && !e.shiftKey && (key === "g" || key === "G")))) {
            e.preventDefault();
            askGoTo();
            return;
        }
        // A shape being drawn: Esc stops.
        if (drawingShape && key === "Escape") {
            e.preventDefault();
            stopDrawing();
            return;
        }
        // The object selected: Delete deletes it, the arrows move it, F2 and Enter edit its text,
        // Esc goes back to the cells.
        if (view && selectedObj && selected && onGrid() && e.altKey && !ctrl && (key === "ArrowLeft" || key === "ArrowRight")) {
            // Alt+← / Alt+→: turned by 15°, as Excel turns a shape.
            e.preventDefault();
            turnObject(selectedObj, key === "ArrowRight" ? 15 : -15);
            return;
        }
        if (view && selectedObj && selected && onGrid() && !e.altKey) {
            const obj = selectedObj;
            const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
            if (key === "Escape") {
                e.preventDefault();
                clearSelection();
                if (sel) setSel(sel, false);
                return;
            }
            if (!ctrl || moves[key]) {
                e.preventDefault();
                if (key === "Delete" || key === "Backspace") deleteObject(obj);
                else if (moves[key]) nudgeObject(obj, selected, moves[key][0], moves[key][1], ctrl);
                else if (key === "F2" || key === "Enter") editObjectText(obj);
                return;
            }
        }
        if (!view || !sel || !onGrid()) return;
        const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
        let done = true;
        if (e.altKey && ctrl && (key === "v" || key === "V")) askPasteSpecial();
        else if (e.altKey && key === "ArrowDown" && listAt(sel.ar, sel.ac)) openList();
        else if (e.altKey) done = false;
        else if (arrows[key]) move(arrows[key][0], arrows[key][1], e.shiftKey, ctrl);
        else if (key === "Tab") move(0, e.shiftKey ? -1 : 1, false);
        else if (key === "Enter" && !ctrl) move(e.shiftKey ? -1 : 1, 0, false);
        else if (key === "Home") {
            const [r, c] = ctrl ? [0, 0] : [sel.kind === "cells" ? sel.fr : sel.ar, 0];
            if (e.shiftKey) setSel({ kind: "cells", nr: sel.nr, nc: sel.nc, fr: r, fc: c }, true);
            else setSel(cellSel(r, c), true);
        } else if (key === "PageDown" || key === "PageUp") {
            const n = pageRows() * (key === "PageDown" ? 1 : -1);
            const from = e.shiftKey && sel.kind === "cells" ? sel.fr : sel.ar;
            const r = Math.max(0, Math.min(view.sheet.maxRows - 1, from + n));
            if (e.shiftKey) setSel({ kind: "cells", nr: sel.nr, nc: sel.nc, fr: r, fc: sel.fc }, true);
            else setSel(cellSel(r, sel.ac), true);
        } else if (ctrl && (key === "a" || key === "A")) setSel({ kind: "all", nr: 0, nc: 0, fr: 0, fc: 0 }, false);
        else if (key === "Escape") setClip(null);
        else if (key === "F2" && e.shiftKey && !ctrl) askNote();
        else if (key === "F2") beginEntry("edit", null, "cell");
        else if (key === "Delete" || key === "Backspace") clearSel();
        else if (ctrl && !e.shiftKey && (key === "z" || key === "Z")) postHistory("undo");
        else if (ctrl && (key === "y" || key === "Y" || (e.shiftKey && (key === "z" || key === "Z")))) postHistory("redo");
        else if (ctrl && !e.shiftKey && (key === "b" || key === "B")) toggleFormat("bold");
        else if (ctrl && !e.shiftKey && (key === "i" || key === "I")) toggleFormat("italic");
        else if (ctrl && !e.shiftKey && (key === "u" || key === "U")) toggleFormat("underline");
        else if (ctrl && !e.shiftKey && (key === "5" || e.code === "Digit5")) toggleFormat("strike");
        else if (ctrl && e.shiftKey && (key === "l" || key === "L")) toggleFilter();
        else if (ctrl && !e.shiftKey && (key === "t" || key === "T" || key === "l" || key === "L")) askTable();
        else if (ctrl && !e.shiftKey && (key === "k" || key === "K")) askLink();
        else if (ctrl && !e.shiftKey && (key === "d" || key === "D")) fill("down");
        else if (ctrl && !e.shiftKey && (key === "r" || key === "R")) fill("right");
        else if (ctrl && !e.shiftKey && key === ";") typeNow(false);
        else if (ctrl && key === ":") typeNow(true);
        else if (ctrl && (key === "+" || e.code === "NumpadAdd" || (e.shiftKey && (key === ";" || e.code === "Semicolon" || e.code === "Equal")))) insertOrDelete(true);
        else if (ctrl && (key === "-" || e.code === "NumpadSubtract" || e.code === "Minus")) insertOrDelete(false);
        else done = false;
        if (done) e.preventDefault();
    });
    // A click on the page outside a control gives the keys back to the grid.
    document.addEventListener("mouseup", (e) => {
        if (!entry && !dialog && !(e.target.closest && e.target.closest("button, input, select, label, textarea, .menu, #side"))) focusGrid();
    });

    // More rows and columns while the sheet is scrolled to its end.
    let pending = false;
    const more = () => {
        if (pending || !view) return;
        pending = true;
        requestAnimationFrame(() => {
            pending = false;
            if (!view) return;
            const before = [view.nrows, view.ncols, view.head];
            view.fill();
            if (view.ncols !== before[1]) refreshOverflow(view, before[1]);
            // Wider row headers (more digits) move the columns: the objects follow.
            if (view.head !== before[2]) placeObjects(view);
            if (view.nrows !== before[0] || view.ncols !== before[1] || view.head !== before[2]) placeFrozenLines(view);
            // Whole rows and columns selected reach as far as the grid is drawn.
            if (view.nrows !== before[0] || view.ncols !== before[1] || view.head !== before[2]) drawBoxes();
            remember({ scroll: { sheet: view.sheet.name, x: $("grid").scrollLeft, y: $("grid").scrollTop } });
        });
    };
    $("grid").addEventListener("scroll", more, { passive: true });
    $("grid").addEventListener("scroll", repin, { passive: true });
    window.addEventListener("resize", more);

    // Columns added after text that runs over them: their gridlines too.
    function refreshOverflow(v, from) {
        for (const key of v.noRight) {
            const c = key % KEY;
            if (c < from) continue;
            const tr = v.trs.get(Math.floor(key / KEY));
            const td = tr && Array.from(tr.children).find((x) => x.c === c);
            if (td) td.classList.add("nogr");
        }
    }

    /**
     * A path of a text warp (flat x,y in pt; k px per pt) measured along its length: at(s) is the
     * point and the unit tangent at the length s (px).
     */
    function track(flat, k) {
        const pts = [];
        for (let i = 0; i + 1 < flat.length; i += 2) pts.push([flat[i] * k, flat[i + 1] * k]);
        const len = [0];
        for (let i = 1; i < pts.length; i++) len.push(len[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
        const total = len[len.length - 1];
        const at = (s) => {
            s = Math.max(0, Math.min(total, s));
            let i = 1;
            while (i < pts.length - 1 && len[i] < s) i++;
            // The tangent of the segment (skipping empty ones).
            let j = i;
            while (j < pts.length - 1 && len[j] - len[j - 1] < 1e-9) j++;
            const [a, b] = [pts[j - 1], pts[j]];
            const d = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
            const seg = len[i] - len[i - 1];
            const f = seg > 1e-9 ? (s - len[i - 1]) / seg : 0;
            const [p, q] = [pts[i - 1], pts[i]];
            return { x: p[0] + (q[0] - p[0]) * f, y: p[1] + (q[1] - p[1]) * f, tx: (b[0] - a[0]) / d, ty: (b[1] - a[1]) / d };
        };
        return { total, at, frac: (u) => at(u * total) };
    }

    /**
     * WordArt's text warps (bodyPr's prstTxWarp), as xljs prints them: the lines of a .warp box
     * are laid out along the warp's paths (data-warp: { w, h, paths } in pt), each character with
     * its own affine transform. Paths in pairs are the top and bottom of a band, and the band's
     * lines are stretched to fill it; a single path (an arch, a circle, the lines of a button)
     * carries its lines standing on it, stretched along it, on the side of the box's center.
     * Lines are shared out among the bands (or single paths) in order.
     */
    function warpText(root) {
        for (const box of root.querySelectorAll(".warp")) {
            let spec;
            try {
                spec = JSON.parse(box.dataset.warp);
            } catch (e) {
                continue;
            }
            const k = box.offsetWidth / spec.w;
            if (!(k > 0) || !spec.paths || !spec.paths.length) continue;
            const paths = spec.paths.map((p) => track(p, k));
            // The lines: paragraphs, split at <br>.
            const lines = [];
            for (const p of box.querySelectorAll("p")) {
                let line = [];
                lines.push(line);
                for (const node of p.childNodes) {
                    if (node.nodeName === "BR") lines.push((line = []));
                    else if (node.nodeType === 1 && node.textContent) line.push(node);
                }
            }
            // Each character as a span of its run, measured on a row as it is set.
            const meas = document.createElement("div");
            meas.style.cssText = "position:absolute;left:0;top:0;visibility:hidden";
            box.append(meas);
            const laid = lines.map((runs) => {
                const row = document.createElement("div");
                row.style.cssText = "position:absolute;left:0;top:0;white-space:pre";
                meas.append(row);
                const chars = [];
                for (const run of runs) {
                    // A run with an inner &nbsp; (empty paragraph) is a blank line.
                    if (run.textContent === " ") continue;
                    for (const ch of Array.from(run.textContent)) {
                        const s = run.cloneNode(false);
                        s.textContent = ch;
                        row.append(s);
                        chars.push(s);
                    }
                }
                const xs = chars.map((s) => s.offsetLeft);
                for (const s of chars) Object.assign(s.style, { position: "absolute", left: "0", top: "0", margin: "0", transformOrigin: "0 0", whiteSpace: "pre" });
                const items = chars.map((s, i) => ({ s, x: xs[i], w: s.offsetWidth, h: s.offsetHeight }));
                const W = Math.max(0, ...items.map((c) => c.x + c.w));
                const H = Math.max(0, ...items.map((c) => c.h));
                return { items, W, H };
            });
            meas.remove();
            if (!laid.some((l) => l.items.length)) continue;
            const pairs = paths.length % 2 === 0;
            const slots = pairs ? paths.length / 2 : paths.length;
            const groups = Array.from({ length: slots }, () => []);
            laid.forEach((l, i) => groups[Math.min(slots - 1, Math.floor((i * slots) / laid.length))].push(l));
            const out = [];
            // Puts the character c (or its part from u0 to u1, clipped to it) with the corners
            // tl, tr (of that part) and bl.
            const place = (c, tl, tr, bl, u0 = 0, u1 = c.w) => {
                if (!(c.w > 0 && c.h > 0 && u1 > u0)) return;
                const a = (tr.x - tl.x) / (u1 - u0),
                    b = (tr.y - tl.y) / (u1 - u0);
                const m = [a, b, (bl.x - tl.x) / c.h, (bl.y - tl.y) / c.h, tl.x - a * u0, tl.y - b * u0];
                if (!m.every(Number.isFinite)) return;
                const s = u0 > 0 || u1 < c.w ? c.s.cloneNode(true) : c.s;
                s.style.transform = `matrix(${m.join(",")})`;
                if (s !== c.s) {
                    // Strips overlap a little so that no seam shows; the outer ones (and all
                    // above and below) are open for outlines, shadows and glows.
                    const l = u0 > 0 ? `${u0 - 0.4}px` : "-100%",
                        r = u1 < c.w ? `${c.w - u1 - 0.4}px` : "-100%";
                    s.style.clipPath = `inset(-100% ${r} -100% ${l})`;
                }
                // A gradient spans the line, not each character.
                if (s.style.backgroundImage) {
                    s.style.backgroundSize = `${c.lineW}px ${c.lineH}px`;
                    s.style.backgroundPosition = `${-c.x}px ${-(c.lineH - c.h)}px`;
                }
                out.push(s);
            };
            const cx = (spec.w * k) / 2,
                cy = (spec.h * k) / 2;
            groups.forEach((group, g) => {
                const m = group.length;
                group.forEach((line, li) => {
                    const { W, H } = line;
                    if (!(W > 0 && H > 0)) return;
                    for (const c of line.items) Object.assign(c, { lineW: W, lineH: H });
                    if (pairs) {
                        // Between the top and bottom paths, this line's share of the band.
                        const [top, bottom] = [paths[2 * g], paths[2 * g + 1]];
                        const pt = (x, y) => {
                            const a = top.frac(x / W),
                                b = bottom.frac(x / W);
                            const t = (li + y / H) / m;
                            return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
                        };
                        for (const c of line.items) {
                            // In vertical strips (about 5pt wide), each through its middle: its
                            // center line and its width at half height.
                            const y0 = H - c.h;
                            const n = Math.max(1, Math.min(16, Math.round(c.w / (5 * k))));
                            for (let i = 0; i < n; i++) {
                                const u0 = (c.w * i) / n,
                                    u1 = (c.w * (i + 1)) / n,
                                    xm = c.x + (u0 + u1) / 2;
                                const tm = pt(xm, y0),
                                    bm = pt(xm, H),
                                    l = pt(c.x + u0, y0 + c.h / 2),
                                    r = pt(c.x + u1, y0 + c.h / 2);
                                const hx = (r.x - l.x) / 2,
                                    hy = (r.y - l.y) / 2;
                                place(c, { x: tm.x - hx, y: tm.y - hy }, { x: tm.x + hx, y: tm.y + hy }, { x: bm.x - hx, y: bm.y - hy }, u0, u1);
                            }
                        }
                        return;
                    }
                    // Standing on the path, stretched along it.
                    const path = paths[g];
                    const kx = path.total / W;
                    const mid = path.at(path.total / 2);
                    const side = mid.ty * (cx - mid.x) - mid.tx * (cy - mid.y);
                    // The distance of the line's top from the path, along "up".
                    const a = side > H / 2 ? (m - li) * H : side < -H / 2 ? -li * H : (m / 2 - li) * H;
                    for (const c of line.items) {
                        // Centered on the path's point under the character's middle, across it
                        // as wide as the chord it takes, upright to it.
                        const p0 = path.at(c.x * kx),
                            p1 = path.at((c.x + c.w) * kx),
                            pm = path.at((c.x + c.w / 2) * kx);
                        const up = { x: pm.ty, y: -pm.tx };
                        const half = Math.hypot(p1.x - p0.x, p1.y - p0.y) / 2;
                        const hx = pm.tx * half,
                            hy = pm.ty * half;
                        const top = a - (H - c.h);
                        const tx = pm.x + up.x * top,
                            ty = pm.y + up.y * top;
                        place(c, { x: tx - hx, y: ty - hy }, { x: tx + hx, y: ty + hy }, { x: tx - hx - up.x * c.h, y: ty - hy - up.y * c.h });
                    }
                });
            });
            box.replaceChildren(...out);
        }
    }

    $("formulas").classList.toggle("on", formulas);
    $("formulas").setAttribute("aria-pressed", String(formulas));
    // Ctrl+` (Ctrl+Shift+@ on a Japanese keyboard), as in Excel.
    document.addEventListener("keydown", (e) => {
        if (e.ctrlKey && !e.altKey && (e.key === "`" || e.code === "Backquote")) {
            e.preventDefault();
            e.stopPropagation();
            showFormulas(!formulas);
        }
    });

    $("file").textContent = data.file;
    const total = Object.values(changes).reduce((n, list) => n + list.length, 0);
    if (data.diff) {
        $("summary").textContent = total || data.diff.structure.length ? `${data.diff.before} からの変更: セル ${total} 個` : `${data.diff.before} からの変更はありません`;
    }
    for (const sheet of data.sheets) {
        const b = el("button", sheet.visible === "visible" ? null : "hidden", sheet.name + (sheet.visible === "visible" ? "" : "（非表示）"));
        b.sheet = sheet;
        // In the editor, the tab's context menu offers the sheet commands (copy, rename, remove).
        b.dataset.vscodeContext = JSON.stringify({ webviewSection: "sheetTab", sheet: sheet.name, preventDefaultContextMenuItems: true });
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
    // The sheet asked for, else the one shown last time, else the first with changes, else the active one.
    const first =
        (data.focus && data.sheets.find((s) => s.name === data.active)) ||
        data.sheets.find((s) => s.name === saved.sheet) ||
        data.sheets.find((s) => (changes[s.name] || []).length) ||
        data.sheets.find((s) => s.name === data.active) ||
        data.sheets.find((s) => s.visible === "visible") ||
        data.sheets[0];
    if (first) {
        const again = saved.sheet === first.name && saved.scroll && saved.scroll.sheet === first.name ? saved.scroll : null;
        draw(first);
        const list = changes[first.name] || [];
        if (list.length && saved.sheet !== first.name) go(first.name, list[0].a);
        // Drawn again (after an edit, at a stop of the debugger): where it was.
        if (again) scrollTo(again.x, again.y);
        if (view) pickObject(view);
        focusGrid();
        // The box of find and replace, open again after a replacement.
        if (finder.open) {
            openFind(finder.replace);
            if (saved.findAfter) findNext(false);
            if (saved.findReplaced) findStatus(`${saved.findReplaced} 件を置換しました。`);
            remember({ findAfter: false, findReplaced: 0 });
        }
        // Fonts that come late change the heights of rows: the objects follow them.
        if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => view && view.sheet === first && placeObjects(view));
    } else {
        $("grid").appendChild(el("div", "empty", "シートがありません"));
    }
})();
