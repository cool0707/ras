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
    const sels = saved.sels || {};
    // What is being typed: { mode: "enter" | "edit", where: "cell" | "bar", original }.
    let entry = null;
    // An edit was posted: nothing more until the page is drawn again, or the edit is refused.
    let busy = false;
    let drag = null;
    // The cells copied or cut: { sheet, kind, r1, c1, r2, c2, cut, text }.
    let clip = saved.clip || null;

    function clearSelection() {
        if (selected) selected.classList.remove("on");
        selected = null;
        selectedObj = null;
        refreshEventButton();
        if (activeTd) activeTd.classList.remove("sel");
        activeTd = null;
        if (view) {
            view.refs.textContent = "";
            view.selbox.hidden = true;
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

    // The name box and the formula bar for the active cell.
    function showCell(sheet, cell, r, c) {
        cell = cell || {};
        view.refs.textContent = "";
        $("name").textContent = address(r, c);
        const fx = $("fx");
        fx.textContent = "";
        if (cell.f) {
            // The references in colors, the same on the grid.
            const refs = references(cell.f);
            const colors = new Map();
            let at = 0;
            for (const ref of refs) {
                const key = `${ref.sheet || ""}!${ref.area.join(",")}`;
                if (!colors.has(key)) colors.set(key, REF_COLORS[colors.size % REF_COLORS.length]);
                const color = colors.get(key);
                fx.appendChild(document.createTextNode(cell.f.slice(at, ref.start)));
                fx.appendChild(el("span", "ref", cell.f.slice(ref.start, ref.end))).style.color = color;
                at = ref.end;
                if (!ref.sheet || ref.sheet === sheet.name) outline(ref.area, color);
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
        refreshEventButton();
        $("name").textContent = obj.name || "";
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
        const shown = typeof cell.x === "number" ? generalFit(cell.x, fits) : null;
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
        $("name").textContent = "";
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
        selbox.hidden = clipbox.hidden = fillbox.hidden = fillh.hidden = true;
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
                    if (text && !/pre-wrap|--shrink|--rot/.test(td.style.cssText) && (!/text-align/.test(td.style.cssText) || (formulas && cell.f)) && !cells.has(r * KEY + next) && !span.has(key)) {
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
                    if (typeof cell.bar.axis === "number") {
                        const axis = td.appendChild(el("div", "cfaxis"));
                        axis.style.left = `calc(2px + (100% - 4px) * ${cell.bar.axis})`;
                    }
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
            }
            if (cell || was || merged) v.byAddress.set(a, td);
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
            tr.appendChild(el("th", null, String(r + 1))).r = r;
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
            v.head = 28 + 8 * Math.max(0, digits - 3);
            corner.style.width = `${v.head}px`;
            table.style.width = `${v.head + v.width}px`;
        }

        // The cells of the sheet, then as many empty ones as the window shows.
        for (let c = 0; c < sheet.cols; c++) addCol(c);
        for (let r = 0; r < sheet.rows; r++) addRow(r);
        sizeHead();
        Object.assign(v, { addRow, addCol, sizeHead });
        v.fill = function fill() {
            let more = false;
            const right = grid.scrollLeft + grid.clientWidth + AHEAD;
            while (v.ncols < sheet.maxCols && v.head + v.width < right) {
                addCol(v.ncols);
                more = true;
            }
            const bottom = grid.scrollTop + grid.clientHeight + AHEAD;
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
        overflow(v);
        placeObjects(v);
        draw.find = (a) => v.byAddress.get(a);
        // The selection the sheet had (A1 at first).
        const was = sels[sheet.name];
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
            const box = layer.appendChild(el("div", `obj ${obj.kind}`));
            Object.assign(box.style, { left: `${x}px`, top: `${y}px`, width: `${x2 - x}px`, height: `${y2 - y}px` });
            box.innerHTML = obj.html;
            if (obj.chart && window.rasChart) {
                try {
                    window.rasChart.draw(box.querySelector(".chartbox") || box, obj.chart);
                } catch (e) {
                    // A chart that cannot be drawn leaves its box empty.
                }
            }
            if (obj.name) box.title = obj.macro ? `${obj.name}（マクロ: ${obj.macro}）` : obj.name;
            box.addEventListener("click", (e) => {
                e.stopPropagation();
                selectObject(box, obj);
            });
            // As the form designer: a double click opens the default event, the menu the others.
            box.addEventListener("dblclick", (e) => {
                e.stopPropagation();
                if (!hasEvents(obj)) return;
                e.preventDefault();
                selectObject(box, obj);
                openEvent(obj, false);
            });
            box.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                selectObject(box, obj);
                const items = eventItems(obj);
                if (items.length) showMenu(e.clientX, e.clientY, items);
            });
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
        sels[sheet.name] = { kind, nr, nc, fr, fc };
        remember({ sels });
        paint(scroll);
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
    }

    // The box of the range, its headers, and the dashed box of the cells copied.
    function drawBoxes() {
        const v = view;
        if (!v) return;
        const put = (div, b) => {
            div.hidden = !b;
            if (b) Object.assign(div.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w + 1}px`, height: `${b.h + 1}px` });
        };
        for (const th of v.hl) th.classList.remove("hs");
        v.hl = [];
        if (sel && !selected) {
            for (const [c, th] of v.ths) if (c >= sel.c1 && c <= sel.c2) v.hl.push(th);
            for (const [r, tr] of v.rows) if (r >= sel.r1 && r <= sel.r2) v.hl.push(tr.firstChild);
            for (const th of v.hl) th.classList.add("hs");
            put(v.selbox, single(sel) ? null : boxOf(sel.r1, sel.c1, sel.r2, sel.c2));
        }
        put(v.clipbox, clip && clip.sheet === v.sheet.name ? boxOf(clip.r1, clip.c1, clip.r2, clip.c2) : null);
        // The fill handle: at the bottom right corner of the cells selected.
        const b = editable && sel && !selected && !entry && sel.kind === "cells" ? boxOf(sel.r1, sel.c1, sel.r2, sel.c2) : null;
        v.fillh.hidden = !b || sel.r2 >= v.nrows || sel.c2 >= v.ncols;
        if (b) Object.assign(v.fillh.style, { left: `${b.x + b.w - 3}px`, top: `${b.y + b.h - 3}px` });
    }

    // Scrolls a cell out from under the headers and into the window (null: that way stays).
    function reveal(r, c) {
        const v = view;
        const g = $("grid");
        const o = (r !== null && c !== null && ownerOf(r, c)) || [r || 0, c || 0, r || 0, c || 0];
        const b = boxOf(o[0], o[1], o[2], o[3]);
        if (!b) return;
        const top = v.table.tHead.offsetHeight;
        if (c !== null) {
            if (b.x - v.head < g.scrollLeft) g.scrollLeft = b.x - v.head;
            else if (b.x + b.w + 1 > g.scrollLeft + g.clientWidth) g.scrollLeft = Math.min(b.x - v.head, b.x + b.w + 1 - g.clientWidth);
        }
        if (r !== null) {
            if (b.y - top < g.scrollTop) g.scrollTop = b.y - top;
            else if (b.y + b.h + 1 > g.scrollTop + g.clientHeight) g.scrollTop = Math.min(b.y - top, b.y + b.h + 1 - g.clientHeight);
        }
    }

    // Scrolls to where the grid was, drawing it that far.
    function scrollTo(x, y) {
        const v = view;
        const g = $("grid");
        const before = [v.nrows, v.ncols, v.head];
        while (v.ncols < v.sheet.maxCols && v.head + v.width < x + g.clientWidth) v.addCol(v.ncols);
        let tall = v.table.offsetHeight;
        while (v.nrows < v.sheet.maxRows && tall < y + g.clientHeight) {
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
            setSel({ kind: "cells", nr, nc, fr: r, fc: c }, true);
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
        if (entry) commit(null);
        focusGrid();
        if (edge) {
            beginResize(edge, e);
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
        if (e.shiftKey && sel && sel.kind === hit.kind && hit.kind !== "all") setSel({ kind: hit.kind, nr: sel.nr, nc: sel.nc, fr: hit.r, fc: hit.c }, false);
        else setSel({ kind: hit.kind, nr: hit.r, nc: hit.c, fr: hit.r, fc: hit.c }, false);
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
        if (!drag) hoverEdge(e);
        if (!drag || !sel) return;
        if (!(e.buttons & 1)) {
            drag = null;
            return;
        }
        const hit = hitOf(document.elementFromPoint(e.clientX, e.clientY));
        if (!hit || hit.kind === "all") return;
        // Along the headers, the cells under the mouse count by their row or column.
        if (drag !== "cells" && hit.kind !== drag && hit.kind !== "cells") return;
        if (drag === "cells" && hit.kind !== "cells") return;
        if (hit.r === sel.fr && hit.c === sel.fc) return;
        setSel({ kind: drag, nr: sel.nr, nc: sel.nc, fr: hit.r, fc: hit.c }, false);
    });
    document.addEventListener("mouseup", (e) => {
        drag = null;
        if (resizing) endResize(e);
        if (filling) endFill();
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
        entry = { mode, where, original };
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
        return true;
    }

    function endEntry() {
        entry = null;
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
        const fill = !!(how && how.fill) && !single(sel);
        const range = fill ? rangeText(sel) : address(sel.ar, sel.ac);
        const sheet = view.sheet.name;
        endEntry();
        if (how && how.move) move(how.move[0], how.move[1], false);
        if (text === e.original && !fill) return;
        postEdit([text === "" ? { op: "clear", sheet, range } : { op: "set", sheet, range, text }]);
    }

    ed.addEventListener("input", () => {
        if (!entry) {
            // The first character typed starts an entry that replaces the cell's content.
            if (!beginEntry("enter", ed.value, "cell")) ed.value = "";
            return;
        }
        $("fx").textContent = ed.value;
        placeEditor();
    });
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
        if (!editable || busy || !sel || !view) return;
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

    document.addEventListener("copy", (e) => {
        if (entry || !onGrid() || !sel || selected) return;
        e.clipboardData.setData("text/plain", copySel(false));
        e.preventDefault();
    });
    document.addEventListener("cut", (e) => {
        if (entry || !onGrid() || !sel || selected) return;
        e.clipboardData.setData("text/plain", copySel(true));
        e.preventDefault();
    });
    document.addEventListener("paste", (e) => {
        if (entry || !onGrid()) return;
        e.preventDefault();
        pasteText(e.clipboardData.getData("text/plain"));
    });
    // From the menu: through the clipboard where the page may use it.
    function menuCopy(cut) {
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
        if (sel && view) postEdit([{ op: "clear", sheet: view.sheet.name, range: rangeText(sel) }]);
    }

    // `how`: "shift" (the cells move right or down / left or up), "rows", "cols".
    function structure(insert, how, across) {
        if (!sel || !view || sel.kind === "all") return;
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
        if (!editable || busy || !sel || sel.kind === "all") return;
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
        if (!sel || !view || sel.kind === "all") return;
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
            if (kind !== "all") {
                items.push(null);
                items.push([header ? "挿入" : "挿入…", () => insertOrDelete(true), "Ctrl++"]);
                items.push([header ? "削除" : "削除…", () => insertOrDelete(false), "Ctrl+-"]);
            }
            items.push(["数式と値のクリア", clearSel, "Delete"]);
            items.push(["書式のクリア", clearFormats, ""]);
            if (kind === "cells") {
                items.push(null);
                if (mergedIn(sel)) items.push(["結合の解除", () => merge("unmerge"), ""]);
                if (!single(sel)) items.push(["セルの結合", () => merge("merge"), ""]);
                items.push(["昇順で並べ替え", () => sortBy(false), ""]);
                items.push(["降順で並べ替え", () => sortBy(true), ""]);
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
    // The button of the bar: the events of the object selected (a menu).
    function refreshEventButton() {
        const b = $("event");
        if (!b) return;
        const on = hasEvents(selectedObj);
        b.disabled = !on;
        b.textContent = on && selectedObj.activex === undefined ? "マクロ" : "イベント";
    }
    function enableEvents() {
        events = true;
        if (!$("event")) {
            const b = el("button", null, "イベント");
            b.id = "event";
            b.type = "button";
            b.title = "選択した図形・コントロールのイベント（マクロ）を開きます。シートのファイルにハンドラーを追加します（図形をダブルクリックすると既定のイベント）";
            b.addEventListener("mousedown", (e) => e.preventDefault());
            b.addEventListener("click", () => {
                const items = eventItems(selectedObj);
                if (!items.length) return;
                const r = b.getBoundingClientRect();
                showMenu(r.left, r.bottom + 1, items);
            });
            $("formulas").before(b);
        }
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
        if (sel && view) postEdit([{ op: "format", sheet: view.sheet.name, range: areaText(sel), set }]);
    }
    function toggleFormat(name) {
        if (editable) format({ [name]: !STATE[name](activeCss()) });
    }
    function clearFormats() {
        if (sel && view) postEdit([{ op: "clearFormats", sheet: view.sheet.name, range: areaText(sel) }]);
    }
    // A merged cell among the cells selected.
    function mergedIn(s) {
        return !!s && view.sheet.merges.some(([a, b, c, d]) => a <= s.r2 && c >= s.r1 && b <= s.c2 && d >= s.c1);
    }
    // `how`: "center" (merged and centered), "across" (each row on its own), "merge", "unmerge".
    function merge(how) {
        if (!sel || !view || sel.kind === "all") return;
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
        if (sel && view) postEdit([{ op: "sort", sheet: view.sheet.name, range: rangeOrCell(), by: colName(sel.ac), desc }]);
    }
    function toggleFilter() {
        if (!editable || !sel || !view) return;
        const sheet = view.sheet.name;
        postEdit([view.sheet.filter ? { op: "filter", sheet, off: true } : { op: "filter", sheet, range: rangeOrCell() }]);
    }
    function hide(hidden) {
        if (!sel || !view) return;
        const sheet = view.sheet.name;
        if (sel.kind === "all") postEdit([{ op: "hide", sheet, range: `1:${MAXR}`, hidden }, { op: "hide", sheet, range: `A:${colName(MAXC - 1)}`, hidden }]);
        else if (sel.kind !== "cells") postEdit([{ op: "hide", sheet, range: rangeText(sel), hidden }]);
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
    // The colors to choose from; `name` is the format ("color", "fill"), `none` what no color is called.
    function openColors(anchor, name, none) {
        const pop = openPop(anchor, "colors");
        popItem(pop, none, () => format({ [name]: null }));
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
                format({ [name]: color.toUpperCase() });
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
            format({ [name]: `#${m[1].toUpperCase()}` });
        });
    }
    const BORDERS = [["all", "格子"], ["outline", "外枠"], ["thick", "太い外枠"], ["top", "上"], ["bottom", "下"], ["left", "左"], ["right", "右"], ["none", "枠なし"]];
    const NUMBER_FORMATS = [["General", "標準"], ["0.00", "数値"], ["#,##0", "桁区切り"], ["¥#,##0", "通貨"], ["0%", "パーセント"], ["yyyy/m/d", "日付"], ["h:mm", "時刻"], ["@", "文字列"]];
    const FONT_SIZES = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

    // The toolbar (where cells are edited): its buttons, by what they do.
    const tools = {};
    function makeTools() {
        const box = el("div");
        box.id = "tools";
        $("bar").after(box);
        // The grid keeps the keys: a click on a button does not take them.
        box.addEventListener("mousedown", (e) => {
            if (!e.target.closest("select")) e.preventDefault();
        });
        const button = (key, label, title, act, cls) => {
            const b = box.appendChild(el("button", cls || null, label));
            b.type = "button";
            b.title = title;
            b.dataset.tool = key;
            b.addEventListener("click", () => {
                if (busy || !sel) return;
                if (entry) commit(null);
                act(b);
                focusGrid();
            });
            tools[key] = b;
            return b;
        };
        const select = (key, title, options) => {
            const s = box.appendChild(el("select"));
            s.title = title;
            s.dataset.tool = key;
            for (const [value, label] of options) s.appendChild(el("option", null, label)).value = value;
            tools[key] = s;
            return s;
        };
        const gap = () => box.appendChild(el("span", "gap"));
        button("bold", "B", "太字（Ctrl+B）", () => toggleFormat("bold"), "b");
        button("italic", "I", "斜体（Ctrl+I）", () => toggleFormat("italic"), "i");
        button("underline", "U", "下線（Ctrl+U）", () => toggleFormat("underline"), "u");
        button("strike", "S", "取り消し線（Ctrl+5）", () => toggleFormat("strike"), "s");
        const size = select("fontSize", "フォント サイズ", FONT_SIZES.map((n) => [String(n), String(n)]));
        size.addEventListener("change", () => {
            format({ fontSize: +size.value });
            focusGrid();
        });
        button("color", "A", "フォントの色", (b) => openColors(b, "color", "自動"), "fc");
        button("fill", "◧", "塗りつぶしの色", (b) => openColors(b, "fill", "塗りつぶしなし"));
        button("border", "田", "罫線", (b) => {
            const pop = openPop(b, "borders");
            for (const [value, label] of BORDERS) popItem(pop, label, () => format({ border: value })).dataset.border = value;
        });
        gap();
        for (const [value, label, title] of [["left", "⇤", "左揃え"], ["center", "↔", "中央揃え"], ["right", "⇥", "右揃え"]]) button(`align-${value}`, label, title, () => format({ align: value }));
        for (const [value, label, title] of [["top", "⤒", "上揃え"], ["middle", "↕", "上下中央揃え"], ["bottom", "⤓", "下揃え"]]) button(`valign-${value}`, label, title, () => format({ valign: value }));
        button("wrap", "↵", "折り返して全体を表示する", () => toggleFormat("wrap"));
        button("merge", "結合", "セルを結合して中央揃え（結合したセルでは結合の解除）", () => merge(sel && single(sel) && mergedIn(sel) ? "unmerge" : "center"));
        button("mergeMore", "▾", "セルの結合", (b) => {
            const pop = openPop(b, "merges");
            for (const [how, label] of [["center", "セルを結合して中央揃え"], ["across", "横方向に結合"], ["merge", "セルの結合"], ["unmerge", "結合の解除"]]) popItem(pop, label, () => merge(how)).dataset.merge = how;
        }, "more");
        gap();
        const number = select("numberFormat", "表示形式", [["", "表示形式"]].concat(NUMBER_FORMATS));
        number.addEventListener("change", () => {
            if (number.value) format({ numberFormat: number.value });
            focusGrid();
        });
        gap();
        button("sortAsc", "A→Z", "昇順で並べ替え", () => sortBy(false));
        button("sortDesc", "Z→A", "降順で並べ替え", () => sortBy(true));
        button("filter", "フィルター", "フィルター（Ctrl+Shift+L）", toggleFilter);
        gap();
        button("clearFormats", "書式のクリア", "書式のクリア", clearFormats);
    }
    // The buttons show the format of the active cell.
    function showTools() {
        if (!tools.bold || !view) return;
        const css = activeCss();
        const cell = (activeTd && activeTd.cell) || {};
        for (const name of Object.keys(STATE)) tools[name].classList.toggle("on", STATE[name](css));
        const align = (/text-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["left", "center", "right"]) tools[`align-${a}`].classList.toggle("on", align === a);
        const valign = (/vertical-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["top", "middle"]) tools[`valign-${a}`].classList.toggle("on", valign === a);
        const size = (/font-size:([\d.]+)pt/.exec(css) || [])[1] || "11";
        if (![...tools.fontSize.options].some((o) => o.value === size)) tools.fontSize.appendChild(el("option", null, size)).value = size;
        tools.fontSize.value = size;
        const code = cell.n || "General";
        tools.numberFormat.value = NUMBER_FORMATS.some(([c]) => c === code) ? code : "";
        tools.merge.classList.toggle("on", !!sel && single(sel) && mergedIn(sel));
        tools.filter.classList.toggle("on", !!view.sheet.filter);
    }
    if (editable) makeTools();

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
        const px = Math.round((r.now - r.start) / (r.cols && formulas ? 2 : 1));
        const sheet = view.sheet.name;
        const range = sized(r);
        postEdit([px <= 0 ? { op: "hide", sheet, range, hidden: true } : { op: "size", sheet, range, px }]);
    }
    // As wide as the widest text of the column's cells, as tall as the tallest content of the row's.
    function autofit(edge) {
        const v = view;
        const sheet = v.sheet;
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
        postEdit([{ op: "size", sheet: sheet.name, range: sized(edge), px: Math.round(px) }]);
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
        const range = rangeText(sel);
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

    // ---- the keys

    // The rows a page key moves by.
    function pageRows() {
        const g = $("grid");
        const room = g.clientHeight - view.table.tHead.offsetHeight;
        return Math.max(1, Math.floor(room / Math.max(1, view.sheet.defaultHeight)) - 1);
    }

    function entryKey(e) {
        const input = entry.where === "bar" ? bar : ed;
        const ctrl = e.ctrlKey || e.metaKey;
        // (An input method takes its own keys.)
        if (e.isComposing || e.keyCode === 229) return;
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
        } else if (entry.mode === "enter" && entry.where === "cell" && /^Arrow/.test(e.key) && !ctrl && !e.shiftKey && !e.altKey) {
            // An entry that was typed over the cell is entered by the arrows, as in Excel.
            e.preventDefault();
            const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
            commit({ move: d });
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
        if (!view || !sel || !onGrid()) return;
        const ctrl = e.ctrlKey || e.metaKey;
        const key = e.key;
        const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
        let done = true;
        if (e.altKey) done = false;
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
        else if (key === "F2") beginEntry("edit", null, "cell");
        else if (key === "Delete" || key === "Backspace") clearSel();
        else if (ctrl && !e.shiftKey && (key === "z" || key === "Z")) postHistory("undo");
        else if (ctrl && (key === "y" || key === "Y" || (e.shiftKey && (key === "z" || key === "Z")))) postHistory("redo");
        else if (ctrl && !e.shiftKey && (key === "b" || key === "B")) toggleFormat("bold");
        else if (ctrl && !e.shiftKey && (key === "i" || key === "I")) toggleFormat("italic");
        else if (ctrl && !e.shiftKey && (key === "u" || key === "U")) toggleFormat("underline");
        else if (ctrl && !e.shiftKey && (key === "5" || e.code === "Digit5")) toggleFormat("strike");
        else if (ctrl && e.shiftKey && (key === "l" || key === "L")) toggleFilter();
        else if (ctrl && !e.shiftKey && (key === "d" || key === "D")) fill("down");
        else if (ctrl && !e.shiftKey && (key === "r" || key === "R")) fill("right");
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
            // Whole rows and columns selected reach as far as the grid is drawn.
            if (view.nrows !== before[0] || view.ncols !== before[1] || view.head !== before[2]) drawBoxes();
            remember({ scroll: { sheet: view.sheet.name, x: $("grid").scrollLeft, y: $("grid").scrollTop } });
        });
    };
    $("grid").addEventListener("scroll", more, { passive: true });
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

    $("formulas").addEventListener("click", () => showFormulas(!formulas));
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
        focusGrid();
        // Fonts that come late change the heights of rows: the objects follow them.
        if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => view && view.sheet === first && placeObjects(view));
    } else {
        $("grid").appendChild(el("div", "empty", "シートがありません"));
    }
})();
