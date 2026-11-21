// The grid of `ras book view` (and of the editor's sheet panel): draws the sheets of DATA.
// DATA = { file, sheets: [sheet_view...], images: [data URL], diff: { before, structure: [text], changes: { sheet: [{ a, before, after }] } } | null, active: name | null, focus?: true }
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
    let selected = null;
    // The sheet as drawn: its table, the rows and columns so far, where they are.
    let view = null;

    function clearSelection() {
        if (selected) selected.classList.remove(selected.classList.contains("obj") ? "on" : "sel");
        selected = null;
        if (view) view.refs.textContent = "";
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

    function select(sheet, td, r, c) {
        clearSelection();
        selected = td;
        td.classList.add("sel");
        const cell = td.cell || {};
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
            const at = selected && selected.r !== undefined ? [selected.r, selected.c] : null;
            draw(current);
            grid.scrollLeft = x;
            grid.scrollTop = y;
            view.fill();
            // The cell selected before, again.
            const tr = at && view.trs.get(at[0]);
            const td = tr && Array.from(tr.children).find((t) => t.c === at[1]);
            if (td) select(current, td, at[0], at[1]);
        }
    }

    function selectObject(box, obj) {
        clearSelection();
        selected = box;
        box.classList.add("on");
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
        const cells = new Map();
        for (const c of sheet.cells) cells.set(c.r * KEY + c.c, c);
        // Cells covered by a merged cell are not drawn; its first cell spans them.
        const span = new Map();
        const covered = new Set();
        for (const [r1, c1, r2, c2] of sheet.merges) {
            span.set(r1 * KEY + c1, [r2, c2]);
            for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.add(r * KEY + c);
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
        head.appendChild(el("th"));
        const body = table.appendChild(el("tbody"));
        wrap.appendChild(table);
        const layer = wrap.appendChild(el("div", "objs"));
        const refs = wrap.appendChild(el("div", "refs"));
        grid.appendChild(wrap);

        const v = (view = {
            sheet,
            table,
            body,
            layer,
            cols: [], // the visible columns drawn, in order
            rows: [], // the visible rows drawn: [row, tr]
            trs: new Map(), // row -> tr
            ths: new Map(), // column -> its header
            ncols: 0, // columns drawn (visible or not)
            nrows: 0,
            width: 0, // of the cells (without the row headers)
            refs, // the boxes of the references of the selected formula
            head: 40, // the width of the row headers
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
                const shown = formulas ? cell.f || cell.v || cell.t : cell.t;
                if (formulas && cell.f) {
                    td.style.textAlign = "left";
                    td.classList.add("fml");
                }
                if (shown !== undefined) {
                    // Text runs over the empty cells to its right, as in Excel.
                    const next = v.cols[i + 1] !== undefined ? v.cols[i + 1] : c + 1;
                    const text = cell.k === "string" || (formulas && cell.f);
                    if (text && !/pre-wrap/.test(td.style.cssText) && (!/text-align/.test(td.style.cssText) || (formulas && cell.f)) && !cells.has(r * KEY + next) && !span.has(key)) {
                        td.classList.add("ov");
                        td.appendChild(el("div", "ovt")).appendChild(el("span", null, shown));
                    } else {
                        td.textContent = shown;
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
            if (filter && filter[0] && filter[1] && r === filter[0][0] && c >= filter[0][1] && c <= filter[1][1]) td.classList.add("filter");
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
            tr.appendChild(el("th", null, String(r + 1)));
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
            v.ths.set(c, head.appendChild(el("th", null, colName(c))));
            const i = v.cols.length;
            v.cols.push(c);
            v.width += w;
            for (const [r, tr] of v.rows) makeCell(tr, r, c, i);
        }

        // The row headers as wide as the largest row number needs.
        function sizeHead() {
            const digits = String(v.nrows).length;
            v.head = Math.max(40, digits * 8 + 12);
            corner.style.width = `${v.head}px`;
            table.style.width = `${v.head + v.width}px`;
        }

        // The cells of the sheet, then as many empty ones as the window shows.
        for (let c = 0; c < sheet.cols; c++) addCol(c);
        for (let r = 0; r < sheet.rows; r++) addRow(r);
        sizeHead();
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
            select(sheet, td, at[0], at[1]);
        }
    }

    $("grid").addEventListener("click", (e) => {
        const td = e.target.closest && e.target.closest("td");
        if (td && view && td.r !== undefined) select(view.sheet, td, td.r, td.c);
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
        draw(first);
        const list = changes[first.name] || [];
        if (list.length && saved.sheet !== first.name) go(first.name, list[0].a);
        // Fonts that come late change the heights of rows: the objects follow them.
        if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => view && view.sheet === first && placeObjects(view));
    } else {
        $("grid").appendChild(el("div", "empty", "シートがありません"));
    }
})();
