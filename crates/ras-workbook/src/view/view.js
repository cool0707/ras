// The grid of `ras book view` (and of the editor's sheet panel): draws the sheets of DATA.
// (A sheet: `filter` is its filter's range, `filterColumns` { offset: { values, blanks } | { other } } the columns filtered.)
// DATA = { editable?: true, host?: "parent", functions?, iconSets?: { set: [SVG] }, file, sheets: [sheet_view...], images: [data URL], diff: { before, structure: [text], changes: { sheet: [{ a, before, after }] } } | null, active: name | null, focus?: true }
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
    // The editor is VS Code's (its webview), or, with `host: "parent"`, the page that holds this
    // one in a frame (the playground): what is posted goes to it as `{ rasView: message }`, and
    // what is kept is kept in the session's storage.
    const parentHost = () => {
        const key = "ras-view-state";
        return {
            postMessage: (m) => window.parent.postMessage({ rasView: m }, "*"),
            getState: () => {
                try {
                    return JSON.parse(sessionStorage.getItem(key) || "null");
                } catch {
                    return null;
                }
            },
            setState: (state) => {
                try {
                    sessionStorage.setItem(key, JSON.stringify(state));
                } catch {
                    // Then it is not kept.
                }
            },
        };
    };
    const vscode =
        typeof acquireVsCodeApi === "function" ? acquireVsCodeApi() : data.host === "parent" && window.parent !== window ? parentHost() : null;
    const saved = (vscode && vscode.getState()) || {};
    const remember = (patch) => {
        Object.assign(saved, patch);
        if (vscode) vscode.setState(saved);
    };

    let current = null;
    // The object (a shape, a control) selected, and the cell of the active cell.
    let selected = null;
    let selectedObj = null;
    // The other objects selected with it (Ctrl or Shift and a click), as Excel selects several.
    let others = [];
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
    // An object moved or sized and sent, until the sheet is drawn again: it back as it was if the
    // edit is not made.
    let undrag = null;
    // The cells copied or cut: { sheet, kind, r1, c1, r2, c2, cut, text }.
    let clip = saved.clip || null;

    function clearSelection() {
        for (const box of [selected, ...others.map((o) => o.box)]) unmark(box);
        selected = null;
        selectedObj = null;
        others = [];
        refreshEventButton();
        refreshShapeTools();
        if (activeTd) activeTd.classList.remove("sel");
        activeTd = null;
        if (view) {
            view.refs.textContent = "";
            view.selbox.hidden = view.selshade.hidden = true;
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

    // `text` into `box`, its references (colorRefs) in their colors (spans of the class `cls`).
    function refSpans(box, text, refs, cls) {
        box.textContent = "";
        let at = 0;
        for (const ref of refs) {
            box.appendChild(document.createTextNode(text.slice(at, ref.start)));
            box.appendChild(el("span", cls, text.slice(ref.start, ref.end))).style.color = ref.color;
            at = ref.end;
        }
        box.appendChild(document.createTextNode(text.slice(at)));
    }

    // The name box and the formula bar for the active cell.
    function showCell(sheet, cell, r, c) {
        cell = cell || {};
        view.refs.textContent = "";
        showName(sheet, r, c);
        const fx = $("fx");
        fx.textContent = "";
        const spill = !cell.f && spillAt(sheet, r, c);
        const anchor = spill && view && view.cells.get(spill[0] * KEY + spill[1]);
        if (cell.f) {
            // The references in colors, the same on the grid.
            refSpans(fx, cell.f, colorRefs(sheet, cell.f, references(cell.f)), "ref");
        } else if (anchor && anchor.f) {
            // A cell a formula spilled into: the formula, grayed, as Excel shows it.
            fx.appendChild(el("span", "spilled", anchor.f));
        } else {
            fx.textContent = cell.v || cell.t || "";
        }
        const was = (changes[sheet.name] || []).find((x) => x.a === address(r, c));
        if (was) fx.appendChild(el("span", "was", `変更前: ${was.before}`));
    }

    // The range a dynamic-array formula spilled into that holds the cell: [r1, c1, r2, c2].
    function spillAt(sheet, r, c) {
        return (sheet.spills || []).find(([r1, c1, r2, c2]) => r >= r1 && r <= r2 && c >= c1 && c <= c2) || null;
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

    // `add` (Ctrl or Shift held): added to the objects selected, or taken out of them if it is
    // one (as Excel selects several shapes and controls; only those that can be moved).
    function selectObject(box, obj, add) {
        if (add && selected && selectedObj && movable(selectedObj) && movable(obj)) {
            if (entry) cancelEntry();
            if (obj === selectedObj || others.includes(obj)) {
                unmark(box);
                if (obj === selectedObj) {
                    const next = others.pop();
                    [selected, selectedObj] = next ? [next.box, next] : [null, null];
                } else others = others.filter((o) => o !== obj);
                if (!selectedObj) {
                    clearSelection();
                    if (sel) setSel(sel, false);
                    return;
                }
            } else {
                others.push(selectedObj);
                [selected, selectedObj] = [box, obj];
                mark(box, obj);
            }
            refreshEventButton();
            refreshShapeTools();
            showObjectName();
            return;
        }
        if (entry) cancelEntry();
        clearSelection();
        selected = box;
        selectedObj = obj;
        mark(box, obj);
        refreshEventButton();
        refreshShapeTools();
        showObjectName();
    }
    function mark(box, obj) {
        box.classList.add("on");
        addHandles(box, obj);
    }
    function unmark(box) {
        if (!box) return;
        box.classList.remove("on");
        box.classList.remove("framed");
        for (const f of box.querySelectorAll(".objframe")) f.remove();
    }
    function showObjectName() {
        const obj = selectedObj;
        $("name").value = obj.name || "";
        $("fx").textContent = obj.macro ? `${eventName(obj)}イベント: ${obj.macro}` : "";
    }
    // The objects selected (the one selected last first); `isSelected`: one of them.
    const selectedObjects = () => (selected && selectedObj ? [selectedObj, ...others] : []);
    const isSelected = (obj) => !!selected && (obj === selectedObj || others.includes(obj));

    // The width of a column and the height of a row (px), drawn or not.
    function width(sheet, c) {
        const k = formulas ? 2 : 1;
        if (c < sheet.cols) return sheet.colWidths[c] * k;
        const w = sheet.moreWidths && sheet.moreWidths[c];
        return (w !== undefined ? w : sheet.defaultWidth) * k;
    }
    function height(sheet, r) {
        if (r < sheet.rows) {
            const h = sheet.rowHeights[r];
            return h === null ? fitHeight(sheet, r) : h;
        }
        const h = sheet.moreHeights && sheet.moreHeights[r];
        return h !== undefined ? h : sheet.rowStyles && sheet.rowStyles[r] !== undefined ? fitHeight(sheet, r) : sheet.defaultHeight;
    }

    // ---- the fonts of the cells, and the heights of the rows Excel fits to them

    // The CSS font-family of an Excel font name (as the workbook writes it: Japanese names also
    // by their English names).
    const FONT_ALIASES = {
        "游ゴシック": "Yu Gothic",
        "游ゴシック Light": "Yu Gothic Light",
        "游ゴシック Medium": "Yu Gothic Medium",
        "游明朝": "Yu Mincho",
        "ＭＳ Ｐゴシック": "MS PGothic",
        "ＭＳ ゴシック": "MS Gothic",
        "ＭＳ Ｐ明朝": "MS PMincho",
        "ＭＳ 明朝": "MS Mincho",
        "メイリオ": "Meiryo",
        "BIZ UDゴシック": "BIZ UDGothic",
        "BIZ UDPゴシック": "BIZ UDPGothic",
        "BIZ UD明朝 Medium": "BIZ UDMincho Medium",
        "BIZ UDP明朝 Medium": "BIZ UDPMincho Medium",
    };
    function fontFamily(name) {
        const quote = (s) => `'${s.replace(/['\\;"]/g, "")}'`;
        let out = quote(name);
        if (FONT_ALIASES[name]) out += `,${quote(FONT_ALIASES[name])}`;
        return out + (/明朝|Mincho/.test(name) ? ",'Yu Mincho',serif" : ",'Yu Gothic',Meiryo,sans-serif");
    }
    // The workbook's font: that of the cells without one of their own.
    const bookFont = (sheet) => sheet.font || { name: "游ゴシック", size: 11, css: fontFamily("游ゴシック") };
    // The font of the cells drawn with the CSS `css`: { family, pt, name }.
    function fontOfCss(sheet, css) {
        const f = bookFont(sheet);
        const family = (/font-family:([^;]+)/.exec(css || "") || [])[1] || f.css;
        const pt = +(/font-size:([\d.]+)pt/.exec(css || "") || [])[1] || f.size;
        const name = (/font-family:'([^']*)'/.exec(css || "") || [])[1] || f.name;
        return { family, pt, name };
    }
    // The ascent and descent of a font `px` high (as the browser rounds them).
    const fontMetrics = new Map();
    function metricsOf(family, px) {
        const k = `${px}|${family}`;
        let m = fontMetrics.get(k);
        if (!m) {
            measurer.font = `${px}px ${family}`;
            const t = measurer.measureText("x");
            m = t.fontBoundingBoxAscent !== undefined ? [t.fontBoundingBoxAscent, t.fontBoundingBoxDescent] : [Math.round(px * 0.9), Math.round(px * 0.25)];
            fontMetrics.set(k, m);
        }
        return m;
    }
    // The height of a line of the cells drawn with `css` (px): the rows Excel fits are no taller,
    // so the text does not make the row grow.
    function lineHeight(sheet, css) {
        const { family, pt } = fontOfCss(sheet, css);
        const [a, d] = metricsOf(family, (pt * 96) / 72);
        return a + d;
    }
    // The height Excel fits a row to for a cell with `css` (px): the font's ascent and descent at
    // its size in whole pixels, and 2 px (Calibri 11: 20, Arial 11: 19, Arial 10: 17). The
    // workbook's font: the sheet's default height.
    function fitOfCss(sheet, css) {
        const f = bookFont(sheet);
        const { family, pt } = fontOfCss(sheet, css);
        if (family === f.css && pt === f.size) return sheet.defaultHeight;
        const [a, d] = metricsOf(family, Math.round((pt * 96) / 72));
        return Math.min(545, a + d + 2);
    }
    // The height Excel fits a row to for a cell with `css` whose text is in `runs` (its
    // characters' fonts, the tallest counts).
    function fitOfRuns(sheet, css, runs) {
        let px = fitOfCss(sheet, css);
        for (const r of runs || []) if (r.f && (r.f.name || r.f.size) && !r.f.va) px = Math.max(px, fitOfCss(sheet, withRunFont(css, r.f)));
        return px;
    }
    // The CSS of a cell with the font name and size of a run.
    function withRunFont(css, f) {
        let out = css || "";
        const put = (prop, value) => (out = new RegExp(`${prop}:[^;]*;?`).test(out) ? out.replace(new RegExp(`${prop}:[^;]*;?`), `${prop}:${value};`) : `${out}${prop}:${value};`);
        if (f.name) put("font-family", fontFamily(f.name));
        if (f.size) put("font-size", `${f.size}pt`);
        return out;
    }
    // ---- text whose characters have formats of their own (rich text): its runs, [{ t, f }],
    // `f` the font of the run ({ name, size, b, i, u, s, color, va }; none: the cell's)

    // The CSS of a run (`deco`: the cell's text-decoration, for the runs of its font: a span
    // cannot take off what its cell draws).
    function runCss(f, deco) {
        if (!f) return deco ? `text-decoration:${deco};` : "";
        let css = "";
        if (f.name) css += `font-family:${fontFamily(f.name)};`;
        if (f.va) css += `vertical-align:${f.va === "subscript" ? "sub" : "super"};font-size:${f.size ? `${(f.size * 0.75).toFixed(2)}pt` : "75%"};`;
        else if (f.size) css += `font-size:${f.size}pt;`;
        css += `font-weight:${f.b ? "bold" : "normal"};font-style:${f.i ? "italic" : "normal"};`;
        const lines = [f.u ? "underline" : "", f.s ? "line-through" : ""].filter(Boolean).join(" ");
        css += `text-decoration:${lines || "none"};`;
        if (f.u && /double/i.test(f.u)) css += "text-decoration-style:double;";
        if (f.color) css += `color:${f.color};`;
        return css;
    }
    // The text of a cell into `node`: its runs as spans (`td`: the cell, whose underline and
    // strikethrough then go to the runs of its font), else as it is.
    function putText(node, text, cell, td) {
        const runs = cell && cell.rt && !formulas && !cell.hv && text === cell.t ? cell.rt : null;
        if (!runs) {
            node.textContent = text;
            return node;
        }
        const deco = td ? td.style.textDecoration || td.style.textDecorationLine : "";
        if (td && deco) td.style.textDecoration = "none";
        node.textContent = "";
        for (const r of runs) node.appendChild(el("span", null, r.t)).style.cssText = runCss(r.f, deco && deco !== "none" ? deco : "");
        return node;
    }
    // A row not given a height: as tall as the tallest font of its cells (merged cells left out,
    // as Excel), or the default height without cells.
    function fitHeight(sheet, r) {
        if (!sheet.fits) {
            const fits = (sheet.fits = new Map());
            const merged = new Set(sheet.merges.map(([r1, c1]) => r1 * KEY + c1));
            const take = (row, css, runs) => fits.set(row, Math.max(fits.get(row) || 0, fitOfRuns(sheet, css, runs)));
            for (const cell of sheet.cells) if (!merged.has(cell.r * KEY + cell.c)) take(cell.r, cell.s ? sheet.styles[cell.s] : "", cell.rt);
            for (const [row, s] of Object.entries(sheet.rowStyles || {})) take(+row, sheet.styles[s]);
        }
        return sheet.fits.get(r) || sheet.defaultHeight;
    }
    // The rows a change of the font of `areas` fits again: [[row, px]] of those it makes taller
    // or shorter. `font(css)` is the CSS of a cell after the change (`null`: as it is).
    function refit(sheet, areaList, font) {
        const custom = new Set(sheet.customRows || []);
        const cells = new Map();
        for (const cell of sheet.cells) {
            if (!cells.has(cell.r)) cells.set(cell.r, []);
            cells.get(cell.r).push(cell);
        }
        const merged = new Set(sheet.merges.map(([r1, c1]) => r1 * KEY + c1));
        const last = Math.max(sheet.rows, view ? view.nrows : 0) - 1;
        const rows = new Set();
        for (const a of areaList) for (let r = a.r1; r <= Math.min(a.r2, last); r++) rows.add(r);
        const out = [];
        for (const r of rows) {
            const now = height(sheet, r);
            if (custom.has(r) || !(now > 0) || (r >= sheet.rows && sheet.moreHeights && sheet.moreHeights[r] !== undefined)) continue;
            const inside = (c) => areaList.some((a) => r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2);
            const rowCss = sheet.rowStyles && sheet.rowStyles[r] !== undefined ? sheet.styles[sheet.rowStyles[r]] : "";
            let px = 0;
            const had = new Set();
            for (const cell of cells.get(r) || []) {
                had.add(cell.c);
                if (merged.has(r * KEY + cell.c)) continue;
                const css = cell.s ? sheet.styles[cell.s] : "";
                px = Math.max(px, fitOfRuns(sheet, inside(cell.c) ? font(css) : css, cell.rt));
            }
            // The empty cells changed: formatted now, they count.
            const empty = areaList.some((a) => r >= a.r1 && r <= a.r2 && a.c2 - a.c1 + 1 > [...had].filter((c) => c >= a.c1 && c <= a.c2).length);
            if (empty) px = Math.max(px, fitOfCss(sheet, font(rowCss)));
            else if (rowCss) px = Math.max(px, fitOfCss(sheet, rowCss));
            if (!px) px = sheet.defaultHeight;
            if (px !== now) out.push([r, px]);
        }
        return out;
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
        red.classList.remove("on");
        red.textContent = "";
        ed.value = "";
        bar.remove();
        if (sheet !== current) others = [];
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
        // The workbook's font, for the cells without one of their own.
        const lines = new Map();
        const lineOf = (css) => {
            if (!lines.has(css)) lines.set(css, lineHeight(sheet, css));
            return lines.get(css);
        };
        table.style.font = `${bookFont(sheet).size}pt ${bookFont(sheet).css}`;
        table.style.lineHeight = `${lineOf("")}px`;
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
        // (The shade of the range: all of it but the active cell, which stays as it is, as Excel's.)
        const selshade = marks.appendChild(el("div", "selshade"));
        // (The border of the range a formula spilled into, while one of its cells is active.)
        const spillbox = marks.appendChild(el("div", "spillbox"));
        const selbox = marks.appendChild(el("div", "selbox"));
        const clipbox = marks.appendChild(el("div", "clipbox"));
        const fillbox = marks.appendChild(el("div", "fillbox"));
        const fillh = marks.appendChild(el("div", "fillhandle"));
        // The button of a list (an input rule) and its message, at the active cell.
        const dropbtn = marks.appendChild(el("div", "dropbtn"));
        dropbtn.appendChild(el("span", "chev"));
        const promptbox = marks.appendChild(el("div", "promptbox"));
        // The lines under the rows and right of the columns frozen.
        const frzh = marks.appendChild(el("div", "frzline h"));
        const frzv = marks.appendChild(el("div", "frzline v"));
        selshade.hidden = spillbox.hidden = selbox.hidden = clipbox.hidden = fillbox.hidden = fillh.hidden = dropbtn.hidden = promptbox.hidden = true;
        frzh.hidden = !fr;
        frzv.hidden = !fc;
        // (Under the marks: the border of the selection stays on the cell while what is typed runs over.)
        wrap.insertBefore(ed, refs);
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
            selshade,
            spillbox,
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
            pats: new Set(), // the cells of pattern fills (td.pat)
        });
        // (Again when the cells move: a row or a column of another size, shown or hidden.)
        if (typeof ResizeObserver === "function") new ResizeObserver(() => alignPatterns(v)).observe(table);

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
                        td.appendChild(el("div", "ovt")).appendChild(putText(el("span"), shown, cell, td));
                    } else if (/pre-wrap/.test(td.style.cssText)) {
                        // Wrapped text is cut at the row's height (a table row would grow to it),
                        // from its first line, as Excel draws it.
                        let h = 0;
                        for (let x = r; x <= (span.has(key) ? span.get(key)[0] : r); x++) h += height(sheet, x);
                        const box = td.appendChild(putText(el("div", "wr"), shown, cell, td));
                        box.style.maxHeight = `${Math.max(0, h - 1)}px`;
                    } else {
                        putText(td, shown, cell, td);
                        const rot = td.style.getPropertyValue("--rot");
                        if (rot) {
                            td.textContent = "";
                            td.appendChild(putText(el("span", "rot"), shown, cell, td)).style.transform = `rotate(${-Number(rot)}deg)`;
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
                        if (cell.bar.ac) axis.style.borderLeftColor = cell.bar.ac;
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
            // Empty cells have their row's or column's format; the lines of text are no taller
            // than the font (not as tall as the browser's "normal") nor than the row.
            const css = cell ? (cell.s && sheet.styles[cell.s]) || "" : cssOf(sheet, r, c, null);
            if (!cell && css) td.style.cssText = css;
            // (Shaded by "ロックされたセルを表示" unless unlocked.)
            if (/--unlocked:1/.test(css)) td.classList.add("unlk");
            // A pattern fill, drawn over the gridlines above and left of it (see view.css), its
            // tiles from the table's corner.
            if (/--fill:/.test(css)) {
                td.classList.add("pat");
                v.pats.add(td);
                alignPatterns(v);
            }
            const merged = span.get(key);
            let line = lineOf(css);
            if (!(merged && merged[0] > r)) line = Math.max(0, Math.min(line, height(sheet, r) - 1));
            if (line !== lineOf("")) td.style.lineHeight = `${line}px`;
            if (v.noRight.has(key)) td.classList.add("nogr");
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

    // The tiles of the pattern fills, from one point (the table's corner) for all the cells, so
    // that a pattern runs on from a cell to the next as Excel draws it: where each cell's box of
    // the pattern (1px above and left of the cell) is, the other way. (Once the cells are laid.)
    function alignPatterns(v) {
        if (v.aligning) return;
        v.aligning = true;
        requestAnimationFrame(() => {
            v.aligning = false;
            const at = [];
            for (const td of v.pats) {
                if (!td.isConnected) v.pats.delete(td);
                else at.push([td, td.offsetLeft + td.clientLeft - 1, td.offsetTop + td.clientTop - 1]);
            }
            for (const [td, x, y] of at) {
                td.style.setProperty("--px", `${-x}px`);
                td.style.setProperty("--py", `${-y}px`);
            }
        });
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
        // (Pattern fills keep their pixels: their tiles are drawn 1/z as large.)
        if (z === 1) v.wrap.style.removeProperty("--pz");
        else v.wrap.style.setProperty("--pz", String(1 / z));
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
    //
    // (Where they are drawn, to the fraction of a pixel: the collapsed borders of the headers put
    // the cells half a pixel off, which offsetLeft and offsetTop round away, and the boxes drawn
    // over the cells, a pixel off theirs, would not stand on the gridlines.)
    function exact(v, node) {
        const w = v.wrap.getBoundingClientRect();
        const r = node.getBoundingClientRect();
        // (The grid's zoom: as getBoundingClientRect has it, whichever way the browser zooms.)
        const k = v.wrap.offsetWidth ? w.width / v.wrap.offsetWidth : 1;
        return { x: (r.left - w.left) / k, y: (r.top - w.top) / k, w: r.width / k, h: r.height / k };
    }
    function columnLeft(v, c) {
        const th = v.ths.get(c);
        if (th) return exact(v, th).x;
        // Hidden or not drawn: after the column drawn before it.
        let k = c - 1;
        while (k >= 0 && !v.ths.has(k)) k--;
        const before = k >= 0 && exact(v, v.ths.get(k));
        let x = before ? before.x + before.w : v.head;
        for (let i = k + 1; i < c; i++) x += width(v.sheet, i);
        return x;
    }
    function rowTop(v, r) {
        const tr = v.trs.get(r);
        if (tr) return exact(v, tr).y;
        // Hidden or not drawn: after the row drawn before it.
        let k = r - 1;
        while (k >= 0 && !v.trs.has(k)) k--;
        const before = k >= 0 && exact(v, v.trs.get(k));
        let y = before ? before.y + before.h : v.table.tHead.offsetHeight;
        for (let i = k + 1; i < r; i++) y += height(v.sheet, i);
        return y;
    }

    function placeObjects(v) {
        undrag = null;
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
            if (obj.name) box.title = obj.macro ? `${obj.name}（${eventName(obj)}イベント: ${obj.macro}）` : obj.name;
            obj.box = box;
            if (movable(obj)) box.classList.add("movable");
            box.addEventListener("click", (e) => {
                e.stopPropagation();
                if (runMode) return;
                // (Ctrl or Shift: the mouse going down added it or took it out. One of several
                // selected stays so, as it was dragged with them.)
                if (e.ctrlKey || e.metaKey || e.shiftKey || (isSelected(obj) && others.length)) return;
                selectObject(box, obj);
            });
            // Dragged: moved (in the panel that edits the workbook).
            box.addEventListener("mousedown", (e) => {
                // Run mode: pressed, as a button is (its events when the mouse goes up on it).
                if (runMode) {
                    if (e.button === 0) {
                        e.preventDefault();
                        runPress = obj;
                        box.classList.add("pressed");
                    }
                    return;
                }
                if (e.button !== 0 || drawingShape) return;
                // Ctrl or Shift: added to the objects selected (or taken out of them).
                if ((e.ctrlKey || e.metaKey || e.shiftKey) && movable(obj) && selected && movable(selectedObj)) {
                    e.preventDefault();
                    selectObject(box, obj, true);
                    focusGrid();
                    return;
                }
                // (What cannot be moved is selected all the same: Delete and the arrows are not
                // for the object selected before it.)
                if (!movable(obj)) {
                    if (selectedObj !== obj) selectObject(box, obj);
                    return;
                }
                e.preventDefault();
                // (One of several selected: moved with the others.)
                if (!isSelected(obj)) selectObject(box, obj);
                focusGrid();
                beginObjectDrag(e, box, obj, "move");
            });
            // As the form designer: a double click opens the default event, the menu the others.
            // Without events, a shape's text is edited.
            box.addEventListener("mouseup", (e) => {
                box.classList.remove("pressed");
                if (!runMode || e.button !== 0 || runPress !== obj) return;
                runPress = null;
                runClick(obj, e);
            });
            box.addEventListener("mouseleave", () => box.classList.remove("pressed"));
            box.addEventListener("dblclick", (e) => {
                e.stopPropagation();
                if (runMode) {
                    e.preventDefault();
                    return;
                }
                if (!hasEvents(obj)) {
                    if (movable(obj) && obj.text !== undefined) {
                        e.preventDefault();
                        selectObject(box, obj);
                        editObjectText(obj, true);
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
                if (runMode) return;
                if (!isSelected(obj)) selectObject(box, obj);
                const items = objectItems(obj);
                if (items.length) showMenu(e.clientX, e.clientY, items);
            });
            // Drawn again (the rows grew, the fonts came): still selected.
            if (selectedObj === obj) selected = box;
            if (selectedObj === obj || others.includes(obj)) mark(box, obj);
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
        if (runMode) runSelectSoon();
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

    // The border of the cells selected (one cell too), 2px as Excel's: 1px over the gridline
    // around them, 1px outside it. (The gridlines are the right and bottom pixels of the cells:
    // those left and above them are the cells' before.) Around the cell while it is typed into
    // too: what is typed past it is shown over the cells next to it, as Excel shows it.
    function borderOf(box) {
        return { x: box.x - 2, y: box.y - 2, w: box.w + 2, h: box.h + 2 };
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
            const one = single(sel) && !areas.length;
            const box = boxOf(sel.r1, sel.c1, sel.r2, sel.c2);
            const range = one ? null : box;
            put(v.selbox, box && !areas.length ? borderOf(box) : range);
            if (!entry) v.selbox.classList.remove("overx", "overy");
            v.selbox.classList.toggle("multi", areas.length > 0);
            put(v.selshade, range);
            const o = ownerOf(sel.ar, sel.ac) || [sel.ar, sel.ac, sel.ar, sel.ac];
            const a = range && boxOf(o[0], o[1], o[2], o[3]);
            if (range) {
                const [W, H] = [range.w + 1, range.h + 1];
                v.selshade.style.clipPath = a
                    ? `polygon(evenodd, 0 0, ${W}px 0, ${W}px ${H}px, 0 ${H}px, 0 0, ${a.x - range.x}px ${a.y - range.y}px, ${a.x - range.x + a.w}px ${a.y - range.y}px, ${a.x - range.x + a.w}px ${a.y - range.y + a.h}px, ${a.x - range.x}px ${a.y - range.y + a.h}px, ${a.x - range.x}px ${a.y - range.y})`
                    : "";
            }
        } else {
            put(v.selshade, null);
            for (const a of areas) {
                const box = v.selbox.parentNode.insertBefore(el("div", "selbox multi more"), v.selbox);
                put(box, boxOf(a.r1, a.c1, a.r2, a.c2));
            }
        }
        // The border of the range the active cell's formula spilled into (or spilled it).
        const sp = sel && !selected && !entry && spillAt(v.sheet, sel.ar, sel.ac);
        const sb = sp && boxOf(sp[0], sp[1], Math.min(sp[2], v.nrows - 1), Math.min(sp[3], v.ncols - 1));
        put(v.spillbox, sb && { x: sb.x - 1, y: sb.y - 1, w: sb.w, h: sb.h });
        // (Where the border of the cells is drawn when they are selected, as Excel: over it.)
        const cb = clip && clip.sheet === v.sheet.name ? boxOf(clip.r1, clip.c1, clip.r2, clip.c2) : null;
        put(v.clipbox, cb && borderOf(cb));
        // The fill handle: at the bottom right corner of the cells selected.
        const b = editable && sel && !selected && !entry && sel.kind === "cells" && !areas.length ? boxOf(sel.r1, sel.c1, sel.r2, sel.c2) : null;
        v.fillh.hidden = !b || sel.r2 >= v.nrows || sel.c2 >= v.ncols;
        // (Centred on the corner of the border.)
        if (b) Object.assign(v.fillh.style, { left: `${b.x + b.w - 3}px`, top: `${b.y + b.h - 3}px` });
        pinTo(v.fillh, pinOf(v, b));
        placeList();
        placeInvalid();
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
        if (e.target === ed || red.contains(e.target)) return;
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
        if (entry) {
            commit(null);
            // (What was typed is asked about: the cell stays as it is.)
            if (entry) return;
        }
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
        if (hit && hit.kind === "cells" && !entry) {
            // Run mode: `BeforeDoubleClick` first; the cell is typed into unless a handler cancels.
            if (runMode && view && sel) {
                postRun({ op: "doubleClick", sheet: view.sheet.name, range: address(sel.ar, sel.ac) }, (m) => {
                    if (!m.cancel && !entry) beginEntry("edit", null, "cell");
                });
                return;
            }
            beginEntry("edit", null, "cell");
        }
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
    // Text whose characters are formatted (some of them red, bigger, struck through, as Excel
    // formats the characters of a cell typed into): typed into a box that shows them, `red`, in
    // place of `ed`. What it holds is `entry.runs` ([{ t, f }]); the box is drawn from them, and
    // what is typed changes them (the browser's own editing is not let to change the box).
    const red = el("div", "editor rich");
    red.contentEditable = "true";
    red.spellcheck = false;
    red.setAttribute("aria-label", "セル");
    red.setAttribute("role", "textbox");
    const runsText = (runs) => runs.map((r) => r.t).join("");
    const sameFont = (a, b) => JSON.stringify(a || null) === JSON.stringify(b || null);
    // (Runs of one font are one; runs without text go.)
    function tidyRuns(runs) {
        const out = [];
        for (const r of runs) {
            if (!r.t) continue;
            const last = out[out.length - 1];
            if (last && sameFont(last.f, r.f)) last.t += r.t;
            else out.push({ t: r.t, f: r.f ? { ...r.f } : null });
        }
        return out;
    }
    // The font of the character at `pos` (past the end: the last one's).
    function fontOfChar(runs, pos) {
        let n = 0;
        for (const r of runs) {
            if (pos < n + r.t.length) return r.f;
            n += r.t.length;
        }
        return runs.length ? runs[runs.length - 1].f : null;
    }
    // The font what is typed at `pos` gets, as in Excel: that of the character before it (the
    // first's at the start).
    const fontAt = (runs, pos) => fontOfChar(runs, Math.max(0, pos - 1));
    // `text` (of the font `f`) in place of the characters start..end.
    function spliceRuns(runs, start, end, text, f) {
        const out = [];
        let n = 0;
        let placed = false;
        const put = () => {
            if (!placed && text) out.push({ t: text, f });
            placed = true;
        };
        for (const r of runs) {
            const a = n;
            const b = (n += r.t.length);
            if (b <= start && !(b === start && a === start)) {
                out.push(r);
                continue;
            }
            if (a >= end) {
                put();
                out.push(r);
                continue;
            }
            if (a < start) out.push({ t: r.t.slice(0, start - a), f: r.f });
            put();
            if (b > end) out.push({ t: r.t.slice(end - a), f: r.f });
        }
        put();
        return tidyRuns(out);
    }
    // The characters start..end changed by `change(font)` (the font of the cell for those of it).
    function restyleRuns(runs, start, end, change, base) {
        const out = [];
        let n = 0;
        for (const r of runs) {
            const a = n;
            const b = (n += r.t.length);
            const from = Math.max(a, start) - a;
            const to = Math.min(b, end) - a;
            if (to <= from) {
                out.push(r);
                continue;
            }
            if (from > 0) out.push({ t: r.t.slice(0, from), f: r.f });
            out.push({ t: r.t.slice(from, to), f: change({ ...(r.f || base) }) });
            if (to < r.t.length) out.push({ t: r.t.slice(to), f: r.f });
        }
        return tidyRuns(out);
    }
    // The runs kept as the text changes from `was` to `now` (in a box that does not show them:
    // the formula bar): what is typed gets the font of the character before it.
    function followRuns(runs, was, now) {
        if (was === now) return runs;
        let a = 0;
        while (a < was.length && a < now.length && was[a] === now[a]) a++;
        let z = 0;
        while (z < was.length - a && z < now.length - a && was[was.length - 1 - z] === now[now.length - 1 - z]) z++;
        return spliceRuns(runs, a, was.length - z, now.slice(a, now.length - z), fontAt(runs, a));
    }
    // The font of the active cell, as a run's (what its characters have unless given another).
    function cellFont() {
        const fm = formatAt(sel.ar, sel.ac);
        const book = bookFont(view.sheet);
        const f = { name: fm.font || book.name, size: fm.size || book.size };
        if (fm.b) f.b = true;
        if (fm.i) f.i = true;
        if (fm.u && fm.u !== "none") f.u = fm.u === true ? "single" : fm.u;
        if (fm.st) f.s = true;
        if (fm.color) f.color = fm.color;
        if (fm.va) f.va = fm.va;
        return f;
    }
    // Whether a format given now is the characters' (selected, or to be typed): while the cell
    // itself is typed into, not a formula (whose characters Excel does not format).
    const charTarget = () => !!entry && entry.where === "cell" && !!view && !!sel && !/^=/.test(entryInput().value);

    // ---- the box: where its caret is, as characters of the text
    function richOffset(node, offset) {
        const r = document.createRange();
        r.selectNodeContents(red);
        try {
            r.setEnd(node, offset);
        } catch (x) {
            return 0;
        }
        return r.toString().length;
    }
    function richSel() {
        const s = document.getSelection();
        if (s && s.rangeCount && red.contains(s.anchorNode) && red.contains(s.focusNode)) {
            const a = richOffset(s.anchorNode, s.anchorOffset);
            const b = richOffset(s.focusNode, s.focusOffset);
            return [Math.min(a, b), Math.max(a, b)];
        }
        return entry && entry.sel ? entry.sel.slice() : [0, 0];
    }
    function richPlace(a, b) {
        const len = runsText(entry.runs).length;
        a = Math.max(0, Math.min(a, len));
        b = Math.max(a, Math.min(b === undefined ? a : b, len));
        entry.sel = [a, b];
        const find = (pos) => {
            const walk = document.createTreeWalker(red, NodeFilter.SHOW_TEXT);
            let n = 0;
            let last = null;
            for (let t = walk.nextNode(); t; t = walk.nextNode()) {
                if (pos <= n + t.data.length) return [t, pos - n];
                n += t.data.length;
                last = t;
            }
            return last ? [last, last.data.length] : [red, 0];
        };
        const s = document.getSelection();
        if (!s) return;
        const r = document.createRange();
        r.setStart(...find(a));
        r.setEnd(...find(b));
        s.removeAllRanges();
        s.addRange(r);
    }
    // The box drawn from the runs, its caret at a..b.
    function richDraw(a, b) {
        red.textContent = "";
        const deco = entry.deco || "";
        for (const r of entry.runs) {
            const span = red.appendChild(el("span", null, r.t));
            span.style.cssText = runCss(r.f, deco);
        }
        // (A line break at the end shows a line after it only with something there.)
        if (/\n$/.test(runsText(entry.runs))) red.appendChild(document.createElement("br"));
        if (document.activeElement === red) richPlace(a, b);
        else entry.sel = [a, b === undefined ? a : b];
    }
    // What was typed (in place of start..end): the box drawn again, the rest as after typing.
    function richReplace(start, end, text) {
        const runs = entry.runs;
        text = text.replace(/\r\n?/g, "\n");
        const p = entry.pending;
        const f = p && p.at === start && start === end ? p.f : end > start ? fontOfChar(runs, start) : fontAt(runs, start);
        entry.undo.push({ runs, sel: richSel() });
        entry.redo = [];
        entry.runs = spliceRuns(runs, start, end, text, f);
        entry.pending = null;
        richDraw(start + text.length);
        richTyped();
    }
    // After the text changed: as `ed`'s input.
    function richTyped() {
        const text = runsText(entry.runs);
        // A formula: its characters are not formatted; typed on in the plain box.
        if (/^=/.test(text)) {
            toPlain();
            return;
        }
        entry.point = null;
        $("fx").textContent = text;
        placeEditor();
        entryRefs();
        showMode();
        updateAssist();
    }
    // The box `entryInput` gives while the characters are formatted (as a textarea's).
    const richBox = {
        get value() {
            return entry && entry.runs ? runsText(entry.runs) : "";
        },
        get selectionStart() {
            return richSel()[0];
        },
        get selectionEnd() {
            return richSel()[1];
        },
        setSelectionRange: (a, b) => richPlace(a, b),
        setRangeText: (text, a, b) => richReplace(a, b, text),
        focus: () => red.focus({ preventScroll: true }),
        select: () => richPlace(0, runsText(entry.runs).length),
        dispatchEvent: () => true,
        getBoundingClientRect: () => red.getBoundingClientRect(),
    };
    // The cell typed into in the box of formats (from the plain box, its text and its caret).
    function toRich() {
        const [a, b] = [ed.selectionStart, ed.selectionEnd];
        if (!entry.runs) entry.runs = tidyRuns([{ t: ed.value, f: null }]);
        entry.rich = true;
        entry.undo = entry.undo || [];
        entry.redo = entry.redo || [];
        ed.value = "";
        ed.classList.remove("on");
        placeEditor();
        red.focus({ preventScroll: true });
        richDraw(a, b);
        richPlace(a, b);
    }
    // Back to the plain box (a formula typed), the runs kept for the text that stays.
    function toPlain() {
        const [a, b] = richSel();
        const text = runsText(entry.runs);
        entry.rich = false;
        entry.lastText = text;
        red.classList.remove("on");
        red.textContent = "";
        ed.value = text;
        placeEditor();
        ed.focus({ preventScroll: true });
        ed.setSelectionRange(a, b);
        $("fx").textContent = text;
        entryRefs();
        showMode();
        updateAssist();
    }
    // Ctrl+Z, Ctrl+Y in the box.
    function richHistory(back) {
        const from = back ? entry.undo : entry.redo;
        const to = back ? entry.redo : entry.undo;
        const step = from.pop();
        if (!step) return;
        to.push({ runs: entry.runs, sel: richSel() });
        entry.runs = step.runs;
        richDraw(step.sel[0], step.sel[1]);
        richTyped();
    }
    red.addEventListener("beforeinput", (e) => {
        if (!entry || !entry.rich) {
            e.preventDefault();
            return;
        }
        // (An input method composes in the box; what it gives is taken when it ends.)
        if (e.inputType === "insertCompositionText" || e.isComposing) return;
        e.preventDefault();
        const t = e.inputType;
        if (t === "historyUndo" || t === "historyRedo") {
            richHistory(t === "historyUndo");
            return;
        }
        let [a, b] = richSel();
        const ranges = e.getTargetRanges ? e.getTargetRanges() : [];
        if (ranges.length && red.contains(ranges[0].startContainer)) {
            const x = richOffset(ranges[0].startContainer, ranges[0].startOffset);
            const y = richOffset(ranges[0].endContainer, ranges[0].endOffset);
            [a, b] = [Math.min(x, y), Math.max(x, y)];
        }
        const text = runsText(entry.runs);
        const paste = () => (e.dataTransfer && e.dataTransfer.getData("text/plain")) || "";
        if (t === "insertText" || t === "insertReplacementText") richReplace(a, b, e.data !== null && e.data !== undefined ? e.data : paste());
        else if (t === "insertLineBreak" || t === "insertParagraph") richReplace(a, b, "\n");
        else if (t === "insertFromPaste" || t === "insertFromDrop" || t === "insertFromYank") richReplace(a, b, paste());
        else if (t.startsWith("delete")) {
            if (a === b) {
                // (Without the ranges: a character, a code point whole.)
                if (/Backward/.test(t) && a > 0) a -= a > 1 && /[\uDC00-\uDFFF]/.test(text[a - 1]) ? 2 : 1;
                else if (/Forward/.test(t) && b < text.length) b += /[\uD800-\uDBFF]/.test(text[b]) ? 2 : 1;
            }
            if (a !== b) richReplace(a, b, "");
        }
    });
    let richComposing = null;
    red.addEventListener("compositionstart", () => {
        if (!entry || !entry.rich) return;
        const [a, b] = richSel();
        const p = entry.pending;
        richComposing = { a, b, text: runsText(entry.runs), runs: entry.runs, f: p && p.at === a && a === b ? p.f : b > a ? fontOfChar(entry.runs, a) : fontAt(entry.runs, a) };
    });
    red.addEventListener("input", () => {
        if (!entry || !entry.rich || !richComposing) return;
        $("fx").textContent = red.textContent;
        placeEditor();
    });
    red.addEventListener("compositionend", () => {
        const c = richComposing;
        richComposing = null;
        if (!entry || !entry.rich || !c) return;
        // What was composed: the box's text between what was before and after the characters
        // it replaced.
        const now = richOffsetText();
        const tail = c.text.length - c.b;
        const ok = now.length >= c.a + tail && now.slice(0, c.a) === c.text.slice(0, c.a) && now.slice(now.length - tail) === c.text.slice(c.b);
        entry.undo.push({ runs: c.runs, sel: [c.a, c.b] });
        entry.redo = [];
        if (ok) {
            const typed = now.slice(c.a, now.length - tail);
            entry.runs = spliceRuns(c.runs, c.a, c.b, typed, c.f);
            entry.pending = null;
            richDraw(c.a + typed.length);
        } else {
            entry.runs = followRuns(c.runs, c.text, now);
            richDraw(now.length);
        }
        richTyped();
    });
    // The box's text as it is now (an input method has typed into it).
    function richOffsetText() {
        const r = document.createRange();
        r.selectNodeContents(red);
        return r.toString();
    }
    // Cut: what is selected goes to the clipboard as text, and from the box.
    red.addEventListener("cut", (e) => {
        if (!entry || !entry.rich) return;
        const [a, b] = richSel();
        e.preventDefault();
        if (a === b) return;
        if (e.clipboardData) e.clipboardData.setData("text/plain", runsText(entry.runs).slice(a, b));
        richReplace(a, b, "");
    });
    red.addEventListener("copy", (e) => {
        if (!entry || !entry.rich) return;
        const [a, b] = richSel();
        e.preventDefault();
        if (e.clipboardData && b > a) e.clipboardData.setData("text/plain", runsText(entry.runs).slice(a, b));
    });
    // Where the caret is, kept (a format chosen from the toolbar, which takes the focus, is the
    // characters selected then); a caret moved away loses the format waiting for what is typed.
    document.addEventListener("selectionchange", () => {
        if (!entry || !entry.rich || document.activeElement !== red || richComposing) return;
        entry.sel = richSel();
        if (entry.pending && (entry.sel[0] !== entry.pending.at || entry.sel[1] !== entry.pending.at)) entry.pending = null;
        // (The toolbar shows the font of the characters there.)
        if (!richToolsSoon) {
            richToolsSoon = true;
            requestAnimationFrame(() => {
                richToolsSoon = false;
                if (entry && entry.rich) showTools();
            });
        }
    });
    let richToolsSoon = false;

    // A format given to the characters selected (none selected: to those typed next), as
    // `format(set)` gives it to the cells: false when it is not one of a character's.
    const CHAR_FORMATS = {
        bold: (f, v) => (v ? (f.b = true) : delete f.b),
        italic: (f, v) => (v ? (f.i = true) : delete f.i),
        underline: (f, v) => (v && v !== "none" ? (f.u = v === true ? "single" : v) : delete f.u),
        strike: (f, v) => (v ? (f.s = true) : delete f.s),
        fontName: (f, v) => v && (f.name = v),
        fontSize: (f, v) => v && (f.size = v),
        color: (f, v) => (v ? (f.color = v) : delete f.color),
        superscript: (f, v) => (v ? (f.va = "superscript") : f.va === "superscript" && delete f.va),
        subscript: (f, v) => (v ? (f.va = "subscript") : f.va === "subscript" && delete f.va),
    };
    function formatChars(set) {
        const keys = Object.keys(set);
        if (!keys.length || !keys.every((k) => CHAR_FORMATS[k]) || !charTarget()) return false;
        if (!entry.rich) toRich();
        const [a, b] = document.activeElement === red || !entry.sel ? richSel() : entry.sel;
        const change = (f) => {
            for (const k of keys) CHAR_FORMATS[k](f, set[k]);
            return f;
        };
        const base = cellFont();
        if (a === b) {
            // What is typed here next.
            entry.pending = { at: a, f: change({ ...(entry.pending && entry.pending.at === a ? entry.pending.f : fontAt(entry.runs, a) || base) }) };
        } else {
            entry.undo.push({ runs: entry.runs, sel: [a, b] });
            entry.redo = [];
            entry.runs = restyleRuns(entry.runs, a, b, change, base);
        }
        red.focus({ preventScroll: true });
        richDraw(a, b);
        richPlace(a, b);
        placeEditor();
        showTools();
        return true;
    }
    // The font of the characters a format is given to (the first selected, or what is typed
    // next), as a run's.
    function charFont() {
        if (!entry.rich) return fontAt(tidyRuns([{ t: entryInput().value, f: null }]), 0) || cellFont();
        const [a, b] = document.activeElement === red || !entry.sel ? richSel() : entry.sel;
        if (a === b && entry.pending && entry.pending.at === a) return entry.pending.f;
        return (a === b ? fontAt(entry.runs, a) : fontOfChar(entry.runs, a)) || cellFont();
    }
    // Ctrl+1 while a cell is typed into: "セルの書式設定" of its characters, the フォント page
    // only (as Excel's): for those selected, or those typed next.
    function askCharFont() {
        if (!entry || !charTarget()) return;
        if (!entry.rich) toRich();
        const kept = richSel();
        entry.sel = kept;
        const f = { ...cellFont(), ...charFont() };
        const form = openDialog("セルの書式設定", "charfont");
        form.appendChild(el("div", "tabs1 tablist")).appendChild(el("span", "on", "フォント"));
        const name = field(form, "フォント名:");
        name.value = f.name || "";
        name.dataset.field = "fontName";
        const styleLine = form.appendChild(el("label"));
        styleLine.appendChild(el("span", "lbl", "スタイル:"));
        const style = styleLine.appendChild(el("select"));
        for (const [v, label] of [["normal", "標準"], ["italic", "斜体"], ["bold", "太字"], ["boldItalic", "太字 斜体"]]) style.appendChild(el("option", null, label)).value = v;
        style.value = f.b && f.i ? "boldItalic" : f.b ? "bold" : f.i ? "italic" : "normal";
        style.dataset.field = "fontStyle";
        const size = field(form, "サイズ:", "number");
        size.min = "1";
        size.max = "409";
        size.step = "0.5";
        size.value = String(f.size || 11);
        size.dataset.field = "fontSize";
        const uLine = form.appendChild(el("label"));
        uLine.appendChild(el("span", "lbl", "下線:"));
        const under = uLine.appendChild(el("select"));
        for (const [v, label] of UNDERLINES) under.appendChild(el("option", null, label)).value = v;
        under.value = f.u || "none";
        under.dataset.field = "underline";
        // 色: the palette's (自動: the cell's color).
        let color = f.color || null;
        const cLine = form.appendChild(el("div", "crow"));
        cLine.appendChild(el("span", "lbl", "色:"));
        const pick = cLine.appendChild(el("button", "pcolor"));
        pick.type = "button";
        pick.dataset.field = "fontColor";
        const swatch = pick.appendChild(el("span", "sw"));
        const label = pick.appendChild(el("span", null));
        const showColor = () => {
            swatch.style.background = color || "#000000";
            label.textContent = color || "自動";
        };
        showColor();
        pick.addEventListener("click", () =>
            openColors(pick, "color", "自動", (c) => {
                color = c;
                showColor();
            }),
        );
        form.appendChild(el("div", "sub", "文字飾り"));
        const strike = check(form, "取り消し線", !!f.s);
        const sup = check(form, "上付き", f.va === "superscript");
        const sub = check(form, "下付き", f.va === "subscript");
        sup.addEventListener("change", () => sup.checked && (sub.checked = false));
        sub.addEventListener("change", () => sub.checked && (sup.checked = false));
        const preview = form.appendChild(el("div", "charpreview"));
        const sample = preview.appendChild(el("span", null, "Aaあぁアァ亜宇"));
        const showSample = () => {
            sample.style.cssText = runCss({
                name: name.value.trim() || f.name,
                size: Math.min(26, +size.value || 11),
                b: /bold/.test(style.value),
                i: /italic/i.test(style.value),
                u: under.value === "none" ? undefined : under.value,
                s: strike.checked,
                color: color || undefined,
                va: sup.checked ? "superscript" : sub.checked ? "subscript" : undefined,
            });
        };
        for (const x of [name, style, size, under, strike, sup, sub]) x.addEventListener("input", showSample);
        pick.addEventListener("click", showSample);
        showSample();
        dialogButtons(form);
        form.cancel = () => {
            closeDialog();
            if (entry) richPlace(kept[0], kept[1]);
        };
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const pt = Number(size.value);
            if (!(pt >= 1 && pt <= 409)) {
                size.setCustomValidity("1 から 409 の数値を指定してください");
                size.reportValidity();
                return;
            }
            closeDialog();
            if (!entry) return;
            entry.sel = kept;
            formatChars({
                fontName: name.value.trim() || f.name,
                fontSize: pt,
                bold: /bold/.test(style.value),
                italic: /italic/i.test(style.value),
                underline: under.value,
                color,
                strike: strike.checked,
                superscript: sup.checked,
                subscript: sub.checked && !sup.checked,
            });
        });
        size.addEventListener("input", () => size.setCustomValidity(""));
        name.focus();
        name.select();
    }
    // The runs a cell is entered with (none: plain text).
    function entryRuns(e, text) {
        if (!e.runs || /^=/.test(text)) return null;
        const runs = tidyRuns(e.runs);
        if (runsText(runs) !== text || !runs.some((r) => r.f)) return null;
        return runs.map((r) => (r.f ? { t: r.t, f: r.f } : { t: r.t }));
    }
    // The dialog open (Excel's), and a dialog over it (the format of a conditional format's rule).
    let dialog = null;
    let subDialog = null;
    // (Not while a dialog is open: it keeps the focus.)
    const focusGrid = () => {
        if (dialog) return;
        if (entry && entry.rich) {
            if (document.activeElement !== red) {
                red.focus({ preventScroll: true });
                if (entry.sel) richPlace(entry.sel[0], entry.sel[1]);
            }
        } else if (!entry || entry.where === "cell") ed.focus({ preventScroll: true });
    };
    const onGrid = () => {
        const a = document.activeElement;
        return !a || a === document.body || a === ed || a === red || a === bar || a === $("grid");
    };

    function placeEditor() {
        if (!view || !sel) return;
        if (ed.parentNode !== view.wrap) view.wrap.insertBefore(ed, view.refs);
        if (red.parentNode !== view.wrap) view.wrap.insertBefore(red, view.refs);
        const o = ownerOf(sel.ar, sel.ac) || [sel.ar, sel.ac, sel.ar, sel.ac];
        const b = boxOf(o[0], o[1], o[2], o[3]);
        if (!b) return;
        const on = !!entry && entry.where === "cell";
        const rich = on && !!entry.rich;
        // (The box typed into: of the characters' formats, or the plain one.)
        const box = rich ? red : ed;
        ed.classList.toggle("on", on && !rich);
        red.classList.toggle("on", rich);
        for (const t of [ed, red]) {
            Object.assign(t.style, { left: `${b.x}px`, top: `${b.y}px`, width: "", height: "", font: "", paddingTop: "", color: "" });
            pinTo(t, pinOf(view, b));
        }
        view.selbox.classList.remove("overx", "overy");
        if (!on) return;
        const style = activeTd ? getComputedStyle(activeTd) : null;
        if (style) box.style.font = style.font;
        if (style && rich) box.style.color = style.color;
        // As wide and as tall as what is typed, the cell at least.
        box.style.width = `${b.w + 1}px`;
        box.style.width = `${Math.max(b.w + 1, box.scrollWidth + 10)}px`;
        // What is typed stands where the cell shows it (at the bottom, unless the cell says
        // otherwise): it does not move when it is entered.
        box.style.paddingTop = "0px";
        box.style.height = "0px";
        const textH = box.scrollHeight;
        const valign = style ? style.verticalAlign : "bottom";
        const room = b.h - textH;
        const pad = room <= 0 || valign === "top" ? 0 : valign === "middle" ? room / 2 : room;
        box.style.paddingTop = `${pad}px`;
        box.style.height = `${pad > 0 ? b.h + 1 : Math.max(b.h + 1, textH + 2)}px`;
        // (Not filled from while typed into: the fill handle is not there.)
        view.fillh.hidden = true;
        // What is typed past the cell runs over the right (bottom) side of its border, as Excel
        // shows it: the border stays on the cell.
        view.selbox.classList.toggle("overx", parseFloat(box.style.width) > b.w + 1);
        view.selbox.classList.toggle("overy", parseFloat(box.style.height) > b.h + 1);
        placePaint();
    }

    // What is typed in a formula, its references in their colors in the cell and in the bar (as
    // Excel colors them while one is typed): a layer over the box typed into, of the same text in
    // the same place, the box's own text not seen (its caret and its selection are).
    const edPaint = el("div", "paint");
    const barPaint = el("div", "paint");
    // While an input method composes, the box shows its own text (the layer would not have it).
    let composing = false;
    function paintEntry(refs) {
        const input = entry && entryInput();
        const on = !!input && input !== richBox && !composing && refs.length > 0;
        for (const t of [ed, bar]) t.classList.toggle("painted", on && t === input);
        const paint = input === bar ? barPaint : edPaint;
        for (const p of [edPaint, barPaint]) if (!on || p !== paint) p.remove();
        if (!on) return;
        if (paint.previousSibling !== input) input.after(paint);
        refSpans(paint, input.value, refs, "");
        placePaint();
    }
    // The layer where the box is, of its font, scrolled as it is.
    function placePaint() {
        const input = entry && entryInput();
        const paint = input === bar ? barPaint : edPaint;
        if (!input || paint.previousSibling !== input) return;
        const s = getComputedStyle(input);
        Object.assign(paint.style, {
            left: `${input.offsetLeft}px`,
            top: `${input.offsetTop}px`,
            width: `${input.offsetWidth}px`,
            height: `${input.offsetHeight}px`,
            padding: s.padding,
            fontFamily: s.fontFamily,
            fontSize: s.fontSize,
            fontWeight: s.fontWeight,
            fontStyle: s.fontStyle,
            lineHeight: s.lineHeight,
            letterSpacing: s.letterSpacing,
            textAlign: s.textAlign,
        });
        // (On the panes frozen, as the box.)
        if (input.dataset.pin) pinTo(paint, input.dataset.pin);
        else pinTo(paint, "");
        paint.scrollLeft = input.scrollLeft;
        paint.scrollTop = input.scrollTop;
    }
    for (const t of [ed, bar]) {
        t.addEventListener("scroll", placePaint);
        t.addEventListener("compositionstart", () => {
            composing = true;
            if (entry) paintEntry([]);
        });
        t.addEventListener("compositionend", () => {
            composing = false;
            if (entry) entryRefs();
        });
    }

    function beginEntry(mode, text, where) {
        if (!editable || busy || !sel || !view) return false;
        if (view.sheet.protection && lockedAt(sel.ar, sel.ac)) {
            showError(PROTECTED);
            return false;
        }
        closeMenu();
        const cell = (activeTd && activeTd.cell) || {};
        const original = cell.f || cell.v || cell.t || "";
        // `point`: the reference being pointed at (clicked, or moved to by the arrows) while a
        // formula is typed: { start, end } in the text, `anchor` and `cur` its corners [r, c].
        entry = { mode, where, original, point: null };
        // The characters' formats of the text edited (not of a text typed over the cell).
        if (text === null && cell.rt && !cell.f) {
            entry.runs = tidyRuns(cell.rt.map((r) => ({ t: r.t, f: r.f || null })));
            entry.originalRuns = JSON.stringify(entry.runs);
            entry.lastText = runsText(entry.runs);
        }
        entry.undo = [];
        entry.redo = [];
        entry.deco = activeTd ? activeTd.style.textDecoration || "" : "";
        if (entry.deco === "none") entry.deco = "";
        if (text === null) text = original;
        const fx = $("fx");
        fx.textContent = "";
        if (where === "bar") {
            fx.appendChild(bar);
            bar.value = text;
            bar.focus();
            bar.setSelectionRange(text.length, text.length);
        } else {
            fx.textContent = text;
            reveal(sel.ar, sel.ac);
            if (entry.runs && entry.runs.some((r) => r.f)) {
                ed.value = "";
                entry.rich = true;
                placeEditor();
                red.focus({ preventScroll: true });
                richDraw(text.length);
                richPlace(text.length);
                placeEditor();
            } else {
                // (What an input method is composing in the box stays as it is.)
                if (ed.value !== text) ed.value = text;
                placeEditor();
                ed.focus({ preventScroll: true });
                ed.setSelectionRange(text.length, text.length);
            }
        }
        document.body.classList.add("typing");
        entryRefs();
        placeList();
        showMode();
        return true;
    }

    function endEntry() {
        const was = entry;
        entry = null;
        richComposing = null;
        red.classList.remove("on");
        red.textContent = "";
        if (was && was.rich && document.activeElement === red) ed.focus({ preventScroll: true });
        closeAssist();
        showMode();
        ed.value = "";
        bar.remove();
        paintEntry([]);
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
        const box = entryInput();
        const text = box.value;
        // What the rule of the cell (データの入力規則) says of it, as Excel checks it: a value it
        // does not allow is asked about (再試行: typed on; キャンセル: undone); one worked out by
        // the editor (a formula, a rule of a formula) is checked there, after it is entered.
        const rule = listAt(sel.ar, sel.ac);
        const verdict = rule && rule.showError && text !== e.original ? judge(rule, text) : null;
        if (verdict === "no") {
            askRefused(rule, (answer) => {
                if (answer === "enter") enter(e, how, text, false);
                else if (answer === "retry") {
                    box.focus();
                    box.select();
                } else cancelEntry();
            });
            return;
        }
        enter(e, how, text, verdict === "editor" && rule.errorStyle === "stop");
    }
    // Enters `text` (`validate`: the editor checks it against the cell's rule).
    function enter(e, how, text, validate) {
        if (!sel || !view) return;
        const fill = !!(how && how.fill) && (!single(sel) || areas.length > 0);
        const at = address(sel.ar, sel.ac);
        const ranges = fill ? allAreas().map(rangeText) : [at];
        const sheet = view.sheet.name;
        // The characters' formats, and whether they changed.
        if (e.runs && e.where === "bar") e.runs = followRuns(e.runs, e.lastText, text);
        const runs = entryRuns(e, text);
        const runsChanged = JSON.stringify(runs ? tidyRuns(runs.map((r) => ({ t: r.t, f: r.f || null }))) : null) !== (e.originalRuns || "null");
        const [ar, ac] = [sel.ar, sel.ac];
        endEntry();
        if (how && how.move) move(how.move[0], how.move[1], false);
        if (text === e.original && !fill && !runsChanged) return;
        const ops = ranges.map((range) => (text === "" ? { op: "clear", sheet, range } : runs ? { op: "set", sheet, range, text, runs } : { op: "set", sheet, range, text }));
        if (validate && text !== "") ops.push({ op: "validate", sheet, range: at });
        // A row Excel fits to its fonts fits to those of the characters too.
        if (runs && view && !(view.sheet.protection && !(view.sheet.protection.allow || {}).formatRows)) {
            const v = view.sheet;
            const cells = (v.cells || []).filter((c) => c.r === ar && !(c.c === ac));
            const css = (c) => (c.s ? v.styles[c.s] : "");
            const mine = cellOf(ar, ac);
            let px = fitOfRuns(v, mine ? css(mine) : cssOf(v, ar, ac, null), runs);
            for (const c of cells) px = Math.max(px, fitOfRuns(v, css(c), c.rt));
            if (!(v.customRows || []).includes(ar) && height(v, ar) > 0 && px > height(v, ar)) ops.push({ op: "fit", sheet, rows: [[ar, px]] });
        }
        postEdit(ops);
    }

    ed.addEventListener("input", () => {
        if (!entry) {
            // The first character typed starts an entry that replaces the cell's content.
            if (!beginEntry("enter", ed.value, "cell")) ed.value = "";
            return;
        }
        // Typing ends the pointing.
        entry.point = null;
        if (entry.runs) {
            entry.runs = followRuns(entry.runs, entry.lastText || "", ed.value);
            entry.lastText = ed.value;
        }
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
            // Goes on in the bar (the characters' formats follow what is typed there).
            const text = entryInput().value;
            if (entry.rich) {
                entry.rich = false;
                entry.lastText = text;
                red.classList.remove("on");
                red.textContent = "";
            }
            entry.where = "bar";
            entry.mode = "edit";
            ed.value = "";
            placeEditor();
            $("fx").textContent = "";
            $("fx").appendChild(bar);
            bar.value = text;
            bar.focus();
            entryRefs();
        } else if (!entry) {
            beginEntry("edit", null, "bar");
        }
    });

    // ---- formulas typed: pointing at cells (a click, a drag, the arrows), F4, the references

    const entryInput = () => (entry && entry.where === "bar" ? bar : entry && entry.rich ? richBox : ed);
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
    // (In the cell and in the bar, the references in the same colors.)
    function entryRefs() {
        if (!view || !entry) return;
        view.refs.textContent = "";
        const text = entryInput().value;
        const refs = /^[=+-]/.test(text) ? colorRefs(view.sheet, text, references(text)) : [];
        const p = entry.point;
        const ref = p && refs.find((r) => r.start === p.start && r.end === p.end);
        if (ref && ref.box) ref.box.classList.add("pointing");
        if (entry.where === "cell") refSpans($("fx"), text, refs, "ref");
        paintEntry(refs);
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
        const r = entryInput().getBoundingClientRect();
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
        if (refusedByProtection(ops)) {
            showError(PROTECTED);
            return false;
        }
        if (runMode) return runEdit(ops);
        if (!keepClip) setClip(null);
        setBusy(true);
        vscode.postMessage({ type: "edit", ops });
        return true;
    }
    function postHistory(type) {
        if (!editable || busy) return;
        if (runMode) {
            showError("実行モードでは、元に戻す・やり直しはできません。");
            return;
        }
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
            setRunMode(!!m.run, m.session);
        } else if (m && m.type === "runDone") {
            runDone(m);
        } else if (m && m.type === "fonts") {
            if (Array.isArray(m.names)) takeFonts(m.names);
        } else if (m && m.type === "idle") {
            // Nothing was done (a question of the editor's was cancelled).
            setBusy(false);
            dragNotMade();
            focusGrid();
        } else if (m && m.type === "error") {
            setBusy(false);
            dragNotMade();
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
    function closeDialog() {
        closeSubDialog();
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
    const asText = (text) => (/^[=+\-@']/.test(text) || (text.trim() !== "" && !isNaN(+text.replace(/[,¥$%\s]/g, ""))) || /^\d+[/-]\d+/.test(text) || /^\s*(明治|大正|昭和|平成|令和|[明大昭平令MTSHRmtshr])(\d+|元)[/.\-年]/.test(text) ? `'${text}` : text);

    // A dialog of the page: its form, its buttons (OK and キャンセル), closed by a click beside it.
    // `cancel`: what closing it without OK does (a dialog opened from another goes back to it).
    function openDialog(title, cls, cancel) {
        closeMenu();
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", `dlg${cls ? ` ${cls}` : ""}`));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, title));
        form.cancel = cancel || closeDialog;
        dialog.addEventListener("mousedown", (e) => {
            if (e.target === dialog) form.cancel();
        });
        return form;
    }
    function dialogButtons(form, okText) {
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", okText || "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", () => (form.cancel || closeDialog)());
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
        // Headings guessed as Excel 2021 guesses them: a column with a text (typed or a formula's)
        // over something that is not one (a number, a date, ...).
        let guess = false;
        if (sel.kind === "cells" && s.r2 > s.r1) {
            for (let c = s.c1; c <= s.c2 && !guess; c++) {
                const x = cellOf(s.r1, c);
                const y = cellOf(s.r1 + 1, c);
                if (x && x.k === "string" && y && y.k !== "string" && (y.t ?? "") !== "") guess = true;
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
            // A new link gets Excel's cell style ハイパーリンク (blue, underlined).
            if (!now) ops.push({ op: "format", sheet: sheet.name, range, set: { style: "Hyperlink" } });
            ops.push(op);
            postEdit(ops);
        });
        addr.addEventListener("input", () => addr.setCustomValidity(""));
        ref.addEventListener("input", () => ref.setCustomValidity(""));
        (now && place.checked ? ref : web.checked ? addr : shown).focus();
    }
    // "ハイパーリンクの削除": the links of the selection go, and the cells get the style 標準.
    function removeLinks() {
        if (!editable || busy || !sel || !view) return;
        const sheet = view.sheet.name;
        postEdit(allAreas().flatMap((a) => [{ op: "link", sheet, range: rangeText(a), off: true }, { op: "format", sheet, range: areaText(a), set: { style: "Normal" } }]));
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
    // The size Excel 2021 gives a shape put with a click (a line goes across a square of 96).
    const clickSize = (kind) =>
        ({ textBox: [144, 36], triangle: [111.36, 96], rightArrow: [102.72, 50.88], leftArrow: [102.72, 50.88], upArrow: [50.88, 102.72], downArrow: [50.88, 102.72] })[kind] || [96, 96];
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
    // An ActiveX control of Microsoft Forms (one whose settings the page has).
    const isActiveX = (obj) => !!obj && obj.kind === "control" && obj.shapeId !== undefined && obj.activex !== undefined && !!obj.form;
    // A control: by its shape; its settings and place change, not its format.
    const isControl = (obj) => isForm(obj) || isActiveX(obj);
    const movable = (obj) => editable && !!obj && (isControl(obj) || (!!obj.name && obj.kind !== "control"));
    // How an edit names an object: a form control by its shape.
    const target = (obj) => (isControl(obj) ? { shape: obj.shapeId } : { name: obj.name });
    // What tells an object from the others after the page is drawn again.
    const keyOf = (obj) => (isControl(obj) ? `#${obj.shapeId}` : obj.name || "");

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
    // The ActiveX controls (Microsoft Forms 2.0) of 開発 > 挿入, in Excel's order: [kind, name,
    // icon (16 × 16)]. Their events are then handled in the sheet's file (`@on`).
    const ACTIVEX_CONTROLS = [
        ["commandButton", "コマンド ボタン", '<rect x="1.5" y="4" width="13" height="8" rx="1"/><path d="M4 10.5h8" stroke-opacity=".4"/>'],
        ["comboBox", "コンボ ボックス", '<rect x="1.5" y="5" width="13" height="6"/><path d="M10.5 5v6M11.8 7.3l.7.9.7-.9"/>'],
        ["checkBox", "チェック ボックス", '<rect x="2.5" y="4.5" width="7" height="7"/><path d="m3.8 8 1.7 1.7 2.7-3.4"/><path d="M11 8h3.5" stroke-opacity=".5"/>'],
        ["listBox", "リスト ボックス", '<rect x="1.5" y="2.5" width="13" height="11"/><path d="M3.5 5.5h6M3.5 8h6M3.5 10.5h6M11 2.5v11"/>'],
        ["textBox", "テキスト ボックス", '<rect x="1.5" y="4.5" width="13" height="7"/><path d="M4 6.5v3"/>'],
        ["scrollBar", "スクロール バー", '<rect x="1.5" y="5" width="13" height="6"/><path d="M4 5v6M12 5v6M6.5 7h3v2h-3z"/>'],
        ["spinButton", "スピン ボタン", '<rect x="5" y="1.5" width="6" height="13"/><path d="M5 8h6M6.8 5.3 8 4l1.2 1.3M6.8 10.7 8 12l1.2-1.3"/>'],
        ["optionButton", "オプション ボタン", '<circle cx="6" cy="8" r="3.5"/><circle cx="6" cy="8" r="1.3" fill="currentColor"/><path d="M11 8h3.5" stroke-opacity=".5"/>'],
        ["label", "ラベル", '<text x="2" y="12" font-size="10" stroke="none" fill="currentColor" font-family="sans-serif">A</text>'],
        ["toggleButton", "トグル ボタン", '<rect x="1.5" y="4" width="13" height="8" rx="1"/><rect x="3.5" y="6" width="9" height="4" fill="currentColor" fill-opacity=".35" stroke="none"/>'],
    ];
    // What is drawn: `control:<kind>` a form control, `activeX:<kind>` an ActiveX control.
    const isControlKind = (kind) => kind.startsWith("control:") || kind.startsWith("activeX:");

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
    // `key`: the object's (`keyOf`), or null for the one the edit adds. (One of several
    // selected: all of them stay selected.)
    function pickAfter(key) {
        if (!view) return;
        const all = selectedObjects().map(keyOf);
        const keys = key ? (all.includes(key) ? all : [key]) : null;
        remember({ pickObj: { sheet: view.sheet.name, keys, before: key ? null : (view.sheet.objects || []).map(keyOf) } });
    }
    function pickObject(v) {
        const p = saved.pickObj;
        if (!p) return;
        remember({ pickObj: null });
        if (p.sheet !== v.sheet.name) return;
        const objects = v.sheet.objects || [];
        const keys = p.keys || (p.key ? [p.key] : null);
        const found = keys ? keys.map((k) => objects.find((o) => keyOf(o) === k)).filter((o) => o && o.box) : [objects.find((o) => keyOf(o) && !(p.before || []).includes(keyOf(o)))].filter((o) => o && o.box);
        if (!found.length) return;
        // (The one selected last is selected again last.)
        const rest = found.slice(1);
        rest.forEach((obj, i) => selectObject(obj.box, obj, i > 0));
        selectObject(found[0].box, found[0], rest.length > 0);
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
        if (kind.startsWith("activeX:")) return `${(ACTIVEX_CONTROLS.find((c) => `activeX:${c[0]}` === kind) || [, kind])[1]}（ActiveX コントロール）`;
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
                    putControl(d.kind, x0, y0, null);
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
            putControl(kind, Math.min(x0, x1), Math.min(y0, y1), { width: Math.round(Math.abs(x1 - x0)), height: Math.round(Math.abs(y1 - y0)) });
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
    // A form control (`control:<kind>`) or an ActiveX control (`activeX:<kind>`) put where it was
    // drawn (`size` null: Excel's size for it).
    function putControl(drawn, x, y, size) {
        const [how, kind] = drawn.split(":");
        const op = { op: how === "activeX" ? "activeX" : "control", sheet: view.sheet.name, control: kind, ...anchorOf(x, y), ...(size || {}) };
        // A form scroll bar drawn wider than high goes across (an ActiveX one turns by its shape).
        if (how === "control" && kind === "scrollBar" && size && size.width > size.height) op.horizontal = true;
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
    // F2, a double click: the text in the pane of the object (a double click: all of it
    // selected).
    function editObjectText(obj, all) {
        if (!movable(obj) || obj.text === undefined || busy || !view) return;
        if (obj.textBox && obj.box) editShapeText(obj, all);
        else if (obj.kind === "control" && obj.box && controlTextSpan(obj)) editControlText(obj, all);
        else openPane("format", "text");
    }
    // The element a control shows its text in (a caption, a label, the value of a text or combo
    // box); none for a list box, whose value is not shown as a text.
    function controlTextSpan(obj) {
        const ctl = obj.box && obj.box.querySelector(".ctl");
        if (!ctl || ctl.classList.contains("list")) return null;
        return ctl.querySelector(":scope > span.label, :scope > span.value, :scope > span:not(.btn):not(.box):not(.radio)");
    }
    // Whether a control's text is a value typed on one line (a text or combo box; a form
    // control's edit box): Enter sets it.
    const oneLineText = (obj) => {
        const t = obj.form && obj.form.type;
        return ["TextBox", "ComboBox", "EditBox", "Drop"].includes(t);
    };
    // The text of a control typed on the control, as Excel's text editing of a button or a
    // label (and an ActiveX control's caption or value in design mode): where it shows its text,
    // in its font. Esc, Ctrl+Enter or a click elsewhere sets it (Enter too, for a value on one
    // line; else Enter starts a new line).
    function editControlText(obj, all) {
        if (shapeEditor) return;
        const box = obj.box;
        const span = controlTextSpan(obj);
        if (!span) return;
        const ctl = span.parentNode;
        const ta = el("textarea", "ctledit");
        ta.spellcheck = false;
        ta.rows = 1;
        ta.value = obj.text;
        if (oneLineText(obj)) ta.wrap = "off";
        const cs = getComputedStyle(span);
        Object.assign(ta.style, { fontFamily: cs.fontFamily, fontSize: cs.fontSize, fontWeight: cs.fontWeight, fontStyle: cs.fontStyle, color: cs.color, textAlign: getComputedStyle(ctl).textAlign });
        // (In the place of its text: the box of a check box stays.)
        if (span.classList.contains("label") && ctl.classList.contains("gbox")) ta.classList.add("gbox");
        span.style.display = "none";
        ctl.insertBefore(ta, span);
        box.classList.add("typing");
        const fit = () => {
            ta.style.height = "auto";
            ta.style.height = `${Math.min(ta.scrollHeight, ctl.clientHeight || ta.scrollHeight)}px`;
        };
        let done = false;
        const finish = (back) => {
            if (done) return;
            done = true;
            shapeEditor = null;
            const text = ta.value;
            ta.remove();
            span.style.display = "";
            box.classList.remove("typing");
            if (text !== obj.text) {
                // (Shown at once, as it will be drawn.)
                span.textContent = text;
                objectEdit(obj, { text });
            }
            if (back) focusGrid();
        };
        ta.addEventListener("input", fit);
        ta.addEventListener("blur", () => finish(false));
        ta.addEventListener("keydown", (e) => {
            e.stopPropagation();
            if (e.key === "Escape" || (e.key === "Enter" && (e.ctrlKey || e.metaKey || (oneLineText(obj) && !e.isComposing)))) {
                e.preventDefault();
                finish(true);
            }
        });
        for (const type of ["mousedown", "click", "dblclick", "contextmenu"]) ta.addEventListener(type, (e) => e.stopPropagation());
        shapeEditor = { obj, ta };
        fit();
        const place = () => {
            if (done) return;
            ta.focus({ preventScroll: true });
            if (all) ta.select();
            else ta.setSelectionRange(ta.value.length, ta.value.length);
        };
        place();
        requestAnimationFrame(() => document.activeElement !== ta && place());
    }
    // The text of a shape typed in the shape, as Excel's: over its text, in its font, where its
    // text box puts it. Enter starts a new line; Esc or a click elsewhere sets it. `all`: its
    // text selected (a double click), else the caret after it (shown in an empty shape too).
    let shapeEditor = null;
    function editShapeText(obj, all) {
        if (shapeEditor) return;
        const box = obj.box;
        const shape = box.querySelector(":scope > .shape, :scope > .refl > .shape") || box;
        const tb = obj.textBox;
        // (A warped text is put straight while it is typed.)
        const old = shape.querySelector(".stext, .warp");
        const cover = shape.appendChild(el("div", "stedit"));
        const [l, t, r, b] = tb.inset;
        const pad = old && old.classList.contains("stext") ? old.style.padding : `${t}pt ${r}pt ${b}pt ${l}pt`;
        const just = { t: "flex-start", ctr: "center", b: "flex-end" }[tb.anchor] || "flex-start";
        Object.assign(cover.style, { padding: pad, justifyContent: just, color: tb.color });
        const ta = cover.appendChild(el("textarea"));
        ta.spellcheck = false;
        // (One line high when empty: the caret where the text box puts its text.)
        ta.rows = 1;
        ta.value = obj.text;
        if (!tb.wrap) ta.wrap = "off";
        // The font of its first run, the alignment of its first paragraph.
        const run = old && old.querySelector("span");
        const para = old && old.querySelector("p");
        if (run) {
            const cs = getComputedStyle(run);
            Object.assign(ta.style, { fontFamily: cs.fontFamily, fontSize: cs.fontSize, fontWeight: cs.fontWeight, fontStyle: cs.fontStyle });
            // (A text filled with a gradient or a picture is see-through: typed in its style's color.)
            if (!/^rgba\(.*,\s*0\)$|^transparent$/.test(cs.color)) ta.style.color = cs.color;
        }
        ta.style.textAlign = para ? para.style.textAlign || "left" : "center";
        if (old) old.style.visibility = "hidden";
        box.classList.add("typing");
        const fit = () => {
            ta.style.height = "auto";
            ta.style.height = `${ta.scrollHeight}px`;
        };
        let done = false;
        const finish = (back) => {
            if (done) return;
            done = true;
            shapeEditor = null;
            const text = ta.value;
            cover.remove();
            box.classList.remove("typing");
            if (old) old.style.visibility = "";
            if (text !== obj.text) objectEdit(obj, { text });
            if (back) focusGrid();
        };
        ta.addEventListener("input", fit);
        ta.addEventListener("blur", () => finish(false));
        ta.addEventListener("keydown", (e) => {
            e.stopPropagation();
            if (e.key === "Escape" || (e.key === "Enter" && (e.ctrlKey || e.metaKey))) {
                e.preventDefault();
                finish(true);
            }
        });
        for (const type of ["mousedown", "click", "dblclick", "contextmenu"]) ta.addEventListener(type, (e) => e.stopPropagation());
        shapeEditor = { obj, ta };
        fit();
        const place = () => {
            if (done) return;
            ta.focus({ preventScroll: true });
            if (all) ta.select();
            else ta.setSelectionRange(ta.value.length, ta.value.length);
        };
        place();
        // (Again once the double click and what the selection redraws are over, so the focus
        // and the caret stay in the shape.)
        requestAnimationFrame(() => document.activeElement !== ta && place());
    }
    function deleteObject(obj) {
        if (!movable(obj) || busy || !view) return;
        // (One of several selected: all of them.)
        const all = isSelected(obj) ? selectedObjects().filter(movable) : [obj];
        postEdit(all.map((o) => ({ op: "object", sheet: view.sheet.name, ...target(o), delete: true })));
    }
    // The box of an object as it is placed (px at 100%; as placed, not rounded as offsetLeft is,
    // which would move it by half a pixel each time).
    const placed = (v, rounded) => (Number.isFinite(parseFloat(v)) ? parseFloat(v) : rounded);
    const boxOfObject = (box) => ({ x: placed(box.style.left, box.offsetLeft), y: placed(box.style.top, box.offsetTop), w: placed(box.style.width, box.offsetWidth), h: placed(box.style.height, box.offsetHeight) });
    // An object moved (its box) to where it now is, or sized.
    function moveObject(obj, b) {
        moveObjects([[obj, b]]);
    }
    // Objects moved (or sized) at once: [[obj, box]].
    function moveObjects(moves) {
        if (!view || !moves.length) return false;
        pickAfter(keyOf(moves[0][0]));
        return postEdit(moves.map(([obj, b]) => ({ op: "object", sheet: view.sheet.name, ...target(obj), ...anchorOf(Math.max(0, b.x), Math.max(0, b.y)), width: Math.max(0, Math.round(b.w)), height: Math.max(0, Math.round(b.h)) })));
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
        // 調整ハンドル: where the shape's adjust values put them.
        const def = obj.adjust;
        if (!def) return;
        const [w, h] = [box.offsetWidth, box.offsetHeight];
        if (!(w > 0 && h > 0)) return;
        const g = shapeGuides(def, adjustValues(def), w * EMU_PX, h * EMU_PX);
        def.ah.forEach((ah, i) => {
            let [x, y] = handlePoint(g, ah);
            if (obj.flipH) x = w - x;
            if (obj.flipV) y = h - y;
            const d = frame.appendChild(el("div", "objhandle adj"));
            Object.assign(d.style, { left: `${x}px`, top: `${y}px` });
            d.title = "調整ハンドル（図形の形を変える）";
            d.addEventListener("mousedown", (e) => {
                if (e.button !== 0) return;
                e.preventDefault();
                e.stopPropagation();
                beginObjectDrag(e, box, obj, "adj");
                if (objDrag) objDrag.handle = i;
            });
        });
    }

    // ---- 調整ハンドル: the adjust values of a preset shape (its avLst), evaluated as DrawingML
    // does (`obj.adjust`: the preset's adjust values, guides, handles and paths, from
    // ECMA-376), changed by dragging its handles

    // EMU in a pixel (at 100%: 96 a inch).
    const EMU_PX = 9525;
    // The adjust values of a shape: the preset's, with those the shape has.
    function adjustValues(def) {
        const out = {};
        for (const [name, f] of def.av) {
            const v = +String(f).replace(/^val\s+/, "");
            out[name] = Number.isFinite(v) ? v : 0;
        }
        return Object.assign(out, def.adj || {});
    }
    // The guides of a shape of w × h (EMU) for adjust values: what its handles and paths refer to.
    function shapeGuides(def, adj, w, h) {
        const v = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss: Math.min(w, h), ls: Math.max(w, h) };
        for (let d = 2; d <= 32; d++) [v[`wd${d}`], v[`hd${d}`], v[`ssd${d}`]] = [w / d, h / d, v.ss / d];
        Object.assign(v, { cd2: 10800000, cd4: 5400000, cd8: 2700000, "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000 });
        Object.assign(v, adj);
        for (const [name, f] of def.gd) {
            const x = guideValue(v, f);
            v[name] = Number.isFinite(x) ? x : 0;
        }
        return v;
    }
    // A guide's formula (as DrawingML's guides: angles in 60000ths of a degree).
    function guideValue(v, fmla) {
        const [op, ...args] = String(fmla).trim().split(/\s+/);
        const a = args.map((s) => guideOf(v, s));
        const x = (i) => a[i] || 0;
        const ang = (t) => ((t / 60000) * Math.PI) / 180;
        const div = (p, q) => (q === 0 ? 0 : p / q);
        switch (op) {
            case "val": return x(0);
            case "*/": return div(x(0) * x(1), x(2));
            case "+-": return x(0) + x(1) - x(2);
            case "+/": return div(x(0) + x(1), x(2));
            case "?:": return x(0) > 0 ? x(1) : x(2);
            case "abs": return Math.abs(x(0));
            case "at2": return ((Math.atan2(x(1), x(0)) * 180) / Math.PI) * 60000;
            case "cat2": return x(0) * Math.cos(Math.atan2(x(2), x(1)));
            case "sat2": return x(0) * Math.sin(Math.atan2(x(2), x(1)));
            case "cos": return x(0) * Math.cos(ang(x(1)));
            case "sin": return x(0) * Math.sin(ang(x(1)));
            case "tan": return x(0) * Math.tan(ang(x(1)));
            case "max": return Math.max(x(0), x(1));
            case "min": return Math.min(x(0), x(1));
            case "mod": return Math.sqrt(x(0) * x(0) + x(1) * x(1) + x(2) * x(2));
            case "pin": return x(1) < x(0) ? x(0) : x(1) > x(2) ? x(2) : x(1);
            case "sqrt": return Math.sqrt(Math.max(0, x(0)));
            default: return 0;
        }
    }
    const guideOf = (v, s) => {
        if (s === undefined || s === null) return 0;
        const n = Number(s);
        return Number.isFinite(n) ? n : v[s] || 0;
    };
    // Where a handle is (px from the shape's box, before it is flipped).
    const handlePoint = (g, ah) => [guideOf(g, ah.pos[0]) / EMU_PX, guideOf(g, ah.pos[1]) / EMU_PX];
    // The adjust values a handle moves ([name, least, greatest]: x and y, or radius and angle).
    function handleRefs(g, ah) {
        const out = [];
        for (const k of ["x", "y", "r", "a"]) {
            const ref = ah[k];
            if (!ref) continue;
            let [lo, hi] = [guideOf(g, ref[1]), guideOf(g, ref[2])];
            if (lo > hi) [lo, hi] = [hi, lo];
            out.push({ name: ref[0], lo, hi });
        }
        return out;
    }
    // The adjust values that bring a handle nearest a point (px from the shape's box): the
    // handle's place is any function of them, so they are looked for (a grid of values, then
    // finer ones around the best).
    function adjustTo(def, adj, ah, w, h, px, py) {
        const [we, he] = [w * EMU_PX, h * EMU_PX];
        const refs = handleRefs(shapeGuides(def, adj, we, he), ah);
        if (!refs.length) return adj;
        const [tx, ty] = [px * EMU_PX, py * EMU_PX];
        const trial = Object.assign({}, adj);
        const cost = (vals) => {
            refs.forEach((r, i) => (trial[r.name] = vals[i]));
            const g = shapeGuides(def, trial, we, he);
            const [x, y] = [guideOf(g, ah.pos[0]) - tx, guideOf(g, ah.pos[1]) - ty];
            return x * x + y * y;
        };
        // (Handles without limits, as a callout's: within ten times the shape around where the
        // value is now.)
        const span = refs.map((r) => {
            const now = adj[r.name] || 0;
            return [Math.max(r.lo, now - 1e6), Math.min(r.hi, now + 1e6)];
        });
        const n = refs.length === 1 ? 64 : 24;
        let best = refs.map((r) => adj[r.name] || 0);
        let bestCost = cost(best);
        let steps = span.map(([lo, hi]) => (hi - lo) / n);
        const scan = (ranges, count) => {
            const grid = ranges.map(([lo, hi]) => Array.from({ length: count + 1 }, (_, i) => lo + ((hi - lo) * i) / count));
            const each = (i, vals) => {
                if (i === grid.length) {
                    const c = cost(vals);
                    if (c < bestCost) [best, bestCost] = [vals.slice(), c];
                    return;
                }
                for (const x of grid[i]) {
                    vals[i] = x;
                    each(i + 1, vals);
                }
            };
            each(0, []);
        };
        scan(span, n);
        for (let round = 0; round < 8; round++) {
            const ranges = best.map((b, i) => [Math.max(span[i][0], b - steps[i]), Math.min(span[i][1], b + steps[i])]);
            scan(ranges, 8);
            steps = steps.map((s) => s / 4);
        }
        const out = Object.assign({}, adj);
        refs.forEach((r, i) => (out[r.name] = Math.round(Math.min(r.hi, Math.max(r.lo, best[i])))));
        return out;
    }
    // The outline of a shape of w × h (px) for adjust values: SVG path data in px (arcs as lines).
    function shapeOutline(def, adj, w, h) {
        const [we, he] = [w * EMU_PX, h * EMU_PX];
        const g = shapeGuides(def, adj, we, he);
        let d = "";
        for (const p of def.path) {
            const kx = (p.w ? we / p.w : 1) / EMU_PX;
            const ky = (p.h ? he / p.h : 1) / EMU_PX;
            const pt = (x, y) => `${(x * kx).toFixed(2)},${(y * ky).toFixed(2)}`;
            let cur = [0, 0];
            let start = [0, 0];
            for (const c of p.c) {
                const v = c.slice(1).map((s) => guideOf(g, s));
                if (c[0] === "M") {
                    cur = start = [v[0], v[1]];
                    d += `M${pt(...cur)} `;
                } else if (c[0] === "L") {
                    cur = [v[0], v[1]];
                    d += `L${pt(...cur)} `;
                } else if (c[0] === "Q") {
                    d += `Q${pt(v[0], v[1])} ${pt(v[2], v[3])} `;
                    cur = [v[2], v[3]];
                } else if (c[0] === "C") {
                    d += `C${pt(v[0], v[1])} ${pt(v[2], v[3])} ${pt(v[4], v[5])} `;
                    cur = [v[4], v[5]];
                } else if (c[0] === "A") {
                    // An arc of the ellipse wR × hR from the angle stAng by swAng (angles as seen).
                    const [wr, hr, st, sw] = v;
                    const rad = (t) => ((t / 60000) * Math.PI) / 180;
                    const param = (t) => Math.atan2(wr * Math.sin(t), hr * Math.cos(t));
                    const p1 = param(rad(st));
                    const center = [cur[0] - wr * Math.cos(p1), cur[1] - hr * Math.sin(p1)];
                    const turns = Math.trunc(sw / 21600000);
                    const rem = rad(sw) - turns * 2 * Math.PI;
                    let p2 = param(rad(st + sw));
                    if (rem > 0) {
                        while (p2 <= p1) p2 += 2 * Math.PI;
                        while (p2 - p1 > 2 * Math.PI) p2 -= 2 * Math.PI;
                    } else if (rem < 0) {
                        while (p2 >= p1) p2 -= 2 * Math.PI;
                        while (p1 - p2 > 2 * Math.PI) p2 += 2 * Math.PI;
                    } else p2 = p1;
                    p2 += turns * 2 * Math.PI;
                    if (!d) d += `M${pt(...cur)} `;
                    const steps = Math.min(1000, Math.max(1, Math.ceil(Math.abs(p2 - p1) / (Math.PI / 48))));
                    for (let i = 1; i <= steps; i++) {
                        const q = p1 + ((p2 - p1) * i) / steps;
                        cur = [center[0] + wr * Math.cos(q), center[1] + hr * Math.sin(q)];
                        d += `L${pt(...cur)} `;
                    }
                } else if (c[0] === "Z") {
                    d += "Z ";
                    cur = start;
                }
            }
        }
        return d.trim();
    }
    // The point of the mouse in a shape's box (px at 100%; before it is turned and flipped).
    function pointInShape(e, d) {
        const z = Z();
        const f = d.from;
        const [dx, dy] = [(e.clientX - d.center[0]) / z, (e.clientY - d.center[1]) / z];
        const turn = ((d.obj.rot || 0) * Math.PI) / 180;
        const [cos, sin] = [Math.cos(turn), Math.sin(turn)];
        let [x, y] = [dx * cos + dy * sin + f.w / 2, -dx * sin + dy * cos + f.h / 2];
        if (d.obj.flipH) x = f.w - x;
        if (d.obj.flipV) y = f.h - y;
        return [x, y];
    }
    // The paths of a shape as drawn (`shape_html`), with the preset's paths they are drawn from
    // ([drawn path, its index in `def.path`]) and the size of its SVG (pt); null when they do
    // not match (then the outline is shown instead).
    function drawnPaths(box, def) {
        const svg = box.querySelector(":scope > .shape > svg, :scope > .refl > .shape > svg");
        const drawn = svg ? [...svg.querySelectorAll(":scope > g > path")] : [];
        const vb = svg && svg.viewBox.baseVal;
        if (!drawn.length || !vb || !(vb.width > 0 && vb.height > 0)) return null;
        // (Paths neither filled nor stroked are not drawn: with the shape's fill or line none.)
        const all = def.path.map((p, i) => i);
        const tries = [all, all.filter((i) => def.path[i].stroke !== false), all.filter((i) => def.path[i].fill !== "none"), all.filter((i) => def.path[i].stroke !== false && def.path[i].fill !== "none")];
        const which = tries.find((t) => t.length === drawn.length);
        if (!which) return null;
        return { paths: drawn.map((el, k) => [el, which[k]]), w: vb.width, h: vb.height, was: drawn.map((el) => el.getAttribute("d")) };
    }
    // A handle dragged: the shape drawn as its new adjust values make it (its handles with it).
    function dragAdjust(e, d) {
        const def = d.obj.adjust;
        const f = d.from;
        const [x, y] = pointInShape(e, d);
        if (!d.adj) d.adj = adjustValues(def);
        const adj = adjustTo(def, d.adj, def.ah[d.handle], f.w, f.h, x, y);
        d.now = { adjust: adj };
        if (d.drawn === undefined) d.drawn = drawnPaths(d.box, def);
        if (d.drawn) {
            const { paths, w, h } = d.drawn;
            for (const [el, i] of paths) el.setAttribute("d", shapeOutline({ ...def, path: [def.path[i]] }, adj, w, h));
        } else {
            if (!d.ghost) {
                d.ghost = makeGhost(false);
                d.ghost.classList.add("adj");
            }
            ghostAt(d.ghost, f.x, f.y, f.x + f.w, f.y + f.h, false);
            const flip = d.obj.flipH || d.obj.flipV ? ` transform="translate(${d.obj.flipH ? f.w : 0},${d.obj.flipV ? f.h : 0}) scale(${d.obj.flipH ? -1 : 1},${d.obj.flipV ? -1 : 1})"` : "";
            d.ghost.innerHTML = `<svg width="${f.w}" height="${f.h}" viewBox="0 0 ${f.w} ${f.h}"><path${flip} d="${shapeOutline(def, adj, f.w, f.h)}"/></svg>`;
            if (d.obj.rot) d.ghost.style.transform = `rotate(${d.obj.rot}deg)`;
        }
        // The handles where the new values put them.
        const g = shapeGuides(def, adj, f.w * EMU_PX, f.h * EMU_PX);
        d.box.querySelectorAll(".objhandle.adj").forEach((el, i) => {
            if (!def.ah[i]) return;
            let [hx, hy] = handlePoint(g, def.ah[i]);
            if (d.obj.flipH) hx = f.w - hx;
            if (d.obj.flipV) hy = f.h - hy;
            Object.assign(el.style, { left: `${hx}px`, top: `${hy}px` });
        });
    }
    // The shape as it was (a drag that changed nothing, or an edit not made).
    function undoAdjust(d) {
        if (d.drawn) d.drawn.paths.forEach(([el], k) => el.setAttribute("d", d.drawn.was[k]));
        const def = d.obj.adjust;
        const f = d.from;
        const g = shapeGuides(def, d.adj || adjustValues(def), f.w * EMU_PX, f.h * EMU_PX);
        d.box.querySelectorAll(".objhandle.adj").forEach((el, i) => {
            if (!def.ah[i]) return;
            let [hx, hy] = handlePoint(g, def.ah[i]);
            if (d.obj.flipH) hx = f.w - hx;
            if (d.obj.flipV) hy = f.h - hy;
            Object.assign(el.style, { left: `${hx}px`, top: `${hy}px` });
        });
    }
    // The adjust values a handle was dragged to (only those that changed).
    function endAdjust(d) {
        const before = d.adj || {};
        const change = {};
        for (const [k, v] of Object.entries(d.now.adjust)) if (Math.round(before[k]) !== v) change[k] = v;
        if (Object.keys(change).length && movable(d.obj) && !busy && view) objectEdit(d.obj, { adjust: change });
        else undoAdjust(d);
    }
    function beginObjectDrag(e, box, obj, how) {
        if (busy || !view) return;
        if (entry) commit(null);
        const r = box.getBoundingClientRect();
        objDrag = { obj, box, how, x: e.clientX, y: e.clientY, from: boxOfObject(box), center: [r.left + r.width / 2, r.top + r.height / 2], ghost: null, now: null };
        // Moved: the other objects selected with it too.
        if (how === "move" && isSelected(obj)) objDrag.group = selectedObjects().filter((o) => o !== obj && movable(o) && o.box).map((o) => ({ obj: o, from: boxOfObject(o.box), ghost: null }));
    }
    function dragObject(e) {
        const d = objDrag;
        if (!(e.buttons & 1)) {
            endObjectDrag();
            return;
        }
        const z = Z();
        let [dx, dy] = [(e.clientX - d.x) / z, (e.clientY - d.y) / z];
        if (!d.now && Math.abs(dx) < 3 && Math.abs(dy) < 3) return;
        if (d.how === "adj") {
            dragAdjust(e, d);
            return;
        }
        const f = d.from;
        let { x, y, w, h } = f;
        const turn = ((d.obj.rot || 0) * Math.PI) / 180;
        if (d.how === "rot") {
            // The angle of the mouse around the object's center (0: above it).
            let deg = (Math.atan2(e.clientY - d.center[1], e.clientX - d.center[0]) * 180) / Math.PI + 90;
            if (e.shiftKey) deg = Math.round(deg / 15) * 15;
            deg = ((Math.round(deg) % 360) + 360) % 360;
            // The object itself turned so. (It is drawn turned as it is: turned by the rest, after
            // where frozen panes pin it.)
            if (d.pin === undefined) d.pin = d.box.style.transform;
            d.now = { rot: deg };
            d.box.style.transform = `${d.pin} rotate(${deg - (d.obj.rot || 0)}deg)`;
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
        // 枠線に合わせる (or Alt held, as in Excel): the sides moved go to the cells' borders.
        if ((saved.snapGrid || e.altKey) && !turn) [x, y, w, h] = snapBox(d.how, x, y, w, h);
        d.now = { x, y, w, h };
        if (d.how === "move") {
            // Moved: the object itself (and the others selected, by as much).
            placeBox(d.box, x, y);
            for (const g of d.group || []) {
                const b = g.from;
                g.now = { x: b.x + x - f.x, y: b.y + y - f.y, w: b.w, h: b.h };
                placeBox(g.obj.box, g.now.x, g.now.y);
            }
            return;
        }
        // Sized: only a see-through copy of it as it will be (the object hidden meanwhile).
        if (!d.ghost) {
            d.ghost = copyOfObject(d.box);
            d.box.classList.add("sizing");
        }
        Object.assign(d.ghost.style, { left: `${x}px`, top: `${y}px`, width: `${Math.max(0, w)}px`, height: `${Math.max(0, h)}px` });
        if (d.obj.chart && window.rasChart && !d.redraw) {
            d.redraw = requestAnimationFrame(() => {
                d.redraw = 0;
                if (!d.ghost.isConnected) return;
                try {
                    window.rasChart.draw(d.ghost.querySelector(".chartbox") || d.ghost, d.obj.chart);
                } catch (e) {
                    // (As it was drawn, stretched or not.)
                }
            });
        }
    }
    function dragNotMade() {
        if (undrag) undrag();
        undrag = null;
    }
    function placeBox(box, x, y) {
        Object.assign(box.style, { left: `${x}px`, top: `${y}px` });
    }
    // A copy of an object's box (without its handles), half seen and not in the mouse's way.
    function copyOfObject(box) {
        const c = box.cloneNode(true);
        c.querySelectorAll(".objframe").forEach((f) => f.remove());
        c.classList.remove("on", "framed", "movable");
        c.classList.add("objcopy");
        c.removeAttribute("title");
        return view.layer.appendChild(c);
    }
    // The objects back as they were (a move or a size not made).
    function unmove(d) {
        if (d.ghost) d.ghost.remove();
        d.box.classList.remove("sizing");
        if (d.pin !== undefined) d.box.style.transform = d.pin;
        placeBox(d.box, d.from.x, d.from.y);
        for (const g of d.group || []) placeBox(g.obj.box, g.from.x, g.from.y);
    }
    // The border of a column (`cols`) or a row nearest a point of the sheet.
    function nearestBorder(cols, p) {
        const v = view;
        const at = cols ? (i) => columnLeft(v, i) : (i) => rowTop(v, i);
        const i = lineAt(cols ? v.cols : v.rows.map((row) => row[0]), at, p);
        const [a, b] = [at(i), at(i + 1)];
        return p - a <= b - p ? a : b;
    }
    function snapBox(how, x, y, w, h) {
        if (how === "move") {
            const [l, r] = [nearestBorder(true, x), nearestBorder(true, x + w)];
            x = Math.abs(l - x) <= Math.abs(r - x - w) ? l : r - w;
            const [t, b] = [nearestBorder(false, y), nearestBorder(false, y + h)];
            y = Math.abs(t - y) <= Math.abs(b - y - h) ? t : b - h;
            return [x, y, w, h];
        }
        if (how.includes("w")) [x, w] = [nearestBorder(true, x), x + w - nearestBorder(true, x)];
        if (how.includes("e")) w = nearestBorder(true, x + w) - x;
        if (how.includes("n")) [y, h] = [nearestBorder(false, y), y + h - nearestBorder(false, y)];
        if (how.includes("s")) h = nearestBorder(false, y + h) - y;
        return [x, y, Math.max(0, w), Math.max(0, h)];
    }
    function endObjectDrag() {
        const d = objDrag;
        objDrag = null;
        if (!d || !d.now) return;
        if (d.redraw) cancelAnimationFrame(d.redraw);
        swallowClick = true;
        setTimeout(() => (swallowClick = false));
        const sized = d.how !== "rot" && d.how !== "adj" && d.how !== "move";
        // (A size sent: its copy stays until the sheet is drawn again, as the object is hidden.)
        if (d.ghost && !sized) d.ghost.remove();
        if (d.how === "rot") {
            if (objectEdit(d.obj, { rotation: d.now.rot })) undrag = () => unmove(d);
            else unmove(d);
        } else if (d.how === "adj") endAdjust(d);
        else if (moveObjects([[d.obj, d.now], ...(d.group || []).filter((g) => g.now).map((g) => [g.obj, g.now])])) undrag = () => unmove(d);
        else unmove(d);
    }

    // The arrow keys move the selected objects by a pixel (Ctrl: by 10).
    function nudgeObjects(dr, dc, far) {
        if (busy) return;
        const step = far ? 10 : 1;
        moveObjects(selectedObjects().filter((o) => movable(o) && o.box).map((o) => {
            const b = boxOfObject(o.box);
            return [o, { x: b.x + dc * step, y: b.y + dr * step, w: b.w, h: b.h }];
        }));
    }

    // ---- 図形の書式: the fill, the line, the rotation, the order of the object selected

    // What can be given a fill (a shape, a picture; not a line, a group or a chart), a line, a
    // rotation (not a chart).
    const canFill = (obj) => movable(obj) && (obj.el === "sp" || obj.el === "pic");
    const canLine = (obj) => movable(obj) && (obj.el === "sp" || obj.el === "pic" || obj.el === "cxnSp");
    const canTurn = (obj) => movable(obj) && !!obj.el && obj.el !== "graphicFrame";
    function objectEdit(obj, change) {
        if (!movable(obj) || busy || !view) return false;
        pickAfter(keyOf(obj));
        return postEdit([{ op: "object", sheet: view.sheet.name, ...target(obj), ...change }]);
    }
    // The objects selected changed at once: `changeOf(obj)` is the change of each (null: not
    // that one, as a fill is not a control's).
    function editSelected(changeOf) {
        if (busy || !view) return;
        const ops = [];
        for (const obj of selectedObjects()) {
            const change = movable(obj) && changeOf(obj);
            if (change) ops.push({ op: "object", sheet: view.sheet.name, ...target(obj), ...change });
        }
        if (!ops.length) return;
        pickAfter(keyOf(selectedObj));
        postEdit(ops);
    }
    // Turned by `by` degrees (clockwise), as 右へ 90 度回転 turns it.
    const turnChange = (obj, by) => canTurn(obj) && { rotation: ((((obj.rot || 0) + by) % 360) + 360) % 360 };
    // 上下反転 / 左右反転: flipped as it is seen, a turned object turned the other way too.
    function flipChange(obj, vertical) {
        if (!canTurn(obj)) return null;
        const change = vertical ? { flipV: !obj.flipV } : { flipH: !obj.flipH };
        if (obj.rot) change.rotation = (360 - obj.rot) % 360;
        return change;
    }

    // ---- 配置 (オブジェクトの配置): the objects selected lined up by their sides or centers, or
    // spread evenly between the first and the last (as Excel, by the boxes they are seen in,
    // turned)

    // The box an object is seen in (turned as it is): its left, top, right and bottom.
    function seenBox(obj) {
        const b = boxOfObject(obj.box);
        const t = ((obj.rot || 0) * Math.PI) / 180;
        const [c, s] = [Math.abs(Math.cos(t)), Math.abs(Math.sin(t))];
        const [hw, hh] = [(b.w * c + b.h * s) / 2, (b.w * s + b.h * c) / 2];
        const [cx, cy] = [b.x + b.w / 2, b.y + b.h / 2];
        return { box: b, l: cx - hw, t: cy - hh, r: cx + hw, b: cy + hh };
    }
    const alignable = () => selectedObjects().filter((o) => movable(o) && o.box);
    // `how`: left, center, right, top, middle, bottom; spreadH, spreadV (three or more).
    function alignObjects(how) {
        const list = alignable().map((obj) => ({ obj, ...seenBox(obj) }));
        if (list.length < (how.startsWith("spread") ? 3 : 2) || busy) return;
        const min = (k) => Math.min(...list.map((o) => o[k]));
        const max = (k) => Math.max(...list.map((o) => o[k]));
        const across = how === "left" || how === "center" || how === "right" || how === "spreadH";
        const [lo, hi] = across ? ["l", "r"] : ["t", "b"];
        const to = new Map();
        if (how === "left" || how === "top") for (const o of list) to.set(o, min(lo));
        else if (how === "right" || how === "bottom") for (const o of list) to.set(o, max(hi) - (o[hi] - o[lo]));
        else if (how === "center" || how === "middle") for (const o of list) to.set(o, (min(lo) + max(hi)) / 2 - (o[hi] - o[lo]) / 2);
        else {
            // The first and the last stay; the space between the others is the same.
            const order = list.slice().sort((a, b) => a[lo] + a[hi] - (b[lo] + b[hi]));
            const [first, last] = [order[0], order[order.length - 1]];
            const sizes = order.reduce((n, o) => n + o[hi] - o[lo], 0);
            const gap = (last[hi] - first[lo] - sizes) / (order.length - 1);
            let at = first[lo];
            for (const o of order) {
                to.set(o, at);
                at += o[hi] - o[lo] + gap;
            }
        }
        const moves = [];
        for (const o of list) {
            const by = to.get(o) - o[lo];
            if (Math.abs(by) < 0.5) continue;
            const b = o.box;
            moves.push([o.obj, across ? { ...b, x: b.x + by } : { ...b, y: b.y + by }]);
        }
        if (moves.length) moveObjects(moves);
    }
    const ORDERS = [
        ["front", "最前面へ移動"],
        ["forward", "前面へ移動"],
        ["backward", "背面へ移動"],
        ["back", "最背面へ移動"],
    ];
    const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
    // 配置's items: [how, icon, label, title], by group.
    const ALIGNS = [
        [
            ["left", "alignLeft", "左揃え", "選択したオブジェクトの左端を揃えます"],
            ["center", "alignCenter", "左右中央揃え", "選択したオブジェクトを左右の中央で揃えます"],
            ["right", "alignRight", "右揃え", "選択したオブジェクトの右端を揃えます"],
        ],
        [
            ["top", "alignTop", "上揃え", "選択したオブジェクトの上端を揃えます"],
            ["middle", "alignMiddle", "上下中央揃え", "選択したオブジェクトを上下の中央で揃えます"],
            ["bottom", "alignBottom", "下揃え", "選択したオブジェクトの下端を揃えます"],
        ],
        [
            ["spreadH", "spreadH", "左右に整列", "選択したオブジェクト（3 つ以上）を左右に等間隔で並べます"],
            ["spreadV", "spreadV", "上下に整列", "選択したオブジェクト（3 つ以上）を上下に等間隔で並べます"],
        ],
    ];
    // The contextual tab of the ribbon: shown while an object that can be changed is selected
    // (a shape, a picture, a chart, a control).
    function refreshShapeTools() {
        renderPane();
        const ribbon = $("ribbon");
        if (!ribbon || !ribbon.showShapeTab) return;
        const obj = selected && movable(selectedObj) ? selectedObj : null;
        ribbon.showShapeTab(!!obj);
        if (!obj) return;
        // (Several selected: what any of them can be given.)
        const all = alignable();
        tools.shapeFill.disabled = tools.shapeFillMore.disabled = !all.some(canFill);
        tools.shapeLine.disabled = tools.shapeLineMore.disabled = !all.some(canLine);
        for (const key of ["forward", "forwardMore", "backward", "backwardMore"]) tools[key].disabled = all.every(isControl);
        tools.rotate.disabled = !all.some(canTurn);
        const b = boxOfObject(selected);
        tools.objHeight.setCm(cmOf(b.h));
        tools.objWidth.setCm(cmOf(b.w));
    }
    function cmOf(px) {
        return Math.round(((px * 2.54) / 96) * 100) / 100;
    }
    // A box of centimeters with its spin buttons (as Excel's サイズ): `set` takes what is typed.
    // (`unit`: another unit, its step and its limit; 線の幅 is points.)
    function sizeBox(into, label, set, unit) {
        const { name = "cm", step = 0.1, min = 0, max = 1000 } = unit || {};
        const wrap = into.appendChild(el("span", "spinbox"));
        const input = wrap.appendChild(el("input"));
        input.type = "text";
        input.title = label;
        input.setAttribute("aria-label", label);
        input.spellcheck = false;
        input.autocomplete = "off";
        let now = null;
        const show = () => (input.value = now === null ? "" : !name || name === "°" || name === "%" ? `${now}${name}` : `${now} ${name}`);
        const take = (cm) => {
            if (!isFinite(cm)) return show();
            cm = Math.round(Math.min(Math.max(cm, min), max) * 100) / 100;
            if (cm === now) return show();
            now = cm;
            show();
            set(cm);
        };
        input.addEventListener("change", () => take(parseFloat(input.value)));
        input.addEventListener("keydown", (e) => {
            if (e.key === "ArrowUp" || e.key === "ArrowDown") {
                e.preventDefault();
                take((parseFloat(input.value) || 0) + (e.key === "ArrowUp" ? step : -step));
            } else if (e.key === "Escape") {
                e.preventDefault();
                show();
                focusGrid();
            } else if (e.key === "Enter") {
                // (Not the grid's Enter too, which would move the active cell.)
                e.preventDefault();
                e.stopPropagation();
                input.blur();
                focusGrid();
            }
        });
        const spins = wrap.appendChild(el("span", "spins"));
        for (const [cls, by, label] of [["chev up", step, "増やす"], ["chev", -step, "減らす"]]) {
            const b = spins.appendChild(el("button"));
            b.appendChild(el("span", cls));
            b.setAttribute("aria-label", label);
            b.type = "button";
            b.tabIndex = -1;
            b.addEventListener("click", () => take((parseFloat(input.value) || 0) + by));
        }
        wrap.setCm = (cm) => {
            now = cm;
            if (document.activeElement !== input) show();
        };
        wrap.input = input;
        return wrap;
    }

    // ---- the pane at the right: the format of the object selected (as the form designer's
    // properties), or the objects of the sheet (オブジェクトの選択と表示)
    // `saved.pane`: "format", "list", or null (closed); the format is shown until it is closed.

    function paneMode() {
        return saved.pane === undefined ? "format" : saved.pane;
    }
    function openPane(mode, focus) {
        remember({ pane: mode });
        renderPane(focus);
    }
    function closePane() {
        remember({ pane: null });
        renderPane();
        focusGrid();
    }
    function objectLabel(obj) {
        // (A form control may have no name: its text, as it shows.)
        return obj.name || (isControl(obj) && obj.text ? obj.text : obj.shapeId !== undefined ? `_x0000_s${obj.shapeId}` : "（名前なし）");
    }
    function renderPane(focus) {
        let pane = $("pane");
        const mode = paneMode();
        const obj = selected && movable(selectedObj) ? selectedObj : null;
        // "cells": セルの書式設定, which gives way to the format of an object selected.
        const shape = !!obj && (mode === "format" || mode === "cells");
        const cells = !obj && mode === "cells" && !!sel;
        const on = editable && !!view && (mode === "list" || shape || cells);
        if (!on) {
            if (pane) pane.hidden = true;
            return;
        }
        if (!pane) {
            pane = el("div");
            pane.id = "pane";
            pane.dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: true });
            $("main").insertBefore(pane, $("grid").nextSibling);
        }
        // (A field being typed into keeps what is typed until the object changes; the cells'
        // formats change only as the page is drawn again.)
        const key = mode === "list" ? "list" : shape ? `format:${view.sheet.name}:${keyOf(obj)}` : `cells:${view.sheet.name}:${[sel.r1, sel.c1, sel.r2, sel.c2, sel.ar, sel.ac, areas.length].join(",")}`;
        if (!pane.hidden && pane.dataset.key === key && (cells || pane.contains(document.activeElement)) && !focus) {
            if (mode === "list") markListed(pane);
            return;
        }
        pane.hidden = false;
        pane.dataset.key = key;
        pane.textContent = "";
        paneGrip(pane);
        const head = pane.appendChild(el("div", "phd"));
        const title = mode === "list" ? "選択" : cells ? "セルの書式設定" : isActiveX(obj) ? "プロパティ" : isControl(obj) ? "コントロールの書式設定" : "図形の書式設定";
        head.appendChild(el("span", "ptitle", title));
        const close = head.appendChild(el("button", "pclose", "×"));
        close.type = "button";
        close.title = "閉じる";
        close.addEventListener("click", closePane);
        const body = pane.appendChild(el("div", "pbody"));
        if (mode === "list") listPane(body);
        else if (cells) cellPane(body, focus);
        else formatPane(body, obj, focus);
        // Where it was scrolled to, as long as it shows the same (the page is drawn again at
        // each edit).
        const scroll = saved.paneScroll;
        if (scroll && scroll.key === key) body.scrollTop = scroll.top;
        body.addEventListener("scroll", () => remember({ paneScroll: { key, top: body.scrollTop } }), { passive: true });
        const target = focus && body.querySelector(`[data-field="${focus}"]`);
        if (target) {
            target.focus({ preventScroll: !!(scroll && scroll.key === key) });
            if (target.select && target.tagName === "INPUT") target.select();
        }
    }
    // The left edge of the pane, dragged: its width (kept with the page's state; a double
    // click: the width it had at first).
    const PANE_WIDTH = 270;
    function paneGrip(pane) {
        const fit = (w) => Math.round(Math.max(220, Math.min(w, Math.max(220, window.innerWidth * 0.6))));
        pane.style.width = `${fit(saved.paneWidth || PANE_WIDTH)}px`;
        const grip = pane.appendChild(el("div", "pgrip"));
        grip.title = "ドラッグで幅を変更（ダブルクリックで元の幅）";
        grip.addEventListener("pointerdown", (e) => {
            if (e.button !== 0) return;
            e.preventDefault();
            grip.setPointerCapture(e.pointerId);
            const [x0, w0] = [e.clientX, pane.getBoundingClientRect().width];
            document.body.classList.add("panesizing");
            // (The grid fills what it is given more of.)
            const move = (ev) => {
                pane.style.width = `${fit(w0 + x0 - ev.clientX)}px`;
                window.dispatchEvent(new Event("resize"));
            };
            const up = () => {
                grip.removeEventListener("pointermove", move);
                grip.removeEventListener("pointerup", up);
                grip.removeEventListener("pointercancel", up);
                document.body.classList.remove("panesizing");
                remember({ paneWidth: parseFloat(pane.style.width) });
            };
            grip.addEventListener("pointermove", move);
            grip.addEventListener("pointerup", up);
            grip.addEventListener("pointercancel", up);
        });
        grip.addEventListener("dblclick", () => {
            pane.style.width = `${PANE_WIDTH}px`;
            remember({ paneWidth: PANE_WIDTH });
        });
    }
    // The objects of the sheet, the front one first (as Excel lists them); a click selects one.
    function listPane(body) {
        const objects = (view.sheet.objects || []).filter((o) => o.name || o.shapeId !== undefined);
        if (!objects.length) {
            body.appendChild(el("div", "pnone", "このシートには図形・コントロールがありません"));
            return;
        }
        const list = body.appendChild(el("div", "plist"));
        for (const obj of objects.slice().reverse()) {
            const b = list.appendChild(el("button", "pitem", objectLabel(obj)));
            b.type = "button";
            b.obj = obj;
            b.title = obj.macro ? `${objectLabel(obj)}（${eventName(obj)}イベント: ${obj.macro}）` : objectLabel(obj);
            b.addEventListener("click", (e) => {
                if (!obj.box) return;
                // (Ctrl or Shift: added to the objects selected, or taken out of them.)
                selectObject(obj.box, obj, e.ctrlKey || e.metaKey || e.shiftKey);
                obj.box.scrollIntoView({ block: "nearest", inline: "nearest" });
                focusGrid();
            });
        }
        markListed(body.parentNode);
    }
    function markListed(pane) {
        for (const b of pane.querySelectorAll(".pitem")) b.classList.toggle("on", isSelected(b.obj));
    }
    // The format of an object, as Excel's 図形の書式設定: its pages by the icons of one row of tabs
    // (文字, 塗りつぶしと線, 効果, サイズとプロパティ: 文字 before 塗りつぶし, as the pages of
    // セルの書式設定 are), their sections folded by their headings; each value is set as it is
    // changed.
    const PANE_ICONS = {
        fill: '<path d="m2.5 7.5 5-5 5 5-5 5z"/><path d="M4 9h7"/><path d="M13.3 9.5c.8 1.2 1.2 2 1.2 2.5a1.2 1.2 0 0 1-2.4 0c0-.5.4-1.3 1.2-2.5z"/><path d="M1.5 15h13" stroke-width="2"/>',
        effects: '<path d="M8 1.8 14.2 6.3 11.8 13.5H4.2L1.8 6.3z"/>',
        size: '<rect x="5" y="5" width="7" height="7"/><path d="M2.5 1.5v13M1.5 2.5h2M1.5 13.5h2M5 2.5h9M13 1.5v2M7.5 8.5h2M8.5 7.5v2"/>',
        text: '<path d="M4 11.5 8 2l4 9.5M5.4 8.3h5.2"/><path d="M2 14.5h12" stroke-width="2"/>',
        code: '<path d="M5.5 4 1.5 8l4 4M10.5 4l4 4-4 4M9.2 2.5l-2.4 11"/>',
    };
    const PANE_PAGES = [
        ["text", "文字"],
        ["fill", "塗りつぶしと線"],
        ["effects", "効果"],
        ["size", "サイズとプロパティ"],
        ["code", "コード"],
    ];
    // The sections folded (by their heading), as long as the page is open.
    const paneFolded = new Set();
    // The properties of an object as the object model (VBA's Shape) has them, as Chrome's
    // developer tools list an element's: a name and its value, an object (Fill, Line, ...)
    // opened by its triangle; what is opened stays so while the pane is.
    const propsOpen = new Set();
    const MSO_SHAPES = { rect: 1, parallelogram: 2, trapezoid: 3, diamond: 4, roundRect: 5, octagon: 6, triangle: 7, rtTriangle: 8, ellipse: 9, hexagon: 10, plus: 11, pentagon: 51, chevron: 52, star5: 92, rightArrow: 33, leftArrow: 34, upArrow: 35, downArrow: 36, can: 13, cube: 14, smileyFace: 17, donut: 18, heart: 21, lightningBolt: 22, sun: 23, moon: 24, cloud: 179, wedgeRectCallout: 105, wedgeRoundRectCallout: 106, wedgeEllipseCallout: 107 };
    function objectModel(obj) {
        const pt = (px) => Math.round(px * 0.75 * 100) / 100;
        const b = obj.box ? boxOfObject(obj.box) : null;
        const p = obj.paint || {};
        const pr = obj.props || {};
        const f = obj.form || {};
        const ax = isActiveX(obj);
        const enumOf = (name, n) => ({ enum: name, n });
        const rgb = (c) => (c ? { color: c } : undefined);
        const tri = (on) => enumOf(on ? "msoTrue" : "msoFalse", on ? -1 : 0);
        const cell = (m) => (m ? { range: `$${colName(m[0])}$${m[2] + 1}` } : undefined);
        const type = obj.kind === "chart" ? enumOf("msoChart", 3) : obj.el === "pic" ? enumOf("msoPicture", 13) : obj.el === "grpSp" || obj.kind === "group" ? enumOf("msoGroup", 6) : obj.el === "cxnSp" ? enumOf("msoLine", 9) : ax ? enumOf("msoOLEControlObject", 12) : isControl(obj) ? enumOf("msoFormControl", 8) : enumOf("msoAutoShape", 1);
        const m = {
            Name: obj.name || "",
            Type: type,
        };
        if (obj.adjust && obj.adjust.prst) {
            const prst = obj.adjust.prst;
            m.AutoShapeType = MSO_SHAPES[prst] ? enumOf(`msoShape${prst[0].toUpperCase()}${prst.slice(1)}`, MSO_SHAPES[prst]) : prst;
        }
        if (b) Object.assign(m, { Left: pt(b.x), Top: pt(b.y), Width: pt(b.w), Height: pt(b.h) });
        m.Rotation = obj.rot || 0;
        if (canTurn(obj)) Object.assign(m, { HorizontalFlip: tri(!!obj.flipH), VerticalFlip: tri(!!obj.flipV) });
        m.Visible = tri(true);
        if (obj.props) {
            m.Placement = { twoCell: enumOf("xlMoveAndSize", 1), oneCell: enumOf("xlMove", 2), absolute: enumOf("xlFreeFloating", 3) }[pr.placement];
            m.Locked = !!pr.locked;
            if (isControl(obj)) m.PrintObject = !!pr.print;
        }
        if (isControl(obj) && f.altText !== undefined) m.AlternativeText = f.altText || "";
        m.TopLeftCell = cell(obj.from);
        m.BottomRightCell = cell(obj.to);
        if (obj.onAction) m.OnAction = obj.macro || "";
        const z = (view.sheet.objects || []).indexOf(obj);
        if (z >= 0) m.ZOrderPosition = z + 1;
        if ("fill" in p && obj.el !== "cxnSp") m.Fill = { Visible: tri(!!p.fill), ForeColor: { RGB: rgb(p.fill || "#FFFFFF") } };
        if ("line" in p) m.Line = { Visible: tri(!!p.line), ForeColor: { RGB: rgb(p.line || "#000000") }, Weight: p.lineWidth };
        if ("shadow" in p) {
            const sh = p.shadow;
            const rad = sh ? (sh.angle * Math.PI) / 180 : 0;
            m.Shadow = sh
                ? { Visible: tri(true), Type: enumOf("msoShadowOuter", 25), ForeColor: { RGB: rgb(sh.color) }, Transparency: sh.transparency / 100, Blur: sh.blur, OffsetX: Math.round(sh.distance * Math.cos(rad) * 100) / 100, OffsetY: Math.round(sh.distance * Math.sin(rad) * 100) / 100 }
                : { Visible: tri(false) };
            m.Glow = p.glow ? { Radius: p.glow.size, Color: { RGB: rgb(p.glow.color) }, Transparency: p.glow.transparency / 100 } : { Radius: 0 };
            m.Reflection = p.reflection ? { Type: enumOf("msoReflectionTypeMixed", -2), Transparency: p.reflection.transparency / 100, Size: p.reflection.size, Offset: p.reflection.distance, Blur: p.reflection.blur } : { Type: enumOf("msoReflectionTypeNone", 0) };
        }
        const tb = obj.textBox;
        if (tb) {
            m.TextFrame2 = {
                HasText: tri(!!obj.text),
                TextRange: { Text: obj.text || "" },
                VerticalAnchor: { t: enumOf("msoAnchorTop", 1), ctr: enumOf("msoAnchorMiddle", 3), b: enumOf("msoAnchorBottom", 4) }[tb.anchor],
                WordWrap: tri(!!tb.wrap),
                MarginLeft: tb.inset[0],
                MarginTop: tb.inset[1],
                MarginRight: tb.inset[2],
                MarginBottom: tb.inset[3],
            };
        }
        if (obj.adjust && obj.adjust.adj) {
            const vals = Object.values(obj.adjust.adj);
            if (vals.length) {
                const a = { Count: vals.length };
                vals.forEach((v, i) => (a[`Item(${i + 1})`] = Math.round((v / 100000) * 10000) / 10000));
                m.Adjustments = a;
            }
        }
        if (isControl(obj) && obj.form) {
            const t = f.type;
            const c = {};
            if (["CheckBox", "Radio", "OptionButton", "ToggleButton"].includes(t)) c.Value = ax ? !!f.checked : f.checked ? enumOf("xlOn", 1) : enumOf("xlOff", -4146);
            if (f.link !== undefined) c.LinkedCell = f.link || "";
            if (["List", "Drop", "ListBox", "ComboBox"].includes(t)) c.ListFillRange = f.items || "";
            if (t === "List" || t === "Drop") c.ListIndex = f.sel || 0;
            if (["Spin", "Scroll", "SpinButton", "ScrollBar"].includes(t)) {
                Object.assign(c, { Value: f.val || 0, Min: f.min || 0, Max: f.max || 0, SmallChange: f.inc || 1 });
                if (t === "Scroll" || t === "ScrollBar") c.LargeChange = f.page || 1;
            }
            if (obj.text !== undefined) c[ax ? (["TextBox", "ListBox", "ComboBox"].includes(t) ? "Text" : "Caption") : "Caption"] = obj.text;
            if (ax && f.enabled !== undefined) {
                if (f.font) c.Font = { Name: f.font.name, Size: f.font.size, Bold: !!f.font.bold, Italic: !!f.font.italic, Underline: !!f.font.underline, Strikethrough: !!f.font.strike };
                c.ForeColor = rgb(f.foreColor);
                c.BackColor = rgb(f.backColor);
                if (!["SpinButton", "ScrollBar"].includes(t)) c.BackStyle = f.opaque ? enumOf("fmBackStyleOpaque", 1) : enumOf("fmBackStyleTransparent", 0);
                c.Enabled = !!f.enabled;
                if (!["ListBox", "ComboBox", "SpinButton", "ScrollBar"].includes(t)) c.WordWrap = !!f.wordWrap;
                if (!["ListBox", "SpinButton", "ScrollBar"].includes(t)) c.AutoSize = !!f.autoSize;
            }
            if (ax) m.OLEFormat = { ProgId: `Forms.${obj.activex || t}.1`, Object: c };
            else m.ControlFormat = c;
        }
        return m;
    }
    function propertyTree(sec, obj) {
        const tree = sec.appendChild(el("div", "pprops"));
        tree.setAttribute("role", "tree");
        const value = (v) => {
            const span = el("span", "pv");
            if (v && v.color) {
                span.appendChild(el("span", "psw")).style.background = v.color;
                const n = parseInt(v.color.slice(1), 16);
                span.appendChild(document.createTextNode(`RGB(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`));
            } else if (v && v.enum) {
                span.textContent = `${v.enum} (${v.n})`;
                span.classList.add("penum");
            } else if (v && v.range) {
                span.textContent = v.range;
            } else if (typeof v === "string") {
                span.textContent = JSON.stringify(v);
                span.classList.add("pstr");
            } else if (typeof v === "boolean") {
                span.textContent = v ? "True" : "False";
                span.classList.add("pbool");
            } else {
                span.textContent = String(v);
                span.classList.add("pnum");
            }
            return span;
        };
        const leaf = (v) => v === null || typeof v !== "object" || v.color || v.enum || v.range;
        const add = (into, model, path, depth) => {
            for (const [k, v] of Object.entries(model)) {
                if (v === undefined) continue;
                const row = into.appendChild(el("div", "prop"));
                row.style.paddingLeft = `${depth * 12 + 4}px`;
                row.setAttribute("role", "treeitem");
                const tw = row.appendChild(el("span", "ptw"));
                // (The values in one column, whatever the depth.)
                row.appendChild(el("span", "pk", k)).style.minWidth = `${Math.max(40, 112 - depth * 12)}px`;
                if (leaf(v)) {
                    row.appendChild(value(v)).title = row.lastChild.textContent;
                    continue;
                }
                const key = `${path}.${k}`;
                const kids = into.appendChild(el("div", "pkids"));
                const open = (on) => {
                    row.classList.toggle("open", on);
                    row.setAttribute("aria-expanded", String(on));
                    kids.hidden = !on;
                    if (on && !kids.firstChild) add(kids, v, key, depth + 1);
                };
                // (Its first values, as the tools preview an object.)
                const brief = (x) => (leaf(x) ? (x && x.color ? x.color : x && x.enum ? x.enum : x && x.range ? x.range : typeof x === "string" ? JSON.stringify(x) : String(x)) : "{…}");
                const preview = `{${Object.entries(v)
                    .filter(([, x]) => x !== undefined)
                    .map(([n, x]) => `${n}: ${brief(x)}`)
                    .join(", ")}}`;
                row.appendChild(el("span", "pv pobj", preview)).title = preview;
                row.classList.add("pbranch");
                tw.classList.add("chev", "right");
                row.addEventListener("click", () => {
                    const on = !row.classList.contains("open");
                    if (on) propsOpen.add(key);
                    else propsOpen.delete(key);
                    open(on);
                });
                open(propsOpen.has(key));
            }
        };
        add(tree, objectModel(obj), "", 0);
    }
    // The parts of a pane of formats (図形の書式設定, セルの書式設定): its pages, shown by the icons
    // of one row of tabs, their sections folded by their headings, and the controls in them.
    function paneKit(body) {
        const tabs = body.parentNode.insertBefore(el("div", "ptabs"), body);
        tabs.setAttribute("role", "tablist");
        const pages = {};
        const page = (key) => {
            if (!pages[key]) pages[key] = body.appendChild(el("div", "ppage"));
            return pages[key];
        };
        const section = (pg, label) => {
            const sec = page(pg).appendChild(el("section", "psec"));
            const h = sec.appendChild(el("button", "phead"));
            h.type = "button";
            h.appendChild(el("span", "pchev"));
            h.appendChild(el("span", null, label));
            const inner = sec.appendChild(el("div", "pinner"));
            const fold = (on) => {
                sec.classList.toggle("folded", on);
                h.setAttribute("aria-expanded", String(!on));
            };
            fold(paneFolded.has(label));
            h.addEventListener("click", () => {
                const on = !sec.classList.contains("folded");
                if (on) paneFolded.add(label);
                else paneFolded.delete(label);
                fold(on);
            });
            return inner;
        };
        const line = (sec, label) => {
            const row = sec.appendChild(el("label", "prow"));
            row.appendChild(el("span", "plabel", label));
            return row;
        };
        // A box committed when it changes (Enter: back to the grid; Esc: what it had).
        const box = (sec, label, field, value, take, type) => {
            const input = line(sec, label).appendChild(el("input"));
            input.type = type || "text";
            input.dataset.field = field;
            input.spellcheck = false;
            input.autocomplete = "off";
            input.value = value;
            input.addEventListener("change", () => take(input.value));
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    e.stopPropagation();
                    input.blur();
                    focusGrid();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    input.value = value;
                    focusGrid();
                }
            });
            return input;
        };
        // Options of one choice, as Excel's radio buttons (塗りつぶしなし, 塗りつぶし (単色), ...).
        let radios = 0;
        const choices = (sec, list, now, pick) => {
            const name = `pr${++radios}`;
            const wrap = sec.appendChild(el("div", "pradios"));
            for (const [value, label] of list) {
                const row = wrap.appendChild(el("label", "pradio"));
                const r = row.appendChild(el("input"));
                r.type = "radio";
                r.name = name;
                r.checked = value === now;
                row.appendChild(el("span", null, label));
                r.addEventListener("change", () => r.checked && pick(value));
            }
        };
        // The button of a color and its palette: `color` null is `none` (自動, 色なし), when
        // the palette has it.
        const colorRow = (sec, label, color, none, set, field) => {
            const row = sec.appendChild(el("div", "prow"));
            row.appendChild(el("span", "plabel", label));
            const cb = row.appendChild(el("button", "pcolor"));
            cb.type = "button";
            cb.title = `${label}の選択`;
            if (field) cb.dataset.field = field;
            const sw = cb.appendChild(el("span", "sw"));
            if (color) sw.style.background = color;
            else {
                // (自動 is black; 色なし is crossed out.)
                if (none === "自動") sw.style.background = "#000000";
                else sw.classList.add("none");
                cb.appendChild(el("span", "pctext", none || ""));
            }
            cb.appendChild(el("span", "pdrop chev"));
            cb.addEventListener("click", () => openColors(cb, null, none, set));
            return cb;
        };
        const checkRow = (sec, label, on, field, take) => {
            const row = sec.appendChild(el("label", "prow pcheck"));
            const input = row.appendChild(el("input"));
            input.type = "checkbox";
            input.checked = !!on;
            input.dataset.field = field;
            row.appendChild(el("span", null, label));
            input.addEventListener("change", () => take(input.checked));
            return input;
        };
        // A list to choose from: [value, label] each.
        const selectRow = (sec, label, list, now, field, take, cls) => {
            const pick = line(sec, label).appendChild(el("select", cls || null));
            pick.dataset.field = field;
            for (const [value, text] of list) pick.appendChild(el("option", null, text)).value = value;
            pick.value = now;
            pick.addEventListener("change", () => take(pick.value));
            return pick;
        };
        // A block of a label over what it has (サンプル, 種類), and a box of a number with its
        // spin buttons.
        const block = (sec, label) => {
            const b = sec.appendChild(el("div", "pblock"));
            if (label) b.appendChild(el("div", "plabel", label));
            return b;
        };
        const spin = (sec, label, field, now, take, unit) => {
            const b = sizeBox(line(sec, label), label, take, unit);
            b.input.dataset.field = field;
            b.setCm(now);
            return b;
        };
        // The tabs of the pages it has (`order`: [key, label], `icons` by key): the one shown
        // last (`savedKey` in the page's state), or that of the field asked for.
        const finish = (order, icons, savedKey, focus) => {
            const shown = order.filter(([key]) => pages[key]);
            const target = focus && body.querySelector(`[data-field="${focus}"]`);
            let now = target ? Object.keys(pages).find((k) => pages[k].contains(target)) : pages[focus] ? focus : saved[savedKey];
            if (target && target.closest(".psec")) target.closest(".psec").classList.remove("folded");
            if (!pages[now]) now = shown[0] && shown[0][0];
            const show = (key) => {
                for (const [k] of shown) {
                    pages[k].hidden = k !== key;
                    tabs.querySelector(`[data-page="${k}"]`).classList.toggle("on", k === key);
                    tabs.querySelector(`[data-page="${k}"]`).setAttribute("aria-selected", String(k === key));
                }
            };
            for (const [key, label] of shown) {
                const tb = tabs.appendChild(el("button", "ptab"));
                tb.type = "button";
                tb.dataset.page = key;
                tb.title = label;
                tb.setAttribute("role", "tab");
                tb.setAttribute("aria-label", label);
                tb.innerHTML = `<svg viewBox="0 0 16 16" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[key]}</svg>`;
                tb.addEventListener("click", () => {
                    remember({ [savedKey]: key });
                    show(key);
                });
            }
            if (now) {
                if (pages[focus]) remember({ [savedKey]: now });
                show(now);
            }
        };
        return { tabs, pages, page, section, line, box, choices, colorRow, checkRow, selectRow, block, spin, finish };
    }
    function formatPane(body, obj, focus) {
        const ax = isActiveX(obj);
        const f = obj.form || {};
        const t = f.type;
        const paint = obj.paint || {};
        const edit = (change) => objectEdit(obj, change);
        const kit = paneKit(body);
        const { section, line, box, choices } = kit;
        body.appendChild(el("div", "pname", objectLabel(obj)));
        // 色: the button of a color and its palette (the color used last when it has none).
        const colorRow = (sec, key, color, none, set) =>
            kit.colorRow(sec, "色", color || lastShape[key], none, (c) => {
                if (c) lastShape[key] = c;
                set(c);
            });

        // 塗りつぶしと線
        if (canFill(obj)) {
            const sec = section("fill", "塗りつぶし");
            const fill = paint.fill || null;
            choices(sec, [["none", "塗りつぶしなし"], ["solid", "塗りつぶし (単色)"]], fill ? "solid" : "none", (v) => edit({ fill: v === "none" ? null : fill || lastShape.fill }));
            if (fill) colorRow(sec, "fill", fill, "塗りつぶしなし", (c) => edit({ fill: c }));
        }
        // An ActiveX control's 背景 (BackStyle, BackColor): 透明 or 不透明 in its color.
        if (ax && f.backColor !== undefined && t !== "SpinButton" && t !== "ScrollBar") {
            const sec = section("fill", "背景（BackStyle・BackColor）");
            choices(sec, [["transparent", "透明（fmBackStyleTransparent）"], ["opaque", "不透明（fmBackStyleOpaque）"]], f.opaque ? "opaque" : "transparent", (v) => edit({ backStyle: v }));
            if (f.opaque) kit.colorRow(sec, "色", f.backColor || null, "既定の色", (c) => edit({ backColor: c }), "backColor");
        } else if (ax && f.backColor !== undefined) {
            kit.colorRow(section("fill", "背景（BackColor）"), "色", f.backColor || null, "既定の色", (c) => edit({ backColor: c }), "backColor");
        }
        if (canLine(obj)) {
            const sec = section("fill", "線");
            const color = paint.line || null;
            choices(sec, [["none", "線なし"], ["solid", "線 (単色)"]], color ? "solid" : "none", (v) => edit({ line: v === "none" ? null : color || lastShape.line }));
            if (color) {
                colorRow(sec, "line", color, "線なし", (c) => edit({ line: c }));
                const w = sizeBox(line(sec, "幅"), "線の幅", (pt) => edit({ lineWidth: pt }), { name: "pt", step: 0.25, max: 1584 });
                w.input.dataset.field = "lineWidth";
                w.setCm(paint.lineWidth ?? 0.75);
            }
        }

        // 効果: those the grid draws (影 (外側), 反射, 光彩), each by its 標準スタイル and its
        // values; the others the file has (内側の影, ぼかし, 3-D, ...) stay as they are.
        if (canLine(obj) && "shadow" in paint) {
            const effectSection = (key, label, presets, values) => {
                const sec = section("effects", label);
                const now = paint[key] || null;
                const row = line(sec, "標準スタイル");
                const pick = row.appendChild(el("select", "pwide"));
                pick.dataset.field = `${key}-preset`;
                const same = (p) => now && Object.keys(p).every((k) => k === "color" || Math.abs((now[k] ?? 0) - p[k]) < 0.01);
                let chosen = now ? "" : "none";
                presets.forEach(([label, p], i) => {
                    pick.appendChild(el("option", null, label)).value = String(i);
                    if (!chosen && same(p)) chosen = String(i);
                });
                pick.insertBefore(el("option", null, "なし"), pick.firstChild).value = "none";
                if (!chosen) {
                    const custom = pick.appendChild(el("option", null, "ユーザー設定"));
                    custom.value = "";
                    custom.disabled = true;
                }
                pick.value = chosen;
                pick.addEventListener("change", () => {
                    if (pick.value === "none") return edit({ [key]: null });
                    const p = { ...presets[+pick.value][1] };
                    if (key !== "reflection") p.color = (now && now.color) || lastShape[key];
                    edit({ [key]: p });
                });
                if (!now) return;
                for (const [field, label, unit] of values) {
                    if (field === "color") {
                        colorRow(sec, key, now.color, null, (c) => c && edit({ [key]: { color: c } }));
                        continue;
                    }
                    const b = sizeBox(line(sec, label), label, (v) => edit({ [key]: { [field]: v } }), unit);
                    b.input.dataset.field = `${key}-${field}`;
                    b.setCm(now[field] ?? 0);
                }
            };
            const pct = { name: "%", step: 1, max: 100 };
            const pts = (max) => ({ name: "pt", step: 1, max });
            const shadow = (angle, distance = 3, blur = 4) => ({ transparency: 60, blur, angle, distance });
            effectSection(
                "shadow",
                "影",
                [
                    ["オフセット: 右下", shadow(45)],
                    ["オフセット: 下", shadow(90)],
                    ["オフセット: 左下", shadow(135)],
                    ["オフセット: 右", shadow(0)],
                    ["オフセット: 中央", shadow(0, 0, 5)],
                    ["オフセット: 左", shadow(180)],
                    ["オフセット: 右上", shadow(315)],
                    ["オフセット: 上", shadow(270)],
                    ["オフセット: 左上", shadow(225)],
                ],
                [
                    ["color", "色"],
                    ["transparency", "透明度", pct],
                    ["blur", "ぼかし", pts(100)],
                    ["angle", "角度", { name: "°", step: 1, max: 359 }],
                    ["distance", "距離", pts(200)],
                ],
            );
            effectSection(
                "reflection",
                "反射",
                [
                    ["反射 (弱): 接触", { transparency: 48, size: 35, distance: 0 }],
                    ["反射 (中): 接触", { transparency: 50, size: 55, distance: 0 }],
                    ["反射 (強): 接触", { transparency: 50, size: 90, distance: 0 }],
                    ["反射 (弱): 4 pt オフセット", { transparency: 48, size: 35, distance: 4 }],
                    ["反射 (中): 4 pt オフセット", { transparency: 50, size: 55, distance: 4 }],
                    ["反射 (強): 4 pt オフセット", { transparency: 50, size: 90, distance: 4 }],
                    ["反射 (弱): 8 pt オフセット", { transparency: 48, size: 35, distance: 8 }],
                    ["反射 (中): 8 pt オフセット", { transparency: 50, size: 55, distance: 8 }],
                    ["反射 (強): 8 pt オフセット", { transparency: 50, size: 90, distance: 8 }],
                ],
                [
                    ["transparency", "透明度", pct],
                    ["size", "サイズ", pct],
                    ["distance", "距離", pts(200)],
                ],
            );
            effectSection(
                "glow",
                "光彩",
                [5, 8, 11, 18].map((size) => [`光彩: ${size} pt`, { transparency: 60, size }]),
                [
                    ["color", "色"],
                    ["size", "サイズ", pts(150)],
                    ["transparency", "透明度", pct],
                ],
            );
        }

        // サイズ (cm), 回転 (°)
        const size = section("size", "サイズ");
        const b = boxOfObject(selected);
        for (const [field, label, px] of [["height", "高さ", b.h], ["width", "幅", b.w]]) {
            const row = line(size, label);
            const spin = sizeBox(row, `図形の${label}`, (cm) => {
                const now = boxOfObject(obj.box || selected);
                const v = (cm * 96) / 2.54;
                moveObject(obj, field === "height" ? { ...now, h: v } : { ...now, w: v });
            });
            spin.input.dataset.field = field;
            spin.setCm(cmOf(px));
        }
        if (canTurn(obj)) {
            const turn = sizeBox(line(size, "回転"), "図形の回転", (deg) => edit({ rotation: ((Math.round(deg) % 360) + 360) % 360 }), { name: "°", step: 1, min: -3600, max: 3600 });
            turn.input.dataset.field = "rotation";
            turn.setCm(Math.round(obj.rot || 0));
            const flips = size.appendChild(el("div", "pbtns"));
            for (const [label, vertical] of [["左右反転", false], ["上下反転", true]]) {
                const fb = flips.appendChild(el("button", null, label));
                fb.type = "button";
                fb.addEventListener("click", () => objectEdit(obj, flipChange(obj, vertical)));
            }
        }
        // (An ActiveX control's AutoSize: its size by what it shows.)
        if (ax && f.enabled !== undefined && !["ListBox", "SpinButton", "ScrollBar"].includes(t)) kit.checkRow(size, "自動サイズ（AutoSize）", f.autoSize, "autoSize", (on) => edit({ autoSize: on }));

        // プロパティ (オブジェクトの位置関係, オブジェクトを印刷する) and 保護 (ロック, 文字列のロック),
        // named as the cells' 保護 is.
        const props = obj.props;
        if (props) {
            const sec = section("size", "プロパティ");
            choices(
                sec,
                [
                    ["twoCell", "セルに合わせて移動やサイズ変更をする"],
                    ["oneCell", "セルに合わせて移動するがサイズ変更はしない"],
                    ["absolute", "セルに合わせて移動やサイズ変更をしない"],
                ],
                props.placement,
                (v) => edit({ placement: v }),
            );
            kit.checkRow(sec, "オブジェクトを印刷する", props.print, "print", (on) => edit({ print: on }));
            const lock = section("size", "保護");
            kit.checkRow(lock, "ロック", props.locked, "locked", (on) => edit({ locked: on }));
            if (props.lockText !== null && props.lockText !== undefined) kit.checkRow(lock, "文字列のロック", props.lockText, "lockText", (on) => edit({ lockText: on }));
        }
        // 代替テキスト (a control's)
        if (isControl(obj) && f.altText !== undefined) {
            const sec = section("size", "代替テキスト");
            const alt = sec.appendChild(el("textarea", "ptext"));
            alt.rows = 3;
            alt.value = f.altText || "";
            alt.dataset.field = "altText";
            alt.title = "Ctrl+Enter で確定";
            alt.addEventListener("change", () => alt.value !== (f.altText || "") && edit({ altText: alt.value }));
            alt.addEventListener("keydown", (e) => {
                if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                    e.preventDefault();
                    alt.blur();
                    focusGrid();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    alt.value = f.altText || "";
                    focusGrid();
                }
            });
        }

        // コントロール: what its kind has (the linked cell, the items, the state, the values)
        if (isControl(obj) && obj.form) {
            const sec = section("size", "コントロール");
            if (["CheckBox", "Radio", "OptionButton", "ToggleButton"].includes(t)) {
                const row = sec.appendChild(el("label", "prow pcheck"));
                const on = row.appendChild(el("input"));
                on.type = "checkbox";
                on.checked = !!f.checked;
                on.dataset.field = "checked";
                row.appendChild(el("span", null, "オン（チェックする）"));
                on.addEventListener("change", () => edit({ checked: on.checked }));
            }
            const ref = (label, key) => {
                const input = box(sec, label, key, f[key] || "", (v) => {
                    v = v.trim().replace(/^=/, "");
                    if (v !== (f[key] || "")) edit({ [key]: v || null });
                });
                input.placeholder = key === "items" ? "$A$1:$A$5" : "$B$2";
            };
            if (["List", "Drop", "ListBox", "ComboBox"].includes(t)) ref(ax ? "ListFillRange" : "入力範囲", "items");
            if (["CheckBox", "Radio", "List", "Drop", "Spin", "Scroll", "OptionButton", "ToggleButton", "TextBox", "ListBox", "ComboBox", "SpinButton", "ScrollBar"].includes(t)) ref(ax ? "LinkedCell" : "リンクするセル", "link");
            const number = (label, key, op, limits) => {
                const input = box(sec, label, op, String(f[key] || 0), (v) => {
                    const n = Math.round(+v);
                    if (v.trim() !== "" && isFinite(n) && n !== (f[key] || 0)) edit({ [op]: n });
                }, "number");
                Object.assign(input, limits ? { min: "0", max: "30000", step: "1" } : { step: "1" });
            };
            // The item selected (a form control's list or combo box: its number; 0 none).
            if (t === "List" || t === "Drop") number("選択している項目", "sel", "selected", true);
            if (["Spin", "Scroll", "SpinButton", "ScrollBar"].includes(t)) {
                const numbers = ax
                    ? [["Value", "val", "value"], ["Min", "min", "min"], ["Max", "max", "max"], ["SmallChange", "inc", "step"]]
                    : [["現在値", "val", "value"], ["最小値", "min", "min"], ["最大値", "max", "max"], ["変化の増分", "inc", "step"]];
                if (t === "Scroll" || t === "ScrollBar") numbers.push([ax ? "LargeChange" : "ページ単位の変化", "page", "page"]);
                for (const [label, key, op] of numbers) number(label, key, op, !ax);
            }
            // An ActiveX control's Enabled (WordWrap is 文字の制御's, AutoSize サイズ's).
            if (ax && f.enabled !== undefined) kit.checkRow(sec, "有効（Enabled）", f.enabled, "enabled", (on) => edit({ enabled: on }));
            if (!sec.querySelector(".prow")) sec.parentNode.remove();
        }

        // コード: its name and how the code refers to it, its events, its properties (オブジェクト
        // モデル: as the object model has them, as a browser's developer tools list an element's)
        {
            const sec = section("code", "オブジェクト");
            const named = (label, value, title) => {
                const row = sec.appendChild(el("div", "prow"));
                row.appendChild(el("span", "plabel", label));
                const v = row.appendChild(el("span", "pcode", value));
                v.title = title;
            };
            named("名前", obj.name || (obj.shapeId !== undefined ? `_x0000_s${obj.shapeId}` : ""), "オブジェクトの名前（Name）");
            const code = view.sheet.codeName || "Sheet1";
            named("コード", obj.activex !== undefined && obj.name ? `${code}.${obj.name}` : `${code}.shapes(${JSON.stringify(obj.name || "")})`, "コードからこのオブジェクトを参照する書き方");
        }
        const items = eventItems(obj);
        if (items.length) {
            const sec = section("code", "イベント");
            if (!ax) sec.appendChild(el("div", "pnote", obj.macro ? `${eventName(obj)}イベント: ${obj.macro}` : `${eventName(obj)}イベントは登録されていません`));
            const btns = sec.appendChild(el("div", "pbtns"));
            for (const [label, act] of items) {
                const mb = btns.appendChild(el("button", null, label));
                mb.type = "button";
                mb.addEventListener("click", act);
            }
        }
        propertyTree(section("code", "オブジェクト モデル"), obj);

        // 文字: the text of a control (its caption or value), then what the cells' 文字 has, in
        // its order and by its names: フォント, 文字飾り, 文字の配置, 文字の制御, and 余白. (A
        // shape's text is typed in the shape.)
        const tb = obj.textBox;
        if (obj.text !== undefined && !tb) {
            const sec = section("text", ax ? (["TextBox", "ListBox", "ComboBox"].includes(t) ? "文字（Value）" : "文字（Caption）") : "文字");
            const text = sec.appendChild(el("textarea", "ptext"));
            text.rows = 4;
            text.value = obj.text;
            text.dataset.field = "text";
            text.title = "Ctrl+Enter で確定";
            const commit = () => {
                if (text.value !== obj.text) edit({ text: text.value });
            };
            text.addEventListener("change", commit);
            text.addEventListener("keydown", (e) => {
                if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                    e.preventDefault();
                    text.blur();
                    focusGrid();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    text.value = obj.text;
                    focusGrid();
                }
            });
        }

        // 文字 > フォント, 文字飾り: a control's font (an ActiveX control's Font, a form button's)
        // and its color (ForeColor)
        if (isControl(obj) && f.font) {
            const fn = f.font;
            const font = (v) => edit({ font: v });
            const sec = section("text", "フォント");
            const row = line(sec, "フォント名");
            const combo = row.appendChild(el("span", "pcombo"));
            const input = combo.appendChild(el("input"));
            input.type = "text";
            input.value = fn.name;
            input.dataset.field = "fontName";
            input.spellcheck = false;
            input.autocomplete = "off";
            const take = (v) => {
                v = v.trim();
                if (v && v !== fn.name) font({ name: v });
            };
            input.addEventListener("change", () => take(input.value));
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    e.stopPropagation();
                    input.blur();
                    focusGrid();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    input.value = fn.name;
                    focusGrid();
                }
            });
            const drop = combo.appendChild(el("button", "pdrop"));
            drop.appendChild(el("span", "chev"));
            drop.type = "button";
            drop.title = "フォントの一覧";
            drop.addEventListener("click", () => openFontList(combo, take, fn.name));
            const style = fn.bold && fn.italic ? "boldItalic" : fn.bold ? "bold" : fn.italic ? "italic" : "normal";
            kit.selectRow(sec, "スタイル", [["normal", "標準"], ["italic", "斜体"], ["bold", "太字"], ["boldItalic", "太字 斜体"]], style, "fontStyle", (v) => font({ bold: v.startsWith("bold"), italic: /italic/i.test(v) }));
            const size = sizeBox(line(sec, "サイズ"), "フォント サイズ", (n) => n > 0 && n !== fn.size && font({ size: n }), { name: "pt", step: 1, min: 1, max: 409 });
            size.input.dataset.field = "fontSize";
            size.setCm(fn.size);
            // (A control's underline is one line or none.)
            kit.selectRow(sec, "下線", UNDERLINES.slice(0, 2), fn.underline ? "single" : "none", "underline", (v) => font({ underline: v === "single" }));
            kit.colorRow(sec, "色", f.foreColor || null, ax ? "既定の色" : "自動", (c) => edit({ foreColor: c }), "foreColor");
            kit.checkRow(section("text", "文字飾り"), "取り消し線", fn.strike, "strike", (on) => font({ strike: on }));
        }

        // 文字 > 文字の配置: a shape's text box (縦位置), a form button's (横位置, 縦位置)
        const button = isControl(obj) && !ax && f.textAlign !== undefined;
        if (tb) {
            kit.selectRow(section("text", "文字の配置"), "縦位置", [["t", "上揃え"], ["ctr", "上下中央揃え"], ["b", "下揃え"]], tb.anchor, "textAnchor", (v) => edit({ textAnchor: v }), "pwide");
        } else if (button) {
            const sec = section("text", "文字の配置");
            const ends = [["justify", "両端揃え"], ["distributed", "均等割り付け"]];
            kit.selectRow(sec, "横位置", [["left", "左揃え"], ["center", "中央揃え"], ["right", "右揃え"], ...ends], f.textAlign, "textAlign", (v) => edit({ textAlign: v }), "pwide");
            kit.selectRow(sec, "縦位置", [["top", "上揃え"], ["center", "上下中央揃え"], ["bottom", "下揃え"], ...ends], f.textVAlign, "textVAlign", (v) => edit({ textVAlign: v }), "pwide");
        }
        // 文字 > 文字の制御: 図形内でテキストを折り返す (a shape's text box), an ActiveX control's
        // WordWrap
        if (tb) kit.checkRow(section("text", "文字の制御"), "図形内でテキストを折り返す", tb.wrap, "textWrap", (on) => edit({ textWrap: on }));
        else if (ax && f.enabled !== undefined && !["ListBox", "ComboBox", "SpinButton", "ScrollBar"].includes(t)) kit.checkRow(section("text", "文字の制御"), "折り返し（WordWrap）", f.wordWrap, "wordWrap", (on) => edit({ wordWrap: on }));
        // 文字 > 余白: 左・右・上・下 (cm) of a shape's text box, or of a form button (自動, or its own)
        if (tb || button) {
            const sec = section("text", "余白");
            const inset = tb ? tb.inset : f.inset;
            if (button) kit.checkRow(sec, "自動", f.insetAuto, "insetAuto", (on) => edit(on ? { insetAuto: true } : { inset: { left: inset[0], top: inset[1], right: inset[2], bottom: inset[3] } }));
            if (tb || !f.insetAuto) {
                const cm = (pt) => Math.round(((pt * 2.54) / 72) * 100) / 100;
                [["left", "左", 0], ["right", "右", 2], ["top", "上", 1], ["bottom", "下", 3]].forEach(([key, label, i]) => {
                    const m = sizeBox(line(sec, label), `${label}余白`, (v) => edit({ inset: { [key]: Math.round(((v * 72) / 2.54) * 100) / 100 } }), { step: 0.05, max: 50 });
                    m.input.dataset.field = `inset-${key}`;
                    m.setCm(cm(inset[i]));
                });
            }
        }

        kit.finish(PANE_PAGES, PANE_ICONS, "panePage", focus);
    }

    // ---- セルの書式設定 (Ctrl+1): the formats of the cells selected in the pane at the right, as
    // Excel's dialog has them (表示形式, 配置, フォント, 罫線, 塗りつぶし, 保護) but with the
    // controls of 図形の書式設定 and its pages: 表示形式, 文字 (フォント and 配置, as an object's
    // 文字), 塗りつぶしと罫線 (as an object's 塗りつぶしと線), プロパティ; each is set as it is
    // changed. Shown for the active cell (the borders: for the range), as Excel's dialog opens.
    // プロパティ has 保護 and what the cells selected are in: the rules of 条件付き書式 and the
    // names defined.

    const CELL_PAGES = [
        ["number", "表示形式"],
        ["font", "文字"],
        ["fill", "塗りつぶしと罫線"],
        ["props", "プロパティ"],
    ];
    const CELL_ICONS = {
        number: '<path d="M6.2 2.5 4.6 13.5M11.4 2.5 9.8 13.5M2.5 5.8h11.5M2 10.2h11.5"/>',
        font: PANE_ICONS.text,
        fill: PANE_ICONS.fill,
        props: '<rect x="2.5" y="1.5" width="11" height="13" rx="1"/><path d="M5 5h1M5 8h1M5 11h1M8 5h3M8 8h3M8 11h3"/>',
    };
    const H_ALIGNS = [
        ["general", "標準"],
        ["left", "左揃え"],
        ["center", "中央揃え"],
        ["right", "右揃え"],
        ["fill", "繰り返し"],
        ["justify", "両端揃え"],
        ["centerContinuous", "選択範囲内で中央"],
        ["distributed", "均等割り付け"],
    ];
    const V_ALIGNS = [
        ["top", "上揃え"],
        ["center", "上下中央揃え"],
        ["bottom", "下揃え"],
        ["justify", "両端揃え"],
        ["distributed", "均等割り付け"],
    ];
    const UNDERLINES = [
        ["none", "なし"],
        ["single", "一重線"],
        ["double", "二重線"],
        ["singleAccounting", "一重線 (会計)"],
        ["doubleAccounting", "二重線 (会計)"],
    ];
    // The patterns of 塗りつぶし > パターンの種類 (as Excel lists them) and their tiles (as the
    // grid draws them: a row of pixels per string, '1' the pattern's color).
    const FILL_PATTERNS = [
        ["solid", "単色", null],
        ["darkGray", "75% 灰色", ["0111", "1101"]],
        ["mediumGray", "50% 灰色", ["10", "01"]],
        ["lightGray", "25% 灰色", ["1000", "0010"]],
        ["gray125", "12.5% 灰色", ["1000", "0000", "0010", "0000"]],
        ["gray0625", "6.25% 灰色", ["10000000", "00000000", "00001000", "00000000"]],
        ["darkHorizontal", "横 縞", ["1", "1", "0", "0"]],
        ["darkVertical", "縦 縞", ["1100"]],
        ["darkDown", "右下がり斜線 縞", ["1100", "0110", "0011", "1001"]],
        ["darkUp", "左下がり斜線 縞", ["0011", "0110", "1100", "1001"]],
        ["darkGrid", "斜線 格子", ["1100", "1100", "0011", "0011"]],
        ["darkTrellis", "極太 斜線 格子", ["1100", "1111", "0011", "1111"]],
        ["lightHorizontal", "実線 横 縞", ["1", "0", "0", "0"]],
        ["lightVertical", "実線 縦 縞", ["1000"]],
        ["lightDown", "実線 右下がり斜線 縞", ["1000", "0100", "0010", "0001"]],
        ["lightUp", "実線 左下がり斜線 縞", ["0001", "0010", "0100", "1000"]],
        ["lightGrid", "実線 横 格子", ["1111", "1000", "1000", "1000"]],
        ["lightTrellis", "実線 斜線 格子", ["1001", "0110", "0110", "1001"]],
    ];
    // A fill as CSS: its color, or its pattern in `fg` (black without one) over `bg` (white).
    function fillCss(pattern, fg, bg) {
        if (pattern === "none") return "";
        if (pattern === "solid") return fg ? `background:${fg};` : "";
        const tile = (FILL_PATTERNS.find((p) => p[0] === pattern) || [])[2];
        if (!tile) return "";
        let d = "";
        tile.forEach((row, y) => [...row].forEach((b, x) => b === "1" && (d += `M${x} ${y}h1v1h-1z`)));
        const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='${tile[0].length}' height='${tile.length}' shape-rendering='crispEdges'><path fill='${fg || "#000000"}' d='${d}'/></svg>`;
        return `background:${bg || "#FFFFFF"} url("data:image/svg+xml,${encodeURIComponent(svg)}");`;
    }
    // The formats of a cell (the sheet's `formats`): its own, else its row's or its column's.
    function formatAt(r, c) {
        const s = view.sheet;
        const list = s.formats || [{}];
        const cell = cellOf(r, c);
        const i = cell ? cell.fm || 0 : s.rowFormats && s.rowFormats[r] !== undefined ? s.rowFormats[r] : s.colFormats && s.colFormats[c] !== undefined ? s.colFormats[c] : 0;
        return list[i] || {};
    }
    // The number format of a cell, [as in the file, as Japanese Excel shows it (when not the
    // same)]: a value's, else (an empty cell) that of its formats, its row's or its column's.
    function numberCodeAt(r, c) {
        const cell = cellOf(r, c);
        const f = cell && cell.t !== undefined ? cell : formatAt(r, c);
        return [f.n, f.sn];
    }
    // The borders of the cells selected, as Excel's 罫線 shows them: the sides of the range (or
    // those its neighbors draw there), the lines between its cells, and the diagonals of its
    // first cell. Each { s: the line's name, c: its color (null: automatic) } or null.
    function bordersOf(s) {
        const at = (r, c) => (r >= 0 && c >= 0 ? formatAt(r, c).borders || {} : {});
        const one = single(s);
        const tl = at(s.r1, s.c1);
        const br = at(s.r2, s.c2);
        return {
            left: tl.left || at(s.r1, s.c1 - 1).right || null,
            top: tl.top || at(s.r1 - 1, s.c1).bottom || null,
            right: br.right || at(s.r2, s.c2 + 1).left || null,
            bottom: br.bottom || at(s.r2 + 1, s.c2).top || null,
            insideV: !one && s.c2 > s.c1 ? tl.right || at(s.r1, s.c1 + 1).left || null : null,
            insideH: !one && s.r2 > s.r1 ? tl.bottom || at(s.r1 + 1, s.c1).top || null : null,
            diagDown: tl.diagDown || null,
            diagUp: tl.diagUp || null,
        };
    }
    // The 線のスタイル and 線の色 the borders are drawn with (the ribbon's menu and the pane's 罫線
    // share them; kept with the page's state, as the page is drawn again at each edit).
    function setPen(change) {
        Object.assign(borderPen, change);
        remember({ borderPen: { ...borderPen } });
        const line = tools.borderLine;
        if (line) {
            const label = (BORDER_LINES.find((l) => l[0] === borderPen.line) || BORDER_LINES[0])[1];
            line.querySelector(".sample").innerHTML = lineSample(borderPen.line, 28);
            line.title = `線のスタイル（${label}）`;
        }
        const color = tools.borderColor;
        if (color) {
            color.querySelector(".bar").style.background = borderPen.color || "#000000";
            color.title = `線の色（${borderPen.color || "自動"}）`;
        }
    }
    // 表示形式 of "セルの書式設定" in the section `sec` of a pane's kit: 分類, サンプル, and what
    // the kind has (as the dialog's page). `o.now`: the format (as the dialog shows it; null:
    // none, with `o.unset`), `o.cell` what サンプル shows, `o.at` where it is (the kind chosen
    // last there is kept); `o.put(code)` sets a format (as in the file), `o.redraw(field)` draws
    // it all again. Gives the list of 種類 ({ list, where }) when it has one.
    function numberSection(kit, sec, o) {
        const { block, spin, selectRow, checkRow } = kit;
        const { cell, at } = o;
        let numberList = null;
        if (o.now === null) {
            const pick = selectRow(sec, "分類", [["", "（設定なし）"]].concat(FORMAT_KINDS.map((k) => [k.key, k.label])), "", "numberKind", (key) => {
                const next = FORMAT_KINDS.find((k) => k.key === key);
                if (!next) return;
                const code = next.code ? next.code({ places: 2, comma: false, symbol: "¥", negative: "minus" }) : next.custom ? "G/標準" : next.list[0];
                remember({ numKind: { at, code, kind: key } });
                o.put(fileCode(code));
            });
            pick.options[0].hidden = true;
            return null;
        }
        {
            const now = o.now;
            // The kind chosen last for this cell, while its format is what it made (ユーザー定義
            // stays so, as the format is one of another kind too).
            const was = saved.numKind;
            const start = kindOf(now);
            let kind = was && was.at === at && was.code === now ? FORMAT_KINDS.find((k) => k.key === was.kind) || start.kind : start.kind;
            const state = { ...start.state };
            const apply = (code, kindKey) => {
                code = code.trim();
                remember({ numKind: { at, code, kind: kindKey || kind.key } });
                if (!code) return;
                if (fileCode(code) !== fileCode(now)) o.put(fileCode(code));
                else o.redraw("numberKind");
            };
            selectRow(sec, "分類", FORMAT_KINDS.map((k) => [k.key, k.label]), kind.key, "numberKind", (key) => {
                const next = FORMAT_KINDS.find((k) => k.key === key);
                // ユーザー定義 starts from the format so far, the others from their first.
                const code = next.code ? next.code(next.key === start.kind.key ? state : { places: 2, comma: false, symbol: "¥", negative: "minus" }) : next.custom ? now : next.list.includes(now) ? now : next.list[0];
                apply(code, key);
            });
            const sample = block(sec, "サンプル").appendChild(el("div", "psample"));
            const t = sampleText(cell, now);
            sample.textContent = t === null ? "（ここでは表示できません）" : t === "" ? " " : t;
            if (kind.places) {
                spin(sec, "小数点以下の桁数", "places", state.places, (n) => apply(kind.code({ ...state, places: Math.round(n) })), { name: "", step: 1, max: 30 });
            }
            if (kind.symbol) selectRow(sec, "記号", [["¥", "¥"], ["$", "$"], ["", "なし"]], state.symbol, "symbol", (v) => apply(kind.code({ ...state, symbol: v })));
            if (kind.comma) checkRow(sec, "桁区切り (,) を使用する", state.comma, "comma", (on) => apply(kind.code({ ...state, comma: on })));
            if (kind.negative) {
                const list = block(sec, "負の数の表示形式").appendChild(el("select", "plistbox"));
                list.size = NEGATIVES.length;
                list.dataset.field = "negative";
                for (const [v, l] of NEGATIVES) {
                    const opt = list.appendChild(el("option", null, kind.key === "currency" ? l.replace(/1,234/, `${state.symbol}1,234`) : l));
                    opt.value = v;
                    if (v.includes("red")) opt.className = "red";
                }
                list.value = state.negative;
                list.addEventListener("change", () => apply(kind.code({ ...state, negative: list.value })));
            }
            if (kind.list) {
                const b = block(sec, kind.custom ? null : "種類");
                let typed = null;
                if (kind.custom) {
                    typed = kit.box(b, "種類", "numberCode", now, (v) => apply(v));
                    typed.classList.add("pcode");
                    typed.placeholder = "G/標準";
                }
                const codes = kind.custom && !kind.list.includes(now) ? [now].concat(kind.list) : kind.list;
                const list = b.appendChild(el("select", "plistbox"));
                list.size = 8;
                list.dataset.field = "numberList";
                // Each by its name, or as Excel's sample (2012/3/14 13:30:55.2) shows in it (* before
                // those of the system), or as it is (ユーザー定義).
                const x = 40982 + 48655.2 / 86400;
                codes.forEach((c, i) => {
                    const file = fileCode(c);
                    const shown = kind.custom ? c : kind.names ? kind.names[i] : (SYSTEM_CODES.includes(c) ? "*" : "") + (formatNumber(/\[[hms]+\]/i.test(file) ? 1 + (x % 1) : x, file === "m/d/yyyy" ? "yyyy/m/d" : file === "m/d/yyyy h:mm" ? "yyyy/m/d h:mm" : file) ?? c);
                    const opt = list.appendChild(el("option", null, shown));
                    opt.value = c;
                    opt.title = c;
                });
                list.value = codes.includes(now) ? now : "";
                const where = `${at}:${kind.key}`;
                numberList = { list, where };
                list.addEventListener("scroll", () => remember({ numberListTop: { where, top: list.scrollTop } }), { passive: true });
                list.addEventListener("change", () => {
                    if (typed) typed.value = list.value;
                    remember({ numberListTop: { where, top: list.scrollTop } });
                    apply(list.value);
                });
            }
            if (kind.note) sec.appendChild(el("div", "pnote", kind.note));
        }
        return numberList;
    }
    function openCellPane(page) {
        if (!editable || !view || !sel) return;
        if (entry) commit(null);
        openPane("cells", page);
    }
    function cellPane(body, focus) {
        const kit = paneKit(body);
        const { section, line, colorRow, checkRow, selectRow, block, spin } = kit;
        const fm = formatAt(sel.ar, sel.ac);
        const range = allAreas()
            .map((a) => areaText(a))
            .join(",");
        body.appendChild(el("div", "pname", range));
        // (A list of the pane is chosen from by a click: the grid then has the focus.)
        const set = (change) => format(change);

        // The list of 種類 (drawn again as a format is chosen in it): where it was scrolled to.
        let numberList = null;
        // 表示形式: 分類, サンプル, and what the kind has (as the dialog's page).
        {
            const [n, sn] = numberCodeAt(sel.ar, sel.ac);
            const now = localCode(/^\[\$-F[48]00\]/i.test(n || "") ? n : sn || n || "General");
            numberList = numberSection(kit, section("number", "表示形式"), {
                now,
                cell: cellOf(sel.ar, sel.ac) || {},
                at: `${view.sheet.name}!${sel.ar}:${sel.ac}`,
                put: (code) => set({ numberFormat: code }),
                redraw: renderPane,
            });
        }

        // 文字: フォント (フォント名, スタイル, サイズ, 下線, 色, 標準フォント), 文字飾り, and
        // プレビュー of them over the fill (set below, with 塗りつぶし).
        const preview = el("div", "ppreview");
        {
            const book = bookFont(view.sheet);
            const name = fm.font || book.name;
            const size = fm.size || book.size;
            const sec = section("font", "フォント");
            const row = line(sec, "フォント名");
            const combo = row.appendChild(el("span", "pcombo"));
            const input = combo.appendChild(el("input"));
            input.type = "text";
            input.value = name;
            input.dataset.field = "fontName";
            input.spellcheck = false;
            input.autocomplete = "off";
            const take = (v) => {
                v = v.trim();
                if (v && v !== name) set({ fontName: v });
            };
            input.addEventListener("change", () => take(input.value));
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    e.stopPropagation();
                    input.blur();
                    focusGrid();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    input.value = name;
                    focusGrid();
                }
            });
            const drop = combo.appendChild(el("button", "pdrop"));
            drop.appendChild(el("span", "chev"));
            drop.type = "button";
            drop.title = "フォントの一覧";
            drop.addEventListener("click", () => openFontList(combo, take, name));
            const style = fm.b && fm.i ? "boldItalic" : fm.b ? "bold" : fm.i ? "italic" : "normal";
            selectRow(sec, "スタイル", [["normal", "標準"], ["italic", "斜体"], ["bold", "太字"], ["boldItalic", "太字 斜体"]], style, "fontStyle", (v) => set({ bold: v.startsWith("bold"), italic: /italic/i.test(v) }));
            spin(sec, "サイズ", "fontSize", size, (n) => set({ fontSize: n }), { name: "pt", step: 1, min: 1, max: 409 });
            selectRow(sec, "下線", UNDERLINES, fm.u || "none", "underline", (v) => set({ underline: v }));
            colorRow(sec, "色", fm.color || null, "自動", (c) => set({ color: c }), "fontColor");
            const standard = name === book.name && size === book.size && !fm.b && !fm.i && !fm.u && !fm.st && !fm.va && !fm.color;
            const std = checkRow(sec, "標準フォント", standard, "standardFont", (on) => {
                if (on) set({ fontName: book.name, fontSize: book.size, bold: false, italic: false, underline: "none", strike: false, superscript: false, color: null });
                else std.checked = true;
            });
            const deco = section("font", "文字飾り");
            checkRow(deco, "取り消し線", fm.st, "strike", (on) => set({ strike: on }));
            checkRow(deco, "上付き", fm.va === "superscript", "superscript", (on) => set({ superscript: on }));
            checkRow(deco, "下付き", fm.va === "subscript", "subscript", (on) => set({ subscript: on }));
            const text = preview.appendChild(el("span", null, name));
            Object.assign(text.style, {
                fontFamily: fontFamily(name),
                fontSize: `${Math.min(size, 26)}pt`,
                fontWeight: fm.b ? "bold" : "normal",
                fontStyle: fm.i ? "italic" : "normal",
                textDecoration: [fm.u ? "underline" : "", fm.st ? "line-through" : ""].join(" ").trim() || "none",
                textDecorationStyle: /double/i.test(fm.u || "") ? "double" : "solid",
                color: fm.color || "#000000",
                verticalAlign: fm.va === "superscript" ? "super" : fm.va === "subscript" ? "sub" : "baseline",
            });
            section("font", "プレビュー").appendChild(preview);
        }

        // 文字 > 配置: 文字の配置, 文字の制御, 方向, 右から左
        {
            const sec = section("font", "文字の配置");
            const h = fm.h || "general";
            selectRow(sec, "横位置", H_ALIGNS, h, "hAlign", (v) => set({ align: v }), "pwide");
            // (An indent makes 標準 and 中央揃え 左揃え, as in Excel.)
            const indented = ["left", "right", "distributed"].includes(h);
            spin(sec, "インデント", "indent", fm.indent || 0, (n) => set(indented ? { indent: Math.round(n) } : { align: "left", indent: Math.round(n) }), { name: "", step: 1, max: 250 });
            selectRow(sec, "縦位置", V_ALIGNS, fm.v || "bottom", "vAlign", (v) => set({ valign: v === "center" ? "middle" : v }), "pwide");
            const ctl = section("font", "文字の制御");
            checkRow(ctl, "折り返して全体を表示する", fm.wrap, "wrap", (on) => set({ wrap: on }));
            checkRow(ctl, "縮小して全体を表示する", fm.shrink, "shrink", (on) => set({ shrink: on })).disabled = !!fm.wrap;
            const merged = mergedIn(sel);
            const m = checkRow(ctl, "セルを結合する", merged, "merge", (on) => merge(on ? "merge" : "unmerge"));
            m.disabled = sel.kind === "all" || areas.length > 0 || (!merged && single(sel));
            const dir = section("font", "方向");
            const rot = fm.rot || 0;
            const vertical = rot === 255;
            checkRow(dir, "縦書き", vertical, "vertical", (on) => set({ orientation: on ? "vertical" : 0 }));
            const angle = spin(dir, "角度", "orientation", vertical ? 0 : rot <= 90 ? rot : 90 - rot, (d) => set({ orientation: Math.round(d) }), { name: "°", step: 1, min: -90, max: 90 });
            angle.input.disabled = vertical;
            const rtl = section("font", "右から左");
            selectRow(rtl, "文字の方向", [["context", "コンテキスト"], ["ltr", "左から右"], ["rtl", "右から左"]], ["context", "ltr", "rtl"][fm.ro || 0] || "context", "readingOrder", (v) => set({ readingOrder: v }), "pwide");
        }

        // 塗りつぶしと罫線 > 塗りつぶし: 背景色, パターンの色, パターンの種類 (shown in 文字's
        // プレビュー too)
        {
            const fill = fm.fill || null;
            const pattern = fill ? fill.pattern : "none";
            const plain = pattern === "none" || pattern === "solid";
            const bg = !fill ? null : pattern === "solid" ? fill.fg : fill.bg;
            const fg = plain ? (saved.patternColor ?? null) : fill.fg;
            const put = (p, color, background) => set({ pattern: { pattern: p, color, background } });
            const sec = section("fill", "塗りつぶし");
            colorRow(sec, "背景色", bg, "色なし", (c) => (plain ? put(c ? "solid" : "none", c, null) : put(pattern, fg, c)), "fillColor");
            colorRow(sec, "パターンの色", fg, "自動", (c) => {
                if (!plain) return put(pattern, c, bg);
                // (No pattern yet: the one chosen next is drawn in it.)
                remember({ patternColor: c });
                renderPane("patternColor");
            }, "patternColor");
            const row = line(sec, "パターンの種類");
            const kind = row.appendChild(el("button", "pcolor"));
            kind.type = "button";
            kind.dataset.field = "pattern";
            const name = plain ? "なし" : (FILL_PATTERNS.find((p) => p[0] === pattern) || [, pattern])[1];
            kind.title = `パターンの種類（${name}）`;
            const sw = kind.appendChild(el("span", "sw ptile"));
            if (plain) {
                sw.classList.add("none");
                kind.appendChild(el("span", "pctext", "なし"));
            } else sw.style.cssText = fillCss(pattern, fg, bg);
            kind.appendChild(el("span", "pdrop chev"));
            kind.addEventListener("click", () => {
                const pop = openPop(kind, "patterns");
                const tiles = pop.appendChild(el("div", "swatches"));
                for (const [key, label] of FILL_PATTERNS) {
                    const b = tiles.appendChild(el("button", "swatch"));
                    b.type = "button";
                    b.title = label;
                    b.dataset.pattern = key;
                    b.style.cssText = key === "solid" ? "background:#FFFFFF" : fillCss(key, "#000000", "#FFFFFF");
                    b.classList.toggle("on", key === pattern || (key === "solid" && plain));
                    b.addEventListener("click", () => {
                        closeMenu();
                        focusGrid();
                        if (key === "solid") put(bg ? "solid" : "none", bg, null);
                        else put(key, fg, bg);
                    });
                }
            });
            preview.style.cssText = fillCss(pattern, pattern === "solid" ? bg : fg, bg);
        }

        // 罫線: one section, as each effect of 図形の書式設定 is: the line and its color (what the
        // borders are drawn with), then where they are drawn: by a preset, by the buttons around
        // the picture of the cells, or by the picture itself clicked.
        {
            const sides = bordersOf(sel);
            const many = { v: !single(sel) && sel.c2 > sel.c1, h: !single(sel) && sel.r2 > sel.r1 };
            const box = section("fill", "罫線");
            const row = line(box, "スタイル");
            const pick = row.appendChild(el("button", "pcolor pline"));
            pick.type = "button";
            pick.dataset.field = "borderLine";
            pick.title = `線のスタイル（${(BORDER_LINES.find((l) => l[0] === borderPen.line) || BORDER_LINES[0])[1]}）`;
            pick.appendChild(el("span", "sample")).innerHTML = lineSample(borderPen.line, 56);
            pick.appendChild(el("span", "pdrop chev"));
            pick.addEventListener("click", () => {
                const pop = openPop(pick, "linestyles");
                for (const [key, label] of BORDER_LINES) {
                    const b = popItem(pop, "", () => {
                        setPen({ line: key });
                        renderPane("borderLine");
                    });
                    b.title = label;
                    b.classList.toggle("on", borderPen.line === key);
                    b.querySelector("span").innerHTML = lineSample(key, 96);
                }
            });
            colorRow(box, "色", borderPen.color, "自動", (c) => {
                setPen({ color: c });
                renderPane("borderColor");
            }, "borderColor");
            const pen = () => ({ line: borderPen.line, color: borderPen.color });
            const presets = box.appendChild(el("div", "prow pborderpresets"));
            presets.appendChild(el("span", "plabel", "プリセット"));
            const preset = (label, icon, change, off) => {
                const b = presets.appendChild(el("button", "pside"));
                b.type = "button";
                b.title = label;
                b.setAttribute("aria-label", label);
                b.appendChild(borderIcon(icon));
                b.disabled = !!off;
                b.addEventListener("click", () => set({ borders: change() }));
            };
            preset("なし", {}, () => ({ left: null, top: null, right: null, bottom: null, insideV: null, insideH: null, diagDown: null, diagUp: null }));
            preset("外枠", { t: "s", b: "s", l: "s", r: "s" }, () => ({ left: pen(), top: pen(), right: pen(), bottom: pen() }));
            preset("内側", { h: "s", v: "s" }, () => Object.assign(many.v ? { insideV: pen() } : {}, many.h ? { insideH: pen() } : {}), !many.v && !many.h);

            const grid = box.appendChild(el("div", "pborders"));
            // A side clicked: drawn with the pen, or taken away when it is drawn so already.
            const toggle = (k) => {
                const now = sides[k];
                const same = now && now.s === borderPen.line && (now.c || "").toUpperCase() === (borderPen.color || "").toUpperCase();
                set({ borders: { [k]: same ? null : pen() } });
            };
            const ICONS = {
                top: { t: "s" },
                insideH: { h: "s" },
                bottom: { b: "s" },
                diagUp: { du: "s" },
                left: { l: "s" },
                insideV: { v: "s" },
                right: { r: "s" },
                diagDown: { dd: "s" },
            };
            const NAMES = { top: "上罫線", insideH: "内側横罫線", bottom: "下罫線", diagUp: "斜め罫線（／）", left: "左罫線", insideV: "内側縦罫線", right: "右罫線", diagDown: "斜め罫線（＼）" };
            const sideButton = (into, k) => {
                const b = into.appendChild(el("button", "pside"));
                b.type = "button";
                b.title = NAMES[k];
                b.setAttribute("aria-label", NAMES[k]);
                b.setAttribute("aria-pressed", String(!!sides[k]));
                b.dataset.side = k;
                b.classList.toggle("on", !!sides[k]);
                b.appendChild(borderIcon(ICONS[k]));
                b.disabled = (k === "insideV" && !many.v) || (k === "insideH" && !many.h);
                b.addEventListener("click", () => toggle(k));
            };
            const left = grid.appendChild(el("div", "pbcol"));
            for (const k of ["top", "insideH", "bottom"]) sideButton(left, k);
            // The picture: a square of the cells (two when the range has more) and their sides.
            const [W, H, M] = [128, 128, 12];
            const [x1, y1, x2, y2] = [M, M, W - M, H - M];
            const [xm, ym] = [(x1 + x2) / 2, (y1 + y2) / 2];
            const stroke = (side, x, y, X, Y) => {
                const b = sides[side];
                if (!b) return "";
                const [w, dash, double] = (BORDER_LINES.find((l) => l[0] === b.s) || BORDER_LINES[0])[2];
                const c = b.c || "#000000";
                if (double) {
                    const [dx, dy] = y === Y ? [0, 1.5] : x === X ? [1.5, 0] : [1.1, -1.1];
                    return `<path d="M${x - dx} ${y - dy}L${X - dx} ${Y - dy}M${x + dx} ${y + dy}L${X + dx} ${Y + dy}" stroke="${c}" stroke-width="1"/>`;
                }
                return `<path d="M${x} ${y}L${X} ${Y}" stroke="${c}" stroke-width="${w}"${dash ? ` stroke-dasharray="${dash}"` : ""}/>`;
            };
            const xs = many.v ? [[x1, xm], [xm, x2]] : [[x1, x2]];
            const ys = many.h ? [[y1, ym], [ym, y2]] : [[y1, y2]];
            let svg = "";
            // The corners and the middles, as Excel marks them.
            for (const [x, y] of [[x1, y1], [x2, y1], [x1, y2], [x2, y2]].concat(many.v ? [[xm, y1], [xm, y2]] : [], many.h ? [[x1, ym], [x2, ym]] : [])) {
                svg += `<path d="M${x - 6} ${y}h-4M${x + 6} ${y}h4M${x} ${y - 6}v-4M${x} ${y + 6}v4" stroke="#9a9a9a" stroke-width="1"/>`;
            }
            for (const [a, b] of xs)
                for (const [c, d] of ys) {
                    svg += stroke("diagDown", a, c, b, d);
                    svg += stroke("diagUp", a, d, b, c);
                }
            svg += stroke("top", x1, y1, x2, y1) + stroke("bottom", x1, y2, x2, y2) + stroke("left", x1, y1, x1, y2) + stroke("right", x2, y1, x2, y2);
            if (many.v) svg += stroke("insideV", xm, y1, xm, y2);
            if (many.h) svg += stroke("insideH", x1, ym, x2, ym);
            const pic = grid.appendChild(el("div", "pbpic"));
            pic.title = "クリックすると、その罫線を引きます（引いてある罫線は消します）";
            pic.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" fill="none" aria-hidden="true">${svg}</svg>`;
            pic.addEventListener("click", (e) => {
                const r = pic.firstChild.getBoundingClientRect();
                const [x, y] = [((e.clientX - r.left) * W) / r.width, ((e.clientY - r.top) * H) / r.height];
                // The nearest line (within 8 px): a side, a line between the cells, a diagonal.
                const near = [
                    ["top", Math.abs(y - y1)],
                    ["bottom", Math.abs(y - y2)],
                    ["left", Math.abs(x - x1)],
                    ["right", Math.abs(x - x2)],
                ];
                if (many.v) near.push(["insideV", Math.abs(x - xm)]);
                if (many.h) near.push(["insideH", Math.abs(y - ym)]);
                const cw = (x2 - x1) / xs.length;
                const ch = (y2 - y1) / ys.length;
                const [u, v] = [(((x - x1) % cw) + cw) % cw, (((y - y1) % ch) + ch) % ch];
                const k = Math.hypot(cw, ch);
                near.push(["diagDown", Math.abs(u * ch - v * cw) / k], ["diagUp", Math.abs(u * ch + v * cw - cw * ch) / k]);
                near.sort((a, b) => a[1] - b[1]);
                if (near[0][1] <= 8) toggle(near[0][0]);
            });
            // ／ in the corner under the buttons of the rows, ＼ at the end of those of the columns.
            sideButton(grid.appendChild(el("div", "pbcorner")), "diagUp");
            const bottom = grid.appendChild(el("div", "pbrow"));
            for (const k of ["left", "insideV", "right", "diagDown"]) sideButton(bottom, k);
            box.appendChild(el("div", "pnote", "選んだスタイルと色で、プリセットや周りのボタン、図のクリックで罫線を引きます（同じ罫線をもう一度クリックすると消えます）。"));
        }

        // プロパティ > 保護: ロック, 表示しない (what they do on a protected sheet)
        {
            const sec = section("props", "保護");
            checkRow(sec, "ロック", !fm.unlocked, "locked", (on) => set({ locked: on }));
            checkRow(sec, "表示しない", fm.hidden, "hidden", (on) => set({ hidden: on }));
            sec.appendChild(el("div", "pnote", "ワークシートを保護しなければ、セルをロックしたり数式を表示しないようにした効果はありません（［校閲］タブの［シートの保護］で保護します）。"));
        }

        // プロパティ > 条件付き書式: the rules on the cells selected, in their order of priority
        // (one clicked: chosen in "条件付き書式ルールの管理").
        {
            const sec = section("props", "条件付き書式");
            const rules = view.sheet.condRules || [];
            const on = rules.map((r, i) => i).filter((i) => touchesSelection(condAreas(rules[i].range)));
            const list = sec.appendChild(el("div", "plist"));
            for (const i of on) {
                const r = rules[i];
                const what = condText(r);
                const where = String(r.range).split(/\s+/).filter(Boolean).join(",");
                const b = list.appendChild(el("button", "pitem pcond"));
                b.type = "button";
                b.dataset.rule = String(i);
                b.title = `${what}\n適用先: ${where}${r.stop ? "\n条件を満たす場合は停止" : ""}`;
                const top = b.appendChild(el("span", "pctop"));
                top.appendChild(condSample(r));
                top.appendChild(el("span", "pcwhat", what));
                b.appendChild(el("span", "pcwhere", `適用先: ${where}`));
                b.addEventListener("click", () => {
                    const state = condManageState();
                    state.chosen = i;
                    manageCond(state);
                });
            }
            if (!on.length) list.appendChild(el("div", "pnone", "選択範囲に条件付き書式はありません。"));
            const btns = sec.appendChild(el("div", "pbtns"));
            const manage = btns.appendChild(el("button", null, "ルールの管理…"));
            manage.type = "button";
            manage.dataset.field = "condManage";
            manage.addEventListener("click", () => manageCond());
        }

        // プロパティ > 名前の定義: the names that refer to the cells selected (one clicked: what it
        // refers to selected).
        {
            const sec = section("props", "名前の定義");
            const named = (data.names || []).map((n) => ({ n, areas: nameAreas(n) })).filter((x) => touchesSelection(x.areas));
            const list = sec.appendChild(el("div", "plist"));
            for (const { n, areas: at } of named) {
                const scope = n.sheet || "ブック";
                const b = list.appendChild(el("button", "pitem pdefname"));
                b.type = "button";
                b.dataset.name = n.name;
                b.title = `${n.name}（範囲: ${scope}）\n参照範囲: =${n.ref}`;
                const top = b.appendChild(el("span", "pctop"));
                top.appendChild(el("span", "pnname", n.name));
                top.appendChild(el("span", "pnscope", scope));
                b.appendChild(el("span", "pcwhere", `=${n.ref}`));
                b.addEventListener("click", () => {
                    selectAreas(at);
                    focusGrid();
                });
            }
            if (!named.length) list.appendChild(el("div", "pnone", "選択範囲を参照する名前はありません。"));
        }

        // (The pages that were 配置, 罫線 and 保護 are those that have them now.)
        const was = { align: "font", border: "fill", protect: "props" }[saved.cellPage];
        if (was) remember({ cellPage: was });
        kit.finish(CELL_PAGES, CELL_ICONS, "cellPage", focus);
        // (Once its page shows: a list hidden is not scrolled.)
        const top = saved.numberListTop;
        if (numberList && top && top.where === numberList.where) numberList.list.scrollTop = top.top;
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
        const tabs = form.appendChild(el("div", "symtabs tablist"));
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
            items.push(["セルの書式設定…", () => openCellPane(), "Ctrl+1"]);
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
            if (item[2]) b.appendChild(el("kbd", null, item[2]));
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
        // Run mode: `BeforeRightClick` first; the menu opens unless a handler cancels.
        if (runMode && view && sel) {
            const [x, y, kind] = [e.clientX, e.clientY, sel.kind];
            postRun({ op: "rightClick", sheet: view.sheet.name, range: allAreas().map(rangeText).join(",") }, (m) => {
                if (!m.cancel) openMenu(x, y, kind);
            });
            return;
        }
        openMenu(e.clientX, e.clientY, sel ? sel.kind : hit.kind);
    });
    document.addEventListener("mousedown", (e) => {
        if (menu && !menu.contains(e.target)) closeMenu();
    });
    window.addEventListener("blur", closeMenu);

    // ---- the events of the objects: their handlers in the sheet's file, written by the editor
    // An ActiveX control's events are `@on("CommandButton1.Click")` functions of the sheet's
    // file; a shape or a form control runs the macro assigned to it (`OnAction`): the one the
    // compiler generates for its `@on("ボタン 1.Click")` functions, or another one. The editor says
    // whether the workbook shown is the project's (`events`), whose sheets have files.

    let events = false;
    const hasEvents = (obj) => events && !!obj && (obj.activex !== undefined || !!obj.onAction);
    // The event a shape's or a form control's function (its OnAction) runs on: a click; a
    // change of the item selected (a list or combo box); of the value (a spin button, a scroll bar).
    function eventName(obj) {
        const t = obj && obj.form && obj.form.type;
        return t === "List" || t === "Drop" ? "選択" : t === "Spin" || t === "Scroll" ? "変更" : "クリック";
    }
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
            eventName: eventName(obj),
            choose,
        });
    }
    function eventItems(obj) {
        if (!hasEvents(obj)) return [];
        if (obj.activex !== undefined) {
            return [
                ["コードの表示", () => openEvent(obj, false), ""],
                ["イベントの登録…", () => openEvent(obj, true), ""],
            ];
        }
        const ev = `${eventName(obj)}イベント`;
        return obj.macro
            ? [
                  ["コードの表示", () => openEvent(obj, false), ""],
                  [`${ev}の登録…`, () => openEvent(obj, true), ""],
              ]
            : [
                  [`${ev}の登録`, () => openEvent(obj, false), ""],
                  [`${ev}に関数を登録…`, () => openEvent(obj, true), ""],
              ];
    }
    // The menu of an object: its text, deleting it, its events.
    function objectItems(obj) {
        const items = [];
        if (movable(obj)) {
            if (obj.text !== undefined) items.push(["テキストの編集", () => editObjectText(obj), "F2"]);
            items.push([isActiveX(obj) ? "プロパティ" : isControl(obj) ? "コントロールの書式設定" : "図形の書式設定", () => openPane("format"), ""]);
            items.push(["削除", () => deleteObject(obj), "Delete"]);
            if (!isControl(obj)) {
                items.push(null);
                for (const [how, label] of ORDERS) items.push([label, () => (isSelected(obj) ? editSelected((o) => !isControl(o) && { order: how }) : objectEdit(obj, { order: how })), ""]);
            }
            if (canTurn(obj)) items.push(["回転", () => openPane("format", "rotation"), ""]);
        }
        const more = eventItems(obj);
        if (items.length && more.length) items.push(null);
        return items.concat(more);
    }
    // The button of the bar: the events of the object selected (a menu).
    // (Only where the objects have events: the project's workbook.)
    function refreshEventButton() {
        const b = $("event");
        if (!b) return;
        const on = hasEvents(selectedObj);
        b.disabled = !on;
        if (tools.macroGroup) tools.macroGroup.hidden = !events;
        if (tools.modeSwitch) tools.modeSwitch.hidden = !events;
    }
    function enableEvents() {
        events = true;
        refreshEventButton();
        renderPane();
    }

    // ---- the run mode: what is done on the sheet raises the events of the program
    // The switch at the right end of the ribbon's tabs (with the project's workbook): 編集 edits
    // the workbook file; 実行 runs the program on the stand-in for Excel, where what the user does
    // raises its events (spec §14.10.10): a sheet chosen (`Deactivate`, `Activate`), cells
    // selected (`SelectionChange`), typed into or cleared (`Change`), double-clicked and
    // right-clicked (`BeforeDoubleClick`, `BeforeRightClick`), an ActiveX control clicked
    // (`GotFocus`, `MouseDown`, `MouseUp`, `Click`, then `LostFocus`), a shape or a form control
    // clicked (its macro). The editor draws the page again when the program changed the workbook.

    let runMode = !!(data.run && editable);
    // The object pressed (the mouse down on it), and the ActiveX control that has the focus.
    let runPress = null;
    // What waits for its answer: { id, after }.
    let runWait = null;
    let runIds = 0;
    let runTimer = 0;
    const RUN_ONLY = "実行モードでは、セルへの入力と削除だけができます。ほかの編集は、リボンの右端のスイッチで「編集」モードに切り替えてから行ってください。";
    function setRunMode(on, session) {
        runMode = on && editable;
        document.body.classList.toggle("run-mode", runMode);
        showMode();
        if (tools.modeSwitch) {
            for (const b of tools.modeSwitch.querySelectorAll("button")) {
                const pressed = (b.dataset.mode === "run") === runMode;
                b.classList.toggle("on", pressed);
                b.setAttribute("aria-pressed", String(pressed));
            }
        }
        if (!runMode) return;
        // A new run: what the program was told (the sheet, the selection, the focus) is told again.
        if (session !== undefined && saved.runSession !== session) remember({ runSession: session, runSheet: null, runSel: null, runFocus: null });
        // (No object stays selected: they are clicked now.)
        if (selectedObj) {
            clearSelection();
            if (sel) paint(false);
        }
        runSelectSoon();
    }
    function askMode(run) {
        if (run === runMode || busy || !events) return;
        if (entry) cancelEntry();
        closeMenu();
        setBusy(true);
        vscode.postMessage({ type: "mode", run });
    }
    // `after`: the page waits for the answer (busy), then calls it with it ({ cancel }).
    function postRun(m, after) {
        // What was selected is told first: `SelectionChange` comes before what is done there.
        // (Not before cells typed into: Enter moved the selection after them, and Excel raises
        // `Change` before that `SelectionChange`.)
        if (m.op !== "select" && m.op !== "activate" && m.op !== "enter") runSelectNow();
        const id = ++runIds;
        if (after) {
            runWait = { id, after };
            setBusy(true);
        }
        vscode.postMessage({ type: "run", id, ...m });
    }
    function runDone(m) {
        if (!runWait || runWait.id !== m.id) return;
        const after = runWait.after;
        runWait = null;
        setBusy(false);
        // (Not while a control is typed into.)
        if (!runAskBox) focusGrid();
        after(m);
    }
    // Cells typed into or cleared (one text for all the ranges); the other edits are refused.
    function runEdit(ops) {
        const cells = ops.filter((o) => o.op !== "validate");
        const text = (o) => (o.op === "clear" ? "" : o.text);
        if (!cells.length || !cells.every((o) => (o.op === "set" || o.op === "clear") && o.sheet === cells[0].sheet && text(o) === text(cells[0]))) {
            showError(RUN_ONLY);
            return false;
        }
        postRun({ op: "enter", sheet: cells[0].sheet, ranges: cells.map((o) => o.range), text: text(cells[0]) }, () => {});
        return true;
    }
    // The selection, told when it has settled (not while the mouse drags it).
    function runSelectSoon() {
        clearTimeout(runTimer);
        runTimer = setTimeout(runSelectNow, 150);
    }
    function runSelectNow() {
        clearTimeout(runTimer);
        if (!runMode || !view || !sel) return;
        if (drag) {
            runSelectSoon();
            return;
        }
        const sheet = view.sheet.name;
        const range = allAreas().map(rangeText).join(",");
        const active = address(sel.ar, sel.ac);
        const key = `${range}|${active}`;
        if (saved.runSheet !== sheet) {
            // Another sheet: `Activate` (its selection is the one it had: no `SelectionChange`).
            remember({ runSheet: sheet, runSel: key, runFocus: null });
            postRun({ op: "activate", sheet });
            return;
        }
        if (saved.runSel === key) return;
        remember({ runSel: key });
        runBlur();
        postRun({ op: "select", sheet, range, active });
    }
    // The ActiveX control that had the focus loses it.
    function runBlur(next) {
        const had = saved.runFocus;
        if (!had || had === next || !view) return;
        remember({ runFocus: next || null });
        postRun({ op: "control", sheet: view.sheet.name, object: had, events: [{ event: "LostFocus" }] });
    }
    // What a click on a control changes, by where it is on it: the part of a spin button or a
    // scroll bar (`more` / `less`, `pageMore` / `pageLess`), the item of a list box.
    function runAction(obj, e) {
        const type = (obj.form && obj.form.type) || "";
        const r = obj.box.getBoundingClientRect();
        const horizontal = r.width > r.height;
        // (Along the axis, from where the value is least: left, or top.)
        const len = horizontal ? r.width : r.height;
        const at = horizontal ? e.clientX - r.left : e.clientY - r.top;
        if (type === "SpinButton" || type === "Spin") {
            // Up (top) and right are more.
            return { part: horizontal ? (at > len / 2 ? "more" : "less") : at < len / 2 ? "more" : "less" };
        }
        if (type === "ScrollBar" || type === "Scroll") {
            const arrow = Math.min(r.width, r.height);
            if (at < arrow) return { part: "less" };
            if (at > len - arrow) return { part: "more" };
            const f = obj.form;
            const span = (f.max || 0) - (f.min || 0);
            const thumb = arrow + (len - 2 * arrow) * (span ? ((f.val || 0) - (f.min || 0)) / span : 0);
            return { part: at < thumb ? "pageLess" : "pageMore" };
        }
        if (type === "ListBox" || type === "List") {
            const span = e.target.closest && e.target.closest(".ctl.list > span");
            if (span) return { index: [...span.parentNode.children].indexOf(span) };
        }
        return null;
    }
    // A text typed into a control, or an item chosen from its list, in a box under it.
    function runAsk(obj, opts, done) {
        closeRunAsk();
        const r = obj.box.getBoundingClientRect();
        const pop = (runAskBox = document.body.appendChild(el("div", "runask")));
        pop.style.left = `${Math.round(r.left)}px`;
        pop.style.top = `${Math.round(r.bottom + 1)}px`;
        pop.style.minWidth = `${Math.round(r.width)}px`;
        let input = null;
        if (opts.text) {
            input = pop.appendChild(el("input"));
            input.type = "text";
            input.value = obj.text || "";
            input.addEventListener("keydown", (e) => {
                e.stopPropagation();
                if (e.key === "Enter") {
                    closeRunAsk();
                    done({ text: input.value });
                } else if (e.key === "Escape") {
                    closeRunAsk();
                    focusGrid();
                }
            });
        }
        (opts.items || []).forEach((item, index) => {
            const b = pop.appendChild(el("button", item === obj.text ? "on" : null, item || " "));
            b.type = "button";
            b.addEventListener("mousedown", (e) => e.preventDefault());
            b.addEventListener("click", () => {
                closeRunAsk();
                done({ index });
            });
        });
        if (input) {
            input.focus();
            input.select();
        }
    }
    let runAskBox = null;
    function closeRunAsk() {
        if (runAskBox) runAskBox.remove();
        runAskBox = null;
    }
    document.addEventListener("mousedown", (e) => {
        if (runAskBox && !runAskBox.contains(e.target)) closeRunAsk();
    });
    // An object clicked: an ActiveX control's events (and what the click changes), a shape's or
    // a form control's macro (a form control changed first: a check box turned on, ...).
    function runClick(obj, e) {
        if (busy || !view) return;
        const sheet = view.sheet.name;
        const type = (obj.form && obj.form.type) || "";
        if (obj.activex !== undefined && obj.name) {
            // Where on it, in points (Excel's X, Y), and the keys held (Shift 1, Ctrl 2, Alt 4).
            const r = obj.box.getBoundingClientRect();
            const scale = obj.box.offsetWidth ? obj.box.offsetWidth / (r.width || 1) : 1;
            const x = Math.round((e.clientX - r.left) * scale * 0.75 * 100) / 100;
            const y = Math.round((e.clientY - r.top) * scale * 0.75 * 100) / 100;
            const shift = (e.shiftKey ? 1 : 0) | (e.ctrlKey ? 2 : 0) | (e.altKey ? 4 : 0);
            const list = [];
            if (saved.runFocus !== obj.name) {
                runBlur(obj.name);
                remember({ runFocus: obj.name });
                list.push({ event: "GotFocus" });
            }
            list.push({ event: "MouseDown", args: [1, shift, x, y] }, { event: "MouseUp", args: [1, shift, x, y] }, { event: "Click" });
            const post = (more) => postRun({ op: "control", sheet, object: obj.name, events: list, ...more }, () => {});
            // A text box is typed into, a combo box's item chosen (or its text typed).
            if (type === "TextBox" || type === "ComboBox") {
                post({ events: list.filter((x) => x.event !== "Click") });
                runAsk(obj, { text: true, items: type === "ComboBox" ? obj.form.list || [] : [] }, (answer) =>
                    postRun({ op: "control", sheet, object: obj.name, events: answer.index !== undefined ? [{ event: "Click" }] : [], action: answer.index !== undefined ? { index: answer.index } : null, text: answer.text }, () => {}),
                );
                return;
            }
            post({ action: runAction(obj, e) });
            return;
        }
        const form = ["CheckBox", "Radio", "List", "Drop", "Spin", "Scroll"].includes(type) && !(obj.form && obj.form.activex);
        if (!form && !obj.onAction && !obj.macro) return;
        runBlur();
        const post = (action) => postRun({ op: "macro", sheet, object: obj.name || "", macro: obj.macro || "", action }, () => {});
        if (type === "Drop") {
            runAsk(obj, { items: obj.form.list || [] }, (answer) => post(answer));
            return;
        }
        post(form ? runAction(obj, e) : null);
    }
    document.addEventListener("mouseup", () => {
        runPress = null;
    });
    if (editable) vscode.postMessage({ type: "ready" });

    // ---- formats, merged cells, sorting and the filter: the toolbar, the menus, the keys

    const rangeOrCell = () => (single(sel) ? address(sel.ar, sel.ac) : rangeText(sel));
    // The cells selected, a merged cell with all its cells (its format is that of all of them).
    const areaText = (s) => (s.kind !== "cells" ? rangeText(s) : s.r1 === s.r2 && s.c1 === s.c2 ? address(s.r1, s.c1) : `${address(s.r1, s.c1)}:${address(s.r2, s.c2)}`);
    // The CSS the active cell is drawn with (its format, as far as the page knows it).
    const activeCss = () => (view && activeTd ? cssOf(view.sheet, activeTd.r, activeTd.c, activeTd.cell) : "");
    // The CSS of a cell: its own format, or (empty, without one) its row's or column's.
    function cssOf(sheet, r, c, cell) {
        if (cell) return (cell.s && sheet.styles[cell.s]) || "";
        const s = sheet.rowStyles && sheet.rowStyles[r] !== undefined ? sheet.rowStyles[r] : sheet.colStyles && sheet.colStyles[c];
        return (s !== undefined && sheet.styles[s]) || "";
    }
    const STATE = {
        bold: (css) => /font-weight:bold/.test(css),
        italic: (css) => /font-style:italic/.test(css),
        underline: (css) => /text-decoration:[^;]*underline/.test(css),
        strike: (css) => /text-decoration:[^;]*line-through/.test(css),
        wrap: (css) => /white-space:pre-wrap/.test(css),
    };
    function format(set) {
        if (!sel || !view) return;
        // While a cell is typed into: the characters'.
        if (formatChars(set)) return;
        const list = allAreas();
        const ops = list.map((a) => ({ op: "format", sheet: view.sheet.name, range: areaText(a), set }));
        // A new font fits the rows to it, as Excel does (those not given a height).
        const font =
            set.style === "Normal"
                ? () => ""
                : set.fontName !== undefined || set.fontSize !== undefined
                  ? (css) => {
                        let out = css || "";
                        const put = (prop, value) => (out = new RegExp(`${prop}:[^;]*;?`).test(out) ? out.replace(new RegExp(`${prop}:[^;]*;?`), `${prop}:${value};`) : `${out}${prop}:${value};`);
                        if (set.fontName !== undefined) put("font-family", fontFamily(set.fontName));
                        if (set.fontSize !== undefined) put("font-size", `${set.fontSize}pt`);
                        return out;
                    }
                  : null;
        postEdit(ops.concat(fitOps(list, font)));
    }
    // The edit that fits the rows of `list` to the font `font` gives their cells (none without one).
    function fitOps(list, font) {
        if (!font || !view) return [];
        const rows = refit(view.sheet, list, font);
        return rows.length ? [{ op: "fit", sheet: view.sheet.name, rows }] : [];
    }
    function toggleFormat(name) {
        if (editable && CHAR_FORMATS[name] && charTarget()) {
            const f = charFont();
            formatChars({ [name]: !{ bold: f.b, italic: f.i, underline: f.u, strike: f.s }[name] });
            return;
        }
        if (editable) format({ [name]: !STATE[name](activeCss()) });
    }
    function clearFormats() {
        if (sel && view) postEdit(allAreas().map((a) => ({ op: "clearFormats", sheet: view.sheet.name, range: areaText(a) })).concat(fitOps(allAreas(), () => "")));
    }
    // ---- ホーム > 条件付き書式: rules put on the cells selected, as Excel's menu and dialogs
    // put them (the first in priority), and cleared from them or from the sheet.

    // The formats Excel offers in its quick dialogs ("書式:"); "ユーザー設定の書式…" opens
    // "セルの書式設定".
    const COND_FORMATS = [
        ["red", "濃い赤の文字、明るい赤の背景", { color: "#9C0006", fill: "#FFC7CE" }],
        ["yellow", "濃い黄色の文字、黄色の背景", { color: "#9C5700", fill: "#FFEB9C" }],
        ["green", "濃い緑の文字、緑の背景", { color: "#006100", fill: "#C6EFCE" }],
        ["redFill", "明るい赤の背景", { fill: "#FFC7CE" }],
        ["redText", "赤の文字", { color: "#9C0006" }],
        ["redBorder", "赤の罫線", { border: "#FF0000" }],
        ["custom", "ユーザー設定の書式…", null],
    ];
    // The kinds of rules of "新しい書式ルール", named and ordered as Excel's.
    const COND_KINDS = [
        ["scale", "セルの値に基づいてすべてのセルを書式設定"],
        ["contains", "指定の値を含むセルだけを書式設定"],
        ["top10", "上位または下位に入る値だけを書式設定"],
        ["aboveAverage", "平均より上または下の値だけを書式設定"],
        ["duplicateValues", "一意の値または重複する値だけを書式設定"],
        ["expression", "数式を使用して、書式設定するセルを決定"],
    ];
    // "セルの値に基づいてすべてのセルを書式設定": its 書式スタイル.
    const COND_STYLES = [
        ["colorScale2", "2 色スケール"],
        ["colorScale3", "3 色スケール"],
        ["dataBar", "データ バー"],
        ["iconSet", "アイコン セット"],
    ];
    // "指定の値を含むセルだけを書式設定": what of the cells.
    const COND_CONTAINS = [
        ["cellIs", "セルの値"],
        ["text", "特定の文字列"],
        ["timePeriod", "日付"],
        ["containsBlanks", "空白"],
        ["notContainsBlanks", "空白なし"],
        ["containsErrors", "エラー"],
        ["notContainsErrors", "エラーなし"],
    ];
    const COND_OPERATORS = [
        ["between", "次の値の間"],
        ["notBetween", "次の値の間以外"],
        ["equal", "次の値に等しい"],
        ["notEqual", "次の値に等しくない"],
        ["greaterThan", "次の値より大きい"],
        ["lessThan", "次の値より小さい"],
        ["greaterThanOrEqual", "次の値以上"],
        ["lessThanOrEqual", "次の値以下"],
    ];
    const COND_TEXTS = [
        ["containsText", "次の値を含む"],
        ["notContainsText", "次の値を含まない"],
        ["beginsWith", "次の値で始まる"],
        ["endsWith", "次の値で終わる"],
    ];
    const COND_PERIODS = [
        ["yesterday", "昨日"],
        ["today", "今日"],
        ["tomorrow", "明日"],
        ["last7Days", "過去 7 日間"],
        ["lastWeek", "先週"],
        ["thisWeek", "今週"],
        ["nextWeek", "来週"],
        ["lastMonth", "先月"],
        ["thisMonth", "今月"],
        ["nextMonth", "来月"],
    ];
    // "平均より上または下の値だけを書式設定": above or below the average (by standard deviations).
    const COND_AVERAGES = [
        ["above", "より上"],
        ["below", "より下"],
        ["aboveEq", "以上"],
        ["belowEq", "以下"],
        ["above1", "標準偏差 1 上"],
        ["below1", "標準偏差 1 下"],
        ["above2", "標準偏差 2 上"],
        ["below2", "標準偏差 2 下"],
        ["above3", "標準偏差 3 上"],
        ["below3", "標準偏差 3 下"],
    ];
    // The data bars and color scales of the menu (Excel's first ones).
    const COND_BARS = [["#638EC6", "青"], ["#63C384", "緑"], ["#FF555A", "赤"], ["#FFB628", "オレンジ"], ["#008AEF", "水色"], ["#D6007B", "紫"]];
    const COND_SCALES = [
        ["#63BE7B", "#FFEB84", "#F8696B", "緑、黄、赤"],
        ["#F8696B", "#FFEB84", "#63BE7B", "赤、黄、緑"],
        ["#63BE7B", "#FCFCFF", "#F8696B", "緑、白、赤"],
        ["#F8696B", "#FCFCFF", "#63BE7B", "赤、白、緑"],
        ["#5A8AC6", "#FCFCFF", "#F8696B", "青、白、赤"],
        ["#F8696B", "#FCFCFF", "#5A8AC6", "赤、白、青"],
        ["#FCFCFF", null, "#F8696B", "白、赤"],
        ["#F8696B", null, "#FCFCFF", "赤、白"],
        ["#63BE7B", null, "#FCFCFF", "緑、白"],
        ["#FCFCFF", null, "#63BE7B", "白、緑"],
        ["#63BE7B", null, "#FFEF9C", "緑、黄"],
        ["#FFEF9C", null, "#63BE7B", "黄、緑"],
    ];
    // The icon sets in the order of Excel's dialog ("アイコン スタイル"), named as Excel names
    // them (3 つの星・3 つの三角形・5 つのボックス: of Excel 2010, written in the sheet's extLst).
    const COND_ICONS = [
        ["3Arrows", "3 つの矢印 (色分け)"],
        ["3ArrowsGray", "3 つの矢印 (灰色)"],
        ["3Flags", "3 つのフラグ"],
        ["3TrafficLights1", "3 つの信号 (枠なし)"],
        ["3TrafficLights2", "3 つの信号 (枠あり)"],
        ["3Signs", "3 つの図形"],
        ["3Symbols", "3 つの記号 (丸囲み)"],
        ["3Symbols2", "3 つの記号 (丸囲みなし)"],
        ["3Stars", "3 つの星"],
        ["3Triangles", "3 つの三角形"],
        ["4Arrows", "4 つの矢印 (色分け)"],
        ["4ArrowsGray", "4 つの矢印 (灰色)"],
        ["4RedToBlack", "赤から黒"],
        ["4Rating", "4 つの評価"],
        ["4TrafficLights", "4 つの信号"],
        ["5Arrows", "5 つの矢印 (色分け)"],
        ["5ArrowsGray", "5 つの矢印 (灰色)"],
        ["5Rating", "5 つの評価"],
        ["5Quarters", "5 つの四分円"],
        ["5Boxes", "5 つのボックス"],
    ];
    // Those of the menu, in Excel's groups and order.
    const COND_ICON_MENU = [
        ["方向", ["3Arrows", "3ArrowsGray", "3Triangles", "4ArrowsGray", "4Arrows", "5ArrowsGray", "5Arrows"]],
        ["図形", ["3TrafficLights1", "3TrafficLights2", "3Signs", "4TrafficLights", "4RedToBlack"]],
        ["インジケーター", ["3Symbols", "3Symbols2", "3Flags"]],
        ["評価", ["3Stars", "4Rating", "5Quarters", "5Rating", "5Boxes"]],
    ];
    const iconName = (set) => (COND_ICONS.find(([k]) => k === set) || [, set])[1];
    // The icons of a set as the grid draws them in the cells (`data.iconSets`, from the icon of
    // the lowest values): the `i`-th, or all of them from the least (`high`: from the greatest).
    const iconSvg = (set, i) => {
        const svg = data.iconSets && data.iconSets[set] && data.iconSets[set][i];
        return svg ? `<span class="ic">${svg}</span>` : "";
    };
    function iconsHtml(set, high) {
        const n = (data.iconSets && data.iconSets[set] && data.iconSets[set].length) || 0;
        const list = Array.from({ length: n }, (_, i) => iconSvg(set, i));
        if (high) list.reverse();
        return `<span class="icons">${list.join("")}</span>`;
    }
    // The cells a rule goes on: every area selected.
    const condRange = () => allAreas().map((a) => (a.kind === "cells" ? areaText(a) : rangeText(a))).join(" ");
    // A rule's range as the dialogs show it (`A1:B5 D2` → `=$A$1:$B$5,$D$2`).
    const condRefText = (range) => `=${String(range).split(/\s+/).filter(Boolean).map((a) => a.replace(/([A-Z]+|\d+)/g, "$$$1")).join(",")}`;
    // `range`: the cells (else those selected).
    function postCond(rule, range) {
        if (!editable || busy || !sel || !view) return;
        if (entry) commit(null);
        postEdit([Object.assign({ op: "conditional", sheet: view.sheet.name, range: range || condRange() }, rule)]);
    }
    // "ルールのクリア": of the cells selected, or (`all`) of the whole sheet.
    function clearCond(all) {
        if (!editable || busy || !view || (!all && !sel)) return;
        if (entry) commit(null);
        const op = { op: "conditional", sheet: view.sheet.name, clear: true };
        if (!all) op.range = condRange();
        postEdit([op]);
    }
    // The value of the active cell, to start a dialog's box with (as Excel does).
    function activeValue() {
        const cell = sel ? cellOf(sel.ar, sel.ac) : null;
        return cell && cell.t !== undefined && !cell.f ? String(cell.v !== undefined ? cell.v : cell.t) : "";
    }

    // The boxes of the dialogs: a text, a list, whether a text is there (said where it is not).
    function condInput(where, value, name) {
        const input = where.appendChild(el("input"));
        input.type = "text";
        input.spellcheck = false;
        input.dataset.field = name;
        input.value = value || "";
        input.addEventListener("input", () => input.setCustomValidity(""));
        return input;
    }
    function condSelect(where, options, value, name) {
        const s = where.appendChild(el("select"));
        s.dataset.field = name;
        for (const [v, label] of options) s.appendChild(el("option", null, label)).value = v;
        if (value !== undefined && value !== null) s.value = value;
        return s;
    }
    function condNeeded(input) {
        if (input.value.trim()) return true;
        input.setCustomValidity("値を入力してください");
        input.reportValidity();
        return false;
    }
    // What the format of a rule `f` draws, on the element `t`: what it sets, over what the
    // element has.
    function condStyle(t, f) {
        t.style.color = f.color || "";
        t.style.background = "";
        if (f.pattern) t.style.cssText += fillCss(f.pattern, f.patternColor || "#000000", f.fill || "#FFFFFF");
        else t.style.background = f.fill || "";
        t.style.fontWeight = f.bold === true ? "bold" : f.bold === false ? "normal" : "";
        t.style.fontStyle = f.italic === true ? "italic" : f.italic === false ? "normal" : "";
        const under = f.underline === true || (typeof f.underline === "string" && f.underline !== "none");
        t.style.textDecoration = [under && "underline", f.strike === true && "line-through"].filter(Boolean).join(" ") || (f.underline === "none" || f.underline === false || f.strike === false ? "none" : "");
        t.style.textDecorationStyle = /double/i.test(typeof f.underline === "string" ? f.underline : "") ? "double" : "";
        t.style.outline = f.border ? `1px solid ${f.border}` : "";
        for (const side of ["left", "right", "top", "bottom"]) {
            const line = f.borders && f.borders[side];
            const css = side[0].toUpperCase() + side.slice(1);
            t.style[`border${css}`] = line ? `1px ${{ hair: "dotted", dotted: "dotted", dashDotDot: "dotted", dashDot: "dashed", dashed: "dashed" }[line.style] || "solid"} ${line.color || "#000000"}` : "";
        }
    }

    function closeSubDialog() {
        if (subDialog) subDialog.remove();
        subDialog = null;
    }
    // A dialog over the dialog open (its form): キャンセル, Escape and a click beside it close
    // it alone (`cancelled()` then).
    function openSubDialog(title, cls, cancelled) {
        closeSubDialog();
        const back = (subDialog = document.body.appendChild(el("div", "dlgback sub")));
        const form = back.appendChild(el("form", `dlg ${cls}`));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, title));
        form.cancel = () => {
            closeSubDialog();
            if (cancelled) cancelled();
        };
        back.addEventListener("mousedown", (e) => {
            if (e.target === back) form.cancel();
        });
        return form;
    }
    // "セルの書式設定" of a rule's format, over the rule's dialog, as the pane's of the cells has
    // it: 表示形式; 文字 (フォント (スタイル・下線・色), 文字飾り (取り消し線)); 塗りつぶしと罫線
    // (塗りつぶし (背景色・パターン), 罫線 (each side, of a line and its color)). Each item is 設定なし (the cell's own
    // shows) until it is set, and taken back to it by the button at the end of its row; クリア
    // takes all back. (As Excel's: a rule's font and size are not set; nor 上付き・下付き.)
    // `done(f)`: the format set (null: none); `cancelled()`: closed without it.
    const COND_FMT_PAGES = [
        ["number", "表示形式"],
        ["font", "文字"],
        ["fill", "塗りつぶしと罫線"],
    ];
    // The lines a rule's border can have (as Excel's dialog of a rule lists them).
    const COND_LINES = ["thin", "hair", "dotted", "dashDotDot", "dashDot", "dashed"];
    const RESET_ICON = '<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.5 6.5A5 5 0 1 1 3 9.5"/><path d="M3.2 3v3.6h3.6"/></svg>';
    function askCondFormat(f0, done, cancelled) {
        const f = JSON.parse(JSON.stringify(f0 || {}));
        // (A border of each side's, as the outline of Excel 2007's is.)
        if (f.border) {
            f.borders = f.borders || {};
            for (const k of ["left", "right", "top", "bottom"]) if (!f.borders[k]) f.borders[k] = { style: "thin", color: f.border };
            delete f.border;
        }
        const form = openSubDialog("セルの書式設定", "condfmtdlg", cancelled);
        const holder = form.appendChild(el("div", "pfmt"));
        const sampleBox = form.appendChild(el("fieldset", "condfmtsample"));
        sampleBox.appendChild(el("legend", null, "プレビュー"));
        const sample = sampleBox.appendChild(el("span", null, "Aaあぁアァ亜宇"));
        // What the borders are drawn with (as the pane's 線のスタイル and 線の色).
        const pen = { line: "thin", color: null };
        const has = (k) => f[k] !== undefined && f[k] !== null;
        // Items set (null: back to 設定なし), and all drawn again (`focus`: the field focused then).
        const put = (change, focus) => {
            for (const [k, v] of Object.entries(change)) {
                if (v === undefined || v === null) delete f[k];
                else f[k] = v;
            }
            if (f.borders && !Object.keys(f.borders).length) delete f.borders;
            draw(focus);
        };
        const sideNames = { top: "上罫線", bottom: "下罫線", left: "左罫線", right: "右罫線" };
        const putSide = (k, line) => {
            const b = Object.assign({}, f.borders || {});
            if (line) b[k] = line;
            else delete b[k];
            put({ borders: b }, k);
        };
        let scrollTop = 0;
        const draw = (focus) => {
            const a = document.activeElement;
            if (!focus && a && holder.contains(a)) focus = a.dataset.field || (a.dataset.side && `side:${a.dataset.side}`);
            const old = holder.querySelector(".pbody");
            if (old) scrollTop = old.scrollTop;
            holder.textContent = "";
            const body = holder.appendChild(el("div", "pbody"));
            const kit = paneKit(body);
            const { section, line, colorRow, checkRow, selectRow } = kit;
            // The button at the end of an item's row that takes it back to 設定なし (shown when
            // the item is set).
            const reset = (control, on, label, change) => {
                const row = control.closest(".prow") || control.parentNode;
                const b = row.appendChild(el("button", "preset"));
                b.type = "button";
                b.title = `${label}を設定なしに戻す`;
                b.setAttribute("aria-label", b.title);
                b.dataset.reset = control.dataset.field;
                b.innerHTML = RESET_ICON;
                b.disabled = !on;
                b.addEventListener("click", () => put(change, control.dataset.field));
                return b;
            };

            // 表示形式
            {
                const sec = section("number", "表示形式");
                const now = has("numberFormat") ? localCode(f.numberFormat) : null;
                const list = numberSection(kit, sec, {
                    now,
                    cell: (sel && cellOf(sel.ar, sel.ac)) || {},
                    at: "condFormat",
                    put: (code) => put({ numberFormat: code }),
                    redraw: (field) => draw(field),
                });
                reset(sec.querySelector('[data-field="numberKind"]'), now !== null, "表示形式", { numberFormat: null });
                if (now === null) sec.appendChild(el("div", "pnote", "表示形式は設定されていません（セルの表示形式のまま表示します）。分類を選ぶと設定します。"));
                const top = saved.numberListTop;
                if (list && top && top.where === list.where) requestAnimationFrame(() => (list.list.scrollTop = top.top));
            }

            // フォント: スタイル, 下線, 色
            {
                const sec = section("font", "フォント");
                const style = !has("bold") && !has("italic") ? "" : f.bold && f.italic ? "boldItalic" : f.bold ? "bold" : f.italic ? "italic" : "normal";
                const st = selectRow(sec, "スタイル", [["", "（設定なし）"], ["normal", "標準"], ["italic", "斜体"], ["bold", "太字"], ["boldItalic", "太字 斜体"]], style, "fontStyle", (v) =>
                    put(v ? { bold: v.startsWith("bold"), italic: /italic/i.test(v) } : { bold: null, italic: null }),
                );
                st.options[0].hidden = true;
                reset(st, style !== "", "スタイル", { bold: null, italic: null });
                const u = f.underline === true ? "single" : f.underline === false ? "none" : f.underline || "";
                const us = selectRow(sec, "下線", [["", "（設定なし）"]].concat(UNDERLINES), u, "underline", (v) => put({ underline: v || null }));
                us.options[0].hidden = true;
                reset(us, u !== "", "下線", { underline: null });
                const fc = colorRow(sec, "色", f.color || null, "設定なし", (c) => put({ color: c }, "fontColor"), "fontColor");
                reset(fc, has("color"), "色", { color: null });
                const deco = section("font", "文字飾り");
                const strike = checkRow(deco, "取り消し線", f.strike === true, "strike", (on) => put({ strike: on }));
                strike.indeterminate = !has("strike");
                reset(strike, has("strike"), "取り消し線", { strike: null });
            }

            // 塗りつぶし: 背景色, パターンの色, パターンの種類
            {
                const sec = section("fill", "塗りつぶし");
                const pattern = f.pattern || null;
                const bg = colorRow(sec, "背景色", f.fill || null, "設定なし", (c) => put({ fill: c }, "fillColor"), "fillColor");
                reset(bg, has("fill"), "背景色", { fill: null });
                const pc = colorRow(sec, "パターンの色", f.patternColor || null, "設定なし", (c) => put({ patternColor: c }, "patternColor"), "patternColor");
                reset(pc, has("patternColor"), "パターンの色", { patternColor: null });
                const row = line(sec, "パターンの種類");
                const kind = row.appendChild(el("button", "pcolor"));
                kind.type = "button";
                kind.dataset.field = "pattern";
                const name = pattern ? (FILL_PATTERNS.find((p) => p[0] === pattern) || [, pattern])[1] : "設定なし";
                kind.title = `パターンの種類（${name}）`;
                const sw = kind.appendChild(el("span", "sw ptile"));
                if (!pattern) {
                    sw.classList.add("none");
                    kind.appendChild(el("span", "pctext", "設定なし"));
                } else sw.style.cssText = fillCss(pattern, f.patternColor || "#000000", f.fill || "#FFFFFF");
                kind.appendChild(el("span", "pdrop chev"));
                kind.addEventListener("click", () => {
                    const pop = openPop(kind, "patterns");
                    const tiles = pop.appendChild(el("div", "swatches"));
                    for (const [key, label] of FILL_PATTERNS) {
                        const b = tiles.appendChild(el("button", "swatch"));
                        b.type = "button";
                        // (単色: no pattern, the background color alone.)
                        b.title = key === "solid" ? "なし（背景色のみ）" : label;
                        b.dataset.pattern = key;
                        b.style.cssText = key === "solid" ? "background:#FFFFFF" : fillCss(key, "#000000", "#FFFFFF");
                        b.classList.toggle("on", key === pattern || (key === "solid" && !pattern));
                        b.addEventListener("click", () => {
                            closeMenu();
                            put({ pattern: key === "solid" ? null : key }, "pattern");
                        });
                    }
                });
                reset(kind, !!pattern, "パターンの種類", { pattern: null });
                if (has("patternColor") && !pattern) sec.appendChild(el("div", "pnote", "パターンの色は、パターンの種類を選ぶと使われます。"));
            }

            // 罫線: the line and its color (what the sides are drawn with), then where: by a
            // preset, by the buttons around the picture of the cell, or by the picture clicked.
            {
                const sides = f.borders || {};
                const box = section("fill", "罫線");
                const row = line(box, "スタイル");
                const pick = row.appendChild(el("button", "pcolor pline"));
                pick.type = "button";
                pick.dataset.field = "borderLine";
                pick.title = `線のスタイル（${(BORDER_LINES.find((l) => l[0] === pen.line) || BORDER_LINES[0])[1]}）`;
                pick.appendChild(el("span", "sample")).innerHTML = lineSample(pen.line, 56);
                pick.appendChild(el("span", "pdrop chev"));
                pick.addEventListener("click", () => {
                    const pop = openPop(pick, "linestyles");
                    for (const [key, label] of BORDER_LINES.filter((l) => COND_LINES.includes(l[0]))) {
                        const b = pop.appendChild(el("button"));
                        b.type = "button";
                        b.appendChild(el("span")).innerHTML = lineSample(key, 96);
                        b.title = label;
                        b.classList.toggle("on", pen.line === key);
                        b.addEventListener("click", () => {
                            closeMenu();
                            pen.line = key;
                            draw("borderLine");
                        });
                    }
                });
                colorRow(box, "色", pen.color, "自動", (c) => {
                    pen.color = c;
                    draw("borderColor");
                }, "borderColor");
                const now = () => (pen.color ? { style: pen.line, color: pen.color } : { style: pen.line });
                const presets = box.appendChild(el("div", "prow pborderpresets"));
                presets.appendChild(el("span", "plabel", "プリセット"));
                const preset = (label, icon, act) => {
                    const b = presets.appendChild(el("button", "pside"));
                    b.type = "button";
                    b.title = label;
                    b.setAttribute("aria-label", label);
                    b.dataset.field = `preset:${label}`;
                    b.appendChild(borderIcon(icon));
                    b.addEventListener("click", act);
                };
                preset("なし（設定なし）", {}, () => put({ borders: null }, "preset:なし（設定なし）"));
                preset("外枠", { t: "s", b: "s", l: "s", r: "s" }, () => put({ borders: { left: now(), right: now(), top: now(), bottom: now() } }, "preset:外枠"));
                // A side clicked: drawn with the pen, or back to 設定なし when it is drawn so already.
                const toggle = (k) => {
                    const s = sides[k];
                    const same = s && s.style === pen.line && (s.color || "").toUpperCase() === (pen.color || "").toUpperCase();
                    putSide(k, same ? null : now());
                };
                const grid = box.appendChild(el("div", "pborders"));
                const sideButton = (into, k, icon) => {
                    const b = into.appendChild(el("button", "pside"));
                    b.type = "button";
                    b.title = sideNames[k];
                    b.setAttribute("aria-label", sideNames[k]);
                    b.setAttribute("aria-pressed", String(!!sides[k]));
                    b.dataset.side = k;
                    b.dataset.field = `side:${k}`;
                    b.classList.toggle("on", !!sides[k]);
                    b.appendChild(borderIcon(icon));
                    b.addEventListener("click", () => toggle(k));
                };
                const left = grid.appendChild(el("div", "pbcol"));
                sideButton(left, "top", { t: "s" });
                sideButton(left, "bottom", { b: "s" });
                const [W, H, M] = [128, 96, 14];
                const [x1, y1, x2, y2] = [M, M, W - M, H - M];
                const stroke = (k, x, y, X, Y) => {
                    const b = sides[k];
                    if (!b) return "";
                    const [w, dash] = (BORDER_LINES.find((l) => l[0] === b.style) || BORDER_LINES[0])[2];
                    return `<path d="M${x} ${y}L${X} ${Y}" stroke="${b.color || "#000000"}" stroke-width="${w}"${dash ? ` stroke-dasharray="${dash}"` : ""}/>`;
                };
                let svg = "";
                for (const [x, y] of [[x1, y1], [x2, y1], [x1, y2], [x2, y2]]) svg += `<path d="M${x - 6} ${y}h-4M${x + 6} ${y}h4M${x} ${y - 6}v-4M${x} ${y + 6}v4" stroke="#9a9a9a" stroke-width="1"/>`;
                svg += `<text x="${W / 2}" y="${H / 2 + 4}" text-anchor="middle" font-size="11" fill="#777">文字列</text>`;
                svg += stroke("top", x1, y1, x2, y1) + stroke("bottom", x1, y2, x2, y2) + stroke("left", x1, y1, x1, y2) + stroke("right", x2, y1, x2, y2);
                const pic = grid.appendChild(el("div", "pbpic"));
                pic.title = "クリックすると、その罫線を引きます（引いてある罫線は設定なしに戻します）";
                pic.innerHTML = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" fill="none" aria-hidden="true">${svg}</svg>`;
                pic.addEventListener("click", (e) => {
                    const r = pic.firstChild.getBoundingClientRect();
                    const [x, y] = [((e.clientX - r.left) * W) / r.width, ((e.clientY - r.top) * H) / r.height];
                    const near = [["top", Math.abs(y - y1)], ["bottom", Math.abs(y - y2)], ["left", Math.abs(x - x1)], ["right", Math.abs(x - x2)]].sort((a, b) => a[1] - b[1]);
                    if (near[0][1] <= 10) toggle(near[0][0]);
                });
                grid.appendChild(el("div", "pbcorner"));
                const bottom = grid.appendChild(el("div", "pbrow"));
                sideButton(bottom, "left", { l: "s" });
                sideButton(bottom, "right", { r: "s" });
                box.appendChild(el("div", "pnote", "選んだスタイルと色で、プリセットや周りのボタン、図のクリックで罫線を引きます（同じ罫線をもう一度クリックすると設定なしに戻ります）。"));
            }

            kit.finish(COND_FMT_PAGES, CELL_ICONS, "condFmtPage", focus);
            body.scrollTop = scrollTop;
            condStyle(sample, f);
            const target = focus && holder.querySelector(`[data-field="${focus}"]`);
            if (target) target.focus({ preventScroll: true });
        };
        draw();
        const buttons = dialogButtons(form);
        const clear = buttons.insertBefore(el("button", "clear", "クリア"), buttons.firstChild);
        clear.type = "button";
        clear.title = "すべての項目を設定なしに戻します";
        clear.addEventListener("click", () => {
            for (const k of Object.keys(f)) delete f[k];
            draw();
        });
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeSubDialog();
            done(Object.keys(f).length ? JSON.parse(JSON.stringify(f)) : null);
        });
        const first = holder.querySelector(".ptab.on") || holder.querySelector(".ptab");
        if (first) first.focus();
    }

    // "負の値と軸の設定" of a data bar, over the rule's dialog, as Excel's: the negative values'
    // fill and border (their own colors, or the positive's), the axis and its color. `n`: the
    // settings ({ negativeFill, negativeFillSame, negativeBorder, negativeBorderSame, axis,
    // axisColor }), `bordered`: whether the bar has a border; `done(n)` with them changed.
    function askBarNegative(n, bordered, done, cancelled) {
        const form = openSubDialog("負の値と軸の設定", "condnegdlg", cancelled);
        const group = (title) => {
            const box = form.appendChild(el("fieldset"));
            box.appendChild(el("legend", null, title));
            return box;
        };
        const radio = (box, name, value, label, on) => {
            const line = box.appendChild(el("label"));
            const r = line.appendChild(el("input"));
            r.type = "radio";
            r.name = name;
            r.value = value;
            r.checked = on;
            line.appendChild(document.createTextNode(label));
            return line;
        };
        const colorAfter = (line, value, field) => {
            const pick = line.appendChild(colorInput(value));
            pick.dataset.field = field;
            return pick;
        };
        // Its own color, or the positive's.
        const choice = (title, name, own, same, color, field, sameLabel) => {
            const box = group(title);
            const pick = colorAfter(radio(box, name, "own", own, !same), color, field);
            radio(box, name, "same", sameLabel, same);
            const sync = () => (pick.disabled = form.querySelector(`input[name=${name}]:checked`).value === "same");
            box.addEventListener("change", sync);
            pick.addEventListener("input", () => {
                form.querySelector(`input[name=${name}][value=own]`).checked = true;
            });
            sync();
            return { box, pick, same: () => form.querySelector(`input[name=${name}]:checked`).value === "same" };
        };
        const fill = choice("負の値のバーの塗りつぶしの色", "negFill", "塗りつぶしの色:", n.negativeFillSame, n.negativeFill, "negativeFill", "正の値のバーと同じ塗りつぶしの色を適用する");
        const border = choice("負の値のバーの枠線の色", "negBorder", "枠線の色:", n.negativeBorderSame, n.negativeBorder, "negativeBorder", "正の値のバーと同じ枠線の色を適用する");
        // (Of a bar with a border only.)
        if (!bordered) for (const x of border.box.querySelectorAll("input, .colorinput")) x.disabled = true;
        const axisBox = group("軸の設定");
        axisBox.appendChild(el("div", "hint", "負の値のバーの外観を変更するには、セルの軸の位置を選択してください:"));
        radio(axisBox, "axis", "automatic", "セルの値に基づく自動調整 (可変の位置に表示)", n.axis === "automatic");
        radio(axisBox, "axis", "middle", "セルの中点", n.axis === "middle");
        radio(axisBox, "axis", "none", "なし (負の値のバーを正の値のバーと同じ方向に表示)", n.axis === "none");
        const axisLine = axisBox.appendChild(el("label", "axiscolor"));
        axisLine.appendChild(el("span", null, "軸の色:"));
        const axisColor = colorAfter(axisLine, n.axisColor, "axisColor");
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeSubDialog();
            done({
                negativeFill: fill.pick.value.toUpperCase(),
                negativeFillSame: fill.same(),
                negativeBorder: border.pick.value.toUpperCase(),
                negativeBorderSame: bordered ? border.same() : n.negativeBorderSame,
                axis: form.querySelector("input[name=axis]:checked").value,
                axisColor: axisColor.value.toUpperCase(),
            });
        });
        form.querySelector("input[name=negFill]:checked").focus();
    }

    // The values a bar, a scale or an icon set goes by (`cfvo`): their kinds, as Excel names them.
    const POINT_KINDS = { auto: "自動", min: "最小値", max: "最大値", num: "数値", percent: "パーセント", formula: "数式", percentile: "百分位" };
    const POINT_LEAST = ["min", "num", "percent", "formula", "percentile"];
    const POINT_BETWEEN = ["num", "percent", "formula", "percentile"];
    const POINT_GREATEST = ["max", "num", "percent", "formula", "percentile"];
    // A data bar's (Excel 2010's 自動 first).
    const POINT_BAR_LEAST = ["auto"].concat(POINT_LEAST);
    const POINT_BAR_GREATEST = ["auto"].concat(POINT_GREATEST);
    // The boxes of a value: its kind (`type`), its value (`value`: none for the least and the
    // greatest), and a color (`pick`) when `color` is given; the caller lays them out.
    // `read()`: `{ type, value }`, or null (said where) when the value is missing.
    function pointFields(kinds, t, color) {
        const type = el("select");
        type.dataset.field = "pointType";
        for (const k of kinds) type.appendChild(el("option", null, POINT_KINDS[k])).value = k;
        type.value = kinds.includes(t.type) ? t.type : kinds[0];
        const value = el("input");
        value.type = "text";
        value.spellcheck = false;
        value.dataset.field = "pointValue";
        value.value = t.value === undefined || t.value === null ? "" : String(t.value);
        const sync = () => {
            const none = type.value === "min" || type.value === "max" || type.value === "auto";
            value.disabled = none;
            value.placeholder = none ? `(${POINT_KINDS[type.value]})` : type.value === "formula" ? "=$A$1" : "";
            if (none) value.value = "";
        };
        type.addEventListener("change", () => {
            if (type.value === "formula" && value.value && !value.value.startsWith("=")) value.value = `=${value.value}`;
            sync();
        });
        value.addEventListener("input", () => value.setCustomValidity(""));
        sync();
        let pick = null;
        if (color) {
            pick = colorInput(color);
            pick.dataset.field = "pointColor";
        }
        const now = () => ({ type: type.value, value: value.value.trim() });
        return {
            type,
            value,
            pick,
            now,
            read: () => {
                const v = value.value.trim();
                if (!value.disabled && !v.replace(/^=/, "").trim()) {
                    value.setCustomValidity("値を入力してください");
                    value.reportValidity();
                    return null;
                }
                if (!value.disabled && !v.startsWith("=") && type.value !== "formula" && !isFinite(+v)) {
                    value.setCustomValidity("数値か、= で始まる数式を入力してください");
                    value.reportValidity();
                    return null;
                }
                return now();
            },
            color: () => pick && pick.value.toUpperCase(),
        };
    }

    // "アイコン スタイル" as Excel's: a list of the icon sets that shows their icons (from the
    // least). The button's `value` is the set chosen; `change` when another is chosen.
    function iconPicker(where, value) {
        const b = where.appendChild(el("button", "condiconpick"));
        b.type = "button";
        b.dataset.field = "icons";
        b.setAttribute("aria-haspopup", "listbox");
        const show = () => {
            b.innerHTML = `${iconsHtml(b.value)}<span class="chev"></span>`;
            b.title = iconName(b.value);
        };
        b.value = value;
        show();
        let list = null;
        const outside = (e) => {
            if (list && !list.contains(e.target) && !b.contains(e.target)) close(false);
        };
        const close = (focus) => {
            if (list) list.remove();
            list = null;
            document.removeEventListener("mousedown", outside, true);
            if (focus) b.focus();
        };
        const choose = (v) => {
            close(true);
            if (v === b.value) return;
            b.value = v;
            show();
            b.dispatchEvent(new Event("change"));
        };
        const step = (by) => {
            const at = COND_ICONS.findIndex(([k]) => k === b.value);
            const to = COND_ICONS[Math.max(0, Math.min(COND_ICONS.length - 1, at + by))][0];
            if (to !== b.value) choose(to);
        };
        const open = () => {
            list = dialog.appendChild(el("div", "condiconlist"));
            list.setAttribute("role", "listbox");
            list.tabIndex = -1;
            let at = Math.max(0, COND_ICONS.findIndex(([k]) => k === b.value));
            const items = COND_ICONS.map(([k, name], i) => {
                const o = list.appendChild(el("div", "opt"));
                o.setAttribute("role", "option");
                o.dataset.value = k;
                o.title = name;
                o.innerHTML = iconsHtml(k);
                o.addEventListener("mousemove", () => mark(i));
                o.addEventListener("click", () => choose(k));
                return o;
            });
            const mark = (i) => {
                at = i;
                items.forEach((o, j) => o.setAttribute("aria-selected", String(j === i)));
            };
            mark(at);
            const r = b.getBoundingClientRect();
            list.style.left = `${r.left}px`;
            list.style.minWidth = `${r.width}px`;
            list.style.top = `${r.bottom + list.offsetHeight > window.innerHeight - 4 ? Math.max(4, r.top - list.offsetHeight) : r.bottom}px`;
            items[at].scrollIntoView({ block: "nearest" });
            list.addEventListener("keydown", (e) => {
                // (Not the dialog's keys: Escape closes the list alone.)
                e.stopPropagation();
                if (e.key === "ArrowDown" || e.key === "ArrowUp") {
                    e.preventDefault();
                    mark(Math.max(0, Math.min(items.length - 1, at + (e.key === "ArrowDown" ? 1 : -1))));
                    items[at].scrollIntoView({ block: "nearest" });
                } else if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    choose(COND_ICONS[at][0]);
                } else if (e.key === "Escape" || e.key === "Tab") {
                    e.preventDefault();
                    close(true);
                }
            });
            document.addEventListener("mousedown", outside, true);
            list.focus();
        };
        b.addEventListener("click", () => (list ? close(true) : open()));
        // As a list box's: ↑↓ choose, Alt+↓ (F4) opens the list.
        b.addEventListener("keydown", (e) => {
            if ((e.altKey && e.key === "ArrowDown") || e.key === "F4") open();
            else if (e.key === "ArrowDown" || e.key === "ArrowUp") step(e.key === "ArrowDown" ? 1 : -1);
            else return;
            e.preventDefault();
        });
        return b;
    }

    // The kind of rule of "新しい書式ルール" a rule is of; its 書式スタイル (of "セルの値に基づいて…");
    // what of the cells it is about (of "指定の値を含むセルだけ…").
    const condKind = (r) =>
        ["dataBar", "colorScale", "iconSet"].includes(r.type)
            ? "scale"
            : ["top10", "aboveAverage", "expression"].includes(r.type)
              ? r.type
              : r.type === "duplicateValues" || r.type === "uniqueValues"
                ? "duplicateValues"
                : "contains";
    const condStyleOf = (r) => (r.type === "colorScale" ? ((r.colors || []).length === 3 ? "colorScale3" : "colorScale2") : r.type);
    const condWhatOf = (r) => (COND_TEXTS.some(([t]) => t === r.type) ? "text" : r.type);

    // Excel's dialogs of 条件付き書式: one of the quick ones of the menu (`o.kind` given), or
    // "新しい書式ルール" / "書式ルールの編集".
    function askCond(o) {
        if (!editable || busy || !sel || !view) return;
        if (entry) commit(null);
        o = o || {};
        if (o.kind) askQuickCond(o);
        else askCondRule(o);
    }

    // The quick dialogs of the menu (セルの強調表示ルール, 上位/下位ルール): `o.kind` the rule (with
    // `o.operator`, `o.bottom`, `o.percent`, `o.below`), `o.prompt` what it asks, `o.title`; the
    // format one of Excel's, or "ユーザー設定の書式…".
    function askQuickCond(o) {
        const form = openDialog(o.title, "conddlg quick");
        form.appendChild(el("div", "condprompt", o.prompt));
        const line = form.appendChild(el("div", "condline"));
        let read = () => ({});
        let first = null;
        if (o.kind === "cellIs") {
            const two = o.operator === "between" || o.operator === "notBetween";
            const a = (first = condInput(line, activeValue(), "value1"));
            let b = null;
            if (two) {
                line.appendChild(el("span", null, "と"));
                b = condInput(line, "", "value2");
            }
            read = () => (!condNeeded(a) || (b && !condNeeded(b)) ? null : { type: "cellIs", operator: o.operator, values: b ? [a.value, b.value] : [a.value] });
        } else if (o.kind === "text") {
            const t = (first = condInput(line, "", "text"));
            read = () => (condNeeded(t) ? { type: "containsText", text: t.value } : null);
        } else if (o.kind === "timePeriod") {
            const p = (first = condSelect(line, COND_PERIODS, "yesterday", "period"));
            read = () => ({ type: "timePeriod", period: p.value });
        } else if (o.kind === "duplicateValues") {
            const d = (first = condSelect(line, [["duplicateValues", "重複"], ["uniqueValues", "一意"]], "duplicateValues", "duplicate"));
            read = () => ({ type: d.value });
        } else if (o.kind === "top10") {
            const n = (first = line.appendChild(el("input")));
            n.type = "number";
            n.min = 1;
            n.max = o.percent ? 100 : 1000;
            n.value = 10;
            n.dataset.field = "rank";
            n.addEventListener("input", () => n.setCustomValidity(""));
            if (o.percent) line.appendChild(el("span", null, "%"));
            read = () => {
                const rank = Math.round(+n.value);
                if (!(rank >= 1 && rank <= +n.max)) {
                    n.setCustomValidity(`1〜${n.max} の数を入力してください`);
                    n.reportValidity();
                    return null;
                }
                const r = { type: "top10", rank };
                if (o.bottom) r.bottom = true;
                if (o.percent) r.percent = true;
                return r;
            };
        } else if (o.kind === "aboveAverage") {
            read = () => (o.below ? { type: "aboveAverage", below: true } : { type: "aboveAverage" });
        }
        const fmtLine = line.appendChild(el("label", "condfmt"));
        fmtLine.appendChild(el("span", null, "書式:"));
        const fmtSel = condSelect(fmtLine, COND_FORMATS, "red", "format");
        // "ユーザー設定の書式…": the format made in "セルの書式設定" (back to the one chosen before
        // when none is made).
        let custom = null;
        let last = fmtSel.value;
        const customize = () =>
            askCondFormat(
                custom,
                (f) => {
                    custom = f;
                    fmtSel.value = f ? "custom" : last;
                    if (f) last = "custom";
                    fmtSel.focus();
                },
                () => {
                    fmtSel.value = last;
                    fmtSel.focus();
                },
            );
        fmtSel.addEventListener("change", () => {
            fmtSel.setCustomValidity("");
            if (fmtSel.value === "custom") customize();
            else last = fmtSel.value;
        });
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const rule = read();
            if (!rule) return;
            const preset = COND_FORMATS.find(([v]) => v === fmtSel.value);
            const f = preset && preset[2] ? preset[2] : custom;
            if (!f) {
                fmtSel.setCustomValidity("書式を選んでください");
                fmtSel.reportValidity();
                return;
            }
            rule.format = f;
            closeDialog();
            postCond(rule);
        });
        (first || fmtSel).focus();
        if (first && first.select) first.select();
    }

    // "新しい書式ルール" (`o.rule`: "書式ルールの編集" of that rule; `o.startKind`, `o.startStyle`: the
    // kind of rule, the 書式スタイル to start with), as Excel's: the kind of rule
    // from Excel's list, the settings of that kind, the format ("書式…"), and (not in Excel's:
    // there a new rule goes on the cells selected) the cells it goes on, "適用先". `o.range`:
    // those cells to start with (else the cells selected); `o.done(rule, range)`: what is done
    // with the rule made (else it is put on the cells); `o.back`: what キャンセル does.
    function askCondRule(o) {
        const init = o.rule || {};
        const initKind = o.rule ? condKind(o.rule) : o.startKind || "scale";
        const initSub = o.rule ? (initKind === "scale" ? condStyleOf(init) : condWhatOf(init)) : null;
        // A setting of the rule edited, for the fields of its kind (and of its `sub`: 書式スタイル,
        // or what of the cells).
        const given = (kind, key, value, sub) => (initKind === kind && (!sub || sub === initSub) && init[key] !== undefined && init[key] !== null ? init[key] : value);
        const form = openDialog(o.rule ? "書式ルールの編集" : "新しい書式ルール", "conddlg", o.back);
        const whereLine = form.appendChild(el("label", "condtop"));
        whereLine.appendChild(el("span", "lbl", "適用先:"));
        const where = condInput(whereLine, condRefText(o.range || condRange()), "range");
        where.title = "ルールを適用するセル（=$A$1:$B$10,$D$2 のように , で区切って複数の範囲も）";
        const kindLine = form.appendChild(el("label", "condtop"));
        kindLine.appendChild(el("span", "lbl", "ルールの種類:"));
        const kinds = condSelect(kindLine, COND_KINDS, initKind, "kind");
        form.appendChild(el("div", "condcaption", "ルールの内容を編集してください:"));
        const body = form.appendChild(el("div", "condbody"));

        // The format, of all the kinds but "セルの値に基づいて…" (kept as the kind changes).
        let fmt = init.format && Object.keys(init.format).length ? Object.assign({}, init.format) : null;
        const foot = body.appendChild(el("div", "condfoot"));
        foot.appendChild(el("span", "lbl", "プレビュー:"));
        const preview = foot.appendChild(el("span", "condpreview"));
        const fmtButton = foot.appendChild(el("button", null, "書式…"));
        fmtButton.type = "button";
        fmtButton.dataset.act = "format";
        const fmtNote = body.appendChild(el("div", "condnote", "書式を設定してください（書式…）。"));
        const showFmt = () => {
            preview.textContent = fmt ? "Aaあぁアァ亜宇" : "書式が設定されていません";
            preview.classList.toggle("none", !fmt);
            condStyle(preview, fmt || {});
        };
        showFmt();
        fmtNote.hidden = true;
        fmtButton.addEventListener("click", () =>
            askCondFormat(
                fmt,
                (f) => {
                    fmt = f;
                    fmtNote.hidden = true;
                    showFmt();
                    fmtButton.focus();
                },
                () => fmtButton.focus(),
            ),
        );

        const prompt = (box, text) => box.appendChild(el("div", "condprompt", text));
        const lineOf = (box) => box.appendChild(el("div", "condline"));
        // The kinds of rules: each fills its box, and gives `read()` (the rule's settings, or null:
        // something is missing, said where).
        const KINDS = {
            scale(box) {
                prompt(box, "セルの値に基づいてすべてのセルを書式設定:");
                const top = lineOf(box);
                const styleLine = top.appendChild(el("label"));
                styleLine.appendChild(el("span", "lbl", "書式スタイル:"));
                const style = condSelect(styleLine, COND_STYLES, (initKind === "scale" && initSub) || o.startStyle || "colorScale2", "style");
                // Each 書式スタイル's boxes, made when first chosen (and kept as it changes).
                const styles = {};
                const showStyle = () => {
                    if (!styles[style.value]) styles[style.value] = STYLES[style.value](box.appendChild(el("div", "condstyle")), top);
                    for (const [k, s] of Object.entries(styles)) {
                        s.box.hidden = k !== style.value;
                        for (const x of s.top) x.hidden = k !== style.value;
                    }
                };
                style.addEventListener("change", showStyle);
                showStyle();
                return { read: () => styles[style.value].read(), focus: () => style.focus() };
            },
            contains(box) {
                prompt(box, "次のセルのみを書式設定:");
                const line = lineOf(box);
                const what = condSelect(line, COND_CONTAINS, initKind === "contains" ? initSub : "cellIs", "what");
                const op = condSelect(line, COND_OPERATORS, given("contains", "operator", "between", "cellIs"), "operator");
                const values = given("contains", "values", null, "cellIs");
                const a = condInput(line, values ? values[0] : "", "value1");
                const and = line.appendChild(el("span", null, "と"));
                const b = condInput(line, values ? values[1] || "" : "", "value2");
                const how = condSelect(line, COND_TEXTS, initSub === "text" ? init.type : "containsText", "textKind");
                const t = condInput(line, given("contains", "text", "", "text"), "text");
                const period = condSelect(line, COND_PERIODS, given("contains", "period", "yesterday", "timePeriod"), "period");
                const two = () => op.value === "between" || op.value === "notBetween";
                const show = () => {
                    const w = what.value;
                    op.hidden = a.hidden = w !== "cellIs";
                    and.hidden = b.hidden = w !== "cellIs" || !two();
                    how.hidden = t.hidden = w !== "text";
                    period.hidden = w !== "timePeriod";
                };
                what.addEventListener("change", show);
                op.addEventListener("change", show);
                show();
                return {
                    read: () => {
                        const w = what.value;
                        if (w === "cellIs") return !condNeeded(a) || (two() && !condNeeded(b)) ? null : { type: "cellIs", operator: op.value, values: two() ? [a.value, b.value] : [a.value] };
                        if (w === "text") return condNeeded(t) ? { type: how.value, text: t.value } : null;
                        if (w === "timePeriod") return { type: "timePeriod", period: period.value };
                        return { type: w };
                    },
                    focus: () => what.focus(),
                };
            },
            top10(box) {
                prompt(box, "次に入る値を書式設定:");
                const line = lineOf(box);
                const which = condSelect(line, [["top", "上位"], ["bottom", "下位"]], given("top10", "bottom", false) ? "bottom" : "top", "which");
                const n = line.appendChild(el("input"));
                n.type = "number";
                n.min = 1;
                n.max = 1000;
                n.value = given("top10", "rank", 10);
                n.dataset.field = "rank";
                n.addEventListener("input", () => n.setCustomValidity(""));
                const pct = check(line, "% (選択範囲に占める割合)", !!given("top10", "percent", false));
                pct.dataset.field = "percent";
                return {
                    read: () => {
                        const rank = Math.round(+n.value);
                        const max = pct.checked ? 100 : 1000;
                        if (!(rank >= 1 && rank <= max)) {
                            n.setCustomValidity(`1〜${max} の数を入力してください`);
                            n.reportValidity();
                            return null;
                        }
                        const r = { type: "top10", rank };
                        if (which.value === "bottom") r.bottom = true;
                        if (pct.checked) r.percent = true;
                        return r;
                    },
                    focus: () => which.focus(),
                };
            },
            aboveAverage(box) {
                prompt(box, "次の値を書式設定:");
                const line = lineOf(box);
                line.appendChild(el("span", null, "選択範囲の平均値"));
                const sd = given("aboveAverage", "stdDev", 0);
                const dir = given("aboveAverage", "below", false) ? "below" : "above";
                const avg = condSelect(line, COND_AVERAGES, sd ? `${dir}${sd}` : `${dir}${given("aboveAverage", "equal", false) ? "Eq" : ""}`, "average");
                return {
                    read: () => {
                        const [, d, x] = /^(above|below)(Eq|\d)?$/.exec(avg.value);
                        const r = { type: "aboveAverage" };
                        if (d === "below") r.below = true;
                        if (x === "Eq") r.equal = true;
                        else if (x) r.stdDev = +x;
                        return r;
                    },
                    focus: () => avg.focus(),
                };
            },
            duplicateValues(box) {
                prompt(box, "すべての値を書式設定:");
                const line = lineOf(box);
                const d = condSelect(line, [["duplicateValues", "重複"], ["uniqueValues", "一意"]], initKind === "duplicateValues" ? init.type : "duplicateValues", "duplicate");
                line.appendChild(el("span", null, "値 (選択範囲内)"));
                return { read: () => ({ type: d.value }), focus: () => d.focus() };
            },
            expression(box) {
                prompt(box, "次の数式を満たす場合に値を書式設定:");
                const f = condInput(box, given("expression", "formula", ""), "formula");
                f.classList.add("wide");
                f.placeholder = "=$A1>100";
                return {
                    read: () => {
                        if (f.value.trim().replace(/^=/, "").trim()) return { type: "expression", formula: f.value.trim() };
                        f.setCustomValidity("数式を入力してください");
                        f.reportValidity();
                        return null;
                    },
                    focus: () => f.focus(),
                };
            },
        };
        // A setting of the 書式スタイル `k` of the rule edited.
        const scaleGiven = (k, key, value) => given("scale", key, value, k);
        // The values (`cfvo`) of a scale or a bar in columns (最小値・中間点・最大値), their
        // 種類・値 (and 色) in rows, as Excel lays them out.
        const columns = (where, heads, fields, withColor) => {
            const grid = where.appendChild(el("div", "condcols"));
            grid.style.gridTemplateColumns = `auto repeat(${heads.length}, minmax(0, 1fr))`;
            grid.appendChild(el("span"));
            for (const h of heads) grid.appendChild(el("span", "head", h));
            const rows = [["種類:", "type"], ["値:", "value"]].concat(withColor ? [["色:", "pick"]] : []);
            for (const [label, key] of rows) {
                grid.appendChild(el("span", "lbl", label));
                for (const f of fields) grid.appendChild(f[key]);
            }
            return grid;
        };
        const previewLine = (where) => {
            const l = where.appendChild(el("div", "condline condscalepreview"));
            l.appendChild(el("span", "lbl", "プレビュー:"));
            return l.appendChild(el("span", "condscalebar"));
        };
        const colorScale = (three) => (box) => {
            const k = three ? "colorScale3" : "colorScale2";
            // (Excel's colors to start with.)
            const was = scaleGiven(k, "colors", null) || (three ? ["#F8696B", "#FFEB84", "#63BE7B"] : ["#FF7128", "#FFEF9C"]);
            const t = scaleGiven(k, "thresholds", null) || (three ? [{ type: "min" }, { type: "percentile", value: "50" }, { type: "max" }] : [{ type: "min" }, { type: "max" }]);
            const fields = [pointFields(POINT_LEAST, t[0], was[0])];
            if (three) fields.push(pointFields(POINT_BETWEEN, t[1], was[1]));
            fields.push(pointFields(POINT_GREATEST, t[t.length - 1], was[was.length - 1]));
            columns(box, three ? ["最小値", "中間点", "最大値"] : ["最小値", "最大値"], fields, true);
            const bar = previewLine(box);
            const paint = () => (bar.style.background = `linear-gradient(to right, ${fields.map((f) => f.color()).join(", ")})`);
            for (const f of fields) f.pick.addEventListener("input", paint);
            paint();
            return {
                box,
                top: [],
                read: () => {
                    const points = fields.map((f) => f.read());
                    if (points.includes(null)) return null;
                    return { type: "colorScale", colors: fields.map((f) => f.color()), thresholds: points };
                },
            };
        };
        const STYLES = {
            colorScale2: colorScale(false),
            colorScale3: colorScale(true),
            dataBar(box, top) {
                const only = check(top, "棒のみ表示", !!scaleGiven("dataBar", "barOnly", false));
                only.dataset.field = "barOnly";
                const t = scaleGiven("dataBar", "thresholds", null) || [{ type: "auto" }, { type: "auto" }];
                const fields = [pointFields(POINT_BAR_LEAST, t[0]), pointFields(POINT_BAR_GREATEST, t[1])];
                columns(box, ["最小値", "最大値"], fields, false);
                // バーの外観: 塗りつぶし and its color, 枠線 and its color (as Excel's, a border of
                // the bar's color to start with).
                box.appendChild(el("div", "condsub", "バーの外観:"));
                const look = box.appendChild(el("div", "condcols"));
                look.style.gridTemplateColumns = "minmax(0, 2.1fr) minmax(0, .8fr) minmax(0, 1.4fr) minmax(0, .8fr)";
                for (const h of ["塗りつぶし", "色", "枠線", "色"]) look.appendChild(el("span", "head", h));
                const fill = condSelect(look, [["gradient", "塗りつぶし (グラデーション)"], ["solid", "塗りつぶし (単色)"]], scaleGiven("dataBar", "solid", false) ? "solid" : "gradient", "barFill");
                const color = (value, name) => {
                    const pick = look.appendChild(colorInput(value));
                    pick.dataset.field = name;
                    return pick;
                };
                const c = color((scaleGiven("dataBar", "colors", null) || ["#638EC6"])[0], "barColor");
                const given = scaleGiven("dataBar", "barBorder", undefined);
                const line = condSelect(look, [["none", "枠線なし"], ["solid", "枠線 (実線)"]], given === false ? "none" : "solid", "barBorder");
                const bc = color(typeof given === "string" ? given : c.value, "barBorderColor");
                // (The border's color follows the bar's until it is chosen.)
                let follow = typeof given !== "string";
                bc.addEventListener("input", () => (follow = false));
                // "負の値と軸…" (its own dialog) and "棒の方向".
                let neg = {
                    negativeFill: scaleGiven("dataBar", "negativeFill", "#FF0000"),
                    negativeFillSame: !!scaleGiven("dataBar", "negativeFillSame", false),
                    negativeBorder: scaleGiven("dataBar", "negativeBorder", "#FF0000"),
                    negativeBorderSame: !!scaleGiven("dataBar", "negativeBorderSame", false),
                    axis: scaleGiven("dataBar", "axis", "automatic"),
                    axisColor: scaleGiven("dataBar", "axisColor", "#000000"),
                };
                const more = box.appendChild(el("div", "condline condbarmore"));
                const negButton = more.appendChild(el("button", null, "負の値と軸…"));
                negButton.type = "button";
                negButton.dataset.act = "negative";
                negButton.addEventListener("click", () =>
                    askBarNegative(
                        neg,
                        line.value !== "none",
                        (n) => {
                            neg = n;
                            negButton.focus();
                        },
                        () => negButton.focus(),
                    ),
                );
                const dirLine = more.appendChild(el("label"));
                dirLine.appendChild(el("span", null, "棒の方向:"));
                const direction = condSelect(dirLine, [["context", "コンテキスト"], ["leftToRight", "左から右"], ["rightToLeft", "右から左"]], scaleGiven("dataBar", "direction", "context"), "direction");
                // (The bar in its box, as in a cell.)
                const bar = previewLine(box).appendChild(el("span", "barfill"));
                const paint = () => {
                    if (follow) bc.value = c.value;
                    bc.disabled = line.value === "none";
                    const rtl = direction.value === "rightToLeft";
                    bar.style.background = fill.value === "solid" ? c.value : `linear-gradient(to ${rtl ? "left" : "right"}, ${c.value}, #fff)`;
                    bar.style.borderColor = line.value === "none" ? "transparent" : bc.value;
                    bar.style.marginLeft = rtl ? "auto" : "";
                };
                for (const x of [fill, c, line, bc, direction]) x.addEventListener(x.tagName === "SELECT" ? "change" : "input", paint);
                paint();
                return {
                    box,
                    top: [only.parentNode],
                    read: () => {
                        const points = fields.map((f) => f.read());
                        if (points.includes(null)) return null;
                        const r = Object.assign({ type: "dataBar", colors: [c.value.toUpperCase()], thresholds: points, barBorder: line.value === "none" ? false : bc.value.toUpperCase(), direction: direction.value }, neg);
                        if (fill.value === "solid") r.solid = true;
                        if (only.checked) r.barOnly = true;
                        return r;
                    },
                };
            },
            iconSet(box, top) {
                const rev = top.appendChild(el("button", "condreverse", "アイコンの順序を逆にする"));
                rev.type = "button";
                rev.dataset.act = "reverse";
                let reversed = !!scaleGiven("iconSet", "reverse", false);
                const line = lineOf(box);
                const sl = line.appendChild(el("label"));
                sl.appendChild(el("span", "lbl", "アイコン スタイル:"));
                const s = iconPicker(sl, scaleGiven("iconSet", "icons", "3TrafficLights1"));
                const only = check(line, "アイコンのみ表示", !!scaleGiven("iconSet", "iconOnly", false));
                only.dataset.field = "iconOnly";
                box.appendChild(el("div", "condsub", "次のルールに従って各アイコンを表示:"));
                const grid = box.appendChild(el("div", "condicongrid"));
                // The values of the icons (the first: where the least starts), kept as the set
                // changes when it has as many icons.
                let points = scaleGiven("iconSet", "thresholds", null);
                let rows = [];
                const draw = () => {
                    const n = data.iconSets[s.value].length;
                    if (rows.length) points = [points[0]].concat(rows.map((r) => r.now()).reverse());
                    if (!points || points.length !== n) points = Array.from({ length: n }, (_, i) => ({ type: "percent", value: String(Math.round((i * 100) / n)), gte: true }));
                    // The icon of the k-th line (the greatest first).
                    const icon = (k) => iconSvg(s.value, reversed ? k : n - 1 - k);
                    rev.setAttribute("aria-pressed", String(reversed));
                    grid.textContent = "";
                    for (const h of ["アイコン", "", "", "値", "種類"]) grid.appendChild(el("span", "head", h));
                    rows = [];
                    const texts = [];
                    // What the icon below one is for: under its value.
                    const relabel = () => {
                        for (let k = 1; k < n; k++) {
                            const p = rows[k - 1].now();
                            texts[k].textContent = `値 ${p.gte ? "<" : "<="} ${p.value === "" ? "?" : p.value}${k < n - 1 ? " および" : ""}`;
                        }
                    };
                    // The greatest first, as Excel lists them.
                    for (let k = 0; k < n; k++) {
                        grid.appendChild(el("span", "icon")).innerHTML = icon(k);
                        texts.push(grid.appendChild(el("span", "what", k === 0 ? "値が" : "")));
                        if (k === n - 1) {
                            texts[k].classList.add("last");
                            break;
                        }
                        const t = points[n - 1 - k];
                        const op = condSelect(grid, [["1", ">="], ["0", ">"]], t.gte === false ? "0" : "1", "gte");
                        const point = pointFields(POINT_BETWEEN, t);
                        grid.appendChild(point.value);
                        grid.appendChild(point.type);
                        for (const x of [op, point.type]) x.addEventListener("change", relabel);
                        point.value.addEventListener("input", relabel);
                        rows.push({
                            read: () => {
                                const x = point.read();
                                return x && Object.assign(x, { gte: op.value === "1" });
                            },
                            now: () => Object.assign(point.now(), { gte: op.value === "1" }),
                        });
                    }
                    relabel();
                };
                s.addEventListener("change", draw);
                rev.addEventListener("click", () => {
                    reversed = !reversed;
                    draw();
                });
                draw();
                return {
                    box,
                    top: [rev],
                    read: () => {
                        const read = rows.map((r) => r.read());
                        if (read.includes(null)) return null;
                        const r = { type: "iconSet", icons: s.value, thresholds: [points[0]].concat(read.reverse()) };
                        if (reversed) r.reverse = true;
                        if (only.checked) r.iconOnly = true;
                        return r;
                    },
                };
            },
        };

        // Each kind's box, made when first chosen (and kept, as Excel keeps what was set).
        const panes = {};
        const showKind = () => {
            const k = kinds.value;
            if (!panes[k]) {
                const box = body.insertBefore(el("div", "condpane"), foot);
                panes[k] = Object.assign({ box }, KINDS[k](box));
            }
            for (const [key, p] of Object.entries(panes)) p.box.hidden = key !== k;
            foot.hidden = k === "scale";
            fmtNote.hidden = true;
        };
        kinds.addEventListener("change", showKind);
        showKind();
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const areas = condAreas(where.value);
            if (!areas) {
                where.setCustomValidity("このシートのセル範囲を指定してください（例: =$A$1:$B$10）");
                where.reportValidity();
                return;
            }
            const pane = panes[kinds.value];
            const rule = pane.read();
            if (!rule) return;
            if (kinds.value !== "scale") {
                if (!fmt) {
                    fmtNote.hidden = false;
                    fmtButton.focus();
                    return;
                }
                rule.format = fmt;
            }
            const range = condRangeText(areas);
            if (o.done) return o.done(rule, range);
            closeDialog();
            postCond(rule, range);
        });
        kinds.focus();
    }

    // ---- "条件付き書式ルールの管理": the rules of the sheet (or of the cells selected) in
    // their order of priority; new ones, edited, deleted, moved up and down, their ranges and
    // "条件を満たす場合は停止" changed. OK writes them all (`conditionalRules`).

    const COND_OPS = { equal: "=", notEqual: "<>", greaterThan: ">", lessThan: "<", greaterThanOrEqual: ">=", lessThanOrEqual: "<=" };
    const COND_TEXT_OPS = { containsText: "を含む", notContainsText: "を含まない", beginsWith: "で始まる", endsWith: "で終わる" };
    // What a rule is, as the dialog lists it.
    function condText(r) {
        const v = r.values || [];
        switch (r.type) {
            case "cellIs":
                if (r.operator === "between" || r.operator === "notBetween") return `セルの値が ${v[0]} と ${v[1]} の間${r.operator === "notBetween" ? "以外" : ""}`;
                return `セルの値 ${COND_OPS[r.operator] || r.operator} ${v[0]}`;
            case "containsText":
            case "notContainsText":
            case "beginsWith":
            case "endsWith":
                return `文字列が "${r.text}" ${COND_TEXT_OPS[r.type]}`;
            case "timePeriod":
                return `日付: ${(COND_PERIODS.find(([p]) => p === r.period) || [, r.period])[1]}`;
            case "containsBlanks":
                return "空白";
            case "notContainsBlanks":
                return "空白なし";
            case "containsErrors":
                return "エラー";
            case "notContainsErrors":
                return "エラーなし";
            case "duplicateValues":
                return "重複する値";
            case "uniqueValues":
                return "一意の値";
            case "top10":
                return `${r.bottom ? "下位" : "上位"} ${r.rank}${r.percent ? "%" : ""}`;
            case "aboveAverage":
                if (r.stdDev) return `平均から標準偏差 ${r.stdDev} ${r.below ? "下" : "上"}`;
                return `平均${r.below ? (r.equal ? "以下" : "より下") : r.equal ? "以上" : "より上"}`;
            case "expression":
                return `数式: ${r.formula}`;
            case "dataBar":
                return "データ バー";
            case "colorScale":
                return "グラデーション カラー スケール";
            case "iconSet":
                return "アイコン セット";
            default:
                return r.type || "（不明なルール）";
        }
    }
    // A sample of what a rule draws.
    function condSample(r) {
        const s = el("span", "condsample");
        if (r.type === "dataBar") {
            const c = (r.colors || ["#638EC6"])[0];
            const b = s.appendChild(el("span", "bar"));
            b.style.background = r.solid ? c : `linear-gradient(to right, ${c}, #fff)`;
            if (r.barBorder) b.style.outline = `1px solid ${r.barBorder}`;
        } else if (r.type === "colorScale") s.innerHTML = `<span class="scale" style="background:linear-gradient(to right, ${(r.colors && r.colors.length > 1 ? r.colors : ["#F8696B", "#63BE7B"]).join(", ")})"></span>`;
        else if (r.type === "iconSet") {
            // (From the least, as the rule has them.)
            s.innerHTML = data.iconSets && data.iconSets[r.icons] ? iconsHtml(r.icons, !!r.reverse) : `<span class="icons">${r.icons || ""}</span>`;
        } else {
            const f = r.format || {};
            const t = s.appendChild(el("span", "fmt", "Aaあぁアァ亜宇"));
            condStyle(t, f);
            if (!Object.keys(f).length) t.textContent = "書式なし";
        }
        return s;
    }
    // The areas of a range of a rule (`A1:B5 D2`), or null when one cannot be read.
    function condAreas(text) {
        const out = [];
        for (const part of String(text).replace(/^=/, "").split(/[\s,]+/).filter(Boolean)) {
            const ref = parseRef(part.replace(/^=/, ""), view.sheet);
            if (!ref || ref.sheet !== view.sheet) return null;
            out.push(ref);
        }
        return out.length ? out : null;
    }
    const condRangeText = (areas) => areas.map((a) => (a.kind === "cells" ? areaText(a) : rangeText(a))).join(" ");
    // Some of `list` (areas of the sheet shown, or null) is among the cells selected.
    const touchesSelection = (list) => !!list && list.some((a) => allAreas().some((s) => a.r1 <= s.r2 && a.r2 >= s.r1 && a.c1 <= s.c2 && a.c2 >= s.c1));
    // The areas a name refers to (`Sheet1!$A$1:$B$5,Sheet1!$D$2`; on the sheet of the name, or
    // the one shown, when they name none), or null when it refers to something else (a
    // formula, a constant).
    function refAreas(n) {
        const home = (n.sheet && sheetByName(n.sheet)) || view.sheet;
        const parts = [];
        let [part, quoted, depth] = ["", false, 0];
        for (const ch of String(n.ref).replace(/^=/, "")) {
            if (ch === "'") quoted = !quoted;
            else if (!quoted && ch === "(") depth++;
            else if (!quoted && ch === ")") depth--;
            if (ch === "," && !quoted && depth === 0) {
                parts.push(part);
                part = "";
            } else part += ch;
        }
        parts.push(part);
        const out = [];
        for (const p of parts) {
            const ref = parseRef(p, home);
            if (!ref) return null;
            out.push(ref);
        }
        return out.length ? out : null;
    }
    // The areas of the sheet shown a name refers to, or null (another sheet, not a range).
    function nameAreas(n) {
        const list = refAreas(n);
        return list && list.every((a) => a.sheet === view.sheet) ? list : null;
    }
    // Selects areas of the sheet shown (the last one with the active cell).
    function selectAreas(list) {
        if (!list || !list.length) return;
        if (entry) commit(null);
        const last = list[list.length - 1];
        const rest = list.slice(0, -1).map((a) => ({ kind: a.kind, r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }));
        const s = last.kind === "rows" ? { kind: "rows", nr: last.r1, nc: 0, fr: last.r2, fc: 0 } : last.kind === "cols" ? { kind: "cols", nr: 0, nc: last.c1, fr: 0, fc: last.c2 } : areaSel(last.r1, last.c1, last.r2, last.c2);
        setSel(Object.assign(s, { areas: rest }), true);
    }

    // The rules of the sheet as "条件付き書式ルールの管理" has them, at first.
    function condManageState() {
        const list = (view.sheet.condRules || []).map((r) => ({ id: r.id, now: Object.assign({}, r), range: r.range, stop: !!r.stop, edited: false }));
        return { list, scope: "selection", chosen: -1, changed: false };
    }
    // `state`: the list as the dialog has it (kept while a rule is edited in its own dialog).
    function manageCond(state) {
        if (!editable || busy || !view) return;
        if (entry) commit(null);
        if (!state) state = condManageState();
        const form = openDialog("条件付き書式ルールの管理", "condmgr");
        const scopeLine = form.appendChild(el("label"));
        scopeLine.appendChild(el("span", "lbl", "書式ルールの表示:"));
        const scope = scopeLine.appendChild(el("select"));
        scope.dataset.field = "scope";
        for (const [v, label] of [["selection", "現在の選択範囲"], ["sheet", "このワークシート"]]) scope.appendChild(el("option", null, label)).value = v;
        scope.value = state.scope;
        const tools = form.appendChild(el("div", "condtools"));
        const tool = (label, key, act) => {
            const b = tools.appendChild(el("button", null, label));
            b.type = "button";
            b.dataset.act = key;
            b.addEventListener("click", act);
            return b;
        };
        const table = form.appendChild(el("div", "condlist"));
        // The rules shown: those of the sheet, or those on the cells selected.
        const shown = () =>
            state.list
                .map((r, i) => i)
                .filter((i) => {
                    if (state.scope === "sheet" || !sel) return true;
                    const areas = condAreas(state.list[i].range) || [];
                    return areas.some((a) => allAreas().some((s) => a.r1 <= s.r2 && a.r2 >= s.r1 && a.c1 <= s.c2 && a.c2 >= s.c1));
                });
        const reopen = () => manageCond(state);
        if (state.chosen === -1 && state.scope === "selection" && !shown().length) {
            state.scope = "sheet";
            scope.value = "sheet";
        }
        const add = tool("新規ルール…", "new", () =>
            askCond({
                back: reopen,
                done: (rule, range) => {
                    // Excel puts a new rule first.
                    state.list.unshift({ id: null, now: rule, range, stop: false, edited: true });
                    state.chosen = 0;
                    state.changed = true;
                    reopen();
                },
            }),
        );
        const edit = tool("ルールの編集…", "edit", () => {
            const r = state.list[state.chosen];
            if (!r || r.now.editable === false) return;
            askCond({
                rule: r.now,
                range: r.range,
                back: reopen,
                done: (rule, range) => {
                    r.now = rule;
                    r.range = range;
                    r.edited = true;
                    state.changed = true;
                    reopen();
                },
            });
        });
        const del = tool("ルールの削除", "delete", () => {
            if (state.chosen < 0) return;
            const was = state.chosen;
            state.list.splice(was, 1);
            const left = shown();
            state.chosen = left.find((i) => i >= was) ?? (left.length ? left[left.length - 1] : -1);
            state.changed = true;
            draw();
        });
        const move = (up) => {
            const list = shown();
            const at = list.indexOf(state.chosen);
            const to = list[up ? at - 1 : at + 1];
            if (at < 0 || to === undefined) return;
            const l = state.list;
            [l[state.chosen], l[to]] = [l[to], l[state.chosen]];
            state.chosen = to;
            state.changed = true;
            draw();
        };
        const upBtn = tool("", "up", () => move(true));
        upBtn.title = "上へ移動（優先順位を上げる）";
        upBtn.className = "iconbtn";
        upBtn.appendChild(el("span", "chev up"));
        const downBtn = tool("", "down", () => move(false));
        downBtn.title = "下へ移動（優先順位を下げる）";
        downBtn.className = "iconbtn";
        downBtn.appendChild(el("span", "chev"));
        function draw() {
            table.textContent = "";
            const head = table.appendChild(el("div", "condhead"));
            for (const t of ["ルール（表示順で適用）", "書式", "適用先", "条件を満たす場合は停止"]) head.appendChild(el("span", null, t));
            const list = shown();
            if (!list.includes(state.chosen)) state.chosen = list.length ? list[0] : -1;
            for (const i of list) {
                const r = state.list[i];
                const row = table.appendChild(el("div", `condrule${i === state.chosen ? " on" : ""}`));
                row.dataset.id = r.id || "new";
                row.appendChild(el("span", "what", condText(r.now)));
                row.appendChild(el("span", "how")).appendChild(condSample(r.now));
                const where = row.appendChild(el("input"));
                where.type = "text";
                where.spellcheck = false;
                where.dataset.field = "range";
                where.value = condRefText(r.range);
                where.addEventListener("change", () => {
                    const areas = condAreas(where.value);
                    if (!areas) {
                        where.setCustomValidity("参照が正しくありません");
                        where.reportValidity();
                        return;
                    }
                    where.setCustomValidity("");
                    r.range = condRangeText(areas);
                    state.changed = true;
                });
                const stopBox = row.appendChild(el("span", "stop")).appendChild(el("input"));
                stopBox.type = "checkbox";
                stopBox.checked = r.stop;
                stopBox.disabled = ["dataBar", "colorScale", "iconSet"].includes(r.now.type);
                stopBox.addEventListener("change", () => {
                    r.stop = stopBox.checked;
                    state.changed = true;
                });
                row.addEventListener("mousedown", () => {
                    state.chosen = i;
                    for (const x of table.querySelectorAll(".condrule.on")) x.classList.remove("on");
                    row.classList.add("on");
                    showTools();
                });
                row.addEventListener("dblclick", (e) => {
                    if (e.target.tagName !== "INPUT") edit.click();
                });
            }
            if (!list.length) table.appendChild(el("div", "condnone", state.scope === "sheet" ? "このシートに条件付き書式はありません。" : "選択範囲に条件付き書式はありません。"));
            showTools();
        }
        function showTools() {
            const list = shown();
            const at = list.indexOf(state.chosen);
            const r = state.list[state.chosen];
            edit.disabled = !r || r.now.editable === false;
            edit.title = r && r.now.editable === false ? "このルール（Excel 2010 以降の形式など）は、このダイアログでは編集できません（削除・移動・適用先の変更はできます）" : "";
            del.disabled = at < 0;
            upBtn.disabled = at <= 0;
            downBtn.disabled = at < 0 || at >= list.length - 1;
            add.disabled = !sel;
        }
        scope.addEventListener("change", () => {
            state.scope = scope.value;
            draw();
        });
        draw();
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const bad = [...table.querySelectorAll("input[data-field=range]")].find((i) => !condAreas(i.value));
            if (bad) {
                bad.setCustomValidity("参照が正しくありません");
                bad.reportValidity();
                return;
            }
            closeDialog();
            if (!state.changed) return;
            const before = new Map((view.sheet.condRules || []).map((r) => [r.id, r]));
            const rules = state.list.map((r) => {
                const out = r.id ? { id: r.id } : {};
                const was = r.id && before.get(r.id);
                if (r.edited) {
                    const rule = Object.assign({}, r.now);
                    delete rule.id;
                    delete rule.range;
                    delete rule.editable;
                    delete rule.stop;
                    Object.assign(out, rule);
                }
                if (!was || r.range !== was.range) out.range = r.range;
                if (!was || r.stop !== !!was.stop || r.edited) out.stop = r.stop;
                return out;
            });
            postEdit([{ op: "conditionalRules", sheet: view.sheet.name, rules }]);
        });
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
    // An item of the menu `pop` that opens a submenu beside it (`fill(sub)` fills it): on the
    // mouse over it, a click, → or Enter (← in the submenu goes back). One open at a time.
    function subMenu(pop, label, fill) {
        const b = pop.appendChild(el("button", "hassub"));
        b.type = "button";
        b.setAttribute("aria-haspopup", "menu");
        b.appendChild(el("span", null, label));
        b.appendChild(el("i", "subarrow"));
        const open = (focus) => {
            if (pop.sub && pop.sub.owner === b) {
                if (focus) pop.sub.focusFirst();
                return;
            }
            if (pop.sub) pop.sub.close();
            const sub = pop.appendChild(el("div", "menu pop sub"));
            sub.setAttribute("role", "menu");
            fill(sub);
            b.classList.add("open");
            b.setAttribute("aria-expanded", "true");
            // Beside the item (on the left when there is no room on the right).
            const r = b.getBoundingClientRect();
            const p = pop.getBoundingClientRect();
            const left = p.right + sub.offsetWidth > window.innerWidth - 4 ? Math.max(0, p.left - sub.offsetWidth) : p.right - 1;
            sub.style.left = `${left}px`;
            sub.style.top = `${Math.max(0, Math.min(r.top - 5, window.innerHeight - sub.offsetHeight - 4))}px`;
            const close = () => {
                sub.remove();
                b.classList.remove("open");
                b.setAttribute("aria-expanded", "false");
                if (pop.sub && pop.sub.owner === b) pop.sub = null;
            };
            const focusFirst = () => {
                const first = sub.querySelector("button");
                if (first) first.focus();
            };
            pop.sub = { owner: b, close, focusFirst };
            sub.addEventListener("keydown", (e) => {
                if (e.key !== "ArrowLeft") return;
                e.preventDefault();
                e.stopPropagation();
                close();
                b.focus();
            });
            if (focus) focusFirst();
        };
        b.addEventListener("mouseenter", () => open(false));
        b.addEventListener("click", () => open(true));
        b.addEventListener("keydown", (e) => {
            if (e.key !== "ArrowRight") return;
            e.preventDefault();
            open(true);
        });
        // The mouse over another item of the menu closes the submenu.
        if (!pop.subWatch) {
            pop.subWatch = true;
            pop.addEventListener("mouseover", (e) => {
                const item = e.target.closest && e.target.closest("button");
                if (item && item.parentNode === pop && pop.sub && pop.sub.owner !== item) pop.sub.close();
            });
        }
        return b;
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
        fillColors(openPop(anchor, "colors"), name, none, pick);
    }
    function fillColors(pop, name, none, pick) {
        if (none) popItem(pop, none, () => (pick ? pick(null) : format({ [name]: null })));
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
        // Any other color: picked on its square (saturation and brightness) and hue, or typed.
        const start = (name && lastColor[name]) || "#4472C4";
        colorPicker(pop, start, (color) => {
            closeMenu();
            focusGrid();
            if (pick) pick(color);
            else useColor(name, color);
        });
    }
    // A color picker (as shadcn's): a square of saturation × brightness, a bar of hue, the color as
    // hex or RGB, OK (`done("#RRGGBB")`).
    function colorPicker(into, start, done) {
        const form = into.appendChild(el("form", "picker"));
        const area = form.appendChild(el("div", "sv"));
        const dot = area.appendChild(el("span", "knob"));
        const hueBar = form.appendChild(el("div", "hue"));
        const hueDot = hueBar.appendChild(el("span", "knob"));
        for (const [k, label] of [[dot, "彩度と明るさ"], [hueDot, "色相"]]) {
            k.tabIndex = 0;
            k.setAttribute("role", "slider");
            k.setAttribute("aria-label", label);
        }
        const row = form.appendChild(el("div", "pickrow"));
        const mode = row.appendChild(el("select"));
        mode.setAttribute("aria-label", "表記");
        for (const [v, label] of [["hex", "Hex"], ["rgb", "RGB"]]) mode.appendChild(el("option", null, label)).value = v;
        const box = row.appendChild(el("label", "pickvalue"));
        const hash = box.appendChild(el("span", "hash", "#"));
        const input = box.appendChild(el("input"));
        input.type = "text";
        input.spellcheck = false;
        input.setAttribute("aria-label", "色");
        const foot = form.appendChild(el("div", "pickfoot"));
        const sample = foot.appendChild(el("span", "picksample"));
        const ok = foot.appendChild(el("button", "ok", "OK"));
        ok.type = "submit";
        // The color as hue (0-360), saturation and value (0-1).
        let [h, sat, val] = hexToHsv(start);
        const hex = () => hsvToHex(h, sat, val);
        const show = (typed) => {
            area.style.background = `linear-gradient(to top, #000, transparent), linear-gradient(to right, #fff, hsl(${h}, 100%, 50%))`;
            dot.style.left = `${sat * 100}%`;
            dot.style.top = `${(1 - val) * 100}%`;
            dot.style.background = hex();
            hueDot.style.left = `${(h / 360) * 100}%`;
            hueDot.style.background = `hsl(${h}, 100%, 50%)`;
            sample.style.background = hex();
            hash.hidden = mode.value !== "hex";
            if (!typed) {
                const c = hex();
                input.value = mode.value === "hex" ? c.slice(1) : [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16)).join(", ");
            }
            input.setCustomValidity("");
        };
        // Dragged (or clicked) on the square or the bar.
        const drag = (target, at) => {
            target.addEventListener("pointerdown", (e) => {
                e.preventDefault();
                target.setPointerCapture(e.pointerId);
                const move = (ev) => {
                    const r = target.getBoundingClientRect();
                    at(Math.min(1, Math.max(0, (ev.clientX - r.left) / r.width)), Math.min(1, Math.max(0, (ev.clientY - r.top) / r.height)));
                    show();
                };
                move(e);
                (target === area ? dot : hueDot).focus();
                const up = () => {
                    target.removeEventListener("pointermove", move);
                    target.removeEventListener("pointerup", up);
                };
                target.addEventListener("pointermove", move);
                target.addEventListener("pointerup", up);
            });
        };
        drag(area, (x, y) => ([sat, val] = [x, 1 - y]));
        drag(hueBar, (x) => (h = Math.min(359.9, x * 360)));
        const keys = (k, step) =>
            k.addEventListener("keydown", (e) => {
                const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, 1], ArrowDown: [0, -1] }[e.key];
                if (!d) return;
                e.preventDefault();
                e.stopPropagation();
                step(d[0] * (e.shiftKey ? 10 : 1), d[1] * (e.shiftKey ? 10 : 1));
                show();
            });
        keys(dot, (dx, dy) => {
            sat = Math.min(1, Math.max(0, sat + dx / 100));
            val = Math.min(1, Math.max(0, val + dy / 100));
        });
        keys(hueDot, (dx) => (h = Math.min(359.9, Math.max(0, h + dx))));
        // Typed: as hex (RRGGBB, RGB) or as R, G, B.
        const typed = () => {
            const t = input.value.trim();
            let c = null;
            const m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec(t);
            if (m) c = m[1].length === 3 ? m[1].replace(/./g, "$&$&") : m[1];
            const n = t.split(/[\s,]+/).filter(Boolean).map(Number);
            if (!c && n.length === 3 && n.every((x) => Number.isInteger(x) && x >= 0 && x <= 255)) c = n.map((x) => x.toString(16).padStart(2, "0")).join("");
            return c && `#${c.toUpperCase()}`;
        };
        input.addEventListener("input", () => {
            const c = typed();
            if (!c) return;
            [h, sat, val] = hexToHsv(c);
            show(true);
        });
        mode.addEventListener("change", () => show());
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            if (document.activeElement === input && !typed()) {
                input.setCustomValidity(mode.value === "hex" ? "RRGGBB の 16 進数で入力してください" : "0 から 255 の数を 3 つ入力してください");
                input.reportValidity();
                return;
            }
            done(hex());
        });
        show();
        return form;
    }
    // A box of a color in a dialog (for <input type="color">, whose picker is the system's): its
    // `value` ("#rrggbb"), `input` and `change` when another is picked in the page's picker.
    function colorInput(start) {
        const b = el("button", "colorinput");
        b.type = "button";
        const sw = b.appendChild(el("span", "sw"));
        let value = "#000000";
        Object.defineProperty(b, "value", {
            get: () => value,
            set: (v) => {
                value = String(v).toLowerCase();
                sw.style.background = value;
                b.title = value.toUpperCase();
            },
        });
        b.value = start || "#000000";
        b.appendChild(el("span", "chev"));
        b.addEventListener("click", () => {
            const pop = openPop(b, "colors pickeronly");
            colorPicker(pop, value, (color) => {
                closeMenu();
                b.value = color;
                b.dispatchEvent(new Event("input", { bubbles: true }));
                b.dispatchEvent(new Event("change", { bubbles: true }));
                b.focus();
            });
        });
        return b;
    }
    function hexToHsv(c) {
        const m = /^#?([0-9a-f]{6})$/i.exec(c || "");
        const [r, g, b] = m ? [0, 2, 4].map((i) => parseInt(m[1].slice(i, i + 2), 16) / 255) : [0, 0, 0];
        const max = Math.max(r, g, b);
        const d = max - Math.min(r, g, b);
        let h = 0;
        if (d) h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
        return [(h * 60 + 360) % 360, max ? d / max : 0, max];
    }
    function hsvToHex(h, s, v) {
        const f = (n) => {
            const k = (n + h / 60) % 6;
            return Math.round((v - v * s * Math.max(0, Math.min(k, 4 - k, 1))) * 255);
        };
        return `#${[f(5), f(3), f(1)].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
    }
    // The borders of the ribbon's menu (as Excel's, in its groups) and the sides each draws: its
    // lines of the 線のスタイル chosen ("s"), thick ("t") or double ("d").
    const BORDERS = [
        [["bottom", "下罫線", { b: "s" }], ["top", "上罫線", { t: "s" }], ["left", "左罫線", { l: "s" }], ["right", "右罫線", { r: "s" }]],
        [["none", "枠なし", {}], ["all", "格子", { t: "s", b: "s", l: "s", r: "s", h: "s", v: "s" }], ["outline", "外枠", { t: "s", b: "s", l: "s", r: "s" }], ["thick", "太い外枠", { t: "t", b: "t", l: "t", r: "t" }]],
        [["doubleBottom", "下二重罫線", { b: "d" }], ["thickBottom", "下太罫線", { b: "t" }], ["topBottom", "上罫線 + 下罫線", { t: "s", b: "s" }], ["topThickBottom", "上罫線 + 下太罫線", { t: "s", b: "t" }], ["topDoubleBottom", "上罫線 + 下二重罫線", { t: "s", b: "d" }]],
        [["diagonalDown", "斜め罫線（＼）", { dd: "s" }], ["diagonalUp", "斜め罫線（／）", { du: "s" }]],
    ];
    // The lines of 線のスタイル (as Excel's list, in the file's names) and how they are drawn:
    // [width, dashes (SVG), double].
    const BORDER_LINES = [
        ["thin", "細線", [1, ""]],
        ["hair", "極細線", [1, "1 1"]],
        ["dotted", "点線", [1, "1 2"]],
        ["dashDotDot", "二点鎖線", [1, "6 2 1 2 1 2"]],
        ["dashDot", "一点鎖線", [1, "6 2 1 2"]],
        ["dashed", "破線", [1, "3 2"]],
        ["double", "二重線", [1, "", true]],
        ["medium", "中線", [2, ""]],
        ["mediumDashed", "中破線", [2, "5 2"]],
        ["mediumDashDot", "中一点鎖線", [2, "7 2 2 2"]],
        ["mediumDashDotDot", "中二点鎖線", [2, "7 2 2 2 2 2"]],
        ["slantDashDot", "斜め一点鎖線", [2, "7 1 3 1"]],
        ["thick", "太線", [3, ""]],
    ];
    // The icon of a border of the menu: the cell's gridlines dotted, the sides it draws on them.
    function borderIcon(sides) {
        const line = (k, x1, y1, x2, y2) => {
            const v = sides[k];
            if (!v) return "";
            if (v === "d") {
                const [dx, dy] = y1 === y2 ? [0, 1.2] : [1.2, 0];
                return `<path d="M${x1 - dx} ${y1 - dy}L${x2 - dx} ${y2 - dy}M${x1 + dx} ${y1 + dy}L${x2 + dx} ${y2 + dy}" stroke-width="1"/>`;
            }
            return `<path d="M${x1} ${y1}L${x2} ${y2}" stroke-width="${v === "t" ? 2.2 : 1.2}"/>`;
        };
        const span = el("span", "ico bico");
        span.innerHTML =
            '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-linecap="square" aria-hidden="true">' +
            '<path d="M2.5 2.5h11v11h-11zM8 2.5v11M2.5 8h11" stroke="#9a9a9a" stroke-width="1" stroke-dasharray="1 1"/>' +
            line("t", 2.5, 2.5, 13.5, 2.5) +
            line("b", 2.5, 13.5, 13.5, 13.5) +
            line("l", 2.5, 2.5, 2.5, 13.5) +
            line("r", 13.5, 2.5, 13.5, 13.5) +
            line("h", 2.5, 8, 13.5, 8) +
            line("v", 8, 2.5, 8, 13.5) +
            line("dd", 2.5, 2.5, 13.5, 13.5) +
            line("du", 2.5, 13.5, 13.5, 2.5) +
            "</svg>";
        return span;
    }
    // A line of 線のスタイル drawn `w` wide, in `color`.
    function lineSample(name, w, color) {
        const [width, dash, double] = (BORDER_LINES.find((l) => l[0] === name) || BORDER_LINES[0])[2];
        const c = color || "currentColor";
        const path = double ? `<path d="M0 6.5h${w}M0 9.5h${w}" stroke-width="1"/>` : `<path d="M0 8h${w}" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ""}/>`;
        return `<svg viewBox="0 0 ${w} 16" width="${w}" height="16" fill="none" stroke="${c}" aria-hidden="true">${path}</svg>`;
    }
    // The formats of the ribbon's list (as Japanese Excel 2021 gives them, in the file's spelling).
    const ACCOUNTING_YEN = '_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* "-"_ ;_ @_ ';
    const NUMBER_FORMATS = [
        ["General", "標準"],
        ["0.00", "数値"],
        ["¥#,##0;[Red]¥-#,##0", "通貨"],
        [ACCOUNTING_YEN, "会計"],
        ["m/d/yyyy", "短い日付形式"],
        ["[$-F800]dddd, mmmm dd, yyyy", "長い日付形式"],
        ["[$-F400]h:mm:ss AM/PM", "時刻"],
        ["0.00%", "パーセンテージ"],
        ["# ?/?", "分数"],
        ["0.00E+00", "指数"],
        ["@", "文字列"],
    ];
    const FONT_SIZES = [6, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
    // The fonts listed until those of the system are known.
    const FONT_NAMES = ["游ゴシック", "游ゴシック Light", "游明朝", "ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ 明朝", "ＭＳ Ｐ明朝", "メイリオ", "Meiryo UI", "BIZ UDゴシック", "BIZ UDPゴシック", "Arial", "Calibri", "Century", "Consolas", "Segoe UI", "Times New Roman", "Verdana"];
    // The fonts installed on the system: VS Code's editor lists them (the page asks it), the
    // playground asks the browser (the Local Font Access API: the browser asks the user the first
    // time the list is opened). Kept with the page's state: the page is drawn again at each edit.
    let systemFonts = Array.isArray(saved.fonts) ? saved.fonts : null;
    let fontsAsked = !!systemFonts;
    let fontsWaiting = null;
    const inVsCode = typeof acquireVsCodeApi === "function";
    function takeFonts(names) {
        const seen = new Set();
        const list = [];
        for (const n of names) {
            const name = String(n || "").trim();
            // (Windows lists the vertical writing of a font as "@" and its name.)
            if (!name || name.startsWith("@") || seen.has(name)) continue;
            seen.add(name);
            list.push(name);
        }
        if (!list.length) return;
        list.sort((a, b) => a.localeCompare(b, "ja"));
        systemFonts = list;
        remember({ fonts: list });
        if (fontsWaiting) fontsWaiting();
    }
    function askFonts() {
        if (fontsAsked) return;
        fontsAsked = true;
        if (inVsCode) vscode.postMessage({ type: "fonts" });
        else if (typeof window.queryLocalFonts === "function") {
            // (Called on the click that opens the list: the browser asks for the permission then.)
            window
                .queryLocalFonts()
                .then((fonts) => takeFonts(fonts.map((f) => f.family)))
                .catch(() => {
                    // Not allowed: the fonts of Office listed.
                });
        }
    }
    // The list of the fonts under the font box: each in its own font; `pick` takes the one chosen.
    function openFontList(anchor, pick, current) {
        askFonts();
        const fill = () => {
            const now = current !== undefined ? current : tools.fontName.value.trim();
            const names = systemFonts || FONT_NAMES;
            pop.textContent = "";
            const list = names.includes(now) || !now ? names : [now].concat(names);
            let on = null;
            for (const name of list) {
                const b = popItem(pop, name, () => pick(name));
                b.dataset.font = name;
                b.style.fontFamily = `"${name.replace(/["\\]/g, "")}", sans-serif`;
                if (name === now) {
                    b.classList.add("on");
                    on = b;
                }
            }
            if (!systemFonts && fontsAsked && !inVsCode && typeof window.queryLocalFonts !== "function") {
                pop.appendChild(el("div", "pnote", "このブラウザーではシステムのフォントを一覧できません（名前を入力すると使えます）"));
            }
            if (on) pop.scrollTop = Math.max(0, on.offsetTop - pop.clientHeight / 2);
        };
        const pop = openPop(anchor, "fonts");
        fontsWaiting = () => {
            if (menu === pop) fill();
        };
        fill();
    }

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
        const now = view && charTarget() ? charFont().size || fontOfCss(view.sheet, activeCss()).pt : view ? fontOfCss(view.sheet, activeCss()).pt : 11;
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
    // ---- ホーム > 書式: the sheets (hidden, shown again, moved or copied, protected) and the
    // cells' lock.

    // "既定の幅": the width of the columns given none, in characters (as Excel counts them).
    function askStandardWidth() {
        if (!editable || busy || !view) return;
        const sheet = view.sheet;
        const px = sheet.defaultWidth || 72;
        const now = px < 13 ? px / 13 : (px - 5) / 8;
        const form = openDialog("標準の幅", "stdwidth");
        const input = field(form, "標準の列幅", "number");
        input.min = "0";
        input.max = "255";
        input.step = "0.01";
        input.value = String(Math.round(now * 100) / 100);
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const width = +input.value;
            if (!(width >= 0 && width <= 255) || input.value.trim() === "") {
                input.setCustomValidity("0 から 255 までの数値を入力してください。");
                input.reportValidity();
                return;
            }
            closeDialog();
            postEdit([{ op: "standardWidth", sheet: sheet.name, width }]);
        });
        input.addEventListener("input", () => input.setCustomValidity(""));
        input.focus();
        input.select();
    }
    // "シートを表示しない": the next visible sheet is shown in its place.
    function hideSheet(sheet) {
        if (!editable || busy) return;
        const visible = data.sheets.filter((s) => s.visible === "visible");
        if (sheet.visible !== "visible") return;
        if (visible.length <= 1) {
            showError("ブックのすべてのシートを非表示にすることはできません。");
            return;
        }
        const i = visible.indexOf(sheet);
        const next = visible[i + 1] || visible[i - 1];
        remember({ sheet: next.name });
        postEdit([{ op: "sheetVisible", sheet: sheet.name, visible: false }]);
    }
    // "シートの再表示": the hidden sheets, any of them shown again.
    function askUnhideSheets() {
        if (!editable || busy) return;
        const hidden = data.sheets.filter((s) => s.visible !== "visible");
        if (!hidden.length) return;
        const form = openDialog("再表示", "unhidedlg");
        form.appendChild(el("div", "lbl", "表示するシートを選択してください:"));
        const list = form.appendChild(el("select"));
        list.multiple = true;
        list.size = Math.min(8, Math.max(4, hidden.length));
        for (const s of hidden) list.appendChild(el("option", null, s.name)).value = s.name;
        list.options[0].selected = true;
        list.addEventListener("dblclick", () => form.requestSubmit());
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const names = [...list.selectedOptions].map((o) => o.value);
            closeDialog();
            if (!names.length) return;
            remember({ sheet: names[0] });
            postEdit(names.map((name) => ({ op: "sheetVisible", sheet: name, visible: true })));
        });
        list.focus();
    }
    // "シートの移動またはコピー": before which sheet (or at the end), moved or a copy.
    function askMoveOrCopy(sheet) {
        if (!editable || busy) return;
        if (entry) commit(null);
        const form = openDialog("シートの移動またはコピー", "movedlg");
        form.appendChild(el("div", "lbl", `選択したシート「${sheet.name}」を移動します。`));
        form.appendChild(el("div", "lbl", "挿入先:"));
        const list = form.appendChild(el("select"));
        list.size = Math.min(10, data.sheets.length + 1);
        for (const s of data.sheets) list.appendChild(el("option", null, s.visible === "visible" ? s.name : `${s.name}（非表示）`)).value = s.name;
        const end = list.appendChild(el("option", null, "(末尾へ移動)"));
        end.value = "";
        // (As Excel: the sheet after it.)
        const at = data.sheets.indexOf(sheet);
        list.selectedIndex = Math.min(at + 1, data.sheets.length);
        const copy = check(form, "コピーを作成する", false);
        const sync = () => (form.querySelector(".lbl").textContent = `選択したシート「${sheet.name}」を${copy.checked ? "コピー" : "移動"}します。`);
        copy.addEventListener("change", sync);
        list.addEventListener("dblclick", () => form.requestSubmit());
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const before = list.value;
            closeDialog();
            if (copy.checked) {
                let n = 2;
                const base = (k) => `${sheet.name.slice(0, 31 - ` (${k})`.length)} (${k})`;
                while (takenName(base(n))) n++;
                const place = before ? { before } : { after: data.sheets[data.sheets.length - 1].name };
                postSheet({ action: "copy", sheet: sheet.name, ...place }, base(n));
            } else {
                if (before === sheet.name || (before === "" && at === data.sheets.length - 1) || (before && data.sheets[at + 1] && data.sheets[at + 1].name === before)) return;
                remember({ sheet: sheet.name });
                postEdit([{ op: "moveSheet", sheet: sheet.name, before }]);
            }
        });
        list.focus();
    }
    // What a protected sheet still allows (the check boxes of "シートの保護"), in Excel's order.
    const PROTECT_ALLOWS = [
        ["selectLockedCells", "ロックされたセル範囲の選択", true],
        ["selectUnlockedCells", "ロックされていないセル範囲の選択", true],
        ["formatCells", "セルの書式設定"],
        ["formatColumns", "列の書式設定"],
        ["formatRows", "行の書式設定"],
        ["insertColumns", "列の挿入"],
        ["insertRows", "行の挿入"],
        ["insertHyperlinks", "ハイパーリンクの挿入"],
        ["deleteColumns", "列の削除"],
        ["deleteRows", "行の削除"],
        ["sort", "並べ替え"],
        ["autoFilter", "オートフィルターの使用"],
        ["pivotTables", "ピボットテーブルとピボットグラフを使う"],
        ["objects", "オブジェクトの編集"],
        ["scenarios", "シナリオの編集"],
    ];
    // "シートの保護": a password (asked again), and what the sheet still allows.
    function askProtect(sheet) {
        if (!editable || busy) return;
        if (entry) commit(null);
        const form = openDialog("シートの保護", "protectdlg");
        const on = check(form, "シートとロックされたセルの内容を保護する", true);
        const password = field(form, "シートの保護を解除するためのパスワード", "password");
        password.autocomplete = "new-password";
        form.appendChild(el("div", "lbl", "このシートのすべてのユーザーに以下を許可します。"));
        const box = form.appendChild(el("div", "allows"));
        const boxes = PROTECT_ALLOWS.map(([key, label, checked]) => [key, check(box, label, !!checked)]);
        const ok = dialogButtons(form).querySelector(".ok");
        on.addEventListener("change", () => (ok.disabled = !on.checked));
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const allow = boxes.filter(([, b]) => b.checked).map(([key]) => key);
            const op = { op: "protect", sheet: sheet.name, on: true, password: password.value, allow };
            if (!password.value) {
                closeDialog();
                postEdit([op]);
                return;
            }
            // "パスワードの確認".
            const again = openDialog("パスワードの確認", "protectdlg");
            const input = field(again, "続行するにはもう一度パスワードを入力してください。", "password");
            again.appendChild(el("div", "note", "注意: 忘れてしまったパスワードは回復できません。パスワードとそれに対応するブック名とシート名を一覧にして、安全な場所に保管することをお勧めします (パスワードについては大文字と小文字が区別されることに注意してください)。"));
            dialogButtons(again);
            again.addEventListener("submit", (e2) => {
                e2.preventDefault();
                closeDialog();
                if (input.value !== password.value) {
                    showError("確認のパスワードが一致しません。");
                    return;
                }
                postEdit([op]);
            });
            input.focus();
        });
        password.focus();
    }
    // "シート保護の解除": with its password, when it has one.
    function askUnprotect(sheet) {
        if (!editable || busy) return;
        if (entry) commit(null);
        const op = { op: "protect", sheet: sheet.name, on: false, password: "" };
        if (!sheet.protection.password) {
            postEdit([op]);
            return;
        }
        const form = openDialog("シート保護の解除", "protectdlg");
        const input = field(form, "パスワード", "password");
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
            postEdit([{ ...op, password: input.value }]);
        });
        input.focus();
    }
    // The cell at (r, c) of the sheet shown is locked (as every cell is, until セルのロック is
    // taken off it).
    const lockedAt = (r, c) => !/--unlocked:1/.test(cssOf(view.sheet, r, c, view.cells.get(r * KEY + c)));
    const cellLocked = () => !!(view && sel && lockedAt(sel.ar, sel.ac));
    // "セルのロック": the selection locked, or no longer (as the active cell is now).
    function toggleLock() {
        if (!sel || !view) return;
        if (view.sheet.protection) {
            showError(PROTECTED);
            return;
        }
        format({ locked: !cellLocked() });
    }
    // "ロックされたセルを表示": the locked cells shaded, or not (kept with the page's state).
    let showLocked = !!saved.showLocked;
    if (showLocked) $("grid").classList.add("showlocked");
    function showLockedCells(on) {
        showLocked = on;
        remember({ showLocked: on });
        $("grid").classList.toggle("showlocked", on);
    }
    // A protected sheet: what an edit would change that it does not allow (Excel's message).
    const PROTECTED = "変更しようとしているセルやグラフは保護されているシート上にあります。変更するには、シートの保護を解除してください。パスワードの入力が必要な場合もあります。";
    // Every cell of the range `text` (as the edits write it; `size` rows by columns from its
    // start when given) is unlocked.
    function unlockedRange(text, size) {
        const sheet = view.sheet;
        const a = parseRef(text, sheet);
        if (!a) return false;
        let { r1, c1, r2, c2 } = a;
        if (size) {
            r2 = Math.max(r2, r1 + size[0] - 1);
            c2 = Math.max(c2, c1 + size[1] - 1);
        }
        // (Past the cells: their rows' and columns' formats; one beyond stands for the rest.)
        r2 = Math.min(r2, (sheet.rows || 0) + 1);
        c2 = Math.min(c2, (sheet.cols || 0) + 1);
        if ((r2 - r1 + 1) * (c2 - c1 + 1) > 500000) return false;
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (lockedAt(r, c)) return false;
        return true;
    }
    // The edits `ops` are refused on the protected sheet shown: what they change is locked, or
    // the protection does not allow it.
    function refusedByProtection(ops) {
        if (!view || !view.sheet.protection) return false;
        const sheet = view.sheet;
        const allow = sheet.protection.allow || {};
        const sizeOf = (text) => {
            const a = text && parseRef(text, sheet);
            return a ? [a.r2 - a.r1 + 1, a.c2 - a.c1 + 1] : null;
        };
        const columns = (text) => /^\$?[A-Za-z]/.test(String(text || ""));
        return ops.some((o) => {
            if (o.sheet && o.sheet !== sheet.name) return false;
            switch (o.op) {
                case "protect":
                case "sheetVisible":
                case "moveSheet":
                case "tabColor":
                case "name":
                    return false;
                case "set":
                case "clear":
                    return !unlockedRange(o.range);
                case "values":
                    return !unlockedRange(o.at, [o.rows.length, Math.max(1, ...o.rows.map((row) => row.length))]);
                case "copy":
                case "pasteSpecial":
                case "autofill":
                    return !unlockedRange(o.to, o.op === "autofill" ? null : sizeOf(o.from));
                case "move":
                    return !unlockedRange(o.from) || !unlockedRange(o.to, sizeOf(o.from));
                case "fill":
                    return !unlockedRange(o.range);
                case "link":
                    return !allow.insertHyperlinks || !unlockedRange(o.range);
                case "format":
                case "clearFormats":
                case "merge":
                case "unmerge":
                case "conditional":
                case "conditionalRules":
                    return !allow.formatCells;
                case "validate":
                    return false;
                case "size":
                case "hide":
                    return !(columns(o.range) ? allow.formatColumns : allow.formatRows);
                case "fit":
                    return !allow.formatRows;
                case "standardWidth":
                    return !allow.formatColumns;
                case "filterBy":
                    return !allow.autoFilter;
                case "sort":
                    return !allow.sort || !unlockedRange(o.range);
                // An object: as its ロック and テキストをロックする say, unless the sheet lets
                // objects be edited.
                case "object": {
                    if (allow.objects) return false;
                    const obj = (sheet.objects || []).find((x) => (o.name ? x.name === o.name : x.shapeId === o.shape));
                    const props = (obj && obj.props) || {};
                    if (props.locked === false) return false;
                    const keys = Object.keys(o).filter((k) => !["op", "sheet", "name", "shape"].includes(k));
                    return !(props.lockText === false && keys.length && keys.every((k) => k === "text"));
                }
                case "shape":
                case "picture":
                case "control":
                case "activeX":
                    return !allow.objects;
                default:
                    return true;
            }
        });
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
        validation: '<rect x="1.5" y="2.5" width="9" height="11"/><path d="M1.5 6h9M1.5 9.5h9M5 2.5v11"/><path d="m9.5 11 2 2 3.5-5" stroke="#107c41" stroke-width="1.6"/>',
        formulas: '<text x="1" y="12" font-size="10" font-style="italic" stroke="none" fill="currentColor" font-family="serif">fx</text><path d="M11 5h4M11 9h4M11 13h4"/>',
        nameManager: '<rect x="1.5" y="2.5" width="13" height="11"/><path d="M1.5 5.5h13M5.5 5.5v8M1.5 8.2h13M1.5 10.9h13"/><path d="M7.5 3.8h5" stroke-width="1.2"/>',
        defineName: '<path d="M1.5 4.5h8l3.5 3.5-3.5 3.5h-8z"/><path d="M3.8 8h4"/>',
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
        selectionPane: '<rect x="1.5" y="1.5" width="13" height="13"/><path d="M9.5 1.5v13M11 4.5h2M11 7h2M11 9.5h2"/><path d="M3.5 4.5 7 8.2l-1.6.3.9 2-1 .4-.9-2L3.5 10z" fill="currentColor" stroke="none"/>',
        align: '<path d="M2.5 1.5v13"/><rect x="2.5" y="3.5" width="9" height="3"/><rect x="2.5" y="9.5" width="6" height="3"/>',
        alignLeft: '<path d="M2.5 1.5v13"/><rect x="4.5" y="3.5" width="9" height="3"/><rect x="4.5" y="9.5" width="6" height="3"/>',
        alignCenter: '<path d="M8 1.5v2M8 6.5v3M8 12.5v2"/><rect x="3" y="3.5" width="10" height="3"/><rect x="5" y="9.5" width="6" height="3"/>',
        alignRight: '<path d="M13.5 1.5v13"/><rect x="2.5" y="3.5" width="9" height="3"/><rect x="5.5" y="9.5" width="6" height="3"/>',
        alignTop: '<path d="M1.5 2.5h13"/><rect x="3.5" y="4.5" width="3" height="9"/><rect x="9.5" y="4.5" width="3" height="6"/>',
        alignMiddle: '<path d="M1.5 8h2M6.5 8h3M12.5 8h2"/><rect x="3.5" y="3" width="3" height="10"/><rect x="9.5" y="5" width="3" height="6"/>',
        alignBottom: '<path d="M1.5 13.5h13"/><rect x="3.5" y="2.5" width="3" height="9"/><rect x="9.5" y="5.5" width="3" height="6"/>',
        spreadH: '<path d="M1.5 1.5v13M14.5 1.5v13"/><rect x="6.5" y="4.5" width="3" height="7"/>',
        spreadV: '<path d="M1.5 1.5h13M1.5 14.5h13"/><rect x="4.5" y="6.5" width="7" height="3"/>',
        snapGrid: '<path d="M1.5 5.5h13M1.5 10.5h13M5.5 1.5v13M10.5 1.5v13" stroke-opacity=".45"/><rect x="5.5" y="5.5" width="5" height="5" fill="#dbe5f1"/>',
        height: '<path d="M8 1.5v13M6 3.5l2-2 2 2M6 12.5l2 2 2-2M3 1.5h10M3 14.5h10"/>',
        width: '<path d="M1.5 8h13M3.5 6l-2 2 2 2M12.5 6l2 2-2 2M1.5 3v10M14.5 3v10"/>',
        condFormat: '<rect x="1.5" y="1.5" width="13" height="13"/><path d="M1.5 5.8h13M1.5 10.2h13M5.5 1.5v13"/><path d="M7 2.9h6.2v1.6H7zM7 7.2h3v1.6H7zM7 11.5h4.6v1.6H7z" fill="currentColor" stroke="none"/>',
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
    // The colors of 図形の塗りつぶし and 図形の枠線 used last.
    const lastShape = { fill: "#4472C4", line: "#000000", shadow: "#000000", glow: "#4472C4" };
    let lastBorder = "all";
    // The 線の色 (null: automatic) and the 線のスタイル the borders of the menu are drawn with.
    const borderPen = { color: null, line: "thin", ...(saved.borderPen || {}) };
    function borderSet(preset) {
        return { border: borderPen.color || borderPen.line !== "thin" ? { preset, line: borderPen.line, color: borderPen.color } : preset };
    }
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
        // A tab wider than the window scrolls: a fade marks each side with buttons out of view,
        // and the wheel scrolls it sideways (its scroll bar is a thin one shown under the mouse).
        const edges = () => {
            const max = body.scrollWidth - body.clientWidth;
            body.classList.toggle("more-l", max > 1 && body.scrollLeft > 1);
            body.classList.toggle("more-r", max > 1 && body.scrollLeft < max - 1);
        };
        body.addEventListener("scroll", edges, { passive: true });
        // `hover` (overflow-x: scroll) repaints the bar; set only while the tab overflows, so nothing moves.
        body.addEventListener("mouseenter", () => body.classList.toggle("hover", body.scrollWidth > body.clientWidth));
        body.addEventListener("mouseleave", () => body.classList.remove("hover"));
        if (typeof ResizeObserver === "function") new ResizeObserver(edges).observe(body);
        body.addEventListener(
            "wheel",
            (e) => {
                if (e.ctrlKey || Math.abs(e.deltaX) >= Math.abs(e.deltaY) || body.scrollWidth <= body.clientWidth) return;
                if (e.target.closest("select")) return;
                e.preventDefault();
                body.scrollLeft += e.deltaMode === 1 ? e.deltaY * 16 : e.deltaY;
            },
            { passive: false },
        );
        const show = (key, keep) => {
            shown = key;
            for (const [k, p] of Object.entries(panels)) {
                p.panel.hidden = k !== key;
                p.tab.classList.toggle("on", k === key);
                p.tab.setAttribute("aria-selected", String(k === key));
            }
            if (!keep) remember({ ribbonTab: key });
            edges();
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
            if (entry && !keep) return;
            act(b);
            if (!menu) closeOpen();
            focusGrid();
        };
        // `o`: { label, title, icon, large, edits, menu: (anchor) => fill it, split, cls, text,
        // keepEntry (what is being typed in the cell stays) }. A large button's label is on the
        // lines `\n` breaks it into, as Excel's ("条件付き\n書式"), its ▾ after the last.
        const button = (key, o) => {
            const wrap = o.split ? into.appendChild(el("div", `rsplit${o.large ? " large" : ""}`)) : into;
            const b = wrap.appendChild(el("button", `rbtn${o.large ? " large" : ""}${o.cls ? ` ${o.cls}` : ""}`));
            b.type = "button";
            b.title = o.title || (o.label || "").replace(/\n/g, "");
            b.dataset.tool = key;
            if (o.icon) b.appendChild(icon(o.icon, o.large));
            if (o.text) b.appendChild(el("span", "glyph", o.text));
            const lbl = o.label && (o.large || o.showLabel) ? b.appendChild(el("span", "lbl", o.large ? o.label : o.label.replace(/\n/g, ""))) : null;
            if (o.menu && !o.split) (o.large && lbl ? lbl : b).appendChild(el("span", "caret chev"));
            if (o.edits && !editable) b.disabled = true;
            const openMenuOf = (anchor) => {
                if (o.edits && busy) return;
                const keep = o.chars && charTarget();
                if (entry && !keep) commit(null);
                if (entry && !keep) return;
                o.menu(anchor);
            };
            if (o.menu && !o.split) b.addEventListener("click", () => openMenuOf(wrap === into ? b : wrap));
            else b.addEventListener("click", () => run(b, o.act, o.edits, o.keepEntry || (o.chars && charTarget())));
            if (o.split) {
                const more = wrap.appendChild(el("button", "rbtn more"));
                more.appendChild(el("span", "chev"));
                more.type = "button";
                more.title = o.moreTitle || o.title || (o.label || "").replace(/\n/g, "");
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
        // The small button at the corner of a group's label (Excel's dialog box launcher).
        const launcher = (title, act) => {
            const label = into.closest(".rgroup").querySelector(".rlabel");
            const b = label.appendChild(el("button", "launch"));
            b.type = "button";
            b.title = title;
            b.innerHTML = '<svg viewBox="0 0 8 8" width="8" height="8" fill="none" stroke="currentColor" stroke-width="1" aria-hidden="true"><path d="M.5 3.5V.5h3M7.5 3.5v4h-4M3 3l4.2 4.2M7.5 5v2.5H5"/></svg>';
            b.addEventListener("click", () => {
                if (entry) commit(null);
                act();
            });
            return b;
        };
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
        group(home, "履歴");
        stack();
        button("undo", { icon: "undo", label: "元に戻す", title: "元に戻す（Ctrl+Z）", edits: true, act: () => postHistory("undo") });
        button("redo", { icon: "redo", label: "やり直し", title: "やり直し（Ctrl+Y）", edits: true, act: () => postHistory("redo") });

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
        button("cut", { icon: "cut", label: "切り取り", title: "切り取り（Ctrl+X）", edits: true, act: () => menuCopy(true) });
        button("copy", { icon: "copy", label: "コピー", title: "コピー（Ctrl+C）", act: () => menuCopy(false) });

        group(home, "フォント");
        const fontRows = into.appendChild(el("div", "rstack wide"));
        row(fontRows);
        launcher("フォント: 右にセルの書式設定（フォント）を表示します", () => openCellPane("font"));
        // A box typed into and a list of every font (a datalist would show only the fonts
        // matching what the box has).
        const fontBox = into.appendChild(el("span", "fontbox"));
        const fontName = fontBox.appendChild(el("input", "fontname"));
        fontName.type = "text";
        fontName.title = "フォント";
        fontName.dataset.tool = "fontName";
        fontName.spellcheck = false;
        fontName.autocomplete = "off";
        const fontDrop = fontBox.appendChild(el("button", "fontdrop"));
        fontDrop.appendChild(el("span", "chev"));
        fontDrop.type = "button";
        fontDrop.title = "フォントの一覧";
        fontDrop.dataset.tool = "fontNameMore";
        if (!editable) fontName.disabled = fontDrop.disabled = true;
        tools.fontName = fontName;
        const applyFont = () => {
            const name = fontName.value.trim();
            if (name && editable && sel) format({ fontName: name });
            focusGrid();
        };
        fontName.addEventListener("change", applyFont);
        const openFonts = () => {
            if (busy) return;
            if (entry && !charTarget()) commit(null);
            openFontList(fontBox, (name) => {
                fontName.value = name;
                applyFont();
            });
        };
        fontDrop.addEventListener("click", openFonts);
        fontName.addEventListener("keydown", (e) => {
            if (e.key === "Escape") {
                e.preventDefault();
                showTools();
                focusGrid();
            } else if (e.key === "ArrowDown" && e.altKey) {
                e.preventDefault();
                openFonts();
            }
        });
        const size = select("fontSize", "フォント サイズ", FONT_SIZES.map((n) => [String(n), String(n)]), "fontsize");
        size.addEventListener("change", () => {
            format({ fontSize: +size.value });
            focusGrid();
        });
        button("fontGrow", { icon: "grow", title: "フォント サイズの拡大", edits: true, chars: true, act: () => stepFontSize(true) });
        button("fontShrink", { icon: "shrink", title: "フォント サイズの縮小", edits: true, chars: true, act: () => stepFontSize(false) });
        row(fontRows);
        button("bold", { text: "B", title: "太字（Ctrl+B）", cls: "b", edits: true, chars: true, act: () => toggleFormat("bold") });
        button("italic", { text: "I", title: "斜体（Ctrl+I）", cls: "i", edits: true, chars: true, act: () => toggleFormat("italic") });
        button("underline", { text: "U", title: "下線（Ctrl+U）", cls: "u", edits: true, chars: true, act: () => toggleFormat("underline") });
        sep();
        // The button shows the border used last (格子 at first) and gives it again, as Excel's.
        const showBorder = (value) => {
            const [, label, sides] = BORDERS.flat().find((b) => b[0] === value);
            const b = tools.border;
            b.title = `罫線（${label}）`;
            const old = b.querySelector(".ico");
            if (old) old.replaceWith(borderIcon(sides));
            else b.insertBefore(borderIcon(sides), b.firstChild);
        };
        button("border", {
            title: "罫線",
            moreTitle: "罫線の種類",
            split: true,
            edits: true,
            act: () => format(borderSet(lastBorder)),
            menu: (anchor) => {
                const pop = openPop(anchor, "borders");
                head(pop, "罫線");
                BORDERS.forEach((group, i) => {
                    if (i) pop.appendChild(el("hr"));
                    for (const [value, label, sides] of group) {
                        const b = item(pop, label, () => {
                            lastBorder = value;
                            showBorder(value);
                            format(borderSet(value));
                        });
                        b.insertBefore(borderIcon(sides), b.firstChild);
                        b.dataset.border = value;
                    }
                });
            },
        });
        showBorder(lastBorder);
        // 線の色 and 線のスタイル: what the borders of the menu are drawn with (Excel has them in
        // the menu).
        const penColor = button("borderColor", {
            icon: "shapeLine",
            title: "線の色",
            cls: "color",
            edits: true,
            menu: (anchor) =>
                openColors(anchor, "borderColor", "自動", (color) => {
                    setPen({ color });
                    renderPane(true);
                }),
        });
        penColor.querySelector(".ico").appendChild(el("span", "bar")).style.background = borderPen.color || "#000000";
        penColor.title = `線の色（${borderPen.color || "自動"}）`;
        const penLine = button("borderLine", {
            title: "線のスタイル（細線）",
            cls: "linestyle",
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "linestyles");
                head(pop, "線のスタイル");
                for (const [name, label] of BORDER_LINES) {
                    const b = item(pop, "", () => {
                        setPen({ line: name });
                        renderPane(true);
                    });
                    b.title = label;
                    b.dataset.line = name;
                    b.classList.toggle("on", borderPen.line === name);
                    b.querySelector("span").innerHTML = lineSample(name, 96);
                }
            },
        });
        penLine.insertBefore(el("span", "sample"), penLine.firstChild).innerHTML = lineSample(borderPen.line, 28);
        penLine.title = `線のスタイル（${(BORDER_LINES.find((l) => l[0] === borderPen.line) || BORDER_LINES[0])[1]}）`;
        sep();
        const colorButton = (key, iconName, title, none) => {
            const b = button(key, {
                icon: iconName,
                title,
                moreTitle: `${title}の選択`,
                split: true,
                edits: true,
                cls: "color",
                chars: key === "color",
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
        launcher("配置の設定: 右にセルの書式設定（配置）を表示します", () => openCellPane("hAlign"));
        for (const [value, ic, title] of [["top", "top", "上揃え"], ["middle", "middle", "上下中央揃え"], ["bottom", "bottom", "下揃え"]]) button(`valign-${value}`, { icon: ic, title, edits: true, act: () => format({ valign: value }) });
        sep();
        button("wrap", { icon: "wrap", label: "折り返して全体を表示する", title: "折り返して全体を表示する", edits: true, act: () => toggleFormat("wrap") });
        row(alignRows);
        for (const [value, ic, title] of [["left", "left", "左揃え"], ["center", "center", "中央揃え"], ["right", "right", "右揃え"]]) button(`align-${value}`, { icon: ic, title, edits: true, act: () => format({ align: value }) });
        sep();
        button("merge", {
            icon: "merge",
            label: "セルを結合して中央揃え",
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
        launcher("表示形式: 右にセルの書式設定（表示形式）を表示します (Ctrl+1)", () => openCellPane("number"));
        const number = select("numberFormat", "表示形式", [["", "ユーザー定義"]].concat(NUMBER_FORMATS, [["more", "その他の表示形式…"]]), "numberformat");
        number.addEventListener("change", () => {
            if (number.value === "more") {
                showTools();
                openCellPane("number");
                return;
            }
            if (number.value) format({ numberFormat: number.value });
            focusGrid();
        });
        row(numberRows);
        button("currency", { icon: "currency", title: "通貨表示形式", edits: true, act: () => format({ numberFormat: "¥#,##0" }) });
        button("percent", { icon: "percent", title: "パーセント スタイル", edits: true, act: () => format({ numberFormat: "0%" }) });
        button("comma", { icon: "comma", title: "桁区切りスタイル", edits: true, act: () => format({ numberFormat: "#,##0" }) });
        sep();
        const decimals = (up) => {
            if (!sel) return;
            const code = stepDecimals(numberCodeAt(sel.ar, sel.ac)[0] || "General", up);
            if (code) format({ numberFormat: code });
        };
        button("decUp", { icon: "decUp", title: "小数点以下の表示桁数を増やす", edits: true, act: () => decimals(true) });
        button("decDown", { icon: "decDown", title: "小数点以下の表示桁数を減らす", edits: true, act: () => decimals(false) });

        group(home, "スタイル");
        button("condFormat", {
            icon: "condFormat",
            label: "条件付き\n書式",
            large: true,
            cls: "two",
            edits: true,
            menu: (anchor) => {
                // As Excel's: the kinds of rules, each a submenu.
                const pop = openPop(anchor, "condmenu");
                const ask = (sub, label, o) => item(sub, label, () => askCond(Object.assign({ title: label.replace(/…$/, "") }, o)));
                // "その他のルール…": "新しい書式ルール", the kind of the submenu chosen.
                const more = (sub, o) => {
                    sub.appendChild(el("hr"));
                    item(sub, "その他のルール…", () => askCond(o));
                };
                const swatch = (where, title, html, act) => {
                    const b = where.appendChild(el("button", "condswatch"));
                    b.type = "button";
                    b.title = title;
                    b.innerHTML = html;
                    b.addEventListener("click", () => {
                        closeMenu();
                        closeOpen();
                        focusGrid();
                        act();
                    });
                    return b;
                };
                const grid = (sub, title, columns) => {
                    if (title) sub.appendChild(el("div", "subhead", title));
                    const g = sub.appendChild(el("div", "condgrid"));
                    g.style.gridTemplateColumns = `repeat(${columns}, auto)`;
                    return g;
                };
                subMenu(pop, "セルの強調表示ルール", (sub) => {
                    ask(sub, "指定の値より大きい…", { kind: "cellIs", operator: "greaterThan", prompt: "次の値より大きいセルを書式設定:" });
                    ask(sub, "指定の値より小さい…", { kind: "cellIs", operator: "lessThan", prompt: "次の値より小さいセルを書式設定:" });
                    ask(sub, "指定の範囲内…", { kind: "cellIs", operator: "between", prompt: "次の値の間にあるセルを書式設定:" });
                    ask(sub, "指定の値に等しい…", { kind: "cellIs", operator: "equal", prompt: "次の値に等しいセルを書式設定:" });
                    ask(sub, "文字列…", { kind: "text", prompt: "次の文字列を含むセルを書式設定:" });
                    ask(sub, "日付…", { kind: "timePeriod", prompt: "次の日付を含むセルを書式設定:" });
                    ask(sub, "重複する値…", { kind: "duplicateValues", prompt: "次の値を含むセルを書式設定:" });
                    more(sub, { startKind: "contains" });
                });
                subMenu(pop, "上位/下位ルール", (sub) => {
                    ask(sub, "上位 10 項目…", { kind: "top10", prompt: "上位に入るセルを書式設定（項目数）:" });
                    ask(sub, "上位 10%…", { kind: "top10", percent: true, prompt: "上位に入るセルを書式設定（%）:" });
                    ask(sub, "下位 10 項目…", { kind: "top10", bottom: true, prompt: "下位に入るセルを書式設定（項目数）:" });
                    ask(sub, "下位 10%…", { kind: "top10", bottom: true, percent: true, prompt: "下位に入るセルを書式設定（%）:" });
                    ask(sub, "平均より上…", { kind: "aboveAverage", prompt: "選択範囲の平均値より上のセルを書式設定:" });
                    ask(sub, "平均より下…", { kind: "aboveAverage", below: true, prompt: "選択範囲の平均値より下のセルを書式設定:" });
                    more(sub, { startKind: "top10" });
                });
                pop.appendChild(el("hr"));
                subMenu(pop, "データ バー", (sub) => {
                    let g = grid(sub, "塗りつぶし (グラデーション)", 3);
                    for (const [color, name] of COND_BARS) swatch(g, `${name}のデータ バー`, `<span class="bar" style="background:linear-gradient(to right, ${color}, #fff);border-color:${color}"></span>`, () => postCond({ type: "dataBar", colors: [color] })).dataset.bar = color;
                    sub.appendChild(el("hr"));
                    g = grid(sub, "塗りつぶし (単色)", 3);
                    for (const [color, name] of COND_BARS) swatch(g, `${name}のデータ バー (単色)`, `<span class="bar" style="background:${color}"></span>`, () => postCond({ type: "dataBar", colors: [color], solid: true })).dataset.solid = color;
                    more(sub, { startKind: "scale", startStyle: "dataBar" });
                });
                subMenu(pop, "カラー スケール", (sub) => {
                    const g = grid(sub, null, 4);
                    for (const [lo, mid, hi, name] of COND_SCALES) {
                        const colors = mid ? [lo, mid, hi] : [lo, hi];
                        // (The menu shows the greatest first, the scale is written from the least.)
                        swatch(g, `${name}のカラー スケール`, `<span class="scale" style="background:linear-gradient(to bottom, ${colors.join(", ")})"></span>`, () => postCond({ type: "colorScale", colors: colors.slice().reverse() })).dataset.scale = colors.join(" ");
                    }
                    more(sub, { startKind: "scale", startStyle: "colorScale2" });
                });
                subMenu(pop, "アイコン セット", (sub) => {
                    // (As Excel's, in its groups, the greatest first.)
                    for (const [title, sets] of COND_ICON_MENU) {
                        const g = grid(sub, title, 2);
                        for (const set of sets) swatch(g, iconName(set), iconsHtml(set, true), () => postCond({ type: "iconSet", icons: set })).dataset.icons = set;
                    }
                    more(sub, { startKind: "scale", startStyle: "iconSet" });
                });
                pop.appendChild(el("hr"));
                item(pop, "新しいルール…", () => askCond());
                subMenu(pop, "ルールのクリア", (sub) => {
                    item(sub, "選択したセルからルールをクリア", () => clearCond(false));
                    item(sub, "シート全体からルールをクリア", () => clearCond(true));
                });
                item(pop, "ルールの管理…", () => manageCond());
            },
        });

        group(home, "セル");
        stack();
        button("insert", {
            icon: "insert",
            label: "挿入",
            showLabel: true,
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
            showLabel: true,
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
            showLabel: true,
            edits: true,
            menu: (anchor) => {
                // (As Excel's ホーム > 書式.)
                const pop = openPop(anchor, "cellformats");
                const sheet = view && view.sheet;
                head(pop, "セルのサイズ");
                item(pop, "行の高さ…", () => askSize(false));
                item(pop, "行の高さの自動調整", () => autofitSel(false));
                item(pop, "列の幅…", () => askSize(true));
                item(pop, "列の幅の自動調整", () => autofitSel(true));
                item(pop, "既定の幅…", askStandardWidth);
                head(pop, "表示設定");
                subMenu(pop, "非表示/再表示", (sub) => {
                    item(sub, "行を表示しない", () => hideWhole(false, true));
                    item(sub, "列を表示しない", () => hideWhole(true, true));
                    item(sub, "シートを表示しない", () => sheet && hideSheet(sheet));
                    sub.appendChild(el("hr"));
                    item(sub, "行の再表示", () => hideWhole(false, false));
                    item(sub, "列の再表示", () => hideWhole(true, false));
                    item(sub, "シートの再表示…", askUnhideSheets).disabled = !data.sheets.some((s) => s.visible !== "visible");
                });
                head(pop, "シートの整理");
                item(pop, "シート名の変更", () => {
                    const b = [...$("tabs").children].find((t) => t.sheet === sheet);
                    if (b) renameTab(b);
                });
                item(pop, "シートの移動またはコピー…", () => sheet && askMoveOrCopy(sheet));
                subMenu(pop, "シート見出しの色", (sub) => {
                    sub.classList.add("colors");
                    fillColors(sub, "tab", "色なし", (color) => sheet && postEdit([{ op: "tabColor", sheet: sheet.name, color }]));
                });
                head(pop, "保護");
                if (sheet && sheet.protection) item(pop, "シート保護の解除…", () => askUnprotect(sheet));
                else item(pop, "シートの保護…", () => sheet && askProtect(sheet));
                const lock = item(pop, "セルのロック", toggleLock);
                lock.classList.add("chk");
                lock.classList.toggle("on", cellLocked());
                lock.disabled = !!(sheet && sheet.protection);
                const show = item(pop, "ロックされたセルを表示", () => showLockedCells(!showLocked));
                show.classList.add("chk");
                show.classList.toggle("on", showLocked);
                item(pop, "セルの書式設定…", () => openCellPane("number"), "Ctrl+1");
            },
        });

        group(home, "編集");
        stack();
        button("fillDown", {
            icon: "fillDown",
            label: "フィル",
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
        stack();
        button("sortFilter", {
            icon: "sortFilter",
            label: "並べ替えとフィルター",
            edits: true,
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
        // The controls of a sheet (Excel's 開発 > 挿入): their events are in the sheet's file.
        group(insertTab, "コントロール");
        button("controlInsert", {
            icon: "controls",
            label: "コントロール",
            large: true,
            cls: "two",
            edits: true,
            menu: (anchor) => {
                const pop = openPop(anchor, "shapes");
                // A row of buttons a group: [heading, controls, what is drawn, what the tip adds].
                const groups = [
                    ["フォーム コントロール", FORM_CONTROLS, "control", "フォーム コントロール"],
                    ["ActiveX コントロール", ACTIVEX_CONTROLS, "activeX", "ActiveX コントロール。イベントはシートのファイルの @on で受け取ります"],
                ];
                for (const [heading, list, how, tip] of groups) {
                    head(pop, heading);
                    const row = pop.appendChild(el("div", "shaperow"));
                    for (const [kind, name, path] of list) {
                        const b = row.appendChild(el("button", "shapebtn"));
                        b.type = "button";
                        b.title = `${name}（${tip}）`;
                        b.dataset.control = `${how}:${kind}`;
                        b.innerHTML = `<svg viewBox="0 0 16 16" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.1" stroke-linejoin="round" aria-hidden="true">${path}</svg>`;
                        b.addEventListener("click", () => {
                            closeMenu();
                            closeOpen();
                            startDrawing(`${how}:${kind}`);
                        });
                    }
                }
            },
        });
        group(insertTab, "テキスト");
        button("textBox", { icon: "textBox", label: "テキスト\nボックス", large: true, cls: "two", title: "テキスト ボックス: シートをドラッグして描きます", edits: true, act: () => startDrawing("textBox") });
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
        button("symbol", { icon: "symbol", label: "記号と\n特殊文字", large: true, cls: "two", title: "記号と特殊文字: 入力中のセル（またはアクティブなセル）に記号を入力します", edits: true, keepEntry: true, act: askSymbol });

        // ---- 数式
        const formulasTab = tab("formulas", "数式");
        group(formulasTab, "ワークシート分析");
        const f = button("formulas", { icon: "formulas", label: "数式の表示", large: true, title: "セルに値ではなく数式を表示します（Ctrl+` / Ctrl+Shift+@）", act: () => showFormulas(!formulas) });
        f.id = "formulas";
        f.setAttribute("aria-pressed", "false");
        group(formulasTab, "定義された名前");
        button("nameManager", { icon: "nameManager", label: "名前の\n管理", large: true, cls: "two", title: "名前の管理（Ctrl+F3）: ブックの名前の一覧を表示し、新しい名前の作成・編集・削除をします", edits: true, act: () => manageNames() });
        stack();
        button("defineName", { icon: "defineName", label: "名前の定義", showLabel: true, title: "名前の定義: 選択したセル（または値や数式）に名前を付けます", edits: true, act: defineName });

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
        group(dataTab, "データ ツール");
        button("validation", {
            icon: "validation",
            label: "データの\n入力規則",
            large: true,
            split: true,
            cls: "two",
            title: "データの入力規則: セルに入力できる値（整数・リスト・日付など）を制限し、入力時メッセージやエラー メッセージを設定します",
            moreTitle: "データの入力規則",
            edits: true,
            act: askValidation,
            menu: (anchor) => {
                const pop = openPop(anchor, "validations");
                item(pop, "データの入力規則…", askValidation);
                item(pop, "無効データのマーク", () => markInvalid(true));
                item(pop, "入力規則マークのクリア", () => markInvalid(false));
            },
        });

        // ---- 表示
        const viewTab = tab("view", "表示");
        group(viewTab, "ズーム");
        button("zoomBox", { icon: "zoom", label: "ズーム", large: true, menu: (anchor) => openZoomBox(anchor) });
        button("zoom100", { icon: "zoom100", label: "100%", large: true, title: "100% に戻す", act: () => setZoom(100) });
        button("zoomSel", { icon: "zoomSel", label: "選択範囲に合わせて\n拡大/縮小", large: true, cls: "two", act: zoomToSelection });
        group(viewTab, "ウィンドウ");
        button("freeze", { icon: "freeze", label: "ウィンドウ枠の\n固定", large: true, cls: "two", edits: true, menu: (anchor) => openFreeze(anchor) });

        // ---- 図形の書式 (shown while a shape, a picture or a chart is selected)
        const shapeTab = tab("shapeFormat", "図形の書式");
        panels.shapeFormat.tab.hidden = true;
        panels.shapeFormat.tab.classList.add("context");
        // (What is chosen is done to each of the objects selected.)
        const onObjs = (changeOf) => () => editSelected(changeOf);
        group(shapeTab, "図形のスタイル");
        const shapeColor = (key, iconName, label, none, more) => {
            const pick = (color) => {
                if (color) {
                    lastShape[key] = color;
                    const bar = tools[key === "fill" ? "shapeFill" : "shapeLine"].querySelector(".bar");
                    if (bar) bar.style.background = color;
                }
                editSelected((obj) => (key === "fill" ? canFill(obj) : canLine(obj)) && { [key]: color });
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
            for (const w of LINE_WIDTHS) item(pop, `${w} pt`, onObjs((obj) => canLine(obj) && { lineWidth: w }));
        });
        group(shapeTab, "配置");
        stack();
        const orderButton = (key, iconName, label, one, all) => {
            button(key, {
                icon: iconName,
                label,
                showLabel: true,
                split: true,
                title: label,
                act: onObjs((obj) => !isControl(obj) && { order: one }),
                menu: (anchor) => {
                    const pop = openPop(anchor, "orders");
                    for (const [how, text] of ORDERS.filter(([how]) => how === one || how === all)) item(pop, text, onObjs((obj) => !isControl(obj) && { order: how }));
                },
            });
        };
        orderButton("forward", "forward", "前面へ移動", "forward", "front");
        orderButton("backward", "backward", "背面へ移動", "backward", "back");
        button("selectionPane", { icon: "selectionPane", label: "オブジェクトの選択と表示", showLabel: true, title: "オブジェクトの選択と表示: シートの図形・コントロールの一覧を右に表示します", act: () => openPane("list") });
        stack();
        button("align", {
            icon: "align",
            label: "配置",
            title: "オブジェクトの配置",
            menu: (anchor) => {
                const pop = openPop(anchor, "aligns");
                // (Two or more objects selected: lined up; three or more: spread.)
                const n = alignable().length;
                ALIGNS.forEach((group, i) => {
                    if (i) pop.appendChild(el("hr"));
                    for (const [how, iconName, label, title] of group) {
                        const b = item(pop, label, () => alignObjects(how));
                        b.insertBefore(icon(iconName), b.firstChild);
                        b.title = title;
                        b.disabled = n < (how.startsWith("spread") ? 3 : 2);
                    }
                });
                pop.appendChild(el("hr"));
                const snap = item(pop, "枠線に合わせる", () => remember({ snapGrid: !saved.snapGrid }));
                snap.insertBefore(icon("snapGrid"), snap.firstChild);
                snap.classList.toggle("on", !!saved.snapGrid);
                snap.setAttribute("aria-pressed", String(!!saved.snapGrid));
                snap.title = "移動・サイズの変更でセルの枠線に合わせます（Alt を押しながらドラッグしても合わせます）";
            },
        });
        button("rotate", {
            icon: "rotate",
            label: "回転",
            menu: (anchor) => {
                const pop = openPop(anchor, "rotates");
                item(pop, "右へ 90 度回転", onObjs((obj) => turnChange(obj, 90)));
                item(pop, "左へ 90 度回転", onObjs((obj) => turnChange(obj, -90)));
                item(pop, "上下反転", onObjs((obj) => flipChange(obj, true)));
                item(pop, "左右反転", onObjs((obj) => flipChange(obj, false)));
                pop.appendChild(el("hr"));
                item(pop, "その他の回転オプション…", () => selectedObj && openPane("format", "rotation"));
            },
        });

        // サイズ: the height and the width (cm, as Excel shows them in Japanese).
        group(shapeTab, "サイズ");
        launcher("書式設定: 右に図形・コントロールの書式設定を表示します", () => openPane("format", "height"));
        stack().classList.add("sizes");
        for (const [key, iconName, label] of [["objHeight", "height", "図形の高さ"], ["objWidth", "width", "図形の幅"]]) {
            row();
            into.appendChild(icon(iconName)).title = label;
            // (Several selected: each made as high, or as wide.)
            const box = sizeBox(into, label, (cm) => {
                const px = (cm * 96) / 2.54;
                moveObjects(alignable().map((obj) => {
                    const b = boxOfObject(obj.box);
                    return [obj, key === "objHeight" ? { ...b, h: px } : { ...b, w: px }];
                }));
            });
            box.dataset.tool = key;
            tools[key] = box;
            into = into.parentNode;
        }

        // コード: the events of the object selected (in the sheet's file; with a project's workbook).
        group(shapeTab, "コード");
        const ev = button("event", {
            icon: "code",
            label: "コードの表示",
            large: true,
            cls: "two",
            title: "選択した図形・コントロールのイベント（クリック・選択など）のコードを開きます。シートのファイルにハンドラーを追加します（図形をダブルクリックすると既定のイベント）",
            menu: (anchor) => {
                const items = eventItems(selectedObj);
                if (!items.length) return;
                const r = anchor.getBoundingClientRect();
                showMenu(r.left, r.bottom + 1, items);
            },
        });
        ev.id = "event";
        ev.disabled = true;
        tools.macroGroup = ev.closest(".rgroup");

        // The mode, at the right end of the tabs (with the project's workbook): 編集 edits the
        // workbook file, 実行 raises the program's events for what is done on the sheet.
        const mode = tabs.appendChild(el("div", "rmode"));
        mode.hidden = true;
        mode.setAttribute("role", "group");
        mode.setAttribute("aria-label", "モード");
        for (const [key, label, title] of [
            ["edit", "編集", "編集モード: ブックのファイルを編集します。図形・コントロールは選択して動かせます（ダブルクリックでイベントのコード）"],
            ["run", "実行", "実行モード: プログラムを Excel の代役で実行し、セルの選択・入力、シートの切り替え、コントロール・図形のクリックでイベントを起こします（ブックのファイルは変えません）"],
        ]) {
            const b = mode.appendChild(el("button", `mode-${key}`));
            b.type = "button";
            b.dataset.mode = key;
            b.title = title;
            b.appendChild(el("span", "mico", key === "run" ? "▶" : "✎"));
            b.appendChild(el("span", null, label));
            b.addEventListener("click", () => askMode(key === "run"));
        }
        tools.modeSwitch = mode;

        if (!panels[shown] || panels[shown].tab.hidden) shown = "home";
        show(shown, true);
        ribbon.classList.toggle("collapsed", collapsed);
        edges();
        ribbon.showShapeTab = (on) => {
            panels.shapeFormat.tab.hidden = !on;
            if (on && saved.ribbonTab === "shapeFormat" && shown !== "shapeFormat") show("shapeFormat", true);
            if (!on && shown === "shapeFormat") show("home", true);
        };
    }
    // The buttons show the format of the active cell.
    function showTools() {
        if (!tools.bold || !view) return;
        const css = activeCss();
        const cell = (activeTd && activeTd.cell) || {};
        for (const name of ["bold", "italic", "underline", "wrap"]) tools[name].classList.toggle("on", STATE[name](css));
        // While characters are formatted: theirs.
        if (entry && entry.rich && charTarget()) {
            const f = charFont();
            tools.bold.classList.toggle("on", !!f.b);
            tools.italic.classList.toggle("on", !!f.i);
            tools.underline.classList.toggle("on", !!f.u);
            const size = String(f.size || fontOfCss(view.sheet, css).pt);
            if (![...tools.fontSize.options].some((o) => o.value === size)) tools.fontSize.appendChild(el("option", null, size)).value = size;
            tools.fontSize.value = size;
            if (document.activeElement !== tools.fontName) tools.fontName.value = f.name || fontOfCss(view.sheet, css).name;
            return;
        }
        const align = (/text-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["left", "center", "right"]) tools[`align-${a}`].classList.toggle("on", align === a);
        const valign = (/vertical-align:(\w+)/.exec(css) || [])[1];
        for (const a of ["top", "middle", "bottom"]) tools[`valign-${a}`].classList.toggle("on", valign === a);
        const font = fontOfCss(view.sheet, css);
        const size = String(font.pt);
        if (![...tools.fontSize.options].some((o) => o.value === size)) tools.fontSize.appendChild(el("option", null, size)).value = size;
        tools.fontSize.value = size;
        if (document.activeElement !== tools.fontName) tools.fontName.value = font.name;
        const code = (sel ? numberCodeAt(sel.ar, sel.ac)[0] : cell.n) || "General";
        tools.numberFormat.value = NUMBER_FORMATS.some(([c]) => c === code) ? code : "";
        tools.merge.classList.toggle("on", !!sel && single(sel) && mergedIn(sel));
        tools.filter.classList.toggle("on", !!view.sheet.filter);
        tools.filterClear.disabled = !editable || !(view.sheet.filter && Object.keys(view.sheet.filterColumns || {}).length);
        tools.freeze.classList.toggle("on", !!(view.fr || view.fc));
    }
    makeRibbon();
    setRunMode(runMode);

    // ---- the widths of columns and the heights of rows: the edges of the headers

    // How near an edge of a header the mouse takes it (px).
    const EDGE = 4;
    let resizing = null;
    const guide = el("div", "guide");
    const sizeTip = el("div", "sizetip");
    // The edge of a header under the mouse: { cols, index } (the column or row before the edge).
    function edgeOf(e) {
        if (!editable || !view || busy) return null;
        const th = e.target.closest && e.target.closest("th");
        if (!th || !view.table.contains(th) || (th.r === undefined && th.c === undefined)) return null;
        const b = th.getBoundingClientRect();
        if (th.c !== undefined) {
            if (b.right - e.clientX <= EDGE) return { cols: true, index: th.c, th };
            const i = view.cols.indexOf(th.c);
            if (e.clientX - b.left <= EDGE && i > 0) {
                // Right of a line with hidden columns before it: the last of them, which the
                // drag shows again (as Excel).
                if (th.c - 1 > view.cols[i - 1]) return { cols: true, index: th.c - 1, th, from: b.left };
                return { cols: true, index: view.cols[i - 1], th: view.ths.get(view.cols[i - 1]) };
            }
            if (i === 0 && th.c > 0 && e.clientX - b.left <= EDGE) return { cols: true, index: th.c - 1, th, from: b.left };
            return null;
        }
        if (b.bottom - e.clientY <= EDGE) return { cols: false, index: th.r, th };
        const i = view.rows.findIndex((x) => x[0] === th.r);
        if (e.clientY - b.top <= EDGE && i > 0) {
            // Under a line with hidden rows above it: the last of them.
            if (th.r - 1 > view.rows[i - 1][0]) return { cols: false, index: th.r - 1, th, from: b.top };
            return { cols: false, index: view.rows[i - 1][0], th: view.rows[i - 1][1].firstChild };
        }
        // The first rows (columns) hidden: under (right of) the headers' line.
        if (i === 0 && th.r > 0 && e.clientY - b.top <= EDGE) return { cols: false, index: th.r - 1, th, from: b.top };
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
        const start = edge.from !== undefined ? edge.from : edge.cols ? b.left : b.top;
        resizing = Object.assign(edge, { start, at: edge.cols ? e.clientX : e.clientY, moved: false });
        document.body.appendChild(guide);
        document.body.appendChild(sizeTip);
        guide.className = edge.cols ? "guide v" : "guide h";
        moveResize(e);
    }
    // The size the edge dragged to gives (px at 100%; columns are drawn twice as wide while
    // formulas are shown).
    const resizedPx = (r) => Math.round((r.now - r.start) / (r.cols && formulas ? 2 : 1) / Z());
    // A column width as Excel gives it (characters: 8 px a digit and 5 px of padding).
    const widthChars = (px) => Math.max(0, Math.trunc((px < 13 ? px / 13 : (px - 5) / 8) * 100 + 0.5) / 100);
    function moveResize(e) {
        const r = resizing;
        const at = Math.max(r.start, r.cols ? e.clientX : e.clientY);
        if (Math.abs(at - r.at) > 1) r.moved = true;
        r.now = at;
        // The line across the grid (from its headers), not over the ribbon or the tabs.
        const g = $("grid").getBoundingClientRect();
        if (r.cols) Object.assign(guide.style, { left: `${at}px`, top: `${g.top}px`, height: `${g.height}px`, width: "" });
        else Object.assign(guide.style, { left: `${g.left}px`, top: `${at}px`, width: `${g.width}px`, height: "" });
        // The size, as Excel shows it: 幅: 8.38 (72 ピクセル) / 高さ: 18.75 (25 ピクセル).
        const px = resizedPx(r);
        sizeTip.textContent = r.cols ? `幅: ${widthChars(px).toFixed(2)} (${px} ピクセル)` : `高さ: ${(px * 0.75).toFixed(2)} (${px} ピクセル)`;
        const w = sizeTip.offsetWidth;
        const h = sizeTip.offsetHeight;
        if (r.cols) Object.assign(sizeTip.style, { left: `${Math.min(at + 4, window.innerWidth - w - 2)}px`, top: `${Math.max(0, g.top - h - 2)}px` });
        else Object.assign(sizeTip.style, { left: `${g.left + 2}px`, top: `${Math.min(at + 4, window.innerHeight - h - 2)}px` });
    }
    function endResize() {
        const r = resizing;
        resizing = null;
        guide.remove();
        sizeTip.remove();
        if (!r.moved || !view) return;
        const px = resizedPx(r);
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
    // The icons of the filter's menu (as Excel's).
    const FILTER_ICONS = {
        asc: '<text x="1" y="7" font-size="7.5" font-weight="bold" fill="currentColor" stroke="none">A</text><text x="1" y="15" font-size="7.5" font-weight="bold" fill="currentColor" stroke="none">Z</text><path d="M12 2v11.5M9.5 11 12 13.5 14.5 11"/>',
        desc: '<text x="1" y="7" font-size="7.5" font-weight="bold" fill="currentColor" stroke="none">Z</text><text x="1" y="15" font-size="7.5" font-weight="bold" fill="currentColor" stroke="none">A</text><path d="M12 2v11.5M9.5 11 12 13.5 14.5 11"/>',
        clear: '<path d="M1.5 2.5h10L8 7v5.5l-2.5-1.5V7z"/><path d="m10 9.5 4 4M14 9.5l-4 4"/>',
    };
    const filterIcon = (name) => {
        const i = el("i", "fico");
        if (name) i.innerHTML = `<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.1" font-family="Segoe UI, Arial, sans-serif">${FILTER_ICONS[name]}</svg>`;
        return i;
    };
    // An item of the filter's menu: its icon, its label and the key (`(S)`) that chooses it.
    const filterItem = (pop, icon, label, key, act) => {
        const b = act ? popItem(pop, label, act) : pop.appendChild(el("button"));
        if (!act) {
            b.type = "button";
            b.disabled = true;
            b.appendChild(el("span", null, label));
        }
        b.insertBefore(filterIcon(icon), b.firstChild);
        if (key) {
            b.querySelector("span").append("(", el("u", null, key), ")");
            b.dataset.key = key.toLowerCase();
        }
        return b;
    };
    // An item that opens a submenu (`fill`), or one shown disabled (without `fill`).
    const filterSub = (pop, label, key, fill) => {
        let b;
        if (fill) b = subMenu(pop, label, fill);
        else {
            b = pop.appendChild(el("button", "hassub"));
            b.type = "button";
            b.disabled = true;
            b.appendChild(el("span", null, label));
            b.appendChild(el("i", "subarrow"));
        }
        b.insertBefore(filterIcon(null), b.firstChild);
        if (key) {
            const span = b.querySelector("span");
            span.append("(", el("u", null, key), ")");
            b.dataset.key = key.toLowerCase();
        }
        return b;
    };
    function openFilter(at) {
        const v = view;
        const sheet = v.sheet;
        const [[r1, c1], [r2]] = v.filter;
        const column = colName(at.c);
        const now = (sheet.filterColumns || {})[at.c - c1] || null;
        // The texts the column shows, each once; empty cells last. Dates go by year, month, day.
        const texts = new Set();
        const days = new Map(); // "yyyy-mm-dd" → { y, m, d, texts }
        // (With a comparison or a top ten, the values checked are the ones of the rows shown.)
        const seen = new Set();
        let blanks = false;
        let seenBlanks = false;
        let numbers = 0;
        for (let r = r1 + 1; r <= r2; r++) {
            const cell = v.cells.get(r * KEY + at.c);
            if (cell && (cell.k === "number" || cell.k === "date")) numbers++;
            if (!cell || !cell.t) {
                blanks = true;
                if (height(sheet, r) > 0) seenBlanks = true;
                continue;
            }
            if (height(sheet, r) > 0) seen.add(cell.t);
            if (cell.k === "date" && typeof cell.x === "number" && cell.x >= 1) {
                const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(cell.x) * 864e5);
                const key = d.toISOString().slice(0, 10);
                if (!days.has(key)) days.set(key, { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), texts: new Set() });
                days.get(key).texts.add(cell.t);
            } else texts.add(cell.t);
        }
        const list = [...texts].sort((a, b) => a.localeCompare(b, "ja", { numeric: true }));
        const dated = days.size > 0 && days.size * 2 >= texts.size + days.size;
        const numeric = !dated && numbers * 2 > texts.size && numbers > 0;
        const allTexts = list.concat([...days.values()].flatMap((x) => [...x.texts]));
        const pop = openPop(at.td, "filterpop");
        const sort = (desc) => postEdit([{ op: "sort", sheet: sheet.name, range: sheet.filter, by: column, desc, header: true }]);
        filterItem(pop, "asc", "昇順", "S", () => sort(false)).dataset.sort = "asc";
        filterItem(pop, "desc", "降順", "O", () => sort(true)).dataset.sort = "desc";
        filterSub(pop, "色で並べ替え", "T");
        pop.appendChild(el("hr"));
        filterSub(pop, "シート ビュー", "V");
        pop.appendChild(el("hr"));
        const head = v.cells.get(r1 * KEY + at.c);
        const name = head && head.t ? `"${head.t}"` : `"(列 ${column})"`;
        const clear = filterItem(pop, "clear", `${name} からフィルターをクリア`, "C", now ? () => postEdit([{ op: "filterBy", sheet: sheet.name, column, clear: true }]) : null);
        if (now) clear.dataset.clear = "1";
        filterSub(pop, "色フィルター", "I");
        // The filters of the column's kind, as Excel's submenu: each opens the comparisons
        // ("オートフィルター オプション") with its comparison chosen.
        const custom = (preset) => () => askCustomFilter(column, now, allTexts, preset);
        const kinds = dated
            ? [
                  ["指定の値に等しい…", ["eq"]],
                  ["指定の値より前…", ["lt"]],
                  ["指定の値より後…", ["gt"]],
                  ["指定の範囲内…", ["ge", "le"]],
              ]
            : numeric
              ? [
                    ["指定の値に等しい…", ["eq"]],
                    ["指定の値に等しくない…", ["ne"]],
                    ["指定の値より大きい…", ["gt"]],
                    ["指定の値以上…", ["ge"]],
                    ["指定の値より小さい…", ["lt"]],
                    ["指定の値以下…", ["le"]],
                    ["指定の範囲内…", ["ge", "le"]],
                ]
              : [
                    ["指定の値に等しい…", ["eq"]],
                    ["指定の値に等しくない…", ["ne"]],
                    ["指定の値で始まる…", ["begins"]],
                    ["指定の値で終わる…", ["ends"]],
                    ["指定の値を含む…", ["contains"]],
                    ["指定の値を含まない…", ["notContains"]],
                ];
        const kindMenu = filterSub(pop, dated ? "日付フィルター" : numeric ? "数値フィルター" : "テキスト フィルター", "F", (sub) => {
            for (const [label, preset] of kinds) popItem(sub, label, custom(preset));
            if (numeric) {
                sub.appendChild(el("hr"));
                popItem(sub, "トップテン…", () => askTop10(column, now)).dataset.top = "1";
            }
            sub.appendChild(el("hr"));
            popItem(sub, "ユーザー設定フィルター…", custom(null)).dataset.custom = "1";
        });
        kindMenu.dataset.kinds = dated ? "date" : numeric ? "number" : "text";
        if (now && (now.custom || now.top)) kindMenu.classList.add("set");
        // The search and the values: a tree, (すべて選択) first, the dates under their year and month.
        const form = pop.appendChild(el("form"));
        const search = form.appendChild(el("input", "search"));
        search.type = "search";
        search.placeholder = "(すべて) の検索";
        search.setAttribute("aria-label", "検索");
        const box = form.appendChild(el("div", "values"));
        box.setAttribute("role", "tree");
        const leaves = [];
        const parents = [];
        const shown = (text) => (!now ? true : now.values ? now.values.some((x) => x.toLowerCase() === text.toLowerCase()) : seen.has(text));
        const row = (parent, label, depth, node) => {
            const line = parent.appendChild(el("div", "node"));
            line.style.paddingLeft = `${2 + depth * 14}px`;
            const twist = line.appendChild(el("span", "twist"));
            const lab = line.appendChild(el("label"));
            const cb = lab.appendChild(el("input"));
            cb.type = "checkbox";
            lab.appendChild(document.createTextNode(label));
            Object.assign(node, { line, box: cb, label });
            if (!node.kids) twist.classList.add("none");
            else {
                const kids = parent.appendChild(el("div", "kids"));
                node.group = kids;
                twist.classList.add("open");
                // `open`: as the user left it; a search shows every parent open for the time.
                const toggle = (open) => {
                    node.open = open;
                    show(open);
                };
                const show = (open) => {
                    twist.classList.toggle("open", open);
                    kids.hidden = !open;
                };
                Object.assign(node, { open: true, toggle, show });
                twist.addEventListener("click", () => toggle(kids.hidden));
                parents.push(node);
            }
            return node;
        };
        const leaf = (parent, label, depth, values, blank) => {
            const node = row(parent, label, depth, { values, blank, search: values.join("\n") || label });
            node.box.checked = blank ? (!now ? true : now.values ? !!now.blanks : seenBlanks) : values.some(shown);
            leaves.push(node);
            return node;
        };
        const all = row(box, "(すべて選択)", 0, { kids: [] });
        all.line.classList.add("all");
        all.group.className = "top";
        all.toggle = all.show = () => {};
        all.line.querySelector(".twist").className = "twist none";
        const walk = (node, f) => {
            f(node);
            for (const k of node.kids || []) walk(k, f);
        };
        // The dates: years, months, days (Excel's grouping of dates).
        const years = new Map();
        for (const day of [...days.values()].sort((a, b) => a.y - b.y || a.m - b.m || a.d - b.d)) {
            if (!years.has(day.y)) years.set(day.y, new Map());
            const months = years.get(day.y);
            if (!months.has(day.m)) months.set(day.m, []);
            months.get(day.m).push(day);
        }
        for (const text of list) all.kids.push(leaf(all.group, text, 0, [text]));
        for (const [y, months] of years) {
            const year = row(all.group, `${y}年`, 0, { kids: [] });
            all.kids.push(year);
            for (const [m, list] of months) {
                const month = row(year.group, `${m}月`, 1, { kids: [] });
                year.kids.push(month);
                month.toggle(false);
                for (const day of list) month.kids.push(leaf(month.group, String(day.d).padStart(2, "0"), 2, [...day.texts]));
            }
        }
        if (blanks) all.kids.push(leaf(all.group, BLANK, 0, [], true));
        const buttons = form.appendChild(el("div", "buttons"));
        const ok = buttons.appendChild(el("button", "ok", "OK"));
        ok.type = "submit";
        const cancel = buttons.appendChild(el("button", null, "キャンセル"));
        cancel.type = "button";
        cancel.addEventListener("click", () => {
            closeMenu();
            focusGrid();
        });
        const visible = (n) => !n.line.hidden;
        const sync = () => {
            // Each parent from its children, the deepest first.
            for (const p of [...parents].reverse()) {
                const kids = p.kids.filter(visible);
                const on = kids.filter((k) => k.box.checked && !k.box.indeterminate).length;
                const some = kids.some((k) => k.box.checked || k.box.indeterminate);
                p.box.checked = kids.length > 0 && on === kids.length;
                p.box.indeterminate = some && on < kids.length;
            }
            ok.disabled = !leaves.some((l) => visible(l) && l.box.checked);
        };
        box.addEventListener("change", (e) => {
            const node = [...parents, ...leaves].find((n) => n.box === e.target);
            if (node && node.kids) walk(node, (n) => {
                if (n.line.hidden) return;
                n.box.checked = node.box.checked;
                n.box.indeterminate = false;
            });
            sync();
        });
        // The search: only the values with the text are listed, all of them checked (Excel's
        // search); an empty search gives the list back as it was.
        const before = new Map(leaves.map((l) => [l, l.box.checked]));
        search.addEventListener("input", () => {
            const q = search.value.trim().toLowerCase();
            for (const l of leaves) {
                const hit = !q || l.search.toLowerCase().includes(q) || l.label.toLowerCase().includes(q);
                l.line.hidden = !hit;
                l.box.checked = q ? hit : before.get(l);
            }
            for (const p of [...parents].reverse()) {
                if (p === all) continue;
                p.line.hidden = !p.kids.some(visible);
                p.show(!!q || p.open);
            }
            all.line.querySelector("label").lastChild.textContent = q ? "(検索結果をすべて選択)" : "(すべて選択)";
            sync();
        });
        sync();
        // The keys of the items: (S), (O), …
        pop.addEventListener("keydown", (e) => {
            if (e.target === search || e.ctrlKey || e.altKey || e.metaKey) return;
            const b = pop.querySelector(`:scope > button[data-key="${CSS.escape(e.key.toLowerCase())}"]:not(:disabled)`);
            if (!b) return;
            e.preventDefault();
            b.click();
        });
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const on = leaves.filter((l) => visible(l) && l.box.checked);
            closeMenu();
            focusGrid();
            if (!on.length) return;
            if (on.length === leaves.length) {
                if (now) postEdit([{ op: "filterBy", sheet: sheet.name, column, clear: true }]);
                return;
            }
            const op = { op: "filterBy", sheet: sheet.name, column, values: on.filter((l) => !l.blank).flatMap((l) => l.values) };
            if (on.some((l) => l.blank)) op.blanks = true;
            postEdit([op]);
        });
        requestAnimationFrame(() => search.focus());
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
    // `preset`: the comparisons chosen first (Excel's 数値フィルター > 指定の値より大きい…).
    function askCustomFilter(column, now, values, preset) {
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
        const had = preset ? preset.map((k) => [k, ""]) : now && now.custom ? [now.custom.a, now.custom.b].filter(Boolean).map(kindOf) : [];
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
    // ---- the number formats of "セルの書式設定" (表示形式, in the pane at the right)

    // The colors of a format as Japanese Excel spells them ([赤]), and as the file does ([Red]).
    const FORMAT_COLORS = [["Black", "黒"], ["Blue", "青"], ["Cyan", "水"], ["Green", "緑"], ["Magenta", "紫"], ["Red", "赤"], ["White", "白"], ["Yellow", "黄"]];
    // A format as the dialog shows it (Japanese Excel's NumberFormatLocal).
    function localCode(code) {
        if (code === "General") return "G/標準";
        let out = code.replace(/General/gi, "G/標準");
        for (const [en, ja] of FORMAT_COLORS) out = out.replace(new RegExp(`\\[${en}\\]`, "gi"), `[${ja}]`);
        return out;
    }
    // A format typed into the dialog as the file keeps it: the system's short date is Excel's
    // built-in one (`m/d/yyyy`, shown `yyyy/m/d`), as when Japanese Excel is given it.
    function fileCode(code) {
        if (/^(G\/標準|general)$/i.test(code)) return "General";
        if (code === "yyyy/m/d") return "m/d/yyyy";
        if (code === "yyyy/m/d h:mm") return "m/d/yyyy h:mm";
        let out = code.split("G/標準").join("General");
        for (const [en, ja] of FORMAT_COLORS) out = out.split(`[${ja}]`).join(`[${en}]`);
        return out;
    }
    // The formats of the 日付 and 時刻 lists (as Japanese Excel 2021 offers them, eras too; not
    // those of 午前/午後, which the cells do not show as Excel does). Those of the system (Excel
    // marks them with *): its short date (`yyyy/m/d`, `m/d/yyyy` in the file), its long date and
    // its time.
    const SYSTEM_CODES = ["yyyy/m/d", "[$-F800]dddd, mmmm dd, yyyy", "[$-F400]h:mm:ss AM/PM"];
    const DATE_CODES = ["yyyy/m/d", "[$-F800]dddd, mmmm dd, yyyy", 'yyyy"年"m"月"d"日"', 'yyyy"年"m"月"', 'm"月"d"日"', "[$-411]ge.m.d", '[$-411]ggge"年"m"月"d"日"', '[$-ja-JP-x-gannen]ggge"年"m"月"d"日"', '[$-411]gge"年"m"月"d"日"', "[$-411]ge/m/d", "yyyy/m/d(aaa)", 'yyyy"年"m"月"d"日"(aaaa)', "m/d", "yy/m/d", "yyyy/mm/dd", "yyyy-mm-dd", "m/d/yy", "d-mmm", "d-mmm-yy", "mmm-yy", "yyyy/m/d h:mm", "yyyy/m/d h:mm AM/PM"];
    const TIME_CODES = ["[$-F400]h:mm:ss AM/PM", "h:mm", "[$-409]h:mm AM/PM", "h:mm:ss", "[$-409]h:mm:ss AM/PM", "[$-409]yyyy/m/d h:mm AM/PM", "yyyy/m/d h:mm", 'h"時"mm"分"', 'h"時"mm"分"ss"秒"', "mm:ss.0", "[h]:mm:ss"];
    // The 分数 and the その他 (Japanese) lists: their formats and the names Excel gives them.
    const FRACTION_CODES = [
        ["# ?/?", "1 桁増加 (1/4)"],
        ["# ??/??", "2 桁増加 (21/25)"],
        ["# ???/???", "3 桁増加 (312/943)"],
        ["# ?/2", "2 分の 1 (1/2)"],
        ["# ?/4", "4 分の 1 (2/4)"],
        ["# ?/8", "8 分の 1 (4/8)"],
        ["# ??/16", "16 分の 1 (8/16)"],
        ["# ?/10", "10 分の 1 (3/10)"],
        ["# ??/100", "100 分の 1 (30/100)"],
    ];
    const SPECIAL_CODES = [
        ["[<=999]000;[<=9999]000-00;000-0000", "郵便番号"],
        ["[<=99999999]####-####;(00) ####-####", "電話番号 (東京)"],
        ["[DBNum1][$-411]G/標準", "漢数字 (十二万三千四百)"],
        ["[DBNum2][$-411]G/標準", "大字 (壱拾弐萬参阡四百)"],
        ["[DBNum3][$-411]0", "全角 (１２３４５)"],
        ["[DBNum3][$-411]#,##0", "全角 桁区切り (１２,３４５)"],
    ];
    const CUSTOM_CODES = ["G/標準", "0", "0.00", "#,##0", "#,##0.00", "#,##0;-#,##0", "#,##0;[赤]-#,##0", "#,##0.00;[赤]-#,##0.00", "¥#,##0;¥-#,##0", "¥#,##0;[赤]¥-#,##0", "0%", "0.00%", "0.00E+00", "##0.0E+0", "# ?/?", "# ??/??", "@"]
        .concat(DATE_CODES, TIME_CODES.filter((c) => !DATE_CODES.includes(c)), SPECIAL_CODES.map((c) => c[0]));
    // The numbers: the places, the separators, and how a negative number shows.
    const NEGATIVES = [
        ["minus", "-1,234", (p) => `${p};-${p}`],
        ["red", "1,234 (赤)", (p) => `${p};[赤]${p}`],
        ["redminus", "-1,234 (赤)", (p) => `${p};[赤]-${p}`],
        ["tri", "△ 1,234", (p) => `${p};"△ "${p}`],
        ["blacktri", "▲ 1,234", (p) => `${p};"▲ "${p}`],
    ];
    const places = (n) => (n > 0 ? `.${"0".repeat(n)}` : "");
    const FORMAT_KINDS = [
        { key: "general", label: "標準", note: "標準の表示形式では、特定の書式は設定されません。", code: () => "G/標準" },
        {
            key: "number",
            label: "数値",
            note: "数値の表示形式は一般的な数値を表すのに使用します。",
            places: true,
            comma: true,
            negative: true,
            code: (s) => {
                const p = `${s.comma ? "#,##0" : "0"}${places(s.places)}`;
                return s.negative === "minus" ? p : NEGATIVES.find((n) => n[0] === s.negative)[2](p);
            },
        },
        {
            key: "currency",
            label: "通貨",
            note: "通貨の表示形式は一般的な金額を表すのに使用します。",
            places: true,
            symbol: true,
            negative: true,
            code: (s) => {
                const p = `${s.symbol}#,##0${places(s.places)}`;
                const neg = { minus: `${s.symbol}-#,##0${places(s.places)}`, red: `[赤]${p}`, redminus: `[赤]${s.symbol}-#,##0${places(s.places)}`, tri: `"△ "${p}`, blacktri: `"▲ "${p}` }[s.negative];
                return `${p};${neg}`;
            },
        },
        {
            key: "accounting",
            label: "会計",
            note: "会計の表示形式は、通貨記号と小数点の位置をそろえて表示します。",
            places: true,
            symbol: true,
            code: (s) => {
                const sym = `${s.symbol}* `;
                const n = `#,##0${places(s.places)}`;
                return `_ ${sym}${n}_ ;_ ${sym}-${n}_ ;_ ${sym}"-"${"?".repeat(s.places)}_ ;_ @_ `;
            },
        },
        { key: "date", label: "日付", note: "日付の表示形式は日付/時刻のシリアル値を日付として表示します。", list: DATE_CODES },
        { key: "time", label: "時刻", note: "時刻の表示形式は日付/時刻のシリアル値を時刻として表示します。", list: TIME_CODES },
        { key: "percent", label: "パーセンテージ", note: "セルの値を 100 倍し、パーセント記号 (%) を付けて表示します。", places: true, code: (s) => `0${places(s.places)}%` },
        { key: "fraction", label: "分数", note: "", list: FRACTION_CODES.map((c) => c[0]), names: FRACTION_CODES.map((c) => c[1]) },
        { key: "scientific", label: "指数", note: "数値を指数で表示します。", places: true, code: (s) => `0${places(s.places)}E+00` },
        { key: "text", label: "文字列", note: "文字列として扱われます。数値も文字列として扱われ、入力したとおりに表示されます。", code: () => "@" },
        { key: "special", label: "その他", note: "その他の表示形式は、郵便番号や電話番号などを表示するときに使用します。", list: SPECIAL_CODES.map((c) => c[0]), names: SPECIAL_CODES.map((c) => c[1]) },
        { key: "custom", label: "ユーザー定義", note: "基になる組み込みの表示形式を選択し、新しい表示形式に変更します。", list: CUSTOM_CODES, custom: true },
    ];
    // The kind (and its settings) a format is one of, as the dialog opens on it.
    function kindOf(code) {
        const state = { places: 2, comma: false, symbol: "¥", negative: "minus" };
        for (const kind of FORMAT_KINDS) {
            if (kind.list) {
                if (!kind.custom && kind.list.includes(code)) return { kind, state, code };
                continue;
            }
            const tries = [{}];
            if (kind.places) for (let p = 0; p <= 10; p++) tries.push({ places: p });
            for (const t of tries) {
                for (const comma of kind.comma ? [false, true] : [false]) {
                    for (const symbol of kind.symbol ? ["¥", "$", ""] : ["¥"]) {
                        for (const [negative] of kind.negative ? NEGATIVES : [["minus"]]) {
                            const s = { ...state, ...t, comma, symbol, negative };
                            if (kind.code(s) === code) return { kind, state: s, code };
                        }
                    }
                }
            }
        }
        return { kind: FORMAT_KINDS[FORMAT_KINDS.length - 1], state, code };
    }
    // What a format shows of the active cell (null when it is not drawn here).
    function sampleText(cell, code) {
        const file = fileCode(code);
        const x = typeof cell.x === "number" ? cell.x : null;
        if (x === null) return cell.t ?? "";
        if (file === "@") return cell.t ?? String(x);
        const shown = formatNumber(x, file === "m/d/yyyy" ? "yyyy/m/d" : file === "m/d/yyyy h:mm" ? "yyyy/m/d h:mm" : file);
        return shown !== null ? shown : /^general$/i.test(file) ? generalFit(x, () => true) || String(x) : null;
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
    // What the fill puts in the cell it is drawn out to (shown by the mouse, as Excel shows it):
    // the same series the workbook continues (`autofill` in patch.rs), for the first cell across
    // the end of the fill. null: nothing to show (a formula, an empty cell, the cells drawn in).
    const fillTip = el("div", "sizetip filltip");
    function numbered(text) {
        const m = /^([\s\S]*?)([0-9]+)([^0-9]*)$/.exec(text);
        if (!m || m[2].length > 15) return null;
        return { head: m[1], n: Number(m[2]), digits: m[2].length, tail: m[3] };
    }
    function fillPreview(f, to) {
        const down = to[1] === f.c1 && to[3] === f.c2;
        const [lo, hi, toLo, toHi] = down ? [f.r1, f.r2, to[0], to[2]] : [f.c1, f.c2, to[1], to[3]];
        const n = hi - lo + 1;
        // The line along the fill: after the cells (down, right), else before them.
        const k = toHi > hi ? toHi - lo : toLo < lo ? toLo - lo : null;
        if (k === null) return null;
        const at = (i) => view.cells.get(down ? (lo + i) * KEY + f.c1 : f.r1 * KEY + lo + i);
        const cells = [];
        for (let i = 0; i < n; i++) cells.push(at(i) || {});
        if (cells.some((c) => c.f)) return null;
        const src = cells[((k % n) + n) % n];
        const code = src.sn || src.n;
        const shown = (x) => statText(Math.round(x * 1e10) / 1e10, code);
        if (cells.every((c) => (c.k === "number" || c.k === "date") && typeof c.x === "number")) {
            const xs = cells.map((c) => c.x);
            if (n >= 2) {
                // The line through the numbers (least squares).
                const mk = (n - 1) / 2;
                const mx = xs.reduce((a, b) => a + b, 0) / n;
                let num = 0,
                    den = 0;
                xs.forEach((x, i) => {
                    num += (i - mk) * (x - mx);
                    den += (i - mk) ** 2;
                });
                const slope = num / den;
                return shown(mx - slope * mk + slope * k);
            }
            if (cells[0].k === "date") return shown(xs[0] + k);
        } else if (cells.every((c) => c.k === "string")) {
            const ts = cells.map((c) => numbered(c.t || ""));
            if (ts.every(Boolean)) {
                const step = n >= 2 ? ts[1].n - ts[0].n : 1;
                const regular = ts.every((t, i) => i === 0 || (t.head === ts[i - 1].head && t.tail === ts[i - 1].tail && t.n - ts[i - 1].n === step));
                if (regular) {
                    const t = ts[0];
                    return t.head + String(Math.abs(t.n + step * k)).padStart(t.digits, "0") + t.tail;
                }
            }
        }
        return src.t === undefined || src.t === "" ? null : src.t;
    }
    function moveFill(e) {
        const f = filling;
        f.to = fillTarget(e);
        const b = boxOf(f.to[0], f.to[1], f.to[2], f.to[3]);
        view.fillbox.hidden = !b;
        if (b) Object.assign(view.fillbox.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w + 1}px`, height: `${b.h + 1}px` });
        pinTo(view.fillbox, pinOf(view, b));
        const text = b ? fillPreview(f, f.to) : null;
        if (text === null) {
            fillTip.remove();
            return;
        }
        document.body.appendChild(fillTip);
        fillTip.textContent = text;
        // At the bottom right of the cells drawn out to.
        const r = view.fillbox.getBoundingClientRect();
        const w = fillTip.offsetWidth;
        const h = fillTip.offsetHeight;
        Object.assign(fillTip.style, { left: `${Math.max(0, Math.min(r.right + 6, window.innerWidth - w - 2))}px`, top: `${Math.max(0, Math.min(r.bottom + 6, window.innerHeight - h - 2))}px` });
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
        fillTip.remove();
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

    // ---- 数式 > 定義された名前: "名前の管理" (Ctrl+F3: the names of the workbook, new ones,
    // edited and deleted; what the one chosen refers to changed below the list; OK writes the
    // changes, an edit `name` each) and "名前の定義" (a new name, written at once).

    // A cell's address (A1, XFD1048576), or R1C1 (R, C, R1C, RC1, ...): no name.
    function cellLike(text) {
        const m = /^([A-Za-z]{1,3})([1-9]\d{0,6})$/.exec(text);
        return (!!m && colNumber(m[1].toUpperCase()) < MAXC && +m[2] <= MAXR) || /^[Rr]\d*[Cc]?\d*$/.test(text) || /^[Cc]\d*$/.test(text);
    }
    // Why a name cannot be one (as Excel takes them), or "".
    function nameError(name) {
        if (!name) return "名前を入力してください。";
        if (name.length > 255 || !/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) || cellLike(name)) {
            return "入力した名前は正しくありません（先頭は文字か _ か \\、ほかは文字・数字・_・.・\\ で、セルの番地と同じ名前は使えません）。";
        }
        return "";
    }
    // A sheet's name in a reference: quoted when it must be ('My Sheet'!A1).
    const sheetPrefix = (name) => `${/^[A-Za-z_　-鿿＀-￯][\w.　-鿿＀-￯]*$/.test(name) && !cellLike(name) ? name : `'${name.replace(/'/g, "''")}'`}!`;
    // An area as an absolute reference ($A$1:$B$5, $3:$5, $B:$D).
    function absArea(a) {
        if (a.kind === "rows") return `$${a.r1 + 1}:$${a.r2 + 1}`;
        if (a.kind === "cols") return `$${colName(a.c1)}:$${colName(a.c2)}`;
        if (a.kind === "all") return `$1:$${MAXR}`;
        const one = `$${colName(a.c1)}$${a.r1 + 1}`;
        return a.r1 === a.r2 && a.c1 === a.c2 ? one : `${one}:$${colName(a.c2)}$${a.r2 + 1}`;
    }
    // The cells selected as 参照範囲 has them (=Sheet1!$A$1:$B$5,Sheet1!$D$2).
    const selectionRef = () => `=${allAreas().map((a) => sheetPrefix(view.sheet.name) + absArea(a)).join(",")}`;
    // What a name comes to, as 名前の管理 lists it (値): the text of its cell, the texts of its
    // cells as an array ({"1","2";"3","4"}), or its constant; "" when it is worked out.
    function nameValue(n) {
        const areas = refAreas(n);
        const ref = String(n.ref).replace(/^=/, "");
        if (!areas) return /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(ref) ? ref : /^"([^"]|"")*"$/.test(ref) ? ref.slice(1, -1).replace(/""/g, '"') : /^(TRUE|FALSE)$/i.test(ref) ? ref.toUpperCase() : "";
        if (areas.length !== 1) return "";
        const a = areas[0];
        const rows = a.r2 - a.r1 + 1;
        const cols = a.c2 - a.c1 + 1;
        if (rows * cols > 200) return "";
        const texts = new Map();
        for (const c of a.sheet.cells || []) if (c.r >= a.r1 && c.r <= a.r2 && c.c >= a.c1 && c.c <= a.c2) texts.set(c.r * KEY + c.c, c.t || "");
        const at = (r, c) => texts.get(r * KEY + c) || "";
        if (rows === 1 && cols === 1) return at(a.r1, a.c1);
        const out = [];
        for (let r = a.r1; r <= a.r2; r++) {
            const row = [];
            for (let c = a.c1; c <= a.c2; c++) row.push(`"${at(r, c)}"`);
            out.push(row.join(","));
        }
        return `{${out.join(";")}}`;
    }
    const scopeLabel = (sheet) => sheet || "ブック";
    const sameName = (a, b) => a.toLowerCase() === b.toLowerCase();

    // "新しい名前" / "名前の編集": `o.item` ({ name, sheet, ref, comment }) edited, or a new one;
    // `o.others`: the names it must not be the same as; `o.done(item)`, `o.back` (キャンセル).
    function askName(o) {
        const it = o.item || {};
        const form = openDialog(o.item ? "名前の編集" : "新しい名前", "namedlg", o.back);
        const name = field(form, "名前:");
        name.dataset.field = "name";
        // (New: the active cell's text, when it is a name, as Excel suggests it.)
        const active = !o.item && sel ? String((cellOf(sel.ar, sel.ac) || {}).t || "").trim().replace(/\s+/g, "_") : "";
        name.value = o.item ? it.name : active && !nameError(active) ? active : "";
        const scopeLine = form.appendChild(el("label"));
        scopeLine.appendChild(el("span", "lbl", "範囲:"));
        const scope = scopeLine.appendChild(el("select"));
        scope.dataset.field = "scope";
        scope.appendChild(el("option", null, "ブック")).value = "";
        for (const sh of data.sheets) scope.appendChild(el("option", null, sh.name)).value = sh.name;
        scope.value = it.sheet || "";
        // (As Excel's: a name keeps its scope.)
        scope.disabled = !!o.item;
        const comment = field(form, "コメント:", "textarea");
        comment.dataset.field = "comment";
        comment.rows = 3;
        comment.value = it.comment || "";
        const ref = field(form, "参照範囲:");
        ref.dataset.field = "ref";
        ref.value = o.item ? `=${String(it.ref).replace(/^=/, "")}` : selectionRef();
        for (const box of [name, ref]) box.addEventListener("input", () => box.setCustomValidity(""));
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const n = name.value.trim();
            const sheet = scope.value || undefined;
            const wrong = nameError(n) || ((o.others || []).some((x) => sameName(x.name, n) && (x.sheet || "") === (sheet || "")) ? "入力された名前は既に存在します。一意の名前を入力してください。" : "");
            if (wrong) {
                name.setCustomValidity(wrong);
                name.reportValidity();
                return;
            }
            const r = ref.value.trim().replace(/^=/, "").trim();
            if (!r) {
                ref.setCustomValidity("参照範囲を入力してください");
                ref.reportValidity();
                return;
            }
            o.done({ name: n, sheet, ref: r, comment: comment.value });
        });
        name.focus();
        name.select();
    }
    // 名前の定義: a new name of the cells selected, written at once.
    function defineName() {
        if (!editable || busy || !view || !sel) return;
        if (entry) commit(null);
        askName({
            others: data.names || [],
            done: (item) => {
                closeDialog();
                postEdit([{ op: "name", sheet: view.sheet.name, name: item.name, scope: item.sheet || null, ref: `=${item.ref}`, comment: item.comment }]);
            },
        });
    }
    // `state`: the names as the dialog has them (kept while one is edited in its own dialog).
    function manageNames(state) {
        if (!editable || busy || !view) return;
        if (entry) commit(null);
        if (!state) {
            const list = (data.names || []).map((n) => ({ was: n, name: n.name, sheet: n.sheet, ref: n.ref, comment: n.comment || "" }));
            state = { list, deleted: [], chosen: list.length ? 0 : -1, changed: false };
        }
        const reopen = () => manageNames(state);
        const form = openDialog("名前の管理", "namemgr");
        const tools = form.appendChild(el("div", "condtools"));
        const tool = (label, key, act) => {
            const b = tools.appendChild(el("button", null, label));
            b.type = "button";
            b.dataset.act = key;
            b.addEventListener("click", act);
            return b;
        };
        const others = (except) => state.list.filter((x) => x !== except);
        tool("新規作成…", "new", () =>
            askName({
                others: others(null),
                back: reopen,
                done: (item) => {
                    state.list.push(Object.assign({ was: null }, item));
                    state.chosen = state.list.length - 1;
                    state.changed = true;
                    reopen();
                },
            }),
        );
        const edit = tool("編集…", "edit", () => {
            const it = state.list[state.chosen];
            if (!it) return;
            askName({
                item: it,
                others: others(it),
                back: reopen,
                done: (item) => {
                    Object.assign(it, { name: item.name, ref: item.ref, comment: item.comment });
                    state.changed = true;
                    reopen();
                },
            });
        });
        const del = tool("削除", "delete", () => {
            const it = state.list[state.chosen];
            if (!it) return;
            if (it.was) state.deleted.push(it.was);
            state.list.splice(state.chosen, 1);
            state.chosen = Math.min(state.chosen, state.list.length - 1);
            state.changed = true;
            draw();
        });
        const table = form.appendChild(el("div", "namelist"));
        const refLine = form.appendChild(el("label", "nameref"));
        refLine.appendChild(el("span", "lbl", "参照範囲:"));
        const refBox = refLine.appendChild(el("input"));
        refBox.type = "text";
        refBox.spellcheck = false;
        refBox.dataset.field = "ref";
        refBox.addEventListener("input", () => refBox.setCustomValidity(""));
        refBox.addEventListener("change", () => {
            const it = state.list[state.chosen];
            const r = refBox.value.trim().replace(/^=/, "").trim();
            if (!it) return;
            if (!r) {
                refBox.setCustomValidity("参照範囲を入力してください");
                refBox.reportValidity();
                return;
            }
            if (r === it.ref) return;
            it.ref = r;
            state.changed = true;
            draw();
        });
        function draw() {
            table.textContent = "";
            const head = table.appendChild(el("div", "namehead"));
            for (const t of ["名前", "値", "参照範囲", "範囲", "コメント"]) head.appendChild(el("span", null, t));
            // (In the order of their names, as Excel lists them.)
            const order = state.list.map((x, i) => i).sort((a, b) => state.list[a].name.localeCompare(state.list[b].name, "ja") || scopeLabel(state.list[a].sheet).localeCompare(scopeLabel(state.list[b].sheet), "ja"));
            for (const i of order) {
                const it = state.list[i];
                const row = table.appendChild(el("div", `namerow${i === state.chosen ? " on" : ""}`));
                row.dataset.name = it.name;
                const cells = [it.name, nameValue(it), `=${it.ref}`, scopeLabel(it.sheet), it.comment || ""];
                cells.forEach((t, k) => {
                    const c = row.appendChild(el("span", null, t));
                    c.title = t;
                    if (k === 0) c.className = "nm";
                });
                row.addEventListener("mousedown", () => {
                    state.chosen = i;
                    for (const x of table.querySelectorAll(".namerow.on")) x.classList.remove("on");
                    row.classList.add("on");
                    showChosen();
                });
                row.addEventListener("dblclick", () => edit.click());
            }
            if (!state.list.length) table.appendChild(el("div", "condnone", "名前は定義されていません。［新規作成］で、セルや値に名前を付けられます。"));
            showChosen();
        }
        function showChosen() {
            const it = state.list[state.chosen];
            refBox.value = it ? `=${it.ref}` : "";
            refBox.disabled = !it;
            edit.disabled = del.disabled = !it;
        }
        draw();
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            // (The reference being typed, not yet taken.)
            if (document.activeElement === refBox) refBox.dispatchEvent(new Event("change"));
            if (!refBox.checkValidity()) return;
            closeDialog();
            if (!state.changed) return;
            const sheet = view.sheet.name;
            const ops = state.deleted.map((n) => ({ op: "name", sheet, delete: true, name: n.name, scope: n.sheet || null }));
            // (Those changed before those added, whose names they may free.)
            const added = [];
            for (const it of state.list) {
                const w = it.was;
                if (w && w.name === it.name && w.ref === it.ref && (w.comment || "") === it.comment) continue;
                const op = { op: "name", sheet, name: it.name, scope: it.sheet || null, ref: `=${it.ref}`, comment: it.comment };
                if (w) {
                    op.old = { name: w.name, scope: w.sheet || null };
                    ops.push(op);
                } else added.push(op);
            }
            ops.push(...added);
            if (ops.length) postEdit(ops);
        });
        if (state.chosen >= 0) {
            const on = table.querySelector(".namerow.on");
            if (on) on.scrollIntoView({ block: "nearest" });
        }
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
        // (Run mode: said before it, as a badge.)
        $("mode").classList.toggle("running", runMode);
    }
    // A number as the status bar shows it: in the number format of the active cell, as Excel
    // 2021 shows it (General in at most 11 characters, without separators).
    function statText(x, code) {
        const shown = code ? formatNumber(x, code, true) : null;
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
    // (`hours`: times without a day as hours on end, as the status bar shows a sum of them.)
    function formatNumber(x, code, hours) {
        if (/^general$/i.test(code)) return null;
        // The system's long date and time, as Japanese Windows has them.
        if (/^\[\$-F800\]/i.test(code)) code = 'yyyy"年"m"月"d"日"';
        else if (/^\[\$-F400\]/i.test(code)) code = "h:mm:ss";
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
        const fraction = /^(?:([#0?]+) +)?([#0?]+)\/([#0?]+|[1-9]\d*)$/.exec(s);
        if (fraction) return sign + formatFraction(Math.abs(x), fraction);
        const toks = tokensOf(s);
        if (!toks) return null;
        if (toks.some((t) => t.k === "date" || t.k === "ampm")) {
            // (Not dates before 1900; fractions of a second only as 0s after the seconds.)
            if (x < 0) return null;
            const after = (i) => toks.slice(0, i).reverse().find((t) => t.k !== "ph");
            const second = (i) => toks.slice(0, i).reverse().find((t) => t.k === "date");
            if (toks.some((t, i) => t.k === "ph" && (t.v !== "0" || (after(i) || {}).k !== "dot" || !/^s/.test((second(i) || {}).v || "")))) return null;
            // In the status bar, times (without a day) are shown as hours on end: 32:24:00.
            const day = toks.some((t) => t.k === "date" && (/^[ydage]/.test(t.v) || /^m{3,}$/.test(t.v)));
            const time = toks.some((t) => t.k === "ampm" || (t.k === "date" && /^[hs]/.test(t.v)));
            return formatDate(x, hours && time && !day ? tokensOf("[h]:mm:ss") : toks);
        }
        const out = formatDigits(Math.abs(x), toks);
        return out === null ? null : sign + out;
    }
    // A fraction (`# ?/?`, `# ??/16`): the nearest one with as many digits in the denominator,
    // or with the denominator given; the integer before it when the format has one.
    function formatFraction(x, [, whole, num, den]) {
        let w = whole ? Math.floor(x) : 0;
        const f = x - w;
        let n, d;
        if (/^\d+$/.test(den)) {
            d = Number(den);
            n = Math.round(f * d);
        } else {
            [n, d] = [Math.round(f), 1];
            for (let t = 2, best = Math.abs(f - n); t < 10 ** den.length; t++) {
                const e = Math.abs(f - Math.round(f * t) / t);
                if (e < best - 1e-12) [n, d, best] = [Math.round(f * t), t, e];
            }
        }
        if (whole && n === d) [w, n] = [w + 1, 0];
        const blank = " ".repeat(num.length + den.length + 1);
        if (whole && n === 0) return `${w} ${blank}`;
        return `${whole && w ? `${w} ` : ""}${String(n).padStart(num.length, " ")}/${/^\d+$/.test(den) ? den : String(d).padEnd(den.length, " ")}`;
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
                    // `[$-ja-JP-x-gannen]`: the first year of an era is 元.
                    if (m[2] !== undefined && /^ja-jp-x-gannen$/i.test(m[2])) toks.gannen = true;
                    else if (m[2] !== undefined && !/^[0-9a-f]{1,4}$/i.test(m[2])) return null;
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
            } else if (/^(general|G\/標準)/i.test(rest)) return null;
            else if ((m = /^(y+|m+|d+|h+|s+|a{3,4}|g+|e+)/i.exec(rest))) {
                // (g: the Japanese era, e: the year in it.)
                toks.push({ k: "date", v: m[1].toLowerCase() });
                i += m[1].length;
            } else if (/[bB]/.test(ch)) return null; // Buddhist years
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
    // The Japanese eras: the first day, the letter (g) and the name (ggg).
    const ERAS = [
        [Date.UTC(1868, 8, 8), "M", "明治"],
        [Date.UTC(1912, 6, 30), "T", "大正"],
        [Date.UTC(1926, 11, 25), "S", "昭和"],
        [Date.UTC(1989, 0, 8), "H", "平成"],
        [Date.UTC(2019, 4, 1), "R", "令和"],
    ];
    // The era of a date, and the year in it (dates before Meiji count as Meiji).
    function eraOf(d) {
        const [start, letter, name] = [...ERAS].reverse().find(([t]) => d.getTime() >= t) || ERAS[0];
        return { letter, name, year: d.getUTCFullYear() - new Date(start).getUTCFullYear() + 1 };
    }
    function formatDate(x, toks) {
        // The 0s after the seconds: their tenths, hundredths and thousandths.
        const places = Math.min(3, toks.filter((t) => t.k === "ph").length);
        let days = Math.floor(x);
        let parts = Math.round((x - days) * 86400 * 10 ** places);
        if (parts >= 86400 * 10 ** places) {
            days++;
            parts -= 86400 * 10 ** places;
        }
        const secs = Math.floor(parts / 10 ** places);
        const frac = String(parts % 10 ** places).padStart(places, "0");
        let firstPh = true;
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
            else if (t.k === "ph") {
                if (firstPh) out += frac;
                firstPh = false;
            } else if (t.k === "date") {
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
                else if (v[0] === "g" || v[0] === "e") {
                    const era = eraOf(d);
                    if (v[0] === "g") out += v.length === 1 ? era.letter : v.length === 2 ? era.name[0] : era.name;
                    else out += toks.gannen && era.year === 1 ? "元" : two(era.year, Math.min(v.length, 2));
                }
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
        const button = !!b && editable && rule.type === "list" && rule.inCellDropdown && !!rule.values && !entry;
        v.dropbtn.hidden = !button;
        if (button) {
            Object.assign(v.dropbtn.style, { left: `${b.x + b.w + 1}px`, top: `${b.y + Math.max(0, b.h - 17)}px` });
            pinTo(v.dropbtn, pinOf(v, b));
        }
        const message = !!b && rule.showInput && !!(rule.promptTitle || rule.prompt);
        v.promptbox.hidden = !message;
        if (message) {
            v.promptbox.textContent = "";
            if (rule.promptTitle) v.promptbox.appendChild(el("b", null, rule.promptTitle));
            if (rule.prompt) v.promptbox.appendChild(el("div", null, rule.prompt));
            Object.assign(v.promptbox.style, { left: `${b.x + Math.min(b.w, 40)}px`, top: `${b.y + b.h + 6}px` });
            pinTo(v.promptbox, pinOf(v, b));
        }
    }
    function openList() {
        if (!editable || busy || !view || !sel) return;
        const rule = listAt(sel.ar, sel.ac);
        if (!rule || rule.type !== "list" || !rule.inCellDropdown || !rule.values) return;
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
    // What a text typed into a cell is, as far as the rules are checked here: { n } a number
    // (a date or a time typed as Excel reads them: its serial), { text }, { bool }, or null: one
    // the editor works out (a formula, a date of an era): checked there.
    function typedValue(text) {
        const t = text.trim();
        if (text.startsWith("'")) return { text: text.slice(1) };
        if (/^[=@]/.test(t) || /^[+-][^\d.]/.test(t)) return null;
        if (/^(true|false)$/i.test(t)) return { bool: /^true$/i.test(t) };
        const plain = t.replace(/[,¥$\s]/g, "");
        const num = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)(%?)$/i.exec(plain);
        if (num && plain !== "") return { n: +num[1] / (num[2] ? 100 : 1) };
        const serial = (y, m, d) => {
            const ms = Date.UTC(y, m - 1, d);
            const back = new Date(ms);
            if (back.getUTCMonth() !== m - 1 || back.getUTCDate() !== d) return null;
            return ms / 86400000 + 25569;
        };
        const time = (s) => {
            const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm)?$/i.exec(s) || /^(\d{1,2})時(?:(\d{1,2})分)?(?:(\d{1,2})秒)?$/.exec(s);
            if (!m) return null;
            let h = +m[1];
            if (m[4]) h = (h % 12) + (/pm/i.test(m[4]) ? 12 : 0);
            return (h * 3600 + +(m[2] || 0) * 60 + +(m[3] || 0)) / 86400;
        };
        const tm = time(t);
        if (tm !== null) return { n: tm };
        const [datePart, timePart] = t.split(/\s+(?=\d{1,2}(?::|時))/);
        const year = new Date().getFullYear();
        const d =
            /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(datePart) ||
            /^(\d{4})年(\d{1,2})月(\d{1,2})日$/.exec(datePart) ||
            ((m) => m && [m[0], year, m[1], m[2]])(/^(\d{1,2})[/-](\d{1,2})$/.exec(datePart) || /^(\d{1,2})月(\d{1,2})日$/.exec(datePart));
        if (d) {
            const day = serial(+d[1], +d[2], +d[3]);
            const at = timePart ? time(timePart) : 0;
            if (day !== null && at !== null) return { n: day + at };
        }
        // (An era's date, a month's name, ...: the editor reads it.)
        if (/^(明治|大正|昭和|平成|令和|[MTSHRmtshr])\s*(\d+|元)/.test(t)) return null;
        return { text };
    }
    // A rule's operand worked out here: a number written in the rule (null: a formula, a
    // reference — the editor works those out).
    const ruleNumber = (f) => (/^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?\s*$/i.test(f || "") ? +f : null);
    // What the rule of the active cell says of the text typed into it: null (allowed), "editor"
    // (the editor checks it after it is entered), or "no".
    function judge(rule, text) {
        if (!rule || rule.type === "any" || text === "") return null;
        if (rule.type === "custom") return "editor";
        const v = typedValue(text);
        if (!v) return "editor";
        if (rule.type === "list") {
            const all = rule.values || [];
            const t = v.n !== undefined ? null : String(v.text !== undefined ? v.text : v.bool ? "TRUE" : "FALSE").toLowerCase();
            const hit = all.some((x) => (v.n !== undefined ? (x.trim() !== "" && +x.replace(/,/g, "") === v.n) || x === text.trim() : x.toLowerCase() === t));
            // (A list of cells shows them formatted: a number typed may be one of them as it is shown.)
            return hit ? null : rule.formula1.startsWith('"') ? "no" : "editor";
        }
        let x;
        if (rule.type === "textLength") x = (v.text !== undefined ? v.text : text.trim()).length;
        else if (v.n === undefined) return "no";
        else x = v.n;
        if (rule.type === "whole" && x !== Math.floor(x)) return "no";
        const a = ruleNumber(rule.formula1);
        const two = rule.operator === "between" || rule.operator === "notBetween";
        const b = two ? ruleNumber(rule.formula2) : 0;
        if (a === null || b === null) return "editor";
        const ok = {
            between: x >= a && x <= b,
            notBetween: x < a || x > b,
            equal: x === a,
            notEqual: x !== a,
            greaterThan: x > a,
            lessThan: x < a,
            greaterThanOrEqual: x >= a,
            lessThanOrEqual: x <= a,
        }[rule.operator];
        return ok === false ? "no" : null;
    }
    // The message of a value the rule does not allow, as Excel's: 再試行 / キャンセル (停止),
    // はい / いいえ / キャンセル (注意: "続けますか?"), OK / キャンセル (情報). `then(answer)`:
    // "enter" (the value is entered), "retry" (typed on), "cancel" (what was typed is undone).
    function askRefused(rule, then) {
        closeDialog();
        const style = rule.errorStyle || "stop";
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", `dlg alertdlg ${style}`));
        form.setAttribute("role", "alertdialog");
        form.appendChild(el("h2", null, rule.errorTitle || "Microsoft Excel"));
        const body = form.appendChild(el("div", "alertbody"));
        body.appendChild(el("span", "alerticon", style === "stop" ? "✕" : style === "warning" ? "!" : "i"));
        const text = body.appendChild(el("div", "msg", rule.error || "この値は、このセルに定義されているデータ入力規則の制限を満たしていません。"));
        if (style === "warning") text.appendChild(el("div", "ask", "続けますか?"));
        const buttons = form.appendChild(el("div", "buttons"));
        const answers = style === "stop" ? [["再試行", "retry"], ["キャンセル", "cancel"]] : style === "warning" ? [["はい", "enter"], ["いいえ", "retry"], ["キャンセル", "cancel"]] : [["OK", "enter"], ["キャンセル", "cancel"]];
        let first = null;
        for (const [label, answer] of answers) {
            const b = buttons.appendChild(el("button", first ? null : "ok", label));
            b.type = "button";
            b.addEventListener("click", () => {
                closeDialog();
                then(answer);
            });
            first = first || b;
        }
        form.cancel = () => {
            closeDialog();
            then("cancel");
        };
        form.addEventListener("keydown", (e) => {
            if (e.key === "Escape") {
                e.preventDefault();
                e.stopPropagation();
                form.cancel();
            }
        });
        form.addEventListener("submit", (e) => e.preventDefault());
        first.focus();
    }

    // ---- データ > データの入力規則

    // The rules of the sheet on the cells selected.
    function rulesOfSelection() {
        const list = view.sheet.lists || [];
        const all = allAreas();
        return list.filter((rule) => rule.ranges.some(([r1, c1, r2, c2]) => all.some((a) => r1 <= a.r2 && r2 >= a.r1 && c1 <= a.c2 && c2 >= a.c1)));
    }
    // A rule's ranges as an edit takes them (`A1:B5 D2`).
    const ruleRange = (rule) => rule.ranges.map(([r1, c1, r2, c2]) => (r1 === r2 && c1 === c2 ? address(r1, c1) : `${address(r1, c1)}:${address(r2, c2)}`)).join(" ");
    const VALIDATION_KINDS = [
        ["any", "すべての値"],
        ["whole", "整数"],
        ["decimal", "小数点数"],
        ["list", "リスト"],
        ["date", "日付"],
        ["time", "時刻"],
        ["textLength", "文字列 (長さ指定)"],
        ["custom", "ユーザー設定"],
    ];
    const VALIDATION_OPERATORS = [
        ["between", "次の値の間"],
        ["notBetween", "次の値の間以外"],
        ["equal", "次の値に等しい"],
        ["notEqual", "次の値に等しくない"],
        ["greaterThan", "次の値より大きい"],
        ["lessThan", "次の値より小さい"],
        ["greaterThanOrEqual", "次の値以上"],
        ["lessThanOrEqual", "次の値以下"],
    ];
    const VALIDATION_STYLES = [
        ["stop", "停止"],
        ["warning", "注意"],
        ["information", "情報"],
    ];
    const IME_MODES = [
        ["noControl", "コントロールなし"],
        ["on", "オン"],
        ["off", "オフ (英語モード)"],
        ["disabled", "無効"],
        ["hiragana", "ひらがな"],
        ["fullKatakana", "全角カタカナ"],
        ["halfKatakana", "半角カタカナ"],
        ["fullAlpha", "全角英数字"],
        ["halfAlpha", "半角英数字"],
    ];
    // A serial as the dialog shows a date (2026/11/24) or a time (9:00:00).
    function serialText(n, kind) {
        if (kind === "date") {
            const d = new Date(Math.round((Math.floor(n) - 25569) * 86400000));
            const date = `${d.getUTCFullYear()}/${d.getUTCMonth() + 1}/${d.getUTCDate()}`;
            return n === Math.floor(n) ? date : `${date} ${serialText(n - Math.floor(n), "time")}`;
        }
        const secs = Math.round((n - Math.floor(n)) * 86400);
        const two = (x) => String(x).padStart(2, "0");
        return `${Math.floor(secs / 3600) + (n >= 1 ? Math.floor(n) * 24 : 0)}:${two(Math.floor(secs / 60) % 60)}:${two(secs % 60)}`;
    }
    // A formula of the file as the dialog shows it, of a rule of `kind`.
    function operandText(f, kind) {
        if (!f) return "";
        if (kind === "list" && f.startsWith('"')) return f.slice(1, f.endsWith('"') ? -1 : undefined).replace(/""/g, '"');
        const n = ruleNumber(f);
        if (n !== null && kind !== "custom" && kind !== "list") return kind === "date" || kind === "time" ? serialText(n, kind) : String(n);
        return `=${f}`;
    }
    // What is typed into the dialog as the file's formula (`ok: false`, and why, when it is not one).
    function operandOf(text, kind) {
        const t = text.trim();
        if (t === "") return { ok: false, why: kind === "list" ? "元の値を入力してください。" : "値を入力してください。" };
        if (t.startsWith("=")) {
            const f = t.slice(1).trim();
            return f ? { ok: true, f } : { ok: false, why: "数式を入力してください。" };
        }
        if (kind === "list") return { ok: true, f: `"${t.replace(/"/g, '""')}"` };
        if (kind === "custom") return { ok: true, f: t };
        const v = typedValue(t);
        if (!v || v.n === undefined) return { ok: false, why: "入力した値は正しくありません。数値、日付、時刻、または = で始まる数式を入力してください。" };
        if ((kind === "whole" || kind === "textLength") && v.n !== Math.floor(v.n)) return { ok: false, why: "入力した値は正しくありません。整数を入力してください。" };
        return { ok: true, f: String(+v.n.toPrecision(15)), n: v.n };
    }
    // "データの入力規則": 設定 (the kind of values, the comparison, its values; a list's values,
    // a formula), 入力時メッセージ, エラー メッセージ, 日本語入力 — of the rule of the cells
    // selected; すべてクリア takes the rules off them.
    function askValidation() {
        if (!editable || busy || !sel || !view) return;
        if (entry) commit(null);
        const sheet = view.sheet;
        if (sheet.protection) {
            showError("保護されたシートでは、データの入力規則を変更できません。シートの保護を解除してください。");
            return;
        }
        const found = rulesOfSelection();
        if (found.length > 1) {
            askYes("選択範囲には複数の種類の入力値の設定が含まれています。現在の設定を消去して続けますか?", () => validationDialog(null, true));
            return;
        }
        const active = listAt(sel.ar, sel.ac);
        validationDialog(found[0] || null, false, active && active === found[0]);
    }
    // A question with OK and キャンセル.
    function askYes(message, then) {
        closeDialog();
        dialog = document.body.appendChild(el("div", "dlgback"));
        const form = dialog.appendChild(el("form", "dlg"));
        form.setAttribute("role", "dialog");
        form.appendChild(el("h2", null, "Microsoft Excel"));
        form.appendChild(el("div", "msg", message));
        dialogButtons(form);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            closeDialog();
            then();
        });
        form.querySelector("button.ok").focus();
    }
    function validationDialog(rule, cleared, sameOffered) {
        const sheet = view.sheet;
        const r = Object.assign(
            { type: "any", operator: "between", formula1: "", formula2: "", allowBlank: true, inCellDropdown: true, showInput: true, promptTitle: "", prompt: "", showError: true, errorStyle: "stop", errorTitle: "", error: "", ime: "noControl" },
            rule || {},
        );
        const form = openDialog("データの入力規則", "validdlg");
        const tabs = form.appendChild(el("div", "fmttabs tablist"));
        const pages = {};
        const showTab = (key) => {
            for (const t of tabs.children) t.classList.toggle("on", t.dataset.tab === key);
            for (const [k, p] of Object.entries(pages)) p.hidden = k !== key;
        };
        for (const [key, label] of [["settings", "設定"], ["input", "入力時メッセージ"], ["error", "エラー メッセージ"], ["ime", "日本語入力"]]) {
            const t = tabs.appendChild(el("button", null, label));
            t.type = "button";
            t.dataset.tab = key;
            t.addEventListener("click", () => showTab(key));
            pages[key] = form.appendChild(el("div", "validpage"));
        }
        const selectOf = (page, label, options, value) => {
            const line = page.appendChild(el("label", "row"));
            line.appendChild(el("span", "lbl", label));
            const s = line.appendChild(el("select"));
            for (const [v, text] of options) s.appendChild(el("option", null, text)).value = v;
            s.value = value;
            return s;
        };
        const textOf = (page, label, value, area) => {
            const line = page.appendChild(el("label", "col"));
            line.appendChild(el("span", "lbl", label));
            const input = line.appendChild(el(area ? "textarea" : "input"));
            if (!area) input.type = "text";
            input.spellcheck = false;
            input.value = value;
            input.addEventListener("input", () => input.setCustomValidity(""));
            return input;
        };

        // 設定
        const set = pages.settings;
        set.appendChild(el("div", "condprompt", "条件の設定"));
        const kind = selectOf(set, "入力値の種類:", VALIDATION_KINDS, r.type);
        const blank = check(set, "空白を無視する", r.allowBlank);
        const op = selectOf(set, "データ:", VALIDATION_OPERATORS, r.operator);
        const f1 = textOf(set, "", operandText(r.formula1, r.type));
        const f2 = textOf(set, "", operandText(r.formula2, r.type));
        const drop = check(set, "ドロップダウン リストから選択する", r.inCellDropdown);
        set.appendChild(el("hr"));
        const same = check(set, "同じ入力規則が設定されたすべてのセルに変更を適用する", false);
        same.disabled = !sameOffered;
        const lineOf = (x) => x.closest("label");
        const labels = {
            between: { "": ["最小値:", "最大値:"], date: ["次の日付から:", "次の日付まで:"], time: ["開始時刻:", "終了時刻:"] },
            one: { "": ["値:"], date: ["日付:"], time: ["時刻:"], textLength: ["長さ:"] },
            min: { "": ["最小値:"], date: ["開始日:"], time: ["開始時刻:"] },
            max: { "": ["最大値:"], date: ["終了日:"], time: ["終了時刻:"] },
        };
        const layout = () => {
            const k = kind.value;
            const compares = ["whole", "decimal", "date", "time", "textLength"].includes(k);
            const o = op.value;
            const two = compares && (o === "between" || o === "notBetween");
            lineOf(op).hidden = !compares;
            blank.disabled = k === "any";
            lineOf(f1).hidden = k === "any";
            lineOf(f2).hidden = !two;
            lineOf(drop).hidden = k !== "list";
            const group = two ? "between" : /^greaterThan/.test(o) ? "min" : /^lessThan/.test(o) ? "max" : "one";
            const names = labels[group][k] || labels[group][""];
            lineOf(f1).querySelector(".lbl").textContent = k === "list" ? "元の値:" : k === "custom" ? "数式:" : names[0];
            lineOf(f2).querySelector(".lbl").textContent = names[1] || "";
        };
        let lastKind = r.type;
        kind.addEventListener("change", () => {
            // (Another kind: the values typed for the other go, as Excel's.)
            if (kind.value !== lastKind) f1.value = f2.value = "";
            lastKind = kind.value;
            layout();
        });
        op.addEventListener("change", layout);
        layout();

        // 入力時メッセージ
        const inp = pages.input;
        const showInput = check(inp, "セルを選択したときに入力時メッセージを表示する", r.showInput);
        inp.appendChild(el("div", "condprompt", "セルを選択したときに表示するメッセージ:"));
        const promptTitle = textOf(inp, "タイトル:", r.promptTitle);
        promptTitle.maxLength = 32;
        const prompt = textOf(inp, "入力時メッセージ:", r.prompt, true);
        prompt.maxLength = 255;

        // エラー メッセージ
        const err = pages.error;
        const showErr = check(err, "無効なデータが入力されたらエラー メッセージを表示する", r.showError);
        err.appendChild(el("div", "condprompt", "無効なデータが入力されたときに表示するエラー メッセージ:"));
        const style = selectOf(err, "スタイル:", VALIDATION_STYLES, r.errorStyle);
        const errorTitle = textOf(err, "タイトル:", r.errorTitle);
        errorTitle.maxLength = 32;
        const error = textOf(err, "エラー メッセージ:", r.error, true);
        error.maxLength = 225;

        // 日本語入力
        const ime = pages.ime;
        ime.appendChild(el("div", "condprompt", "IME"));
        const imeMode = selectOf(ime, "日本語入力:", IME_MODES, r.ime);

        const buttons = dialogButtons(form);
        const clearAll = buttons.insertBefore(el("button", "clear", "すべてクリア(C)"), buttons.firstChild);
        clearAll.type = "button";
        clearAll.addEventListener("click", () => {
            kind.value = lastKind = "any";
            op.value = "between";
            f1.value = f2.value = "";
            blank.checked = drop.checked = showInput.checked = showErr.checked = true;
            promptTitle.value = prompt.value = errorTitle.value = error.value = "";
            style.value = "stop";
            imeMode.value = "noControl";
            layout();
        });
        const fail = (input, why) => {
            showTab("settings");
            input.setCustomValidity(why);
            input.reportValidity();
        };
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            const range = same.checked && rule ? ruleRange(rule) : condRange();
            const k = kind.value;
            const edit = { op: "validation", sheet: sheet.name, range };
            const nothing = k === "any" && !promptTitle.value && !prompt.value && !errorTitle.value && !error.value && imeMode.value === "noControl";
            if (nothing) {
                closeDialog();
                // (Nothing set: the cells have no rule — nothing to do when they had none.)
                if (rule || cleared) postEdit([Object.assign(edit, { clear: true })]);
                return;
            }
            Object.assign(edit, {
                type: k,
                allowBlank: blank.checked,
                inCellDropdown: drop.checked,
                showInput: showInput.checked,
                promptTitle: promptTitle.value,
                prompt: prompt.value,
                showError: showErr.checked,
                errorStyle: style.value,
                errorTitle: errorTitle.value,
                error: error.value,
                ime: imeMode.value,
            });
            if (k !== "any") {
                const a = operandOf(f1.value, k);
                if (!a.ok) return fail(f1, a.why);
                edit.formula1 = a.f;
                if (["whole", "decimal", "date", "time", "textLength"].includes(k)) {
                    edit.operator = op.value;
                    if (op.value === "between" || op.value === "notBetween") {
                        const b = operandOf(f2.value, k);
                        if (!b.ok) return fail(f2, b.why);
                        if (a.n !== undefined && b.n !== undefined && b.n < a.n) return fail(f2, "最大値は最小値以上でなければなりません。");
                        edit.formula2 = b.f;
                    }
                }
            }
            closeDialog();
            postEdit([edit]);
        });
        showTab("settings");
        kind.focus();
    }

    // "無効データのマーク": the cells with values their rules do not allow, circled in red (as
    // Excel, up to 255), until "入力規則マークのクリア".
    function markInvalid(on) {
        remember({ invalidMarks: on });
        if (on && view && !(view.sheet.invalid || []).length && (view.sheet.lists || []).length) showError("無効なデータはありません。");
        drawBoxes();
    }
    function placeInvalid() {
        const v = view;
        if (!v) return;
        for (const m of v.wrap.querySelectorAll(".invalidmark")) m.remove();
        if (!saved.invalidMarks) return;
        const marks = v.selbox.parentNode;
        for (const [r, c] of (v.sheet.invalid || []).slice(0, 255)) {
            const o = ownerOf(r, c) || [r, c, r, c];
            const b = boxOf(o[0], o[1], o[2], o[3]);
            if (!b) continue;
            const m = marks.insertBefore(el("div", "invalidmark"), v.selbox);
            Object.assign(m.style, { left: `${b.x - 4}px`, top: `${b.y - 3}px`, width: `${b.w + 8}px`, height: `${b.h + 6}px` });
            pinTo(m, pinOf(v, b));
        }
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
        const tabList = tabs.appendChild(el("div", "tablist"));
        const tab = (label, on, act) => {
            const b = tabList.appendChild(el("button", on ? "on" : null, label));
            b.type = "button";
            b.addEventListener("click", act);
        };
        tab("検索", !finder.replace, () => openFind(false));
        if (editable) tab("置換", finder.replace, () => openFind(true));
        const close = tabs.appendChild(el("button", "x iconbtn", "×"));
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
        const input = entryInput();
        const ctrl = e.ctrlKey || e.metaKey;
        // (An input method takes its own keys.)
        if (e.isComposing || e.keyCode === 229) return;
        // The characters' formats: Ctrl+B, I, U, 5 (取り消し線), Ctrl+1 (セルの書式設定: フォント).
        const k = e.key.length === 1 ? e.key.toLowerCase() : e.key;
        if (ctrl && !e.altKey && !e.shiftKey && charTarget() && (["b", "i", "u", "5", "1"].includes(k) || e.code === "Digit5" || e.code === "Digit1" || e.code === "Numpad1")) {
            e.preventDefault();
            if (k === "1" || e.code === "Digit1" || e.code === "Numpad1") askCharFont();
            else toggleFormat({ b: "bold", i: "italic", u: "underline" }[k] || "strike");
            return;
        }
        if (entry.rich && ctrl && !e.altKey && (k === "z" || k === "y")) {
            e.preventDefault();
            richHistory(k === "z" && !e.shiftKey);
            return;
        }
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
            // (As キャンセル: a dialog opened from another goes back to it; a menu over it
            // closes alone.)
            if (e.key === "Escape" && menu) {
                e.preventDefault();
                closeMenu();
            } else if (e.key === "Escape") {
                e.preventDefault();
                const form = subDialog ? subDialog.querySelector("form") : dialog.querySelector("form");
                ((form && form.cancel) || closeDialog)();
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
        // Ctrl+F3: 名前の管理.
        if (view && editable && ctrl && key === "F3" && !e.altKey && !e.shiftKey) {
            e.preventDefault();
            manageNames();
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
            editSelected((o) => turnChange(o, key === "ArrowRight" ? 15 : -15));
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
            // A character typed: into the shape's text (the input method's too).
            if (!ctrl && editable && !busy && obj.textBox && movable(obj) && (key.length === 1 || key === "Process" || e.keyCode === 229)) {
                editShapeText(obj);
                return;
            }
            if (!ctrl || moves[key]) {
                e.preventDefault();
                if (key === "Delete" || key === "Backspace") deleteObject(obj);
                else if (moves[key]) nudgeObjects(moves[key][0], moves[key][1], ctrl);
                else if (key === "F2" || key === "Enter") editObjectText(obj);
                return;
            }
        }
        if (!view || !sel || !onGrid()) return;
        // A character typed while the box of the active cell has lost the focus (the page drawn
        // again in a frame that was not focused, an input method started while an edit was
        // posted): the box takes it, so that it is typed into the cell, not lost.
        if (editable && !busy && !ctrl && !e.altKey && document.activeElement !== ed && (key.length === 1 || key === "Process" || e.keyCode === 229)) {
            ed.focus({ preventScroll: true });
            return;
        }
        const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
        let done = true;
        if (e.altKey && ctrl && (key === "v" || key === "V")) askPasteSpecial();
        else if (e.altKey && key === "ArrowDown" && (listAt(sel.ar, sel.ac) || {}).type === "list") openList();
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
        else if (ctrl && !e.shiftKey && (key === "1" || e.code === "Digit1" || e.code === "Numpad1")) openCellPane();
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
        // (Nor while a control of the run mode is typed into.)
        if (!entry && !dialog && !runAskBox && !(e.target.closest && e.target.closest("button, input, select, label, textarea, .menu, #side"))) focusGrid();
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

    const total = Object.values(changes).reduce((n, list) => n + list.length, 0);
    if (data.diff) {
        $("summary").textContent = total || data.diff.structure.length ? `${data.diff.before} からの変更: セル ${total} 個` : `${data.diff.before} からの変更はありません`;
    }
    // ---- the tabs of the sheets: a click shows one. Where the workbook is edited, their menu
    // inserts, removes, renames and copies sheets and colors a tab, a double click on one
    // renames it there, and ＋ inserts a sheet, as Excel's tabs do. The editor makes the change
    // (in the project's workbook, with the files of objects/) and draws the page again.

    // What is wrong with `name` for a sheet's tab (`but`: the sheet that has it now).
    function sheetNameProblem(name, but) {
        if (!name) return "シート名を入力してください。";
        if ([...name].length > 31 || /[:\\/?*[\]]/.test(name) || name.startsWith("'") || name.endsWith("'")) return "シート名は 31 文字までで、: \\ / ? * [ ] は使えません（前後を ' にもできません）。";
        if (data.sheets.some((s) => s !== but && s.name.toLowerCase() === name.toLowerCase())) return `「${name}」という名前のシートは既にあります。`;
        return "";
    }
    function codeNameProblem(code) {
        if (!code) return "";
        if (!/^\p{L}[\p{L}\p{N}_]*$/u.test(code) || [...code].length > 31) return "コード名は、文字で始まり、文字・数字・_ だけの 31 文字までの名前にしてください。";
        if (data.sheets.some((s) => (s.codeName || "").toLowerCase() === code.toLowerCase())) return `コード名 ${code} は、ほかのシートが使っています。`;
        return "";
    }
    const takenName = (name) => data.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase() || (s.codeName || "").toLowerCase() === name.toLowerCase());
    // `m`: { action: "add" | "copy" | "remove" | "rename", sheet, name, codeName, color };
    // `show`: the sheet shown when the page is drawn again.
    function postSheet(m, show) {
        if (!editable || busy) return;
        if (runMode) {
            showError(RUN_ONLY);
            return;
        }
        setClip(null);
        setBusy(true);
        if (show) remember({ sheet: show });
        vscode.postMessage({ type: "sheet", ...m });
    }
    // "挿入": a new sheet after `after`, its name, code name and color asked at once.
    function askNewSheet(after) {
        if (!editable || busy) return;
        if (entry) commit(null);
        const form = openDialog("シートの挿入", "sheetdlg");
        let n = data.sheets.length + 1;
        while (takenName(`Sheet${n}`)) n++;
        const name = field(form, "シート名");
        name.value = `Sheet${n}`;
        name.maxLength = 31;
        const code = field(form, "コード名");
        code.placeholder = "空欄なら自動で付けます";
        code.maxLength = 31;
        const line = form.appendChild(el("div", "tabcolor"));
        line.appendChild(el("span", "lbl", "見出しの色"));
        const swatches = line.appendChild(el("div", "swatches"));
        let color = null;
        const swatch = (c, title) => {
            const b = swatches.appendChild(el("button", c ? "swatch" : "swatch none"));
            b.type = "button";
            b.title = title;
            if (c) b.style.background = c;
            b.classList.toggle("on", c === color);
            b.addEventListener("click", () => {
                color = c;
                for (const x of swatches.children) x.classList.toggle("on", x === b);
            });
        };
        swatch(null, "色なし");
        for (const c of PALETTE) swatch(c.toUpperCase(), c.toUpperCase());
        form.appendChild(el("div", "hint", "コード名は、プログラムからシートを指す名前です（Sheet1.Range(\"A1\") の Sheet1、objects/Sheet1.ras）。"));
        dialogButtons(form);
        const validate = () => {
            name.setCustomValidity(sheetNameProblem(name.value.trim()));
            code.setCustomValidity(codeNameProblem(code.value.trim()));
        };
        name.addEventListener("input", validate);
        code.addEventListener("input", validate);
        form.addEventListener("submit", (e) => {
            e.preventDefault();
            validate();
            if (!form.reportValidity()) return;
            closeDialog();
            const n = name.value.trim();
            postSheet({ action: "add", sheet: after.name, name: n, codeName: code.value.trim(), color }, n);
        });
        name.focus();
        name.select();
    }
    // "名前の変更" (a double click on the tab): the name edited on the tab; Enter (or a click
    // elsewhere) changes it, Esc does not.
    function renameTab(b) {
        if (!editable || busy || b.querySelector("input")) return;
        if (entry) commit(null);
        closeMenu();
        const sheet = b.sheet;
        const input = el("input", "tabname");
        input.type = "text";
        input.value = sheet.name;
        input.maxLength = 31;
        input.spellcheck = false;
        const fit = () => (input.style.width = `${Math.max(3, [...input.value].length + 1)}em`);
        fit();
        b.textContent = "";
        b.appendChild(input);
        let done = false;
        const finish = (change) => {
            if (done) return;
            const name = input.value.trim();
            const problem = change && name !== sheet.name ? sheetNameProblem(name, sheet) : "";
            if (problem && change === "enter") {
                input.setCustomValidity(problem);
                input.reportValidity();
                return;
            }
            done = true;
            labelTab(b);
            if (problem) showError(problem);
            else if (change && name !== sheet.name) {
                b.firstChild.textContent = name;
                postSheet({ action: "rename", sheet: sheet.name, name }, name);
            }
            focusGrid();
        };
        input.addEventListener("keydown", (e) => {
            e.stopPropagation();
            if (e.isComposing || e.keyCode === 229) return;
            if (e.key === "Enter") {
                e.preventDefault();
                finish("enter");
            } else if (e.key === "Escape") {
                e.preventDefault();
                finish(null);
            }
        });
        input.addEventListener("input", () => {
            input.setCustomValidity("");
            fit();
        });
        input.addEventListener("blur", () => finish("blur"));
        for (const type of ["click", "dblclick", "mousedown", "contextmenu"]) input.addEventListener(type, (e) => e.stopPropagation());
        input.focus();
        input.select();
    }
    function removeSheet(sheet) {
        if (data.sheets.filter((s) => s !== sheet && s.visible === "visible").length === 0) {
            showError("表示されているシートがほかに無いため、削除できません。");
            return;
        }
        const others = data.sheets.filter((s) => s !== sheet);
        const next = others[Math.min(data.sheets.indexOf(sheet), others.length - 1)];
        postSheet({ action: "remove", sheet: sheet.name, codeName: sheet.codeName || "" }, next && next.name);
    }
    function copySheet(sheet) {
        let n = 2;
        const base = (k) => `${sheet.name.slice(0, 31 - ` (${k})`.length)} (${k})`;
        while (takenName(base(n))) n++;
        postSheet({ action: "copy", sheet: sheet.name }, base(n));
    }
    function openTabMenu(x, y, b) {
        const sheet = b.sheet;
        showMenu(x, y, [
            ["挿入…", () => askNewSheet(sheet), ""],
            ["削除", () => removeSheet(sheet), ""],
            ["名前の変更", () => renameTab(b), ""],
            ["コピーを作成", () => copySheet(sheet), ""],
            ["移動またはコピー…", () => askMoveOrCopy(sheet), ""],
            null,
            sheet.protection ? ["シート保護の解除…", () => askUnprotect(sheet), ""] : ["シートの保護…", () => askProtect(sheet), ""],
            ["シート見出しの色…", () => openColors(b, "tab", "色なし", (color) => postEdit([{ op: "tabColor", sheet: sheet.name, color }])), ""],
            null,
            ["非表示", () => hideSheet(sheet), ""],
            ...(data.sheets.some((s) => s.visible !== "visible") ? [["再表示…", askUnhideSheets, ""]] : []),
        ]);
    }
    // A tab colored: filled with its color made a little lighter (the name in black or white,
    // whichever reads), the tab shown with a tint of it (underlined, as every tab shown, in the accent).
    function paintTab(b, color) {
        const m = /^#?([0-9a-f]{6})$/i.exec(color || "");
        if (!m) return;
        const rgb = [0, 2, 4].map((i) => parseInt(m[1].slice(i, i + 2), 16));
        const mix = (t) => "#" + rgb.map((c) => Math.round(c + (255 - c) * t).toString(16).padStart(2, "0")).join("");
        const light = rgb.map((c) => c + (255 - c) * 0.25);
        const lum = light.map((c) => ((c /= 255) <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)).reduce((a, c, i) => a + c * [0.2126, 0.7152, 0.0722][i], 0);
        b.style.setProperty("--tab-bg", mix(0.25));
        b.style.setProperty("--tab-ink", lum > 0.4 ? "#222" : "#fff");
        b.style.setProperty("--tab-tint", mix(0.88));
        b.classList.add("colored");
    }
    // The tab's name, and the number of the cells changed.
    function labelTab(b) {
        const sheet = b.sheet;
        b.textContent = "";
        b.appendChild(el("span", null, sheet.name));
        if (sheet.visible !== "visible") b.appendChild(document.createTextNode("（非表示）"));
        const n = (changes[sheet.name] || []).length;
        if (n) b.appendChild(el("span", "badge", String(n)));
    }
    for (const sheet of data.sheets) {
        const b = el("button", sheet.visible === "visible" ? null : "hidden");
        b.type = "button";
        b.sheet = sheet;
        labelTab(b);
        if (sheet.tabColor) {
            paintTab(b, sheet.tabColor);
        }
        b.addEventListener("click", () => {
            if (!b.querySelector("input")) draw(sheet);
        });
        if (editable) {
            b.title = "ダブルクリックで名前を変更";
            b.dataset.vscodeContext = JSON.stringify({ preventDefaultContextMenuItems: true });
            b.addEventListener("dblclick", (e) => {
                e.preventDefault();
                renameTab(b);
            });
            b.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (busy) return;
                if (view && view.sheet !== sheet) draw(sheet);
                openTabMenu(e.clientX, e.clientY, b);
            });
        }
        $("tabs").appendChild(b);
    }
    if (editable) {
        // Excel's ⊕ right after the last tab (kept in view while the tabs scroll).
        const add = $("tabs").appendChild(el("button", "addsheet"));
        add.type = "button";
        add.title = "新しいシート";
        add.setAttribute("aria-label", "新しいシート");
        add.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" aria-hidden="true"><circle cx="8" cy="8" r="6.5"/><path d="M8 4.5v7M4.5 8h7"/></svg>';
        add.addEventListener("click", () => {
            const after = (view && view.sheet) || data.sheets[data.sheets.length - 1];
            if (after) askNewSheet(after);
        });
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
        if (view && !runMode) pickObject(view);
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
