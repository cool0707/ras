// Excel's charts on the grid of `ras book view`, drawn as SVG from what `chart.rs` reads of them:
// column and bar (clustered, stacked, 100%), line, area, pie and doughnut, scatter and bubble,
// radar; their title, legend, axes with gridlines, data labels, a secondary value axis. The
// look is Excel's since 2013 (gray text, light gridlines, no axis line for values).
"use strict";
window.rasChart = (function () {
    const TEXT = "#595959";
    const LINE = "#d9d9d9";
    const FONT = '"Yu Gothic", "Yu Gothic UI", Meiryo, sans-serif';
    const ctx = document.createElement("canvas").getContext("2d");
    const measure = (text, px, bold) => {
        ctx.font = `${bold ? "bold " : ""}${px}px ${FONT}`;
        return ctx.measureText(String(text)).width;
    };
    const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
    const num = (v) => (Math.round(v * 100) / 100).toString();

    // ---- number formats (the common ones: decimals, thousands, percent, literal text, dates)

    const pad = (n, w) => String(n).padStart(w, "0");
    function dateText(serial, code) {
        const ms = Math.round((serial - 25569) * 86400000);
        const d = new Date(ms);
        const parts = { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), h: d.getUTCHours(), n: d.getUTCMinutes(), s: d.getUTCSeconds() };
        let out = "";
        let lastH = false;
        const re = /(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|"[^"]*"|\\.|.)/gi;
        for (const [tok] of code.matchAll(re)) {
            const t = tok.toLowerCase();
            if (t === "yyyy") out += parts.y;
            else if (t === "yy") out += pad(parts.y % 100, 2);
            else if (t === "dd") out += pad(parts.d, 2);
            else if (t === "d") out += parts.d;
            else if (t === "ddd" || t === "dddd") out += "日月火水木金土"[d.getUTCDay()];
            else if (t === "hh" || t === "h") {
                out += t === "hh" ? pad(parts.h, 2) : parts.h;
                lastH = true;
                continue;
            } else if (t === "mm" || t === "m") out += lastH ? pad(parts.n, 2) : t === "mm" ? pad(parts.m, 2) : parts.m;
            else if (t === "mmm" || t === "mmmm") out += parts.m + "月";
            else if (t === "ss" || t === "s") out += pad(parts.s, 2);
            else if (tok.startsWith('"')) out += tok.slice(1, -1);
            else if (tok.startsWith("\\")) out += tok.slice(1);
            else out += tok;
            if (t !== "") lastH = lastH && /^[:\s]$/.test(tok);
        }
        return out;
    }
    function format(v, code) {
        if (v === null || v === undefined || !isFinite(v)) return "";
        code = (code || "").trim();
        if (!code || /^general$/i.test(code)) return String(Number(v.toPrecision(10)));
        const sections = code.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/);
        let sec = sections[0];
        let neg = false;
        if (v < 0 && sections.length > 1) {
            sec = sections[1];
            v = -v;
        } else if (v === 0 && sections.length > 2) sec = sections[2];
        else if (v < 0) neg = true;
        // [Red], [$¥-411], [>100]: the currency's symbol only.
        sec = sec.replace(/\[\$([^\]-]*)[^\]]*\]/g, '"$1"').replace(/\[[^\]]*\]/g, "");
        const bare = sec.replace(/"[^"]*"|\\./g, "");
        if (/[ymdhs]/i.test(bare) && !/[#0?]/.test(bare)) return dateText(v, sec);
        if (/%/.test(bare)) v *= 100;
        const m = /[#0?,.]*[#0?][#0?,.]*/.exec(sec.replace(/"[^"]*"/g, (s) => "\u0000".repeat(s.length)));
        if (!m) return sec.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1");
        const pat = m[0];
        const dot = pat.indexOf(".");
        const decimals = dot < 0 ? 0 : (pat.slice(dot + 1).match(/[#0?]/g) || []).length;
        const thousands = /[#0?],[#0?]/.test(pat);
        // Commas at the end divide by a thousand each.
        const scale = (/(,+)(?:\.|$)/.exec(pat.replace(/[^#0?,.]/g, "")) || ["", ""])[1].length;
        v /= Math.pow(1000, scale);
        let s = v.toFixed(decimals);
        if (thousands) {
            const [i, f] = s.split(".");
            s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (f !== undefined ? "." + f : "");
        }
        const lit = (t) => t.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1").replace(/[_*]./g, "");
        return (neg ? "-" : "") + lit(sec.slice(0, m.index)) + s + lit(sec.slice(m.index + pat.length));
    }

    // ---- scales

    // An axis from lo to hi: from 0 unless the values are close together (Excel: their range is
    // less than a sixth of the largest).
    function niceScale(lo, hi, length, axis, zero, spacing = 36) {
        if (zero && lo >= 0 && hi - lo >= hi / 6) lo = 0;
        if (zero && hi <= 0 && hi - lo >= -lo / 6) hi = 0;
        if (axis && axis.min !== undefined) lo = axis.min;
        if (axis && axis.max !== undefined) hi = axis.max;
        if (!(hi > lo)) {
            hi = lo + (lo === 0 ? 1 : Math.abs(lo));
        }
        let step = axis && axis.majorUnit;
        if (!(step > 0)) {
            const target = Math.max(2, Math.min(10, Math.floor(length / spacing)));
            const raw = (hi - lo) / target;
            const mag = Math.pow(10, Math.floor(Math.log10(raw)));
            step = [1, 2, 5, 10].map((k) => k * mag).find((k) => k >= raw) || 10 * mag;
        }
        if (!(axis && axis.min !== undefined)) lo = Math.floor(lo / step + 1e-9) * step;
        if (!(axis && axis.max !== undefined)) hi = Math.ceil(hi / step - 1e-9) * step;
        if (hi <= lo) hi = lo + step;
        const ticks = [];
        for (let t = lo; t <= hi + step * 1e-6 && ticks.length < 100; t += step) ticks.push(Math.abs(t) < step * 1e-9 ? 0 : t);
        return { lo, hi, ticks };
    }

    // A smooth line through points (Catmull-Rom as Bézier curves).
    function smoothPath(pts) {
        let d = `M${num(pts[0][0])},${num(pts[0][1])}`;
        for (let i = 0; i + 1 < pts.length; i++) {
            const p0 = pts[i - 1] || pts[i],
                p1 = pts[i],
                p2 = pts[i + 1],
                p3 = pts[i + 2] || p2;
            const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
            const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
            d += ` C${num(c1[0])},${num(c1[1])} ${num(c2[0])},${num(c2[1])} ${num(p2[0])},${num(p2[1])}`;
        }
        return d;
    }

    function draw(box, spec) {
        const W = box.clientWidth,
            H = box.clientHeight;
        if (!(W > 20 && H > 20)) return;
        const px = (spec.textSize || 9) * (4 / 3);
        const out = [];
        const text = (x, y, s, o = {}) =>
            out.push(
                `<text x="${num(x)}" y="${num(y)}" font-size="${num(o.size || px)}"${o.anchor ? ` text-anchor="${o.anchor}"` : ""}${o.bold ? ' font-weight="bold"' : ""}${o.rotate ? ` transform="rotate(${o.rotate} ${num(x)} ${num(y)})"` : ""} fill="${o.color || TEXT}" dominant-baseline="${o.base || "middle"}">${esc(s)}</text>`,
            );
        const groups = spec.groups || [];
        const allSeries = groups.flatMap((g) => g.series.map((s) => ({ s, g })));
        const autoColor = (i) => (spec.autoColors || [])[i % Math.max(1, (spec.autoColors || []).length)] || "#4472c4";
        const pieGroup = groups.find((g) => g.kind === "pie" || g.kind === "doughnut");
        // Colors of the points of a series drawn with varied colors (a pie, a single series).
        const pointColor = (s, g, i) => {
            const p = (s.points || []).find((x) => x.idx === i);
            if (p && p.color) return p.color;
            return g.varyColors && (g.kind === "pie" || g.kind === "doughnut" || g.series.length === 1) ? autoColor(i) : s.color || autoColor(0);
        };

        // The title: its text, or the name of the only series (as Excel titles it).
        let title = typeof spec.title === "string" ? spec.title : null;
        if (spec.title === true || (!spec.title && !spec.autoTitleDeleted && allSeries.length === 1)) title = allSeries.length === 1 ? allSeries[0].s.name : "グラフ タイトル";
        let top = 8,
            left = 8,
            right = W - 8,
            bottom = H - 8;
        if (title) {
            const size = (spec.titleSize || 14) * (4 / 3);
            const lines = String(title).split("\n");
            lines.forEach((line, i) => text(W / 2, top + size * (i + 0.6), line, { size, anchor: "middle" }));
            top += size * lines.length + 6;
        }

        // The legend.
        const vary = pieGroup && pieGroup.varyColors !== false;
        const entries = vary
            ? (pieGroup.series[0] ? (pieGroup.series[0].cat || pieGroup.series[0].val.map((_, i) => String(i + 1))).map((c, i) => ({ name: c, color: pointColor(pieGroup.series[0], pieGroup, i), kind: "box" })) : [])
            : allSeries.map(({ s, g }, i) => ({
                  name: s.name,
                  color: s.color || autoColor(i),
                  kind: g.kind === "line" || g.kind === "radar" || (g.kind === "scatter" && !s.noFill) ? "line" : "box",
                  // A series without fill (the start of a Gantt chart's bars) is not in the legend.
                  hidden: s.noFill && (g.kind === "bar" || g.kind === "area"),
              }));
        const shown = entries.filter((e) => !e.hidden);
        if (spec.legend && shown.length) {
            const pos = spec.legend;
            const item = (e) => measure(e.name, px) + 18;
            if (pos === "r" || pos === "l" || pos === "tr") {
                const w = Math.min(W * 0.4, Math.max(...shown.map(item)) + 8);
                const x0 = pos === "l" ? left : right - w;
                const total = shown.length * (px + 6);
                let y = pos === "tr" ? top + 4 : Math.max(top, (top + bottom) / 2 - total / 2);
                for (const e of shown) {
                    legendKey(out, x0 + 4, y + px / 2, e);
                    text(x0 + 20, y + px / 2, e.name);
                    y += px + 6;
                }
                if (pos === "l") left += w + 4;
                else right -= w + 4;
            } else {
                // Above or below: in rows, centered.
                const rows = [[]];
                let w = 0;
                for (const e of shown) {
                    const iw = item(e) + 10;
                    if (w + iw > W - 16 && rows[rows.length - 1].length) {
                        rows.push([]);
                        w = 0;
                    }
                    rows[rows.length - 1].push(e);
                    w += iw;
                }
                const h = rows.length * (px + 6);
                let y = pos === "t" ? top : bottom - h;
                for (const row of rows) {
                    const rw = row.reduce((a, e) => a + item(e) + 10, -10);
                    let x = W / 2 - rw / 2;
                    for (const e of row) {
                        legendKey(out, x, y + px / 2, e);
                        text(x + 16, y + px / 2, e.name);
                        x += item(e) + 10;
                    }
                    y += px + 6;
                }
                if (pos === "t") top += h + 4;
                else bottom -= h + 4;
            }
        }
        const area = { left, top, right, bottom };
        if (pieGroup) drawPie(out, area, pieGroup, pointColor, text, px);
        else if (groups[0].kind === "radar") drawRadar(out, area, groups.filter((g) => g.kind === "radar"), spec, text, px, autoColor);
        else drawPlot(out, area, groups, spec, text, px, pointColor, autoColor);

        const fill = spec.fill === "none" ? "none" : spec.fill || "#fff";
        const border = spec.border === "none" ? "none" : spec.border || LINE;
        box.innerHTML = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="position:absolute;inset:0;font-family:${esc(FONT)}"><rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" fill="${fill}" stroke="${border}"/>${out.join("")}</svg>`;
    }

    function legendKey(out, x, y, e) {
        if (e.kind === "line") {
            out.push(`<line x1="${num(x)}" y1="${num(y)}" x2="${num(x + 14)}" y2="${num(y)}" stroke="${e.color}" stroke-width="2"/>`);
        } else {
            out.push(`<rect x="${num(x + 3)}" y="${num(y - 4)}" width="8" height="8" fill="${e.color}"/>`);
        }
    }

    function drawPie(out, a, g, pointColor, text, px) {
        const s = g.series[0];
        if (!s) return;
        const vals = (s.val || []).map((v) => (v > 0 ? v : 0));
        const total = vals.reduce((x, y) => x + y, 0);
        if (!(total > 0)) return;
        const cx = (a.left + a.right) / 2,
            cy = (a.top + a.bottom) / 2;
        const r = Math.max(4, Math.min(a.right - a.left, a.bottom - a.top) / 2 - 6);
        const hole = g.kind === "doughnut" ? r * Math.min(90, Math.max(10, +(g.holeSize || 50))) / 100 : 0;
        let ang = ((+(g.firstSliceAng || 0) - 90) * Math.PI) / 180;
        const at = (rr, t) => [cx + rr * Math.cos(t), cy + rr * Math.sin(t)];
        vals.forEach((v, i) => {
            if (!(v > 0)) return;
            const sweep = (v / total) * Math.PI * 2;
            const a2 = ang + sweep;
            const large = sweep > Math.PI ? 1 : 0;
            const color = pointColor(s, g, i);
            if (sweep >= Math.PI * 2 - 1e-9) {
                out.push(`<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(r)}" fill="${color}"/>`);
                if (hole) out.push(`<circle cx="${num(cx)}" cy="${num(cy)}" r="${num(hole)}" fill="#fff"/>`);
            } else {
                const [x1, y1] = at(r, ang),
                    [x2, y2] = at(r, a2);
                let d = `M${num(x1)},${num(y1)} A${num(r)},${num(r)} 0 ${large} 1 ${num(x2)},${num(y2)}`;
                if (hole) {
                    const [x3, y3] = at(hole, a2),
                        [x4, y4] = at(hole, ang);
                    d += ` L${num(x3)},${num(y3)} A${num(hole)},${num(hole)} 0 ${large} 0 ${num(x4)},${num(y4)} Z`;
                } else d += ` L${num(cx)},${num(cy)} Z`;
                out.push(`<path d="${d}" fill="${color}" stroke="#fff" stroke-width="1"/>`);
            }
            if (s.labels) {
                const parts = [];
                if (s.labels.cat && s.cat) parts.push(s.cat[i]);
                if (s.labels.val) parts.push((s.valText || [])[i] || String(v));
                if (s.labels.percent) parts.push(Math.round((v / total) * 100) + "%");
                const [lx, ly] = at(hole ? (r + hole) / 2 : r * 0.65, ang + sweep / 2);
                text(lx, ly, parts.join(" "), { anchor: "middle", color: "#fff" });
            }
            ang = a2;
        });
    }

    function drawRadar(out, a, groups, spec, text, px, autoColor) {
        const series = groups.flatMap((g) => g.series.map((s) => ({ s, g })));
        const cats = (series[0] && (series[0].s.cat || series[0].s.val.map((_, i) => String(i + 1)))) || [];
        const n = cats.length;
        if (n < 3) return;
        const cx = (a.left + a.right) / 2,
            cy = (a.top + a.bottom) / 2 + px / 2;
        const r = Math.max(4, Math.min(a.right - a.left, a.bottom - a.top) / 2 - px * 1.8);
        const values = series.flatMap(({ s }) => (s.val || []).filter((v) => v !== null));
        const sc = niceScale(Math.min(0, ...values), Math.max(...values, 0), r, (spec.axes || []).find((x) => x.kind === "val"), true);
        const at = (i, v) => {
            const t = -Math.PI / 2 + (i * 2 * Math.PI) / n;
            const rr = ((v - sc.lo) / (sc.hi - sc.lo)) * r;
            return [cx + rr * Math.cos(t), cy + rr * Math.sin(t)];
        };
        for (const t of sc.ticks) {
            const pts = cats.map((_, i) => at(i, t).map(num).join(","));
            out.push(`<polygon points="${pts.join(" ")}" fill="none" stroke="${LINE}"/>`);
        }
        cats.forEach((c, i) => {
            const [x, y] = at(i, sc.hi);
            out.push(`<line x1="${num(cx)}" y1="${num(cy)}" x2="${num(x)}" y2="${num(y)}" stroke="${LINE}"/>`);
            const [lx, ly] = at(i, sc.hi + (sc.hi - sc.lo) * 0.12);
            text(lx, ly, c, { anchor: Math.abs(lx - cx) < 4 ? "middle" : lx > cx ? "start" : "end" });
        });
        series.forEach(({ s, g }, k) => {
            const color = s.color || autoColor(k);
            const pts = (s.val || []).map((v, i) => at(i, v || 0));
            const filled = g.radarStyle === "filled";
            out.push(`<polygon points="${pts.map((p) => p.map(num).join(",")).join(" ")}" fill="${filled ? color : "none"}"${filled ? ' fill-opacity="0.6"' : ""} stroke="${color}" stroke-width="2"/>`);
            if (s.marker && !filled) for (const [x, y] of pts) out.push(`<circle cx="${num(x)}" cy="${num(y)}" r="3" fill="${color}"/>`);
        });
    }

    function drawPlot(out, a, groups, spec, text, px, pointColor, autoColor) {
        const axes = spec.axes || [];
        const axis = (id) => axes.find((x) => x.id === id) || {};
        const first = groups[0];
        const xy = first.kind === "scatter" || first.kind === "bubble";
        const horizontal = first.kind === "bar" && first.barDir === "bar";
        const catAxis = axis(first.axes[0]);
        const valOf = (g) => axis(g.axes[1]);
        const primaryVal = valOf(first);
        const secondary = groups.filter((g) => g.axes[1] && g.axes[1] !== first.axes[1] && !xy);
        // Categories: of the first series that has them, else 1, 2, ...
        const lengthOf = (s) => Math.max((s.val || s.y || []).length, (s.cat || s.x || []).length);
        const count = Math.max(1, ...groups.flatMap((g) => g.series.map(lengthOf)));
        const withCats = groups.flatMap((g) => g.series).find((s) => s.cat && s.cat.length);
        const cats = withCats ? withCats.cat : Array.from({ length: count }, (_, i) => String(i + 1));

        // The range of values of the groups on an axis (stacked: sums; 100%: 0 to 1).
        const extent = (gs) => {
            let lo = Infinity,
                hi = -Infinity;
            for (const g of gs) {
                if (g.grouping === "percentStacked") {
                    lo = Math.min(lo, 0);
                    hi = Math.max(hi, 1);
                    continue;
                }
                if (g.grouping === "stacked") {
                    for (let i = 0; i < count; i++) {
                        let pos = 0,
                            neg = 0;
                        for (const s of g.series) {
                            const v = (s.val || [])[i];
                            if (v > 0) pos += v;
                            else if (v < 0) neg += v;
                            if (v !== null && v !== undefined) lo = Math.min(lo, v);
                        }
                        if (neg < 0) lo = Math.min(lo, neg);
                        hi = Math.max(hi, pos);
                    }
                    continue;
                }
                for (const s of g.series) for (const v of s.val || s.y || []) if (v !== null && v !== undefined) (lo = Math.min(lo, v)), (hi = Math.max(hi, v));
            }
            return lo <= hi ? [lo, hi] : [0, 1];
        };
        const percent = (g) => g.grouping === "percentStacked";
        const valFormat = (g, ax) => (percent(g) ? "0%" : ax.format && !ax.sourceLinked ? ax.format : (g.series.find((s) => s.format) || {}).format || ax.format || "");

        // Room for the labels of the axes.
        const valLabelsWidth = (sc, fmt) => Math.max(...sc.ticks.map((t) => measure(format(t, fmt), px))) + 6;
        let plot = { ...a };
        const catTitle = catAxis.title,
            valTitle = primaryVal.title;
        const lenV = horizontal ? plot.right - plot.left : plot.bottom - plot.top;
        let sc1, sc2, scX;
        if (xy) {
            const xs = groups.flatMap((g) => g.series.flatMap((s) => (s.x && typeof s.x[0] === "number" ? s.x : (s.y || []).map((_, i) => i + 1))));
            const vx = xs.filter((v) => v !== null);
            let [ylo, yhi] = extent(groups.map((g) => ({ ...g, series: g.series.map((s) => ({ val: s.y })) })));
            let [xlo, xhi] = [Math.min(...vx), Math.max(...vx)];
            // Bubbles: room for them inside the axes.
            if (first.kind === "bubble") {
                const [dx, dy] = [(xhi - xlo || 1) * 0.15, (yhi - ylo || 1) * 0.15];
                [xlo, xhi, ylo, yhi] = [xlo - dx, xhi + dx, ylo - dy, yhi + dy];
            }
            scX = niceScale(xlo, xhi, plot.right - plot.left, catAxis, true, 80);
            sc1 = niceScale(ylo, yhi, lenV, primaryVal, true);
        } else {
            sc1 = niceScale(...extent(groups.filter((g) => !secondary.includes(g))), lenV, primaryVal, true, horizontal ? 80 : 36);
            if (secondary.length) sc2 = niceScale(...extent(secondary), lenV, valOf(secondary[0]), true);
        }
        const fmt1 = xy ? valFormat({ series: groups[0].series }, primaryVal) : valFormat(first, primaryVal);
        const fmt2 = secondary.length ? valFormat(secondary[0], valOf(secondary[0])) : "";
        const fmtX = xy ? catAxis.format || "" : "";
        if (valTitle) horizontal ? (plot.bottom -= px + 6) : (plot.left += px + 6);
        if (catTitle) horizontal ? (plot.left += px + 6) : (plot.bottom -= px + 6);
        const showVal = !primaryVal.deleted && !primaryVal.noLabels;
        const showCat = !catAxis.deleted && !catAxis.noLabels;
        if (horizontal) {
            if (showCat) plot.left += Math.min((a.right - a.left) * 0.4, Math.max(...cats.map((c) => measure(c, px))) + 8);
            if (showVal) plot.bottom -= px + 6;
        } else {
            if (showVal) plot.left += valLabelsWidth(sc1, fmt1);
            if (sc2 && !valOf(secondary[0]).deleted) plot.right -= valLabelsWidth(sc2, fmt2);
            if (showCat) {
                const labels = xy ? scX.ticks.map((t) => format(t, fmtX)) : cats;
                const band = (plot.right - plot.left) / (xy ? Math.max(1, labels.length - 1) : count);
                const wide = Math.max(...labels.map((c) => measure(c, px)));
                plot.rotate = !xy && wide > band * 0.95;
                plot.bottom -= plot.rotate ? Math.min((a.bottom - a.top) * 0.4, wide * 0.72 + px) : px + 6;
            }
        }
        if (!(plot.right - plot.left > 10 && plot.bottom - plot.top > 10)) return;
        const P = plot;
        // Value → position along the value axis; category index → the middle of its band.
        const vpos = (sc, v, reversed) => {
            const f = (v - sc.lo) / (sc.hi - sc.lo);
            const g = reversed ? 1 - f : f;
            return horizontal ? P.left + g * (P.right - P.left) : P.bottom - g * (P.bottom - P.top);
        };
        const catReversed = !!catAxis.reversed;
        const band = (horizontal ? P.bottom - P.top : P.right - P.left) / count;
        const cpos = (i) => {
            const k = catReversed ? count - 1 - i : i;
            return horizontal ? P.bottom - (k + 0.5) * band : P.left + (k + 0.5) * band;
        };
        const xpos = (v) => P.left + ((v - scX.lo) / (scX.hi - scX.lo)) * (P.right - P.left);

        // Gridlines and the labels of the value axes.
        const r1 = !!primaryVal.reversed;
        for (const t of sc1.ticks) {
            const p = vpos(sc1, t, r1);
            if (primaryVal.gridlines !== false && (primaryVal.gridlines || !xy)) {
                out.push(horizontal ? `<line x1="${num(p)}" y1="${num(P.top)}" x2="${num(p)}" y2="${num(P.bottom)}" stroke="${LINE}"/>` : `<line x1="${num(P.left)}" y1="${num(p)}" x2="${num(P.right)}" y2="${num(p)}" stroke="${LINE}"/>`);
            }
            if (showVal) horizontal ? text(p, P.bottom + px / 2 + 4, format(t, fmt1), { anchor: "middle" }) : text(P.left - 4, p, format(t, fmt1), { anchor: "end" });
        }
        if (sc2 && !valOf(secondary[0]).deleted) for (const t of sc2.ticks) text(P.right + 4, vpos(sc2, t, !!valOf(secondary[0]).reversed), format(t, fmt2), { anchor: "start" });
        if (xy) {
            for (const t of scX.ticks) {
                const p = xpos(t);
                if (catAxis.gridlines) out.push(`<line x1="${num(p)}" y1="${num(P.top)}" x2="${num(p)}" y2="${num(P.bottom)}" stroke="${LINE}"/>`);
                if (showCat) text(p, P.bottom + px / 2 + 4, format(t, fmtX), { anchor: "middle" });
            }
        } else if (showCat) {
            cats.forEach((c, i) => {
                const p = cpos(i);
                if (horizontal) text(P.left - 4, p, c, { anchor: "end" });
                else if (P.rotate) text(p + px / 3, P.bottom + 4, c, { anchor: "end", rotate: -45, base: "hanging" });
                else text(p, P.bottom + px / 2 + 4, c, { anchor: "middle" });
            });
        }
        // The category axis line (at the value 0, or the edge).
        const base = vpos(sc1, Math.min(Math.max(0, sc1.lo), sc1.hi), r1);
        if (!catAxis.deleted) out.push(horizontal ? `<line x1="${num(base)}" y1="${num(P.top)}" x2="${num(base)}" y2="${num(P.bottom)}" stroke="${LINE}"/>` : `<line x1="${num(P.left)}" y1="${num(xy ? P.bottom : base)}" x2="${num(P.right)}" y2="${num(xy ? P.bottom : base)}" stroke="${LINE}"/>`);
        if (valTitle) horizontal ? text((P.left + P.right) / 2, a.bottom - px / 2, valTitle, { anchor: "middle" }) : text(a.left + px / 2, (P.top + P.bottom) / 2, valTitle, { anchor: "middle", rotate: -90 });
        if (catTitle) horizontal ? text(a.left + px / 2, (P.top + P.bottom) / 2, catTitle, { anchor: "middle", rotate: -90 }) : text((P.left + P.right) / 2, a.bottom - px / 2, catTitle, { anchor: "middle" });

        const labelsOf = [];
        groups.forEach((g) => {
            const sc = secondary.includes(g) ? sc2 : sc1;
            const rev = !!valOf(g).reversed;
            const vp = (v) => vpos(sc, v, rev);
            const n = g.series.length;
            if (g.kind === "bar") {
                const stacked = g.grouping === "stacked" || g.grouping === "percentStacked";
                const gap = +(g.gapWidth === undefined ? 150 : g.gapWidth) / 100;
                const overlap = stacked ? 1 : +(g.overlap || 0) / 100;
                const w = band / (n - (n - 1) * overlap + gap);
                const groupWidth = w * (n - (n - 1) * overlap);
                const pos = new Array(count).fill(0),
                    neg = new Array(count).fill(0);
                const sums = Array.from({ length: count }, (_, i) => g.series.reduce((t, s) => t + Math.abs((s.val || [])[i] || 0), 0));
                g.series.forEach((s, k) => {
                    (s.val || []).forEach((v, i) => {
                        if (v === null || v === undefined) return;
                        let a0 = 0,
                            a1 = v;
                        if (stacked) {
                            if (g.grouping === "percentStacked") v = sums[i] ? v / sums[i] : 0;
                            if (v >= 0) (a0 = pos[i]), (a1 = pos[i] += v);
                            else (a0 = neg[i]), (a1 = neg[i] += v);
                        }
                        if (s.noFill) return;
                        // The first series first along the categories (bottom up for bars).
                        const step = stacked ? 0 : k * w * (1 - overlap);
                        const c = horizontal !== catReversed ? cpos(i) + groupWidth / 2 - w - step : cpos(i) - groupWidth / 2 + step;
                        const p0 = vp(a0),
                            p1 = vp(a1);
                        const color = pointColor(s, g, i);
                        const [x, y, ww, hh] = horizontal ? [Math.min(p0, p1), c, Math.abs(p1 - p0), w] : [c, Math.min(p0, p1), w, Math.abs(p1 - p0)];
                        out.push(`<rect x="${num(x)}" y="${num(y)}" width="${num(Math.max(ww, 0.5))}" height="${num(Math.max(hh, 0.5))}" fill="${color}"/>`);
                        if (s.labels && s.labels.val) labelsOf.push(horizontal ? [x + ww + 3, y + hh / 2, (s.valText || [])[i], "start"] : [x + ww / 2, y - px / 2 - 2, (s.valText || [])[i], "middle"]);
                    });
                });
            } else if (g.kind === "line" || g.kind === "area") {
                const stacked = g.grouping === "stacked" || g.grouping === "percentStacked";
                const acc = new Array(count).fill(0);
                const sums = Array.from({ length: count }, (_, i) => g.series.reduce((t, s) => t + ((s.val || [])[i] || 0), 0));
                const layers = g.series.map((s, k) => {
                    const lower = acc.slice();
                    const pts = (s.val || []).map((v, i) => {
                        if (v === null || v === undefined) return stacked ? [cpos(i), vp(acc[i])] : null;
                        if (g.grouping === "percentStacked") v = sums[i] ? v / sums[i] : 0;
                        if (stacked) acc[i] += v;
                        return horizontal ? [vp(stacked ? acc[i] : v), cpos(i)] : [cpos(i), vp(stacked ? acc[i] : v)];
                    });
                    return { s, k, pts, lower };
                });
                // Areas: the later ones in front, from the axis (or the one below).
                for (const { s, k, pts, lower } of layers) {
                    const color = s.color || autoColor(k);
                    const runs = [];
                    let run = [];
                    pts.forEach((p, i) => (p ? run.push([p, i]) : (run.length && runs.push(run), (run = []))));
                    if (run.length) runs.push(run);
                    for (const r of runs) {
                        const line = r.map(([p]) => p);
                        if (g.kind === "area") {
                            const back = r.map(([, i]) => [cpos(i), vp(stacked ? lower[i] : Math.max(sc.lo, Math.min(0, sc.hi)))]).reverse();
                            if (!s.noFill) out.push(`<polygon points="${line.concat(back).map((p) => p.map(num).join(",")).join(" ")}" fill="${color}"/>`);
                        } else if (!s.noFill && !g.stock && line.length > 1) {
                            const d = s.smooth ? smoothPath(line) : "M" + line.map((p) => p.map(num).join(",")).join(" L");
                            out.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="${num((s.width || 2.25) * (4 / 3))}" stroke-linejoin="round" stroke-linecap="round"/>`);
                        }
                        if (g.kind === "line" && (s.marker || g.stock)) for (const [x, y] of line) out.push(`<circle cx="${num(x)}" cy="${num(y)}" r="3.5" fill="${color}"/>`);
                    }
                    if (s.labels && s.labels.val) pts.forEach((p, i) => p && labelsOf.push([p[0], p[1] - px * 0.8, (s.valText || [])[i], "middle"]));
                }
            } else if (g.kind === "scatter" || g.kind === "bubble") {
                const sizes = g.series.flatMap((s) => s.size || []).filter((v) => v > 0);
                const maxSize = Math.max(1, ...sizes);
                const maxR = Math.min(P.right - P.left, P.bottom - P.top) / 8;
                g.series.forEach((s, k) => {
                    const color = s.color || autoColor(k);
                    const xs = s.x && typeof s.x[0] === "number" ? s.x : (s.y || []).map((_, i) => i + 1);
                    const pts = (s.y || []).map((v, i) => (v === null || xs[i] === null || xs[i] === undefined ? null : [xpos(xs[i]), vp(v)]));
                    if (g.kind === "bubble") {
                        pts.forEach((p, i) => p && out.push(`<circle cx="${num(p[0])}" cy="${num(p[1])}" r="${num(Math.sqrt(Math.max(0, (s.size || [])[i] || 0) / maxSize) * maxR)}" fill="${color}" fill-opacity="0.75" stroke="#fff"/>`));
                        return;
                    }
                    const line = pts.filter(Boolean);
                    if (!s.noFill && line.length > 1 && !/^marker$/i.test(g.scatterStyle || "")) {
                        const d = s.smooth ? smoothPath(line) : "M" + line.map((p) => p.map(num).join(",")).join(" L");
                        out.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="${num((s.width || 2.25) * (4 / 3))}"/>`);
                    }
                    if (s.marker) for (const [x, y] of line) out.push(`<circle cx="${num(x)}" cy="${num(y)}" r="3.5" fill="${color}"/>`);
                    if (s.labels && s.labels.val) pts.forEach((p, i) => p && labelsOf.push([p[0], p[1] - px * 0.8, (s.yText || [])[i], "middle"]));
                });
            }
        });
        for (const [x, y, s, anchor] of labelsOf) if (s) text(x, y, s, { anchor });
    }

    return { draw, format };
})();
