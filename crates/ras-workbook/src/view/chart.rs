//! Excel's charts for the grid: what the page needs to draw one (`drawChart` in view.js): its
//! title, legend, chart groups (kind, grouping, series with their colors) and axes, with the
//! values of the series read from the cells they refer to as they are now (so that a chart shows
//! what a run changed), or from the chart's cache when the cells cannot be read.
//!
//! How the chart XML is read (the automatic colors of series, the sources of values) follows
//! xljs (`src/api/xlsx_chart.rs`), which draws Excel's charts with AG Charts; the grid draws
//! them as SVG itself.

use std::collections::HashMap;

use serde_json::{Map, Value as Json, json};

use super::drawing::{Env, Parts, base_color, color_of, modified, mods_of, office_color};
use super::xml::Node;

/// A cell as a chart reads it: its number (None for text and empty cells), what it shows, its
/// number format.
#[derive(Clone, Debug, Default)]
pub struct Point {
    pub number: Option<f64>,
    pub text: String,
    pub format: String,
}

/// The automatic colors of a chart's series and points, as Excel 2013 and later picks them: the
/// chart's color style (xl/charts/colorsN.xml, the palette chosen with "Change Colors") or else
/// the theme's accent1 to accent6. After the last color they repeat with the next variation
/// (darker, lighter, ...).
struct AutoColors {
    colors: Vec<String>,
    variations: Vec<Vec<(String, f64)>>,
}

impl AutoColors {
    fn new(parts: &Parts, chart_part: &str, theme: &HashMap<String, String>) -> AutoColors {
        let style = parts
            .rels(chart_part)
            .into_values()
            .find(|(t, _)| t.ends_with("/chartColorStyle"))
            .and_then(|(_, p)| parts.xml(&p));
        if let Some(st) = style {
            let colors: Vec<String> = st
                .children
                .iter()
                .filter(|c| matches!(c.name.as_str(), "srgbClr" | "schemeClr" | "sysClr"))
                .filter_map(|c| {
                    modified(&base_color(c, theme)?, &mods_of(c)).map(|x| x[1..].to_string())
                })
                .collect();
            if !colors.is_empty() {
                return AutoColors {
                    colors,
                    variations: st.all("variation").map(mods_of).collect(),
                };
            }
        }
        // Office's default color style (colors1.xml of a new chart).
        let m = |v: &[(&str, f64)]| {
            v.iter()
                .map(|(k, x)| (k.to_string(), *x))
                .collect::<Vec<_>>()
        };
        AutoColors {
            colors: (1..=6)
                .map(|i| format!("accent{i}"))
                .filter_map(|a| {
                    theme
                        .get(&a)
                        .cloned()
                        .or_else(|| office_color(&a).map(str::to_string))
                })
                .collect(),
            variations: vec![
                vec![],
                m(&[("lumMod", 0.6)]),
                m(&[("lumMod", 0.8), ("lumOff", 0.2)]),
                m(&[("lumMod", 0.8)]),
                m(&[("lumMod", 0.6), ("lumOff", 0.4)]),
                m(&[("lumMod", 0.5)]),
                m(&[("lumMod", 0.7), ("lumOff", 0.3)]),
                m(&[("lumMod", 0.7)]),
                m(&[("lumMod", 0.5), ("lumOff", 0.5)]),
            ],
        }
    }

    /// The color of the i-th series (its c:idx) or point.
    fn nth(&self, i: usize) -> Option<String> {
        let n = self.colors.len();
        if n == 0 {
            return None;
        }
        let none = Vec::new();
        let var = if self.variations.is_empty() {
            &none
        } else {
            &self.variations[(i / n) % self.variations.len()]
        };
        modified(&self.colors[i % n], var)
    }
}

/// The areas of a series reference ("'売上 2026'!$B$2:$B$13", "(Sheet1!$A$1,Sheet1!$A$3)").
fn areas(f: &str) -> Vec<String> {
    let f = f.trim();
    let f = f
        .strip_prefix('(')
        .and_then(|x| x.strip_suffix(')'))
        .unwrap_or(f);
    let mut parts = Vec::new();
    let (mut cur, mut quoted) = (String::new(), false);
    for ch in f.chars() {
        match ch {
            '\'' => {
                quoted = !quoted;
                cur.push(ch);
            }
            ',' if !quoted => parts.push(std::mem::take(&mut cur)),
            _ => cur.push(ch),
        }
    }
    parts.push(cur);
    parts
        .into_iter()
        .map(|p| p.trim().to_string())
        .filter(|p| !p.is_empty())
        .collect()
}

/// The points of c:cat, c:val, c:xVal, c:yVal, c:bubbleSize or c:tx: from the cells they refer
/// to (down a column or along a row; of an area more than one cell wide and tall, the innermost
/// level of categories, its last column or row), else from the cache.
fn source(n: &Node, env: &Env) -> Vec<Point> {
    let Some(r) = n.children.iter().find(|c| {
        matches!(
            c.name.as_str(),
            "numRef" | "strRef" | "multiLvlStrRef" | "numLit" | "strLit"
        )
    }) else {
        return Vec::new();
    };
    let formula = r
        .child("f")
        .map(|f| f.text.trim().to_string())
        .unwrap_or_default();
    if !formula.is_empty() {
        let mut out = Vec::new();
        let mut read = true;
        for area in areas(&formula) {
            let Some(grid) = (env.range)(&area) else {
                read = false;
                break;
            };
            let (rows, cols) = (grid.len(), grid.first().map(Vec::len).unwrap_or(0));
            if rows == 0 || cols == 0 {
                continue;
            }
            if cols == 1 || rows >= cols {
                out.extend(grid.iter().map(|row| row[cols - 1].clone()));
            } else {
                out.extend(grid[rows - 1].iter().cloned());
            }
        }
        if read && !out.is_empty() {
            return out;
        }
    }
    // The cache (and literals): points by index.
    let cache = r
        .children
        .iter()
        .find(|c| {
            matches!(
                c.name.as_str(),
                "numCache" | "strCache" | "multiLvlStrCache"
            )
        })
        .unwrap_or(r);
    let cache = if cache.name == "multiLvlStrCache" {
        cache.all("lvl").next().unwrap_or(cache)
    } else {
        cache
    };
    let format = cache
        .child("formatCode")
        .map(|f| f.text.trim().to_string())
        .unwrap_or_default();
    let count = cache
        .child("ptCount")
        .and_then(|c| c.attr("val"))
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(0)
        .min(100_000);
    let numeric = matches!(r.name.as_str(), "numRef" | "numLit");
    let mut out = vec![Point::default(); count];
    for p in cache.all("pt") {
        let idx = p
            .attr("idx")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(out.len())
            .min(100_000);
        if idx >= out.len() {
            out.resize(idx + 1, Point::default());
        }
        let text = p.child("v").map(|v| v.text.clone()).unwrap_or_default();
        let number = if numeric {
            text.trim().parse::<f64>().ok()
        } else {
            None
        };
        out[idx] = Point {
            number,
            text,
            format: p
                .attr("formatCode")
                .map(str::to_string)
                .unwrap_or_else(|| format.clone()),
        };
    }
    out
}

/// The text of a title or a series name (c:tx: rich text, a value or a cell).
fn text_of(tx: &Node, env: &Env) -> Option<String> {
    if let Some(rich) = tx.find("rich") {
        let lines: Vec<String> = rich
            .all("p")
            .map(|p| {
                let mut runs = Vec::new();
                p.find_all("t", &mut runs);
                runs.iter().map(|t| t.text.as_str()).collect::<String>()
            })
            .collect();
        return Some(lines.join("\n").trim().to_string());
    }
    if let Some(v) = tx.child("v") {
        return Some(v.text.clone());
    }
    let parts: Vec<String> = source(tx, env)
        .into_iter()
        .map(|p| p.text)
        .filter(|t| !t.is_empty())
        .collect();
    (!parts.is_empty()).then(|| parts.join(" "))
}

/// The fill (or, for lines, the line) of a c:spPr: its color, or that it has none.
fn paint(sp: Option<&Node>, line: bool, theme: &HashMap<String, String>) -> (Option<String>, bool) {
    let Some(sp) = sp else { return (None, false) };
    let src = if line {
        sp.child("ln").unwrap_or(sp)
    } else {
        sp
    };
    if src.child("noFill").is_some() {
        return (None, true);
    }
    let color = src.child("solidFill").and_then(|f| color_of(f, theme));
    let color = color.or_else(|| {
        src.child("gradFill")
            .and_then(|g| g.find("gs"))
            .and_then(|gs| color_of(gs, theme))
    });
    (color, false)
}

/// The size (pt) of the text of a title or the chart (`txPr`'s or the rich text's defRPr).
fn text_size(n: Option<&Node>) -> Option<f64> {
    let d = n?.find("defRPr")?;
    d.attr("sz")
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| v / 100.0)
}

/// Whether data labels show values, percentages, category names (c:dLbls).
fn labels(d: Option<&Node>) -> Option<Json> {
    let d = d?;
    let on = |k: &str| {
        d.child(k)
            .and_then(|c| c.attr("val"))
            .is_some_and(|v| v == "1" || v == "true")
    };
    let (val, pct, cat) = (on("showVal"), on("showPercent"), on("showCatName"));
    (val || pct || cat).then(|| json!({ "val": val, "percent": pct, "cat": cat, "pos": d.child("dLblPos").and_then(|p| p.attr("val")).unwrap_or("") }))
}

const GROUPS: &[&str] = &[
    "barChart",
    "bar3DChart",
    "lineChart",
    "line3DChart",
    "pieChart",
    "pie3DChart",
    "ofPieChart",
    "doughnutChart",
    "areaChart",
    "area3DChart",
    "scatterChart",
    "bubbleChart",
    "radarChart",
    "stockChart",
];

/// What the page draws of the chart part `part`; None for the charts it does not draw (chartEx:
/// waterfall, treemap, ...; surface charts).
pub(crate) fn chart_spec(
    parts: &Parts,
    part: &str,
    theme: &HashMap<String, String>,
    env: &Env,
) -> Option<Json> {
    let root = parts.xml(part)?;
    let chart = root.child("chart")?;
    let auto = AutoColors::new(parts, part, theme);
    let mut out = Map::new();
    let on = |n: Option<&Node>| {
        n.and_then(|c| c.attr("val"))
            .is_none_or(|v| v == "1" || v == "true")
    };
    if let Some(t) = chart.child("title") {
        let text = t.child("tx").and_then(|tx| text_of(tx, env));
        out.insert(
            "title".into(),
            text.map(Json::String).unwrap_or(Json::Bool(true)),
        );
        if let Some(s) = text_size(t.child("tx").or(t.child("txPr"))) {
            out.insert("titleSize".into(), json!(s));
        }
    }
    out.insert(
        "autoTitleDeleted".into(),
        json!(chart.child("autoTitleDeleted").is_some_and(|c| on(Some(c)))),
    );
    if let Some(l) = chart.child("legend") {
        out.insert(
            "legend".into(),
            json!(
                l.child("legendPos")
                    .and_then(|p| p.attr("val"))
                    .unwrap_or("r")
            ),
        );
    }
    if let Some(s) = text_size(root.child("txPr")) {
        out.insert("textSize".into(), json!(s));
    }
    // The chart's own fill and border (none: transparent, as Excel shows them).
    let space = root.child("spPr");
    let (fill, no_fill) = paint(space, false, theme);
    let (line, no_line) = paint(space, true, theme);
    out.insert(
        "fill".into(),
        json!(if no_fill {
            "none".to_string()
        } else {
            fill.unwrap_or_else(|| "#ffffff".into())
        }),
    );
    out.insert(
        "border".into(),
        json!(if no_line {
            "none".to_string()
        } else {
            line.unwrap_or_else(|| "#d9d9d9".into())
        }),
    );
    let plot = chart.child("plotArea")?;
    if plot.child("surfaceChart").is_some() || plot.child("surface3DChart").is_some() {
        return None;
    }
    let mut groups = Vec::new();
    for g in plot
        .children
        .iter()
        .filter(|c| GROUPS.contains(&c.name.as_str()))
    {
        let line = matches!(
            g.name.as_str(),
            "lineChart" | "line3DChart" | "scatterChart" | "radarChart" | "stockChart"
        );
        let group_labels = labels(g.child("dLbls"));
        let mut series = Vec::new();
        for (i, s) in g.all("ser").enumerate() {
            let mut o = Map::new();
            o.insert(
                "name".into(),
                s.child("tx")
                    .and_then(|tx| text_of(tx, env))
                    .map(Json::String)
                    .unwrap_or_else(|| json!(format!("系列{}", i + 1))),
            );
            // Lines take their color from the line; markers-only scatter from the marker.
            let marker_none = s
                .child("marker")
                .and_then(|m| m.child("symbol"))
                .and_then(|m| m.attr("val"))
                == Some("none");
            let (mut color, no_fill) = paint(s.child("spPr"), line, theme);
            if color.is_none() && line {
                color = paint(
                    s.child("marker").and_then(|m| m.child("spPr")),
                    false,
                    theme,
                )
                .0;
            }
            let idx = s
                .child("idx")
                .and_then(|c| c.attr("val"))
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(i);
            if color.is_none() && (!no_fill || line) {
                color = auto.nth(idx);
            }
            o.insert("color".into(), json!(color));
            o.insert("noFill".into(), json!(no_fill));
            o.insert("marker".into(), json!(!marker_none));
            o.insert(
                "smooth".into(),
                json!(s.child("smooth").is_some_and(|c| on(Some(c)))),
            );
            if let Some(w) = s
                .child("spPr")
                .and_then(|p| p.child("ln"))
                .and_then(|l| l.attr("w"))
                .and_then(|v| v.parse::<f64>().ok())
            {
                o.insert("width".into(), json!(w / 12700.0));
            }
            let points: Vec<Json> = s
                .all("dPt")
                .filter_map(|p| {
                    let i = p.child("idx")?.attr("val")?.parse::<u32>().ok()?;
                    let (c, _) = paint(p.child("spPr"), false, theme);
                    Some(json!({ "idx": i, "color": c }))
                })
                .collect();
            if !points.is_empty() {
                o.insert("points".into(), json!(points));
            }
            if let Some(l) = labels(s.child("dLbls")).or(group_labels.clone()) {
                o.insert("labels".into(), l);
            }
            for (k, name) in [
                ("cat", "cat"),
                ("val", "val"),
                ("x", "xVal"),
                ("y", "yVal"),
                ("size", "bubbleSize"),
            ] {
                let Some(n) = s.child(name) else { continue };
                let pts = source(n, env);
                let text = matches!(k, "cat")
                    || (k == "x" && pts.iter().any(|p| p.number.is_none() && !p.text.is_empty()));
                if text {
                    o.insert(
                        k.into(),
                        json!(pts.iter().map(|p| p.text.clone()).collect::<Vec<_>>()),
                    );
                } else {
                    o.insert(
                        k.into(),
                        json!(pts.iter().map(|p| p.number).collect::<Vec<_>>()),
                    );
                    if matches!(k, "val" | "y") {
                        // What the cells show, for the labels of the points.
                        o.insert(
                            format!("{k}Text"),
                            json!(pts.iter().map(|p| p.text.clone()).collect::<Vec<_>>()),
                        );
                        if let Some(f) = pts
                            .iter()
                            .map(|p| p.format.as_str())
                            .find(|f| !f.is_empty() && *f != "General")
                        {
                            o.insert("format".into(), json!(f));
                        }
                    }
                }
            }
            series.push(Json::Object(o));
        }
        let mut o = Map::new();
        let kind = match g.name.as_str() {
            "barChart" | "bar3DChart" => "bar",
            "lineChart" | "line3DChart" | "stockChart" => "line",
            "pieChart" | "pie3DChart" | "ofPieChart" => "pie",
            "doughnutChart" => "doughnut",
            "areaChart" | "area3DChart" => "area",
            "scatterChart" => "scatter",
            "bubbleChart" => "bubble",
            _ => "radar",
        };
        o.insert("kind".into(), json!(kind));
        for k in [
            "barDir",
            "grouping",
            "scatterStyle",
            "holeSize",
            "radarStyle",
            "gapWidth",
            "overlap",
            "firstSliceAng",
        ] {
            if let Some(v) = g.child(k).and_then(|c| c.attr("val")) {
                o.insert(k.into(), json!(v));
            }
        }
        // Lines of a stock chart are its high-low lines: points only.
        if g.name == "stockChart" {
            o.insert("stock".into(), json!(true));
        }
        o.insert(
            "varyColors".into(),
            json!(g.child("varyColors").is_some_and(|c| on(Some(c)))),
        );
        o.insert(
            "axes".into(),
            json!(
                g.all("axId")
                    .filter_map(|a| a.attr("val"))
                    .collect::<Vec<_>>()
            ),
        );
        o.insert("series".into(), json!(series));
        groups.push(Json::Object(o));
    }
    if groups.is_empty() {
        return None;
    }
    out.insert("groups".into(), json!(groups));
    // The automatic colors of points (pies whose slices have no color of their own).
    out.insert(
        "autoColors".into(),
        json!((0..32).filter_map(|i| auto.nth(i)).collect::<Vec<_>>()),
    );
    let mut axes = Vec::new();
    for a in plot
        .children
        .iter()
        .filter(|c| matches!(c.name.as_str(), "catAx" | "dateAx" | "valAx" | "serAx"))
    {
        let val = |k: &str| a.child(k).and_then(|c| c.attr("val"));
        let mut o = Map::new();
        o.insert("kind".into(), json!(a.name.trim_end_matches("Ax")));
        o.insert("id".into(), json!(val("axId").unwrap_or("")));
        o.insert("pos".into(), json!(val("axPos").unwrap_or("")));
        o.insert(
            "deleted".into(),
            json!(a.child("delete").is_some_and(|c| on(Some(c)))),
        );
        o.insert(
            "gridlines".into(),
            json!(a.child("majorGridlines").is_some()),
        );
        if let Some(t) = a.child("title") {
            o.insert(
                "title".into(),
                json!(t.child("tx").and_then(|tx| text_of(tx, env))),
            );
        }
        if let Some(nf) = a.child("numFmt") {
            o.insert("format".into(), json!(nf.attr("formatCode").unwrap_or("")));
            o.insert(
                "sourceLinked".into(),
                json!(matches!(nf.attr("sourceLinked"), Some("1" | "true"))),
            );
        }
        if val("tickLblPos") == Some("none") {
            o.insert("noLabels".into(), json!(true));
        }
        if let Some(sc) = a.child("scaling") {
            o.insert(
                "reversed".into(),
                json!(sc.child("orientation").and_then(|c| c.attr("val")) == Some("maxMin")),
            );
            for k in ["min", "max", "logBase"] {
                if let Some(v) = sc
                    .child(k)
                    .and_then(|c| c.attr("val"))
                    .and_then(|v| v.parse::<f64>().ok())
                {
                    o.insert(k.into(), json!(v));
                }
            }
        }
        if let Some(v) = val("majorUnit").and_then(|v| v.parse::<f64>().ok()) {
            o.insert("majorUnit".into(), json!(v));
        }
        axes.push(Json::Object(o));
    }
    out.insert("axes".into(), json!(axes));
    Some(Json::Object(out))
}
