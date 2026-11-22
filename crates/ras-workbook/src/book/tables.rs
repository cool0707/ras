//! The look of tables (ListObjects) in the grid: the table style's fills, fonts and borders laid
//! under the cells' own formats, as xljs lays them for printing (`src/api/xlsx_table.rs`
//! there). Custom styles come from styles.xml (`tableStyles` and their `dxfs`); the built-in
//! ones (TableStyleLight1 … TableStyleDark11) from an approximation of Excel's presets by family.

use std::collections::HashMap;

use umya_spreadsheet::{Color, Worksheet};

use crate::view::xml::Node;

/// A border line: style and color (None: automatic).
#[derive(Debug, Clone, PartialEq)]
pub struct Side {
    pub style: String,
    pub color: Option<Color>,
}

/// The parts of a differential format a table style uses.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Dxf {
    pub fill: Option<Color>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<Color>,
    /// left, right, top, bottom, vertical (inside), horizontal (inside).
    pub sides: [Option<Side>; 6],
}

/// A table style: its elements (type, band size, format).
pub type StyleDef = Vec<(String, u32, Dxf)>;

/// What a table shows of its style: first column, last column, row stripes, column stripes.
pub type Shown = (bool, bool, bool, bool);

/// The custom table styles of a workbook (its styles.xml, as a tree) by name.
pub fn custom_styles(root: &Node) -> HashMap<String, StyleDef> {
    let mut out = HashMap::new();
    let dxfs: Vec<Dxf> = root
        .child("dxfs")
        .map(|d| d.all("dxf").map(dxf).collect())
        .unwrap_or_default();
    let Some(ts) = root.child("tableStyles") else {
        return out;
    };
    for s in ts.all("tableStyle") {
        let Some(name) = s.attr("name") else { continue };
        let els = s
            .all("tableStyleElement")
            .filter_map(|e| {
                let d = dxfs.get(e.attr("dxfId")?.parse::<usize>().ok()?)?.clone();
                let size = e.attr("size").and_then(|v| v.parse().ok()).unwrap_or(1);
                Some((e.attr("type")?.to_string(), size, d))
            })
            .collect();
        out.insert(name.to_string(), els);
    }
    out
}

fn color(n: &Node) -> Option<Color> {
    let mut c = Color::default();
    if n.attr("auto").is_some_and(|v| v == "1" || v == "true") {
        return None;
    }
    if let Some(t) = n.attr("theme").and_then(|v| v.parse().ok()) {
        c.set_theme_index(t);
    } else if let Some(rgb) = n.attr("rgb") {
        let rgb = if rgb.len() == 6 {
            format!("FF{rgb}")
        } else {
            rgb.to_string()
        };
        if rgb.len() != 8 || !rgb.bytes().all(|b| b.is_ascii_hexdigit()) {
            return None;
        }
        c.set_argb_str(rgb);
    } else {
        let i = n.attr("indexed").and_then(|v| v.parse().ok())?;
        if i == 64 {
            // The system foreground.
            return None;
        }
        c.set_indexed(i);
    }
    if let Some(t) = n.attr("tint").and_then(|v| v.parse().ok()) {
        c.set_tint(t);
    }
    Some(c)
}

fn dxf(n: &Node) -> Dxf {
    let mut d = Dxf::default();
    let on = |e: Option<&Node>| e.map(|e| !matches!(e.attr("val"), Some("0" | "false")));
    if let Some(f) = n.child("font") {
        d.bold = on(f.child("b"));
        d.italic = on(f.child("i"));
        d.color = f.child("color").and_then(color);
    }
    if let Some(p) = n.child("fill").and_then(|f| f.child("patternFill")) {
        let t = p.attr("patternType").unwrap_or("solid");
        if t != "none" {
            // A dxf's solid fill is its background color.
            d.fill = p
                .child("bgColor")
                .and_then(color)
                .or_else(|| p.child("fgColor").and_then(color));
        }
    }
    if let Some(b) = n.child("border") {
        for (i, name) in ["left", "right", "top", "bottom", "vertical", "horizontal"]
            .into_iter()
            .enumerate()
        {
            if let Some(s) = b.child(name)
                && let Some(st) = s.attr("style")
            {
                d.sides[i] = Some(Side {
                    style: st.to_string(),
                    color: s.child("color").and_then(color),
                });
            }
        }
    }
    d
}

// ---- the built-in styles

fn theme(i: u32, tint: f64) -> Option<Color> {
    let mut c = Color::default();
    c.set_theme_index(i);
    if tint != 0.0 {
        c.set_tint(tint);
    }
    Some(c)
}

fn side(style: &str, c: Option<Color>) -> Option<Side> {
    Some(Side {
        style: style.into(),
        color: c,
    })
}

fn el(t: &str, d: Dxf) -> (String, u32, Dxf) {
    (t.to_string(), 1, d)
}

/// Excel's built-in table styles, approximated by family: Light 1-7, 8-14, 15-21, Medium 1-7,
/// 8-14, 15-21, 22-28, Dark 1-7, 8-11. Within a family the first uses the text color and the
/// next six the accents 1-6.
pub fn builtin(name: &str) -> Option<StyleDef> {
    let (family, n) = if let Some(v) = name.strip_prefix("TableStyleLight") {
        ("L", v.parse::<u32>().ok()?)
    } else if let Some(v) = name.strip_prefix("TableStyleMedium") {
        ("M", v.parse::<u32>().ok()?)
    } else {
        (
            "D",
            name.strip_prefix("TableStyleDark")?.parse::<u32>().ok()?,
        )
    };
    if n == 0 {
        return None;
    }
    // 0: the text color, 1-6: the accents.
    let i = (n - 1) % 7;
    let group = (n - 1) / 7;
    let base = if i == 0 { 1 } else { 3 + i };
    let c = |tint: f64| {
        if i == 0 && tint > 0.0 {
            theme(0, -(1.0 - tint) * 0.5)
        } else {
            theme(base, tint)
        }
    };
    let white = || theme(0, 0.0);
    let bold = |d: Dxf| Dxf {
        bold: Some(true),
        ..d
    };
    let header_solid = || Dxf {
        fill: c(0.0),
        color: white(),
        bold: Some(true),
        ..Default::default()
    };
    let total_double = || Dxf {
        bold: Some(true),
        sides: [None, None, side("double", c(0.0)), None, None, None],
        ..Default::default()
    };
    let stripe = |tint: f64| Dxf {
        fill: c(tint),
        ..Default::default()
    };
    let mut out: StyleDef = Vec::new();
    match (family, group) {
        ("L", 0) => {
            out.push(el(
                "wholeTable",
                Dxf {
                    color: if i == 0 {
                        theme(1, 0.0)
                    } else {
                        theme(base, -0.25)
                    },
                    sides: [
                        None,
                        None,
                        side("thin", c(0.0)),
                        side("thin", c(0.0)),
                        None,
                        None,
                    ],
                    ..Default::default()
                },
            ));
            out.push(el("firstRowStripe", stripe(0.8)));
            out.push(el("firstColumnStripe", stripe(0.8)));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            out.push(el(
                "headerRow",
                Dxf {
                    bold: Some(true),
                    sides: [None, None, None, side("thin", c(0.0)), None, None],
                    ..Default::default()
                },
            ));
            out.push(el("totalRow", total_double()));
        }
        ("L", 1) => {
            let t = || side("thin", c(0.0));
            out.push(el(
                "wholeTable",
                Dxf {
                    sides: [t(), t(), t(), t(), None, None],
                    ..Default::default()
                },
            ));
            out.push(el(
                "firstRowStripe",
                Dxf {
                    sides: [None, None, t(), t(), None, None],
                    ..Default::default()
                },
            ));
            out.push(el(
                "firstColumnStripe",
                Dxf {
                    sides: [t(), t(), None, None, None, None],
                    ..Default::default()
                },
            ));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            out.push(el("headerRow", header_solid()));
            out.push(el("totalRow", total_double()));
        }
        ("L", _) => {
            let t = || side("thin", c(0.0));
            out.push(el(
                "wholeTable",
                Dxf {
                    sides: [t(), t(), t(), t(), t(), t()],
                    ..Default::default()
                },
            ));
            out.push(el("firstRowStripe", stripe(0.8)));
            out.push(el("firstColumnStripe", stripe(0.8)));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            out.push(el(
                "headerRow",
                Dxf {
                    bold: Some(true),
                    sides: [None, None, None, side("medium", c(0.0)), None, None],
                    ..Default::default()
                },
            ));
            out.push(el("totalRow", total_double()));
        }
        ("M", 0) => {
            let t = || side("thin", c(0.4));
            out.push(el(
                "wholeTable",
                Dxf {
                    sides: [t(), t(), t(), t(), None, t()],
                    ..Default::default()
                },
            ));
            out.push(el("firstRowStripe", stripe(0.8)));
            out.push(el("firstColumnStripe", stripe(0.8)));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            let s = || side("thin", c(0.0));
            out.push(el(
                "headerRow",
                Dxf {
                    sides: [s(), s(), s(), s(), None, None],
                    ..header_solid()
                },
            ));
            out.push(el("totalRow", total_double()));
        }
        ("M", 1) => {
            let w = || side("thin", white());
            out.push(el(
                "wholeTable",
                Dxf {
                    fill: c(0.8),
                    sides: [None, None, None, None, w(), w()],
                    ..Default::default()
                },
            ));
            out.push(el("firstRowStripe", stripe(0.6)));
            out.push(el("firstColumnStripe", stripe(0.6)));
            out.push(el("firstColumn", header_solid()));
            out.push(el("lastColumn", header_solid()));
            out.push(el(
                "headerRow",
                Dxf {
                    sides: [None, None, None, side("thick", white()), None, None],
                    ..header_solid()
                },
            ));
            out.push(el(
                "totalRow",
                Dxf {
                    sides: [None, None, side("thick", white()), None, None, None],
                    ..header_solid()
                },
            ));
        }
        ("M", 2) => {
            let t = || side("thin", theme(1, 0.0));
            out.push(el(
                "wholeTable",
                Dxf {
                    sides: [t(), t(), t(), t(), None, t()],
                    ..Default::default()
                },
            ));
            let gray = || Dxf {
                fill: theme(0, -0.15),
                ..Default::default()
            };
            out.push(el("firstRowStripe", gray()));
            out.push(el("firstColumnStripe", gray()));
            out.push(el("firstColumn", header_solid()));
            out.push(el("lastColumn", header_solid()));
            out.push(el(
                "headerRow",
                Dxf {
                    sides: [None, None, None, side("thick", theme(1, 0.0)), None, None],
                    ..header_solid()
                },
            ));
            out.push(el(
                "totalRow",
                Dxf {
                    bold: Some(true),
                    sides: [None, None, side("double", theme(1, 0.0)), None, None, None],
                    ..Default::default()
                },
            ));
        }
        ("M", _) => {
            let t = || side("thin", c(0.4));
            out.push(el(
                "wholeTable",
                Dxf {
                    fill: c(0.8),
                    sides: [t(), t(), t(), t(), t(), t()],
                    ..Default::default()
                },
            ));
            out.push(el("firstRowStripe", stripe(0.6)));
            out.push(el("firstColumnStripe", stripe(0.6)));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            out.push(el("headerRow", bold(Dxf::default())));
            out.push(el("totalRow", total_double()));
        }
        ("D", 0) => {
            let dark = |t: f64| {
                if i == 0 {
                    theme(1, if t < -0.3 { 0.15 } else { 0.35 })
                } else {
                    theme(base, t)
                }
            };
            out.push(el(
                "wholeTable",
                Dxf {
                    fill: dark(-0.25),
                    color: white(),
                    ..Default::default()
                },
            ));
            let band = || Dxf {
                fill: dark(-0.5),
                ..Default::default()
            };
            out.push(el("firstRowStripe", band()));
            out.push(el("firstColumnStripe", band()));
            out.push(el(
                "firstColumn",
                Dxf {
                    bold: Some(true),
                    fill: dark(-0.5),
                    sides: [None, side("medium", white()), None, None, None, None],
                    ..Default::default()
                },
            ));
            out.push(el(
                "lastColumn",
                Dxf {
                    bold: Some(true),
                    fill: dark(-0.5),
                    sides: [side("medium", white()), None, None, None, None, None],
                    ..Default::default()
                },
            ));
            out.push(el(
                "headerRow",
                Dxf {
                    bold: Some(true),
                    fill: theme(1, 0.0),
                    sides: [None, None, None, side("medium", white()), None, None],
                    ..Default::default()
                },
            ));
            out.push(el(
                "totalRow",
                Dxf {
                    bold: Some(true),
                    fill: dark(-0.5),
                    sides: [None, None, side("double", white()), None, None, None],
                    ..Default::default()
                },
            ));
        }
        ("D", _) => {
            let accent = [1, 4, 6, 8][(i as usize).min(3)];
            out.push(el(
                "wholeTable",
                Dxf {
                    fill: theme(0, -0.15),
                    ..Default::default()
                },
            ));
            let band = || Dxf {
                fill: theme(accent, 0.4),
                ..Default::default()
            };
            out.push(el("firstRowStripe", band()));
            out.push(el("firstColumnStripe", band()));
            out.push(el("firstColumn", bold(Dxf::default())));
            out.push(el("lastColumn", bold(Dxf::default())));
            out.push(el(
                "headerRow",
                Dxf {
                    bold: Some(true),
                    color: white(),
                    fill: theme(1, 0.0),
                    ..Default::default()
                },
            ));
            out.push(el(
                "totalRow",
                Dxf {
                    bold: Some(true),
                    sides: [None, None, side("double", theme(1, 0.0)), None, None, None],
                    ..Default::default()
                },
            ));
        }
        _ => return None,
    }
    Some(out)
}

// ---- a style laid over the cells

/// What a table style gives one cell. Sides: top, right, bottom, left.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Overlay {
    pub fill: Option<Color>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<Color>,
    pub sides: [Option<Side>; 4],
}

/// A rectangle of cells (0-based, inclusive): c1, r1, c2, r2.
type Area = (u32, u32, u32, u32);

impl Overlay {
    /// Lays an element's format over the cell (c, r) of the region `reg`.
    fn apply(&mut self, d: &Dxf, c: u32, r: u32, reg: Area) {
        if d.fill.is_some() {
            self.fill = d.fill.clone();
        }
        if d.bold.is_some() {
            self.bold = d.bold;
        }
        if d.italic.is_some() {
            self.italic = d.italic;
        }
        if d.color.is_some() {
            self.color = d.color.clone();
        }
        let [l, rt, t, b, v, h] = &d.sides;
        let pick = |edge: bool, outer: &Option<Side>, inner: &Option<Side>| {
            if edge { outer.clone() } else { inner.clone() }
        };
        let sides = [
            pick(r == reg.1, t, h),
            pick(c == reg.2, rt, v),
            pick(r == reg.3, b, h),
            pick(c == reg.0, l, v),
        ];
        for (i, s) in sides.into_iter().enumerate() {
            if s.is_some() {
                self.sides[i] = s;
            }
        }
    }
}

/// What the tables' styles give the cells of a sheet, by (row, column). The styles are looked
/// up in `custom` (the workbook's own) and then among the built-in ones.
/// `flags`: what the file says a table (by its name) shows — first column, last column, row
/// stripes, column stripes — when umya-spreadsheet does not read it (it reads the flags of a
/// style with a name as all off).
pub fn overlays(
    ws: &Worksheet,
    custom: &HashMap<String, StyleDef>,
    flags: &dyn Fn(&str) -> Option<Shown>,
) -> HashMap<(u32, u32), Overlay> {
    let mut out: HashMap<(u32, u32), Overlay> = HashMap::new();
    for t in ws.tables() {
        let (a, b) = t.area();
        let (c1, r1, c2, r2) = (a.col_num(), a.row_num(), b.col_num(), b.row_num());
        if c1 == 0 || r1 == 0 || c2 < c1 || r2 < r1 || (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000 {
            continue;
        }
        let (c1, r1, c2, r2) = (c1 - 1, r1 - 1, c2 - 1, r2 - 1);
        let info = t.style_info();
        let Some(name) = info.map(|i| i.name().to_string()) else {
            continue;
        };
        let Some(def) = custom.get(&name).cloned().or_else(|| builtin(&name)) else {
            continue;
        };
        let (first_col, last_col, row_stripes, col_stripes) =
            flags(t.name()).unwrap_or(match info {
                Some(i) => (
                    i.is_show_first_col(),
                    i.is_show_last_col(),
                    i.is_show_row_stripes(),
                    i.is_show_col_stripes(),
                ),
                None => (false, false, true, false),
            });
        // (The heading row's count is not kept: tables nearly always have one.)
        let totals = t.totals_row_count();
        let body_r1 = r1 + 1;
        let body_r2 = r2.saturating_sub(totals);
        let get = |ty: &str| def.iter().find(|e| e.0 == ty);
        // The elements, from the lowest priority to the highest.
        let mut regions: Vec<(&Dxf, Area)> = Vec::new();
        if let Some(e) = get("wholeTable") {
            regions.push((&e.2, (c1, r1, c2, r2)));
        }
        if col_stripes {
            let s1 = get("firstColumnStripe").map(|e| e.1.max(1)).unwrap_or(1);
            let s2 = get("secondColumnStripe").map(|e| e.1.max(1)).unwrap_or(1);
            let (mut c, mut first) = (c1, true);
            while c <= c2 {
                let end = (c + if first { s1 } else { s2 } - 1).min(c2);
                if let Some(e) = get(if first {
                    "firstColumnStripe"
                } else {
                    "secondColumnStripe"
                }) {
                    regions.push((&e.2, (c, body_r1, end, body_r2)));
                }
                c = end + 1;
                first = !first;
            }
        }
        if row_stripes {
            let s1 = get("firstRowStripe").map(|e| e.1.max(1)).unwrap_or(1);
            let s2 = get("secondRowStripe").map(|e| e.1.max(1)).unwrap_or(1);
            let (mut r, mut first) = (body_r1, true);
            while r <= body_r2 {
                let end = (r + if first { s1 } else { s2 } - 1).min(body_r2);
                if let Some(e) = get(if first {
                    "firstRowStripe"
                } else {
                    "secondRowStripe"
                }) {
                    regions.push((&e.2, (c1, r, c2, end)));
                }
                r = end + 1;
                first = !first;
            }
        }
        if last_col && let Some(e) = get("lastColumn") {
            regions.push((&e.2, (c2, r1, c2, r2)));
        }
        if first_col && let Some(e) = get("firstColumn") {
            regions.push((&e.2, (c1, r1, c1, r2)));
        }
        if let Some(e) = get("headerRow") {
            regions.push((&e.2, (c1, r1, c2, r1)));
        }
        if totals > 0
            && let Some(e) = get("totalRow")
        {
            regions.push((&e.2, (c1, r2, c2, r2)));
        }
        for (ty, (c, r)) in [("firstHeaderCell", (c1, r1)), ("lastHeaderCell", (c2, r1))] {
            if let Some(e) = get(ty) {
                regions.push((&e.2, (c, r, c, r)));
            }
        }
        if totals > 0 {
            for (ty, (c, r)) in [("firstTotalCell", (c1, r2)), ("lastTotalCell", (c2, r2))] {
                if let Some(e) = get(ty) {
                    regions.push((&e.2, (c, r, c, r)));
                }
            }
        }
        for (d, reg) in regions {
            if reg.0 > reg.2 || reg.1 > reg.3 {
                continue;
            }
            for r in reg.1..=reg.3 {
                for c in reg.0..=reg.2 {
                    out.entry((r, c)).or_default().apply(d, c, r, reg);
                }
            }
        }
    }
    out
}
