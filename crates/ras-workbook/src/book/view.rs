//! A sheet as people see it in Excel, for the grid of `ras book view` and the editor: what the
//! cells show, with their main formats (font, fill, borders, alignment), the sizes of the rows
//! and columns, merged cells and the filter.
//!
//! How Excel's formats become CSS (theme colors, the kinds of border lines, column widths in
//! pixels, the alignment of values without one) follows xljs's printing of sheets
//! (`src/api/xlsx_print.rs` there), without its pages: this is the screen, not the paper.
//! The objects of the sheet's drawing (shapes, pictures, charts, controls) are drawn over the
//! cells (`crate::view::drawing`). Of the conditional formats, color scales and data bars are
//! drawn (`super::cond`); table styles are not.

use std::collections::HashMap;

use serde_json::{Value as Json, json};
use umya_spreadsheet::EnumTrait;
use umya_spreadsheet::Style;
use umya_spreadsheet::drawingml::anchor::Anchor;

use super::*;

/// The rows and columns of a sheet drawn at most.
pub const VIEW_ROWS: u32 = 500;
pub const VIEW_COLS: u32 = 100;

fn css_color(
    c: &umya_spreadsheet::Color,
    theme: &umya_spreadsheet::drawing::Theme,
) -> Option<String> {
    if c.automatic() {
        return None;
    }
    let s = c.argb_with_theme(theme);
    if s.is_empty() || s == "00000000" || s.len() < 6 || !s.bytes().all(|b| b.is_ascii_hexdigit()) {
        return None;
    }
    Some(format!("#{}", &s[s.len() - 6..]))
}

/// A font-family list for an Excel font name (Japanese names also by their English names).
fn font_family(name: &str) -> String {
    let alias = match name {
        "游ゴシック" => "Yu Gothic",
        "游ゴシック Light" => "Yu Gothic Light",
        "游ゴシック Medium" => "Yu Gothic Medium",
        "游明朝" => "Yu Mincho",
        "ＭＳ Ｐゴシック" => "MS PGothic",
        "ＭＳ ゴシック" => "MS Gothic",
        "ＭＳ Ｐ明朝" => "MS PMincho",
        "ＭＳ 明朝" => "MS Mincho",
        "メイリオ" => "Meiryo",
        "BIZ UDゴシック" => "BIZ UDGothic",
        "BIZ UDPゴシック" => "BIZ UDPGothic",
        "BIZ UD明朝 Medium" => "BIZ UDMincho Medium",
        "BIZ UDP明朝 Medium" => "BIZ UDPMincho Medium",
        _ => "",
    };
    let quote = |s: &str| format!("'{}'", s.replace(['\'', '\\', ';', '"'], ""));
    let mut out = quote(name);
    if !alias.is_empty() {
        out.push(',');
        out.push_str(&quote(alias));
    }
    out.push_str(if name.contains("明朝") || name.contains("Mincho") {
        ",'Yu Mincho',serif"
    } else {
        ",'Yu Gothic',Meiryo,sans-serif"
    });
    out
}

/// A border line as CSS: (width px, style).
fn line_css(style: &str) -> Option<(u32, &'static str)> {
    Some(match style {
        "thin" => (1, "solid"),
        "hair" | "dotted" => (1, "dotted"),
        "dashed" | "dashDot" | "dashDotDot" => (1, "dashed"),
        "medium" => (2, "solid"),
        "mediumDashed" | "mediumDashDot" | "slantDashDot" | "mediumDashDotDot" => (2, "dashed"),
        "thick" => (3, "solid"),
        "double" => (3, "double"),
        _ => return None,
    })
}

fn n(v: f64) -> String {
    let s = format!("{v:.2}");
    s.trim_end_matches('0').trim_end_matches('.').to_string()
}

/// The CSS of a cell's formats. `kind`: what its value is (the alignment without one follows it).
/// The tile of a pattern fill, a row of pixels per string ('1': the pattern's color).
fn pattern_tile(t: &str) -> Option<&'static [&'static str]> {
    Some(match t {
        "gray0625" => &["10000000", "00000000", "00001000", "00000000"],
        "gray125" => &["1000", "0000", "0010", "0000"],
        "lightGray" => &["1000", "0010"],
        "mediumGray" => &["10", "01"],
        "darkGray" => &["0111", "1101"],
        "darkHorizontal" => &["1", "1", "0", "0"],
        "darkVertical" => &["1100"],
        "darkDown" => &["1100", "0110", "0011", "1001"],
        "darkUp" => &["0011", "0110", "1100", "1001"],
        "darkGrid" => &["1100", "1100", "0011", "0011"],
        "darkTrellis" => &["1100", "1111", "0011", "1111"],
        "lightHorizontal" => &["1", "0", "0", "0"],
        "lightVertical" => &["1000"],
        "lightDown" => &["1000", "0100", "0010", "0001"],
        "lightUp" => &["0001", "0010", "0100", "1000"],
        "lightGrid" => &["1111", "1000", "1000", "1000"],
        "lightTrellis" => &["1001", "0110", "0110", "1001"],
        _ => return None,
    })
}

/// A pattern's tile as a CSS image in `color` (#RRGGBB).
fn pattern_image(tile: &[&str], color: &str) -> String {
    let (w, h) = (tile[0].len(), tile.len());
    let mut d = String::new();
    for (y, row) in tile.iter().enumerate() {
        for (x, _) in row.bytes().enumerate().filter(|(_, b)| *b == b'1') {
            d.push_str(&format!("M{x} {y}h1v1h-1z"));
        }
    }
    format!(
        "url(\"data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' width='{w}' height='{h}' shape-rendering='crispEdges'><path fill='{}' d='{d}'/></svg>\")",
        color.replace('#', "%23")
    )
}

/// A border side as CSS (`side`: top, right, bottom, left).
fn side_css(side: &str, style: &str, color: Option<String>) -> String {
    match line_css(style) {
        Some((w, line)) => format!(
            "border-{side}:{w}px {line} {};",
            color.unwrap_or_else(|| "#000".into())
        ),
        None => String::new(),
    }
}

/// What a table's style gives a cell, as CSS put before the cell's own (which wins).
fn table_css(t: &super::tables::Overlay, theme: &umya_spreadsheet::drawing::Theme) -> String {
    let mut css = String::new();
    if let Some(c) = t.fill.as_ref().and_then(|c| css_color(c, theme)) {
        css.push_str(&format!(
            "background:{c};box-shadow:-1px 0 {c},0 -1px {c},-1px -1px {c};"
        ));
    }
    if t.bold == Some(true) {
        css.push_str("font-weight:bold;");
    }
    if t.italic == Some(true) {
        css.push_str("font-style:italic;");
    }
    if let Some(c) = t.color.as_ref().and_then(|c| css_color(c, theme)) {
        css.push_str(&format!("color:{c};"));
    }
    for (i, name) in ["top", "right", "bottom", "left"].into_iter().enumerate() {
        if let Some(s) = &t.sides[i] {
            css.push_str(&side_css(
                name,
                &s.style,
                s.color.as_ref().and_then(|c| css_color(c, theme)),
            ));
        }
    }
    css
}

/// What the conditional formats give a cell, as CSS put after the cell's own (which they
/// change).
fn cond_css(fx: &super::cond::CondFx, own: &Style) -> String {
    let mut css = String::new();
    if let Some(c) = &fx.fill {
        css.push_str(&format!(
            "background:{c};box-shadow:-1px 0 {c},0 -1px {c},-1px -1px {c};"
        ));
    }
    if let Some(c) = &fx.color {
        css.push_str(&format!("color:{c};"));
    }
    if let Some(b) = fx.bold {
        css.push_str(if b {
            "font-weight:bold;"
        } else {
            "font-weight:normal;"
        });
    }
    if let Some(i) = fx.italic {
        css.push_str(if i {
            "font-style:italic;"
        } else {
            "font-style:normal;"
        });
    }
    if fx.underline.is_some() || fx.strike.is_some() {
        let f = own.font();
        let underline = fx.underline.unwrap_or_else(|| {
            f.is_some_and(|f| {
                let u = f.font_underline().val().value_string();
                !u.is_empty() && u != "none"
            })
        });
        let strike = fx
            .strike
            .unwrap_or_else(|| f.is_some_and(|f| f.strikethrough()));
        let mut deco = Vec::new();
        if underline {
            deco.push("underline");
        }
        if strike {
            deco.push("line-through");
        }
        css.push_str(&format!(
            "text-decoration:{};",
            if deco.is_empty() {
                "none".to_string()
            } else {
                deco.join(" ")
            }
        ));
    }
    for (i, name) in ["left", "right", "top", "bottom"].into_iter().enumerate() {
        if let Some(l) = &fx.sides[i] {
            css.push_str(&side_css(name, &l.style, l.color.clone()));
        }
    }
    css
}

fn style_css(st: &Style, theme: &umya_spreadsheet::drawing::Theme, kind: &str) -> String {
    let mut css = String::new();
    if let Some(pf) = st.fill().and_then(|f| f.pattern_fill()) {
        let t = pf.pattern_type().value_string();
        if !(t == "none" || t.is_empty()) {
            let fg = pf.foreground_color().and_then(|c| css_color(c, theme));
            if t == "solid" {
                if let Some(c) = fg {
                    // Over the gridlines around it too, as Excel fills cells.
                    css.push_str(&format!(
                        "background:{c};box-shadow:-1px 0 {c},0 -1px {c},-1px -1px {c};"
                    ));
                }
            } else if let Some(tile) = pattern_tile(t) {
                // The pattern in its color (black without one) over the cell's color (white).
                let fg = fg.unwrap_or_else(|| "#000000".into());
                let bg = pf
                    .background_color()
                    .and_then(|c| css_color(c, theme))
                    .unwrap_or_else(|| "#FFFFFF".into());
                css.push_str(&format!(
                    "background:{bg} {};box-shadow:-1px 0 {bg},0 -1px {bg},-1px -1px {bg};",
                    pattern_image(tile, &fg)
                ));
            }
        }
    }
    if let Some(f) = st.font() {
        if !f.name().is_empty() {
            css.push_str(&format!("font-family:{};", font_family(f.name())));
        }
        // (Always: the grid's own font is the workbook's, of any size.)
        if f.size() > 0.0 {
            css.push_str(&format!("font-size:{}pt;", n(f.size())));
        }
        if f.bold() {
            css.push_str("font-weight:bold;");
        }
        if f.italic() {
            css.push_str("font-style:italic;");
        }
        let underline = f.font_underline().val().value_string().to_string();
        let mut deco = Vec::new();
        if !underline.is_empty() && underline != "none" {
            deco.push("underline");
        }
        if f.strikethrough() {
            deco.push("line-through");
        }
        if !deco.is_empty() {
            css.push_str(&format!("text-decoration:{};", deco.join(" ")));
        }
        if let Some(c) = css_color(f.color(), theme).filter(|c| c != "#000000") {
            css.push_str(&format!("color:{c};"));
        }
    }
    if let Some(b) = st.borders() {
        for (side, bd) in [
            ("top", b.top()),
            ("right", b.right()),
            ("bottom", b.bottom()),
            ("left", b.left()),
        ] {
            if let Some((w, line)) = line_css(bd.border_style()) {
                let color = bd
                    .color()
                    .as_ref()
                    .and_then(|c| css_color(c, theme))
                    .unwrap_or_else(|| "#000".into());
                css.push_str(&format!("border-{side}:{w}px {line} {color};"));
            }
        }
    }
    let (mut h, mut v, mut wrap, mut indent) = (String::new(), String::new(), false, 0);
    let (mut rotation, mut shrink) = (0, false);
    if let Some(a) = st.alignment() {
        h = a.horizontal().value_string().to_string();
        v = a.vertical().value_string().to_string();
        wrap = a.wrap_text();
        indent = a.indent();
        rotation = a.text_rotation();
        shrink = a.shrink_to_fit();
    }
    let align = match h.as_str() {
        "center" | "centerContinuous" => "center",
        "right" => "right",
        "left" | "fill" | "justify" | "distributed" => "left",
        // Without one: numbers and dates to the right, booleans and errors in the middle.
        _ => match kind {
            "number" | "date" => "right",
            "boolean" | "error" => "center",
            _ => "left",
        },
    };
    if align != "left" {
        css.push_str(&format!("text-align:{align};"));
    }
    match v.as_str() {
        "top" => css.push_str("vertical-align:top;"),
        "center" | "justify" | "distributed" => css.push_str("vertical-align:middle;"),
        _ => {}
    }
    if wrap {
        css.push_str("white-space:pre-wrap;overflow-wrap:anywhere;");
    } else if shrink {
        // The page makes the font smaller until the text fits.
        css.push_str("--shrink:1;");
    }
    // Text at an angle: 1-90 degrees up, 91-180 down (as 1-90), 255 stacked letters.
    match rotation {
        255 => css.push_str("writing-mode:vertical-rl;text-orientation:upright;"),
        1..=90 => css.push_str(&format!("--rot:{rotation};")),
        91..=180 => css.push_str(&format!("--rot:{};", 90_i64 - i64::from(rotation))),
        _ => {}
    }
    if indent > 0 {
        css.push_str(&format!(
            "padding-{}:{}px;",
            if align == "right" { "right" } else { "left" },
            3 + indent * 9
        ));
    }
    css
}

/// What the grid draws of a sheet's drawing: the parts of the file, and what the page shares
/// among its sheets (the ids of SVG markers, the pictures).
pub struct Drawing<'a> {
    pub parts: &'a crate::view::drawing::Parts,
    pub ids: &'a std::cell::Cell<usize>,
    pub images: &'a std::cell::RefCell<crate::view::drawing::Images>,
}

/// The sizes of a sheet's columns and rows in pixels, as Excel shows them at 100%.
pub(crate) struct Sizes {
    default_w: f64,
    default_h: f64,
    widths: HashMap<u32, f64>,
    heights: HashMap<u32, f64>,
    /// The rows (1-based) the file gives a height (or hides); the others Excel fits to their
    /// fonts. And those given it by hand (customHeight).
    fixed: std::collections::HashSet<u32>,
    custom: Vec<u32>,
}

impl Sizes {
    /// The sizes of the sheet `ws`. Column widths are kept as `ColumnWidth` has them
    /// (characters, without the padding).
    pub(crate) fn of(ws: &umya_spreadsheet::Worksheet) -> Sizes {
        let px = super::format::width_px;
        let fmt = ws.sheet_format_properties();
        let default_w = if fmt.default_column_width() > 0.0 {
            px(fmt.default_column_width())
        } else {
            72.0
        };
        let default_h = if fmt.default_row_height() > 0.0 {
            (fmt.default_row_height() * 96.0 / 72.0).round()
        } else {
            25.0
        };
        let mut widths: HashMap<u32, f64> = HashMap::new();
        for c in ws.column_dimensions() {
            widths.insert(
                c.col_num(),
                // (Width 0: hidden, as Excel shows it.)
                if c.hidden() || c.width() <= 0.0 {
                    0.0
                } else if c.width() > 0.0 {
                    px(c.width())
                } else {
                    default_w
                },
            );
        }
        let mut heights: HashMap<u32, f64> = HashMap::new();
        let mut fixed = std::collections::HashSet::new();
        let mut custom = Vec::new();
        for r in ws.row_dimensions() {
            // (Height 0 given: hidden, as Excel shows it.)
            let zero = r.custom_height() && r.height() <= 0.0;
            if r.hidden() || zero || r.height() > 0.0 {
                fixed.insert(r.row_num());
            }
            if r.custom_height() && r.row_num() > 0 {
                custom.push(r.row_num() - 1);
            }
            if r.hidden() || zero {
                heights.insert(r.row_num(), 0.0);
            } else if r.height() > 0.0 {
                heights.insert(r.row_num(), (r.height() * 96.0 / 72.0).round());
            }
        }
        widths.retain(|_, w| *w != default_w);
        heights.retain(|_, h| *h != default_h);
        custom.sort();
        Sizes {
            default_w,
            default_h,
            widths,
            heights,
            fixed,
            custom,
        }
    }

    /// The width of the column `c` (0-based).
    fn w(&self, c: u32) -> f64 {
        self.widths.get(&(c + 1)).copied().unwrap_or(self.default_w)
    }

    fn h(&self, r: u32) -> f64 {
        self.heights
            .get(&(r + 1))
            .copied()
            .unwrap_or(self.default_h)
    }

    /// The left of the column `c`: the widths before it.
    pub(crate) fn x(&self, c: u32) -> f64 {
        let custom: f64 = self
            .widths
            .iter()
            .filter(|(k, _)| **k <= c)
            .map(|(_, w)| *w)
            .sum();
        let n = self.widths.keys().filter(|k| **k <= c).count() as f64;
        custom + (c as f64 - n) * self.default_w
    }

    pub(crate) fn y(&self, r: u32) -> f64 {
        let custom: f64 = self
            .heights
            .iter()
            .filter(|(k, _)| **k <= r)
            .map(|(_, h)| *h)
            .sum();
        let n = self.heights.keys().filter(|k| **k <= r).count() as f64;
        custom + (r as f64 - n) * self.default_h
    }

    /// The box of an anchor in pixels: (left, top, right, bottom).
    fn anchor(&self, an: &Anchor) -> Option<(f64, f64, f64, f64)> {
        use umya_spreadsheet::drawingml::EMU_PX;
        let at = |m: (u32, i64, u32, i64)| {
            (
                self.x(m.0) + m.1 as f64 / EMU_PX,
                self.y(m.2) + m.3 as f64 / EMU_PX,
            )
        };
        let (x, y) = match (an.from, an.pos) {
            (Some(f), _) => at(f),
            (None, Some((x, y))) => (x as f64 / EMU_PX, y as f64 / EMU_PX),
            _ => return None,
        };
        let (x2, y2) = match (an.to, an.ext) {
            (Some(t), _) => at(t),
            (None, Some((cx, cy))) => (x + cx as f64 / EMU_PX, y + cy as f64 / EMU_PX),
            _ => return None,
        };
        (x2 >= x && y2 >= y).then_some((x, y, x2, y2))
    }

    /// Where a point (pixels) is as a drawing's anchor has it: the column, the offset in it,
    /// the row, the offset in it (pixels). On a line between two cells, the later cell.
    pub(crate) fn marker(&self, x: f64, y: f64) -> (u32, f64, u32, f64) {
        let (mut col, mut left) = (0, 0.0);
        while col + 1 < MAX_COLS && left + self.w(col) <= x {
            left += self.w(col);
            col += 1;
        }
        let (mut row, mut top) = (0, 0.0);
        while row + 1 < MAX_ROWS && top + self.h(row) <= y {
            top += self.h(row);
            row += 1;
        }
        (col, x - left, row, y - top)
    }

    /// The cell at a point (pixels).
    fn cell_at(&self, x: f64, y: f64) -> (u32, u32) {
        let (mut col, mut left) = (0, 0.0);
        while col + 1 < MAX_COLS && left + self.w(col) <= x {
            left += self.w(col);
            col += 1;
        }
        let (mut row, mut top) = (0, 0.0);
        while row + 1 < MAX_ROWS && top + self.h(row) <= y {
            top += self.h(row);
            row += 1;
        }
        (row, col)
    }
}

impl Book {
    /// The sizes of the columns and rows of the sheet `si` in pixels, as the grid draws them.
    pub(crate) fn sizes(&self, si: usize) -> Option<Sizes> {
        self.file.sheet(si).ok().map(Sizes::of)
    }

    /// The sheet for the grid: `rows` × `cols` of it at most, from its first cell, with the
    /// objects of its drawing.
    pub fn sheet_view(
        &mut self,
        si: usize,
        rows: u32,
        cols: u32,
        drawing: Option<&Drawing>,
    ) -> Result<Json> {
        let info = self.sheet_summary(si)?;
        let mut used = info
            .used_range
            .unwrap_or(RangeRef::cell(CellRef::new(0, 0)));
        // Merged cells are drawn even when they are empty.
        for m in info.merged.iter().filter_map(|m| RangeRef::parse(m)) {
            used = used.union(&m);
        }
        let (sizes, grid_lines) = {
            let Ok(ws) = self.file.sheet(si) else {
                return Err("シートがありません".into());
            };
            (Sizes::of(ws), super::sheet_view::grid_lines(ws))
        };
        // The objects of the drawing; the grid reaches to them.
        let objects = match drawing {
            Some(d) => self.drawn(si, &info.name, d, &sizes),
            None => Vec::new(),
        };
        for o in &objects {
            if let Some((_, _, x2, y2)) = sizes.anchor(&o.anchor) {
                let (r, c) = sizes.cell_at(x2, y2);
                used = used.union(&RangeRef::cell(CellRef::new(r, c)));
            }
        }
        // From A1, so that the grid reads as the sheet does.
        let end = CellRef::new(
            used.end.row.min(rows.saturating_sub(1)),
            used.end.col.min(cols.saturating_sub(1)),
        );
        let shown = RangeRef::new(CellRef::new(0, 0), end);
        let truncated = end != used.end;
        let cells = self.cells(si, shown)?;
        let mut by_at: HashMap<(u32, u32), &CellInfo> =
            cells.iter().map(|c| ((c.at.row, c.at.col), c)).collect();

        let theme = self.file.theme().clone();
        let mut styles: Vec<String> = vec![String::new()];
        let mut style_index: HashMap<String, usize> = HashMap::new();
        style_index.insert(String::new(), 0);
        let mut intern = |css: String| -> usize {
            if let Some(i) = style_index.get(&css) {
                return *i;
            }
            styles.push(css.clone());
            style_index.insert(css, styles.len() - 1);
            styles.len() - 1
        };
        // The color scales and data bars of the conditional formats.
        let cond = self.cond_effects(si, shown, &|c| css_color(c, &theme));
        // What the cells show in the number formats the conditional formats give them.
        let mut cond_texts: HashMap<(u32, u32), String> = HashMap::new();
        for ((row, col), fx) in &cond {
            if let Some(code) = &fx.num_fmt
                && by_at.contains_key(&(*row, *col))
            {
                let v = self.value(si, CellRef::new(*row, *col));
                cond_texts.insert((*row, *col), self.display(&v, code));
            }
        }
        // The icons of icon sets, once each: "set:index" -> SVG.
        let mut icons = serde_json::Map::new();
        // The workbook's own table styles.
        let custom = super::tables::custom_styles(&self.file);
        // The colors number formats give what the cells show (`[Red]`).
        let mut format_colors: HashMap<(u32, u32), String> = HashMap::new();
        for c in &cells {
            if let Some(color) = self.format_color(si, c.at) {
                format_colors.insert((c.at.row, c.at.col), color);
            }
        }
        let mut out_cells: Vec<Json> = Vec::new();
        let col_w: Vec<f64> = (0..=end.col).map(|c| sizes.w(c)).collect();
        let font = self.default_font();
        // (null: a row Excel fits to its fonts, which the page works out.)
        let row_h: Vec<Option<f64>> = (0..=end.row)
            .map(|r| sizes.fixed.contains(&(r + 1)).then(|| sizes.h(r)))
            .collect();
        {
            let Ok(ws) = self.file.sheet(si) else {
                return Err("シートがありません".into());
            };
            // The cells with a value, then those with only formats (a fill, a border).
            let plain = Style::default();
            // A cell without a format of its own has the Normal style's alignment (its style is
            // empty).
            let normal = self.file.normal_style().alignment().map(|a| {
                let mut st = Style::default();
                st.set_alignment(a.clone());
                st
            });
            // The tables' styles under the cells' own formats.
            let tables = super::tables::overlays(ws, &custom);
            let mut at: Vec<(u32, u32)> = by_at.keys().copied().collect();
            for c in ws.cells() {
                let p = at_of(c);
                if shown.contains(p) && !by_at.contains_key(&(p.row, p.col)) && *c.style() != plain
                {
                    at.push((p.row, p.col));
                }
            }
            // Empty cells a table or a conditional format draws.
            for (row, col) in tables.keys().chain(cond.keys()) {
                if shown.contains(CellRef::new(*row, *col)) {
                    at.push((*row, *col));
                }
            }
            at.sort();
            at.dedup();
            for (row, col) in at {
                let info = by_at.remove(&(row, col));
                let kind = info
                    .map(|c| crate::dump::type_name(&c.value))
                    .unwrap_or("empty");
                let mut style = match ws.cell((col + 1, row + 1)) {
                    Some(c) => c.style().clone(),
                    None => self.style_at(si, CellRef::new(row, col)),
                };
                if let Some(n) = normal.as_ref().filter(|_| style == plain) {
                    style = n.clone();
                }
                let mut css = tables
                    .get(&(row, col))
                    .map(|t| table_css(t, &theme))
                    .unwrap_or_default();
                css.push_str(&style_css(&style, &theme, kind));
                if let Some(color) = format_colors.get(&(row, col)) {
                    css.push_str(&format!("color:#{color};"));
                }
                let fx = cond.get(&(row, col));
                if let Some(fx) = fx {
                    css.push_str(&cond_css(fx, &style));
                }
                let s = intern(css);
                let mut cell = json!({ "r": row, "c": col });
                if let Some(b) = fx.and_then(|fx| fx.bar.as_ref()) {
                    // A gradient fades to white at the bar's end, as Excel draws it.
                    let bg = if b.gradient {
                        let light = super::cond::mix(&b.fill, "#FFFFFF", 0.9);
                        let to = if b.leftward { "left" } else { "right" };
                        format!("linear-gradient(to {to},{},{light})", b.fill)
                    } else {
                        b.fill.clone()
                    };
                    cell["bar"] = json!({ "x": b.start, "w": b.len, "bg": bg, "axis": b.axis });
                    if let Some(bd) = &b.border {
                        cell["bar"]["bd"] = json!(bd);
                    }
                }
                if let Some((set, i)) = fx.and_then(|fx| fx.icon.as_ref())
                    && let Some(svg) = super::cond::icon_svg(set, *i)
                {
                    let key = format!("{set}:{i}");
                    icons.entry(key.clone()).or_insert_with(|| json!(svg));
                    cell["ic"] = json!(key);
                }
                if fx.is_some_and(|fx| fx.hide_value) {
                    cell["hv"] = json!(true);
                }
                if s != 0 {
                    cell["s"] = json!(s);
                }
                if let Some(c) = info {
                    cell["t"] = json!(cond_texts.get(&(row, col)).unwrap_or(&c.text));
                    if let Some(f) = &c.formula {
                        cell["f"] = json!(f);
                    }
                    if c.number_format != "General" {
                        cell["n"] = json!(c.number_format);
                        // The format Japanese Excel shows it by, when not the same (built-in
                        // formats: `yyyy/m/d` for `m/d/yyyy`), for the status bar.
                        let shown = self.shown_format_code(si, CellRef::new(row, col));
                        if shown != c.number_format {
                            cell["sn"] = json!(shown);
                        }
                    }
                    // The number (a date's serial): for the page to fit a General number to its
                    // column, as Excel does, and to sum the cells selected (the status bar).
                    if let Val::Num(x) | Val::Date(x) = c.value {
                        cell["x"] = json!(x);
                    }
                    cell["k"] = json!(kind);
                }
                out_cells.push(cell);
            }
        }
        // The formats of whole rows and columns: their empty cells have them (a row's before a
        // column's, as Excel).
        let (mut row_styles, mut col_styles) = (serde_json::Map::new(), serde_json::Map::new());
        if let Ok(ws) = self.file.sheet(si) {
            let plain = Style::default();
            for r in ws.row_dimensions() {
                if r.row_num() > 0 && r.row_num() <= end.row + 1 + 20_000 && *r.style() != plain {
                    let s = intern(style_css(r.style(), &theme, "empty"));
                    row_styles.insert((r.row_num() - 1).to_string(), json!(s));
                }
            }
            for c in ws.column_dimensions() {
                if c.col_num() > 0 && *c.style() != plain {
                    let s = intern(style_css(c.style(), &theme, "empty"));
                    col_styles.insert((c.col_num() - 1).to_string(), json!(s));
                }
            }
        }
        // What is typed into the cells (the formula bar): dates and numbers as values, not as shown.
        for (cell, c) in out_cells
            .iter_mut()
            .filter(|c| c.get("t").is_some())
            .zip(cells.iter())
        {
            let typed = match &c.value {
                Val::Date(d) => Some(self.date_text(*d)),
                Val::Num(x) if c.text != format!("{x}") => Some(format!("{x}")),
                _ => None,
            };
            if let (Some(v), None) = (typed, &c.formula) {
                cell["v"] = json!(v);
            }
        }
        // The sizes past the cells drawn, for the empty grid the page adds as it scrolls.
        let beyond = |m: &HashMap<u32, f64>, from: u32| -> serde_json::Map<String, Json> {
            let mut keys: Vec<&u32> = m.keys().filter(|k| **k > from).collect();
            keys.sort();
            keys.into_iter()
                .take(20_000)
                .map(|k| ((k - 1).to_string(), json!(m[k])))
                .collect()
        };
        let objects: Vec<Json> = objects
            .iter()
            .map(|o| {
                use umya_spreadsheet::drawingml::EMU_PX;
                let px = |v: i64| (v as f64 / EMU_PX * 100.0).round() / 100.0;
                let marker = |m: (u32, i64, u32, i64)| json!([m.0, px(m.1), m.2, px(m.3)]);
                let mut j = json!({ "kind": o.kind, "html": o.html });
                if let Some(f) = o.anchor.from {
                    j["from"] = marker(f);
                }
                if let Some(t) = o.anchor.to {
                    j["to"] = marker(t);
                }
                if let Some((x, y)) = o.anchor.pos {
                    j["pos"] = json!([px(x), px(y)]);
                }
                if let Some((w, h)) = o.anchor.ext {
                    j["ext"] = json!([px(w), px(h)]);
                }
                if !o.name.is_empty() {
                    j["name"] = json!(o.name);
                }
                if !o.macro_name.is_empty() {
                    j["macro"] = json!(o.macro_name);
                }
                if let Some(class) = o.activex {
                    j["activex"] = json!(class);
                }
                if o.assignable {
                    j["onAction"] = json!(true);
                }
                if let Some(id) = o.shape_id {
                    j["shapeId"] = json!(id);
                }
                if let Some(text) = &o.text {
                    j["text"] = json!(text);
                }
                if !o.element.is_empty() {
                    j["el"] = json!(o.element);
                }
                if o.rot != 0.0 {
                    j["rot"] = json!(o.rot);
                }
                if o.flip.0 {
                    j["flipH"] = json!(true);
                }
                if o.flip.1 {
                    j["flipV"] = json!(true);
                }
                if let Some(f) = &o.form {
                    j["form"] = f.clone();
                }
                if let Some(c) = &o.chart {
                    j["chart"] = c.clone();
                }
                j
            })
            .collect();
        let merges: Vec<Json> = info
            .merged
            .iter()
            .filter_map(|m| RangeRef::parse(m))
            .filter(|m| shown.contains(m.start))
            .map(|m| {
                json!([
                    m.start.row,
                    m.start.col,
                    m.end.row.min(end.row),
                    m.end.col.min(end.col)
                ])
            })
            .collect();
        // The columns the filter filters by now, for the buttons of its headings.
        let filter_columns: serde_json::Map<String, Json> = self
            .model
            .sheet(si)
            .and_then(|s| s.autofilter.as_ref())
            .map(|f| {
                f.criteria
                    .iter()
                    .map(|(offset, c)| {
                        let shown = match c {
                            gridcraft_model::FilterCriterion::Values { values, blanks } => {
                                json!({ "values": values, "blanks": blanks })
                            }
                            gridcraft_model::FilterCriterion::Custom { a, b, and } => {
                                json!({ "custom": { "a": [a.0, a.1], "b": b.as_ref().map(|b| json!([b.0, b.1])), "and": and } })
                            }
                            gridcraft_model::FilterCriterion::Top10 {
                                bottom,
                                percent,
                                count,
                            } => {
                                json!({ "top": { "bottom": bottom, "percent": percent, "count": count } })
                            }
                            _ => json!({ "other": true }),
                        };
                        (offset.to_string(), shown)
                    })
                    .collect()
            })
            .unwrap_or_default();
        // The buttons of the tables' headings: [row, column, filtered (0 / 1)].
        let table_filters: Vec<Json> = match self.file.sheet(si) {
            Ok(ws) => ws
                .tables()
                .iter()
                .filter_map(super::sheet_view::table_filter)
                .flat_map(|(r, hidden, filtered)| {
                    (r.start.col..=r.end.col)
                        .map(move |c| (r.start.row, c, c - r.start.col))
                        .filter(move |(_, _, i)| !hidden.contains(i))
                        .map(move |(row, c, i)| json!([row, c, u8::from(filtered.contains(&i))]))
                })
                .collect(),
            _ => Vec::new(),
        };
        let (notes, lists) = self.notes_and_lists(si, &info.name)?;
        let links = self.links(si);
        // The panes frozen (rows, columns), as the file keeps them.
        let screen = self.file.sheet(si).ok();
        let frozen = screen.and_then(super::sheet_view::frozen);
        Ok(json!({
            "frozen": frozen.map(|(r, c)| json!([r, c])),
            "zoom": screen.and_then(super::sheet_view::zoom),
            // The active cell and the range selected when the file was saved.
            "selection": screen
                .and_then(super::sheet_view::selection)
                .map(|(a, r)| json!([a.row, a.col, r.start.row, r.start.col, r.end.row, r.end.col])),
            "notes": notes,
            "links": links,
            "lists": lists,
            "name": info.name,
            "codeName": info.code_name,
            "visible": crate::dump::visible_name(info.visible),
            "usedRange": info.used_range.map(|r| r.a1()),
            "rows": end.row + 1,
            "cols": end.col + 1,
            "truncated": truncated,
            "colWidths": col_w,
            "rowHeights": row_h,
            "rowStyles": row_styles,
            "colStyles": col_styles,
            "customRows": sizes.custom.iter().take(20_000).collect::<Vec<_>>(),
            // The workbook's font (the Normal style's): cells without a font of their own.
            "font": { "name": font.0, "size": font.1, "css": font_family(&font.0) },
            // Past the cells: the default sizes, the others, and how far the grid goes.
            "defaultWidth": sizes.default_w,
            "defaultHeight": sizes.default_h,
            "moreWidths": beyond(&sizes.widths, end.col + 1),
            "moreHeights": beyond(&sizes.heights, end.row + 1),
            "maxRows": if end.row < used.end.row { end.row + 1 } else { MAX_ROWS },
            "maxCols": if end.col < used.end.col { end.col + 1 } else { MAX_COLS },
            "gridLines": grid_lines,
            "objects": objects,
            "styles": styles,
            "icons": icons,
            "cells": out_cells,
            "merges": merges,
            "filter": info.filter,
            "filterColumns": filter_columns,
            "tableFilters": table_filters,
            "tabColor": self.tab_color(si).map(|c| { let c = c as u32; format!("#{:02x}{:02x}{:02x}", c & 0xff, (c >> 8) & 0xff, (c >> 16) & 0xff) }),
        }))
    }

    /// The notes (comments) of the sheet: `{ r, c, text, author }`; and its input rules that
    /// choose from a list, with what the list holds: `{ ranges: [[r1, c1, r2, c2]], values,
    /// title, prompt, stop, all }` (`values` absent: the list's button is hidden; `stop`: an
    /// entry not among `all` is refused, with that message, empty for Excel's own).
    /// The workbook's default font (the Normal style's name and size).
    fn default_font(&self) -> (String, f64) {
        self.file
            .default_font()
            .filter(|(name, _)| !name.is_empty())
            .map(|(name, size)| (name, if size > 0.0 { size } else { 11.0 }))
            .unwrap_or(("游ゴシック".into(), 11.0))
    }

    fn notes_and_lists(&mut self, si: usize, sheet: &str) -> Result<(Vec<Json>, Vec<Json>)> {
        let Ok(ws) = self.file.sheet(si) else {
            return Err("シートがありません".into());
        };
        let notes: Vec<Json> = ws
            .comments()
            .iter()
            .map(|c| {
                let t = c.text();
                let text = t
                    .text()
                    .map(|t| t.value().to_string())
                    .or_else(|| t.rich_text().map(|r| r.text().into_owned()))
                    .unwrap_or_default();
                json!({
                    "r": c.coordinate().row_num().saturating_sub(1),
                    "c": c.coordinate().col_num().saturating_sub(1),
                    "text": text,
                    "author": c.author(),
                })
            })
            .collect();
        // (formula1, ranges, the button shown, title, prompt, the error that stops an entry)
        type Rule = (String, Vec<RangeRef>, bool, String, String, Option<String>);
        let mut rules: Vec<Rule> = Vec::new();
        if let Some(dv) = ws.data_validations() {
            for v in dv.data_validation_list() {
                if v.get_type().value_string() != "list" {
                    continue;
                }
                let ranges = v
                    .sequence_of_references()
                    .range_collection()
                    .iter()
                    .filter_map(|r| RangeRef::parse(&r.range().replace('$', "")))
                    .collect();
                let (title, prompt) = if v.show_input_message() {
                    (v.prompt_title().to_string(), v.prompt().to_string())
                } else {
                    (String::new(), String::new())
                };
                // An entry not in the list is refused (the style "stop"; the others ask, and
                // Excel enters it when told to).
                let stop = (v.show_error_message() && v.error_style().value_string() == "stop")
                    .then(|| v.error_message().to_string());
                // (`showDropDown` is true when the button is *not* shown.)
                rules.push((
                    v.formula1().to_string(),
                    ranges,
                    !v.show_drop_down(),
                    title,
                    prompt,
                    stop,
                ));
            }
        }
        let mut lists = Vec::new();
        for (formula, ranges, button, title, prompt, stop) in rules {
            let mut j = json!({
                "ranges": ranges
                    .iter()
                    .map(|r| json!([r.start.row, r.start.col, r.end.row, r.end.col]))
                    .collect::<Vec<_>>(),
            });
            let values = self.list_values(si, sheet, &formula);
            if let Some(message) = stop {
                j["stop"] = json!(message);
                j["all"] = json!(values);
            }
            if button {
                j["values"] = json!(values);
            }
            if !title.is_empty() {
                j["title"] = json!(title);
            }
            if !prompt.is_empty() {
                j["prompt"] = json!(prompt);
            }
            lists.push(j);
        }
        Ok((notes, lists))
    }

    /// The hyperlinks of the sheet's cells: `{ r, c, url, tip }`, or `{ r, c, to, tip }` for a
    /// place in the workbook (`Sheet2!A1`, a name).
    fn links(&self, si: usize) -> Vec<Json> {
        let Ok(ws) = self.file.sheet(si) else {
            return Vec::new();
        };
        let mut out: Vec<Json> = ws
            .cells()
            .iter()
            .filter_map(|c| {
                let h = c.hyperlink()?;
                if h.url().is_empty() {
                    return None;
                }
                let at = at_of(c);
                let mut j = json!({ "r": at.row, "c": at.col });
                j[if h.location() { "to" } else { "url" }] = json!(h.url());
                if !h.tooltip().is_empty() {
                    j["tip"] = json!(h.tooltip());
                }
                Some(j)
            })
            .collect();
        out.sort_by_key(|j| (j["r"].as_u64(), j["c"].as_u64()));
        out
    }

    /// What a list of an input rule offers: `"a,b,c"`, or the cells of a reference or a name
    /// (`$A$1:$A$5`, `Sheet2!$A$1:$A$3`, `Items`), at most 1000 of them, empty ones left out.
    fn list_values(&mut self, si: usize, sheet: &str, formula: &str) -> Vec<String> {
        let f = formula.trim().trim_start_matches('=');
        if let Some(inner) = f.strip_prefix('"') {
            let inner = inner
                .strip_suffix('"')
                .unwrap_or(inner)
                .replace("\"\"", "\"");
            return inner
                .split(',')
                .map(str::trim)
                .filter(|s| !s.is_empty())
                .map(str::to_string)
                .collect();
        }
        // A name: what it refers to (the sheet's own first).
        let mut text = f.to_string();
        let names = self.names();
        let named = names
            .iter()
            .filter(|(n, _, _)| n.eq_ignore_ascii_case(f))
            .min_by_key(|(_, scope, _)| if *scope == Some(si) { 0 } else { 1 });
        if let Some((_, _, refers)) = named {
            text = refers.trim().trim_start_matches('=').to_string();
        }
        let (on, body) = gridcraft_core::split_sheet(&text);
        let at = match on {
            Some(s) => match self.sheet_index(s.trim_start_matches("[0]")) {
                Some(i) => i,
                None => return Vec::new(),
            },
            None => self.sheet_index(sheet).unwrap_or(si),
        };
        let Some(range) = RangeRef::parse(&body.replace('$', "")).filter(|r| r.count() <= 100_000)
        else {
            return Vec::new();
        };
        let Ok(cells) = self.cells(at, range) else {
            return Vec::new();
        };
        let mut cells: Vec<CellInfo> = cells;
        cells.sort_by_key(|c| (c.at.row, c.at.col));
        cells
            .into_iter()
            .map(|c| c.text)
            .filter(|t| !t.is_empty())
            .take(1000)
            .collect()
    }

    /// The objects of the sheet's drawing, with what its controls need of the cells.
    fn drawn(
        &mut self,
        si: usize,
        sheet: &str,
        d: &Drawing,
        sizes: &Sizes,
    ) -> Vec<crate::view::drawing::Obj> {
        use crate::view::drawing::{Env, sheet_objects};
        let font = self.default_font();
        let me = std::cell::RefCell::new(&mut *self);
        // A reference of a control (`$A$1:$A$5`, `Sheet2!A1:A3`): its sheet and its range.
        let resolve = |r: &str| -> Option<(usize, RangeRef)> {
            let (sheet, body) = gridcraft_core::split_sheet(r.trim().trim_start_matches('='));
            let at = match sheet {
                Some(s) => me.borrow().sheet_index(s.trim_start_matches("[0]"))?,
                None => si,
            };
            Some((at, RangeRef::parse(&body.replace('$', ""))?))
        };
        let texts = |r: &str| -> Vec<String> {
            let Some((at, range)) = resolve(r).filter(|(_, g)| g.count() <= 1000) else {
                return Vec::new();
            };
            let Ok(cells) = me.borrow_mut().cells(at, range) else {
                return Vec::new();
            };
            let by: HashMap<(u32, u32), String> = cells
                .into_iter()
                .map(|c| ((c.at.row, c.at.col), c.text))
                .collect();
            range
                .iter()
                .map(|p| by.get(&(p.row, p.col)).cloned().unwrap_or_default())
                .collect()
        };
        let value = |r: &str| -> Option<String> {
            let (at, range) = resolve(r)?;
            let cells = me
                .borrow_mut()
                .cells(at, RangeRef::cell(range.start))
                .ok()?;
            cells.into_iter().next().map(|c| match c.value {
                Val::Bool(b) => {
                    if b {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    }
                }
                _ => c.text,
            })
        };
        let range = |r: &str| -> Option<Vec<Vec<crate::view::chart::Point>>> {
            use crate::view::chart::Point;
            let (at, range) = resolve(r).filter(|(_, g)| g.count() <= 100_000)?;
            let cells = me.borrow_mut().cells(at, range).ok()?;
            let by: HashMap<(u32, u32), CellInfo> = cells
                .into_iter()
                .map(|c| ((c.at.row, c.at.col), c))
                .collect();
            Some(
                (range.start.row..=range.end.row)
                    .map(|row| {
                        (range.start.col..=range.end.col)
                            .map(|col| match by.get(&(row, col)) {
                                Some(c) => Point {
                                    number: match c.value {
                                        Val::Num(x) | Val::Date(x) => Some(x),
                                        _ => None,
                                    },
                                    text: c.text.clone(),
                                    format: c.number_format.clone(),
                                },
                                None => Point::default(),
                            })
                            .collect()
                    })
                    .collect(),
            )
        };
        let size = |an: &Anchor| {
            sizes
                .anchor(an)
                .map(|(x, y, x2, y2)| ((x2 - x) * 0.75, (y2 - y) * 0.75))
                .filter(|(w, h)| *w > 0.0 && *h > 0.0)
        };
        let env = Env {
            size: &size,
            texts: &texts,
            value: &value,
            range: &range,
            font,
        };
        sheet_objects(d.parts, sheet, &env, d.ids, d.images)
    }
}
