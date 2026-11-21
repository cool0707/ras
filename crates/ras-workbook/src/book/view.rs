//! A sheet as people see it in Excel, for the grid of `ras book view` and the editor: what the
//! cells show, with their main formats (font, fill, borders, alignment), the sizes of the rows
//! and columns, merged cells and the filter.
//!
//! How Excel's formats become CSS (theme colors, the kinds of border lines, column widths in
//! pixels, the alignment of values without one) follows xljs's printing of sheets
//! (`src/api/xlsx_print.rs` there), without its pages: this is the screen, not the paper.
//! Conditional formats, table styles, pictures and charts are not drawn.

use std::collections::HashMap;

use serde_json::{Value as Json, json};
use umya_spreadsheet::EnumTrait;
use umya_spreadsheet::Style;

use super::*;

/// The rows and columns of a sheet drawn at most.
pub const VIEW_ROWS: u32 = 500;
pub const VIEW_COLS: u32 = 100;

fn css_color(c: &umya_spreadsheet::Color, theme: &umya_spreadsheet::drawing::Theme) -> Option<String> {
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
        _ => "",
    };
    let quote = |s: &str| format!("'{}'", s.replace(['\'', '\\', ';', '"'], ""));
    let mut out = quote(name);
    if !alias.is_empty() {
        out.push(',');
        out.push_str(&quote(alias));
    }
    out.push_str(if name.contains("明朝") || name.contains("Mincho") { ",'Yu Mincho',serif" } else { ",'Yu Gothic',Meiryo,sans-serif" });
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
fn style_css(st: &Style, theme: &umya_spreadsheet::drawing::Theme, kind: &str) -> String {
    let mut css = String::new();
    if let Some(pf) = st.fill().and_then(|f| f.pattern_fill()) {
        let t = pf.pattern_type().value_string();
        if !(t == "none" || t.is_empty()) {
            let fg = pf.foreground_color().and_then(|c| css_color(c, theme));
            let color = if t == "solid" { fg } else { fg.or_else(|| pf.background_color().and_then(|c| css_color(c, theme))) };
            if let Some(c) = color.filter(|_| t != "gray125") {
                css.push_str(&format!("background:{c};"));
            }
        }
    }
    if let Some(f) = st.font() {
        if !f.name().is_empty() {
            css.push_str(&format!("font-family:{};", font_family(f.name())));
        }
        if f.size() > 0.0 && f.size() != 11.0 {
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
        for (side, bd) in [("top", b.top()), ("right", b.right()), ("bottom", b.bottom()), ("left", b.left())] {
            if let Some((w, line)) = line_css(bd.border_style()) {
                let color = bd.color().as_ref().and_then(|c| css_color(c, theme)).unwrap_or_else(|| "#000".into());
                css.push_str(&format!("border-{side}:{w}px {line} {color};"));
            }
        }
    }
    let (mut h, mut v, mut wrap, mut indent) = (String::new(), String::new(), false, 0);
    if let Some(a) = st.alignment() {
        h = a.horizontal().value_string().to_string();
        v = a.vertical().value_string().to_string();
        wrap = a.wrap_text();
        indent = a.indent();
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
    }
    if indent > 0 {
        css.push_str(&format!("padding-{}:{}px;", if align == "right" { "right" } else { "left" }, 3 + indent * 9));
    }
    css
}

impl Book {
    /// The sheet for the grid: `rows` × `cols` of it at most, from its first cell.
    pub fn sheet_view(&mut self, si: usize, rows: u32, cols: u32) -> Result<Json> {
        let info = self.sheet_summary(si)?;
        let mut used = info.used_range.unwrap_or(RangeRef::cell(CellRef::new(0, 0)));
        // Merged cells are drawn even when they are empty.
        for m in info.merged.iter().filter_map(|m| RangeRef::parse(m)) {
            used = used.union(&m);
        }
        // From A1, so that the grid reads as the sheet does.
        let end = CellRef::new(used.end.row.min(rows.saturating_sub(1)), used.end.col.min(cols.saturating_sub(1)));
        let shown = RangeRef::new(CellRef::new(0, 0), end);
        let truncated = end != used.end;
        let cells = self.cells(si, shown)?;
        let mut by_at: HashMap<(u32, u32), &CellInfo> = cells.iter().map(|c| ((c.at.row, c.at.col), c)).collect();

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
        let mut out_cells: Vec<Json> = Vec::new();
        // Column widths (in characters) to pixels as Excel does, with the widest digit 8 px wide
        // (the default Japanese fonts at 11pt).
        let mdw = 8.0_f64;
        let px = |w: f64| ((256.0 * w + (128.0 / mdw).trunc()) / 256.0 * mdw).trunc();
        let (mut col_w, mut row_h): (Vec<f64>, Vec<f64>) = (Vec::new(), Vec::new());
        {
            let Ok(ws) = self.file.sheet(si) else { return Err("シートがありません".into()) };
            let fmt = ws.sheet_format_properties();
            let default_w = if fmt.default_column_width() > 0.0 { px(fmt.default_column_width()) } else { 72.0 };
            let default_h = if fmt.default_row_height() > 0.0 { (fmt.default_row_height() * 96.0 / 72.0).round() } else { 25.0 };
            let mut widths: HashMap<u32, f64> = HashMap::new();
            for c in ws.column_dimensions() {
                widths.insert(c.col_num(), if c.hidden() { 0.0 } else if c.width() > 0.0 { px(c.width()) } else { default_w });
            }
            let mut heights: HashMap<u32, f64> = HashMap::new();
            for r in ws.row_dimensions() {
                if r.hidden() {
                    heights.insert(r.row_num(), 0.0);
                } else if r.height() > 0.0 {
                    heights.insert(r.row_num(), (r.height() * 96.0 / 72.0).round());
                }
            }
            col_w.extend((0..=end.col).map(|c| widths.get(&(c + 1)).copied().unwrap_or(default_w)));
            row_h.extend((0..=end.row).map(|r| heights.get(&(r + 1)).copied().unwrap_or(default_h)));

            // The cells with a value, then those with only formats (a fill, a border).
            let plain = Style::default();
            let mut at: Vec<(u32, u32)> = by_at.keys().copied().collect();
            for c in ws.cells() {
                let p = at_of(c);
                if shown.contains(p) && !by_at.contains_key(&(p.row, p.col)) && *c.style() != plain {
                    at.push((p.row, p.col));
                }
            }
            at.sort();
            for (row, col) in at {
                let info = by_at.remove(&(row, col));
                let kind = info.map(|c| crate::dump::type_name(&c.value)).unwrap_or("empty");
                let style = match ws.cell((col + 1, row + 1)) {
                    Some(c) => c.style().clone(),
                    None => self.style_at(si, CellRef::new(row, col)),
                };
                let s = intern(style_css(&style, &theme, kind));
                let mut cell = json!({ "r": row, "c": col });
                if s != 0 {
                    cell["s"] = json!(s);
                }
                if let Some(c) = info {
                    cell["t"] = json!(c.text);
                    if let Some(f) = &c.formula {
                        cell["f"] = json!(f);
                    }
                    if c.number_format != "General" {
                        cell["n"] = json!(c.number_format);
                    }
                    cell["k"] = json!(kind);
                }
                out_cells.push(cell);
            }
        }
        // What is typed into the cells (the formula bar): dates and numbers as values, not as shown.
        for (cell, c) in out_cells.iter_mut().filter(|c| c.get("t").is_some()).zip(cells.iter()) {
            let typed = match &c.value {
                Val::Date(d) => Some(self.date_text(*d)),
                Val::Num(x) if c.text != format!("{x}") => Some(format!("{x}")),
                _ => None,
            };
            if let (Some(v), None) = (typed, &c.formula) {
                cell["v"] = json!(v);
            }
        }
        let merges: Vec<Json> = info.merged.iter().filter_map(|m| RangeRef::parse(m)).filter(|m| shown.contains(m.start)).map(|m| json!([m.start.row, m.start.col, m.end.row.min(end.row), m.end.col.min(end.col)])).collect();
        Ok(json!({
            "name": info.name,
            "codeName": info.code_name,
            "visible": crate::dump::visible_name(info.visible),
            "usedRange": info.used_range.map(|r| r.a1()),
            "rows": end.row + 1,
            "cols": end.col + 1,
            "truncated": truncated,
            "colWidths": col_w,
            "rowHeights": row_h,
            "styles": styles,
            "cells": out_cells,
            "merges": merges,
            "filter": info.filter,
            "tabColor": self.tab_color(si).map(|c| { let c = c as u32; format!("#{:02x}{:02x}{:02x}", c & 0xff, (c >> 8) & 0xff, (c >> 16) & 0xff) }),
        }))
    }
}
