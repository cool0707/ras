//! `ras book view`: a workbook drawn as a grid, the way Excel shows it, in one HTML page — for
//! people working without Excel. The cells a run changed are marked when another workbook is
//! given to compare with. The editor shows the same page in a panel.

use std::cell::{Cell, RefCell};

use serde_json::{Value as Json, json};

use crate::book::Book;
use crate::book::view::Drawing;
use crate::diff::{self, Diff};

pub mod chart;
mod controls;
pub mod drawing;
mod emf;
mod geom;
#[cfg(test)]
mod tests;
mod xml;

pub use drawing::Parts;

const CSS: &str = include_str!("view/view.css");
const JS: &str = include_str!("view/view.js");
const CHART_JS: &str = include_str!("view/chart.js");

/// What the page draws: the sheets of `book` with the drawings of the workbook file `source`
/// (its bytes) and, with `changes`, what differs from another workbook.
pub fn data(
    book: &mut Book,
    file: &str,
    source: Option<&[u8]>,
    limits: Option<(u32, u32)>,
    changes: Option<(&str, &Diff)>,
) -> Result<Json, String> {
    let (rows, cols) = limits.unwrap_or((crate::book::MAX_ROWS, crate::book::MAX_COLS));
    let parts = source.map(Parts::read);
    let (ids, images) = (Cell::new(0), RefCell::new(drawing::Images::default()));
    let drawing = parts.as_ref().map(|parts| Drawing {
        parts,
        ids: &ids,
        images: &images,
    });
    let mut sheets = Vec::new();
    for si in 0..book.sheet_count() {
        sheets.push(book.sheet_view(si, rows, cols, drawing.as_ref())?);
    }
    let diff = changes.map(|(before, d)| {
        let mut by_sheet = serde_json::Map::new();
        for c in &d.cells {
            let side = |s: &Option<diff::Side>| {
                s.as_ref()
                    .map(diff::Side::shown)
                    .unwrap_or_else(|| "（空）".into())
            };
            let list = by_sheet.entry(c.sheet.clone()).or_insert_with(|| json!([]));
            if let Some(list) = list.as_array_mut() {
                list.push(
                    json!({ "a": c.at.a1(), "before": side(&c.before), "after": side(&c.after) }),
                );
            }
        }
        json!({ "before": before, "structure": diff::structure_lines(d), "changes": by_sheet })
    });
    // The names defined, for the name box and "ジャンプ" (F5): those of a sheet with its name.
    let sheet_names: Vec<String> = sheets
        .iter()
        .map(|s| s["name"].as_str().unwrap_or_default().to_string())
        .collect();
    let names: Vec<Json> = book
        .names()
        .into_iter()
        // Excel's own names (_xlnm.Print_Area, _FilterDatabase) are not offered.
        .filter(|(n, _, _)| !n.starts_with("_xlnm.") && !n.starts_with("_Filter"))
        .map(|(name, scope, refers)| {
            let mut j = json!({ "name": name, "ref": refers.trim_start_matches('=') });
            if let Some(s) = scope.and_then(|i| sheet_names.get(i)) {
                j["sheet"] = json!(s);
            }
            j
        })
        .collect();
    Ok(
        json!({ "file": file, "sheets": sheets, "names": names, "diff": diff, "active": Json::Null, "images": images.into_inner().list }),
    )
}

fn escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

/// The page for `data`: everything in one file (no network, no other files).
pub fn page(data: &Json) -> String {
    // `</script>` inside the data would end the script element.
    let embedded = data
        .to_string()
        .replace("</", "<\\/")
        .replace("<!--", "<\\!--");
    let title = escape(data["file"].as_str().unwrap_or("workbook"));
    format!(
        "<!doctype html>\n<html lang=\"ja\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>{title}</title>\n<style>\n{CSS}</style>\n</head>\n<body>\n\
         <div id=\"top\"><span class=\"file\" id=\"file\"></span><span class=\"note\" id=\"summary\"></span></div>\n\
         <div id=\"bar\"><input id=\"name\" type=\"text\" list=\"names\" autocomplete=\"off\" spellcheck=\"false\" aria-label=\"名前ボックス\" title=\"セル番地や名前を入力して Enter で移動します（Ctrl+G でジャンプ）\"><datalist id=\"names\"></datalist><div id=\"fx\"></div><button id=\"formulas\" type=\"button\" title=\"セルに値ではなく数式を表示します（Ctrl+` / Ctrl+Shift+@）\" aria-pressed=\"false\">数式の表示</button></div>\n\
         <div id=\"main\"><div id=\"grid\"></div><div id=\"side\"></div></div>\n\
         <div id=\"note\" hidden></div>\n<div id=\"tabs\"></div>\n<div id=\"status\"><span id=\"mode\"></span><span id=\"stats\"></span></div>\n\
         <script>window.DATA = {embedded};</script>\n<script>\n{CHART_JS}</script>\n<script>\n{JS}</script>\n</body>\n</html>\n"
    )
}
