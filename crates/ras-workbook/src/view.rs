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
#[cfg(test)]
pub(crate) mod tests;
pub(crate) mod xml;

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

/// The functions the stand-in works out, for completing formulas typed into the grid:
/// `[name, signature, description]` by name.
pub fn functions() -> Json {
    let mut list: Vec<(&str, &str, &str)> = gridcraft_functions::all()
        .iter()
        .map(|f| (f.name, f.signature, f.description))
        .collect();
    // Those the evaluator works out itself (they take references, names, LAMBDAs).
    for name in gridcraft_calc::SPECIAL_FUNCTIONS {
        if gridcraft_functions::lookup(name).is_some() {
            continue;
        }
        let (sig, desc) = SPECIAL
            .iter()
            .find(|(n, _, _)| n == name)
            .map(|(_, s, d)| (*s, *d))
            .unwrap_or(("", ""));
        list.push((name, sig, desc));
    }
    list.sort_by_key(|f| f.0);
    list.dedup_by_key(|f| f.0);
    Json::Array(
        list.into_iter()
            .map(|(name, sig, desc)| {
                let sig = if sig.is_empty() {
                    format!("{name}(...)")
                } else {
                    sig.to_string()
                };
                json!([name, sig, desc])
            })
            .collect(),
    )
}

/// The signatures of the functions the evaluator works out itself, as Excel's help gives them.
const SPECIAL: &[(&str, &str, &str)] = &[
    (
        "AGGREGATE",
        "AGGREGATE(function_num, options, ref1, [ref2], ...)",
        "Returns an aggregate in a list or database, optionally ignoring hidden rows and errors.",
    ),
    (
        "AREAS",
        "AREAS(reference)",
        "Returns the number of areas in a reference.",
    ),
    (
        "BYCOL",
        "BYCOL(array, lambda)",
        "Applies a LAMBDA to each column and returns an array of the results.",
    ),
    (
        "BYROW",
        "BYROW(array, lambda)",
        "Applies a LAMBDA to each row and returns an array of the results.",
    ),
    (
        "CELL",
        "CELL(info_type, [reference])",
        "Returns information about the formatting, location or contents of a cell.",
    ),
    (
        "COLUMN",
        "COLUMN([reference])",
        "Returns the column number of a reference.",
    ),
    (
        "FORMULATEXT",
        "FORMULATEXT(reference)",
        "Returns the formula at the given reference as text.",
    ),
    (
        "INDIRECT",
        "INDIRECT(ref_text, [a1])",
        "Returns the reference specified by a text string.",
    ),
    (
        "ISFORMULA",
        "ISFORMULA(reference)",
        "Returns TRUE if the reference is to a cell that contains a formula.",
    ),
    (
        "ISOMITTED",
        "ISOMITTED(argument)",
        "Checks whether an argument of a LAMBDA is missing.",
    ),
    (
        "ISREF",
        "ISREF(value)",
        "Returns TRUE if the value is a reference.",
    ),
    (
        "LAMBDA",
        "LAMBDA([parameter1, ...], calculation)",
        "Creates a function that can be called with arguments.",
    ),
    (
        "LET",
        "LET(name1, name_value1, calculation_or_name2, ...)",
        "Assigns names to calculation results.",
    ),
    (
        "MAKEARRAY",
        "MAKEARRAY(rows, cols, lambda)",
        "Returns an array of the given size, its values worked out by a LAMBDA.",
    ),
    (
        "MAP",
        "MAP(array1, lambda_or_array2, ...)",
        "Returns an array of the values of the arrays mapped by a LAMBDA.",
    ),
    (
        "OFFSET",
        "OFFSET(reference, rows, cols, [height], [width])",
        "Returns a reference offset from a given reference.",
    ),
    (
        "PHONETIC",
        "PHONETIC(reference)",
        "Extracts the phonetic (furigana) characters from a text string.",
    ),
    (
        "REDUCE",
        "REDUCE([initial_value], array, lambda)",
        "Reduces an array to an accumulated value by applying a LAMBDA to each value.",
    ),
    (
        "ROW",
        "ROW([reference])",
        "Returns the row number of a reference.",
    ),
    (
        "SCAN",
        "SCAN([initial_value], array, lambda)",
        "Applies a LAMBDA to each value of an array and returns the values in between.",
    ),
    (
        "SHEET",
        "SHEET([value])",
        "Returns the sheet number of the referenced sheet.",
    ),
    (
        "SHEETS",
        "SHEETS([reference])",
        "Returns the number of sheets in a reference.",
    ),
    (
        "SUBTOTAL",
        "SUBTOTAL(function_num, ref1, [ref2], ...)",
        "Returns a subtotal in a list or database.",
    ),
    (
        "TEXT",
        "TEXT(value, format_text)",
        "Formats a number and converts it to text.",
    ),
    (
        "TRIMRANGE",
        "TRIMRANGE(range, [trim_rows], [trim_cols])",
        "Leaves out the empty rows and columns at the edges of a range.",
    ),
];

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
         <div id=\"ribbon\"></div>\n\
         <div id=\"bar\"><input id=\"name\" type=\"text\" list=\"names\" autocomplete=\"off\" spellcheck=\"false\" aria-label=\"名前ボックス\" title=\"セル番地や名前を入力して Enter で移動します（Ctrl+G でジャンプ）\"><datalist id=\"names\"></datalist><span id=\"fxlabel\" aria-hidden=\"true\">fx</span><div id=\"fx\"></div></div>\n\
         <div id=\"main\"><div id=\"grid\"></div><div id=\"side\"></div></div>\n\
         <div id=\"note\" hidden></div>\n<div id=\"tabs\"></div>\n<div id=\"status\"><span id=\"mode\"></span><span id=\"summary\"></span><span id=\"stats\"></span><span id=\"zoom\"></span></div>\n\
         <script>window.DATA = {embedded};</script>\n<script>\n{CHART_JS}</script>\n<script>\n{JS}</script>\n</body>\n</html>\n"
    )
}
