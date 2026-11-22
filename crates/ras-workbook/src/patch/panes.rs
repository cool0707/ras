//! The panes of a sheet frozen ("ウィンドウ枠の固定"), or no longer: the `pane` of the sheet's
//! first `sheetView`, and its selections, written as Excel writes them.

use std::collections::{BTreeMap, HashSet};

use gridcraft_core::CellRef;

use super::scan::{self, Edits, El};
use super::{Pack, Result, WORKBOOK, sheet_parts};

/// The file with the rows above `at` and the columns to its left frozen on the sheet `sheet`
/// (None, or A1: none). The sheet is shown from A1.
pub(super) fn freeze(bytes: &[u8], sheet: &str, at: Option<CellRef>) -> Result<Vec<u8>> {
    let at = at.filter(|a| a.row > 0 || a.col > 0);
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let part = parts
        .get(sheet)
        .or_else(|| {
            parts
                .iter()
                .find(|(n, _)| n.to_lowercase() == sheet.to_lowercase())
                .map(|(_, p)| p)
        })
        .and_then(|p| pack.find(p))
        .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
    let doc = pack
        .text(&part)?
        .ok_or_else(|| format!("{part} を読めません"))?;
    let root = scan::tree(&doc, &["sheetData", "extLst"])
        .ok_or_else(|| format!("{part} を読み取れません"))?;
    let p = scan::prefix(root.name);
    let pane = at.map(|a| pane_xml(p, a)).unwrap_or_default();
    let mut edits = Edits::default();
    match root.child("sheetViews").and_then(|v| v.child("sheetView")) {
        Some(view) => rewrite(&doc, view, &pane, at, &mut edits),
        None if at.is_none() => return Ok(bytes.to_vec()),
        None => {
            // Before everything but sheetPr and dimension, as the schema orders them.
            let after = root
                .children
                .iter()
                .take_while(|c| matches!(c.local(), "sheetPr" | "dimension"))
                .last()
                .map_or(root.open.end, |c| c.end);
            edits.insert(
                after,
                format!("<{p}sheetViews><{p}sheetView workbookViewId=\"0\">{pane}</{p}sheetView></{p}sheetViews>"),
            );
        }
    }
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut replace = BTreeMap::new();
    replace.insert(part, edits.apply(&doc).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// The pane and the selections of a sheet frozen at `at`, as Excel writes them.
fn pane_xml(p: &str, at: CellRef) -> String {
    let a = at.a1();
    let (x, y) = (at.col, at.row);
    let active = match (x > 0, y > 0) {
        (true, true) => "bottomRight",
        (false, true) => "bottomLeft",
        _ => "topRight",
    };
    let mut s = format!("<{p}pane");
    if x > 0 {
        s.push_str(&format!(" xSplit=\"{x}\""));
    }
    if y > 0 {
        s.push_str(&format!(" ySplit=\"{y}\""));
    }
    s.push_str(&format!(
        " topLeftCell=\"{a}\" activePane=\"{active}\" state=\"frozen\"/>"
    ));
    if x > 0 && y > 0 {
        s.push_str(&format!(
            "<{p}selection pane=\"topRight\"/><{p}selection pane=\"bottomLeft\"/>"
        ));
    }
    s.push_str(&format!(
        "<{p}selection pane=\"{active}\" activeCell=\"{a}\" sqref=\"{a}\"/>"
    ));
    s
}

/// The sheet's view with `pane` in place of its pane and selections (the other children stay).
fn rewrite(doc: &str, view: &El, pane: &str, at: Option<CellRef>, edits: &mut Edits) {
    // Shown from A1: the rows frozen are those from the top.
    if scan::attr(doc, view.open.clone(), "topLeftCell").is_some() {
        scan::set_attr(doc, view.open.clone(), "topLeftCell", "A1", edits);
    }
    if view.open.end == view.end {
        // `<sheetView .../>`: made a start tag with the pane in it.
        if at.is_some() {
            // (The tag on its own, as a document: its topLeftCell made A1.)
            let tag = &doc[view.open.clone()];
            let mut inner = Edits::default();
            if scan::attr(tag, 0..tag.len(), "topLeftCell").is_some() {
                scan::set_attr(tag, 0..tag.len(), "topLeftCell", "A1", &mut inner);
            }
            let tag = inner.apply(tag);
            let head = tag.trim_end_matches('>').trim_end_matches('/').trim_end();
            edits.replace(view.open.clone(), format!("{head}>{pane}</{}>", view.name));
        }
        return;
    }
    let mut first = true;
    for child in &view.children {
        if matches!(child.local(), "pane" | "selection") {
            edits.replace(child.span(), if first { pane } else { "" });
            first = false;
        }
    }
    if first {
        edits.insert(view.inner.start, pane);
    }
}
