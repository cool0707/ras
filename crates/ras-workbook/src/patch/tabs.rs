//! The sheets of the workbook as their tabs change them: added, copied, removed and renamed
//! (as `ras sheet` does it in the file, [`ras_ovba::edit_sheet_bytes`]), and the color of a tab
//! ("シート見出しの色": the `tabColor` of the sheet's `sheetPr`).

use std::collections::{BTreeMap, HashSet};

use ras_ovba::{OvbaError, SheetEdit};

use super::scan::{self, Edits};
use super::{Pack, Result, WORKBOOK, sheet_parts};

/// The file with the sheets changed as `edit` says.
pub(super) fn edit(bytes: &[u8], edit: &SheetEdit) -> Result<Vec<u8>> {
    match ras_ovba::edit_sheet_bytes(bytes.to_vec(), edit) {
        Ok((Some(out), _)) => Ok(out),
        Ok((None, _)) => Ok(bytes.to_vec()),
        Err(OvbaError::Sheet(e)) => Err(e.japanese()),
        Err(e) => Err(format!("シートを変更できませんでした: {e}")),
    }
}

/// The file with the tab of the sheet `sheet` given the color `color` (`#RRGGBB`; None: none).
pub(super) fn color(bytes: &[u8], sheet: &str, color: Option<&str>) -> Result<Vec<u8>> {
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
    let tab = color.map(|c| {
        format!(
            "<{p}tabColor rgb=\"FF{}\"/>",
            c.trim_start_matches('#').to_uppercase()
        )
    });
    let mut edits = Edits::default();
    match root.child("sheetPr") {
        // `<sheetPr .../>`: made a start tag with the color in it.
        Some(pr) if pr.open.end == pr.end => {
            if let Some(tab) = tab {
                let head = doc[pr.open.clone()]
                    .trim_end_matches('>')
                    .trim_end_matches('/')
                    .trim_end();
                edits.replace(pr.open.clone(), format!("{head}>{tab}</{}>", pr.name));
            }
        }
        // The first of its children, as the schema orders them.
        Some(pr) => match (pr.child("tabColor"), tab) {
            (Some(old), tab) => edits.replace(old.span(), tab.unwrap_or_default()),
            (None, Some(tab)) => edits.insert(pr.inner.start, tab),
            (None, None) => {}
        },
        // The first child of the worksheet.
        None => {
            if let Some(tab) = tab {
                edits.insert(root.open.end, format!("<{p}sheetPr>{tab}</{p}sheetPr>"));
            }
        }
    }
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut replace = BTreeMap::new();
    replace.insert(part, edits.apply(&doc).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// A color as the edit gives it: `#RRGGBB` (`RRGGBB`), or none.
pub(super) fn color_of(value: Option<&serde_json::Value>) -> Result<Option<String>> {
    match value {
        None | Some(serde_json::Value::Null) => Ok(None),
        Some(serde_json::Value::String(s)) if s.is_empty() => Ok(None),
        Some(serde_json::Value::String(s)) => {
            let hex = s.trim().trim_start_matches('#');
            if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
                Ok(Some(format!("#{}", hex.to_uppercase())))
            } else {
                Err(format!("色 `{s}` は #RRGGBB で指定してください"))
            }
        }
        Some(other) => Err(format!("色 `{other}` は #RRGGBB で指定してください")),
    }
}
