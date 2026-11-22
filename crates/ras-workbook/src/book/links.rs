//! The hyperlinks of the cells: `Range.Hyperlinks`, `Worksheet.Hyperlinks` and `Hyperlink` of the
//! stand-in. They are kept where the file keeps them (umya-spreadsheet's cells), so they are saved
//! with the workbook and move with the cells inserted and deleted.
//!
//! A link of the file is either an address (a web page, a file: `r:id`), or a place in the
//! workbook (`location`). umya-spreadsheet holds one text for the two, so an address with a place
//! in it is kept as `address#place`, as a browser reads it.

use umya_spreadsheet::{Hyperlink, Worksheet};

use super::*;

/// A hyperlink of a cell, as `Hyperlink` shows it.
#[derive(Clone, Debug, PartialEq)]
pub struct LinkInfo {
    pub at: CellRef,
    /// Where it goes: a URL, a file (`Address`); empty for a place in the workbook.
    pub address: String,
    /// The place in the workbook or in the address (`SubAddress`: `Sheet2!A1`, a name).
    pub sub_address: String,
    /// What shows while the mouse is on it (`ScreenTip`).
    pub tip: String,
}

fn info(at: CellRef, h: &Hyperlink) -> LinkInfo {
    let (address, sub_address) = if h.location() {
        (String::new(), h.url().to_string())
    } else {
        match h.url().split_once('#') {
            Some((a, s)) => (a.to_string(), s.to_string()),
            None => (h.url().to_string(), String::new()),
        }
    };
    LinkInfo {
        at,
        address,
        sub_address,
        tip: h.tooltip().to_string(),
    }
}

/// The cell at `at` again, without its hyperlink (umya-spreadsheet cannot take one away).
fn unlink(ws: &mut Worksheet, at: CellRef) {
    let Some(c) = ws.cell(coord(at)) else {
        return;
    };
    if c.hyperlink().is_none() {
        return;
    }
    let (value, style) = (c.cell_value().clone(), c.style().clone());
    let (cell_meta, value_meta, phonetic) =
        (c.cell_meta_index(), c.value_meta_index(), c.show_phonetic());
    ws.remove_cell(coord(at));
    let c = ws.cell_mut(coord(at));
    c.set_cell_value(value);
    c.set_style(style);
    c.set_cell_meta_index(cell_meta);
    c.set_value_meta_index(value_meta);
    c.set_show_phonetic(phonetic);
}

impl Book {
    /// The hyperlinks of the cells of `within` on sheet `si`, by rows.
    pub fn hyperlinks(&self, si: usize, within: RangeRef) -> Result<Vec<LinkInfo>> {
        self.check_sheet(si)?;
        let Ok(ws) = self.file.sheet(si) else {
            return Ok(Vec::new());
        };
        let mut out: Vec<LinkInfo> = ws
            .cells()
            .iter()
            .filter_map(|c| {
                let h = c.hyperlink()?;
                let at = at_of(c);
                within.contains(at).then(|| info(at, h))
            })
            .collect();
        out.sort_by_key(|l| (l.at.row, l.at.col));
        Ok(out)
    }

    /// Gives the cell `at` of sheet `si` a hyperlink (in place of the one it had): to `address`
    /// (with `sub_address` in it), or to the place `sub_address` of the workbook.
    pub fn set_hyperlink(
        &mut self,
        si: usize,
        at: CellRef,
        address: &str,
        sub_address: &str,
        tip: &str,
    ) -> Result<()> {
        self.check_sheet(si)?;
        if address.is_empty() && sub_address.is_empty() {
            return Err("ハイパーリンクのアドレスがありません".into());
        }
        let ws = self
            .file
            .sheet_mut(si)
            .map_err(|_| "シートがありません".to_string())?;
        let h = ws.cell_mut(coord(at)).hyperlink_mut();
        if address.is_empty() {
            h.set_url(sub_address);
            h.set_location(true);
        } else {
            h.set_url(if sub_address.is_empty() {
                address.to_string()
            } else {
                format!("{address}#{sub_address}")
            });
            h.set_location(false);
        }
        h.set_tooltip(tip);
        Ok(())
    }

    /// Takes the hyperlinks of the cells of `within` away; returns how many there were.
    pub fn remove_hyperlinks(&mut self, si: usize, within: RangeRef) -> Result<u32> {
        let cells: Vec<CellRef> = self
            .hyperlinks(si, within)?
            .into_iter()
            .map(|l| l.at)
            .collect();
        if let Ok(ws) = self.file.sheet_mut(si) {
            for at in &cells {
                unlink(ws, *at);
            }
        }
        Ok(cells.len() as u32)
    }
}
