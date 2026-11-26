//! The addon's functions: a `Book` class over [`crate::book::Book`].
//!
//! How values, cells and ranges cross the boundary is `ras_workbook::js`'s, which the browser's
//! `Book` (`crates/ras-wasm`) uses too: keep the two to the same methods (a test of
//! `ras_workbook::js` checks it).
//!
//! Rows and columns are 1-based here, as in Excel. Values cross the boundary as the COM addon
//! passes them (`ras-com/src/com.rs`), so that the run time converts both the same way:
//!
//! | JavaScript               | cell                       |
//! |--------------------------|----------------------------|
//! | `null`                   | empty                      |
//! | number / string / boolean| number / text / boolean    |
//! | `{ $t: "date", v }`      | a number shown as a date   |
//! | `{ $t: "error", v }`     | an error (`CVErr` number)  |
//!
//! Errors are thrown with the message as the user should read it; the run time raises them as
//! run-time error 1004.

use std::path::Path;

use napi::{Error, Result};
use napi_derive::napi;
use ras_workbook::js::{self, areas, bounds, to_json, value_json};
use ras_workbook::{CellRef, RangeRef};
use serde_json::{Value as Json, json};

use ras_workbook::book::{self, FilterCriterion, Read, Val};

fn err(msg: impl Into<String>) -> Error {
    Error::from_reason(msg.into())
}

// The conversions are the ones the browser's `Book` uses too (`ras_workbook::js`).
fn sheet(b: &book::Book, i: u32) -> Result<usize> {
    js::sheet(b, i).map_err(err)
}

fn cell(row: f64, col: f64) -> Result<CellRef> {
    js::cell(row, col).map_err(err)
}

fn range(r1: f64, c1: f64, r2: f64, c2: f64) -> Result<RangeRef> {
    js::range(r1, c1, r2, c2).map_err(err)
}

fn from_json(v: &Json) -> Result<Val> {
    js::from_json(v).map_err(err)
}

fn rows_from_json(rows: &Json) -> Result<Vec<Vec<Val>>> {
    js::rows_from_json(rows).map_err(err)
}

fn range_json(v: &Json, sheet: bool) -> Result<Option<(usize, RangeRef)>> {
    js::range_json(v, sheet).map_err(err)
}

fn criterion(v: &Json) -> Result<Option<FilterCriterion>> {
    js::criterion(v).map_err(err)
}

fn what(s: &str) -> Result<Read> {
    js::what(s).map_err(err)
}

/// A workbook the program works on instead of one open in Excel.
#[napi(js_name = "Book")]
pub struct JsBook {
    inner: book::Book,
}

// A range crosses the boundary as its four corners, so some methods take many arguments.
#[allow(clippy::too_many_arguments)]
#[napi]
impl JsBook {
    /// Opens an `.xlsx` / `.xlsm` file.
    #[napi(factory)]
    pub fn open(path: String) -> Result<JsBook> {
        Ok(JsBook {
            inner: book::Book::open(Path::new(&path)).map_err(err)?,
        })
    }

    /// Opens a workbook from the bytes of a file (`toBytes`): `ras test` puts a workbook back
    /// to what it was before a test this way (spec §14.12.5).
    #[napi(factory)]
    pub fn from_bytes(data: napi::bindgen_prelude::Buffer) -> Result<JsBook> {
        Ok(JsBook {
            inner: book::Book::open_bytes(&data).map_err(err)?,
        })
    }

    /// The controls of the sheets `sheets` of the workbook file `path` (ActiveX and form
    /// controls: what they are and hold), for the stand-in's objects of the controls.
    #[napi]
    pub fn controls(path: String, sheets: Vec<String>) -> Result<Json> {
        let data = std::fs::read(&path).map_err(|e| err(format!("{path}: {e}")))?;
        Ok(ras_workbook::view::controls_state(&data, &sheets))
    }

    /// The bytes of the file `save` would write.
    #[allow(clippy::wrong_self_convention)]
    #[napi]
    pub fn to_bytes(&mut self) -> Result<napi::bindgen_prelude::Buffer> {
        self.inner.to_bytes().map(Into::into).map_err(err)
    }

    /// A new workbook with one sheet.
    #[napi(factory)]
    pub fn create() -> JsBook {
        JsBook {
            inner: book::Book::new_book(),
        }
    }

    #[napi]
    pub fn sheet_count(&self) -> u32 {
        self.inner.sheet_count() as u32
    }

    /// `{ name, codeName, visible }` of a sheet (0-based).
    #[napi]
    pub fn sheet_info(&self, i: u32) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(
            json!({ "name": self.inner.sheet_name(i), "codeName": self.inner.code_name(i), "visible": self.inner.visible(i) }),
        )
    }

    #[napi]
    pub fn sheet_index(&self, name: String) -> Option<u32> {
        self.inner.sheet_index(&name).map(|i| i as u32)
    }

    #[napi]
    pub fn set_visible(&mut self, i: u32, v: i32) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_visible(i, v).map_err(err)
    }

    #[napi]
    pub fn rename_sheet(&mut self, i: u32, name: String) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.rename_sheet(i, &name).map_err(err)
    }

    /// Adds a sheet before position `at` (0-based); returns its index.
    #[napi]
    pub fn add_sheet(&mut self, at: u32, name: Option<String>) -> Result<u32> {
        self.inner
            .add_sheet(at as usize, name.as_deref())
            .map(|i| i as u32)
            .map_err(err)
    }

    #[napi]
    pub fn delete_sheet(&mut self, i: u32) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.delete_sheet(i).map_err(err)
    }

    #[napi]
    pub fn move_sheet(&mut self, from: u32, to: u32) -> Result<()> {
        let from = sheet(&self.inner, from)?;
        self.inner.move_sheet(from, to as usize).map_err(err)
    }

    /// The cells of a range, row by row: `what` is `Value`, `Value2`, `Formula`, `Formula2`,
    /// `Text` or `NumberFormat`.
    #[napi]
    pub fn get(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        what_: String,
    ) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        let rows = self
            .inner
            .get(i, range(r1, c1, r2, c2)?, what(&what_)?)
            .map_err(err)?;
        Ok(js::rows_json(&rows))
    }

    /// `Range.Value = rows` (a value, or an array laid over the range).
    #[napi]
    pub fn set_values(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        rows: Json,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_values(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?)
            .map_err(err)
    }

    /// `Range.Formula = rows` (`dynamic`: `Range.Formula2`).
    #[napi]
    pub fn set_formulas(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        rows: Json,
        dynamic: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_formulas(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?, dynamic)
            .map_err(err)
    }

    /// `Range.FormulaR1C1 = rows`.
    #[napi(js_name = "setFormulasR1C1")]
    pub fn set_formulas_r1c1(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        rows: Json,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_formulas_r1c1(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?)
            .map_err(err)
    }

    /// `Range.Copy Destination`: from a range of sheet `i` to the cells from (`row`, `col`) of sheet `to`.
    #[napi]
    pub fn copy(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        to: u32,
        row: f64,
        col: f64,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let to = sheet(&self.inner, to)?;
        self.inner
            .copy(i, range(r1, c1, r2, c2)?, to, cell(row, col)?)
            .map_err(err)
    }

    /// Pastes a range of sheet `i` at (`row`, `col`) of sheet `to`, as `Range.PasteSpecial` does:
    /// `what` is `XlPasteType`, `op` `XlPasteSpecialOperation`.
    #[napi]
    pub fn paste(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        to: u32,
        row: f64,
        col: f64,
        what_: i32,
        op: i32,
        skip_blanks: bool,
        transpose: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let to = sheet(&self.inner, to)?;
        let p = js::paste_of(what_, op, skip_blanks, transpose).map_err(err)?;
        self.inner
            .paste(i, range(r1, c1, r2, c2)?, to, cell(row, col)?, &p)
            .map_err(err)
    }

    /// `Range.Cut Destination`: moves a range of sheet `i` to (`row`, `col`) of sheet `to`.
    #[napi]
    pub fn move_cells(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        to: u32,
        row: f64,
        col: f64,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let to = sheet(&self.inner, to)?;
        self.inner
            .move_cells(i, range(r1, c1, r2, c2)?, to, cell(row, col)?)
            .map_err(err)
    }

    /// `Range.Insert`: `right` shifts the cells right (down otherwise); `from_after` formats the
    /// new cells as the ones below or to the right.
    #[napi]
    pub fn insert_cells(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        right: bool,
        from_after: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .insert_cells(i, range(r1, c1, r2, c2)?, right, from_after)
            .map_err(err)
    }

    /// `Range.Delete`: `left` shifts the cells left (up otherwise).
    #[napi]
    pub fn delete_cells(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        left: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .delete_cells(i, range(r1, c1, r2, c2)?, left)
            .map_err(err)
    }

    /// `Range.Sort` of a range without its headings: `keys` is `[[column, descending], ...]`
    /// (the row, when `across` sorts columns).
    #[napi]
    pub fn sort(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        keys: Json,
        across: bool,
        match_case: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let keys = js::sort_keys(&keys);
        self.inner
            .sort(i, range(r1, c1, r2, c2)?, &keys, across, match_case)
            .map_err(err)
    }

    /// Whether the first row of a range looks like headings (`Header:=xlGuess`).
    #[napi]
    pub fn guess_header(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64) -> Result<bool> {
        let i = sheet(&self.inner, i)?;
        Ok(self.inner.guess_header(i, range(r1, c1, r2, c2)?))
    }

    /// `Range.RemoveDuplicates` of a range without its headings, by the columns `columns` (of
    /// the sheet). Returns how many rows went.
    #[napi]
    pub fn remove_duplicates(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        columns: Vec<u32>,
    ) -> Result<u32> {
        let i = sheet(&self.inner, i)?;
        let columns: Vec<u32> = columns
            .into_iter()
            .filter_map(|c| c.checked_sub(1))
            .collect();
        self.inner
            .remove_duplicates(i, range(r1, c1, r2, c2)?, &columns)
            .map_err(err)
    }

    /// The filter of a sheet: `{ range: [top, left, bottom, right], fields: [...] }` (the fields
    /// it filters by, counted from 1), or null.
    #[napi]
    pub fn auto_filter(&self, i: u32) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(js::auto_filter_json(self.inner.auto_filter(i)))
    }

    /// Puts the filter of a sheet on a range (`[top, left, bottom, right]`), or removes it (null).
    #[napi]
    pub fn set_auto_filter(&mut self, i: u32, r: Json) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let r = range_json(&r, false)?.map(|(_, r)| r);
        self.inner.set_auto_filter(i, r).map_err(err)
    }

    /// Filters by field `field` (from 1) of the filter, or no longer (null).
    #[napi]
    pub fn set_filter(&mut self, i: u32, field: u32, c: Json) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let field = field
            .checked_sub(1)
            .ok_or_else(|| err("AutoFilter の Field が正しくありません。"))?;
        self.inner.set_filter(i, field, criterion(&c)?).map_err(err)
    }

    #[napi]
    pub fn filter_mode(&self, i: u32) -> Result<bool> {
        let i = sheet(&self.inner, i)?;
        Ok(self.inner.filter_mode(i))
    }

    #[napi]
    pub fn show_all(&mut self, i: u32) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.show_all(i).map_err(err)
    }

    /// `Range.AdvancedFilter`: `criteria` and `copy_to` are `[sheet, top, left, bottom, right]` or null.
    #[napi]
    pub fn advanced_filter(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        criteria: Json,
        copy_to: Json,
        unique: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .advanced_filter(
                i,
                range(r1, c1, r2, c2)?,
                range_json(&criteria, true)?,
                range_json(&copy_to, true)?,
                unique,
            )
            .map_err(err)
    }

    /// Hides or shows the rows (`columns`: the columns) `lo` to `hi`.
    #[napi]
    pub fn set_hidden(
        &mut self,
        i: u32,
        columns: bool,
        lo: u32,
        hi: u32,
        hidden: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_hidden(
                i,
                columns,
                lo.saturating_sub(1),
                hi.saturating_sub(1),
                hidden,
            )
            .map_err(err)
    }

    /// Whether the rows (columns) `lo` to `hi` are hidden: null when only some are.
    #[napi]
    pub fn hidden(&self, i: u32, columns: bool, lo: u32, hi: u32) -> Result<Option<bool>> {
        let i = sheet(&self.inner, i)?;
        Ok(self
            .inner
            .hidden(i, columns, lo.saturating_sub(1), hi.saturating_sub(1)))
    }

    /// The cells of a range that are not hidden, as `[[top, left, bottom, right], ...]`.
    #[napi]
    pub fn visible_areas(&self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(areas(&self.inner.visible_areas(i, range(r1, c1, r2, c2)?)))
    }

    /// `Range.SpecialCells(type, value)`: the areas found (none: an empty array).
    #[napi]
    pub fn special_cells(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        type_: i32,
        value: i32,
    ) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        let what = js::special(type_, value).map_err(err)?;
        Ok(areas(
            &self
                .inner
                .special_cells(i, range(r1, c1, r2, c2)?, what)
                .map_err(err)?,
        ))
    }

    /// A format of the cells of a range (`font.bold`, `interior.color`, `borders.9.linestyle`,
    /// `columnwidth`, ...): null when the cells differ.
    #[napi]
    pub fn format(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        prop: String,
    ) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(self
            .inner
            .format_of(i, range(r1, c1, r2, c2)?, &prop)
            .map_err(err)?
            .as_ref()
            .map(to_json)
            .unwrap_or(Json::Null))
    }

    #[napi]
    pub fn set_format(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        prop: String,
        value: Json,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_format(i, range(r1, c1, r2, c2)?, &prop, &from_json(&value)?)
            .map_err(err)
    }

    #[napi]
    pub fn clear_formats(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .clear_formats(i, range(r1, c1, r2, c2)?)
            .map_err(err)
    }

    /// `Range.Merge` (`across`: each row on its own).
    #[napi]
    pub fn merge(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        across: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .merge_cells(i, range(r1, c1, r2, c2)?, across)
            .map_err(err)
    }

    /// `Range.MergeArea` of a cell.
    #[napi]
    pub fn merge_area(&self, i: u32, row: f64, col: f64) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(bounds(self.inner.merge_area(i, cell(row, col)?)))
    }

    /// The colour of a sheet's tab (as VBA numbers colours), or null.
    #[napi]
    pub fn tab_color(&self, i: u32) -> Result<Option<f64>> {
        let i = sheet(&self.inner, i)?;
        Ok(self.inner.tab_color(i))
    }

    #[napi]
    pub fn set_tab_color(&mut self, i: u32, color: Option<f64>) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_tab_color(i, color);
        Ok(())
    }

    /// `Worksheet.Copy`: a copy of sheet `i` before position `at`; returns its index.
    #[napi]
    pub fn copy_sheet(&mut self, i: u32, at: u32) -> Result<u32> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .copy_sheet(i, at as usize)
            .map(|n| n as u32)
            .map_err(err)
    }

    /// `Application.Calculation`: manual stops formulas from following what is written.
    #[napi]
    pub fn set_manual(&mut self, manual: bool) {
        self.inner.set_manual(manual);
    }

    /// Whether the workbook was saved with manual calculation.
    #[napi]
    pub fn saved_manual(&self) -> bool {
        self.inner.saved_manual()
    }

    /// `Application.Calculate`: the formulas that wait to be calculated.
    #[napi]
    pub fn recalculate(&mut self) {
        self.inner.recalculate();
    }

    /// Calculates what waits (unless calculation is manual) and returns the sheets whose
    /// formulas were calculated, for the `Calculate` events.
    #[napi]
    pub fn calculated_sheets(&mut self) -> Vec<u32> {
        self.inner
            .calculated_sheets()
            .into_iter()
            .map(|i| i as u32)
            .collect()
    }

    /// Whether the engine has a worksheet function.
    #[napi]
    pub fn is_function(name: String) -> bool {
        book::Book::is_function(&name)
    }

    #[napi]
    pub fn clear(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        formats: bool,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .clear(i, range(r1, c1, r2, c2)?, formats)
            .map_err(err)
    }

    #[napi]
    pub fn set_number_format(
        &mut self,
        i: u32,
        r1: f64,
        c1: f64,
        r2: f64,
        c2: f64,
        code: String,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_number_format(i, range(r1, c1, r2, c2)?, &code)
            .map_err(err)
    }

    /// `[top, left, bottom, right]` of the used range, or null when the sheet is empty.
    #[napi]
    pub fn used_range(&mut self, i: u32) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(self.inner.used_range(i).map(bounds).unwrap_or(Json::Null))
    }

    /// `Range.End(direction)` from a cell: `[row, column]`. The direction is `XlDirection`.
    #[napi]
    pub fn end(&mut self, i: u32, row: f64, col: f64, direction: i32) -> Result<Vec<u32>> {
        let i = sheet(&self.inner, i)?;
        let dir = js::direction(direction).map_err(err)?;
        let at = self.inner.end(i, cell(row, col)?, dir);
        Ok(vec![at.row + 1, at.col + 1])
    }

    #[napi]
    pub fn current_region(&mut self, i: u32, row: f64, col: f64) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(bounds(self.inner.current_region(i, cell(row, col)?)))
    }

    /// A formula worked out on a sheet: a value, or rows of values.
    #[napi]
    pub fn evaluate(&mut self, i: u32, formula: String) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(value_json(&self.inner.evaluate(i, &formula)))
    }

    /// The range a reference or name refers to: `[sheet, top, left, bottom, right]`, or null.
    #[napi]
    pub fn resolve(&mut self, i: u32, text: String) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(js::resolved(self.inner.resolve(i, &text)))
    }

    /// The hyperlinks of the cells of a range: `[{ row, col, address, subAddress, tip }]`, by rows.
    #[napi]
    pub fn hyperlinks(&self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        let list = self
            .inner
            .hyperlinks(i, range(r1, c1, r2, c2)?)
            .map_err(err)?;
        Ok(js::hyperlinks_json(list))
    }

    #[napi]
    pub fn set_hyperlink(
        &mut self,
        i: u32,
        row: f64,
        col: f64,
        address: String,
        sub_address: String,
        tip: String,
    ) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .set_hyperlink(i, cell(row, col)?, &address, &sub_address, &tip)
            .map_err(err)
    }

    /// Takes the hyperlinks of the cells of a range away; returns how many there were.
    #[napi]
    pub fn remove_hyperlinks(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64) -> Result<u32> {
        let i = sheet(&self.inner, i)?;
        self.inner
            .remove_hyperlinks(i, range(r1, c1, r2, c2)?)
            .map_err(err)
    }

    /// `[{ name, sheet, refersTo }]`: `sheet` is the index of the sheet a name belongs to, or -1.
    #[napi]
    pub fn names(&self) -> Json {
        js::names_json(&self.inner)
    }

    #[napi]
    pub fn add_name(&mut self, name: String, scope: i32, refers_to: String) -> Result<()> {
        let scope = if scope < 0 {
            None
        } else {
            Some(sheet(&self.inner, scope as u32)?)
        };
        self.inner.add_name(&name, scope, &refers_to).map_err(err)
    }

    #[napi]
    pub fn delete_name(&mut self, name: String, scope: i32) {
        self.inner.delete_name(
            &name,
            if scope < 0 {
                None
            } else {
                Some(scope as usize)
            },
        );
    }

    #[napi]
    pub fn calculate(&mut self) {
        self.inner.calculate();
    }

    /// Saves to `path`; returns the path written (`.xlsx` for a workbook without macros).
    #[napi]
    pub fn save(&mut self, path: String) -> Result<String> {
        self.inner
            .save(Path::new(&path))
            .map(|p| p.to_string_lossy().into_owned())
            .map_err(err)
    }
}
