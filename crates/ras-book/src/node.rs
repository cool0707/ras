//! The addon's functions: a `Book` class over [`crate::book::Book`].
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

use ras_workbook::{CellRef, RangeRef, Value};
use napi::{Error, Result};
use napi_derive::napi;
use serde_json::{Value as Json, json};

use ras_workbook::book::{self, Dir, MAX_COLS, MAX_ROWS, Read, Val};

fn err(msg: impl Into<String>) -> Error {
    Error::from_reason(msg.into())
}

fn sheet(b: &book::Book, i: u32) -> Result<usize> {
    let i = i as usize;
    if i < b.sheet_count() { Ok(i) } else { Err(err("インデックスが有効範囲にありません")) }
}

fn cell(row: f64, col: f64) -> Result<CellRef> {
    if !(1.0..=MAX_ROWS as f64).contains(&row) || !(1.0..=MAX_COLS as f64).contains(&col) || row.fract() != 0.0 || col.fract() != 0.0 {
        return Err(err(format!("セル ({row}, {col}) はシートの外です")));
    }
    Ok(CellRef::new(row as u32 - 1, col as u32 - 1))
}

fn range(r1: f64, c1: f64, r2: f64, c2: f64) -> Result<RangeRef> {
    Ok(RangeRef::new(cell(r1, c1)?, cell(r2, c2)?))
}

fn bounds(r: RangeRef) -> Json {
    json!([r.start.row + 1, r.start.col + 1, r.end.row + 1, r.end.col + 1])
}

fn to_json(v: &Val) -> Json {
    match v {
        Val::Empty => Json::Null,
        Val::Num(n) => json!(n),
        Val::Text(s) => json!(s),
        Val::Bool(b) => json!(b),
        Val::Date(n) => json!({ "$t": "date", "v": n }),
        Val::Err(n) => json!({ "$t": "error", "v": n }),
    }
}

fn from_json(v: &Json) -> Result<Val> {
    Ok(match v {
        Json::Null => Val::Empty,
        Json::Bool(b) => Val::Bool(*b),
        Json::Number(n) => Val::Num(n.as_f64().unwrap_or(0.0)),
        Json::String(s) => Val::Text(s.clone()),
        Json::Object(o) => {
            let n = o.get("v").and_then(Json::as_f64).unwrap_or(0.0);
            match o.get("$t").and_then(Json::as_str) {
                Some("date") => Val::Date(n),
                Some("error") => Val::Err(n as i32),
                Some("nothing") | None => Val::Empty,
                Some(_) => return Err(err("セルに入れられない値です")),
            }
        }
        Json::Array(_) => return Err(err("セルに配列は入れられません")),
    })
}

fn rows_from_json(rows: &Json) -> Result<Vec<Vec<Val>>> {
    match rows {
        Json::Array(rs) if rs.iter().all(Json::is_array) && !rs.is_empty() => {
            rs.iter().map(|r| r.as_array().map(|r| r.iter().map(from_json).collect::<Result<Vec<_>>>()).unwrap_or_else(|| Ok(Vec::new()))).collect()
        }
        // A one-dimensional array is a row, as in VBA.
        Json::Array(r) if !r.is_empty() => Ok(vec![r.iter().map(from_json).collect::<Result<Vec<_>>>()?]),
        Json::Array(_) => Err(err("空の配列はセルに入れられません")),
        v => Ok(vec![vec![from_json(v)?]]),
    }
}

fn value_json(v: &Value) -> Json {
    match v {
        Value::Array(a) => Json::Array(
            (0..a.rows).map(|r| Json::Array((0..a.cols).map(|c| to_json(&book::val_of(a.get(r, c).unwrap_or(&Value::Empty), false))).collect())).collect(),
        ),
        _ => to_json(&book::val_of(v, false)),
    }
}

fn what(s: &str) -> Result<Read> {
    Ok(match s {
        "Value" => Read::Value,
        "Value2" => Read::Value2,
        "Formula" => Read::Formula,
        "Formula2" => Read::Formula2,
        "FormulaR1C1" => Read::FormulaR1C1,
        "Text" => Read::Text,
        "NumberFormat" => Read::NumberFormat,
        _ => return Err(err(format!("{s} は読めません"))),
    })
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
        Ok(JsBook { inner: book::Book::open(Path::new(&path)).map_err(err)? })
    }

    /// A new workbook with one sheet.
    #[napi(factory)]
    pub fn create() -> JsBook {
        JsBook { inner: book::Book::new_book() }
    }

    #[napi]
    pub fn sheet_count(&self) -> u32 {
        self.inner.sheet_count() as u32
    }

    /// `{ name, codeName, visible }` of a sheet (0-based).
    #[napi]
    pub fn sheet_info(&self, i: u32) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        Ok(json!({ "name": self.inner.sheet_name(i), "codeName": self.inner.code_name(i), "visible": self.inner.visible(i) }))
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
        self.inner.add_sheet(at as usize, name.as_deref()).map(|i| i as u32).map_err(err)
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
    pub fn get(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, what_: String) -> Result<Json> {
        let i = sheet(&self.inner, i)?;
        let rows = self.inner.get(i, range(r1, c1, r2, c2)?, what(&what_)?).map_err(err)?;
        Ok(Json::Array(rows.iter().map(|r| Json::Array(r.iter().map(to_json).collect())).collect()))
    }

    /// `Range.Value = rows` (a value, or an array laid over the range).
    #[napi]
    pub fn set_values(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, rows: Json) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_values(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?).map_err(err)
    }

    /// `Range.Formula = rows` (`dynamic`: `Range.Formula2`).
    #[napi]
    pub fn set_formulas(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, rows: Json, dynamic: bool) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_formulas(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?, dynamic).map_err(err)
    }

    /// `Range.FormulaR1C1 = rows`.
    #[napi(js_name = "setFormulasR1C1")]
    pub fn set_formulas_r1c1(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, rows: Json) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_formulas_r1c1(i, range(r1, c1, r2, c2)?, &rows_from_json(&rows)?).map_err(err)
    }

    /// `Range.Copy Destination`: from a range of sheet `i` to the cells from (`row`, `col`) of sheet `to`.
    #[napi]
    pub fn copy(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, to: u32, row: f64, col: f64) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        let to = sheet(&self.inner, to)?;
        self.inner.copy(i, range(r1, c1, r2, c2)?, to, cell(row, col)?).map_err(err)
    }

    /// Whether the engine has a worksheet function.
    #[napi]
    pub fn is_function(name: String) -> bool {
        book::Book::is_function(&name)
    }

    #[napi]
    pub fn clear(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, formats: bool) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.clear(i, range(r1, c1, r2, c2)?, formats).map_err(err)
    }

    #[napi]
    pub fn set_number_format(&mut self, i: u32, r1: f64, c1: f64, r2: f64, c2: f64, code: String) -> Result<()> {
        let i = sheet(&self.inner, i)?;
        self.inner.set_number_format(i, range(r1, c1, r2, c2)?, &code).map_err(err)
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
        let dir = match direction {
            -4162 => Dir::Up,
            -4121 => Dir::Down,
            -4159 => Dir::Left,
            -4161 => Dir::Right,
            _ => return Err(err("End の方向が正しくありません")),
        };
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
        Ok(match self.inner.resolve(i, &text) {
            Some((s, r)) => json!([s, r.start.row + 1, r.start.col + 1, r.end.row + 1, r.end.col + 1]),
            None => Json::Null,
        })
    }

    /// `[{ name, sheet, refersTo }]`: `sheet` is the index of the sheet a name belongs to, or -1.
    #[napi]
    pub fn names(&self) -> Json {
        Json::Array(
            self.inner
                .names()
                .into_iter()
                .map(|(name, scope, formula)| json!({ "name": name, "sheet": scope.map(|s| s as i64).unwrap_or(-1), "refersTo": format!("={formula}") }))
                .collect(),
        )
    }

    #[napi]
    pub fn add_name(&mut self, name: String, scope: i32, refers_to: String) -> Result<()> {
        let scope = if scope < 0 { None } else { Some(sheet(&self.inner, scope as u32)?) };
        self.inner.add_name(&name, scope, &refers_to).map_err(err)
    }

    #[napi]
    pub fn delete_name(&mut self, name: String, scope: i32) {
        self.inner.delete_name(&name, if scope < 0 { None } else { Some(scope as usize) });
    }

    #[napi]
    pub fn calculate(&mut self) {
        self.inner.calculate();
    }

    /// Saves to `path`; returns the path written (`.xlsx` for a workbook without macros).
    #[napi]
    pub fn save(&mut self, path: String) -> Result<String> {
        self.inner.save(Path::new(&path)).map(|p| p.to_string_lossy().into_owned()).map_err(err)
    }
}
