//! The boundary between the workbook and the JavaScript run time (`ras-excel.js`): how values,
//! cells and ranges cross it, and the methods of the `Book` the run time calls.
//!
//! Two bindings give the run time its `Book`: the Node addon (`crates/ras-book`, one N-API
//! method each) and the browser's (`crates/ras-wasm`, which passes every call to [`call`]).
//! Both convert with the functions here, so that the run time sees the same `Book` in either.
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
//! Errors are the message as the user should read it; the run time raises them as run-time
//! error 1004.

use serde_json::{Value as Json, json};

use crate::book::{
    self, Book, Dir, FilterCriterion, MAX_COLS, MAX_ROWS, Paste, PasteOp, PasteWhat, Read, Result,
    SortKey, Special, Val,
};
use crate::{CellRef, RangeRef, Value};

pub fn sheet(b: &Book, i: u32) -> Result<usize> {
    let i = i as usize;
    if i < b.sheet_count() {
        Ok(i)
    } else {
        Err("インデックスが有効範囲にありません".into())
    }
}

pub fn cell(row: f64, col: f64) -> Result<CellRef> {
    if !(1.0..=MAX_ROWS as f64).contains(&row)
        || !(1.0..=MAX_COLS as f64).contains(&col)
        || row.fract() != 0.0
        || col.fract() != 0.0
    {
        return Err(format!("セル ({row}, {col}) はシートの外です"));
    }
    Ok(CellRef::new(row as u32 - 1, col as u32 - 1))
}

pub fn range(r1: f64, c1: f64, r2: f64, c2: f64) -> Result<RangeRef> {
    Ok(RangeRef::new(cell(r1, c1)?, cell(r2, c2)?))
}

pub fn bounds(r: RangeRef) -> Json {
    json!([
        r.start.row + 1,
        r.start.col + 1,
        r.end.row + 1,
        r.end.col + 1
    ])
}

pub fn to_json(v: &Val) -> Json {
    match v {
        Val::Empty => Json::Null,
        Val::Num(n) => json!(n),
        Val::Text(s) => json!(s),
        Val::Bool(b) => json!(b),
        Val::Date(n) => json!({ "$t": "date", "v": n }),
        Val::Err(n) => json!({ "$t": "error", "v": n }),
    }
}

pub fn from_json(v: &Json) -> Result<Val> {
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
                Some(_) => return Err("セルに入れられない値です".into()),
            }
        }
        Json::Array(_) => return Err("セルに配列は入れられません".into()),
    })
}

pub fn rows_from_json(rows: &Json) -> Result<Vec<Vec<Val>>> {
    match rows {
        Json::Array(rs) if rs.iter().all(Json::is_array) && !rs.is_empty() => rs
            .iter()
            .map(|r| {
                r.as_array()
                    .map(|r| r.iter().map(from_json).collect::<Result<Vec<_>>>())
                    .unwrap_or_else(|| Ok(Vec::new()))
            })
            .collect(),
        // A one-dimensional array is a row, as in VBA.
        Json::Array(r) if !r.is_empty() => {
            Ok(vec![r.iter().map(from_json).collect::<Result<Vec<_>>>()?])
        }
        Json::Array(_) => Err("空の配列はセルに入れられません".into()),
        v => Ok(vec![vec![from_json(v)?]]),
    }
}

pub fn rows_json(rows: &[Vec<Val>]) -> Json {
    Json::Array(
        rows.iter()
            .map(|r| Json::Array(r.iter().map(to_json).collect()))
            .collect(),
    )
}

pub fn value_json(v: &Value) -> Json {
    match v {
        Value::Array(a) => Json::Array(
            (0..a.rows)
                .map(|r| {
                    Json::Array(
                        (0..a.cols)
                            .map(|c| {
                                to_json(&book::val_of(a.get(r, c).unwrap_or(&Value::Empty), false))
                            })
                            .collect(),
                    )
                })
                .collect(),
        ),
        _ => to_json(&book::val_of(v, false)),
    }
}

pub fn areas(list: &[RangeRef]) -> Json {
    Json::Array(list.iter().map(|r| bounds(*r)).collect())
}

/// `[top, left, bottom, right]` (and a sheet before them, with `sheet`), or null.
pub fn range_json(v: &Json, sheet: bool) -> Result<Option<(usize, RangeRef)>> {
    let Json::Array(a) = v else { return Ok(None) };
    let n: Vec<f64> = a.iter().map(|x| x.as_f64().unwrap_or(0.0)).collect();
    let (si, n) = if sheet {
        (
            n.first().copied().unwrap_or(0.0) as usize,
            &n[1.min(n.len())..],
        )
    } else {
        (0, &n[..])
    };
    if n.len() != 4 {
        return Err("範囲の指定が正しくありません".into());
    }
    Ok(Some((si, range(n[0], n[1], n[2], n[3])?)))
}

/// What a column of a filter shows: `{ kind: "values", values, blanks }`,
/// `{ kind: "custom", a: [op, value], b: [op, value] | null, and }`,
/// `{ kind: "top", bottom, percent, count }`.
pub fn criterion(v: &Json) -> Result<Option<FilterCriterion>> {
    if v.is_null() {
        return Ok(None);
    }
    let text = |x: &Json| x.as_str().map(str::to_string).unwrap_or_default();
    let flag = |k: &str| v.get(k).and_then(Json::as_bool).unwrap_or(false);
    let test = |x: &Json| -> Option<(String, String)> {
        x.as_array()
            .filter(|a| a.len() == 2)
            .map(|a| (text(&a[0]), text(&a[1])))
    };
    let wrong = || "フィルターの条件が正しくありません".to_string();
    Ok(Some(match v.get("kind").and_then(Json::as_str) {
        Some("values") => FilterCriterion::Values {
            values: v
                .get("values")
                .and_then(Json::as_array)
                .map(|a| a.iter().map(text).collect())
                .unwrap_or_default(),
            blanks: flag("blanks"),
        },
        Some("custom") => FilterCriterion::Custom {
            a: v.get("a").and_then(test).ok_or_else(wrong)?,
            b: v.get("b").and_then(test),
            and: flag("and"),
        },
        Some("top") => FilterCriterion::Top10 {
            bottom: flag("bottom"),
            percent: flag("percent"),
            count: v.get("count").and_then(Json::as_u64).unwrap_or(10) as u32,
        },
        _ => return Err(wrong()),
    }))
}

pub fn what(s: &str) -> Result<Read> {
    Ok(match s {
        "Value" => Read::Value,
        "Value2" => Read::Value2,
        "Formula" => Read::Formula,
        "Formula2" => Read::Formula2,
        "FormulaR1C1" => Read::FormulaR1C1,
        "Text" => Read::Text,
        "NumberFormat" => Read::NumberFormat,
        _ => return Err(format!("{s} は読めません")),
    })
}

/// `Range.Sort` keys: `[[column, descending], ...]`, columns from 1.
pub fn sort_keys(keys: &Json) -> Vec<SortKey> {
    keys.as_array()
        .map(|a| {
            a.iter()
                .filter_map(|k| {
                    Some(SortKey {
                        index: (k.get(0)?.as_f64()? as u32).checked_sub(1)?,
                        descending: k.get(1)?.as_bool()?,
                    })
                })
                .collect()
        })
        .unwrap_or_default()
}

pub fn paste_of(what_: i32, op: i32, skip_blanks: bool, transpose: bool) -> Result<Paste> {
    let what = PasteWhat::of(what_)
        .ok_or("この形式の貼り付け（コメント・入力規則など）には対応していません")?;
    Ok(Paste {
        what,
        op: PasteOp::of(op),
        skip_blanks,
        transpose,
    })
}

/// `Range.SpecialCells(type, value)`.
pub fn special(type_: i32, value: i32) -> Result<Special> {
    let mask = if value <= 0 { 23 } else { value as u32 };
    Ok(match type_ {
        2 => Special::Constants(mask),
        -4123 => Special::Formulas(mask),
        4 => Special::Blanks,
        11 => Special::LastCell,
        12 => Special::Visible,
        _ => return Err("この種類の SpecialCells には対応していません".into()),
    })
}

/// `XlDirection`.
pub fn direction(direction: i32) -> Result<Dir> {
    Ok(match direction {
        -4162 => Dir::Up,
        -4121 => Dir::Down,
        -4159 => Dir::Left,
        -4161 => Dir::Right,
        _ => return Err("End の方向が正しくありません".into()),
    })
}

pub fn resolved(found: Option<(usize, RangeRef)>) -> Json {
    match found {
        Some((s, r)) => json!([
            s,
            r.start.row + 1,
            r.start.col + 1,
            r.end.row + 1,
            r.end.col + 1
        ]),
        None => Json::Null,
    }
}

pub fn auto_filter_json(found: Option<(RangeRef, Vec<u32>)>) -> Json {
    match found {
        Some((r, fields)) => {
            json!({ "range": bounds(r), "fields": fields.iter().map(|f| f + 1).collect::<Vec<_>>() })
        }
        None => Json::Null,
    }
}

pub fn hyperlinks_json(list: Vec<book::LinkInfo>) -> Json {
    Json::Array(
        list.into_iter()
            .map(|l| json!({ "row": l.at.row + 1, "col": l.at.col + 1, "address": l.address, "subAddress": l.sub_address, "tip": l.tip }))
            .collect(),
    )
}

pub fn names_json(b: &Book) -> Json {
    Json::Array(
        b.names()
            .into_iter()
            .map(|(name, scope, formula)| json!({ "name": name, "sheet": scope.map(|s| s as i64).unwrap_or(-1), "refersTo": format!("={formula}") }))
            .collect(),
    )
}

/// The methods of `Book` that [`call`] takes (the instance methods of the Node addon but
/// `toBytes` and `save`, whose bytes and files the binding handles itself).
pub const METHODS: &[&str] = &[
    "sheetCount",
    "sheetInfo",
    "sheetIndex",
    "setVisible",
    "renameSheet",
    "addSheet",
    "deleteSheet",
    "moveSheet",
    "get",
    "setValues",
    "setFormulas",
    "setFormulasR1C1",
    "copy",
    "paste",
    "moveCells",
    "insertCells",
    "deleteCells",
    "sort",
    "guessHeader",
    "removeDuplicates",
    "autoFilter",
    "setAutoFilter",
    "setFilter",
    "filterMode",
    "showAll",
    "advancedFilter",
    "setHidden",
    "hidden",
    "visibleAreas",
    "specialCells",
    "format",
    "setFormat",
    "clearFormats",
    "merge",
    "mergeArea",
    "tabColor",
    "setTabColor",
    "copySheet",
    "setManual",
    "savedManual",
    "recalculate",
    "calculatedSheets",
    "clear",
    "setNumberFormat",
    "usedRange",
    "end",
    "currentRegion",
    "evaluate",
    "resolve",
    "hyperlinks",
    "setHyperlink",
    "removeHyperlinks",
    "names",
    "addName",
    "deleteName",
    "calculate",
];

/// The arguments of a call, read as the Node addon reads them.
struct Args<'a>(&'a [Json]);

impl Args<'_> {
    fn at(&self, i: usize) -> &Json {
        self.0.get(i).unwrap_or(&Json::Null)
    }
    fn f(&self, i: usize) -> f64 {
        self.at(i).as_f64().unwrap_or(0.0)
    }
    fn u(&self, i: usize) -> u32 {
        self.f(i).max(0.0) as u32
    }
    fn i(&self, i: usize) -> i32 {
        self.f(i) as i32
    }
    fn b(&self, i: usize) -> bool {
        self.at(i).as_bool().unwrap_or(false)
    }
    fn s(&self, i: usize) -> String {
        match self.at(i) {
            Json::String(s) => s.clone(),
            Json::Null => String::new(),
            v => v.to_string(),
        }
    }
    /// The sheet at `i` and the range of the four numbers after it.
    fn sheet_range(&self, b: &Book, i: usize) -> Result<(usize, RangeRef)> {
        Ok((
            sheet(b, self.u(i))?,
            range(self.f(i + 1), self.f(i + 2), self.f(i + 3), self.f(i + 4))?,
        ))
    }
    fn sheet_cell(&self, b: &Book, i: usize) -> Result<(usize, CellRef)> {
        Ok((sheet(b, self.u(i))?, cell(self.f(i + 1), self.f(i + 2))?))
    }
}

/// Calls the method `method` of a `Book` with the arguments `args` (as JSON), as the Node
/// addon's method of that name does.
pub fn call(b: &mut Book, method: &str, args: &[Json]) -> Result<Json> {
    let a = Args(args);
    let none = Ok(Json::Null);
    match method {
        "sheetCount" => Ok(json!(b.sheet_count())),
        "sheetInfo" => {
            let i = sheet(b, a.u(0))?;
            Ok(
                json!({ "name": b.sheet_name(i), "codeName": b.code_name(i), "visible": b.visible(i) }),
            )
        }
        "sheetIndex" => Ok(json!(b.sheet_index(&a.s(0)))),
        "setVisible" => {
            let i = sheet(b, a.u(0))?;
            b.set_visible(i, a.i(1)).map(|()| Json::Null)
        }
        "renameSheet" => {
            let i = sheet(b, a.u(0))?;
            b.rename_sheet(i, &a.s(1)).map(|()| Json::Null)
        }
        "addSheet" => {
            let name = a.at(1).as_str().map(str::to_string);
            b.add_sheet(a.u(0) as usize, name.as_deref())
                .map(|i| json!(i))
        }
        "deleteSheet" => {
            let i = sheet(b, a.u(0))?;
            b.delete_sheet(i).map(|()| Json::Null)
        }
        "moveSheet" => {
            let i = sheet(b, a.u(0))?;
            b.move_sheet(i, a.u(1) as usize).map(|()| Json::Null)
        }
        "get" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.get(i, r, what(&a.s(5))?).map(|rows| rows_json(&rows))
        }
        "setValues" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.set_values(i, r, &rows_from_json(a.at(5))?)
                .map(|()| Json::Null)
        }
        "setFormulas" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.set_formulas(i, r, &rows_from_json(a.at(5))?, a.b(6))
                .map(|()| Json::Null)
        }
        "setFormulasR1C1" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.set_formulas_r1c1(i, r, &rows_from_json(a.at(5))?)
                .map(|()| Json::Null)
        }
        "copy" => {
            let (i, r) = a.sheet_range(b, 0)?;
            let (to, at) = a.sheet_cell(b, 5)?;
            b.copy(i, r, to, at).map(|()| Json::Null)
        }
        "paste" => {
            let (i, r) = a.sheet_range(b, 0)?;
            let (to, at) = a.sheet_cell(b, 5)?;
            let p = paste_of(a.i(8), a.i(9), a.b(10), a.b(11))?;
            b.paste(i, r, to, at, &p).map(|()| Json::Null)
        }
        "moveCells" => {
            let (i, r) = a.sheet_range(b, 0)?;
            let (to, at) = a.sheet_cell(b, 5)?;
            b.move_cells(i, r, to, at).map(|()| Json::Null)
        }
        "insertCells" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.insert_cells(i, r, a.b(5), a.b(6)).map(|()| Json::Null)
        }
        "deleteCells" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.delete_cells(i, r, a.b(5)).map(|()| Json::Null)
        }
        "sort" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.sort(i, r, &sort_keys(a.at(5)), a.b(6), a.b(7))
                .map(|()| Json::Null)
        }
        "guessHeader" => {
            let (i, r) = a.sheet_range(b, 0)?;
            Ok(json!(b.guess_header(i, r)))
        }
        "removeDuplicates" => {
            let (i, r) = a.sheet_range(b, 0)?;
            let columns: Vec<u32> = a
                .at(5)
                .as_array()
                .map(|c| {
                    c.iter()
                        .filter_map(|x| (x.as_f64()? as u32).checked_sub(1))
                        .collect()
                })
                .unwrap_or_default();
            b.remove_duplicates(i, r, &columns).map(|n| json!(n))
        }
        "autoFilter" => {
            let i = sheet(b, a.u(0))?;
            Ok(auto_filter_json(b.auto_filter(i)))
        }
        "setAutoFilter" => {
            let i = sheet(b, a.u(0))?;
            let r = range_json(a.at(1), false)?.map(|(_, r)| r);
            b.set_auto_filter(i, r).map(|()| Json::Null)
        }
        "setFilter" => {
            let i = sheet(b, a.u(0))?;
            let field = a
                .u(1)
                .checked_sub(1)
                .ok_or("AutoFilter の Field が正しくありません。")?;
            b.set_filter(i, field, criterion(a.at(2))?)
                .map(|()| Json::Null)
        }
        "filterMode" => {
            let i = sheet(b, a.u(0))?;
            Ok(json!(b.filter_mode(i)))
        }
        "showAll" => {
            let i = sheet(b, a.u(0))?;
            b.show_all(i).map(|()| Json::Null)
        }
        "advancedFilter" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.advanced_filter(
                i,
                r,
                range_json(a.at(5), true)?,
                range_json(a.at(6), true)?,
                a.b(7),
            )
            .map(|()| Json::Null)
        }
        "setHidden" => {
            let i = sheet(b, a.u(0))?;
            b.set_hidden(
                i,
                a.b(1),
                a.u(2).saturating_sub(1),
                a.u(3).saturating_sub(1),
                a.b(4),
            )
            .map(|()| Json::Null)
        }
        "hidden" => {
            let i = sheet(b, a.u(0))?;
            Ok(json!(b.hidden(
                i,
                a.b(1),
                a.u(2).saturating_sub(1),
                a.u(3).saturating_sub(1)
            )))
        }
        "visibleAreas" => {
            let (i, r) = a.sheet_range(b, 0)?;
            Ok(areas(&b.visible_areas(i, r)))
        }
        "specialCells" => {
            let (i, r) = a.sheet_range(b, 0)?;
            let what = special(a.i(5), a.i(6))?;
            b.special_cells(i, r, what).map(|l| areas(&l))
        }
        "format" => {
            let (i, r) = a.sheet_range(b, 0)?;
            Ok(b.format_of(i, r, &a.s(5))?
                .as_ref()
                .map(to_json)
                .unwrap_or(Json::Null))
        }
        "setFormat" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.set_format(i, r, &a.s(5), &from_json(a.at(6))?)
                .map(|()| Json::Null)
        }
        "clearFormats" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.clear_formats(i, r).map(|()| Json::Null)
        }
        "merge" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.merge_cells(i, r, a.b(5)).map(|()| Json::Null)
        }
        "mergeArea" => {
            let (i, at) = a.sheet_cell(b, 0)?;
            Ok(bounds(b.merge_area(i, at)))
        }
        "tabColor" => {
            let i = sheet(b, a.u(0))?;
            Ok(json!(b.tab_color(i)))
        }
        "setTabColor" => {
            let i = sheet(b, a.u(0))?;
            b.set_tab_color(i, a.at(1).as_f64());
            none
        }
        "copySheet" => {
            let i = sheet(b, a.u(0))?;
            b.copy_sheet(i, a.u(1) as usize).map(|n| json!(n))
        }
        "setManual" => {
            b.set_manual(a.b(0));
            none
        }
        "savedManual" => Ok(json!(b.saved_manual())),
        "recalculate" => {
            b.recalculate();
            none
        }
        "calculatedSheets" => Ok(json!(b.calculated_sheets())),
        "clear" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.clear(i, r, a.b(5)).map(|()| Json::Null)
        }
        "setNumberFormat" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.set_number_format(i, r, &a.s(5)).map(|()| Json::Null)
        }
        "usedRange" => {
            let i = sheet(b, a.u(0))?;
            Ok(b.used_range(i).map(bounds).unwrap_or(Json::Null))
        }
        "end" => {
            let (i, at) = a.sheet_cell(b, 0)?;
            let at = b.end(i, at, direction(a.i(3))?);
            Ok(json!([at.row + 1, at.col + 1]))
        }
        "currentRegion" => {
            let (i, at) = a.sheet_cell(b, 0)?;
            Ok(bounds(b.current_region(i, at)))
        }
        "evaluate" => {
            let i = sheet(b, a.u(0))?;
            Ok(value_json(&b.evaluate(i, &a.s(1))))
        }
        "resolve" => {
            let i = sheet(b, a.u(0))?;
            Ok(resolved(b.resolve(i, &a.s(1))))
        }
        "hyperlinks" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.hyperlinks(i, r).map(hyperlinks_json)
        }
        "setHyperlink" => {
            let (i, at) = a.sheet_cell(b, 0)?;
            b.set_hyperlink(i, at, &a.s(3), &a.s(4), &a.s(5))
                .map(|()| Json::Null)
        }
        "removeHyperlinks" => {
            let (i, r) = a.sheet_range(b, 0)?;
            b.remove_hyperlinks(i, r).map(|n| json!(n))
        }
        "names" => Ok(names_json(b)),
        "addName" => {
            let scope = a.i(1);
            let scope = if scope < 0 {
                None
            } else {
                Some(sheet(b, scope as u32)?)
            };
            b.add_name(&a.s(0), scope, &a.s(2)).map(|()| Json::Null)
        }
        "deleteName" => {
            let scope = a.i(1);
            b.delete_name(&a.s(0), (scope >= 0).then_some(scope as usize));
            none
        }
        "calculate" => {
            b.calculate();
            none
        }
        _ => Err(format!("Book に {method} はありません")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every method of the Node addon (`crates/ras-book/src/node.rs`) is one [`call`] takes, so
    /// that the run time finds the same `Book` in the browser.
    #[test]
    fn call_has_every_method_of_the_addon() {
        let node = include_str!("../../ras-book/src/node.rs");
        let mut missing = Vec::new();
        let mut lines = node.lines().peekable();
        while let Some(line) = lines.next() {
            if !line.trim_start().starts_with("#[napi") || line.contains("js_name = \"Book\"") {
                continue;
            }
            let js_name = line
                .split("js_name = \"")
                .nth(1)
                .and_then(|r| r.split('"').next())
                .map(str::to_string);
            // The attribute of the `impl` block.
            if lines
                .peek()
                .is_some_and(|l| l.trim_start().starts_with("impl"))
            {
                continue;
            }
            let factory = line.contains("factory");
            let Some(sig) = lines.find(|l| l.contains("pub fn ")) else {
                break;
            };
            let name = sig
                .split("pub fn ")
                .nth(1)
                .and_then(|r| r.split('(').next())
                .unwrap_or_default();
            let camel = js_name.unwrap_or_else(|| {
                let mut out = String::new();
                let mut up = false;
                for ch in name.chars() {
                    if ch == '_' {
                        up = true;
                    } else if up {
                        out.extend(ch.to_uppercase());
                        up = false;
                    } else {
                        out.push(ch);
                    }
                }
                out
            });
            // Opening and saving files (and reading the controls of one), and the engine's
            // functions, are the bindings' own.
            if factory
                || matches!(
                    camel.as_str(),
                    "toBytes" | "save" | "isFunction" | "controls"
                )
            {
                continue;
            }
            if !METHODS.contains(&camel.as_str()) {
                missing.push(camel);
            }
        }
        assert!(missing.is_empty(), "js::call lacks {missing:?}");
        let mut b = Book::new_book();
        for m in METHODS {
            assert_ne!(
                call(&mut b, m, &[]).err().as_deref(),
                Some(format!("Book に {m} はありません").as_str()),
                "{m} is in METHODS but not in call"
            );
        }
    }

    #[test]
    fn values_cross_as_the_addon_passes_them() {
        let mut b = Book::new_book();
        call(
            &mut b,
            "setValues",
            &[
                json!(0),
                json!(1),
                json!(1),
                json!(1),
                json!(2),
                json!([[1, "a"]]),
            ],
        )
        .unwrap();
        let got = call(
            &mut b,
            "get",
            &[
                json!(0),
                json!(1),
                json!(1),
                json!(1),
                json!(2),
                json!("Value"),
            ],
        )
        .unwrap();
        assert_eq!(got, json!([[1.0, "a"]]));
        assert_eq!(
            call(&mut b, "usedRange", &[json!(0)]).unwrap(),
            json!([1, 1, 1, 2])
        );
    }
}
