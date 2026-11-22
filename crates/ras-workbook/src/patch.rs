//! Edits of a workbook file made as Excel makes them — cells typed into, cleared, copied and
//! moved, cells, rows and columns inserted and deleted — and written back into the file
//! (`ras book edit`, the grid of the editor).
//!
//! The edit itself is the stand-in's ([`Book`]): it knows what moves, which formulas follow and
//! what they come to. Saving the stand-in's workbook would lose what it does not hold (the
//! controls of the sheets, the signature of the VBA project, ...), so the file is not written
//! from it. Two snapshots of its cells — the workbook as read, and after the edit — tell which
//! cells changed, and only those are written again into the parts of the file; every other
//! byte of the package stays what it was.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Cursor, Read as _, Write as _};
use std::ops::Range;

use gridcraft_core::{CellRef, RangeRef};
use gridcraft_formula::adjust::{Axis, Edit};
use serde_json::Value as Json;

use crate::book::snapshot::{CellState, Raw, Snapshot, Styles};
use crate::book::{Book, MAX_COLS, MAX_ROWS, Paste, PasteOp, PasteWhat, SortKey, Val};

mod macros;
mod panes;
mod scan;
mod structure;
mod styles;
#[cfg(test)]
mod tests;

use scan::Edits;

type Result<T> = std::result::Result<T, String>;

/// The most cells one entry (`set`, `values`, `copy`, `fill`) writes.
const MAX_CELLS: u64 = 200_000;

/// One edit, as the grid asks for it.
#[derive(Clone, Debug, PartialEq)]
pub enum Op {
    /// Text typed into every cell of `range` (a formula's relative references move from the
    /// first cell, as with Ctrl+Enter).
    Set {
        sheet: String,
        range: RangeRef,
        text: String,
    },
    /// Texts typed into the cells from `at` on, row by row (text pasted from outside).
    Values {
        sheet: String,
        at: CellRef,
        rows: Vec<Vec<String>>,
    },
    /// The contents of the cells go; their formats stay.
    Clear { sheet: String, range: RangeRef },
    /// Cells inserted at `range`: those there move down (`right`: to the right). Whole rows
    /// and columns when `range` spans the sheet.
    Insert {
        sheet: String,
        range: RangeRef,
        right: bool,
    },
    /// The cells of `range` deleted: those below move up (`left`: those to the right, left).
    Delete {
        sheet: String,
        range: RangeRef,
        left: bool,
    },
    /// The cells of `from` pasted at `to` (repeated over it when it is a multiple of `from`).
    Copy {
        sheet: String,
        from_sheet: String,
        from: RangeRef,
        to: RangeRef,
    },
    /// What of the cells of `from` is pasted at `to` ("形式を選択して貼り付け"): formulas, values,
    /// formats, ..., with an operation on the cells there, empty cells skipped, rows and
    /// columns swapped.
    PasteSpecial {
        sheet: String,
        from_sheet: String,
        from: RangeRef,
        to: RangeRef,
        paste: Paste,
    },
    /// The cells of `from` moved to `to` (cut and paste).
    Move {
        sheet: String,
        from_sheet: String,
        from: RangeRef,
        to: CellRef,
    },
    /// The first row of `range` copied into its other rows (`right`: the first column into
    /// the other columns).
    Fill {
        sheet: String,
        range: RangeRef,
        right: bool,
    },
    /// Formats given to the cells of `range` (`set`: what the grid's toolbar sets).
    Format {
        sheet: String,
        range: RangeRef,
        set: Vec<(String, Json)>,
    },
    /// The cells get the default format, and are no longer merged.
    ClearFormats { sheet: String, range: RangeRef },
    /// The width of whole columns, or the height of whole rows, in pixels as the grid draws.
    Size {
        sheet: String,
        range: RangeRef,
        columns: bool,
        px: f64,
    },
    /// Whole rows or columns hidden, or shown again.
    Hide {
        sheet: String,
        range: RangeRef,
        columns: bool,
        hidden: bool,
    },
    /// The cells of `range` merged (`across`: each of its rows on its own), or taken apart.
    Merge {
        sheet: String,
        range: RangeRef,
        across: bool,
        on: bool,
    },
    /// The rows of `range` sorted by the columns of `keys` (each descending or not), the first
    /// first. A single cell stands for the data around it; `header`: its first row is headings
    /// (None: guessed).
    Sort {
        sheet: String,
        range: RangeRef,
        keys: Vec<(u32, bool)>,
        header: Option<bool>,
    },
    /// The sheet's filter put on `range` (a single cell: the data around it), or taken away.
    Filter {
        sheet: String,
        range: Option<RangeRef>,
    },
    /// The rows shown by a column of the filter: those whose cell shows one of a list of values
    /// (and the empty ones), passes one or two comparisons, or is among the largest or smallest;
    /// None: no longer filtered by the column.
    FilterBy {
        sheet: String,
        column: u32,
        show: Option<crate::book::FilterCriterion>,
    },
    /// The fill handle: `from` continued over `to` (a series of numbers or dates, texts that
    /// end in a number, formulas with their references moved; anything else repeated).
    Autofill {
        sheet: String,
        from: RangeRef,
        to: RangeRef,
    },
    /// The macro (`Sheet1.button1Click`) a shape or a form control runs when it is clicked:
    /// Excel's "マクロの登録". The object by its name, or a form control by its shape (`shape`:
    /// 1025 for `_x0000_s1025`).
    Macro {
        sheet: String,
        object: String,
        shape: Option<u32>,
        name: String,
    },
    /// The rows above `at` and the columns to its left frozen ("ウィンドウ枠の固定"); None (or
    /// A1): no longer frozen.
    Freeze { sheet: String, at: Option<CellRef> },
}

impl Op {
    fn structural(&self) -> bool {
        matches!(
            self,
            Op::Insert { .. } | Op::Delete { .. } | Op::Macro { .. } | Op::Freeze { .. }
        )
    }
}

fn range_of(o: &Json, key: &str) -> Result<RangeRef> {
    let text = o
        .get(key)
        .and_then(Json::as_str)
        .ok_or_else(|| format!("`{key}` がありません"))?;
    RangeRef::parse(&text.replace('$', "")).ok_or_else(|| {
        format!("範囲 `{text}` を読み取れません（A1:B2、3:5、B:D のように指定してください）")
    })
}

/// Whole rows (`3:5`) or whole columns (`B:D`: true) of `range`.
fn lines_of(o: &Json) -> Result<(RangeRef, bool)> {
    let range = range_of(o, "range")?;
    let text = o.get("range").and_then(Json::as_str).unwrap_or_default();
    let columns = text
        .trim_start_matches('$')
        .starts_with(|c: char| c.is_ascii_alphabetic());
    let whole = if columns {
        range.is_full_cols()
    } else {
        range.is_full_rows()
    };
    if !whole {
        return Err(format!(
            "範囲 `{text}` は、行（3:5）か列（B:D）で指定してください"
        ));
    }
    Ok((range, columns))
}

/// A column given by its letters (`"B"`).
fn column_of(o: &Json, key: &str) -> Result<u32> {
    let text = o
        .get(key)
        .and_then(Json::as_str)
        .ok_or_else(|| format!("`{key}` がありません"))?;
    gridcraft_core::addr::letters_to_col(&text.trim().to_ascii_uppercase())
        .ok_or_else(|| format!("列 `{text}` を読み取れません（B のように指定してください）"))
}

/// What a column of the filter shows: `values` (and `blanks`), `criteria` (one or two
/// comparisons `[op, value]`, op one of = <> > >= < <=, the value with `*` and `?`; `and`), or
/// `top` / `bottom` (how many, `percent`).
fn criterion_of(o: &Json) -> Result<crate::book::FilterCriterion> {
    use crate::book::FilterCriterion as C;
    let text = |v: &Json| match v {
        Json::String(s) => s.clone(),
        other => other.to_string(),
    };
    if let Some(list) = o.get("criteria").and_then(Json::as_array) {
        let test = |v: &Json| -> Result<(String, String)> {
            let pair = v.as_array().filter(|p| p.len() == 2);
            let (op, value) = pair
                .map(|p| (text(&p[0]), text(&p[1])))
                .ok_or("`criteria` は [演算子, 値] の配列です（例: [[\">=\", \"100\"]]）")?;
            if !["=", "<>", ">", ">=", "<", "<="].contains(&op.as_str()) {
                return Err(format!("演算子 `{op}` は指定できません（= <> > >= < <=）"));
            }
            Ok((op, value))
        };
        let a = test(list.first().ok_or("`criteria` が空です")?)?;
        let b = list.get(1).map(test).transpose()?;
        if list.len() > 2 {
            return Err("`criteria` は 2 つまでです".into());
        }
        return Ok(C::Custom {
            a,
            b,
            and: o.get("and").and_then(Json::as_bool).unwrap_or(true),
        });
    }
    for (key, bottom) in [("top", false), ("bottom", true)] {
        if let Some(n) = o.get(key) {
            let percent = o.get("percent").and_then(Json::as_bool).unwrap_or(false);
            let count = n
                .as_u64()
                .filter(|n| (1..=if percent { 100 } else { 500 }).contains(n))
                .ok_or(format!(
                    "`{key}` は 1〜500 の数（`percent` なら 1〜100）です"
                ))?;
            return Ok(C::Top10 {
                bottom,
                percent,
                count: count as u32,
            });
        }
    }
    Ok(C::Values {
        values: o
            .get("values")
            .and_then(Json::as_array)
            .ok_or("`values`・`criteria`・`top`・`bottom` のどれかを指定してください")?
            .iter()
            .map(text)
            .collect(),
        blanks: o.get("blanks").and_then(Json::as_bool).unwrap_or(false),
    })
}

/// The edits of a JSON value: one object, or an array of them. `sheet`: the sheet of those that
/// name none.
pub fn parse_ops(json: &Json, sheet: Option<&str>) -> Result<Vec<Op>> {
    let list: Vec<&Json> = match json {
        Json::Array(a) => a.iter().collect(),
        other => vec![other],
    };
    let mut out = Vec::new();
    for o in list {
        let kind = o
            .get("op")
            .and_then(Json::as_str)
            .ok_or("編集の `op` がありません")?;
        let text = |key: &str| o.get(key).and_then(Json::as_str).map(str::to_string);
        let sheet = text("sheet")
            .or_else(|| sheet.map(str::to_string))
            .ok_or("シートが指定されていません（`sheet`）")?;
        let from_sheet = text("fromSheet").unwrap_or_else(|| sheet.clone());
        let shift = text("shift").unwrap_or_default();
        out.push(match kind {
            "set" => Op::Set {
                sheet,
                range: range_of(o, "range")?,
                text: text("text").ok_or("`text` がありません")?,
            },
            "values" => {
                let at = range_of(o, "at")?.start;
                let rows = o
                    .get("rows")
                    .and_then(Json::as_array)
                    .ok_or("`rows` がありません")?
                    .iter()
                    .map(|row| {
                        row.as_array()
                            .map(|cells| {
                                cells
                                    .iter()
                                    .map(|c| match c {
                                        Json::String(s) => s.clone(),
                                        Json::Null => String::new(),
                                        other => other.to_string(),
                                    })
                                    .collect()
                            })
                            .ok_or("`rows` は行の配列の配列です")
                    })
                    .collect::<std::result::Result<Vec<Vec<String>>, _>>()?;
                Op::Values { sheet, at, rows }
            }
            "clear" => Op::Clear {
                sheet,
                range: range_of(o, "range")?,
            },
            "insert" | "delete" => {
                let range = range_of(o, "range")?;
                // Whole rows go down (up), whole columns right (left), whatever is asked.
                let across = if range.is_full_rows() && !range.is_full_cols() {
                    false
                } else if range.is_full_cols() && !range.is_full_rows() {
                    true
                } else {
                    match shift.as_str() {
                        "" | "down" | "up" => false,
                        "right" | "left" => true,
                        other => return Err(format!("`shift` の `{other}` は指定できません")),
                    }
                };
                if kind == "insert" {
                    Op::Insert {
                        sheet,
                        range,
                        right: across,
                    }
                } else {
                    Op::Delete {
                        sheet,
                        range,
                        left: across,
                    }
                }
            }
            "copy" => Op::Copy {
                sheet,
                from_sheet,
                from: range_of(o, "from")?,
                to: range_of(o, "to")?,
            },
            "pasteSpecial" => Op::PasteSpecial {
                sheet,
                from_sheet,
                from: range_of(o, "from")?,
                to: range_of(o, "to")?,
                paste: Paste {
                    what: match text("paste").as_deref().unwrap_or("all") {
                        "all" => PasteWhat::All,
                        "allExceptBorders" => PasteWhat::AllExceptBorders,
                        "formulas" => PasteWhat::Formulas,
                        "formulasAndNumberFormats" => PasteWhat::FormulasAndNumberFormats,
                        "values" => PasteWhat::Values,
                        "valuesAndNumberFormats" => PasteWhat::ValuesAndNumberFormats,
                        "formats" => PasteWhat::Formats,
                        "columnWidths" => PasteWhat::ColumnWidths,
                        other => return Err(format!("`paste` の `{other}` は指定できません（all / allExceptBorders / formulas / formulasAndNumberFormats / values / valuesAndNumberFormats / formats / columnWidths）")),
                    },
                    op: match text("operation").as_deref().unwrap_or("none") {
                        "none" => PasteOp::None,
                        "add" => PasteOp::Add,
                        "subtract" => PasteOp::Subtract,
                        "multiply" => PasteOp::Multiply,
                        "divide" => PasteOp::Divide,
                        other => return Err(format!("`operation` の `{other}` は指定できません（none / add / subtract / multiply / divide）")),
                    },
                    skip_blanks: o.get("skipBlanks").and_then(Json::as_bool).unwrap_or(false),
                    transpose: o.get("transpose").and_then(Json::as_bool).unwrap_or(false),
                },
            },
            "move" => Op::Move {
                sheet,
                from_sheet,
                from: range_of(o, "from")?,
                to: range_of(o, "to")?.start,
            },
            "fill" => Op::Fill {
                sheet,
                range: range_of(o, "range")?,
                right: match text("dir").as_deref() {
                    None | Some("down") => false,
                    Some("right") => true,
                    Some(other) => return Err(format!("`dir` の `{other}` は指定できません")),
                },
            },
            "format" => Op::Format {
                sheet,
                range: range_of(o, "range")?,
                set: o
                    .get("set")
                    .and_then(Json::as_object)
                    .ok_or("`set` がありません")?
                    .iter()
                    .map(|(k, v)| (k.clone(), v.clone()))
                    .collect(),
            },
            "clearFormats" => Op::ClearFormats {
                sheet,
                range: range_of(o, "range")?,
            },
            "size" => {
                let (range, columns) = lines_of(o)?;
                Op::Size {
                    sheet,
                    range,
                    columns,
                    px: o
                        .get("px")
                        .and_then(Json::as_f64)
                        .filter(|px| (0.0..=4000.0).contains(px))
                        .ok_or("`px` がありません")?,
                }
            }
            "hide" => {
                let (range, columns) = lines_of(o)?;
                Op::Hide {
                    sheet,
                    range,
                    columns,
                    hidden: o.get("hidden").and_then(Json::as_bool).unwrap_or(true),
                }
            }
            "merge" | "unmerge" => Op::Merge {
                sheet,
                range: range_of(o, "range")?,
                across: o.get("across").and_then(Json::as_bool).unwrap_or(false),
                on: kind == "merge",
            },
            "sort" => Op::Sort {
                sheet,
                range: range_of(o, "range")?,
                keys: match o.get("keys").and_then(Json::as_array) {
                    // `keys`: [{ "by": "B", "desc": true }, ...] (Excel takes 64 at most).
                    Some(list) => {
                        if list.is_empty() || list.len() > 64 {
                            return Err("`keys` には 1〜64 個の列を指定してください".into());
                        }
                        list.iter()
                            .map(|k| {
                                Ok((
                                    column_of(k, "by")?,
                                    k.get("desc").and_then(Json::as_bool).unwrap_or(false),
                                ))
                            })
                            .collect::<Result<Vec<_>>>()?
                    }
                    None => vec![(
                        column_of(o, "by")?,
                        o.get("desc").and_then(Json::as_bool).unwrap_or(false),
                    )],
                },
                header: o.get("header").and_then(Json::as_bool),
            },
            "filter" => Op::Filter {
                sheet,
                range: if o.get("off").and_then(Json::as_bool).unwrap_or(false) {
                    None
                } else {
                    Some(range_of(o, "range")?)
                },
            },
            "filterBy" => Op::FilterBy {
                sheet,
                column: column_of(o, "column")?,
                show: if o.get("clear").and_then(Json::as_bool).unwrap_or(false) {
                    None
                } else {
                    Some(criterion_of(o)?)
                },
            },
            "autofill" => Op::Autofill {
                sheet,
                from: range_of(o, "from")?,
                to: range_of(o, "to")?,
            },
            "macro" => {
                let name = text("macro").unwrap_or_default();
                let fine = !name.is_empty()
                    && name
                        .split('.')
                        .all(|p| !p.is_empty() && p.chars().all(|c| c.is_alphanumeric() || c == '_'));
                if !fine {
                    return Err(format!("マクロの名前 `{name}` は指定できません（Module1.macro1 のように指定してください）"));
                }
                let object = text("object").unwrap_or_default();
                let shape = o
                    .get("shape")
                    .and_then(Json::as_u64)
                    .and_then(|n| u32::try_from(n).ok());
                if object.is_empty() && shape.is_none() {
                    return Err("`object`（オブジェクトの名前）か `shape` がありません".into());
                }
                Op::Macro {
                    sheet,
                    object,
                    shape,
                    name,
                }
            }
            "freeze" => Op::Freeze {
                sheet,
                at: if o.get("off").and_then(Json::as_bool).unwrap_or(false) {
                    None
                } else {
                    let at = range_of(o, "at")?;
                    if !at.is_single() {
                        return Err("`at` はセル 1 つで指定してください（B2 なら 1 行目と A 列を固定します）".into());
                    }
                    Some(at.start)
                },
            },
            other => return Err(format!("編集 `{other}` はありません（set / values / clear / insert / delete / copy / move / fill / autofill / pasteSpecial / format / clearFormats / size / hide / merge / unmerge / sort / filter / filterBy / macro / freeze）")),
        });
    }
    Ok(out)
}

/// The workbook file `bytes` (.xlsx / .xlsm) with `ops` made, in order.
pub fn apply(bytes: &[u8], ops: &[Op]) -> Result<Vec<u8>> {
    let mut bytes = bytes.to_vec();
    // Edits that move no cells are made together; one that does (an insertion, a deletion) on
    // its own: where the cells of the file go is then one plain mapping.
    let mut at = 0;
    while at < ops.len() {
        let mut end = at + 1;
        if !ops[at].structural() {
            while end < ops.len() && !ops[end].structural() {
                end += 1;
            }
        }
        bytes = match &ops[at] {
            // The drawing, not the cells.
            Op::Macro {
                sheet,
                object,
                shape,
                name,
            } => macros::assign(&bytes, sheet, object, *shape, name)?,
            Op::Freeze { sheet, at } => panes::freeze(&bytes, sheet, *at)?,
            _ => apply_group(&bytes, &ops[at..end])?,
        };
        at = end;
    }
    Ok(bytes)
}

/// The cells an insertion or a deletion moves, as an edit of the formulas' references.
fn edit_of(range: RangeRef, across: bool, insert: bool) -> Edit {
    let axis = if across { Axis::Cols } else { Axis::Rows };
    let whole = match axis {
        Axis::Rows => range.is_full_rows(),
        Axis::Cols => range.is_full_cols(),
    };
    let (at, count) = match axis {
        Axis::Rows => (range.start.row, range.height()),
        Axis::Cols => (range.start.col, range.width()),
    };
    match (whole, insert) {
        (true, true) => Edit::Insert { axis, at, count },
        (true, false) => Edit::Delete { axis, at, count },
        (false, true) => Edit::InsertCells { axis, range },
        (false, false) => Edit::DeleteCells { axis, range },
    }
}

fn sheet_index(book: &Book, name: &str) -> Result<usize> {
    book.find_sheet(name)
        .ok_or_else(|| format!("シート `{name}` がありません"))
}

fn sized(at: CellRef, height: u32, width: u32) -> Result<RangeRef> {
    at.offset(height as i64 - 1, width as i64 - 1)
        .map(|end| RangeRef::new(at, end))
        .ok_or_else(|| "貼り付け先がシートの外です".to_string())
}

fn limit(cells: u64) -> Result<()> {
    if cells > MAX_CELLS {
        return Err(format!(
            "範囲が大きすぎます（一度に書き込めるのは {MAX_CELLS} セルまでです）"
        ));
    }
    Ok(())
}

/// `from` pasted at `to`, repeated over it when it is a multiple of `from` (as Excel pastes).
fn copy(book: &mut Book, fi: usize, from: RangeRef, ti: usize, to: RangeRef) -> Result<RangeRef> {
    paste(book, fi, from, ti, to, &Paste::all())
}

/// What `p` says of `from` pasted at `to`, repeated over it when it is a multiple of `from`
/// (turned, when the rows and columns are swapped).
fn paste(
    book: &mut Book,
    fi: usize,
    from: RangeRef,
    ti: usize,
    to: RangeRef,
    p: &Paste,
) -> Result<RangeRef> {
    let (h, w) = if p.transpose {
        (from.width(), from.height())
    } else {
        (from.height(), from.width())
    };
    let (rows, cols) = if to.height() % h == 0 && to.width() % w == 0 {
        (to.height() / h, to.width() / w)
    } else {
        (1, 1)
    };
    let target = sized(to.start, h * rows, w * cols)?;
    if u64::from(rows) * u64::from(cols) > 10_000 {
        return Err("貼り付け先が大きすぎます".into());
    }
    for r in 0..rows {
        for c in 0..cols {
            let at = CellRef::new(to.start.row + r * h, to.start.col + c * w);
            book.paste(fi, from, ti, at, p)?;
        }
    }
    Ok(target)
}

/// The data a single cell stands for: the filter's range when the cell is in it, else the
/// cells around it; a range of several cells is itself.
fn around(book: &mut Book, si: usize, range: RangeRef) -> RangeRef {
    if !range.is_single() {
        return range;
    }
    match book.auto_filter(si) {
        Some((filter, _)) if filter.contains(range.start) => filter,
        _ => book.current_region(si, range.start),
    }
}

/// VBA's colour number (blue, green, red) of `#RRGGBB`.
fn colour(text: &str) -> Result<f64> {
    let hex = text.trim().trim_start_matches('#');
    let rgb = u32::from_str_radix(hex, 16)
        .ok()
        .filter(|_| hex.len() == 6)
        .ok_or_else(|| format!("色 `{text}` を読み取れません（#RRGGBB で指定してください）"))?;
    Ok(f64::from(
        ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | (rgb >> 16),
    ))
}

/// One format of the toolbar given to the cells of `range`.
fn format(book: &mut Book, si: usize, range: RangeRef, key: &str, value: &Json) -> Result<()> {
    const NONE: f64 = -4142.0;
    let wrong = || format!("書式 `{key}` の値 {value} は指定できません");
    let flag = || value.as_bool().map(Val::Bool).ok_or_else(wrong);
    let text = || value.as_str().ok_or_else(wrong);
    let mut set = |prop: &str, v: Val| book.set_format(si, range, prop, &v);
    match key {
        "bold" => set("font.bold", flag()?),
        "italic" => set("font.italic", flag()?),
        "underline" => set("font.underline", flag()?),
        "strike" => set("font.strikethrough", flag()?),
        "wrap" => set("wraptext", flag()?),
        "fontName" => set("font.name", Val::Text(text()?.to_string())),
        "fontSize" => set(
            "font.size",
            Val::Num(
                value
                    .as_f64()
                    .filter(|s| (1.0..=409.0).contains(s))
                    .ok_or_else(wrong)?,
            ),
        ),
        "color" => match value {
            // Automatic: black.
            Json::Null => set("font.color", Val::Num(0.0)),
            _ => set("font.color", Val::Num(colour(text()?)?)),
        },
        "fill" => match value {
            Json::Null => set("interior.pattern", Val::Num(NONE)),
            _ => set("interior.color", Val::Num(colour(text()?)?)),
        },
        "align" => set(
            "horizontalalignment",
            Val::Num(match text()? {
                "left" => -4131.0,
                "center" => -4108.0,
                "right" => -4152.0,
                "general" => 1.0,
                _ => return Err(wrong()),
            }),
        ),
        "valign" => set(
            "verticalalignment",
            Val::Num(match text()? {
                "top" => -4160.0,
                "middle" => -4108.0,
                "bottom" => -4107.0,
                _ => return Err(wrong()),
            }),
        ),
        "numberFormat" => book.set_number_format(si, range, text()?),
        "border" => {
            // `Borders(index)`: 7 left, 8 top, 9 bottom, 10 right, 11 and 12 the lines inside.
            const THIN: f64 = 2.0;
            const THICK: f64 = 4.0;
            let (sides, weight): (&[&str], f64) = match text()? {
                "all" => (&["7", "8", "9", "10", "11", "12"], THIN),
                "outline" => (&["7", "8", "9", "10"], THIN),
                "thick" => (&["7", "8", "9", "10"], THICK),
                "top" => (&["8"], THIN),
                "bottom" => (&["9"], THIN),
                "left" => (&["7"], THIN),
                "right" => (&["10"], THIN),
                "none" => (&["7", "8", "9", "10", "11", "12"], NONE),
                _ => return Err(wrong()),
            };
            for side in sides {
                if weight == NONE {
                    set(&format!("borders.{side}.linestyle"), Val::Num(NONE))?;
                } else {
                    set(&format!("borders.{side}.linestyle"), Val::Num(1.0))?;
                    set(&format!("borders.{side}.weight"), Val::Num(weight))?;
                }
            }
            Ok(())
        }
        _ => Err(format!("書式 `{key}` はありません")),
    }
}

/// A text with a number in it (the last one): what is before it, the number, how many digits
/// it has, and what is after it.
fn numbered(text: &str) -> Option<(&str, i64, usize, &str)> {
    let end = text.rfind(|c: char| c.is_ascii_digit())? + 1;
    let digits = text[..end]
        .chars()
        .rev()
        .take_while(char::is_ascii_digit)
        .count();
    if digits > 15 {
        return None;
    }
    let start = end - digits;
    Some((
        &text[..start],
        text[start..end].parse().ok()?,
        digits,
        &text[end..],
    ))
}

/// The fill handle: the cells of `from` continued over `to`, which is `from` made longer one way.
fn autofill(book: &mut Book, si: usize, from: RangeRef, to: RangeRef) -> Result<()> {
    use crate::book::Read;
    if !to.contains_range(&from) || to == from {
        return Err("オートフィルの範囲が正しくありません".into());
    }
    let down = to.start.col == from.start.col && to.end.col == from.end.col;
    let across = to.start.row == from.start.row && to.end.row == from.end.row;
    if down == across {
        return Err("オートフィルは、上下か左右のどちらかにだけ行えます".into());
    }
    limit(to.count())?;
    // Along the fill: `n` lines of the source, the lines of `to` at -before..n+after.
    let (n, before, after) = if down {
        (
            from.height() as i64,
            (from.start.row - to.start.row) as i64,
            (to.end.row - from.end.row) as i64,
        )
    } else {
        (
            from.width() as i64,
            (from.start.col - to.start.col) as i64,
            (to.end.col - from.end.col) as i64,
        )
    };
    let lines: Vec<i64> = (-before..0).rev().chain(n..n + after).collect();
    let cell = |k: i64, j: u32| {
        if down {
            CellRef::new((from.start.row as i64 + k) as u32, from.start.col + j)
        } else {
            CellRef::new(from.start.row + j, (from.start.col as i64 + k) as u32)
        }
    };
    let width = if down { from.width() } else { from.height() };
    let typed = book.get(si, from, Read::Formula)?;
    let values = book.get(si, from, Read::Value)?;
    let source = |k: i64, j: u32| -> (usize, usize) {
        if down {
            (k as usize, j as usize)
        } else {
            (j as usize, k as usize)
        }
    };
    // What each line across the fill continues as: a series of numbers, of numbered texts, or
    // (None) its cells again.
    enum Series {
        Numbers(f64, f64),
        Texts(String, i64, i64, usize, String),
    }
    let mut series: Vec<Option<Series>> = Vec::new();
    for j in 0..width {
        let cells: Vec<(&Val, &Val)> = (0..n)
            .map(|k| {
                let (r, c) = source(k, j);
                (&typed[r][c], &values[r][c])
            })
            .collect();
        let formula = cells
            .iter()
            .any(|(t, _)| matches!(t, Val::Text(s) if s.starts_with('=')));
        let numbers: Option<Vec<f64>> = cells
            .iter()
            .map(|(_, v)| match v {
                Val::Num(x) | Val::Date(x) => Some(*x),
                _ => None,
            })
            .collect();
        let dates = cells.iter().all(|(_, v)| matches!(v, Val::Date(_)));
        let texts: Option<Vec<(&str, i64, usize, &str)>> = cells
            .iter()
            .map(|(_, v)| match v {
                Val::Text(s) => numbered(s),
                _ => None,
            })
            .collect();
        series.push(if formula {
            None
        } else if let Some(xs) = numbers {
            if xs.len() >= 2 {
                // The line through the numbers (least squares), as Excel continues them.
                let m = xs.len() as f64;
                let mean_k = (m - 1.0) / 2.0;
                let mean_x = xs.iter().sum::<f64>() / m;
                let (mut num, mut den) = (0.0, 0.0);
                for (k, x) in xs.iter().enumerate() {
                    num += (k as f64 - mean_k) * (x - mean_x);
                    den += (k as f64 - mean_k).powi(2);
                }
                let slope = num / den;
                Some(Series::Numbers(mean_x - slope * mean_k, slope))
            } else if dates {
                // One date: the days after it.
                Some(Series::Numbers(xs[0], 1.0))
            } else {
                None
            }
        } else if let Some(ts) = texts {
            let step = if ts.len() >= 2 { ts[1].1 - ts[0].1 } else { 1 };
            let regular = ts
                .windows(2)
                .all(|w| w[0].0 == w[1].0 && w[0].3 == w[1].3 && w[1].1 - w[0].1 == step);
            regular.then(|| {
                Series::Texts(
                    ts[0].0.to_string(),
                    ts[0].1,
                    step,
                    ts[0].2,
                    ts[0].3.to_string(),
                )
            })
        } else {
            None
        });
    }
    for k in lines {
        // The cells of the source again, with their formats; formulas with their references moved.
        let from_k = k.rem_euclid(n);
        let line = RangeRef::new(cell(from_k, 0), cell(from_k, width - 1));
        book.copy(si, line, si, cell(k, 0))?;
        for (j, s) in series.iter().enumerate() {
            let at = RangeRef::cell(cell(k, j as u32));
            match s {
                Some(Series::Numbers(first, step)) => {
                    // (Sums of tenths are not exact: 0.1, 0.2, 0.30000000000000004.)
                    let v = first + step * k as f64;
                    let v = (v * 1e10).round() / 1e10;
                    book.set_values(si, at, &[vec![Val::Num(v)]])?;
                }
                Some(Series::Texts(head, first, step, digits, tail)) => {
                    let v = (first + step * k).abs();
                    let text = format!("{head}{v:0digits$}{tail}");
                    book.set_values(si, at, &[vec![Val::Text(text)]])?;
                }
                None => {}
            }
        }
    }
    Ok(())
}

/// Makes one edit in the stand-in. `written`: the cells it wrote to; `moved`: the cells it
/// moved, for an insertion or a deletion.
fn run(
    book: &mut Book,
    op: &Op,
    written: &mut Vec<(usize, RangeRef)>,
    moved: &mut Option<(usize, Edit)>,
) -> Result<()> {
    match op {
        Op::Set { sheet, range, text } => {
            let si = sheet_index(book, sheet)?;
            limit(range.count())?;
            if text.is_empty() {
                book.clear(si, *range, false)?;
            } else {
                book.set_formulas(si, *range, &[vec![Val::Text(text.clone())]], true)?;
            }
            written.push((si, *range));
        }
        Op::Values { sheet, at, rows } => {
            let si = sheet_index(book, sheet)?;
            let width = rows.iter().map(Vec::len).max().unwrap_or(0);
            if width == 0 {
                return Ok(());
            }
            let range = sized(*at, rows.len() as u32, width as u32)?;
            limit(range.count())?;
            let rows: Vec<Vec<Val>> = rows
                .iter()
                .map(|row| {
                    (0..width)
                        .map(|i| Val::Text(row.get(i).cloned().unwrap_or_default()))
                        .collect()
                })
                .collect();
            // Each text as it is: a formula keeps the references it was written with.
            for (r, row) in rows.iter().enumerate() {
                for (c, v) in row.iter().enumerate() {
                    let at = CellRef::new(at.row + r as u32, at.col + c as u32);
                    book.set_formulas(si, RangeRef::cell(at), &[vec![v.clone()]], true)?;
                }
            }
            written.push((si, range));
        }
        Op::Clear { sheet, range } => {
            let si = sheet_index(book, sheet)?;
            book.clear(si, *range, false)?;
            written.push((si, *range));
        }
        Op::Insert {
            sheet,
            range,
            right,
        } => {
            let si = sheet_index(book, sheet)?;
            book.insert_cells(si, *range, *right, false)?;
            *moved = Some((si, edit_of(*range, *right, true)));
        }
        Op::Delete { sheet, range, left } => {
            let si = sheet_index(book, sheet)?;
            book.delete_cells(si, *range, *left)?;
            *moved = Some((si, edit_of(*range, *left, false)));
        }
        Op::Copy {
            sheet,
            from_sheet,
            from,
            to,
        } => {
            let (fi, ti) = (sheet_index(book, from_sheet)?, sheet_index(book, sheet)?);
            let target = copy(book, fi, *from, ti, *to)?;
            written.push((ti, target));
        }
        Op::PasteSpecial {
            sheet,
            from_sheet,
            from,
            to,
            paste: p,
        } => {
            let (fi, ti) = (sheet_index(book, from_sheet)?, sheet_index(book, sheet)?);
            let target = paste(book, fi, *from, ti, *to, p)?;
            written.push((ti, target));
        }
        Op::Move {
            sheet,
            from_sheet,
            from,
            to,
        } => {
            let (fi, ti) = (sheet_index(book, from_sheet)?, sheet_index(book, sheet)?);
            let target = sized(*to, from.height(), from.width())?;
            book.move_cells(fi, *from, ti, *to)?;
            written.push((fi, *from));
            written.push((ti, target));
        }
        Op::Format { sheet, range, set } => {
            let si = sheet_index(book, sheet)?;
            for (key, value) in set {
                format(book, si, *range, key, value)?;
            }
        }
        Op::ClearFormats { sheet, range } => {
            let si = sheet_index(book, sheet)?;
            book.clear_formats(si, *range)?;
        }
        Op::Size {
            sheet,
            range,
            columns,
            px,
        } => {
            let si = sheet_index(book, sheet)?;
            if *columns {
                let chars = crate::book::width_of_px(*px);
                book.set_format(si, *range, "columnwidth", &Val::Num(chars))?;
            } else {
                book.set_format(si, *range, "rowheight", &Val::Num(px * 0.75))?;
            }
        }
        Op::Hide {
            sheet,
            range,
            columns,
            hidden,
        } => {
            let si = sheet_index(book, sheet)?;
            let (lo, hi) = if *columns {
                (range.start.col, range.end.col)
            } else {
                (range.start.row, range.end.row)
            };
            book.set_hidden(si, *columns, lo, hi, *hidden)?;
        }
        Op::Merge {
            sheet,
            range,
            across,
            on,
        } => {
            let si = sheet_index(book, sheet)?;
            if !*on {
                book.set_format(si, *range, "mergecells", &Val::Bool(false))?;
                return Ok(());
            }
            limit(range.count())?;
            // A merged cell keeps what its first cell holds, as in Excel.
            let areas: Vec<RangeRef> = if *across {
                (range.start.row..=range.end.row)
                    .map(|r| {
                        RangeRef::new(
                            CellRef::new(r, range.start.col),
                            CellRef::new(r, range.end.col),
                        )
                    })
                    .collect()
            } else {
                vec![*range]
            };
            for area in areas {
                if area.width() > 1 {
                    let rest = RangeRef::new(
                        CellRef::new(area.start.row, area.start.col + 1),
                        CellRef::new(area.start.row, area.end.col),
                    );
                    book.clear(si, rest, false)?;
                }
                if area.height() > 1 {
                    let rest =
                        RangeRef::new(CellRef::new(area.start.row + 1, area.start.col), area.end);
                    book.clear(si, rest, false)?;
                }
            }
            book.merge_cells(si, *range, *across)?;
            written.push((si, *range));
        }
        Op::Sort {
            sheet,
            range,
            keys,
            header,
        } => {
            let si = sheet_index(book, sheet)?;
            let mut rows = around(book, si, *range);
            if keys
                .iter()
                .any(|(by, _)| !(rows.start.col..=rows.end.col).contains(by))
            {
                return Err("並べ替えの列が、並べ替える範囲の外です".into());
            }
            if header.unwrap_or_else(|| book.guess_header(si, rows)) {
                if rows.height() < 2 {
                    return Ok(());
                }
                rows.start.row += 1;
            }
            let keys: Vec<SortKey> = keys
                .iter()
                .map(|(index, descending)| SortKey {
                    index: *index,
                    descending: *descending,
                })
                .collect();
            book.sort(si, rows, &keys, false, false)?;
            written.push((si, rows));
        }
        Op::Filter { sheet, range } => {
            let si = sheet_index(book, sheet)?;
            let range = range.map(|r| {
                if r.is_single() {
                    book.current_region(si, r.start)
                } else {
                    r
                }
            });
            book.set_auto_filter(si, range)?;
        }
        Op::FilterBy {
            sheet,
            column,
            show,
        } => {
            let si = sheet_index(book, sheet)?;
            let Some((range, _)) = book.auto_filter(si) else {
                return Err("オートフィルターが設定されていません。".into());
            };
            let Some(offset) = column
                .checked_sub(range.start.col)
                .filter(|o| *o < range.width())
            else {
                return Err("フィルターの範囲の外の列です".into());
            };
            book.set_filter(si, offset, show.clone())?;
        }
        Op::Autofill { sheet, from, to } => {
            let si = sheet_index(book, sheet)?;
            autofill(book, si, *from, *to)?;
            written.push((si, *to));
        }
        // (Made on the file by `apply`: it changes no cell.)
        Op::Macro { .. } | Op::Freeze { .. } => {}
        Op::Fill {
            sheet,
            range,
            right,
        } => {
            let si = sheet_index(book, sheet)?;
            let (first, rest) = if *right {
                if range.width() < 2 {
                    return Ok(());
                }
                (
                    RangeRef::new(range.start, CellRef::new(range.end.row, range.start.col)),
                    RangeRef::new(
                        CellRef::new(range.start.row, range.start.col + 1),
                        range.end,
                    ),
                )
            } else {
                if range.height() < 2 {
                    return Ok(());
                }
                (
                    RangeRef::new(range.start, CellRef::new(range.start.row, range.end.col)),
                    RangeRef::new(
                        CellRef::new(range.start.row + 1, range.start.col),
                        range.end,
                    ),
                )
            };
            let target = copy(book, si, first, si, rest)?;
            written.push((si, target));
        }
    }
    Ok(())
}

fn apply_group(bytes: &[u8], ops: &[Op]) -> Result<Vec<u8>> {
    let mut styles = Styles::default();
    // The workbook as it was read, saved the way the edited one is: what differs between the
    // two is the edit, and nothing else of the file is written again.
    let before = Book::open_bytes(bytes)?.snapshot(&mut styles, &[]);
    let mut book = Book::open_bytes(bytes)?;
    let (mut written, mut moved) = (Vec::new(), None);
    for op in ops {
        run(&mut book, op, &mut written, &mut moved)?;
    }
    let after = book.snapshot(&mut styles, &written);
    let plain = styles.plain();
    let out = write(
        bytes,
        &before,
        &after,
        &styles,
        plain,
        moved.as_ref().map(|(si, e)| (*si, e)),
    )?;
    verify(&out, &after)?;
    Ok(out)
}

/// Reads the new file back: its cells are those of the edited workbook, or nothing is written.
fn verify(bytes: &[u8], want: &Snapshot) -> Result<()> {
    let mut styles = Styles::default();
    let got = Book::open_bytes(bytes)
        .map_err(|e| {
            format!("書き換えたブックを読み戻せませんでした（{e}）。ファイルは変更していません")
        })?
        .snapshot(&mut styles, &[]);
    let blank = CellState {
        formula: None,
        value: Raw::Empty,
        style: 0,
    };
    let mut wrong = Vec::new();
    for (si, sheet) in want.sheets.iter().enumerate() {
        let other = got.sheets.get(si);
        let keys: HashSet<&CellRef> = sheet
            .cells
            .keys()
            .chain(other.iter().flat_map(|o| o.cells.keys()))
            .collect();
        for at in keys {
            let a = sheet.cells.get(at).unwrap_or(&blank);
            let b = other.and_then(|o| o.cells.get(at)).unwrap_or(&blank);
            // A formula is worked out again when the file is read: only that it is one.
            let same = match (&a.formula, &b.formula) {
                (Some(_), Some(_)) => true,
                (None, None) => a.value == b.value,
                _ => false,
            };
            if !same {
                wrong.push(format!("{}!{}", sheet.name, at.a1()));
            }
        }
    }
    if wrong.is_empty() {
        return Ok(());
    }
    wrong.sort();
    wrong.truncate(5);
    Err(format!(
        "書き換えたブックのセルが編集の結果と一致しませんでした（{}）。ファイルは変更していません",
        wrong.join(", ")
    ))
}

// ---------------------------------------------------------------------------- the package

/// The parts of the file.
pub(crate) struct Pack<'a> {
    zip: zip::ZipArchive<Cursor<&'a [u8]>>,
    names: Vec<String>,
}

impl<'a> Pack<'a> {
    fn open(bytes: &'a [u8]) -> Result<Pack<'a>> {
        let zip = zip::ZipArchive::new(Cursor::new(bytes))
            .map_err(|e| format!("ブックを開けません（{e}）"))?;
        let names = (0..zip.len())
            .map(|i| zip.name_for_index(i).unwrap_or("").to_string())
            .collect();
        Ok(Pack { zip, names })
    }

    /// The entry of a part (part names do not differ by case).
    pub(crate) fn find(&self, part: &str) -> Option<String> {
        let part = part.trim_start_matches('/');
        self.names
            .iter()
            .find(|n| n.as_str() == part)
            .or_else(|| self.names.iter().find(|n| n.eq_ignore_ascii_case(part)))
            .cloned()
    }

    pub(crate) fn names(&self) -> &[String] {
        &self.names
    }

    /// A part as text. An error when it is not UTF-8 (its positions could not be kept).
    pub(crate) fn text(&mut self, part: &str) -> Result<Option<String>> {
        let Some(name) = self.find(part) else {
            return Ok(None);
        };
        let mut file = self
            .zip
            .by_name(&name)
            .map_err(|e| format!("{name}: {e}"))?;
        let mut data = Vec::with_capacity(file.size().min(1 << 28) as usize);
        file.read_to_end(&mut data)
            .map_err(|e| format!("{name}: {e}"))?;
        String::from_utf8(data)
            .map(Some)
            .map_err(|_| format!("{name} は UTF-8 ではないため、書き換えられません"))
    }

    /// The relationships of a part: (id, type, the part it points to).
    pub(crate) fn rels(&mut self, part: &str) -> Result<Vec<(String, String, String)>> {
        let (dir, file) = part.rsplit_once('/').unwrap_or(("", part));
        let Some(doc) = self.text(&format!("{dir}/_rels/{file}.rels"))? else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close() && t.local() == "Relationship") {
            let span = tag.start..tag.end;
            if scan::attr(&doc, span.clone(), "TargetMode").as_deref() == Some("External") {
                continue;
            }
            let (Some(id), Some(target)) = (
                scan::attr(&doc, span.clone(), "Id"),
                scan::attr(&doc, span.clone(), "Target"),
            ) else {
                continue;
            };
            let kind = scan::attr(&doc, span, "Type").unwrap_or_default();
            out.push((id, kind, resolve(dir, &target)));
        }
        Ok(out)
    }

    /// The file with the parts of `replace` given new contents and those of `remove` left out;
    /// the others are copied as they are stored.
    fn rewrite(
        &mut self,
        replace: &BTreeMap<String, Vec<u8>>,
        remove: &HashSet<String>,
    ) -> Result<Vec<u8>> {
        use zip::write::SimpleFileOptions;
        let failed = |e: zip::result::ZipError| format!("ブックを書き出せません（{e}）");
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        writer.set_raw_comment(self.zip.comment().to_vec().into_boxed_slice());
        for index in 0..self.zip.len() {
            let name = self.names[index].clone();
            if remove.contains(&name) {
                continue;
            }
            let Some(data) = replace.get(&name) else {
                let file = self.zip.by_index_raw(index).map_err(failed)?;
                writer.raw_copy_file(file).map_err(failed)?;
                continue;
            };
            let options = {
                let original = self.zip.by_index_raw(index).map_err(failed)?;
                let mut options =
                    SimpleFileOptions::default().compression_method(match original.compression() {
                        zip::CompressionMethod::Stored => zip::CompressionMethod::Stored,
                        _ => zip::CompressionMethod::Deflated,
                    });
                if let Some(time) = original.last_modified() {
                    options = options.last_modified_time(time);
                }
                options
            };
            writer.start_file(name.as_str(), options).map_err(failed)?;
            writer
                .write_all(data)
                .map_err(|e| format!("ブックを書き出せません（{e}）"))?;
        }
        Ok(writer.finish().map_err(failed)?.into_inner())
    }
}

/// The part a relationship of a part in `dir` points to.
fn resolve(dir: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut parts: Vec<&str> = dir.split('/').filter(|p| !p.is_empty()).collect();
    for p in target.split('/') {
        match p {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

const WORKBOOK: &str = "xl/workbook.xml";

/// The parts of the sheets, by the names on their tabs.
fn sheet_parts(pack: &mut Pack, workbook: &str) -> Result<HashMap<String, String>> {
    let rels = pack.rels(WORKBOOK)?;
    let mut out = HashMap::new();
    for tag in scan::tags(workbook, 0).filter(|t| !t.is_close() && t.local() == "sheet") {
        let span = tag.start..tag.end;
        let (Some(name), Some(id)) = (
            scan::attr(workbook, span.clone(), "name"),
            scan::attr(workbook, span, "id"),
        ) else {
            continue;
        };
        if let Some((_, _, part)) = rels.iter().find(|r| r.0 == id) {
            out.insert(name, part.clone());
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------- the cells of a sheet

/// A cell of the file.
struct SrcCell {
    at: CellRef,
    /// Its format (the index of its `xf`).
    s: u32,
    span: Range<usize>,
    open: Range<usize>,
    /// Its formula is shared with other cells: it is written out on its own when the sheet is.
    shared: bool,
    /// `cm`: the cell metadata of a dynamic-array formula.
    cm: Option<String>,
}

struct SrcRow {
    r: u32,
    open: Range<usize>,
    cells: Vec<SrcCell>,
}

#[derive(Default)]
struct SheetData {
    /// The namespace prefix of the sheet's elements (`x:`), usually none.
    prefix: String,
    rows: Vec<SrcRow>,
    /// The part has no cells that can be read (a chart sheet): it is left as it is.
    unread: bool,
}

fn read_sheet_data(doc: &str) -> Option<SheetData> {
    let mut tags = scan::tags(doc, 0);
    let head = tags.by_ref().find(|t| t.local() == "sheetData")?;
    let mut data = SheetData {
        prefix: scan::prefix(head.name).to_string(),
        ..Default::default()
    };
    if head.is_empty() {
        return Some(data);
    }
    let (mut next_row, mut next_col) = (0u32, 0u32);
    while let Some(tag) = tags.next() {
        if tag.is_close() {
            if tag.local() == "sheetData" {
                return Some(data);
            }
            continue;
        }
        let span = tag.start..tag.end;
        match tag.local() {
            "row" => {
                let r = scan::attr(doc, span.clone(), "r")
                    .and_then(|r| r.parse::<u32>().ok())
                    .filter(|r| *r > 0)
                    .map_or(next_row, |r| r - 1);
                next_row = r + 1;
                next_col = 0;
                data.rows.push(SrcRow {
                    r,
                    open: span,
                    cells: Vec::new(),
                });
            }
            "c" => {
                let row = data.rows.last_mut()?;
                let at = scan::attr(doc, span.clone(), "r")
                    .and_then(|r| CellRef::parse(&r))
                    .unwrap_or(CellRef::new(row.r, next_col));
                next_col = at.col + 1;
                let mut cell = SrcCell {
                    at,
                    s: scan::attr(doc, span.clone(), "s")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(0),
                    span: span.clone(),
                    open: span.clone(),
                    shared: false,
                    cm: scan::attr(doc, span, "cm"),
                };
                if tag.is_open() {
                    // To its end tag; what it holds does not nest cells.
                    loop {
                        let inner = tags.next()?;
                        if inner.is_close() && inner.name == tag.name {
                            cell.span.end = inner.end;
                            break;
                        }
                        if !inner.is_close()
                            && inner.local() == "f"
                            && scan::attr(doc, inner.start..inner.end, "t").as_deref()
                                == Some("shared")
                        {
                            cell.shared = true;
                        }
                    }
                }
                row.cells.push(cell);
            }
            _ => {}
        }
    }
    None
}

/// Where an insertion or a deletion puts a cell (None: it is deleted).
pub(crate) fn map_cell(p: CellRef, edit: &Edit) -> Option<CellRef> {
    let along = |v: u32, at: u32, count: u32, insert: bool| -> Option<u32> {
        if insert {
            Some(if v >= at { v + count } else { v })
        } else if v >= at + count {
            Some(v - count)
        } else if v >= at {
            None
        } else {
            Some(v)
        }
    };
    match edit {
        Edit::Insert { axis, at, count } | Edit::Delete { axis, at, count } => {
            let insert = matches!(edit, Edit::Insert { .. });
            match axis {
                Axis::Rows => along(p.row, *at, *count, insert).map(|r| CellRef::new(r, p.col)),
                Axis::Cols => along(p.col, *at, *count, insert).map(|c| CellRef::new(p.row, c)),
            }
        }
        Edit::InsertCells { axis, range } | Edit::DeleteCells { axis, range } => {
            let insert = matches!(edit, Edit::InsertCells { .. });
            match axis {
                Axis::Rows if (range.start.col..=range.end.col).contains(&p.col) => {
                    along(p.row, range.start.row, range.height(), insert)
                        .map(|r| CellRef::new(r, p.col))
                }
                Axis::Cols if (range.start.row..=range.end.row).contains(&p.row) => {
                    along(p.col, range.start.col, range.width(), insert)
                        .map(|c| CellRef::new(p.row, c))
                }
                _ => Some(p),
            }
        }
        _ => Some(p),
    }
    .filter(|q| q.row < MAX_ROWS && q.col < MAX_COLS)
}

/// The attributes of a start tag without those named, as they are written.
fn attrs_without(doc: &str, tag: Range<usize>, without: &[&str]) -> String {
    let mut out = String::new();
    for (name, value) in scan::attrs(doc, tag) {
        if !without.contains(&name) {
            out.push_str(&format!(" {name}=\"{}\"", &doc[value]));
        }
    }
    out
}

/// The attributes of a tag as they are written (` a="1" b="2"`), to change some of them.
pub(crate) struct Attrs(Vec<(String, String)>);

impl Attrs {
    pub(crate) fn parse(text: &str) -> Attrs {
        let tag = format!("<x{text}>");
        Attrs(
            scan::attrs(&tag, 0..tag.len())
                .into_iter()
                .map(|(name, value)| (name.to_string(), tag[value].to_string()))
                .collect(),
        )
    }

    /// Sets an attribute to `value` (as it is written: escaped).
    pub(crate) fn set(&mut self, name: &str, value: &str) {
        match self.0.iter_mut().find(|(n, _)| n == name) {
            Some((_, v)) => *v = value.to_string(),
            None => self.0.push((name.to_string(), value.to_string())),
        }
    }

    pub(crate) fn remove(&mut self, name: &str) {
        self.0.retain(|(n, _)| n != name);
    }

    pub(crate) fn text(&self) -> String {
        self.0
            .iter()
            .map(|(n, v)| format!(" {n}=\"{v}\""))
            .collect()
    }
}

pub(crate) fn number(n: f64) -> String {
    let s = format!("{n}");
    if s.len() > 24 { format!("{n:e}") } else { s }
}

/// A cell written anew.
fn cell_xml(p: &str, at: CellRef, s: u32, state: &CellState, cm: Option<&str>) -> String {
    let mut open = format!("<{p}c r=\"{}\"", at.a1());
    if s != 0 {
        open.push_str(&format!(" s=\"{s}\""));
    }
    let mut body = String::new();
    if let Some(f) = &state.formula {
        if f.area.is_some() || f.dynamic {
            let area = f.area.clone().unwrap_or_else(|| at.a1());
            body.push_str(&format!(
                "<{p}f t=\"array\" ref=\"{}\">{}</{p}f>",
                scan::escape(&area),
                scan::escape(&f.text)
            ));
            if let (true, Some(cm)) = (f.dynamic, cm) {
                open.push_str(&format!(" cm=\"{}\"", scan::escape(cm)));
            }
        } else {
            body.push_str(&format!("<{p}f>{}</{p}f>", scan::escape(&f.text)));
        }
    }
    let value = |v: &str| format!("<{p}v>{v}</{p}v>");
    match &state.value {
        Raw::Empty => {}
        Raw::Num(n) => body.push_str(&value(&number(*n))),
        Raw::Bool(b) => {
            open.push_str(" t=\"b\"");
            body.push_str(&value(if *b { "1" } else { "0" }));
        }
        Raw::Err(e) => {
            open.push_str(" t=\"e\"");
            body.push_str(&value(&scan::escape(e)));
        }
        Raw::Text(t) if state.formula.is_some() => {
            open.push_str(" t=\"str\"");
            body.push_str(&format!(
                "<{p}v xml:space=\"preserve\">{}</{p}v>",
                scan::escape(t)
            ));
        }
        Raw::Text(t) => {
            open.push_str(" t=\"inlineStr\"");
            body.push_str(&format!(
                "<{p}is><{p}t xml:space=\"preserve\">{}</{p}t></{p}is>",
                scan::escape(t)
            ));
        }
    }
    if body.is_empty() {
        format!("{open}/>")
    } else {
        format!("{open}>{body}</{p}c>")
    }
}

/// A cell of the file as it is, at `at`, with the format `s` when one is given.
fn verbatim(doc: &str, cell: &SrcCell, at: CellRef, s: Option<u32>) -> String {
    let text = &doc[cell.span.clone()];
    let s = s.filter(|s| *s != cell.s);
    if cell.at == at && s.is_none() {
        return text.to_string();
    }
    let open = 0..cell.open.len();
    let mut edits = Edits::default();
    if cell.at != at {
        scan::set_attr(text, open.clone(), "r", &at.a1(), &mut edits);
    }
    if let Some(s) = s {
        scan::set_attr(text, open, "s", &s.to_string(), &mut edits);
    }
    edits.apply(text)
}

/// A format to add to the file's.
pub(crate) enum NewXf {
    /// The `xf` with this index, with another number format.
    Derived(u32, String),
    /// A format of the workbook (an index into [`Styles`]) written out whole.
    Fresh(usize),
}

/// The formats of the file (`xf` indexes) for the formats of the cells.
struct Formats<'a> {
    styles: &'a Styles,
    known: HashMap<usize, u32>,
    new: Vec<NewXf>,
    /// The number of `xf` the file has.
    count: u32,
}

impl Formats<'_> {
    /// The `xf` for the format `style` of a cell that had the format `was` as the `xf` `s`.
    fn of(&mut self, style: usize, was: Option<(usize, u32)>) -> u32 {
        if let Some((old, s)) = was
            && old == style
        {
            return s;
        }
        if let Some(s) = self.known.get(&style) {
            return *s;
        }
        // The same format with another number format (a date typed into a cell) is made from
        // the format the file has, which may hold more than the workbook reads of it.
        let count = self.count;
        let base = was
            .filter(|(old, _)| self.styles.same_but_number_format(*old, style))
            .map(|(_, s)| s)
            .or_else(|| {
                let mut like: Vec<(usize, u32)> = self
                    .known
                    .iter()
                    .filter(|(k, _)| self.styles.same_but_number_format(**k, style))
                    .map(|(k, s)| (*k, *s))
                    .collect();
                like.sort();
                like.first().map(|(_, s)| *s)
            })
            .filter(|s| *s < count);
        let new = match base {
            Some(base) => NewXf::Derived(base, self.styles.number_format(style)),
            None => NewXf::Fresh(style),
        };
        let at = self
            .new
            .iter()
            .position(|n| match (n, &new) {
                (NewXf::Derived(a, x), NewXf::Derived(b, y)) => a == b && x == y,
                (NewXf::Fresh(a), NewXf::Fresh(b)) => a == b,
                _ => false,
            })
            .unwrap_or_else(|| {
                self.new.push(new);
                self.new.len() - 1
            });
        let s = self.count + at as u32;
        self.known.insert(style, s);
        s
    }
}

/// The sheet's XML with the cells of `after`; None when nothing of it changes.
#[allow(clippy::too_many_arguments)]
fn rewrite_sheet(
    doc: &str,
    data: &SheetData,
    si: usize,
    before: &Snapshot,
    after: &Snapshot,
    edit: Option<&Edit>,
    moved: Option<(&str, &Edit)>,
    formats: &mut Formats,
    plain: usize,
) -> Result<Option<String>> {
    let (b, a) = (&before.sheets[si], &after.sheets[si]);
    // What whole rows and columns are given, and the filter: where no cells moved (rows and
    // columns inserted take theirs from their neighbours, in the XML).
    let still = edit.is_none();
    let rows_changed = still && a.rows != b.rows;
    let cols_changed = still && a.cols != b.cols;
    let filter_changed = still && a.filter != b.filter;
    let defaults_changed =
        still && (a.default_width != b.default_width || a.default_height != b.default_height);
    let layout_changed = rows_changed || cols_changed || filter_changed || defaults_changed;
    let merges_changed = {
        let set = |m: &[RangeRef]| m.iter().map(RangeRef::a1).collect::<HashSet<String>>();
        set(&b.merges) != set(&a.merges)
    };
    let cells_changed = edit.is_some()
        || rows_changed
        || a.cells.len() != b.cells.len()
        || a.cells.iter().any(|(at, c)| b.cells.get(at) != Some(c));
    if !cells_changed && !merges_changed && !layout_changed && moved.is_none() {
        return Ok(None);
    }
    if data.unread {
        return if cells_changed || merges_changed || layout_changed {
            Err(format!(
                "シート `{}` のセルを読み取れないため、書き換えられません",
                a.name
            ))
        } else {
            Ok(None)
        };
    }
    let root = scan::tree(doc, &["sheetData"]).ok_or("シートの XML を読み取れません")?;
    let mut edits = Edits::default();
    let p = data.prefix.as_str();

    let mut used: Option<RangeRef> = None;
    if cells_changed {
        let map = |at: CellRef| match edit {
            Some(e) => map_cell(at, e),
            None => Some(at),
        };
        // Where the cells of the file are after the edit, and what the cells were.
        let mut sources: HashMap<CellRef, &SrcCell> = HashMap::new();
        for cell in data.rows.iter().flat_map(|r| &r.cells) {
            if let Some(q) = map(cell.at) {
                sources.insert(q, cell);
            }
        }
        let mut old_at: HashMap<CellRef, CellRef> = HashMap::new();
        for at in b.cells.keys() {
            if let Some(q) = map(*at) {
                old_at.insert(q, *at);
            }
        }
        let mut rows: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
        // (In the order of the sheet: the formats added to the file are then the same every time.)
        let mut cells: Vec<(&CellRef, &CellState)> = a.cells.iter().collect();
        cells.sort_by_key(|(q, _)| (q.row, q.col));
        for (q, state) in cells {
            let src = sources.get(q).copied();
            let old = old_at.get(q).and_then(|at| b.cells.get(at));
            let same = old == Some(state);
            let opaque = state.formula.as_ref().is_some_and(|f| f.opaque);
            // Only its format changed: the cell as it is written, with the other format.
            let restyled = old.filter(|o| o.formula == state.formula && o.value == state.value);
            let xml = match src {
                Some(cell) if opaque || (same && !cell.shared) => verbatim(doc, cell, *q, None),
                Some(cell) if restyled.is_some() && !cell.shared => {
                    let was = restyled.map(|o| (o.style, cell.s));
                    verbatim(doc, cell, *q, Some(formats.of(state.style, was)))
                }
                _ => {
                    // A cell of the file with something in it that the workbook did not read
                    // would be written over without it.
                    if let Some(cell) = src
                        && old.is_none()
                        && ["<v", "<f", "<is"].iter().any(|tag| {
                            doc[cell.span.clone()].contains(&format!("{tag}>"))
                                || doc[cell.span.clone()].contains(&format!("{tag} "))
                        })
                    {
                        return Err(format!(
                            "シート `{}` のセル {} を読み取れていないため、書き換えられません（Excel で保存し直すと編集できることがあります）",
                            a.name,
                            cell.at.a1()
                        ));
                    }
                    let was = match (old, src) {
                        (Some(old), Some(cell)) => Some((old.style, cell.s)),
                        _ => None,
                    };
                    let s = formats.of(state.style, was);
                    if s == 0 && state.formula.is_none() && state.value == Raw::Empty {
                        continue;
                    }
                    cell_xml(p, *q, s, state, src.and_then(|c| c.cm.as_deref()))
                }
            };
            rows.entry(q.row).or_default().insert(q.col, xml);
        }
        // Cells of the file the workbook does not hold (empty ones without a format).
        for (q, cell) in &sources {
            if !a.cells.contains_key(q) && !b.cells.contains_key(&cell.at) {
                rows.entry(q.row)
                    .or_default()
                    .insert(q.col, verbatim(doc, cell, *q, None));
            }
        }
        // The rows: their heights and formats go where their cells go.
        let mut row_attrs: BTreeMap<u32, String> = BTreeMap::new();
        for row in &data.rows {
            let to = match edit {
                Some(Edit::Insert {
                    axis: Axis::Rows, ..
                })
                | Some(Edit::Delete {
                    axis: Axis::Rows, ..
                }) => map_cell(CellRef::new(row.r, 0), edit.expect("matched")).map(|q| q.row),
                _ => Some(row.r),
            };
            if let Some(to) = to {
                let attrs = attrs_without(doc, row.open.clone(), &["r", "spans"]);
                if !attrs.is_empty() {
                    row_attrs.insert(to, attrs);
                }
            }
        }
        if let Some(Edit::Insert {
            axis: Axis::Rows,
            at,
            count,
        }) = edit
        {
            // The rows inserted are as high as the row above them, and formatted as it is.
            let from = if *at == 0 { 0 } else { at - 1 };
            if let Some(row) = data.rows.iter().find(|r| r.r == from) {
                let all = scan::attrs(doc, row.open.clone());
                let custom = all
                    .iter()
                    .any(|(n, v)| *n == "customHeight" && matches!(&doc[v.clone()], "1" | "true"));
                let mut attrs = String::new();
                for (name, value) in all {
                    let keep = match name {
                        "s" | "customFormat" => true,
                        "ht" | "customHeight" => custom,
                        _ => false,
                    };
                    if keep {
                        attrs.push_str(&format!(" {name}=\"{}\"", &doc[value]));
                    }
                }
                if !attrs.is_empty() {
                    for n in *at..at + count {
                        row_attrs.insert(n, attrs.clone());
                    }
                }
            }
        }
        if rows_changed {
            let rows: std::collections::BTreeSet<u32> =
                a.rows.keys().chain(b.rows.keys()).copied().collect();
            for r in rows {
                let (was, now) = (b.rows.get(&r), a.rows.get(&r));
                if was == now {
                    continue;
                }
                let mut attrs = Attrs::parse(row_attrs.get(&r).map_or("", String::as_str));
                let custom = now.is_some_and(|n| n.custom);
                if custom {
                    attrs.set("ht", &number(now.map_or(0.0, |n| n.height)));
                    attrs.set("customHeight", "1");
                } else if was.is_some_and(|w| w.custom) {
                    attrs.remove("ht");
                    attrs.remove("customHeight");
                }
                if now.is_some_and(|n| n.hidden) {
                    attrs.set("hidden", "1");
                } else {
                    attrs.remove("hidden");
                }
                let style = now.map_or(plain, |n| n.style);
                if style != was.map_or(plain, |w| w.style) {
                    if style == plain {
                        attrs.remove("s");
                        attrs.remove("customFormat");
                    } else {
                        attrs.set("s", &formats.of(style, None).to_string());
                        attrs.set("customFormat", "1");
                    }
                }
                let text = attrs.text();
                if text.is_empty() {
                    row_attrs.remove(&r);
                } else {
                    row_attrs.insert(r, text);
                }
            }
        }
        let numbers: std::collections::BTreeSet<u32> =
            rows.keys().chain(row_attrs.keys()).copied().collect();
        let mut out = String::new();
        for r in numbers {
            let attrs = row_attrs.get(&r).map(String::as_str).unwrap_or("");
            match rows.get(&r) {
                Some(cells) if !cells.is_empty() => {
                    out.push_str(&format!("<{p}row r=\"{}\"{attrs}>", r + 1));
                    for (c, xml) in cells {
                        out.push_str(xml);
                        let at = RangeRef::cell(CellRef::new(r, *c));
                        used = Some(used.map_or(at, |u| u.union(&at)));
                    }
                    out.push_str(&format!("</{p}row>"));
                }
                _ => out.push_str(&format!("<{p}row r=\"{}\"{attrs}/>", r + 1)),
            }
        }
        let sheet_data = root
            .child("sheetData")
            .ok_or("シートの XML に sheetData がありません")?;
        edits.replace(
            sheet_data.span(),
            if out.is_empty() {
                format!("<{p}sheetData/>")
            } else {
                format!("<{p}sheetData>{out}</{p}sheetData>")
            },
        );
        if let Some(dimension) = root.child("dimension") {
            let area = used.map_or_else(|| "A1".to_string(), |u| u.a1_cells());
            scan::set_attr(doc, dimension.open.clone(), "ref", &area, &mut edits);
        }
    }
    // (What is put in at the same place goes in the order asked for: the order of a sheet's
    // elements.)
    if filter_changed {
        structure::filter(doc, &root, p, a.filter.as_ref(), &mut edits);
    }
    if merges_changed {
        structure::merges(&root, p, &a.merges, &mut edits);
    }
    if cols_changed || defaults_changed {
        let columns: std::collections::BTreeSet<u32> =
            a.cols.keys().chain(b.cols.keys()).copied().collect();
        let mut changes = Vec::new();
        for c in columns {
            let (was, now) = (b.cols.get(&c), a.cols.get(&c));
            if was == now {
                continue;
            }
            let style = now.map_or(plain, |n| n.style);
            let xf = (style != plain).then(|| formats.of(style, None));
            changes.push((
                c,
                now.map(|n| {
                    // (Read and written back, a width is not to the last digit what it was.)
                    let kept = was.is_some_and(|w| w.sized && (w.width - n.width).abs() < 0.01);
                    let width = if kept {
                        structure::Width::Kept
                    } else if n.sized {
                        structure::Width::Own(n.width)
                    } else {
                        structure::Width::Other
                    };
                    (width, n.hidden, xf)
                }),
            ));
        }
        if defaults_changed {
            structure::defaults(doc, &root, p, a.default_width, a.default_height, &mut edits);
        }
        // (9.140625: the width of a column of a new workbook, in the file's units.)
        let other = if a.default_width > 0.0 {
            a.default_width
        } else {
            9.140625
        };
        structure::columns_set(doc, &root, p, &changes, other, &mut edits);
    }
    if let Some((target, moved)) = moved {
        structure::sheet(
            doc,
            &root,
            &a.name,
            edit.is_some(),
            target,
            moved,
            &mut edits,
        );
    }
    if edits.is_empty() {
        return Ok(None);
    }
    Ok(Some(edits.apply(doc)))
}

trait A1Cells {
    fn a1_cells(&self) -> String;
}

impl A1Cells for RangeRef {
    /// `A1:B2` also for whole rows and columns (the `ref` of a sheet's dimension).
    fn a1_cells(&self) -> String {
        if self.is_single() {
            self.start.a1()
        } else {
            format!("{}:{}", self.start.a1(), self.end.a1())
        }
    }
}

// ---------------------------------------------------------------------------- the file

fn write(
    bytes: &[u8],
    before: &Snapshot,
    after: &Snapshot,
    styles: &Styles,
    plain: usize,
    moved: Option<(usize, &Edit)>,
) -> Result<Vec<u8>> {
    if before.sheets.len() != after.sheets.len() {
        return Err("シートの数が変わる編集は行えません".into());
    }
    let mut pack = Pack::open(bytes)?;
    let workbook = pack
        .text(WORKBOOK)?
        .ok_or("ブックに xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let mut docs: Vec<(String, String, SheetData)> = Vec::new();
    for sheet in &after.sheets {
        let part = parts
            .get(&sheet.name)
            .and_then(|p| pack.find(p))
            .ok_or_else(|| format!("シート `{}` のパートが見つかりません", sheet.name))?;
        let doc = pack
            .text(&part)?
            .ok_or_else(|| format!("{part} を読めません"))?;
        let data = read_sheet_data(&doc).unwrap_or(SheetData {
            unread: true,
            ..Default::default()
        });
        docs.push((part, doc, data));
    }
    let style_part = pack.find("xl/styles.xml");
    let style_doc = match &style_part {
        Some(part) => pack.text(part)?,
        None => None,
    };
    // The formats of the file by the formats of the cells that have them.
    let mut formats = Formats {
        styles,
        // (Cells without a format have the first one.)
        known: HashMap::from([(plain, 0)]),
        new: Vec::new(),
        count: style_doc.as_deref().map_or(0, styles::xf_count),
    };
    for (si, (_, doc, data)) in docs.iter().enumerate() {
        let sheet = &before.sheets[si];
        for cell in data.rows.iter().flat_map(|r| &r.cells) {
            if let Some(state) = sheet.cells.get(&cell.at) {
                formats.known.entry(state.style).or_insert(cell.s);
            }
        }
        // The formats of whole rows and columns, for the cells typed into them.
        for row in &data.rows {
            if let (Some(state), Some(s)) = (
                sheet.rows.get(&row.r),
                scan::attr(doc, row.open.clone(), "s").and_then(|v| v.parse::<u32>().ok()),
            ) {
                formats.known.entry(state.style).or_insert(s);
            }
        }
        let end = doc.find("sheetData").unwrap_or(doc.len());
        for t in scan::tags(&doc[..end], 0).filter(|t| !t.is_close() && t.local() == "col") {
            let span = t.start..t.end;
            let (Some(n), Some(s)) = (
                scan::attr(doc, span.clone(), "min").and_then(|v| v.parse::<u32>().ok()),
                scan::attr(doc, span, "style").and_then(|v| v.parse::<u32>().ok()),
            ) else {
                continue;
            };
            if let Some(state) = n.checked_sub(1).and_then(|n| sheet.cols.get(&n)) {
                formats.known.entry(state.style).or_insert(s);
            }
        }
    }

    let mut replace: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut remove: HashSet<String> = HashSet::new();
    let target = moved.map(|(si, e)| (after.sheets[si].name.as_str(), e));
    let mut touched = false;
    let known = formats.known.clone();
    // Twice at most: a format to add that the file turns out to have (one added by an earlier
    // edit, which the workbook reads back as another) is used in its place the second time.
    let mut same: Vec<(usize, u32)> = Vec::new();
    let mut style_out = None;
    for pass in 0..2 {
        formats.known = known.clone();
        formats.known.extend(same.iter().copied());
        formats.new.clear();
        for (si, (part, doc, data)) in docs.iter().enumerate() {
            let edit = moved.filter(|(at, _)| *at == si).map(|(_, e)| e);
            if let Some(new) = rewrite_sheet(
                doc,
                data,
                si,
                before,
                after,
                edit,
                target,
                &mut formats,
                plain,
            )? {
                scan::tree(&new, &["sheetData"])
                    .ok_or_else(|| format!("{part} を書き換えられませんでした"))?;
                replace.insert(part.clone(), new.into_bytes());
                touched = true;
            }
        }
        if formats.new.is_empty() {
            break;
        }
        let (Some(part), Some(doc)) = (&style_part, &style_doc) else {
            return Err("ブックに xl/styles.xml がありません".into());
        };
        let (new, found) = styles::with_formats(doc, &formats.new, styles)?;
        if found.is_empty() || pass == 1 {
            style_out = Some((part.clone(), new));
            break;
        }
        same = found;
    }
    if let Some((part, doc)) = style_out {
        replace.insert(part, doc.into_bytes());
    }
    if let Some((si, edit)) = moved {
        let sheet = &after.sheets[si].name;
        let names: Vec<String> = after.sheets.iter().map(|s| s.name.clone()).collect();
        let sheets: Vec<(String, String)> = docs
            .iter()
            .enumerate()
            .map(|(i, d)| (names[i].clone(), d.0.clone()))
            .collect();
        structure::parts(&mut pack, &sheets, sheet, edit, &mut replace)?;
        touched = true;
    }
    if !touched && before.names == after.names {
        // Nothing to write: the file as it is.
        return Ok(bytes.to_vec());
    }
    structure::workbook(
        &mut pack,
        &workbook,
        before,
        after,
        moved,
        &mut replace,
        &mut remove,
    )?;
    pack.rewrite(&replace, &remove)
}
