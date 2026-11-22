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
use crate::book::{Book, MAX_COLS, MAX_ROWS, Val};

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
}

impl Op {
    fn structural(&self) -> bool {
        matches!(self, Op::Insert { .. } | Op::Delete { .. })
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
            other => return Err(format!("編集 `{other}` はありません（set / values / clear / insert / delete / copy / move / fill）")),
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
        bytes = apply_group(&bytes, &ops[at..end])?;
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
    let (h, w) = (from.height(), from.width());
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
            book.copy(fi, from, ti, at)?;
        }
    }
    Ok(target)
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
    let out = write(
        bytes,
        &before,
        &after,
        &styles,
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

fn number(n: f64) -> String {
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

/// A cell of the file as it is, at `at`.
fn verbatim(doc: &str, cell: &SrcCell, at: CellRef) -> String {
    let text = &doc[cell.span.clone()];
    if cell.at == at {
        return text.to_string();
    }
    let mut edits = Edits::default();
    scan::set_attr(text, 0..cell.open.len(), "r", &at.a1(), &mut edits);
    edits.apply(text)
}

/// The formats of the file (`xf` indexes) for the formats of the cells.
struct Formats<'a> {
    styles: &'a Styles,
    known: HashMap<usize, u32>,
    /// The formats to add: the `xf` they are made from and their number format.
    new: Vec<(u32, String)>,
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
        // The same format with another number format (a date typed into a cell): made from it.
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
            });
        let s = match base {
            Some(base) => {
                let code = self.styles.number_format(style);
                let at = match self.new.iter().position(|n| n.0 == base && n.1 == code) {
                    Some(i) => i,
                    None => {
                        self.new.push((base, code));
                        self.new.len() - 1
                    }
                };
                self.count + at as u32
            }
            // A format the file has nowhere: the cell keeps the one it had.
            None => was.map_or(0, |(_, s)| s),
        };
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
) -> Result<Option<String>> {
    let (b, a) = (&before.sheets[si], &after.sheets[si]);
    let merges_changed = {
        let set = |m: &[RangeRef]| m.iter().map(RangeRef::a1).collect::<HashSet<String>>();
        set(&b.merges) != set(&a.merges)
    };
    let cells_changed = edit.is_some()
        || a.cells.len() != b.cells.len()
        || a.cells.iter().any(|(at, c)| b.cells.get(at) != Some(c));
    if !cells_changed && !merges_changed && moved.is_none() {
        return Ok(None);
    }
    if data.unread {
        return if cells_changed || merges_changed {
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
        for (q, state) in &a.cells {
            let src = sources.get(q).copied();
            let old = old_at.get(q).and_then(|at| b.cells.get(at));
            let same = old == Some(state);
            let opaque = state.formula.as_ref().is_some_and(|f| f.opaque);
            let xml = match src {
                Some(cell) if opaque || (same && !cell.shared) => verbatim(doc, cell, *q),
                _ => {
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
                    .insert(q.col, verbatim(doc, cell, *q));
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
    if merges_changed {
        structure::merges(&root, p, &a.merges, &mut edits);
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
        known: HashMap::new(),
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
        for (styles_of, tag, key, attr) in [
            (&sheet.row_styles, "row", "r", "s"),
            (&sheet.col_styles, "col", "min", "style"),
        ] {
            if styles_of.is_empty() {
                continue;
            }
            // The formats of whole rows and columns, for the cells typed into them.
            let end = if tag == "col" {
                doc.find("sheetData").unwrap_or(doc.len())
            } else {
                doc.len()
            };
            for t in scan::tags(&doc[..end], 0).filter(|t| !t.is_close() && t.local() == tag) {
                let span = t.start..t.end;
                let (Some(n), Some(s)) = (
                    scan::attr(doc, span.clone(), key).and_then(|v| v.parse::<u32>().ok()),
                    scan::attr(doc, span, attr).and_then(|v| v.parse::<u32>().ok()),
                ) else {
                    continue;
                };
                if let Some(style) = n.checked_sub(1).and_then(|n| styles_of.get(&n)) {
                    formats.known.entry(*style).or_insert(s);
                }
            }
        }
    }

    let mut replace: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut remove: HashSet<String> = HashSet::new();
    let target = moved.map(|(si, e)| (after.sheets[si].name.as_str(), e));
    let mut touched = false;
    for (si, (part, doc, data)) in docs.iter().enumerate() {
        let edit = moved.filter(|(at, _)| *at == si).map(|(_, e)| e);
        if let Some(new) = rewrite_sheet(doc, data, si, before, after, edit, target, &mut formats)?
        {
            scan::tree(&new, &["sheetData"])
                .ok_or_else(|| format!("{part} を書き換えられませんでした"))?;
            replace.insert(part.clone(), new.into_bytes());
            touched = true;
        }
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
    if !formats.new.is_empty() {
        let (Some(part), Some(doc)) = (&style_part, &style_doc) else {
            return Err("ブックに xl/styles.xml がありません".into());
        };
        replace.insert(
            part.clone(),
            styles::with_formats(doc, &formats.new)?.into_bytes(),
        );
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
