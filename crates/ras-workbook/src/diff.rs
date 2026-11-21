//! `ras book diff`: what differs between two workbooks, cell by cell (value, formula, number
//! format) and in structure (sheets added, removed and renamed; rows and columns inserted and
//! deleted, as far as they can be told from the cells). `ras run` shows what a run changed in
//! the same form.

use std::collections::{BTreeMap, HashMap};

use serde_json::{Value as Json, json};

use crate::book::{Book, CellInfo};
use crate::dump::{literal, type_name};
use crate::{CellRef, RangeRef};

/// One side of a changed cell.
#[derive(Clone, Debug, PartialEq)]
pub struct Side {
    /// The value as `ras book dump` writes it.
    pub value: String,
    pub kind: &'static str,
    pub formula: Option<String>,
    pub number_format: String,
    r1c1: Option<String>,
}

impl Side {
    fn shown(&self) -> String {
        match &self.formula {
            Some(f) => format!("{f}（{}）", self.value),
            None => self.value.clone(),
        }
    }

    fn json(&self) -> Json {
        let mut o = json!({ "value": self.value, "type": self.kind, "numberFormat": self.number_format });
        if let Some(f) = &self.formula {
            o["formula"] = json!(f);
        }
        o
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellChange {
    pub sheet: String,
    /// The cell in the second workbook.
    pub at: CellRef,
    pub before: Option<Side>,
    pub after: Option<Side>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Structure {
    SheetAdded(String),
    SheetRemoved(String),
    SheetRenamed { from: String, to: String },
    /// Rows or columns inserted (`count` of them before the 0-based `at`) or deleted.
    Lines { sheet: String, columns: bool, inserted: bool, at: u32, count: u32 },
}

#[derive(Clone, Debug, PartialEq)]
pub struct NameChange {
    pub name: String,
    pub before: Option<String>,
    pub after: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Diff {
    pub structure: Vec<Structure>,
    pub cells: Vec<CellChange>,
    pub names: Vec<NameChange>,
}

impl Diff {
    pub fn is_empty(&self) -> bool {
        self.structure.is_empty() && self.cells.is_empty() && self.names.is_empty()
    }

    pub fn total(&self) -> usize {
        self.structure.len() + self.cells.len() + self.names.len()
    }
}

type Cells = BTreeMap<(u32, u32), Side>;

fn read_sheet(book: &mut Book, si: usize) -> Result<Cells, String> {
    let all = RangeRef::new(CellRef::new(0, 0), CellRef::new(crate::book::MAX_ROWS - 1, crate::book::MAX_COLS - 1));
    let cells: Vec<CellInfo> = book.cells(si, all)?;
    let mut out = Cells::new();
    for c in cells {
        let value = literal(book, &c);
        out.insert((c.at.row, c.at.col), Side { value, kind: type_name(&c.value), formula: c.formula, number_format: c.number_format, r1c1: c.r1c1 });
    }
    Ok(out)
}

/// Whether two cells are the same. `moved`: the cells were shifted by inserted or deleted
/// lines, where a formula is the same when its R1C1 form is (its A1 text follows the shift).
fn same(a: &Side, b: &Side, moved: bool) -> bool {
    if a.number_format != b.number_format || a.formula.is_some() != b.formula.is_some() {
        return false;
    }
    match (&a.formula, moved) {
        (Some(_), true) => a.r1c1 == b.r1c1 || a.formula == b.formula,
        (Some(_), false) => a.formula == b.formula && a.value == b.value,
        (None, _) => a.value == b.value && a.kind == b.kind,
    }
}

/// What a line (row or column) holds, formulas left out: they change when lines move.
fn signatures(cells: &Cells, columns: bool) -> Vec<u64> {
    use std::hash::{Hash, Hasher};
    let mut lines: BTreeMap<u32, std::collections::hash_map::DefaultHasher> = BTreeMap::new();
    let mut last = 0;
    for ((row, col), c) in cells {
        let (line, pos) = if columns { (*col, *row) } else { (*row, *col) };
        let h = lines.entry(line).or_default();
        pos.hash(h);
        match &c.formula {
            Some(_) => "=".hash(h),
            None => c.value.hash(h),
        }
        last = last.max(line);
    }
    if cells.is_empty() {
        return Vec::new();
    }
    (0..=last).map(|i| lines.get(&i).map(|h| h.finish()).unwrap_or(0)).collect()
}

/// One block of lines inserted or deleted between `a` and `b`: (inserted, at, count).
fn one_block(a: &[u64], b: &[u64]) -> Option<(bool, u32, u32)> {
    if a.len() == b.len() {
        return None;
    }
    let (short, long, inserted) = if a.len() < b.len() { (a, b, true) } else { (b, a, false) };
    let count = long.len() - short.len();
    let prefix = short.iter().zip(long).take_while(|(x, y)| x == y).count();
    let suffix = short[prefix..].iter().rev().zip(long.iter().rev()).take_while(|(x, y)| x == y).count();
    // Lines added after the last one are new cells, not an insertion: something has to follow
    // the block. Lines that differ besides the block are told as changed cells; the block is
    // taken when most of the shorter side matches around it.
    if suffix == 0 || (prefix + suffix) * 2 < short.len() {
        return None;
    }
    Some((inserted, prefix as u32, count as u32))
}

fn shift(cells: &Cells, columns: bool, at: u32, by: i64) -> Cells {
    let mut out = Cells::new();
    for ((row, col), c) in cells {
        let line = if columns { *col } else { *row };
        if by < 0 && line >= at && (line as i64) < at as i64 - by {
            continue;
        }
        let moved = if line >= at { (line as i64 + by) as u32 } else { line };
        out.insert(if columns { (*row, moved) } else { (moved, *col) }, c.clone());
    }
    out
}

fn count_changes(a: &Cells, b: &Cells, moved: bool) -> usize {
    let mut n = 0;
    for (k, x) in a {
        match b.get(k) {
            Some(y) if same(x, y, moved) => {}
            _ => n += 1,
        }
    }
    n + b.keys().filter(|k| !a.contains_key(k)).count()
}

fn sheet_diff(name: &str, a: &Cells, b: &Cells, out: &mut Diff) {
    let (mut a, mut moved) = (a.clone(), false);
    for columns in [false, true] {
        let direct = count_changes(&a, b, moved);
        if direct == 0 {
            break;
        }
        let Some((inserted, at, count)) = one_block(&signatures(&a, columns), &signatures(b, columns)) else { continue };
        // The cells of the first workbook, moved to where they are in the second one.
        let candidate = shift(&a, columns, at, if inserted { count as i64 } else { -(count as i64) });
        let removed: usize = if inserted { 0 } else { a.len() - candidate.len() };
        let added_in_block = if inserted { b.keys().filter(|(r, c)| (at..at + count).contains(if columns { c } else { r })).count() } else { 0 };
        let aligned = count_changes(&candidate, b, true).saturating_sub(added_in_block);
        if aligned * 2 < direct.saturating_sub(removed).max(1) {
            out.structure.push(Structure::Lines { sheet: name.to_string(), columns, inserted, at, count });
            // The cells of inserted lines are told as new cells.
            a = candidate;
            moved = true;
        }
    }
    let mut keys: Vec<(u32, u32)> = a.keys().chain(b.keys()).copied().collect();
    keys.sort();
    keys.dedup();
    for k in keys {
        let (x, y) = (a.get(&k), b.get(&k));
        if let (Some(x), Some(y)) = (x, y) {
            if same(x, y, moved) {
                continue;
            }
        }
        out.cells.push(CellChange { sheet: name.to_string(), at: CellRef::new(k.0, k.1), before: x.cloned(), after: y.cloned() });
    }
}

/// Compares two workbooks.
pub fn diff(a: &mut Book, b: &mut Book) -> Result<Diff, String> {
    let mut out = Diff::default();
    let names = |book: &Book| (0..book.sheet_count()).map(|i| book.sheet_name(i).unwrap_or_default().to_string()).collect::<Vec<_>>();
    let (an, bn) = (names(a), names(b));
    let mut cells_a: HashMap<String, Cells> = HashMap::new();
    for (i, n) in an.iter().enumerate() {
        cells_a.insert(n.clone(), read_sheet(a, i)?);
    }
    let mut cells_b: HashMap<String, Cells> = HashMap::new();
    for (i, n) in bn.iter().enumerate() {
        cells_b.insert(n.clone(), read_sheet(b, i)?);
    }
    let mut removed: Vec<&String> = an.iter().filter(|n| !bn.contains(n)).collect();
    let mut added: Vec<&String> = bn.iter().filter(|n| !an.contains(n)).collect();
    // A sheet that is gone and a new one with the same cells (or, when there is one of each,
    // at the same place) are one sheet with a new name.
    let mut pairs: Vec<(String, String)> = an.iter().filter(|n| bn.contains(n)).map(|n| (n.clone(), n.clone())).collect();
    let mut renamed = Vec::new();
    for r in removed.clone() {
        let hit = added.iter().position(|x| count_changes(&cells_a[r], &cells_b[*x], false) == 0).or_else(|| {
            (removed.len() == 1 && added.len() == 1 && an.iter().position(|n| n == r) == bn.iter().position(|n| n == added[0])).then_some(0)
        });
        if let Some(i) = hit {
            let to = added.remove(i);
            removed.retain(|x| *x != r);
            renamed.push((r.clone(), to.clone()));
        }
    }
    for n in &removed {
        out.structure.push(Structure::SheetRemoved((*n).clone()));
    }
    for n in &added {
        out.structure.push(Structure::SheetAdded((*n).clone()));
    }
    for (from, to) in &renamed {
        out.structure.push(Structure::SheetRenamed { from: from.clone(), to: to.clone() });
        pairs.push((from.clone(), to.clone()));
    }
    // In the order of the sheets of the second workbook.
    pairs.sort_by_key(|(_, to)| bn.iter().position(|n| n == to));
    for (from, to) in &pairs {
        sheet_diff(to, &cells_a[from], &cells_b[to], &mut out);
    }
    for n in &added {
        sheet_diff(n, &Cells::new(), &cells_b[*n], &mut out);
    }
    let scoped = |book: &Book| {
        book.names().into_iter().map(|(name, scope, f)| (scope.and_then(|i| book.sheet_name(i)).map(|s| format!("{s}!{name}")).unwrap_or(name), format!("={f}"))).collect::<BTreeMap<String, String>>()
    };
    let (na, nb) = (scoped(a), scoped(b));
    let mut keys: Vec<&String> = na.keys().chain(nb.keys()).collect();
    keys.sort();
    keys.dedup();
    for k in keys {
        if na.get(k) != nb.get(k) {
            out.names.push(NameChange { name: k.clone(), before: na.get(k).cloned(), after: nb.get(k).cloned() });
        }
    }
    Ok(out)
}

fn col_name(col: u32) -> String {
    CellRef::new(0, col).a1().trim_end_matches('1').to_string()
}

fn lines_text(columns: bool, inserted: bool, at: u32, count: u32) -> String {
    let span = match (columns, count) {
        (false, 1) => format!("行 {}", at + 1),
        (false, _) => format!("行 {}〜{}", at + 1, at + count),
        (true, 1) => format!("列 {}", col_name(at)),
        (true, _) => format!("列 {}〜{}", col_name(at), col_name(at + count - 1)),
    };
    format!("{span} を{}（推定）", if inserted { "挿入" } else { "削除" })
}

/// The changes as text; `max_cells`: the changed cells written at most.
pub fn text(d: &Diff, max_cells: usize) -> String {
    if d.is_empty() {
        return "差分はありません\n".into();
    }
    let mut out = String::new();
    for s in &d.structure {
        out.push_str(&match s {
            Structure::SheetAdded(n) => format!("+ シート「{n}」を追加\n"),
            Structure::SheetRemoved(n) => format!("- シート「{n}」を削除\n"),
            Structure::SheetRenamed { from, to } => format!("~ シート「{from}」の名前を「{to}」に変更\n"),
            Structure::Lines { sheet, columns, inserted, at, count } => format!("~ {sheet}: {}\n", lines_text(*columns, *inserted, *at, *count)),
        });
    }
    let mut sheet = "";
    for c in d.cells.iter().take(max_cells) {
        if c.sheet != sheet {
            sheet = &c.sheet;
            out.push_str(&format!("{sheet}:\n"));
        }
        let side = |s: &Option<Side>| s.as_ref().map(Side::shown).unwrap_or_else(|| "（空）".into());
        let mut line = format!("  {}\t{} → {}", c.at.a1(), side(&c.before), side(&c.after));
        if let (Some(x), Some(y)) = (&c.before, &c.after) {
            if x.number_format != y.number_format {
                line.push_str(&format!("\t表示形式 {} → {}", x.number_format, y.number_format));
            }
        } else if let Some(y) = c.after.as_ref().filter(|y| y.number_format != "General") {
            line.push_str(&format!("\t表示形式 {}", y.number_format));
        }
        out.push_str(&line);
        out.push('\n');
    }
    if d.cells.len() > max_cells {
        out.push_str(&format!("  …ほか {} 個のセル\n", d.cells.len() - max_cells));
    }
    for n in &d.names {
        out.push_str(&match (&n.before, &n.after) {
            (None, Some(a)) => format!("+ 名前 {} {a}\n", n.name),
            (Some(b), None) => format!("- 名前 {} {b}\n", n.name),
            (b, a) => format!("~ 名前 {} {} → {}\n", n.name, b.clone().unwrap_or_default(), a.clone().unwrap_or_default()),
        });
    }
    out.push_str(&format!("差分 {} 件（構造 {}、セル {}、名前 {}）\n", d.total(), d.structure.len(), d.cells.len(), d.names.len()));
    out
}

pub fn to_json(d: &Diff) -> Json {
    let structure: Vec<Json> = d
        .structure
        .iter()
        .map(|s| match s {
            Structure::SheetAdded(n) => json!({ "kind": "sheetAdded", "sheet": n }),
            Structure::SheetRemoved(n) => json!({ "kind": "sheetRemoved", "sheet": n }),
            Structure::SheetRenamed { from, to } => json!({ "kind": "sheetRenamed", "from": from, "sheet": to }),
            Structure::Lines { sheet, columns, inserted, at, count } => json!({
                "kind": match (columns, inserted) { (false, true) => "rowsInserted", (false, false) => "rowsDeleted", (true, true) => "columnsInserted", (true, false) => "columnsDeleted" },
                "sheet": sheet, "at": at + 1, "count": count, "estimated": true,
            }),
        })
        .collect();
    let cells: Vec<Json> =
        d.cells.iter().map(|c| json!({ "sheet": c.sheet, "address": c.at.a1(), "before": c.before.as_ref().map(Side::json), "after": c.after.as_ref().map(Side::json) })).collect();
    let names: Vec<Json> = d.names.iter().map(|n| json!({ "name": n.name, "before": n.before, "after": n.after })).collect();
    json!({
        "summary": { "structure": d.structure.len(), "cells": d.cells.len(), "names": d.names.len(), "total": d.total() },
        "structure": structure,
        "cells": cells,
        "names": names,
    })
}
