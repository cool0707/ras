//! Sorting and filtering: `Range.Sort`, `Range.RemoveDuplicates`, `Range.AutoFilter`,
//! `Range.AdvancedFilter`, hidden rows and columns, and `Range.SpecialCells`.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use gridcraft_core::{CellRef, RangeRef, Value, compare_numbers, compare_text, sort_compare};
use gridcraft_model as gm;
use umya_spreadsheet::Style;

use super::paste::{Paste, PasteOp, PasteWhat};
use super::{Book, Result, coord};

/// One key of a sort.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SortKey {
    /// The column of the key (the row, when sorting columns), on the sheet.
    pub index: u32,
    pub descending: bool,
}

/// What `Range.SpecialCells` looks for (`XlCellType`).
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Special {
    /// Cells with a value that is not a formula; the mask is `XlSpecialCellsValue`.
    Constants(u32),
    /// Cells with a formula, by what they show.
    Formulas(u32),
    Blanks,
    LastCell,
    /// Cells of rows and columns that are not hidden.
    Visible,
}

/// The bits of `XlSpecialCellsValue` for a value.
fn kind_bit(v: &Value) -> u32 {
    match v {
        Value::Number(_) => 1,
        Value::Text(_) => 2,
        Value::Bool(_) => 4,
        Value::Error(_) => 16,
        _ => 0,
    }
}

/// Cells (in row order) as the fewest rectangles of whole runs.
fn areas(cells: impl IntoIterator<Item = CellRef>) -> Vec<RangeRef> {
    let mut runs: Vec<RangeRef> = Vec::new();
    for at in cells {
        match runs.last_mut() {
            Some(r) if r.end.row == at.row && r.end.col + 1 == at.col => r.end.col = at.col,
            _ => runs.push(RangeRef::cell(at)),
        }
    }
    let mut out: Vec<RangeRef> = Vec::new();
    let mut open: HashMap<(u32, u32), usize> = HashMap::new();
    for run in runs {
        let key = (run.start.col, run.end.col);
        match open.get(&key) {
            Some(&i) if out[i].end.row + 1 == run.start.row => out[i].end.row = run.start.row,
            _ => {
                open.insert(key, out.len());
                out.push(run);
            }
        }
    }
    out
}

/// `*` and `?` of a filter (with `~` before one meant as itself), without regard to case.
fn wildcard(pattern: &str, text: &str) -> bool {
    fn go(p: &[char], t: &[char]) -> bool {
        match p.first() {
            None => t.is_empty(),
            Some('*') => (0..=t.len()).any(|i| go(&p[1..], &t[i..])),
            Some('?') => !t.is_empty() && go(&p[1..], &t[1..]),
            Some('~') if p.len() > 1 && matches!(p[1], '*' | '?' | '~') => {
                t.first() == Some(&p[1]) && go(&p[2..], &t[1..])
            }
            Some(c) => t.first() == Some(c) && go(&p[1..], &t[1..]),
        }
    }
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    let t: Vec<char> = text.to_lowercase().chars().collect();
    go(&p, &t)
}

/// A comparison of a filter, split into its operator and what it compares with (`">=10"`).
pub fn split_criterion(text: &str) -> (&str, &str) {
    for op in ["<>", ">=", "<=", "=", ">", "<"] {
        if let Some(rest) = text.strip_prefix(op) {
            return (op, rest);
        }
    }
    ("", text)
}

/// A cell as filters see it.
struct Shown {
    value: Value,
    text: String,
}

impl Book {
    fn shown(&mut self, si: usize, at: CellRef) -> Shown {
        let value = self
            .model
            .sheet(si)
            .map(|s| s.value(at))
            .unwrap_or_default();
        let text = self.text(si, at, &value);
        Shown { value, text }
    }

    /// The number `text` is when typed into a cell (`10`, `2026/1/31`, `50%`).
    fn typed_number(&self, text: &str) -> Option<f64> {
        match gridcraft_core::parse::parse_input(text.trim(), self.model.date_system).value {
            Value::Number(n) => Some(n),
            _ => None,
        }
    }

    /// Whether a cell passes one comparison of a filter. `prefix`: text without an operator is
    /// what the cell begins with (`Range.AdvancedFilter`), not all of it.
    fn passes(&self, cell: &Shown, op: &str, val: &str, prefix: bool) -> bool {
        let number = self.typed_number(val);
        match op {
            "" | "=" => {
                if val.is_empty() {
                    return op.is_empty() || cell.text.is_empty();
                }
                if let (Some(n), Value::Number(x)) = (number, &cell.value)
                    && compare_numbers(*x, n) == Ordering::Equal
                {
                    return true;
                }
                if op.is_empty() && prefix && number.is_none() {
                    wildcard(&format!("{val}*"), &cell.text)
                } else {
                    wildcard(val, &cell.text)
                }
            }
            "<>" => {
                if val.is_empty() {
                    return !cell.text.is_empty();
                }
                !self.passes(cell, "=", val, false)
            }
            _ => {
                let order = match (&cell.value, number) {
                    (Value::Number(x), Some(n)) => compare_numbers(*x, n),
                    (Value::Text(t), None) => compare_text(t, val),
                    _ => return false,
                };
                match op {
                    ">" => order == Ordering::Greater,
                    ">=" => order != Ordering::Less,
                    "<" => order == Ordering::Less,
                    _ => order != Ordering::Greater,
                }
            }
        }
    }

    // ---- hidden rows and columns

    fn hide(&mut self, si: usize, columns: bool, index: u32, hidden: bool) -> bool {
        let Some(sh) = self.model.sheet_mut(si) else {
            return false;
        };
        let lines = if columns { &mut sh.cols } else { &mut sh.rows };
        let was = lines.get(&index).is_some_and(|l| l.hidden);
        if was == hidden {
            return false;
        }
        if hidden {
            lines.entry(index).or_default().hidden = true;
        } else if let Some(l) = lines.get_mut(&index) {
            l.hidden = false;
            if *l == gm::LineInfo::default() {
                lines.remove(&index);
            }
        }
        if let Ok(ws) = self.file.sheet_mut(si) {
            if columns {
                ws.column_dimension_by_number_mut(index + 1)
                    .set_hidden(hidden);
            } else {
                ws.row_dimension_mut(index + 1).set_hidden(hidden);
            }
        }
        true
    }

    /// `Rows.Hidden = hidden` (`columns`: `Columns.Hidden`) for the lines `lo..=hi`. Hiding goes
    /// no further than the used range: there is nothing to hide below it.
    pub fn set_hidden(
        &mut self,
        si: usize,
        columns: bool,
        lo: u32,
        hi: u32,
        hidden: bool,
    ) -> Result<()> {
        self.check_sheet(si)?;
        let lines: Vec<u32> = if hidden {
            let used = self.used_range(si);
            let last = used
                .map_or(lo, |u| if columns { u.end.col } else { u.end.row })
                .max(lo);
            (lo..=hi.min(last)).collect()
        } else {
            let sh = self.model.sheet(si).expect("checked");
            let lines = if columns { &sh.cols } else { &sh.rows };
            lines
                .range(lo..=hi)
                .filter(|(_, l)| l.hidden)
                .map(|(i, _)| *i)
                .collect()
        };
        let mut changed = false;
        for i in lines {
            changed |= self.hide(si, columns, i, hidden);
        }
        if changed {
            // SUBTOTAL and AGGREGATE leave hidden rows out.
            self.rebuild = true;
        }
        Ok(())
    }

    /// The row `row` (0-based) fit to its fonts: `pt` high, as Excel keeps a row it fits (no
    /// customHeight). A row given a height, or hidden, is left as it is.
    pub fn fit_row(&mut self, si: usize, row: u32, pt: f64) -> Result<()> {
        self.check_sheet(si)?;
        let ws = self.file.sheet_mut(si).map_err(|e| e.to_string())?;
        if ws
            .row_dimension(row + 1)
            .is_some_and(|r| r.custom_height() || r.hidden())
        {
            return Ok(());
        }
        ws.row_dimension_mut(row + 1)
            .set_height(pt)
            .set_custom_height(false);
        Ok(())
    }

    /// Whether the lines `lo..=hi` are hidden: None when some are and some are not.
    pub fn hidden(&self, si: usize, columns: bool, lo: u32, hi: u32) -> Option<bool> {
        let sh = self.model.sheet(si)?;
        let lines = if columns { &sh.cols } else { &sh.rows };
        let n = lines.range(lo..=hi).filter(|(_, l)| l.hidden).count() as u64;
        if n == 0 {
            Some(false)
        } else if n == (hi - lo) as u64 + 1 {
            Some(true)
        } else {
            None
        }
    }

    /// The cells of `r` in rows and columns that are not hidden, as rectangles.
    pub fn visible_areas(&self, si: usize, r: RangeRef) -> Vec<RangeRef> {
        let Some(sh) = self.model.sheet(si) else {
            return vec![r];
        };
        fn bands(
            hidden: &std::collections::BTreeMap<u32, gm::LineInfo>,
            lo: u32,
            hi: u32,
        ) -> Vec<(u32, u32)> {
            let mut out = Vec::new();
            let mut from = lo;
            for (i, _) in hidden.range(lo..=hi).filter(|(_, l)| l.hidden) {
                if *i > from {
                    out.push((from, *i - 1));
                }
                from = *i + 1;
            }
            if from <= hi {
                out.push((from, hi));
            }
            out
        }
        let rows = bands(&sh.rows, r.start.row, r.end.row);
        let cols = bands(&sh.cols, r.start.col, r.end.col);
        rows.iter()
            .flat_map(|(r0, r1)| {
                cols.iter().map(move |(c0, c1)| {
                    RangeRef::new(CellRef::new(*r0, *c0), CellRef::new(*r1, *c1))
                })
            })
            .collect()
    }

    /// `Range.SpecialCells`: the cells of `r` of a kind, as rectangles (empty when there are none).
    pub fn special_cells(
        &mut self,
        si: usize,
        r: RangeRef,
        what: Special,
    ) -> Result<Vec<RangeRef>> {
        self.check_sheet(si)?;
        self.flush();
        let used = self.used_range(si);
        Ok(match what {
            Special::Visible => self.visible_areas(si, r),
            Special::LastCell => used
                .map(|u| vec![RangeRef::cell(u.end)])
                .unwrap_or_else(|| vec![RangeRef::cell(CellRef::new(0, 0))]),
            Special::Blanks => {
                let Some(r) = used.and_then(|u| u.intersection(&r)) else {
                    return Ok(Vec::new());
                };
                if r.count() > 5_000_000 {
                    return Err("範囲が大きすぎます".into());
                }
                areas(r.iter().filter(|at| !self.filled(si, *at)))
            }
            Special::Constants(mask) | Special::Formulas(mask) => {
                let formulas = matches!(what, Special::Formulas(_));
                let sh = self.model.sheet(si).expect("checked");
                areas(
                    sh.cells
                        .iter_range(r)
                        .filter(|(_, c)| {
                            c.formula.is_some() == formulas
                                && (formulas || !c.value.is_empty())
                                && kind_bit(&c.value.scalar()) & mask != 0
                        })
                        .map(|(at, _)| at),
                )
            }
        })
    }

    // ---- sorting

    /// Lays the lines (rows, or columns when `across`) of `r` out again: each (target, source)
    /// of `order` puts what was in line `source` into line `target` (None: nothing). The relative
    /// references of formulas move with them, as when Excel sorts, and the formats of the cells.
    fn rearrange(&mut self, si: usize, r: RangeRef, order: &[(u32, Option<u32>)], across: bool) {
        let line = |i: u32| {
            if across {
                RangeRef::new(CellRef::new(r.start.row, i), CellRef::new(r.end.row, i))
            } else {
                RangeRef::new(CellRef::new(i, r.start.col), CellRef::new(i, r.end.col))
            }
        };
        let order: Vec<(u32, Option<u32>)> = order
            .iter()
            .copied()
            .filter(|(t, s)| Some(*t) != *s)
            .collect();
        // Read first: lines are sources and targets.
        let mut cells: HashMap<u32, Vec<(CellRef, gm::Cell)>> = HashMap::new();
        let mut styles: HashMap<u32, Vec<(CellRef, Style)>> = HashMap::new();
        let mut rich: HashMap<u32, Vec<(CellRef, umya_spreadsheet::RichText)>> = HashMap::new();
        for (_, s) in &order {
            let Some(s) = *s else { continue };
            if cells.contains_key(&s) {
                continue;
            }
            cells.insert(
                s,
                self.model
                    .sheet(si)
                    .map(|sh| {
                        sh.cells
                            .iter_range(line(s))
                            .map(|(at, c)| (at, c.clone()))
                            .collect()
                    })
                    .unwrap_or_default(),
            );
            styles.insert(s, self.styles_in(si, line(s)));
            rich.insert(s, self.rich_in(si, line(s)));
        }
        for (t, s) in order {
            let gone = self
                .model
                .sheet_mut(si)
                .map(|sh| sh.take_cells(line(t)))
                .unwrap_or_default();
            for (at, _) in gone {
                self.touch(si, at);
            }
            self.reset_styles(si, line(t));
            self.clear_rich(si, line(t));
            let Some(s) = s else { continue };
            let d = t as i64 - s as i64;
            let (dr, dc) = if across { (0, d) } else { (d, 0) };
            for (at, c) in cells.get(&s).into_iter().flatten() {
                let Some(to) = at.offset(dr, dc) else {
                    continue;
                };
                let mut c = c.clone();
                if let Some(f) = c.formula.as_ref() {
                    let mut new = match f.expr() {
                        Some(e) => gm::Formula::from_expr(
                            gridcraft_formula::adjust::shift_relative(e, dr, dc),
                        ),
                        None => (**f).clone(),
                    };
                    new.array = None;
                    new.dynamic = f.dynamic;
                    c.formula = Some(Arc::new(new));
                    c.value = Value::Empty;
                }
                if let Some(sh) = self.model.sheet_mut(si) {
                    sh.set_cell(to, c);
                }
                self.touch(si, to);
            }
            if let Ok(ws) = self.file.sheet_mut(si) {
                for (at, style) in styles.get(&s).into_iter().flatten() {
                    if let Some(to) = at.offset(dr, dc) {
                        ws.cell_mut(coord(to)).set_style(style.clone());
                    }
                }
            }
            for (at, rt) in rich.get(&s).into_iter().flatten() {
                if let Some(to) = at.offset(dr, dc) {
                    self.put_rich(si, to, rt.clone());
                }
            }
        }
    }

    /// Whether the first row of `r` looks like headings: text over a column of something else.
    pub fn guess_header(&mut self, si: usize, r: RangeRef) -> bool {
        self.flush();
        if r.height() < 2 {
            return false;
        }
        let Some(sh) = self.model.sheet(si) else {
            return false;
        };
        let mut differs = false;
        for col in r.start.col..=r.end.col {
            let head = sh.value(CellRef::new(r.start.row, col));
            let next = sh.value(CellRef::new(r.start.row + 1, col));
            match head {
                Value::Empty => {}
                Value::Text(_) => differs |= !matches!(next, Value::Text(_) | Value::Empty),
                _ => return false,
            }
        }
        differs
    }

    /// `Range.Sort`: sorts the rows of `r` (the columns, when `across`) by `keys`, as Excel orders
    /// them: numbers, text, FALSE and TRUE, errors; empty cells last either way. Hidden rows stay
    /// where they are. `r` is without the row of headings.
    pub fn sort(
        &mut self,
        si: usize,
        r: RangeRef,
        keys: &[SortKey],
        across: bool,
        match_case: bool,
    ) -> Result<()> {
        self.check_sheet(si)?;
        self.flush();
        // Nothing lies below the used range.
        let Some(r) = self.used_range(si).and_then(|u| u.intersection(&r)) else {
            return Ok(());
        };
        let sh = self.model.sheet(si).expect("checked");
        if sh.merges.iter().any(|m| m.intersects(&r)) {
            return Err(
                "この操作を行うには、すべての結合セルを同じサイズにする必要があります。".into(),
            );
        }
        if sh.cells.iter_range(r).any(|(at, c)| {
            c.formula
                .as_ref()
                .and_then(|f| f.array)
                .is_some_and(|a| !a.is_single() && a.start == at)
        }) {
            return Err("配列の一部を変更することはできません。".into());
        }
        let lines: Vec<u32> = if across {
            (r.start.col..=r.end.col)
                .filter(|c| !sh.is_col_hidden(*c))
                .collect()
        } else {
            (r.start.row..=r.end.row)
                .filter(|row| !sh.is_row_hidden(*row))
                .collect()
        };
        let at = |line: u32, key: u32| {
            if across {
                CellRef::new(key, line)
            } else {
                CellRef::new(line, key)
            }
        };
        let values: HashMap<u32, Vec<Value>> = lines
            .iter()
            .map(|l| {
                (
                    *l,
                    keys.iter()
                        .map(|k| sh.value(at(*l, k.index)).scalar())
                        .collect(),
                )
            })
            .collect();
        let mut sorted = lines.clone();
        sorted.sort_by(|a, b| {
            for (i, k) in keys.iter().enumerate() {
                let (x, y) = (&values[a][i], &values[b][i]);
                let mut o = sort_compare(x, y);
                if o == Ordering::Equal
                    && match_case
                    && let (Value::Text(s), Value::Text(t)) = (x, y)
                {
                    // Lower case before upper case, as Excel sorts with MatchCase.
                    o = t.as_ref().cmp(s.as_ref());
                }
                if k.descending && !x.is_empty() && !y.is_empty() {
                    o = o.reverse();
                }
                if o != Ordering::Equal {
                    return o;
                }
            }
            Ordering::Equal
        });
        let order: Vec<(u32, Option<u32>)> = lines
            .into_iter()
            .zip(sorted.into_iter().map(Some))
            .collect();
        self.rearrange(si, r, &order, across);
        Ok(())
    }

    /// `Range.RemoveDuplicates`: of the rows of `r` (without its headings) that have the same in
    /// the columns `columns` (of the sheet), the first stays; the rows left move up within `r`.
    /// Returns how many rows went.
    pub fn remove_duplicates(&mut self, si: usize, r: RangeRef, columns: &[u32]) -> Result<u32> {
        self.check_sheet(si)?;
        self.flush();
        let Some(r) = self
            .used_range(si)
            .and_then(|u| u.intersection(&r))
            .map(|x| {
                RangeRef::new(
                    CellRef::new(r.start.row, r.start.col),
                    CellRef::new(x.end.row, r.end.col),
                )
            })
        else {
            return Ok(0);
        };
        let sh = self.model.sheet(si).expect("checked");
        let mut seen: HashSet<Vec<String>> = HashSet::new();
        let mut kept: Vec<u32> = Vec::new();
        for row in r.start.row..=r.end.row {
            let key: Vec<String> = columns
                .iter()
                .map(|c| match sh.value(CellRef::new(row, *c)).scalar() {
                    Value::Text(t) => format!("t{}", t.to_lowercase()),
                    Value::Number(n) => format!("n{n}"),
                    other => format!("o{}", other.display()),
                })
                .collect();
            if seen.insert(key) {
                kept.push(row);
            }
        }
        let removed = r.height() - kept.len() as u32;
        if removed > 0 {
            let order: Vec<(u32, Option<u32>)> = (r.start.row..=r.end.row)
                .enumerate()
                .map(|(i, t)| (t, kept.get(i).copied()))
                .collect();
            self.rearrange(si, r, &order, false);
        }
        Ok(removed)
    }

    // ---- filters

    /// The filter of a sheet: its range, and the columns (counted from its first) it filters by.
    pub fn auto_filter(&self, si: usize) -> Option<(RangeRef, Vec<u32>)> {
        let f = self.model.sheet(si)?.autofilter.as_ref()?;
        Some((f.range, f.criteria.iter().map(|(c, _)| *c).collect()))
    }

    /// Puts the filter of a sheet on `r` (its first row is the headings), or takes it away
    /// (None). Rows a filter had hidden show again.
    pub fn set_auto_filter(&mut self, si: usize, r: Option<RangeRef>) -> Result<()> {
        self.check_sheet(si)?;
        self.show_all(si)?;
        if let Some(sh) = self.model.sheet_mut(si) {
            sh.autofilter = r.map(|range| gm::AutoFilter {
                range,
                criteria: Vec::new(),
            });
        }
        Ok(())
    }

    /// Filters by a column of the filter (`offset` from its first column), or no longer (None),
    /// and hides the rows that do not pass.
    pub fn set_filter(
        &mut self,
        si: usize,
        offset: u32,
        criterion: Option<gm::FilterCriterion>,
    ) -> Result<()> {
        self.check_sheet(si)?;
        let Some(f) = self.model.sheet_mut(si).and_then(|s| s.autofilter.as_mut()) else {
            return Err("オートフィルターが設定されていません。".into());
        };
        if offset >= f.range.width() {
            return Err("AutoFilter の Field がフィルターの範囲の外です。".into());
        }
        f.criteria.retain(|(c, _)| *c != offset);
        if let Some(c) = criterion {
            f.criteria.push((offset, c));
            f.criteria.sort_by_key(|(c, _)| *c);
        }
        self.apply_filter(si);
        Ok(())
    }

    /// Whether rows of the sheet are filtered (`Worksheet.FilterMode`).
    pub fn filter_mode(&self, si: usize) -> bool {
        let Some(sh) = self.model.sheet(si) else {
            return false;
        };
        sh.autofilter
            .as_ref()
            .is_some_and(|f| !f.criteria.is_empty())
            || self.advanced.contains_key(&sh.uid.get())
    }

    /// `Worksheet.ShowAllData`: the rows filters hide show again; the filter stays.
    pub fn show_all(&mut self, si: usize) -> Result<()> {
        self.check_sheet(si)?;
        let uid = self
            .model
            .sheet(si)
            .map(|s| s.uid.get())
            .unwrap_or_default();
        if let Some(r) = self.advanced.remove(&uid) {
            self.set_hidden(si, false, r.start.row, r.end.row, false)?;
        }
        let filtered = self
            .model
            .sheet_mut(si)
            .and_then(|s| s.autofilter.as_mut())
            .is_some_and(|f| {
                let any = !f.criteria.is_empty();
                f.criteria.clear();
                any
            });
        if filtered {
            self.apply_filter(si);
        }
        Ok(())
    }

    /// Hides the rows of the filter that do not pass what it filters by, and shows the others.
    fn apply_filter(&mut self, si: usize) {
        self.flush();
        let Some(f) = self.model.sheet(si).and_then(|s| s.autofilter.clone()) else {
            return;
        };
        let last = self
            .used_range(si)
            .map_or(f.range.start.row, |u| u.end.row)
            .min(f.range.end.row);
        let rows: Vec<u32> = (f.range.start.row + 1..=last).collect();
        let mut pass = vec![true; rows.len()];
        for (offset, criterion) in &f.criteria {
            let col = f.range.start.col + offset;
            let cells: Vec<Shown> = rows
                .iter()
                .map(|row| self.shown(si, CellRef::new(*row, col)))
                .collect();
            match criterion {
                gm::FilterCriterion::Values { values, blanks } => {
                    let wanted: HashSet<String> = values.iter().map(|v| v.to_lowercase()).collect();
                    for (i, c) in cells.iter().enumerate() {
                        pass[i] &= if c.text.is_empty() {
                            *blanks
                        } else {
                            wanted.contains(&c.text.to_lowercase())
                        };
                    }
                }
                gm::FilterCriterion::Custom { a, b, and } => {
                    for (i, c) in cells.iter().enumerate() {
                        let first = self.passes(c, &a.0, &a.1, false);
                        pass[i] &= match b {
                            Some(b) if *and => first && self.passes(c, &b.0, &b.1, false),
                            Some(b) => first || self.passes(c, &b.0, &b.1, false),
                            None => first,
                        };
                    }
                }
                gm::FilterCriterion::Top10 {
                    bottom,
                    percent,
                    count,
                } => {
                    let mut numbers: Vec<f64> = cells
                        .iter()
                        .filter_map(|c| {
                            if let Value::Number(n) = c.value {
                                Some(n)
                            } else {
                                None
                            }
                        })
                        .collect();
                    numbers.sort_by(|x, y| {
                        if *bottom {
                            compare_numbers(*x, *y)
                        } else {
                            compare_numbers(*y, *x)
                        }
                    });
                    let n = if *percent {
                        (numbers.len() as u64 * *count as u64 / 100).max(1) as usize
                    } else {
                        *count as usize
                    };
                    let limit = numbers.get(n.clamp(1, numbers.len().max(1)) - 1).copied();
                    for (i, c) in cells.iter().enumerate() {
                        pass[i] &= match (&c.value, limit) {
                            (Value::Number(x), Some(l)) if *bottom => {
                                compare_numbers(*x, l) != Ordering::Greater
                            }
                            (Value::Number(x), Some(l)) => compare_numbers(*x, l) != Ordering::Less,
                            _ => false,
                        };
                    }
                }
                gm::FilterCriterion::AboveAverage(above) => {
                    let numbers: Vec<f64> = cells
                        .iter()
                        .filter_map(|c| {
                            if let Value::Number(n) = c.value {
                                Some(n)
                            } else {
                                None
                            }
                        })
                        .collect();
                    let mean = numbers.iter().sum::<f64>() / numbers.len().max(1) as f64;
                    for (i, c) in cells.iter().enumerate() {
                        pass[i] &= matches!(c.value, Value::Number(x) if (x > mean) == *above && x != mean);
                    }
                }
                _ => {}
            }
        }
        let mut changed = false;
        for (row, show) in rows.into_iter().zip(pass) {
            changed |= self.hide(si, false, row, !show);
        }
        // Rows below what is used are shown.
        if last < f.range.end.row {
            let below: Vec<u32> = self
                .model
                .sheet(si)
                .map(|s| {
                    s.rows
                        .range(last + 1..=f.range.end.row)
                        .filter(|(_, l)| l.hidden)
                        .map(|(i, _)| *i)
                        .collect()
                })
                .unwrap_or_default();
            for row in below {
                changed |= self.hide(si, false, row, false);
            }
        }
        if changed {
            self.rebuild = true;
        }
    }

    /// `Range.AdvancedFilter`: the rows of the list `list` (its first row is the headings) that
    /// pass the criteria of `criteria` (headings, and under them rows of comparisons: a row
    /// passes when every comparison of one of them holds), and only the first of rows that show
    /// the same when `unique`. They are copied to `copy_to` (all columns, or the ones its first
    /// row names), or the other rows of the list are hidden.
    pub fn advanced_filter(
        &mut self,
        si: usize,
        list: RangeRef,
        criteria: Option<(usize, RangeRef)>,
        copy_to: Option<(usize, RangeRef)>,
        unique: bool,
    ) -> Result<()> {
        self.check_sheet(si)?;
        self.flush();
        let used = self.used_range(si).unwrap_or(list);
        let list = RangeRef::new(
            list.start,
            CellRef::new(
                list.end.row.min(used.end.row.max(list.start.row)),
                list.end.col.min(used.end.col.max(list.start.col)),
            ),
        );
        let heads: Vec<String> = (list.start.col..=list.end.col)
            .map(|c| {
                self.shown(si, CellRef::new(list.start.row, c))
                    .text
                    .to_lowercase()
            })
            .collect();
        // Rows of (column of the list, operator, what it compares with).
        let mut wanted: Vec<Vec<(u32, String, String)>> = Vec::new();
        if let Some((ci, cr)) = criteria {
            self.check_sheet(ci)?;
            let cused = self.used_range(ci).unwrap_or(cr);
            let cr = RangeRef::new(
                cr.start,
                CellRef::new(cr.end.row.min(cused.end.row.max(cr.start.row)), cr.end.col),
            );
            let mut columns: Vec<Option<u32>> = Vec::new();
            for c in cr.start.col..=cr.end.col {
                let name = self
                    .shown(ci, CellRef::new(cr.start.row, c))
                    .text
                    .to_lowercase();
                columns.push(
                    heads
                        .iter()
                        .position(|h| !name.is_empty() && *h == name)
                        .map(|i| list.start.col + i as u32),
                );
            }
            for row in cr.start.row + 1..=cr.end.row {
                let mut tests = Vec::new();
                for (i, c) in (cr.start.col..=cr.end.col).enumerate() {
                    let text = self.shown(ci, CellRef::new(row, c)).text;
                    if text.is_empty() {
                        continue;
                    }
                    let Some(col) = columns[i] else {
                        return Err("AdvancedFilter の検索条件範囲の見出しが、リストの見出しにありません（数式による条件には対応していません）。".into());
                    };
                    let (op, val) = split_criterion(&text);
                    tests.push((col, op.to_string(), val.to_string()));
                }
                wanted.push(tests);
            }
        }
        let mut seen: HashSet<Vec<String>> = HashSet::new();
        let mut rows: Vec<(u32, bool)> = Vec::new();
        for row in list.start.row + 1..=list.end.row {
            let mut pass = wanted.is_empty();
            for tests in &wanted {
                let mut all = true;
                for (col, op, val) in tests {
                    let cell = self.shown(si, CellRef::new(row, *col));
                    all &= self.passes(&cell, op, val, true);
                }
                if all {
                    pass = true;
                    break;
                }
            }
            if pass && unique {
                let key: Vec<String> = (list.start.col..=list.end.col)
                    .map(|c| self.shown(si, CellRef::new(row, c)).text)
                    .collect();
                pass = seen.insert(key);
            }
            rows.push((row, pass));
        }
        let Some((ti, to)) = copy_to else {
            // In place: the other rows are hidden until ShowAllData.
            let mut changed = false;
            for (row, pass) in rows {
                changed |= self.hide(si, false, row, !pass);
            }
            if let Some(sh) = self.model.sheet(si) {
                self.advanced.insert(
                    sh.uid.get(),
                    RangeRef::new(CellRef::new(list.start.row + 1, list.start.col), list.end),
                );
            }
            if changed {
                self.rebuild = true;
            }
            return Ok(());
        };
        self.check_sheet(ti)?;
        // The columns to copy: those the first row of the target names, or all of them.
        let named: Vec<String> = (to.start.col..=to.end.col)
            .map(|c| {
                self.shown(ti, CellRef::new(to.start.row, c))
                    .text
                    .to_lowercase()
            })
            .collect();
        let columns: Vec<u32> = if named.iter().all(String::is_empty) {
            (list.start.col..=list.end.col).collect()
        } else {
            let mut out = Vec::new();
            for name in named.iter().filter(|n| !n.is_empty()) {
                let Some(i) = heads.iter().position(|h| h == name) else {
                    return Err(
                        "AdvancedFilter の抽出範囲の見出しが、リストの見出しにありません。".into(),
                    );
                };
                out.push(list.start.col + i as u32);
            }
            out
        };
        let values = Paste {
            what: PasteWhat::Values,
            op: PasteOp::None,
            skip_blanks: false,
            transpose: false,
        };
        let formats = Paste {
            what: PasteWhat::Formats,
            ..values
        };
        let mut target = to.start.row;
        for row in std::iter::once(list.start.row)
            .chain(rows.into_iter().filter(|(_, p)| *p).map(|(r, _)| r))
        {
            for (i, col) in columns.iter().enumerate() {
                let from = RangeRef::cell(CellRef::new(row, *col));
                let Some(at) = CellRef::new(target, to.start.col).offset(0, i as i64) else {
                    continue;
                };
                self.paste(si, from, ti, at, &values)?;
                self.paste(si, from, ti, at, &formats)?;
            }
            target += 1;
            if target >= super::MAX_ROWS {
                break;
            }
        }
        Ok(())
    }
}
