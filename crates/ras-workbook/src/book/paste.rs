//! Copying and moving cells: `Range.Copy`, `Range.Cut`, and pasting what was cut or copied
//! (`Worksheet.Paste`, `Range.PasteSpecial`). Cells are pasted with their formats, the relative
//! references of copied formulas move with them, and the references to cells that were cut
//! follow them, as in Excel.

use std::collections::BTreeMap;
use std::sync::Arc;

use gridcraft_core::{CellRef, RangeRef, Value};
use gridcraft_formula::adjust::Edit;
use gridcraft_formula::{Expr, SheetSel};
use gridcraft_model as gm;
use umya_spreadsheet::Style;

use super::{Book, Result, coord};

/// What `PasteSpecial` pastes (`XlPasteType`).
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum PasteWhat {
    All,
    AllExceptBorders,
    Formulas,
    FormulasAndNumberFormats,
    Values,
    ValuesAndNumberFormats,
    Formats,
    ColumnWidths,
}

impl PasteWhat {
    /// `XlPasteType` (None: what the stand-in does not paste — comments, validation).
    pub fn of(n: i32) -> Option<PasteWhat> {
        Some(match n {
            -4104 | 13 | 14 => PasteWhat::All,
            7 => PasteWhat::AllExceptBorders,
            -4123 => PasteWhat::Formulas,
            11 => PasteWhat::FormulasAndNumberFormats,
            -4163 => PasteWhat::Values,
            12 => PasteWhat::ValuesAndNumberFormats,
            -4122 => PasteWhat::Formats,
            8 => PasteWhat::ColumnWidths,
            _ => return None,
        })
    }

    fn formulas(self) -> bool {
        matches!(
            self,
            PasteWhat::All
                | PasteWhat::AllExceptBorders
                | PasteWhat::Formulas
                | PasteWhat::FormulasAndNumberFormats
        )
    }

    fn values(self) -> bool {
        matches!(self, PasteWhat::Values | PasteWhat::ValuesAndNumberFormats)
    }

    fn styles(self) -> bool {
        matches!(
            self,
            PasteWhat::All | PasteWhat::AllExceptBorders | PasteWhat::Formats
        )
    }

    fn number_formats(self) -> bool {
        matches!(
            self,
            PasteWhat::FormulasAndNumberFormats | PasteWhat::ValuesAndNumberFormats
        )
    }
}

/// `XlPasteSpecialOperation`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum PasteOp {
    None,
    Add,
    Subtract,
    Multiply,
    Divide,
}

impl PasteOp {
    pub fn of(n: i32) -> PasteOp {
        match n {
            2 => PasteOp::Add,
            3 => PasteOp::Subtract,
            4 => PasteOp::Multiply,
            5 => PasteOp::Divide,
            _ => PasteOp::None,
        }
    }

    fn symbol(self) -> &'static str {
        match self {
            PasteOp::Add => "+",
            PasteOp::Subtract => "-",
            PasteOp::Multiply => "*",
            PasteOp::Divide => "/",
            PasteOp::None => "",
        }
    }
}

/// How cells are pasted.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Paste {
    pub what: PasteWhat,
    pub op: PasteOp,
    /// Empty cells of the source leave the cells they are pasted on as they are.
    pub skip_blanks: bool,
    /// Rows become columns.
    pub transpose: bool,
}

impl Paste {
    pub fn all() -> Paste {
        Paste {
            what: PasteWhat::All,
            op: PasteOp::None,
            skip_blanks: false,
            transpose: false,
        }
    }
}

/// A cell to paste.
#[derive(Default)]
struct Source {
    own: Option<gm::Cell>,
    /// What the cell shows (its own value, or what a formula spilled into it).
    shown: Value,
    style: Option<Style>,
}

impl Book {
    /// The formats of the stored cells of `r` (cells with the default format are left out).
    pub(super) fn styles_in(&self, si: usize, r: RangeRef) -> Vec<(CellRef, Style)> {
        let Ok(ws) = self.file.sheet(si) else {
            return Vec::new();
        };
        let plain = Style::default();
        if r.count() <= 4096 {
            return r
                .iter()
                .filter_map(|at| {
                    ws.cell(coord(at))
                        .filter(|c| *c.style() != plain)
                        .map(|c| (at, c.style().clone()))
                })
                .collect();
        }
        ws.cells()
            .iter()
            .map(|c| (super::at_of(c), c))
            .filter(|(at, c)| r.contains(*at) && *c.style() != plain)
            .map(|(at, c)| (at, c.style().clone()))
            .collect()
    }

    /// Gives the stored cells of `r` the default format.
    pub(super) fn reset_styles(&mut self, si: usize, r: RangeRef) {
        let stored: Vec<CellRef> = self
            .styles_in(si, r)
            .into_iter()
            .map(|(at, _)| at)
            .collect();
        if let Ok(ws) = self.file.sheet_mut(si) {
            for at in stored {
                ws.cell_mut(coord(at)).set_style(Style::default());
            }
        }
    }

    /// The merged cells of the file that lie within `r`.
    fn merges_in(&self, si: usize, r: RangeRef) -> Vec<RangeRef> {
        let Ok(ws) = self.file.sheet(si) else {
            return Vec::new();
        };
        ws.merge_cells()
            .iter()
            .filter_map(|m| RangeRef::parse(&m.range().replace('$', "")))
            .filter(|m| r.contains_range(m))
            .collect()
    }

    /// Takes the merged cells that touch `r` apart.
    pub(super) fn unmerge(&mut self, si: usize, r: RangeRef) {
        if let Ok(ws) = self.file.sheet_mut(si) {
            ws.merge_cells_mut().retain(|m| {
                RangeRef::parse(&m.range().replace('$', "")).is_none_or(|m| !m.intersects(&r))
            });
        }
        if let Some(sh) = self.model.sheet_mut(si) {
            sh.merges.retain(|m| !m.intersects(&r));
        }
    }

    /// Merges the cells of `r`.
    pub(super) fn merge(&mut self, si: usize, r: RangeRef) {
        self.unmerge(si, r);
        if r.is_single() {
            return;
        }
        if let Ok(ws) = self.file.sheet_mut(si) {
            ws.add_merge_cells(r.a1());
        }
        if let Some(sh) = self.model.sheet_mut(si) {
            sh.merges.push(r);
        }
    }

    /// `Range.Copy Destination`: the cells of `from` (on sheet `si`) with their formats, into the
    /// cells from `to` (on sheet `ti`) on. The relative references of the formulas move with
    /// them, as when Excel pastes.
    pub fn copy(&mut self, si: usize, from: RangeRef, ti: usize, to: CellRef) -> Result<()> {
        self.paste(si, from, ti, to, &Paste::all())
    }

    /// Pastes the cells of `from` (on sheet `si`) into the cells from `to` (on sheet `ti`) on,
    /// as `Range.PasteSpecial` does.
    pub fn paste(
        &mut self,
        si: usize,
        from: RangeRef,
        ti: usize,
        to: CellRef,
        p: &Paste,
    ) -> Result<()> {
        self.check_sheet(si)?;
        self.check_sheet(ti)?;
        let (h, w) = (from.height() as i64, from.width() as i64);
        let (th, tw) = if p.transpose { (w, h) } else { (h, w) };
        let Some(end) = to.offset(th - 1, tw - 1) else {
            return Err("貼り付け先がシートの外です".into());
        };
        let target = RangeRef::new(to, end);
        // Where a cell of the source goes.
        let place = |at: CellRef| -> CellRef {
            let (dr, dc) = (at.row - from.start.row, at.col - from.start.col);
            if p.transpose {
                CellRef::new(to.row + dc, to.col + dr)
            } else {
                CellRef::new(to.row + dr, to.col + dc)
            }
        };
        if p.what == PasteWhat::ColumnWidths {
            return self.paste_widths(si, from, ti, to);
        }
        self.flush();

        // Everything is read first: the ranges may overlap.
        let content = p.what.formulas() || p.what.values();
        let formats = p.what.styles() || p.what.number_formats();
        let mut cells: BTreeMap<CellRef, Source> = BTreeMap::new();
        if content && let Some(sh) = self.model.sheet(si) {
            for (at, c) in sh.cells.iter_range(from) {
                let e = cells.entry(at).or_default();
                e.shown = c.value.clone();
                e.own = Some(c.clone());
            }
            for (at, v) in sh
                .spill
                .range(from.start..=from.end)
                .filter(|(at, _)| from.contains(**at))
            {
                // A spilled value is pasted as a value, unless its formula is pasted too (and
                // spills again).
                let anchored = sh.spill_anchor(*at).is_some_and(|a| from.contains(a));
                if p.what.values() || !anchored {
                    let e = cells.entry(*at).or_default();
                    e.shown = v.clone();
                    if !p.what.values() {
                        e.own = Some(gm::Cell::value(v.clone()));
                    }
                }
            }
        }
        if formats {
            for (at, style) in self.styles_in(si, from) {
                cells.entry(at).or_default().style = Some(style);
            }
        }
        let merges = if p.what.styles() {
            self.merges_in(si, from)
        } else {
            Vec::new()
        };

        // The cells pasted on that the source has nothing for.
        if content && p.op == PasteOp::None && !p.skip_blanks {
            let gone: Vec<CellRef> = self
                .model
                .sheet(ti)
                .map(|s| s.cells.iter_range(target).map(|(at, _)| at).collect())
                .unwrap_or_default();
            if let Some(sh) = self.model.sheet_mut(ti) {
                for at in &gone {
                    sh.remove_cell(*at);
                }
            }
            for at in gone {
                self.touch(ti, at);
            }
        }
        if formats && !p.skip_blanks {
            let kept: Vec<(CellRef, Style)> = self.styles_in(ti, target);
            if let Ok(ws) = self.file.sheet_mut(ti) {
                for (at, old) in kept {
                    let mut plain = Style::default();
                    match p.what {
                        PasteWhat::AllExceptBorders => {
                            if let Some(b) = old.borders() {
                                plain.set_borders(b.clone());
                            }
                        }
                        PasteWhat::FormulasAndNumberFormats | PasteWhat::ValuesAndNumberFormats => {
                            plain = old;
                            plain.remove_number_format();
                        }
                        _ => {}
                    }
                    ws.cell_mut(coord(at)).set_style(plain);
                }
            }
        }
        if p.what.styles() {
            self.unmerge(ti, target);
        }

        let source = self
            .model
            .sheet(si)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        // A formula of the source as it is pasted at `t`: its relative references moved (rows
        // and columns swapped when transposing).
        let moved = |f: &gm::Formula, at: CellRef, t: CellRef| -> gm::Formula {
            let e = f.expr().map(|e| {
                if p.transpose {
                    transpose_refs(e, &source, from, to, at, t)
                } else {
                    let (dr, dc) = (t.row as i64 - at.row as i64, t.col as i64 - at.col as i64);
                    gridcraft_formula::adjust::shift_relative(e, dr, dc)
                }
            });
            let mut formula = match e {
                Some(e) => gm::Formula::from_expr(e),
                None => f.clone(),
            };
            formula.array = None;
            formula.dynamic = f.dynamic;
            formula
        };
        for (at, src) in cells {
            let t = place(at);
            let blank = src
                .own
                .as_ref()
                .is_none_or(|c| c.formula.is_none() && c.value.is_empty())
                && src.shown.is_empty();
            if content && !(blank && (p.skip_blanks || p.op != PasteOp::None)) {
                if p.op != PasteOp::None {
                    // Formulas pasted with an operation are joined to what the cell has
                    // (`=100+(B1*2)`); values pasted are worked out.
                    match src.own.as_ref().and_then(|c| c.formula.as_ref()) {
                        Some(f) if p.what.formulas() => {
                            let f = moved(f, at, t);
                            self.paste_operation_formula(ti, t, &f.text, p.op);
                        }
                        _ => self.paste_operation(ti, t, &src.shown, p.op),
                    }
                } else if let Some(sh) = self.model.sheet_mut(ti) {
                    match &src.own {
                        Some(c) if p.what.formulas() && c.formula.is_some() => {
                            let formula = moved(c.formula.as_ref().expect("checked"), at, t);
                            let style = sh.style_id(t);
                            sh.set_cell(
                                t,
                                gm::Cell {
                                    value: Value::Empty,
                                    formula: Some(Arc::new(formula)),
                                    style,
                                    py: None,
                                },
                            );
                        }
                        _ if !src.shown.is_empty() => sh.set_value(t, src.shown.scalar()),
                        _ => {
                            sh.remove_cell(t);
                        }
                    }
                }
                self.touch(ti, t);
            }
            if formats && !(p.skip_blanks && blank && src.style.is_none()) {
                let Ok(ws) = self.file.sheet_mut(ti) else {
                    continue;
                };
                match (p.what, src.style) {
                    (
                        PasteWhat::FormulasAndNumberFormats | PasteWhat::ValuesAndNumberFormats,
                        Some(style),
                    ) => {
                        let code = style
                            .number_format()
                            .map(|f| f.format_code().to_string())
                            .unwrap_or_default();
                        if !code.is_empty() {
                            ws.style_mut(coord(t))
                                .number_format_mut()
                                .set_format_code(code);
                        }
                    }
                    (PasteWhat::AllExceptBorders, Some(mut style)) => {
                        match ws.cell(coord(t)).and_then(|c| c.style().borders().cloned()) {
                            Some(b) => style.set_borders(b),
                            None => style.remove_borders(),
                        };
                        ws.cell_mut(coord(t)).set_style(style);
                    }
                    (PasteWhat::All | PasteWhat::Formats, Some(style)) => {
                        ws.cell_mut(coord(t)).set_style(style);
                    }
                    _ => {}
                }
            }
        }
        for m in merges {
            let (a, z) = (place(m.start), place(m.end));
            self.merge(ti, RangeRef::new(a, z));
        }
        Ok(())
    }

    /// `PasteSpecial Operation:=`: the number pasted is added to (taken from, ...) what the cell has.
    fn paste_operation(&mut self, ti: usize, t: CellRef, v: &Value, op: PasteOp) {
        let Value::Number(n) = v.scalar() else { return };
        let Some(sh) = self.model.sheet_mut(ti) else {
            return;
        };
        let old = sh.cell(t).cloned();
        match old {
            Some(c) if c.formula.is_some() => {
                let f = c.formula.as_ref().expect("checked");
                let text = format!(
                    "({}){}{}",
                    f.text,
                    op.symbol(),
                    gridcraft_core::number_to_text(n)
                );
                let mut formula = gm::Formula::new(&text);
                formula.dynamic = f.dynamic;
                sh.set_cell(
                    t,
                    gm::Cell {
                        value: Value::Empty,
                        formula: Some(Arc::new(formula)),
                        style: c.style,
                        py: None,
                    },
                );
            }
            other => {
                let x = match other.map(|c| c.value) {
                    None | Some(Value::Empty) => 0.0,
                    Some(Value::Number(x)) => x,
                    Some(_) => return,
                };
                let y = match op {
                    PasteOp::Add => Value::number(x + n),
                    PasteOp::Subtract => Value::number(x - n),
                    PasteOp::Multiply => Value::number(x * n),
                    PasteOp::Divide if n == 0.0 => Value::Error(gridcraft_core::CellError::Div0),
                    PasteOp::Divide => Value::number(x / n),
                    PasteOp::None => return,
                };
                sh.set_value(t, y);
            }
        }
    }

    /// `PasteSpecial Operation:=` with a formula pasted (`formula`, without `=`): Excel joins it
    /// to what the cell has, `=(<its formula>)+(<formula>)` or `=<its number>+(<formula>)` (an
    /// empty cell as 0). Text and other values are left as they are.
    fn paste_operation_formula(&mut self, ti: usize, t: CellRef, formula: &str, op: PasteOp) {
        let Some(sh) = self.model.sheet_mut(ti) else {
            return;
        };
        let old = sh.cell(t).cloned();
        let left = match &old {
            Some(c) if c.formula.is_some() => {
                format!("({})", c.formula.as_ref().expect("checked").text)
            }
            None => "0".to_string(),
            Some(c) => match c.value {
                Value::Empty => "0".to_string(),
                Value::Number(x) => gridcraft_core::number_to_text(x),
                _ => return,
            },
        };
        let text = format!("{left}{}({formula})", op.symbol());
        let style = sh.style_id(t);
        sh.set_cell(
            t,
            gm::Cell {
                value: Value::Empty,
                formula: Some(Arc::new(gm::Formula::new(&text))),
                style,
                py: None,
            },
        );
    }

    /// `PasteSpecial xlPasteColumnWidths`.
    fn paste_widths(&mut self, si: usize, from: RangeRef, ti: usize, to: CellRef) -> Result<()> {
        let widths: Vec<Option<f64>> = (from.start.col..=from.end.col)
            .map(|c| {
                self.file
                    .sheet(si)
                    .ok()
                    .and_then(|ws| ws.column_dimension_by_number(c + 1))
                    .map(|d| d.width())
            })
            .collect();
        let Ok(ws) = self.file.sheet_mut(ti) else {
            return Ok(());
        };
        for (i, w) in widths.into_iter().enumerate() {
            let col = to.col + i as u32 + 1;
            if col > super::MAX_COLS {
                break;
            }
            match w {
                Some(w) => {
                    ws.column_dimension_by_number_mut(col).set_width(w);
                }
                None if ws.column_dimension_by_number(col).is_some() => {
                    ws.column_dimension_by_number_mut(col)
                        .set_width(super::format::DEFAULT_COLUMN_WIDTH);
                }
                None => {}
            }
        }
        Ok(())
    }

    /// `Range.Cut Destination` (and pasting what was cut): the cells of `from` (on sheet `si`)
    /// move to the cells from `to` (on sheet `ti`) on, with their formats. Their formulas keep
    /// the cells they refer to, and every reference to a cell that moved follows it.
    pub fn move_cells(&mut self, si: usize, from: RangeRef, ti: usize, to: CellRef) -> Result<()> {
        self.check_sheet(si)?;
        self.check_sheet(ti)?;
        let Some(end) = to.offset(from.height() as i64 - 1, from.width() as i64 - 1) else {
            return Err("貼り付け先がシートの外です".into());
        };
        let target = RangeRef::new(to, end);
        if si == ti && from == target {
            return Ok(());
        }
        let (dr, dc) = (
            to.row as i64 - from.start.row as i64,
            to.col as i64 - from.start.col as i64,
        );
        let source = self
            .model
            .sheet(si)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        let dest = self
            .model
            .sheet(ti)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        let edit = if si == ti {
            Edit::Move {
                from,
                to_row: to.row,
                to_col: to.col,
            }
        } else {
            Edit::MoveToSheet {
                from,
                to_sheet: dest.clone(),
                to_row: to.row,
                to_col: to.col,
            }
        };
        self.adjust_refs(&source, &edit);

        let cells = self
            .model
            .sheet_mut(si)
            .map(|sh| sh.take_cells(from))
            .unwrap_or_default();
        let styles = self.styles_in(si, from);
        let merges = self.merges_in(si, from);
        self.reset_styles(si, from);
        self.unmerge(si, from);
        self.reset_styles(ti, target);
        self.unmerge(ti, target);
        if let Some(sh) = self.model.sheet_mut(ti) {
            sh.take_cells(target);
            for (at, mut c) in cells {
                let Some(t) = at.offset(dr, dc) else { continue };
                if si != ti
                    && let Some(f) = c.formula.as_ref()
                    && let Some(e) = f.expr()
                {
                    // On another sheet now: the cells it read on its own sheet are named, and
                    // those that moved with it are its own.
                    let e = e.map(&mut |x| match x {
                        Expr::Ref(mut r) => {
                            r.sheet = match r.sheet {
                                SheetSel::Current => SheetSel::Named(source.clone()),
                                SheetSel::Named(n) if n.eq_ignore_ascii_case(&dest) => {
                                    SheetSel::Current
                                }
                                other => other,
                            };
                            Expr::Ref(r)
                        }
                        other => other,
                    });
                    let mut new = gm::Formula::from_expr(e);
                    new.array = f.array;
                    new.dynamic = f.dynamic;
                    c.formula = Some(Arc::new(new));
                }
                if let Some(f) = c.formula.as_ref()
                    && let Some(area) = f.array
                {
                    let mut new = (**f).clone();
                    new.array = area
                        .start
                        .offset(dr, dc)
                        .zip(area.end.offset(dr, dc))
                        .map(|(a, z)| RangeRef::new(a, z));
                    c.formula = Some(Arc::new(new));
                }
                sh.set_cell(t, c);
            }
        }
        if let Ok(ws) = self.file.sheet_mut(ti) {
            for (at, style) in styles {
                if let Some(t) = at.offset(dr, dc) {
                    ws.cell_mut(coord(t)).set_style(style);
                }
            }
        }
        for m in merges {
            if let Some((a, z)) = m.start.offset(dr, dc).zip(m.end.offset(dr, dc)) {
                self.merge(ti, RangeRef::new(a, z));
            }
        }
        for i in [si, ti] {
            if let Some(sh) = self.model.sheet_mut(i) {
                sh.spill.clear();
                sh.spill_ranges.clear();
            }
        }
        self.pending.clear();
        self.rebuild = true;
        Ok(())
    }
}

/// The references of a formula copied from `at` and pasted transposed at `t` (the cells of
/// `from`, on the sheet `source`, pasted from `to` on), as Excel 2021 writes them:
///
/// * A reference within `from` (on the same sheet), `$` or not, goes where that cell was
///   pasted, with the `$` of its row and its column swapped (`=$B2` → `=G$10`).
/// * Any other reference without a `$` keeps where it is from the cell, with rows and columns
///   swapped (`=A1` in C3 → `=F9` in H11). Pushed off the sheet, it becomes `#REF!`.
/// * Whole rows and columns with a relative part become the columns and rows they swap to,
///   from where the first one goes to the sheet's last row or column (`=SUM(2:2)` in B3 →
///   `=SUM(E$4:E$1048576)` in F5).
/// * The rest (with a `$` somewhere, outside `from`) is left as it is.
fn transpose_refs(
    e: Expr,
    source: &str,
    from: RangeRef,
    to: CellRef,
    at: CellRef,
    t: CellRef,
) -> Expr {
    use gridcraft_core::{CellError, MAX_COLS, MAX_ROWS};
    use gridcraft_formula::{Anchor, RefKind};
    let here = |s: &SheetSel| match s {
        SheetSel::Current => true,
        SheetSel::Named(n) => gridcraft_formula::same_name(n, source),
        SheetSel::Span(..) => false,
    };
    // A cell `row`, `col` relative to the cell copied, swapped, from the cell pasted.
    let swap = |row: i64, col: i64| -> (i64, i64) {
        (
            t.row as i64 + (col - at.col as i64),
            t.col as i64 + (row - at.row as i64),
        )
    };
    let inside = |row: i64, col: i64| -> Option<(u32, u32)> {
        (row >= 0 && col >= 0 && row < MAX_ROWS as i64 && col < MAX_COLS as i64)
            .then_some((row as u32, col as u32))
    };
    // A cell of `from`: where it was pasted.
    let placed = |a: &Anchor| Anchor {
        row: to.row + (a.col - from.start.col),
        col: to.col + (a.row - from.start.row),
        row_abs: a.col_abs,
        col_abs: a.row_abs,
    };
    let relative = |a: &Anchor| -> Option<Anchor> {
        let (row, col) = swap(a.row as i64, a.col as i64);
        let (row, col) = inside(row, col)?;
        Some(Anchor { row, col, ..*a })
    };
    let clip = |row: i64, col: i64, row_abs: bool, col_abs: bool| {
        let (row, col) = swap(row, col);
        Anchor {
            row: row.clamp(0, MAX_ROWS as i64 - 1) as u32,
            col: col.clamp(0, MAX_COLS as i64 - 1) as u32,
            row_abs,
            col_abs,
        }
    };
    let free = |a: &Anchor| !a.row_abs && !a.col_abs;
    e.map(&mut |x| {
        let Expr::Ref(r) = x else { return x };
        let within = here(&r.sheet) && from.contains_range(&r.range());
        let kind = match &r.kind {
            RefKind::Cell(a) if within => Some(RefKind::Cell(placed(a))),
            RefKind::Range(a, b) if within => Some(RefKind::Range(placed(a), placed(b))),
            RefKind::Cell(a) if free(a) => relative(a).map(RefKind::Cell),
            RefKind::Range(a, b) if free(a) && free(b) => relative(a)
                .zip(relative(b))
                .map(|(a, b)| RefKind::Range(a, b)),
            RefKind::Rows(r0, a0, r1, a1) if !*a0 || !*a1 => Some(RefKind::Range(
                clip(*r0 as i64, 0, true, *a0),
                Anchor {
                    row: MAX_ROWS - 1,
                    ..clip(*r1 as i64, 0, true, *a1)
                },
            )),
            RefKind::Cols(c0, a0, c1, a1) if !*a0 || !*a1 => Some(RefKind::Range(
                clip(0, *c0 as i64, *a0, true),
                Anchor {
                    col: MAX_COLS - 1,
                    ..clip(0, *c1 as i64, *a1, true)
                },
            )),
            other => Some(other.clone()),
        };
        match kind {
            Some(kind) => Expr::Ref(gridcraft_formula::Reference {
                sheet: r.sheet,
                kind,
            }),
            None => Expr::Error(CellError::Ref),
        }
    })
}
