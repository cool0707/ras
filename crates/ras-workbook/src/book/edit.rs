//! Cells inserted and deleted (`Range.Insert`, `Range.Delete`), as Excel does: the cells below
//! or to the right move, and so does every reference to them.
//!
//! Both copies of the workbook are edited: the engine's model (cells and formulas) here, the
//! file (formats, sizes, merged cells, the ranges of validations and conditional formats, ...)
//! by umya-spreadsheet's `Workbook::edit_lines`.

use std::sync::Arc;

use gridcraft_core::{CellRef, RangeRef};
use gridcraft_formula::Expr;
use gridcraft_formula::adjust::{Axis, Edit, adjust, adjust_if_changed};
use gridcraft_model as gm;
use gridcraft_xlsx_formulas::{Names, fmla};
use umya_spreadsheet::helper::rewrite::LineEdit;

use super::{Book, MAX_COLS, MAX_ROWS, Result, coord, layout};

/// The name columns inserted into a table get (with a number), as Japanese Excel names them.
const TABLE_COLUMN: &str = "列";

/// A formula of the file with the references to `sheet` moved by `edit`; None when none moves.
pub(crate) fn adjusted(host: Option<&str>, text: &str, sheet: &str, edit: &Edit) -> Option<String> {
    let body = text.trim();
    let body = body.strip_prefix('=').unwrap_or(body);
    if body.is_empty() {
        return None;
    }
    let e = gridcraft_formula::parse(&fmla::from_file(body)).ok()?;
    // A host no sheet has: references without a sheet are left alone.
    let new = adjust_if_changed(&e, host.unwrap_or("\u{0}"), sheet, edit)?;
    Some(fmla::text_to_file(&gridcraft_formula::print(&new)))
}

/// Where an edit of a sheet moves a range of it (None: all of it is deleted).
pub(crate) fn moved(r: RangeRef, edit: &Edit) -> Option<RangeRef> {
    let e = gridcraft_formula::parse(&r.a1()).ok()?;
    match adjust(e, "s", "s", edit) {
        Expr::Ref(x) => Some(x.range()),
        _ => None,
    }
}

impl Book {
    /// `Range.Insert`: cells inserted at `r`, the cells there and below moving down (`right`: the
    /// cells there and to their right moving right). Whole rows (columns) when `r` spans the
    /// sheet. The new cells get the format of the cells above or to the left (`from_after`:
    /// below or to the right), as Excel gives them.
    pub fn insert_cells(
        &mut self,
        si: usize,
        r: RangeRef,
        right: bool,
        from_after: bool,
    ) -> Result<()> {
        self.restructure(
            si,
            r,
            if right { Axis::Cols } else { Axis::Rows },
            true,
            from_after,
        )
    }

    /// `Range.Delete`: the cells of `r` go, the cells below moving up (`left`: those to the right
    /// moving left).
    pub fn delete_cells(&mut self, si: usize, r: RangeRef, left: bool) -> Result<()> {
        self.restructure(
            si,
            r,
            if left { Axis::Cols } else { Axis::Rows },
            false,
            false,
        )
    }

    fn restructure(
        &mut self,
        si: usize,
        r: RangeRef,
        axis: Axis,
        insert: bool,
        from_after: bool,
    ) -> Result<()> {
        self.check_sheet(si)?;
        let whole = match axis {
            Axis::Rows => r.is_full_rows(),
            Axis::Cols => r.is_full_cols(),
        };
        let (at, count, max) = match axis {
            Axis::Rows => (r.start.row, r.height(), MAX_ROWS),
            Axis::Cols => (r.start.col, r.width(), MAX_COLS),
        };
        // The other axis: the columns of rows inserted, the rows of columns inserted.
        let band = match axis {
            Axis::Rows => (r.start.col, r.end.col),
            Axis::Cols => (r.start.row, r.end.row),
        };
        // Everything the edit moves or removes.
        let reach = match axis {
            Axis::Rows => RangeRef::new(r.start, CellRef::new(MAX_ROWS - 1, r.end.col)),
            Axis::Cols => RangeRef::new(r.start, CellRef::new(r.end.row, MAX_COLS - 1)),
        };
        self.check_edit(si, r, reach, axis, insert, whole)?;

        let name = self
            .model
            .sheet(si)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        let edit = match (whole, insert) {
            (true, true) => Edit::Insert { axis, at, count },
            (true, false) => Edit::Delete { axis, at, count },
            (false, true) => Edit::InsertCells { axis, range: r },
            (false, false) => Edit::DeleteCells { axis, range: r },
        };
        // Formulas first: they are found where they are now.
        self.adjust_refs(&name, &edit);
        if let Some(sh) = self.model.sheet_mut(si) {
            let n = count as i64;
            if !insert {
                sh.cells.take_range(r);
            }
            let from = if insert { at } else { at + count };
            let delta = if insert { n } else { -n };
            if from < max {
                match (axis, whole) {
                    (Axis::Rows, true) => sh.cells.shift_rows(from, delta),
                    (Axis::Cols, true) => sh.cells.shift_cols(from, delta),
                    (Axis::Rows, false) => sh.cells.shift_rows_in_cols(band.0, band.1, from, delta),
                    (Axis::Cols, false) => sh.cells.shift_cols_in_rows(band.0, band.1, from, delta),
                }
            }
            // Array formulas move with their cells.
            let arrays: Vec<(CellRef, gm::Formula)> = sh
                .cells
                .iter()
                .filter_map(|(at, c)| {
                    let f = c.formula.as_ref()?;
                    let area = f.array?;
                    let new = moved(area, &edit).filter(|x| *x != area)?;
                    let mut f = (**f).clone();
                    f.array = Some(new);
                    Some((at, f))
                })
                .collect();
            for (at, f) in arrays {
                if let Some(c) = sh.cells.get_mut(at) {
                    c.formula = Some(Arc::new(f));
                }
            }
            // What formulas spilled is laid out again when they are calculated.
            sh.spill.clear();
            sh.spill_ranges.clear();
            // The filter, with the columns it filters by.
            if let Some(f) = sh.autofilter.take() {
                let first = |range: RangeRef| CellRef::new(range.start.row, range.start.col);
                sh.autofilter = moved(f.range, &edit).map(|range| gm::AutoFilter {
                    range,
                    criteria: f
                        .criteria
                        .into_iter()
                        .filter_map(|(off, c)| {
                            let old = CellRef::new(f.range.start.row, f.range.start.col + off);
                            let new = moved(RangeRef::cell(old), &edit)?;
                            Some((new.start.col.checked_sub(first(range).col)?, c))
                        })
                        .collect(),
                });
            }
        }

        let line = match (axis, insert) {
            (Axis::Rows, true) => LineEdit::insert_rows(at + 1, count),
            (Axis::Rows, false) => LineEdit::remove_rows(at + 1, count),
            (Axis::Cols, true) => LineEdit::insert_columns(at + 1, count),
            (Axis::Cols, false) => LineEdit::remove_columns(at + 1, count),
        };
        let line = if whole {
            line
        } else {
            line.within(band.0 + 1, band.1 + 1)
        };
        let line = line.with_table_column(TABLE_COLUMN);
        self.file
            .edit_lines(&name, &line, &mut |host, text| {
                adjusted(host, text, &name, &edit)
            })
            .map_err(|e| format!("セルを移動できません（{e}）"))?;
        if insert {
            self.format_inserted(
                si,
                axis,
                at,
                count,
                if whole { None } else { Some(band) },
                from_after,
            );
        }
        self.relayout(si);
        self.names = Names::default();
        self.pending.clear();
        self.rebuild = true;
        Ok(())
    }

    /// What Excel refuses: cells with content pushed off the sheet, and array formulas, merged
    /// cells and tables cut apart.
    fn check_edit(
        &self,
        si: usize,
        r: RangeRef,
        reach: RangeRef,
        axis: Axis,
        insert: bool,
        whole: bool,
    ) -> Result<()> {
        let Some(sh) = self.model.sheet(si) else {
            return Ok(());
        };
        if insert {
            let tail = match axis {
                Axis::Rows => RangeRef::new(
                    CellRef::new(MAX_ROWS - r.height().min(MAX_ROWS), reach.start.col),
                    reach.end,
                ),
                Axis::Cols => RangeRef::new(
                    CellRef::new(reach.start.row, MAX_COLS - r.width().min(MAX_COLS)),
                    reach.end,
                ),
            };
            if sh
                .cells
                .iter_range(tail)
                .any(|(_, c)| c.formula.is_some() || !c.value.is_empty())
            {
                return Err("空でないセルをワークシートの外に押し出すことはできません。".into());
            }
        }
        let (at, end) = match axis {
            Axis::Rows => (r.start.row, r.end.row),
            Axis::Cols => (r.start.col, r.end.col),
        };
        // An area the edit takes apart: some of it moves (or goes) and some of it stays.
        let apart = |area: &RangeRef| -> bool {
            if whole {
                let (lo, hi) = match axis {
                    Axis::Rows => (area.start.row, area.end.row),
                    Axis::Cols => (area.start.col, area.end.col),
                };
                if insert {
                    at > lo && at <= hi
                } else {
                    at <= hi && end >= lo && !(at <= lo && end >= hi)
                }
            } else {
                area.intersects(&reach)
                    && (!reach.contains_range(area)
                        || (!insert && area.intersects(&r) && !r.contains_range(area)))
            }
        };
        for (at, c) in sh.cells.iter() {
            let Some(area) = c
                .formula
                .as_ref()
                .and_then(|f| f.array)
                .filter(|a| !a.is_single() && a.start == at)
            else {
                continue;
            };
            if apart(&area) {
                return Err("配列の一部を変更することはできません。".into());
            }
        }
        if !whole {
            if sh
                .merges
                .iter()
                .any(|m| m.intersects(&reach) && !reach.contains_range(m))
            {
                return Err("この操作は結合したセルには行えません。".into());
            }
            if sh
                .tables
                .iter()
                .any(|t| t.range.intersects(&reach) && !reach.contains_range(&t.range))
            {
                return Err(
                    "テーブル内のセルを移動しようとしているため、この操作は実行できません。".into(),
                );
            }
        }
        Ok(())
    }

    /// Rewrites the references of every formula and name to the cells of sheet `target` that
    /// `edit` moves.
    pub(super) fn adjust_refs(&mut self, target: &str, edit: &Edit) {
        let names: Vec<String> = self.model.sheets.iter().map(|s| s.name.clone()).collect();
        for (si, host) in names.iter().enumerate() {
            let Some(sh) = self.model.sheet_mut(si) else {
                continue;
            };
            let changes: Vec<(CellRef, gm::Formula)> = sh
                .cells
                .iter()
                .filter_map(|(at, c)| {
                    let old = c.formula.as_ref()?;
                    let e = old.with_expr(|e| adjust_if_changed(e, host, target, edit))??;
                    let mut new = gm::Formula::from_expr(e);
                    new.array = old.array;
                    new.dynamic = old.dynamic;
                    Some((at, new))
                })
                .collect();
            for (at, formula) in changes {
                if let Some(c) = sh.cells.get_mut(at) {
                    c.formula = Some(Arc::new(formula));
                }
            }
        }
        for n in &mut self.model.names {
            let host = n
                .scope
                .and_then(|s| names.get(s))
                .map(String::as_str)
                .unwrap_or("\u{0}");
            if let Ok(e) = gridcraft_formula::parse(&n.formula)
                && let Some(new) = adjust_if_changed(&e, host, target, edit)
            {
                n.formula = gridcraft_formula::print(&new);
            }
        }
        self.names = Names::default();
        self.rebuild = true;
    }

    /// Gives the cells inserted the format of the line before them (above, or to the left), or
    /// of the line after them (`from_after`, and when they are first), as Excel does.
    fn format_inserted(
        &mut self,
        si: usize,
        axis: Axis,
        at: u32,
        count: u32,
        band: Option<(u32, u32)>,
        from_after: bool,
    ) {
        let max = if axis == Axis::Rows {
            MAX_ROWS
        } else {
            MAX_COLS
        };
        let from = if from_after || at == 0 {
            at + count
        } else {
            at - 1
        };
        if from >= max {
            return;
        }
        let Ok(ws) = self.file.sheet_mut(si) else {
            return;
        };
        let inside = |other: u32| band.is_none_or(|(lo, hi)| lo <= other && other <= hi);
        // (row, column) of a line and its place along it.
        let cell = |line: u32, other: u32| {
            if axis == Axis::Rows {
                CellRef::new(line, other)
            } else {
                CellRef::new(other, line)
            }
        };
        let styles: Vec<(u32, umya_spreadsheet::Style)> = ws
            .cells()
            .iter()
            .filter_map(|c| {
                let (row, col) = (c.coordinate().row_num() - 1, c.coordinate().col_num() - 1);
                let (line, other) = if axis == Axis::Rows {
                    (row, col)
                } else {
                    (col, row)
                };
                (line == from && inside(other) && *c.style() != umya_spreadsheet::Style::default())
                    .then(|| (other, c.style().clone()))
            })
            .collect();
        for n in at..at + count {
            for (other, style) in &styles {
                ws.cell_mut(coord(cell(n, *other))).set_style(style.clone());
            }
        }
        if band.is_some() {
            return;
        }
        if axis == Axis::Cols {
            if let Some(src) = ws.column_dimension_by_number(from + 1).cloned() {
                for n in at..at + count {
                    let mut c = src.clone();
                    c.set_col_num(n + 1);
                    c.set_hidden(false);
                    *ws.column_dimension_by_number_mut(n + 1) = c;
                }
            }
        } else if let Some(src) = ws.row_dimension(from + 1) {
            let (height, custom, style) = (src.height(), src.custom_height(), src.style().clone());
            for n in at..at + count {
                let row = ws.row_dimension_mut(n + 1);
                if custom {
                    row.set_height(height);
                    row.set_custom_height(true);
                }
                if style != umya_spreadsheet::Style::default() {
                    row.set_style(style.clone());
                }
            }
        }
    }

    /// Takes what formulas depend on besides cells from the file again (merged cells, hidden rows
    /// and columns, tables), after the file was edited. The filter is the model's own.
    pub(super) fn relayout(&mut self, si: usize) {
        let Ok(ws) = self.file.sheet(si) else { return };
        let mut id = 0;
        let fresh = layout(ws, &mut id);
        let mut next = self
            .model
            .sheets
            .iter()
            .flat_map(|s| s.tables.iter().map(|t| t.id))
            .max()
            .unwrap_or(0);
        let Some(sh) = self.model.sheet_mut(si) else {
            return;
        };
        sh.merges = fresh.merges;
        sh.rows = fresh.rows;
        sh.cols = fresh.cols;
        let old = std::mem::take(&mut sh.tables);
        sh.tables = fresh
            .tables
            .into_iter()
            .map(|mut t| {
                t.id = match old.iter().find(|o| o.name.eq_ignore_ascii_case(&t.name)) {
                    Some(o) => o.id,
                    None => {
                        next += 1;
                        next
                    }
                };
                t
            })
            .collect();
    }
}
