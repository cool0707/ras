//! How a sheet is shown, as the workbook keeps it (its first `sheetView`): the gridlines, the
//! zoom, the panes frozen, the cells selected; and the filter buttons of its tables.

use gridcraft_core::{CellRef, RangeRef};
use umya_spreadsheet::structs::{PaneStateValues, SheetView};
use umya_spreadsheet::{EnumTrait, Table, Worksheet};

fn view(ws: &Worksheet) -> Option<&SheetView> {
    ws.sheets_views().sheet_view_list().first()
}

/// Whether the sheet shows its gridlines.
pub(crate) fn grid_lines(ws: &Worksheet) -> bool {
    view(ws).is_none_or(SheetView::show_grid_lines)
}

/// The zoom (%) the sheet is shown at, when it is not 100.
pub(crate) fn zoom(ws: &Worksheet) -> Option<u32> {
    let z = view(ws)?.zoom_scale();
    ((10..=400).contains(&z) && z != 100).then_some(z)
}

/// The rows and columns the sheet keeps on the screen (its panes frozen): as many as are above
/// and to the left of the pane, counted from A1 (rows scrolled away above a frozen pane count
/// as frozen too).
pub(crate) fn frozen(ws: &Worksheet) -> Option<(u32, u32)> {
    let view = view(ws)?;
    let pane = view.pane()?;
    if !matches!(
        pane.state(),
        PaneStateValues::Frozen | PaneStateValues::FrozenSplit
    ) {
        return None;
    }
    let split = |v: f64| if v > 0.0 { v.round() as u32 } else { 0 };
    let (rows, cols) = (split(pane.vertical_split()), split(pane.horizontal_split()));
    if rows == 0 && cols == 0 {
        return None;
    }
    let top = CellRef::parse(&view.top_left_cell().replace('$', "")).unwrap_or(CellRef::new(0, 0));
    Some((
        if rows > 0 { top.row + rows } else { 0 },
        if cols > 0 { top.col + cols } else { 0 },
    ))
}

/// The cells selected on the sheet when it was saved, in its active pane: the active cell and
/// the range around it.
pub(crate) fn selection(ws: &Worksheet) -> Option<(CellRef, RangeRef)> {
    let view = view(ws)?;
    let pane = view
        .pane()
        .map(|p| p.active_pane().value_string().to_string());
    let sel = view.selection().iter().find(|s| match &pane {
        Some(p) => s.pane().value_string() == p,
        None => true,
    })?;
    let active = CellRef::parse(&sel.active_cell()?.to_string().replace('$', ""))?;
    let range = sel
        .sequence_of_references()
        .range_collection()
        .iter()
        .filter_map(|r| RangeRef::parse(&r.range().replace('$', "")))
        .find(|r| r.contains(active))
        .unwrap_or(RangeRef::cell(active));
    Some((active, range))
}

/// The filter of a table: its range, and the columns (from 0) whose button is hidden and those
/// it filters by.
pub(crate) fn table_filter(t: &Table) -> Option<(RangeRef, Vec<u32>, Vec<u32>)> {
    let filter = t.auto_filter();
    let range = RangeRef::parse(&filter.range().range().replace('$', ""))?;
    let cols = |pick: &dyn Fn(&umya_spreadsheet::structs::FilterColumn) -> bool| -> Vec<u32> {
        filter
            .filter_columns()
            .iter()
            .filter(|c| pick(c))
            .map(|c| c.column_id())
            .collect()
    };
    Some((
        range,
        cols(&|c| c.hidden_button()),
        cols(&|c| c.is_filtered()),
    ))
}
