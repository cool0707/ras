//! What follows the cells when they move (rows, columns and cells inserted and deleted): the
//! ranges and formulas written elsewhere in the file — merged cells, column widths, validations,
//! conditional formats, filters, hyperlinks, defined names, tables, the anchors of drawings and
//! controls, comments, the references of charts.

use std::collections::{BTreeMap, HashSet};

use gridcraft_core::{CellRef, RangeRef};
use gridcraft_formula::adjust::{Axis, Edit};

use super::scan::{self, Edits, El};
use super::{Pack, Result, WORKBOOK};
use crate::book::snapshot::Snapshot;
use crate::book::{adjusted, moved};

/// The cells of a sheet that an edit moves.
struct Shift<'a> {
    /// The name of the sheet edited.
    target: &'a str,
    edit: &'a Edit,
}

impl Shift<'_> {
    /// Where a range of the sheet is after the edit (as it was written, without `$`); None
    /// when all of it is deleted.
    fn range(&self, text: &str) -> Option<String> {
        let Some(r) = RangeRef::parse(&text.replace('$', "")) else {
            return Some(text.to_string());
        };
        let new = moved(r, self.edit)?;
        Some(if new == r {
            text.to_string()
        } else if new.is_single() && !text.contains(':') {
            new.start.a1()
        } else if new.is_full_rows() || new.is_full_cols() {
            new.a1()
        } else {
            format!("{}:{}", new.start.a1(), new.end.a1())
        })
    }

    /// A list of ranges (`A1:B2 D4`); None when all of them are deleted.
    fn sqref(&self, text: &str) -> Option<String> {
        let kept: Vec<String> = text
            .split_whitespace()
            .filter_map(|r| self.range(r))
            .collect();
        (!kept.is_empty()).then(|| kept.join(" "))
    }

    /// A formula written on the sheet `host` (None: where references name their sheet) with
    /// its references moved; None when none moves.
    fn formula(&self, host: Option<&str>, text: &str) -> Option<String> {
        let (eq, body) = match text.strip_prefix('=') {
            Some(body) => ("=", body),
            None => ("", text),
        };
        adjusted(host, body, self.target, self.edit).map(|new| format!("{eq}{new}"))
    }

    /// A row or column number after the edit, and whether it was deleted (it is then the
    /// first one after the deleted ones). Only whole rows and columns move what is anchored.
    fn line(&self, rows: bool, v: u32) -> (u32, bool) {
        match self.edit {
            Edit::Insert { axis, at, count } if (*axis == Axis::Rows) == rows => {
                (if v >= *at { v + count } else { v }, false)
            }
            Edit::Delete { axis, at, count } if (*axis == Axis::Rows) == rows => {
                if v >= at + count {
                    (v - count, false)
                } else if v >= *at {
                    (*at, true)
                } else {
                    (v, false)
                }
            }
            _ => (v, false),
        }
    }

    fn moves_lines(&self) -> bool {
        matches!(self.edit, Edit::Insert { .. } | Edit::Delete { .. })
    }
}

/// Sets the text of a leaf element.
fn set_text(el: &El, text: &str, edits: &mut Edits) {
    if el.inner.is_empty() && el.end == el.open.end {
        // `<x/>` has no place for text; it is left.
        return;
    }
    edits.replace(el.inner.clone(), scan::escape(text));
}

fn adjust_text(doc: &str, el: &El, s: &Shift, host: Option<&str>, edits: &mut Edits) {
    if el.children.is_empty()
        && !el.inner.is_empty()
        && let Some(new) = s.formula(host, &el.text(doc))
    {
        set_text(el, &new, edits);
    }
}

fn adjust_attr(doc: &str, el: &El, name: &str, s: &Shift, host: Option<&str>, edits: &mut Edits) {
    if let Some(span) = scan::attr_span(doc, el.open.clone(), name)
        && let Some(new) = s.formula(host, &scan::unescape(&doc[span.clone()]))
    {
        edits.replace(span, scan::escape(&new));
    }
}

/// Elements that go when everything in them is gone.
const CONTAINERS: &[&str] = &[
    "dataValidations",
    "hyperlinks",
    "protectedRanges",
    "ignoredErrors",
    "conditionalFormattings",
    "sparklines",
    "sparklineGroup",
    "sparklineGroups",
    "ext",
    "extLst",
];

/// What became of an element.
#[derive(Clone, Copy, PartialEq)]
enum Fate {
    Kept,
    Removed,
    /// Its range is gone: so is the element it is the range of.
    HostGone,
}

/// The markers of an anchor (`from`, `to`: column, offset, row, offset) after the edit.
/// `together`: the object moves without changing its size.
fn shift_markers(
    doc: &str,
    from: Option<&El>,
    to: Option<&El>,
    together: bool,
    s: &Shift,
    edits: &mut Edits,
) {
    if !s.moves_lines() {
        return;
    }
    let number = |el: &El| el.text(doc).trim().parse::<u32>().ok();
    let mut delta: [i64; 2] = [0, 0];
    for (i, marker) in [from, to].into_iter().enumerate() {
        let Some(marker) = marker else { continue };
        for (k, (line, offset, rows)) in [("col", "colOff", false), ("row", "rowOff", true)]
            .into_iter()
            .enumerate()
        {
            let Some(el) = marker.child(line) else {
                continue;
            };
            let Some(v) = number(el) else { continue };
            let (new, collapsed) = if i == 1 && together {
                ((v as i64 + delta[k]).max(0) as u32, false)
            } else {
                s.line(rows, v)
            };
            if i == 0 {
                delta[k] = new as i64 - v as i64;
            }
            if new != v {
                set_text(el, &new.to_string(), edits);
            }
            if collapsed && let Some(off) = marker.child(offset) {
                set_text(off, "0", edits);
            }
        }
    }
}

fn walk(doc: &str, el: &El, s: &Shift, host: &str, own: bool, edits: &mut Edits) -> Fate {
    let name = el.local();
    let truthy = |attr: &str| matches!(el.attr(doc, attr).as_deref(), Some("1") | Some("true"));
    match name {
        // Written from the cells and the merged cells of the workbook.
        "sheetData" | "mergeCells" | "dimension" | "sheetViews" | "cols" => return Fate::Kept,
        "conditionalFormatting" | "dataValidation" | "protectedRange" | "ignoredError" if own => {
            if let Some(span) = scan::attr_span(doc, el.open.clone(), "sqref") {
                let old = scan::unescape(&doc[span.clone()]);
                match s.sqref(&old) {
                    None => {
                        edits.replace(el.span(), "");
                        return Fate::Removed;
                    }
                    Some(new) if new != old => edits.replace(span, scan::escape(&new)),
                    _ => {}
                }
            }
        }
        "autoFilter" | "sortState" | "sortCondition" | "hyperlink" if own => {
            if let Some(span) = scan::attr_span(doc, el.open.clone(), "ref") {
                let old = scan::unescape(&doc[span.clone()]);
                match s.range(&old) {
                    None => {
                        edits.replace(el.span(), "");
                        return Fate::Removed;
                    }
                    Some(new) => {
                        if name == "autoFilter" {
                            filter_columns(doc, el, &old, &new, s, edits);
                        }
                        if new != old {
                            edits.replace(span, scan::escape(&new));
                        }
                    }
                }
            }
        }
        // The formulas of conditional formats and validations (`xm:f` in the extensions).
        "formula" | "formula1" | "formula2" | "f" => {
            adjust_text(doc, el, s, Some(host), edits);
            return Fate::Kept;
        }
        "sqref" if own && el.children.is_empty() => {
            let old = el.text(doc);
            match s.sqref(&old) {
                None => return Fate::HostGone,
                Some(new) if new != old => set_text(el, &new, edits),
                _ => {}
            }
            return Fate::Kept;
        }
        // The cells a control is linked to and takes its list from.
        "controlPr" | "objectPr" => {
            for attr in ["linkedCell", "listFillRange"] {
                adjust_attr(doc, el, attr, s, Some(host), edits);
            }
        }
        // Where a control lies on the sheet.
        "anchor" if own => {
            if truthy("moveWithCells") {
                shift_markers(
                    doc,
                    el.child("from"),
                    el.child("to"),
                    !truthy("sizeWithCells"),
                    s,
                    edits,
                );
            }
            return Fate::Kept;
        }
        _ => {}
    }
    let fates: Vec<Fate> = el
        .children
        .iter()
        .map(|c| walk(doc, c, s, host, own, edits))
        .collect();
    if fates.contains(&Fate::HostGone) {
        if matches!(
            name,
            "conditionalFormatting" | "dataValidation" | "sparkline"
        ) {
            edits.replace(el.span(), "");
            return Fate::Removed;
        }
        return Fate::HostGone;
    }
    let left = fates.iter().filter(|f| **f == Fate::Kept).count();
    if CONTAINERS.contains(&name) && !fates.is_empty() && left == 0 {
        edits.replace(el.span(), "");
        return Fate::Removed;
    }
    if name == "dataValidations" && left != fates.len() && el.attr(doc, "count").is_some() {
        scan::set_attr(doc, el.open.clone(), "count", &left.to_string(), edits);
    }
    Fate::Kept
}

/// The columns a filter filters by (`colId`: counted from its first column) after it moved.
fn filter_columns(doc: &str, filter: &El, old: &str, new: &str, s: &Shift, edits: &mut Edits) {
    let (Some(old), Some(new)) = (
        RangeRef::parse(&old.replace('$', "")),
        RangeRef::parse(&new.replace('$', "")),
    ) else {
        return;
    };
    for column in filter
        .children
        .iter()
        .filter(|c| c.local() == "filterColumn")
    {
        let Some(id) = column
            .attr(doc, "colId")
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };
        let head = RangeRef::cell(CellRef::new(old.start.row, old.start.col + id));
        match moved(head, s.edit).and_then(|at| at.start.col.checked_sub(new.start.col)) {
            None => edits.replace(column.span(), ""),
            Some(to) if to != id => {
                scan::set_attr(doc, column.open.clone(), "colId", &to.to_string(), edits);
            }
            _ => {}
        }
    }
}

/// The widths and formats of the columns (`<cols>`), when whole columns are inserted or
/// deleted: each column keeps its own, and the columns inserted are as the one before them.
fn columns(doc: &str, cols: &El, s: &Shift, edits: &mut Edits) {
    let (at, count, insert) = match s.edit {
        Edit::Insert {
            axis: Axis::Cols,
            at,
            count,
        } => (*at, *count, true),
        Edit::Delete {
            axis: Axis::Cols,
            at,
            count,
        } => (*at, *count, false),
        _ => return,
    };
    let p = scan::prefix(cols.name);
    // Runs of columns (0-based, inclusive) with what they are given.
    let mut runs: Vec<(u32, u32, String)> = Vec::new();
    for col in cols.children.iter().filter(|c| c.local() == "col") {
        let number = |name: &str| {
            col.attr(doc, name)
                .and_then(|v| v.parse::<u32>().ok())
                .filter(|v| *v > 0)
        };
        let (Some(min), Some(max)) = (number("min"), number("max")) else {
            continue;
        };
        let attrs = super::attrs_without(doc, col.open.clone(), &["min", "max"]);
        runs.push((min - 1, (max - 1).min(super::MAX_COLS - 1), attrs));
    }
    let mut out: Vec<(u32, u32, String)> = Vec::new();
    for (min, max, attrs) in &runs {
        let (min, max) = (*min, *max);
        if insert {
            if max < at {
                out.push((min, max, attrs.clone()));
            } else if min >= at {
                out.push((min + count, max + count, attrs.clone()));
            } else {
                out.push((min, at - 1, attrs.clone()));
                out.push((at + count, max + count, attrs.clone()));
            }
        } else {
            let end = at + count;
            if min < at {
                out.push((min, max.min(at - 1), attrs.clone()));
            }
            if max >= end {
                out.push((min.max(end) - count, max - count, attrs.clone()));
            }
        }
    }
    if insert {
        // As the column before them (the one after, for the first columns), shown.
        let from = if at == 0 { 0 } else { at - 1 };
        if let Some((_, _, attrs)) = runs
            .iter()
            .find(|(min, max, _)| (*min..=*max).contains(&from))
        {
            let attrs = attrs
                .replace(" hidden=\"1\"", "")
                .replace(" hidden=\"true\"", "");
            out.push((at, at + count - 1, attrs));
        }
    }
    out.retain(|(min, max, _)| min <= max && *min < super::MAX_COLS);
    out.sort_by_key(|r| r.0);
    // Columns side by side with the same properties are one run.
    let mut joined: Vec<(u32, u32, String)> = Vec::new();
    for run in out {
        match joined.last_mut() {
            Some(last) if last.1 + 1 == run.0 && last.2 == run.2 => last.1 = run.1,
            _ => joined.push(run),
        }
    }
    if joined.is_empty() {
        edits.replace(cols.span(), "");
        return;
    }
    let mut xml = format!("<{p}cols>");
    for (min, max, attrs) in joined {
        xml.push_str(&format!(
            "<{p}col min=\"{}\" max=\"{}\"{attrs}/>",
            min + 1,
            (max + 1).min(super::MAX_COLS)
        ));
    }
    xml.push_str(&format!("</{p}cols>"));
    edits.replace(cols.span(), xml);
}

/// The sheet's XML outside its cells, after an edit of the sheet `target`. `own`: this is that
/// sheet (its ranges move); other sheets have formulas that refer to it.
pub(super) fn sheet(
    doc: &str,
    root: &El,
    host: &str,
    own: bool,
    target: &str,
    edit: &Edit,
    edits: &mut Edits,
) {
    let s = Shift { target, edit };
    for child in &root.children {
        if own && child.local() == "cols" {
            columns(doc, child, &s, edits);
        } else {
            walk(doc, child, &s, host, own, edits);
        }
    }
}

/// The merged cells of a sheet, written from the workbook's.
pub(super) fn merges(root: &El, p: &str, merges: &[RangeRef], edits: &mut Edits) {
    let mut list: Vec<&RangeRef> = merges.iter().filter(|m| !m.is_single()).collect();
    list.sort_by_key(|m| (m.start.row, m.start.col));
    let xml = if list.is_empty() {
        String::new()
    } else {
        let mut xml = format!("<{p}mergeCells count=\"{}\">", list.len());
        for m in list {
            xml.push_str(&format!(
                "<{p}mergeCell ref=\"{}:{}\"/>",
                m.start.a1(),
                m.end.a1()
            ));
        }
        xml + &format!("</{p}mergeCells>")
    };
    if let Some(el) = root.child("mergeCells") {
        edits.replace(el.span(), xml);
        return;
    }
    if xml.is_empty() {
        return;
    }
    // Its place among the elements of a sheet: after the last of those that come before it.
    const BEFORE: &[&str] = &[
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
    ];
    if let Some(last) = root
        .children
        .iter()
        .rev()
        .find(|c| BEFORE.contains(&c.local()))
    {
        // (After the element: the cells before it may be replaced as a whole.)
        edits.replace(last.end..last.end, xml);
    }
}

/// The elements named `name` anywhere under `el`.
fn find_all<'a, 'd>(el: &'a El<'d>, name: &str, out: &mut Vec<&'a El<'d>>) {
    for c in &el.children {
        if c.local() == name {
            out.push(c);
        }
        find_all(c, name, out);
    }
}

fn store(
    replace: &mut BTreeMap<String, Vec<u8>>,
    part: &str,
    doc: &str,
    edits: Edits,
) -> Result<()> {
    if edits.is_empty() {
        return Ok(());
    }
    let new = edits.apply(doc);
    scan::tree(&new, &[]).ok_or_else(|| format!("{part} を書き換えられませんでした"))?;
    replace.insert(part.to_string(), new.into_bytes());
    Ok(())
}

/// A drawing (shapes, pictures, charts): the objects anchored to cells move and stretch with
/// them, as each is set to (`editAs`).
fn drawing(doc: &str, s: &Shift, edits: &mut Edits) {
    let Some(root) = scan::tree(doc, &[]) else {
        return;
    };
    let mut anchors = Vec::new();
    find_all(&root, "twoCellAnchor", &mut anchors);
    find_all(&root, "oneCellAnchor", &mut anchors);
    for a in anchors {
        let mode = a.attr(doc, "editAs").unwrap_or_else(|| "twoCell".into());
        if mode == "absolute" {
            continue;
        }
        shift_markers(
            doc,
            a.child("from"),
            a.child("to"),
            mode == "oneCell",
            s,
            edits,
        );
    }
}

/// A VML drawing (form controls, the boxes of notes): anchors, the cells of notes and the
/// cells controls are linked to.
fn vml(doc: &str, s: &Shift, host: &str, own: bool, edits: &mut Edits) {
    let Some(root) = scan::tree(doc, &[]) else {
        return;
    };
    let mut shapes = Vec::new();
    find_all(&root, "shape", &mut shapes);
    for shape in shapes {
        let Some(data) = shape.child("ClientData") else {
            continue;
        };
        for name in ["FmlaLink", "FmlaRange", "FmlaGroup", "FmlaTxbx"] {
            if let Some(el) = data.child(name) {
                adjust_text(doc, el, s, Some(host), edits);
            }
        }
        if !own {
            continue;
        }
        let number = |name: &str| {
            data.child(name)
                .and_then(|el| el.text(doc).trim().parse::<u32>().ok())
        };
        // A note belongs to a cell: it goes where the cell goes, and with it.
        if data.attr(doc, "ObjectType").as_deref() == Some("Note")
            && let (Some(row), Some(col)) = (number("Row"), number("Column"))
        {
            match moved(RangeRef::cell(CellRef::new(row, col)), s.edit) {
                None => {
                    edits.replace(shape.span(), "");
                    continue;
                }
                Some(to) => {
                    if to.start.row != row
                        && let Some(el) = data.child("Row")
                    {
                        set_text(el, &to.start.row.to_string(), edits);
                    }
                    if to.start.col != col
                        && let Some(el) = data.child("Column")
                    {
                        set_text(el, &to.start.col.to_string(), edits);
                    }
                }
            }
        }
        // `<x:MoveWithCells/>` says that the object does *not* move with the cells.
        let off = |name: &str| {
            data.child(name)
                .is_some_and(|el| !el.text(doc).trim().eq_ignore_ascii_case("false"))
        };
        if off("MoveWithCells") || !s.moves_lines() {
            continue;
        }
        let Some(anchor) = data.child("Anchor") else {
            continue;
        };
        let text = anchor.text(doc);
        let mut v: Vec<u32> = text
            .split(',')
            .filter_map(|n| n.trim().parse().ok())
            .collect();
        if v.len() != 8 {
            continue;
        }
        let together = off("SizeWithCells");
        // column, offset, row, offset — of the top left corner, then of the bottom right one.
        for (i, rows) in [(0, false), (2, true)] {
            let (from, gone) = s.line(rows, v[i]);
            let delta = from as i64 - v[i] as i64;
            v[i] = from;
            if gone {
                v[i + 1] = 0;
            }
            if together {
                v[i + 4] = (v[i + 4] as i64 + delta).max(0) as u32;
            } else {
                let (to, gone) = s.line(rows, v[i + 4]);
                v[i + 4] = to;
                if gone {
                    v[i + 5] = 0;
                }
            }
        }
        let new = v.iter().map(u32::to_string).collect::<Vec<_>>().join(", ");
        if new != text.trim() {
            set_text(anchor, &new, edits);
        }
    }
}

/// Elements with the cell they belong to in `ref` (comments, threaded comments).
fn refs(doc: &str, name: &str, s: &Shift, edits: &mut Edits) {
    let Some(root) = scan::tree(doc, &[]) else {
        return;
    };
    let mut list = Vec::new();
    find_all(&root, name, &mut list);
    for el in list {
        let Some(span) = scan::attr_span(doc, el.open.clone(), "ref") else {
            continue;
        };
        let old = scan::unescape(&doc[span.clone()]);
        match s.range(&old) {
            None => edits.replace(el.span(), ""),
            Some(new) if new != old => edits.replace(span, scan::escape(&new)),
            _ => {}
        }
    }
}

/// A table of the sheet: its range follows the rows; its columns cannot be cut.
fn table(doc: &str, s: &Shift, edits: &mut Edits) -> Result<()> {
    let Some(root) = scan::tree(doc, &[]) else {
        return Ok(());
    };
    let Some(old) = root
        .attr(doc, "ref")
        .and_then(|r| RangeRef::parse(&r.replace('$', "")))
    else {
        return Ok(());
    };
    let name = root.attr(doc, "displayName").unwrap_or_default();
    let refused = |what: &str| {
        Err(format!(
            "テーブル `{name}` の{what}は、ファイルの書き換えでは行えません（Excel でブックを開いて行ってください）"
        ))
    };
    let Some(new) = moved(old, s.edit) else {
        return refused("削除");
    };
    if new.width() != old.width() {
        return refused("列の挿入・削除");
    }
    let headers = root.attr(doc, "headerRowCount").is_none_or(|v| v != "0");
    if headers && moved(RangeRef::cell(old.start), s.edit).is_none() {
        return refused("見出しの行の削除");
    }
    if new.height() < 2 && headers {
        return refused("すべての行の削除");
    }
    if new != old {
        scan::set_attr(doc, root.open.clone(), "ref", &new.a1_plain(), edits);
    }
    let mut ranged = Vec::new();
    find_all(&root, "autoFilter", &mut ranged);
    find_all(&root, "sortState", &mut ranged);
    find_all(&root, "sortCondition", &mut ranged);
    for el in ranged {
        let Some(span) = scan::attr_span(doc, el.open.clone(), "ref") else {
            continue;
        };
        let old = scan::unescape(&doc[span.clone()]);
        match s.range(&old) {
            None => edits.replace(el.span(), ""),
            Some(new) if new != old => edits.replace(span, scan::escape(&new)),
            _ => {}
        }
    }
    Ok(())
}

trait Plain {
    fn a1_plain(&self) -> String;
}

impl Plain for RangeRef {
    fn a1_plain(&self) -> String {
        format!("{}:{}", self.start.a1(), self.end.a1())
    }
}

/// The parts besides the sheets that an edit of the sheet `target` concerns. `sheets`: the
/// sheets (name, part) of the workbook.
pub(super) fn parts(
    pack: &mut Pack,
    sheets: &[(String, String)],
    target: &str,
    edit: &Edit,
    replace: &mut BTreeMap<String, Vec<u8>>,
) -> Result<()> {
    let s = Shift { target, edit };
    let mut done: HashSet<String> = HashSet::new();
    for (name, part) in sheets {
        let own = name == target;
        for (_, kind, to) in pack.rels(part)? {
            let Some(to) = pack.find(&to) else { continue };
            let kind = kind.rsplit('/').next().unwrap_or("").to_string();
            if kind == "pivotTable" && own {
                return Err("ピボットテーブルのあるシートでのセル・行・列の挿入と削除は、ファイルの書き換えでは行えません（Excel でブックを開いて行ってください）".into());
            }
            let wanted = match kind.as_str() {
                "vmlDrawing" | "ctrlProp" => true,
                "drawing" | "comments" | "threadedComment" | "table" => own,
                _ => false,
            };
            if !wanted || !done.insert(to.clone()) {
                continue;
            }
            let Some(doc) = pack.text(&to)? else { continue };
            let mut edits = Edits::default();
            match kind.as_str() {
                "drawing" => drawing(&doc, &s, &mut edits),
                "vmlDrawing" => vml(&doc, &s, name, own, &mut edits),
                "comments" => refs(&doc, "comment", &s, &mut edits),
                "threadedComment" => refs(&doc, "threadedComment", &s, &mut edits),
                "table" => table(&doc, &s, &mut edits)?,
                _ => {
                    // The properties of a form control: the cells it is linked to.
                    if let Some(root) = scan::tree(&doc, &[]) {
                        for attr in ["fmlaLink", "fmlaRange", "fmlaGroup", "fmlaTxbx"] {
                            adjust_attr(&doc, &root, attr, &s, Some(name), &mut edits);
                        }
                    }
                }
            }
            store(replace, &to, &doc, edits)?;
        }
    }
    // The references of charts (always with their sheet), and the sources of pivot tables.
    let names: Vec<String> = pack.names().to_vec();
    for part in names {
        let chart = part.starts_with("xl/charts/") && part.ends_with(".xml");
        let cache = part.starts_with("xl/pivotCache/pivotCacheDefinition");
        if !(chart || cache) || part.contains("/_rels/") {
            continue;
        }
        let Some(doc) = pack.text(&part)? else {
            continue;
        };
        let Some(root) = scan::tree(&doc, &[]) else {
            continue;
        };
        let mut edits = Edits::default();
        if chart {
            let mut list = Vec::new();
            find_all(&root, "f", &mut list);
            for el in list {
                adjust_text(&doc, el, &s, None, &mut edits);
            }
        } else {
            let mut list = Vec::new();
            find_all(&root, "worksheetSource", &mut list);
            for el in list {
                if el.attr(&doc, "sheet").as_deref() != Some(target) {
                    continue;
                }
                if let Some(span) = scan::attr_span(&doc, el.open.clone(), "ref") {
                    let old = scan::unescape(&doc[span.clone()]);
                    if let Some(new) = s.range(&old).filter(|n| *n != old) {
                        edits.replace(span, scan::escape(&new));
                    }
                }
            }
        }
        store(replace, &part, &doc, edits)?;
    }
    Ok(())
}

/// The workbook part: the defined names follow the cells; Excel is told to calculate the
/// formulas when it opens the file, and to build the calculation chain anew.
pub(super) fn workbook(
    pack: &mut Pack,
    doc: &str,
    before: &Snapshot,
    after: &Snapshot,
    edit: Option<(usize, &Edit)>,
    replace: &mut BTreeMap<String, Vec<u8>>,
    remove: &mut HashSet<String>,
) -> Result<()> {
    let root = scan::tree(doc, &[]).ok_or("xl/workbook.xml を読み取れません")?;
    let p = scan::prefix(root.name);
    let mut edits = Edits::default();
    if let Some(names) = root.child("definedNames") {
        let shift = edit.map(|(si, edit)| Shift {
            target: &after.sheets[si].name,
            edit,
        });
        for name in names.children.iter().filter(|c| c.local() == "definedName") {
            let scope = name
                .attr(doc, "localSheetId")
                .and_then(|i| i.parse::<usize>().ok());
            let key = name.attr(doc, "name").unwrap_or_default();
            let find = |snapshot: &'_ Snapshot| {
                snapshot
                    .names
                    .iter()
                    .find(|n| n.1 == scope && n.0.eq_ignore_ascii_case(&key))
                    .map(|n| n.2.clone())
            };
            // A name the edit changed in the workbook (cells moved, or cut and pasted) is
            // written from it; the others (the print areas and the like) follow the cells.
            match (find(before), find(after)) {
                (Some(old), Some(new)) if old != new => set_text(name, &new, &mut edits),
                _ => {
                    if let Some(s) = &shift {
                        let host = scope
                            .and_then(|i| after.sheets.get(i))
                            .map(|sheet| sheet.name.as_str());
                        adjust_text(doc, name, s, host, &mut edits);
                    }
                }
            }
        }
    }
    // The values written are the stand-in's: Excel works the formulas out again.
    match root.child("calcPr") {
        Some(calc) => scan::set_attr(doc, calc.open.clone(), "fullCalcOnLoad", "1", &mut edits),
        None => {
            const BEFORE: &[&str] = &[
                "sheets",
                "functionGroups",
                "externalReferences",
                "definedNames",
            ];
            if let Some(last) = root
                .children
                .iter()
                .rev()
                .find(|c| BEFORE.contains(&c.local()))
            {
                edits.replace(
                    last.end..last.end,
                    format!("<{p}calcPr fullCalcOnLoad=\"1\"/>"),
                );
            }
        }
    }
    store(replace, WORKBOOK, doc, edits)?;

    // The calculation chain lists the cells with formulas: Excel builds it anew without it.
    let chain = pack
        .rels(WORKBOOK)?
        .into_iter()
        .find(|r| r.1.ends_with("/calcChain"));
    let Some((id, _, part)) = chain else {
        return Ok(());
    };
    let Some(part) = pack.find(&part) else {
        return Ok(());
    };
    remove.insert(part.clone());
    let rels_part = "xl/_rels/workbook.xml.rels";
    if let Some(rels) = pack.text(rels_part)? {
        let mut edits = Edits::default();
        for tag in scan::tags(&rels, 0).filter(|t| !t.is_close() && t.local() == "Relationship") {
            if scan::attr(&rels, tag.start..tag.end, "Id").as_deref() == Some(id.as_str()) {
                edits.replace(tag.start..tag.end, "");
            }
        }
        let name = pack.find(rels_part).unwrap_or_else(|| rels_part.into());
        store(replace, &name, &rels, edits)?;
    }
    let types_part = "[Content_Types].xml";
    if let Some(types) = pack.text(types_part)? {
        let mut edits = Edits::default();
        for tag in scan::tags(&types, 0).filter(|t| !t.is_close() && t.local() == "Override") {
            let name = scan::attr(&types, tag.start..tag.end, "PartName").unwrap_or_default();
            if name.trim_start_matches('/').eq_ignore_ascii_case(&part) {
                edits.replace(tag.start..tag.end, "");
            }
        }
        let name = pack.find(types_part).unwrap_or_else(|| types_part.into());
        store(replace, &name, &types, edits)?;
    }
    Ok(())
}
