//! What the cells of the workbook hold, cell by cell, for writing an edit back into the file it
//! was read from (`crate::patch`): two snapshots — the workbook as it was read and the workbook
//! after the edit — tell which cells of the file are written again.

use std::collections::HashMap;

use gridcraft_core::{CellRef, RangeRef};
use umya_spreadsheet::{CellFormula, CellFormulaValues, CellRawValue, Style};

use gridcraft_model as gm;

use super::format::{DEFAULT_COLUMN_WIDTH, width_to_file};
use super::{Book, at_of};

/// The value a cell is saved with.
#[derive(Clone, Debug, PartialEq)]
pub enum Raw {
    Empty,
    Num(f64),
    Bool(bool),
    /// An error value as the file writes it (`#N/A`).
    Err(String),
    Text(String),
}

/// The formula of a cell as the file has it (no leading `=`).
#[derive(Clone, Debug, PartialEq)]
pub struct Formula {
    pub text: String,
    /// The cells of an array formula (`A1:B2`).
    pub area: Option<String>,
    /// A dynamic-array formula.
    pub dynamic: bool,
    /// A formula that is not text (a data table): its cell is never written again.
    pub opaque: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellState {
    pub formula: Option<Formula>,
    pub value: Raw,
    /// Its format: an index into [`Styles`].
    pub style: usize,
}

/// The formats met so far, each once.
#[derive(Default)]
pub struct Styles {
    list: Vec<Style>,
}

impl Styles {
    pub fn intern(&mut self, style: &Style) -> usize {
        match self.list.iter().position(|s| s == style) {
            Some(i) => i,
            None => {
                self.list.push(style.clone());
                self.list.len() - 1
            }
        }
    }

    /// The number format of a format (`General` when it has none).
    pub fn number_format(&self, id: usize) -> String {
        self.list
            .get(id)
            .and_then(|s| s.number_format())
            .map(|f| f.format_code().to_string())
            .filter(|c| !c.is_empty())
            .unwrap_or_else(|| "General".into())
    }

    /// The format of cells that are given none.
    pub fn plain(&mut self) -> usize {
        self.intern(&Style::default())
    }

    pub fn get(&self, id: usize) -> Option<&Style> {
        self.list.get(id)
    }

    /// Whether two formats differ in nothing but their number formats.
    pub fn same_but_number_format(&self, a: usize, b: usize) -> bool {
        let (Some(a), Some(b)) = (self.list.get(a), self.list.get(b)) else {
            return false;
        };
        let (mut a, mut b) = (a.clone(), b.clone());
        a.remove_number_format();
        b.remove_number_format();
        a == b
    }
}

/// What a whole column is given.
#[derive(Clone, Debug, PartialEq)]
pub struct ColState {
    /// Its width as the file has it.
    pub width: f64,
    /// It has a width of its own (not the one columns have when none is set).
    pub sized: bool,
    pub hidden: bool,
    pub style: usize,
}

/// What a whole row is given.
#[derive(Clone, Debug, PartialEq)]
pub struct RowState {
    /// Its height in points (0: the sheet's default). Without `custom`, the height Excel fit
    /// the row to (its tallest font), kept as the file has it.
    pub height: f64,
    pub custom: bool,
    pub hidden: bool,
    pub style: usize,
}

/// The filter of a sheet.
#[derive(Clone, Debug, PartialEq)]
pub struct FilterState {
    pub range: RangeRef,
    /// The columns it filters by (counted from its first), each with its criterion as the
    /// file writes it (the elements inside `filterColumn`); None for one that is kept as the
    /// file has it.
    pub columns: Vec<(u32, Option<String>)>,
}

#[derive(Default)]
pub struct SheetState {
    pub name: String,
    pub cells: HashMap<CellRef, CellState>,
    pub merges: Vec<RangeRef>,
    /// The columns and rows that are given something (by 0-based column and row).
    pub cols: HashMap<u32, ColState>,
    pub rows: HashMap<u32, RowState>,
    /// The width of the other columns as the file has it (0: none given), and the height of
    /// the other rows.
    pub default_width: f64,
    pub default_height: f64,
    pub filter: Option<FilterState>,
}

fn attr(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// A criterion of a filter as the file writes it.
fn criterion_xml(c: &gm::FilterCriterion) -> Option<String> {
    Some(match c {
        gm::FilterCriterion::Values { values, blanks } => {
            let mut xml = String::from(if *blanks {
                "<filters blank=\"1\">"
            } else {
                "<filters>"
            });
            for v in values {
                xml.push_str(&format!("<filter val=\"{}\"/>", attr(v)));
            }
            xml + "</filters>"
        }
        gm::FilterCriterion::Custom { a, b, and } => {
            let test = |(op, val): &(String, String)| {
                let operator = match op.as_str() {
                    "<" => " operator=\"lessThan\"",
                    "<=" => " operator=\"lessThanOrEqual\"",
                    "<>" => " operator=\"notEqual\"",
                    ">=" => " operator=\"greaterThanOrEqual\"",
                    ">" => " operator=\"greaterThan\"",
                    _ => "",
                };
                format!("<customFilter{operator} val=\"{}\"/>", attr(val))
            };
            format!(
                "<customFilters{}>{}{}</customFilters>",
                if *and { " and=\"1\"" } else { "" },
                test(a),
                b.as_ref().map(test).unwrap_or_default()
            )
        }
        gm::FilterCriterion::Top10 {
            bottom,
            percent,
            count,
        } => format!(
            "<top10{}{} val=\"{count}\"/>",
            if *bottom { " top=\"0\"" } else { "" },
            if *percent { " percent=\"1\"" } else { "" }
        ),
        _ => return None,
    })
}

#[derive(Default)]
pub struct Snapshot {
    pub sheets: Vec<SheetState>,
    /// The defined names: (name, the sheet of a sheet's own name, formula as the file writes it).
    pub names: Vec<(String, Option<usize>, String)>,
}

fn formula_of(f: &CellFormula) -> Formula {
    let text = f.text().trim();
    let text = text.strip_prefix('=').unwrap_or(text).to_string();
    let array = *f.formula_type() == CellFormulaValues::Array;
    Formula {
        opaque: *f.formula_type() == CellFormulaValues::DataTable || text.is_empty(),
        area: array
            .then(|| f.reference().replace('$', ""))
            .filter(|r| !r.is_empty()),
        dynamic: f.dynamic_array(),
        text,
    }
}

impl Book {
    /// The cells as `save` would write them. `overwritten`: the cells (sheet, range) that were
    /// written to since the workbook was read — a formula of the file there that the engine does
    /// not hold is gone; elsewhere such a formula (one the engine could not work out as Excel
    /// does, kept by its saved value) is still the cell's.
    pub fn snapshot(&mut self, styles: &mut Styles, overwritten: &[(usize, RangeRef)]) -> Snapshot {
        self.flush();
        let mut kept: HashMap<(usize, CellRef), Formula> = HashMap::new();
        for (si, ws) in self.file.sheet_collection().iter().enumerate() {
            let sheet = self.model.sheet(si);
            for c in ws.cells() {
                let Some(f) = c.formula_obj() else { continue };
                let at = at_of(c);
                if overwritten.iter().any(|(s, r)| *s == si && r.contains(at))
                    || sheet
                        .and_then(|s| s.cell(at))
                        .is_some_and(|m| m.formula.is_some())
                {
                    continue;
                }
                kept.insert((si, at), formula_of(f));
            }
        }
        self.store();
        let plain = Style::default();
        let mut out = Snapshot {
            names: self
                .model
                .names
                .iter()
                .map(|n| {
                    (
                        n.name.clone(),
                        n.scope,
                        gridcraft_xlsx_formulas::fmla::text_to_file(&n.formula),
                    )
                })
                .collect(),
            ..Default::default()
        };
        for (si, ws) in self.file.sheet_collection().iter().enumerate() {
            let mut sheet = SheetState {
                name: ws.name().to_string(),
                ..Default::default()
            };
            for c in ws.cells() {
                let at = at_of(c);
                let formula = match c.formula_obj() {
                    Some(f) => Some(formula_of(f)),
                    None => kept.remove(&(si, at)),
                };
                let value = match c.raw_value() {
                    CellRawValue::Empty => Raw::Empty,
                    CellRawValue::Numeric(n) => Raw::Num(*n),
                    CellRawValue::Bool(b) => Raw::Bool(*b),
                    CellRawValue::Error(e) => Raw::Err(e.to_string()),
                    other => {
                        let s = other.to_string();
                        if s.is_empty() && formula.is_none() {
                            Raw::Empty
                        } else {
                            Raw::Text(s)
                        }
                    }
                };
                sheet.cells.insert(
                    at,
                    CellState {
                        formula,
                        value,
                        style: styles.intern(c.style()),
                    },
                );
            }
            sheet.merges = ws
                .merge_cells()
                .iter()
                .filter_map(|m| RangeRef::parse(&m.range().replace('$', "")))
                .collect();
            let format = ws.sheet_format_properties();
            let default = format.default_column_width();
            sheet.default_width = if default > 0.0 {
                width_to_file(default)
            } else {
                0.0
            };
            sheet.default_height = format.default_row_height();
            for c in ws.column_dimensions() {
                // (A column with nothing of its own is one the workbook only looked at.)
                let untouched = default <= 0.0
                    && c.width() == DEFAULT_COLUMN_WIDTH
                    && !c.hidden()
                    && !c.best_fit()
                    && *c.style() == plain;
                if c.col_num() > 0 && !untouched {
                    sheet.cols.insert(
                        c.col_num() - 1,
                        ColState {
                            width: width_to_file(c.width()),
                            sized: default > 0.0 || c.width() != DEFAULT_COLUMN_WIDTH,
                            hidden: c.hidden(),
                            style: styles.intern(c.style()),
                        },
                    );
                }
            }
            for r in ws.row_dimensions() {
                let state = RowState {
                    height: r.height(),
                    custom: r.custom_height(),
                    hidden: r.hidden(),
                    style: styles.intern(r.style()),
                };
                if r.row_num() > 0
                    && (state.custom || state.height > 0.0 || state.hidden || *r.style() != plain)
                {
                    sheet.rows.insert(r.row_num() - 1, state);
                }
            }
            sheet.filter = self
                .model
                .sheet(si)
                .and_then(|s| s.autofilter.as_ref())
                .map(|f| FilterState {
                    range: f.range,
                    columns: f
                        .criteria
                        .iter()
                        .map(|(offset, c)| (*offset, criterion_xml(c)))
                        .collect(),
                });
            out.sheets.push(sheet);
        }
        // Formulas of the file whose cells the engine holds nothing for (no saved value).
        for ((si, at), formula) in kept {
            if let Some(sheet) = out.sheets.get_mut(si) {
                sheet.cells.insert(
                    at,
                    CellState {
                        formula: Some(formula),
                        value: Raw::Empty,
                        style: styles.intern(&self.style_at(si, at)),
                    },
                );
            }
        }
        out
    }
}
