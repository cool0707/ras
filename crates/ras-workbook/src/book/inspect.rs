//! What a workbook holds, read for `ras book dump` / `ras book diff`: the sheets, their cells
//! (value, formula, number format, what the cell shows) and what else is on them (merged cells,
//! the filter, tables, input rules, conditional formats), and the workbook's names.

use std::collections::BTreeMap;

use umya_spreadsheet::EnumTrait;

use super::*;

/// A cell with content.
#[derive(Clone, Debug, PartialEq)]
pub struct CellInfo {
    pub at: CellRef,
    pub value: Val,
    /// The formula as dynamic-array Excel shows it (`=A1*2`).
    pub formula: Option<String>,
    /// The formula in R1C1 form: the same for formulas filled down or across.
    pub r1c1: Option<String>,
    pub number_format: String,
    /// What the cell shows.
    pub text: String,
}

/// Cells in a row or a column that hold the same formula (the same in R1C1 form).
#[derive(Clone, Debug, PartialEq)]
pub struct FormulaRegion {
    pub range: RangeRef,
    /// The formula of the first cell.
    pub formula: String,
    pub r1c1: String,
    pub count: usize,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TableInfo {
    pub name: String,
    pub range: String,
    pub columns: Vec<String>,
    pub totals_row: bool,
}

/// An input rule (data validation) or a conditional format.
#[derive(Clone, Debug, PartialEq)]
pub struct RuleInfo {
    pub ranges: String,
    /// `list`, `whole`, ... / `cellIs`, `expression`, `colorScale`, ...
    pub kind: String,
    pub operator: String,
    pub formulas: Vec<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SheetInfo {
    pub index: usize,
    pub name: String,
    pub code_name: Option<String>,
    /// -1 visible, 0 hidden, 2 very hidden.
    pub visible: i32,
    pub used_range: Option<RangeRef>,
    pub cell_count: usize,
    pub formula_count: usize,
    pub merged: Vec<String>,
    pub filter: Option<String>,
    pub tables: Vec<TableInfo>,
    pub validations: Vec<RuleInfo>,
    pub conditional: Vec<RuleInfo>,
}

fn sqref(s: &umya_spreadsheet::SequenceOfReferences) -> String {
    s.range_collection()
        .iter()
        .map(|r| r.range())
        .collect::<Vec<_>>()
        .join(",")
}

impl Book {
    /// The cells of `r` with content, row by row.
    pub fn cells(&mut self, si: usize, r: RangeRef) -> Result<Vec<CellInfo>> {
        self.check_sheet(si)?;
        self.flush();
        let mut at: Vec<CellRef> = Vec::new();
        if let Some(sh) = self.model.sheet(si) {
            for (c, cell) in sh.cells.iter() {
                if r.contains(c) && (cell.formula.is_some() || !cell.value.is_empty()) {
                    at.push(c);
                }
            }
            for (c, v) in sh.spill.iter() {
                if r.contains(*c) && !v.is_empty() {
                    at.push(*c);
                }
            }
        }
        at.sort_by_key(|c| (c.row, c.col));
        at.dedup();
        let mut out = Vec::with_capacity(at.len());
        for c in at {
            let has_formula = self
                .model
                .sheet(si)
                .and_then(|s| s.cell(c))
                .is_some_and(|x| x.formula.is_some());
            let text_of = |v: Val| match v {
                Val::Text(s) => s,
                _ => String::new(),
            };
            let value = self.read(si, c, Read::Value);
            let (formula, r1c1) = if has_formula {
                (
                    Some(text_of(self.read(si, c, Read::Formula2))),
                    Some(text_of(self.read(si, c, Read::FormulaR1C1))),
                )
            } else {
                (None, None)
            };
            let number_format = self.number_format_code(si, c);
            let text = text_of(self.read(si, c, Read::Text));
            out.push(CellInfo {
                at: c,
                value,
                formula,
                r1c1,
                number_format,
                text,
            });
        }
        Ok(out)
    }

    /// A serial date as text: `2026-04-01`, with the time when it has one.
    pub fn date_text(&mut self, serial: f64) -> String {
        let code = if serial.fract() == 0.0 {
            "yyyy-mm-dd"
        } else if serial < 1.0 {
            "hh:mm:ss"
        } else {
            "yyyy-mm-dd hh:mm:ss"
        };
        self.display(&Value::Number(serial), code)
    }

    /// The formulas of a sheet, those filled down or across written once.
    pub fn formula_regions(&mut self, si: usize) -> Result<Vec<FormulaRegion>> {
        let all = RangeRef::new(CellRef::new(0, 0), CellRef::new(MAX_ROWS - 1, MAX_COLS - 1));
        let cells: Vec<CellInfo> = self
            .cells(si, all)?
            .into_iter()
            .filter(|c| c.formula.is_some())
            .collect();
        let by_at: BTreeMap<(u32, u32), &CellInfo> =
            cells.iter().map(|c| ((c.at.row, c.at.col), c)).collect();
        let mut taken: std::collections::HashSet<(u32, u32)> = std::collections::HashSet::new();
        let mut out = Vec::new();
        for c in &cells {
            if taken.contains(&(c.at.row, c.at.col)) {
                continue;
            }
            let same = |row: u32, col: u32, taken: &std::collections::HashSet<(u32, u32)>| {
                !taken.contains(&(row, col))
                    && by_at.get(&(row, col)).is_some_and(|o| o.r1c1 == c.r1c1)
            };
            let mut down = c.at.row;
            while same(down + 1, c.at.col, &taken) {
                down += 1;
            }
            let mut right = c.at.col;
            while same(c.at.row, right + 1, &taken) {
                right += 1;
            }
            // The longer run; a block filled both ways comes out as its columns.
            let end = if down - c.at.row >= right - c.at.col {
                CellRef::new(down, c.at.col)
            } else {
                CellRef::new(c.at.row, right)
            };
            for row in c.at.row..=end.row {
                for col in c.at.col..=end.col {
                    taken.insert((row, col));
                }
            }
            let range = RangeRef::new(c.at, end);
            out.push(FormulaRegion {
                range,
                formula: c.formula.clone().unwrap_or_default(),
                r1c1: c.r1c1.clone().unwrap_or_default(),
                count: range.count() as usize,
            });
        }
        Ok(out)
    }

    /// The sheet and what is on it besides the cells.
    pub fn sheet_summary(&mut self, si: usize) -> Result<SheetInfo> {
        self.check_sheet(si)?;
        self.flush();
        let used_range = self.used_range(si);
        let (mut cell_count, mut formula_count) = (0, 0);
        if let Some(sh) = self.model.sheet(si) {
            for (_, cell) in sh.cells.iter() {
                if cell.formula.is_some() {
                    formula_count += 1;
                    cell_count += 1;
                } else if !cell.value.is_empty() {
                    cell_count += 1;
                }
            }
        }
        let mut info = SheetInfo {
            index: si,
            name: self.sheet_name(si).unwrap_or_default().to_string(),
            code_name: self.code_name(si),
            visible: self.visible(si),
            used_range,
            cell_count,
            formula_count,
            merged: Vec::new(),
            filter: None,
            tables: Vec::new(),
            validations: Vec::new(),
            conditional: Vec::new(),
        };
        let Ok(ws) = self.file.sheet(si) else {
            return Ok(info);
        };
        info.merged = ws.merge_cells().iter().map(|r| r.range()).collect();
        info.filter = ws.auto_filter().map(|f| f.range().range());
        for t in ws.tables() {
            let (a, b) = t.area();
            let name = if t.display_name().is_empty() {
                t.name()
            } else {
                t.display_name()
            };
            info.tables.push(TableInfo {
                name: name.to_string(),
                range: format!(
                    "{}:{}",
                    a.to_string().replace('$', ""),
                    b.to_string().replace('$', "")
                ),
                columns: t.columns().iter().map(|c| c.name().to_string()).collect(),
                totals_row: t.totals_row_shown(),
            });
        }
        if let Some(dv) = ws.data_validations() {
            for v in dv.data_validation_list() {
                let kind = match v.get_type().value_string() {
                    "iso_8859_8_i" => "none",
                    k => k,
                };
                let mut formulas = Vec::new();
                for f in [v.formula1(), v.formula2()] {
                    if !f.is_empty() {
                        formulas.push(f.to_string());
                    }
                }
                // `between` is what the file leaves out; it means nothing for a list.
                let operator = if matches!(kind, "list" | "custom" | "none") {
                    String::new()
                } else {
                    v.operator().value_string().to_string()
                };
                info.validations.push(RuleInfo {
                    ranges: sqref(v.sequence_of_references()),
                    kind: kind.to_string(),
                    operator,
                    formulas,
                });
            }
        }
        for cf in ws.conditional_formatting_collection() {
            let ranges = sqref(cf.sequence_of_references());
            for rule in cf.conditional_collection() {
                let kind = rule.get_type().value_string().to_string();
                let mut formulas: Vec<String> = rule
                    .formula_collection()
                    .iter()
                    .map(|f| f.address_str())
                    .filter(|s| !s.is_empty())
                    .collect();
                if formulas.is_empty() {
                    if let Some(f) = rule.formula() {
                        let s = f.address_str();
                        if !s.is_empty() {
                            formulas.push(s);
                        }
                    }
                }
                if !rule.text().is_empty() {
                    formulas.push(format!("\"{}\"", rule.text()));
                }
                let operator = if kind == "cellIs" {
                    rule.operator().value_string().to_string()
                } else {
                    String::new()
                };
                info.conditional.push(RuleInfo {
                    ranges: ranges.clone(),
                    kind,
                    operator,
                    formulas,
                });
            }
        }
        Ok(info)
    }

    /// The sheet a name, a code name or a 1-based number means.
    pub fn find_sheet(&self, key: &str) -> Option<usize> {
        if let Some(i) = self.sheet_index(key) {
            return Some(i);
        }
        let n = self.sheet_count();
        if let Some(i) = (0..n).find(|&i| {
            self.code_name(i)
                .is_some_and(|c| c.eq_ignore_ascii_case(key))
        }) {
            return Some(i);
        }
        key.parse::<usize>()
            .ok()
            .filter(|i| (1..=n).contains(i))
            .map(|i| i - 1)
    }
}
