//! Input rules (データの入力規則) worked out for the cells: whether a value is one a rule
//! allows, as Excel checks what is typed into a cell and marks the data that is not
//! ("無効データのマーク").

use gridcraft_core::{CellError, CellRef, RangeRef, Value};
use gridcraft_xlsx_formulas::fmla;
use serde_json::{Value as Json, json};

use super::Book;
use crate::patch::Validation;

/// What Excel says of a value a rule does not allow, when the rule has no message of its own.
pub const REFUSED: &str =
    "この値は、このセルに定義されているデータ入力規則の制限を満たしていません。";

/// At most this many cells of a sheet are checked for "無効データのマーク".
const CHECKED: usize = 20_000;

impl Book {
    /// The value of the formula `formula` (as the file has it, written for the cell `base`) for
    /// the cell `at`: its relative references moved with it.
    fn rule_operand(&mut self, si: usize, formula: &str, base: CellRef, at: CellRef) -> Value {
        let body = fmla::from_file(formula.trim().trim_start_matches('='));
        let Ok(e) = gridcraft_formula::parse(&body) else {
            return Value::Error(CellError::Name);
        };
        let dr = at.row as i64 - base.row as i64;
        let dc = at.col as i64 - base.col as i64;
        let e = gridcraft_formula::adjust::shift_relative_wrapping(e, dr, dc);
        self.flush();
        gridcraft_calc::recalc::evaluate_expr(&self.model, si, at, &e)
    }

    /// The rule allows the value of the cell `at` (of sheet `si`).
    pub(crate) fn allowed(&mut self, si: usize, rule: &Validation, at: CellRef) -> bool {
        let v = self.value(si, at);
        let blank = matches!(v, Value::Empty) || matches!(&v, Value::Text(s) if s.is_empty());
        if rule.kind.is_empty() || blank {
            return true;
        }
        let base = rule.ranges.first().map_or(at, |r| r.start);
        if let Value::Error(_) = v {
            return false;
        }
        let number = |x: &Value| match x {
            Value::Number(n) => Some(*n),
            _ => None,
        };
        let subject = match rule.kind.as_str() {
            "whole" => match number(&v) {
                Some(n) if n.fract() == 0.0 => n,
                _ => return false,
            },
            "decimal" | "date" | "time" => match number(&v) {
                Some(n) => n,
                None => return false,
            },
            "textLength" => match &v {
                Value::Bool(b) => {
                    if *b {
                        4.0
                    } else {
                        5.0
                    }
                }
                other => other.display().chars().count() as f64,
            },
            "list" => return self.in_list(si, rule, base, at, &v),
            "custom" => {
                return match self.rule_operand(si, &rule.formula1, base, at).scalar() {
                    Value::Bool(b) => b,
                    Value::Number(n) => n != 0.0,
                    _ => false,
                };
            }
            _ => return true,
        };
        let mut bound = |f: &str| number(&self.rule_operand(si, f, base, at).scalar());
        let Some(a) = bound(&rule.formula1) else {
            return false;
        };
        match rule.op() {
            "between" | "notBetween" => {
                let Some(b) = bound(&rule.formula2) else {
                    return false;
                };
                let inside = subject >= a && subject <= b;
                if rule.op() == "between" {
                    inside
                } else {
                    !inside
                }
            }
            "equal" => subject == a,
            "notEqual" => subject != a,
            "greaterThan" => subject > a,
            "lessThan" => subject < a,
            "greaterThanOrEqual" => subject >= a,
            "lessThanOrEqual" => subject <= a,
            _ => true,
        }
    }

    /// The value is one of the rule's list (its text without case, a number as a number).
    fn in_list(
        &mut self,
        si: usize,
        rule: &Validation,
        base: CellRef,
        at: CellRef,
        v: &Value,
    ) -> bool {
        let f = rule.formula1.trim();
        let items: Vec<Value> = if let Some(inner) = f.strip_prefix('"') {
            let inner = inner
                .strip_suffix('"')
                .unwrap_or(inner)
                .replace("\"\"", "\"");
            inner
                .split(',')
                .map(|s| Value::Text(s.trim().into()))
                .collect()
        } else {
            match self.rule_operand(si, f, base, at) {
                Value::Array(a) => a.iter().cloned().collect(),
                other => vec![other],
            }
        };
        let text = v.display();
        items.iter().any(|item| match (item, v) {
            (Value::Number(a), Value::Number(b)) => a == b,
            (Value::Empty, _) | (Value::Error(_), _) => false,
            (item, _) => item.display().to_lowercase() == text.to_lowercase(),
        })
    }

    /// The rule of the sheet (`rules`) the cell `at` has: the first that has it, as Excel
    /// keeps one rule a cell.
    fn rule_at(rules: &[Validation], at: CellRef) -> Option<&Validation> {
        rules
            .iter()
            .find(|r| r.ranges.iter().any(|x| x.contains(at)))
    }

    /// The message that stops what was entered into the cells of `range`: of the first a rule
    /// of the style "停止" with "エラー メッセージを表示する" does not allow.
    pub(crate) fn refused(
        &mut self,
        si: usize,
        rules: &[Validation],
        range: RangeRef,
    ) -> Option<String> {
        for at in range.iter().take(CHECKED) {
            let Some(rule) = Self::rule_at(rules, at) else {
                continue;
            };
            let stops = rule.show_error && matches!(rule.error_style.as_str(), "" | "stop");
            if stops && !self.allowed(si, rule, at) {
                let mut message = if rule.error.is_empty() {
                    REFUSED.to_string()
                } else {
                    rule.error.clone()
                };
                if !rule.error_title.is_empty() {
                    message = format!("{}: {message}", rule.error_title);
                }
                return Some(message);
            }
        }
        None
    }

    /// The cells of the sheet with a value its rules do not allow ("無効データのマーク"): `[r, c]`.
    pub(crate) fn invalid_cells(&mut self, si: usize, rules: &[Validation]) -> Vec<Json> {
        self.flush();
        let mut cells: Vec<CellRef> = match self.model.sheet(si) {
            Some(sh) => sh
                .cells
                .iter()
                .filter(|(_, cell)| cell.formula.is_some() || !cell.value.is_empty())
                .map(|(c, _)| c)
                .filter(|c| {
                    rules
                        .iter()
                        .any(|r| r.ranges.iter().any(|x| x.contains(*c)))
                })
                .collect(),
            None => return Vec::new(),
        };
        cells.sort_by_key(|c| (c.row, c.col));
        cells.truncate(CHECKED);
        let mut out = Vec::new();
        for at in cells {
            if let Some(rule) = Self::rule_at(rules, at) {
                if !self.allowed(si, rule, at) {
                    out.push(json!([at.row, at.col]));
                }
            }
        }
        out
    }

    /// The input rules of the sheet for the grid: each as [`Validation::to_json`] has it, with
    /// (a list) the values its button lists (`values`).
    pub(crate) fn rules_view(&mut self, si: usize, sheet: &str, rules: &[Validation]) -> Vec<Json> {
        rules
            .iter()
            .map(|r| {
                let mut j = r.to_json();
                if r.kind == "list" {
                    j["values"] = json!(self.list_values(si, sheet, &r.formula1));
                }
                j
            })
            .collect()
    }
}
