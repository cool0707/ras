//! The color scales and data bars of a sheet's conditional formats, worked out for the grid as
//! xljs works them out for printing (`src/api/xlsx_cond.rs` there). The rules that give a
//! format (cellIs, expression, top10, ...) and icon sets are not drawn.

use std::collections::HashMap;

use umya_spreadsheet::{
    Color, ConditionalFormatValueObject, ConditionalFormattingRule, EnumTrait, Worksheet,
};

use super::*;

/// A data bar of a cell, in fractions of the cell's width.
#[derive(Debug, Clone, PartialEq)]
pub struct Bar {
    pub start: f64,
    pub len: f64,
    /// #RRGGBB.
    pub fill: String,
    pub gradient: bool,
    /// The bar grows to the left (a negative value).
    pub leftward: bool,
    /// The axis between negative and positive values, in fractions of the cell's width.
    pub axis: Option<f64>,
}

/// What the color scales and data bars give a cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CondFx {
    /// The color of a color scale, #RRGGBB.
    pub fill: Option<String>,
    pub bar: Option<Bar>,
    /// The data bar is shown without the value.
    pub hide_value: bool,
}

struct Rule<'a> {
    ranges: Vec<RangeRef>,
    priority: i32,
    rule: &'a ConditionalFormattingRule,
    /// The Excel 2010 rule of a data bar (its lengths, axis, gradient).
    ext: Option<&'a ConditionalFormattingRule>,
}

fn ranges(sqref: &str) -> Vec<RangeRef> {
    sqref
        .split_whitespace()
        .filter_map(|p| RangeRef::parse(&p.replace('$', "")))
        .collect()
}

/// The color scale and data bar rules of a sheet, in priority order.
fn rules(ws: &Worksheet) -> Vec<Rule<'_>> {
    let mut ext: Vec<Option<(&ConditionalFormattingRule, Vec<RangeRef>)>> = Vec::new();
    for cf in ws.conditional_formatting_2010_collection() {
        let rs = ranges(&cf.sequence_of_references().get_sqref());
        for r in cf.conditional_collection() {
            ext.push(Some((r, rs.clone())));
        }
    }
    let mut out = Vec::new();
    for cf in ws.conditional_formatting_collection() {
        let rs = ranges(&cf.sequence_of_references().get_sqref());
        for r in cf.conditional_collection() {
            let id = r.ext_id();
            let x = (!id.is_empty())
                .then(|| {
                    ext.iter_mut()
                        .find(|e| e.as_ref().is_some_and(|(x, _)| x.id() == id))
                        .and_then(Option::take)
                        .map(|(x, _)| x)
                })
                .flatten();
            out.push(Rule {
                ranges: rs.clone(),
                priority: r.priority(),
                rule: r,
                ext: x,
            });
        }
    }
    for (x, rs) in ext.into_iter().flatten() {
        out.push(Rule {
            ranges: rs,
            priority: x.priority(),
            rule: x,
            ext: Some(x),
        });
    }
    out.retain(|r| {
        !r.ranges.is_empty() && matches!(r.rule.get_type().value_string(), "colorScale" | "dataBar")
    });
    // Excel numbers its rules from 1; none comes last.
    out.sort_by_key(|r| {
        if r.priority == 0 {
            i32::MAX
        } else {
            r.priority
        }
    });
    out
}

/// A color that is given (not automatic, not the system foreground).
fn given(c: &Color) -> bool {
    c.has_value() && !c.automatic() && c.indexed() != 64
}

/// PERCENTILE.INC of sorted values.
fn percentile(sorted: &[f64], p: f64) -> f64 {
    if sorted.is_empty() {
        return 0.0;
    }
    let k = p.clamp(0.0, 1.0) * (sorted.len() - 1) as f64;
    let i = k.floor() as usize;
    let f = k - i as f64;
    if i + 1 < sorted.len() {
        sorted[i] + f * (sorted[i + 1] - sorted[i])
    } else {
        sorted[i]
    }
}

/// A threshold's value among the numbers of a rule (sorted); None for a formula.
fn threshold(v: &ConditionalFormatValueObject, nums: &[f64]) -> Option<f64> {
    let (min, max) = (*nums.first()?, *nums.last()?);
    let given = || v.val().trim().parse::<f64>().ok();
    Some(match v.get_type().value_string() {
        "min" => min,
        "max" => max,
        "autoMin" => min.min(0.0),
        "autoMax" => max.max(0.0),
        "percent" => min + given()? / 100.0 * (max - min),
        "percentile" => percentile(nums, given()? / 100.0),
        _ => given()?,
    })
}

fn rgb(css: &str) -> Option<(f64, f64, f64)> {
    let h = css.trim_start_matches('#');
    if h.len() != 6 {
        return None;
    }
    let p = |i: usize| u8::from_str_radix(&h[i..i + 2], 16).ok().map(f64::from);
    Some((p(0)?, p(2)?, p(4)?))
}

/// a and b mixed: t = 0 is a, 1 is b.
pub fn mix(a: &str, b: &str, t: f64) -> String {
    match (rgb(a), rgb(b)) {
        (Some(x), Some(y)) => {
            let c = |p: f64, q: f64| (p + (q - p) * t).round().clamp(0.0, 255.0) as u8;
            format!("#{:02X}{:02X}{:02X}", c(x.0, y.0), c(x.1, y.1), c(x.2, y.2))
        }
        _ => a.to_string(),
    }
}

impl Book {
    /// What the color scales and data bars of sheet `si` give its cells within `shown`.
    /// `css` turns a color into #RRGGBB.
    pub(super) fn cond_effects(
        &mut self,
        si: usize,
        shown: RangeRef,
        css: &dyn Fn(&Color) -> Option<String>,
    ) -> HashMap<(u32, u32), CondFx> {
        let mut out: HashMap<(u32, u32), CondFx> = HashMap::new();
        let used = self
            .sheet_summary(si)
            .ok()
            .and_then(|s| s.used_range)
            .unwrap_or(RangeRef::cell(CellRef::new(0, 0)));
        let Ok(ws) = self.file.sheet(si) else {
            return out;
        };
        // The rules, read first: the values are read from the book afterwards.
        struct Work {
            cells: Vec<CellRef>,
            kind: String,
            cfvo: Vec<ConditionalFormatValueObject>,
            colors: Vec<String>,
            bar: Option<(String, f64, f64, bool, String, bool)>,
        }
        let mut work = Vec::new();
        for r in rules(ws) {
            let mut cells = Vec::new();
            for range in &r.ranges {
                // Clipped to the used part of the sheet (a whole column holds no more).
                let Some(clip) = range.intersection(&used) else {
                    continue;
                };
                for row in clip.start.row..=clip.end.row {
                    for col in clip.start.col..=clip.end.col {
                        cells.push(CellRef::new(row, col));
                    }
                }
            }
            let kind = r.rule.get_type().value_string().to_string();
            let mut w = Work {
                cells,
                kind,
                cfvo: Vec::new(),
                colors: Vec::new(),
                bar: None,
            };
            if let Some(cs) = r.rule.color_scale() {
                w.cfvo = cs.cfvo_collection().to_vec();
                w.colors = cs
                    .color_collection()
                    .iter()
                    .filter(|c| given(c))
                    .filter_map(css)
                    .collect();
            }
            if let Some(db) = r.rule.data_bar() {
                let x = r.ext.and_then(|x| x.data_bar());
                let ext = x.is_some();
                let x = x.unwrap_or(db);
                w.cfvo = x.cfvo_collection().to_vec();
                let color = db
                    .color_collection()
                    .first()
                    .or_else(|| x.color_collection().first())
                    .filter(|c| given(c))
                    .and_then(css)
                    .unwrap_or_else(|| "#638EC6".into());
                // Excel 2007 bars run from 10% to 90% of the cell; Excel 2010 ones say.
                let min_len = x.min_length().map_or(10.0, f64::from) / 100.0;
                let max_len = x.max_length().map_or(90.0, f64::from) / 100.0;
                let axis = match x.axis_position() {
                    "" => "automatic".to_string(),
                    a => a.to_string(),
                };
                w.bar = Some((color, min_len, max_len, x.gradient(), axis, ext));
                if !db.show_value() {
                    w.kind = "dataBar-only".into();
                }
            }
            work.push(w);
        }
        for w in work {
            let vals: Vec<Option<f64>> = w
                .cells
                .iter()
                .map(|c| match self.value(si, *c) {
                    gridcraft_core::Value::Number(n) => Some(n),
                    _ => None,
                })
                .collect();
            let mut nums: Vec<f64> = vals.iter().flatten().copied().collect();
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            if w.kind == "colorScale" {
                let th: Option<Vec<f64>> = w.cfvo.iter().map(|v| threshold(v, &nums)).collect();
                let Some(th) = th else { continue };
                if th.len() < 2 || w.colors.len() != th.len() {
                    continue;
                }
                let k = th.len() - 1;
                for (c, x) in w.cells.iter().zip(&vals) {
                    let Some(x) = *x else { continue };
                    let color = if x <= th[0] {
                        w.colors[0].clone()
                    } else if x >= th[k] {
                        w.colors[k].clone()
                    } else {
                        let i = (0..k).find(|&i| x <= th[i + 1]).unwrap_or(k - 1);
                        let span = th[i + 1] - th[i];
                        let t = if span > 0.0 { (x - th[i]) / span } else { 0.0 };
                        mix(&w.colors[i], &w.colors[i + 1], t)
                    };
                    let fx = out.entry((c.row, c.col)).or_default();
                    if fx.fill.is_none() {
                        fx.fill = Some(color);
                    }
                }
            } else if let Some((fill, min_l, max_l, gradient, axis, ext)) = &w.bar {
                let lo = w.cfvo.first().and_then(|v| threshold(v, &nums));
                let hi = w.cfvo.get(1).and_then(|v| threshold(v, &nums));
                let (Some(lo), Some(hi)) = (lo, hi) else {
                    continue;
                };
                for (c, x) in w.cells.iter().zip(&vals) {
                    let Some(x) = *x else { continue };
                    let bar = if *ext && axis != "none" && (lo < 0.0 || axis == "middle") {
                        // Bars grow both ways from the axis (negative ones in red).
                        let (a, neg, pos) = if axis == "middle" {
                            let m = lo.abs().max(hi.abs());
                            (0.5, m, m)
                        } else {
                            let (n, p) = (-lo.min(0.0), hi.max(0.0));
                            (if n + p > 0.0 { n / (n + p) } else { 0.5 }, n, p)
                        };
                        let x = x.clamp(lo.min(0.0), hi.max(0.0));
                        if x >= 0.0 {
                            let len = if pos > 0.0 {
                                (1.0 - a) * (x / pos)
                            } else {
                                0.0
                            };
                            Bar {
                                start: a,
                                len,
                                fill: fill.clone(),
                                gradient: *gradient,
                                leftward: false,
                                axis: Some(a),
                            }
                        } else {
                            let len = if neg > 0.0 { a * (-x / neg) } else { 0.0 };
                            Bar {
                                start: a - len,
                                len,
                                fill: "#FF0000".into(),
                                gradient: *gradient,
                                leftward: true,
                                axis: Some(a),
                            }
                        }
                    } else {
                        let t = if hi > lo {
                            ((x - lo) / (hi - lo)).clamp(0.0, 1.0)
                        } else {
                            0.0
                        };
                        Bar {
                            start: 0.0,
                            len: min_l + t * (max_l - min_l),
                            fill: fill.clone(),
                            gradient: *gradient,
                            leftward: false,
                            axis: None,
                        }
                    };
                    let fx = out.entry((c.row, c.col)).or_default();
                    if fx.bar.is_none() {
                        fx.bar = Some(bar);
                        fx.hide_value |= w.kind == "dataBar-only";
                    }
                }
            }
        }
        out.retain(|(r, c), _| shown.contains(CellRef::new(*r, *c)));
        out
    }
}
