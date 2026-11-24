//! The conditional formats of a sheet worked out for the grid, as xljs works them out for
//! printing (`src/api/xlsx_cond.rs` there): the rules that give a format (cell values, formulas,
//! top / bottom, above / below average, duplicates, text, blanks, errors, dates), color scales,
//! data bars and icon sets. The rules are those umya-spreadsheet reads, the Excel 2010 ones
//! (x14) too, in priority order; a rule with "stop if true" stops the later ones of its cells.

use std::collections::HashMap;

use umya_spreadsheet::{
    Color, ConditionalFormatValueObject, ConditionalFormattingRule, EnumTrait, Style, Worksheet,
};

use super::*;

/// A data bar of a cell, in fractions of the cell's width.
#[derive(Debug, Clone, PartialEq)]
pub struct Bar {
    pub start: f64,
    pub len: f64,
    /// #RRGGBB.
    pub fill: String,
    pub border: Option<String>,
    pub gradient: bool,
    /// The bar grows to the left (a negative value, or right to left).
    pub leftward: bool,
    /// The axis between negative and positive values, in fractions of the cell's width, and its
    /// color (#RRGGBB; None: black).
    pub axis: Option<f64>,
    pub axis_color: Option<String>,
}

/// A border line a format gives: its style (`thin`, `double`, ...) and #RRGGBB (None: black).
#[derive(Debug, Clone, PartialEq)]
pub struct Line {
    pub style: String,
    pub color: Option<String>,
}

/// What the conditional formats give a cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CondFx {
    /// A fill (a color scale's or a format's), #RRGGBB.
    pub fill: Option<String>,
    pub color: Option<String>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<bool>,
    pub num_fmt: Option<String>,
    /// left, right, top, bottom.
    pub sides: [Option<Line>; 4],
    pub bar: Option<Bar>,
    /// An icon: its set and its index in the set, from the lowest values.
    pub icon: Option<(String, usize)>,
    /// The value is not shown (a data bar or an icon set that shows only the bar / the icon).
    pub hide_value: bool,
    stopped: bool,
}

/// A threshold of a color scale, data bar or icon set.
#[derive(Debug, Clone, Default)]
struct Cfvo {
    /// min, max, num, percent, percentile, formula, autoMin, autoMax.
    kind: String,
    val: String,
    /// The value itself reaches the threshold (icon sets).
    gte: bool,
}

#[derive(Debug, Clone, Default)]
struct BarDef {
    color: Option<Color>,
    /// Lengths in percent of the cell for the lowest and the highest values.
    min_len: f64,
    max_len: f64,
    gradient: bool,
    border: Option<Color>,
    negative: Option<Color>,
    negative_border: Option<Color>,
    negative_same: bool,
    negative_border_same: bool,
    /// automatic, middle, none.
    axis: String,
    axis_color: Option<Color>,
    rtl: bool,
    /// Written with the Excel 2010 extension (negative values grow from an axis).
    ext: bool,
}

#[derive(Debug, Clone, Default)]
struct IconDef {
    set: String,
    reverse: bool,
    /// Icons chosen one by one (set, index), from the lowest values.
    custom: Vec<(String, usize)>,
}

/// A differential format: what a rule's format changes.
#[derive(Debug, Clone, Default)]
struct Dxf {
    fill: Option<Color>,
    color: Option<Color>,
    bold: Option<bool>,
    italic: Option<bool>,
    strike: Option<bool>,
    underline: Option<bool>,
    num_fmt: Option<String>,
    /// left, right, top, bottom.
    sides: [Option<(String, Option<Color>)>; 4],
}

#[derive(Debug, Clone, Default)]
struct Rule {
    ranges: Vec<RangeRef>,
    kind: String,
    priority: i32,
    stop: bool,
    operator: String,
    text: String,
    formulas: Vec<String>,
    rank: u32,
    percent: bool,
    bottom: bool,
    above_average: bool,
    equal_average: bool,
    std_dev: u32,
    time_period: String,
    dxf: Option<Dxf>,
    cfvo: Vec<Cfvo>,
    colors: Vec<Color>,
    bar: Option<BarDef>,
    icons: Option<IconDef>,
    show_value: bool,
}

fn ranges(sqref: &str) -> Vec<RangeRef> {
    sqref
        .split_whitespace()
        .filter_map(|p| RangeRef::parse(&p.replace('$', "")))
        .collect()
}

/// A color that is given (not automatic, not the system foreground).
fn given(c: &Color) -> Option<Color> {
    (c.has_value() && !c.automatic() && c.indexed() != 64).then(|| c.clone())
}

fn cfvo(v: &ConditionalFormatValueObject) -> Cfvo {
    Cfvo {
        kind: v.get_type().value_string().to_string(),
        val: v.val().to_string(),
        gte: v.gte(),
    }
}

fn style_dxf(st: &Style) -> Dxf {
    let mut d = Dxf::default();
    if let Some(f) = st.font() {
        d.bold = f.bold().then_some(true);
        d.italic = f.italic().then_some(true);
        d.strike = f.strikethrough().then_some(true);
        d.underline = (f.font_underline().val().value_string() != "none").then_some(true);
        d.color = given(f.color());
    }
    d.num_fmt = st.numbering_format().map(|n| n.format_code().to_string());
    // A dxf's solid fill is its background color (and its pattern type is not given).
    if let Some(p) = st.fill().and_then(|f| f.pattern_fill()) {
        d.fill = p
            .background_color()
            .and_then(given)
            .or_else(|| p.foreground_color().and_then(given));
    }
    if let Some(b) = st.borders() {
        for (i, s) in [b.left(), b.right(), b.top(), b.bottom()]
            .into_iter()
            .enumerate()
        {
            let style = s.border_style();
            if !style.is_empty() && style != "none" {
                d.sides[i] = Some((style.to_string(), s.color().as_ref().and_then(given)));
            }
        }
    }
    d
}

fn read_rule(r: &ConditionalFormattingRule, x14: bool) -> Rule {
    let kind = r.get_type().value_string().to_string();
    let mut rule = Rule {
        operator: if kind == "cellIs" {
            r.operator().value_string().to_string()
        } else {
            String::new()
        },
        time_period: if kind == "timePeriod" {
            r.time_period().value_string().to_string()
        } else {
            String::new()
        },
        kind,
        // Excel numbers its rules from 1; none comes last.
        priority: match r.priority() {
            0 => i32::MAX,
            p => p,
        },
        stop: r.stop_if_true(),
        text: r.text().to_string(),
        rank: match r.rank() {
            0 => 10,
            n => n,
        },
        percent: r.percent(),
        bottom: r.bottom(),
        above_average: r.above_average(),
        equal_average: r.equal_average(),
        std_dev: r.std_dev().max(0) as u32,
        formulas: r
            .formula_collection()
            .iter()
            .map(|f| f.address_str().trim().to_string())
            .collect(),
        dxf: r.style().map(style_dxf),
        show_value: true,
        ..Default::default()
    };
    if rule.formulas.is_empty()
        && let Some(f) = r.formula()
    {
        let s = f.address_str();
        if !s.trim().is_empty() {
            rule.formulas.push(s.trim().to_string());
        }
    }
    if let Some(cs) = r.color_scale() {
        rule.cfvo = cs.cfvo_collection().iter().map(cfvo).collect();
        rule.colors = cs.color_collection().iter().filter_map(given).collect();
    }
    if let Some(db) = r.data_bar() {
        rule.cfvo = db.cfvo_collection().iter().map(cfvo).collect();
        rule.show_value = db.show_value();
        rule.bar = Some(BarDef {
            color: db
                .color_collection()
                .first()
                .and_then(given)
                .or_else(|| db.fill_color().and_then(given)),
            // Excel 2007 bars run from 10% to 90% of the cell; Excel 2010 ones say.
            min_len: db.min_length().map_or(10.0, f64::from),
            max_len: db.max_length().map_or(90.0, f64::from),
            gradient: db.gradient(),
            border: if x14 && db.border() {
                db.border_color().and_then(given)
            } else {
                None
            },
            negative: db.negative_fill_color().and_then(given),
            negative_border: db.negative_border_color().and_then(given),
            negative_same: db.negative_bar_color_same_as_positive(),
            negative_border_same: db.negative_bar_border_color_same_as_positive(),
            axis: match db.axis_position() {
                "" => "automatic".to_string(),
                a => a.to_string(),
            },
            axis_color: db.axis_color().and_then(given),
            rtl: db.direction() == "rightToLeft",
            ext: x14,
        });
    }
    if let Some(is) = r.icon_set() {
        rule.cfvo = is.cfvo_collection().iter().map(cfvo).collect();
        rule.show_value = is.show_value();
        let custom = if is.custom() {
            is.cf_icon_collection()
                .iter()
                .map(|i| (i.icon_set().to_string(), i.icon_id() as usize))
                .collect()
        } else {
            Vec::new()
        };
        rule.icons = Some(IconDef {
            set: is.icon_set().to_string(),
            reverse: is.reverse(),
            custom,
        });
    }
    rule
}

/// The rules of a sheet (the Excel 2010 ones too), in priority order.
fn sheet_rules(ws: &Worksheet) -> Vec<Rule> {
    // The Excel 2010 rules, to join the data bars written in both places by their id.
    let mut ext: Vec<Option<(&ConditionalFormattingRule, Vec<RangeRef>)>> = Vec::new();
    for cf in ws.conditional_formatting_2010_collection() {
        let rs = ranges(&cf.sequence_of_references().get_sqref());
        for r in cf.conditional_collection() {
            ext.push(Some((r, rs.clone())));
        }
    }
    let mut rules = Vec::new();
    for cf in ws.conditional_formatting_collection() {
        let rs = ranges(&cf.sequence_of_references().get_sqref());
        for r in cf.conditional_collection() {
            let mut rule = read_rule(r, false);
            rule.ranges = rs.clone();
            let id = r.ext_id();
            if !id.is_empty()
                && let Some((x, _)) = ext
                    .iter_mut()
                    .find(|e| e.as_ref().is_some_and(|(x, _)| x.id() == id))
                    .and_then(Option::take)
            {
                let more = read_rule(x, true);
                if let Some(mut bar) = more.bar {
                    // The color stays in the first rule unless the extension has one.
                    if bar.color.is_none() {
                        bar.color = rule.bar.as_ref().and_then(|b| b.color.clone());
                    }
                    rule.bar = Some(bar);
                    rule.cfvo = more.cfvo;
                }
            }
            rules.push(rule);
        }
    }
    for (x, rs) in ext.into_iter().flatten() {
        let mut rule = read_rule(x, true);
        rule.ranges = rs;
        rules.push(rule);
    }
    rules.retain(|r| !r.ranges.is_empty());
    rules.sort_by_key(|r| r.priority);
    rules
}

fn num(v: &Value) -> Option<f64> {
    match v {
        Value::Number(n) => Some(*n),
        _ => None,
    }
}

fn truthy(v: &Value) -> bool {
    match v {
        Value::Bool(b) => *b,
        Value::Number(n) => *n != 0.0,
        Value::Text(s) => s.eq_ignore_ascii_case("true"),
        _ => false,
    }
}

/// Excel's comparison of a cell's value with an operand: numbers below text below booleans,
/// text without case; an empty cell is 0 or "".
fn cmp(a: &Value, b: &Value) -> Option<std::cmp::Ordering> {
    use Value::*;
    let rank = |v: &Value| match v {
        Number(_) => 0,
        Text(_) => 1,
        Bool(_) => 2,
        _ => 3,
    };
    match (a, b) {
        (Error(_), _) | (_, Error(_)) | (Array(_), _) | (_, Array(_)) => None,
        (Empty, Empty) => Some(std::cmp::Ordering::Equal),
        (Empty, Number(_)) => cmp(&Number(0.0), b),
        (Empty, Text(_)) => cmp(&Value::Text("".into()), b),
        (Empty, Bool(_)) => cmp(&Bool(false), b),
        (_, Empty) => cmp(b, a).map(|o| o.reverse()),
        (Number(x), Number(y)) => x.partial_cmp(y),
        (Text(x), Text(y)) => Some(x.to_lowercase().cmp(&y.to_lowercase())),
        (Bool(x), Bool(y)) => Some(x.cmp(y)),
        _ => Some(rank(a).cmp(&rank(b))),
    }
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

/// Whether the date `d` (days since 1970-01-01) is in the period `period` of today's `t`.
fn in_period(d: i64, t: i64, period: &str) -> bool {
    use gridcraft_core::date::civil_from_days;
    // Weeks start on Sunday (1970-01-01 was a Thursday).
    let week = |x: i64| x - (x + 4).rem_euclid(7);
    let month = |x: i64, add: i32| {
        let (y, m, _) = civil_from_days(x);
        let n = y * 12 + m as i32 - 1 + add;
        (n.div_euclid(12), n.rem_euclid(12) as u32 + 1)
    };
    let (dy, dm, _) = civil_from_days(d);
    match period {
        "today" => d == t,
        "yesterday" => d == t - 1,
        "tomorrow" => d == t + 1,
        "last7Days" => (t - 6..=t).contains(&d),
        "thisWeek" => week(d) == week(t),
        "lastWeek" => week(d) == week(t) - 7,
        "nextWeek" => week(d) == week(t) + 7,
        "thisMonth" => (dy, dm) == month(t, 0),
        "lastMonth" => (dy, dm) == month(t, -1),
        "nextMonth" => (dy, dm) == month(t, 1),
        _ => false,
    }
}

enum Eff {
    Dxf,
    Fill(String),
    Bar(Bar),
    Icon(Option<(String, usize)>),
}

impl Book {
    /// What the conditional formats of sheet `si` give its cells within `shown`. `css` turns a
    /// color into #RRGGBB.
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
        let rules = match self.file.sheet(si) {
            Ok(ws) => sheet_rules(ws),
            Err(_) => return out,
        };
        let mut values: HashMap<CellRef, Value> = HashMap::new();
        for rule in &rules {
            // The cells of the rule, clipped to the used part of the sheet (and to 200,000).
            let mut cells: Vec<CellRef> = Vec::new();
            let mut seen = std::collections::HashSet::new();
            for range in &rule.ranges {
                let Some(clip) = range.intersection(&used) else {
                    continue;
                };
                for at in clip.iter() {
                    if cells.len() >= 200_000 {
                        break;
                    }
                    if seen.insert(at) {
                        cells.push(at);
                    }
                }
            }
            if cells.is_empty() {
                continue;
            }
            let vals: Vec<Value> = cells
                .iter()
                .map(|at| {
                    values
                        .entry(*at)
                        .or_insert_with(|| self.value(si, *at).scalar())
                        .clone()
                })
                .collect();
            let mut nums: Vec<f64> = vals.iter().filter_map(num).collect();
            nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
            let base = rule.ranges[0].start;
            for (at, eff) in self.hits(si, rule, &cells, &vals, &nums, base, css) {
                let fx = out.entry((at.row, at.col)).or_default();
                if fx.stopped {
                    continue;
                }
                match eff {
                    Eff::Dxf => {
                        if let Some(d) = &rule.dxf {
                            add_dxf(fx, d, css);
                        }
                    }
                    Eff::Fill(c) => {
                        if fx.fill.is_none() {
                            fx.fill = Some(c);
                        }
                    }
                    // (A bar and an icon together, as Excel draws them: the icon over the bar,
                    // both over a color scale's fill; of each, the first rule's.)
                    Eff::Bar(bar) => {
                        if fx.bar.is_none() {
                            fx.bar = Some(bar);
                            fx.hide_value |= !rule.show_value;
                        }
                    }
                    Eff::Icon(i) => {
                        if fx.icon.is_none() {
                            fx.icon = i;
                            fx.hide_value |= !rule.show_value;
                        }
                    }
                }
                if rule.stop {
                    fx.stopped = true;
                }
            }
        }
        out.retain(|(r, c), f| {
            shown.contains(CellRef::new(*r, *c))
                && *f
                    != CondFx {
                        stopped: f.stopped,
                        ..Default::default()
                    }
        });
        out
    }

    /// A threshold's value among the numbers of a rule (sorted).
    fn threshold(&mut self, si: usize, v: &Cfvo, nums: &[f64]) -> Option<f64> {
        let (min, max) = (*nums.first()?, *nums.last()?);
        let mut given = || -> Option<f64> {
            v.val
                .trim()
                .parse::<f64>()
                .ok()
                .or_else(|| num(&self.evaluate(si, &v.val).scalar()))
        };
        Some(match v.kind.as_str() {
            "min" => min,
            "max" => max,
            "autoMin" => min.min(0.0),
            "autoMax" => max.max(0.0),
            "percent" => min + given()? / 100.0 * (max - min),
            "percentile" => percentile(nums, given()? / 100.0),
            _ => given()?,
        })
    }

    #[allow(clippy::too_many_arguments)]
    fn hits(
        &mut self,
        si: usize,
        rule: &Rule,
        cells: &[CellRef],
        vals: &[Value],
        nums: &[f64],
        base: CellRef,
        css: &dyn Fn(&Color) -> Option<String>,
    ) -> Vec<(CellRef, Eff)> {
        let mut out = Vec::new();
        match rule.kind.as_str() {
            "colorScale" => {
                let th: Option<Vec<f64>> = rule
                    .cfvo
                    .iter()
                    .map(|v| self.threshold(si, v, nums))
                    .collect();
                let colors: Vec<String> = rule.colors.iter().filter_map(css).collect();
                let Some(th) = th else { return out };
                if th.len() < 2 || colors.len() != th.len() {
                    return out;
                }
                let k = th.len() - 1;
                for (at, v) in cells.iter().zip(vals) {
                    let Some(x) = num(v) else { continue };
                    let c = if x <= th[0] {
                        colors[0].clone()
                    } else if x >= th[k] {
                        colors[k].clone()
                    } else {
                        let i = (0..k).find(|&i| x <= th[i + 1]).unwrap_or(k - 1);
                        let span = th[i + 1] - th[i];
                        let t = if span > 0.0 { (x - th[i]) / span } else { 0.0 };
                        mix(&colors[i], &colors[i + 1], t)
                    };
                    out.push((*at, Eff::Fill(c)));
                }
            }
            "dataBar" => {
                let Some(def) = &rule.bar else { return out };
                let lo = rule.cfvo.first().and_then(|v| self.threshold(si, v, nums));
                let hi = rule.cfvo.get(1).and_then(|v| self.threshold(si, v, nums));
                let (Some(lo), Some(hi)) = (lo, hi) else {
                    return out;
                };
                let fill = def
                    .color
                    .as_ref()
                    .and_then(css)
                    .unwrap_or_else(|| "#638EC6".into());
                let neg_fill = if def.negative_same {
                    fill.clone()
                } else {
                    def.negative
                        .as_ref()
                        .and_then(css)
                        .unwrap_or_else(|| "#FF0000".into())
                };
                let border = def.border.as_ref().and_then(css);
                let neg_border = if def.negative_border_same {
                    border.clone()
                } else {
                    def.negative_border
                        .as_ref()
                        .and_then(css)
                        .or(border.clone())
                };
                let (min_l, max_l) = (def.min_len / 100.0, def.max_len / 100.0);
                let axis_color = def.axis_color.as_ref().and_then(css);
                for (at, v) in cells.iter().zip(vals) {
                    let Some(x) = num(v) else { continue };
                    let mut bar =
                        if def.ext && def.axis != "none" && (lo < 0.0 || def.axis == "middle") {
                            // Bars grow both ways from the axis.
                            let (a, neg, pos) = if def.axis == "middle" {
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
                                    border: border.clone(),
                                    gradient: def.gradient,
                                    leftward: false,
                                    axis: Some(a),
                                    axis_color: axis_color.clone(),
                                }
                            } else {
                                let len = if neg > 0.0 { a * (-x / neg) } else { 0.0 };
                                Bar {
                                    start: a - len,
                                    len,
                                    fill: neg_fill.clone(),
                                    border: neg_border.clone(),
                                    gradient: def.gradient,
                                    leftward: true,
                                    axis: Some(a),
                                    axis_color: axis_color.clone(),
                                }
                            }
                        } else {
                            let t = if hi > lo {
                                ((x - lo) / (hi - lo)).clamp(0.0, 1.0)
                            } else {
                                0.0
                            };
                            // (Without an axis, a negative value's bar the same way, in its colors.)
                            let negative = def.ext && x < 0.0;
                            Bar {
                                start: 0.0,
                                len: min_l + t * (max_l - min_l),
                                fill: if negative { &neg_fill } else { &fill }.clone(),
                                border: if negative { &neg_border } else { &border }.clone(),
                                gradient: def.gradient,
                                leftward: false,
                                axis: None,
                                axis_color: None,
                            }
                        };
                    if def.rtl {
                        bar.start = 1.0 - bar.start - bar.len;
                        bar.leftward = !bar.leftward;
                        bar.axis = bar.axis.map(|a| 1.0 - a);
                    }
                    out.push((*at, Eff::Bar(bar)));
                }
            }
            "iconSet" => {
                let Some(def) = &rule.icons else { return out };
                let th: Vec<Option<f64>> = rule
                    .cfvo
                    .iter()
                    .map(|v| self.threshold(si, v, nums))
                    .collect();
                let n = th.len();
                if n == 0 {
                    return out;
                }
                for (at, v) in cells.iter().zip(vals) {
                    let Some(x) = num(v) else { continue };
                    let mut idx = 0;
                    for i in (1..n).rev() {
                        if let Some(t) = th[i]
                            && (if rule.cfvo[i].gte { x >= t } else { x > t })
                        {
                            idx = i;
                            break;
                        }
                    }
                    if def.reverse {
                        idx = n - 1 - idx;
                    }
                    let icon = match def.custom.get(idx) {
                        Some((set, _)) if set == "NoIcons" => None,
                        Some((set, i)) => Some((set.clone(), *i)),
                        None if def.set == "NoIcons" => None,
                        None => Some((def.set.clone(), idx)),
                    };
                    out.push((*at, Eff::Icon(icon)));
                }
            }
            _ => {
                // The rules that give the format of the rule.
                let formulas: Vec<Option<Expr>> = rule
                    .formulas
                    .iter()
                    .map(|f| {
                        gridcraft_formula::parse(&fmla::from_file(f.trim_start_matches('='))).ok()
                    })
                    .collect();
                let avg = if nums.is_empty() {
                    0.0
                } else {
                    nums.iter().sum::<f64>() / nums.len() as f64
                };
                let sd = if nums.len() > 1 {
                    (nums.iter().map(|x| (x - avg).powi(2)).sum::<f64>() / (nums.len() - 1) as f64)
                        .sqrt()
                } else {
                    0.0
                };
                let cut = (rule.kind == "top10" && !nums.is_empty()).then(|| {
                    let k = if rule.percent {
                        ((nums.len() as f64 * rule.rank as f64 / 100.0).floor() as usize).max(1)
                    } else {
                        rule.rank.max(1) as usize
                    };
                    let k = k.min(nums.len());
                    if rule.bottom {
                        nums[k - 1]
                    } else {
                        nums[nums.len() - k]
                    }
                });
                let key = |v: &Value| match v {
                    Value::Number(n) => Some(format!("n{n}")),
                    Value::Text(s) if !s.is_empty() => Some(format!("s{}", s.to_lowercase())),
                    Value::Bool(b) => Some(format!("b{b}")),
                    _ => None,
                };
                let mut counts: HashMap<String, usize> = HashMap::new();
                if matches!(rule.kind.as_str(), "duplicateValues" | "uniqueValues") {
                    for v in vals {
                        if let Some(k) = key(v) {
                            *counts.entry(k).or_default() += 1;
                        }
                    }
                }
                let needle = rule.text.to_lowercase();
                let today = (gridcraft_calc::now_serial().floor() as i64) - 25569;
                let date1904 = self.model.date_system == DateSystem::D1904;
                for (at, v) in cells.iter().zip(vals) {
                    let mut operand = |i: usize| -> Value {
                        match formulas.get(i) {
                            Some(Some(e)) => {
                                let dr = at.row as i64 - base.row as i64;
                                let dc = at.col as i64 - base.col as i64;
                                let e = gridcraft_formula::adjust::shift_relative_wrapping(
                                    e.clone(),
                                    dr,
                                    dc,
                                );
                                self.flush();
                                gridcraft_calc::recalc::evaluate_expr(&self.model, si, *at, &e)
                                    .scalar()
                            }
                            _ => Value::Error(CellError::Value),
                        }
                    };
                    let text = || match v {
                        Value::Empty | Value::Error(_) => String::new(),
                        other => other.display().to_lowercase(),
                    };
                    let blank = matches!(v, Value::Empty)
                        || matches!(v, Value::Text(s) if s.trim().is_empty());
                    let hit = match rule.kind.as_str() {
                        "cellIs" => {
                            use std::cmp::Ordering::*;
                            let a = operand(0);
                            let o = cmp(v, &a);
                            match rule.operator.as_str() {
                                "lessThan" => o == Some(Less),
                                "lessThanOrEqual" => matches!(o, Some(Less | Equal)),
                                "equal" => o == Some(Equal),
                                "notEqual" => o.is_some_and(|o| o != Equal),
                                "greaterThanOrEqual" => matches!(o, Some(Greater | Equal)),
                                "greaterThan" => o == Some(Greater),
                                "between" | "notBetween" => {
                                    let z = operand(1);
                                    // The two bounds in either order.
                                    let (lo, hi) = if cmp(&a, &z) == Some(Greater) {
                                        (z, a)
                                    } else {
                                        (a, z)
                                    };
                                    let inside = matches!(cmp(v, &lo), Some(Greater | Equal))
                                        && matches!(cmp(v, &hi), Some(Less | Equal));
                                    if rule.operator == "between" {
                                        inside
                                    } else {
                                        cmp(v, &lo).is_some() && !inside
                                    }
                                }
                                _ => false,
                            }
                        }
                        "expression" => truthy(&operand(0)),
                        "containsText" => text().contains(&needle),
                        "notContainsText" => !text().contains(&needle),
                        "beginsWith" => text().starts_with(&needle),
                        "endsWith" => text().ends_with(&needle),
                        "containsBlanks" => blank,
                        "notContainsBlanks" => !blank,
                        "containsErrors" => matches!(v, Value::Error(_)),
                        "notContainsErrors" => !matches!(v, Value::Error(_)),
                        "top10" => match (num(v), cut) {
                            (Some(x), Some(k)) => {
                                if rule.bottom {
                                    x <= k
                                } else {
                                    x >= k
                                }
                            }
                            _ => false,
                        },
                        "aboveAverage" => match num(v) {
                            Some(x) if !nums.is_empty() => {
                                let d = rule.std_dev as f64 * sd;
                                match (rule.above_average, rule.equal_average) {
                                    (true, false) => x > avg + d,
                                    (true, true) => x >= avg + d,
                                    (false, false) => x < avg - d,
                                    (false, true) => x <= avg - d,
                                }
                            }
                            _ => false,
                        },
                        "duplicateValues" => {
                            key(v).is_some_and(|k| counts.get(&k).copied().unwrap_or(0) > 1)
                        }
                        "uniqueValues" => {
                            key(v).is_some_and(|k| counts.get(&k).copied().unwrap_or(0) == 1)
                        }
                        "timePeriod" => num(v).is_some_and(|x| {
                            // Days since 1970-01-01 (serial 25569 in the 1900 system).
                            let d = x.floor() as i64 + if date1904 { 1462 } else { 0 } - 25569;
                            in_period(d, today, &rule.time_period)
                        }),
                        _ => false,
                    };
                    if hit {
                        out.push((*at, Eff::Dxf));
                    }
                }
            }
        }
        out
    }
}

fn add_dxf(fx: &mut CondFx, d: &Dxf, css: &dyn Fn(&Color) -> Option<String>) {
    if fx.fill.is_none() {
        fx.fill = d.fill.as_ref().and_then(css);
    }
    if fx.color.is_none() {
        fx.color = d.color.as_ref().and_then(css);
    }
    fx.bold = fx.bold.or(d.bold);
    fx.italic = fx.italic.or(d.italic);
    fx.strike = fx.strike.or(d.strike);
    fx.underline = fx.underline.or(d.underline);
    if fx.num_fmt.is_none() {
        fx.num_fmt = d.num_fmt.clone();
    }
    for i in 0..4 {
        if fx.sides[i].is_none() {
            fx.sides[i] = d.sides[i].as_ref().map(|(style, c)| Line {
                style: style.clone(),
                color: c.as_ref().and_then(css),
            });
        }
    }
}

// ---- icons (as Excel 2021 / 365 draws them, 16 x 16 px)

/// The colors of Excel's icons: (fill, edge).
type Paint = (&'static str, &'static str);
const RED: Paint = ("#D65532", "#9E381C");
const YELLOW: Paint = ("#EAC282", "#A4802B");
const GREEN: Paint = ("#68A490", "#326F5B");
const GRAY: Paint = ("#808080", "#646262");
const DARK: Paint = ("#505050", "#343434");
/// Of "赤から黒".
const LIGHT: Paint = ("#B1B1B1", "#787878");
const PINK: Paint = ("#EDB2A3", "#C98E80");
const RED_2: Paint = ("#D86344", "#B14629");
/// The parts of the ratings and boxes, on and off.
const BLUE: Paint = ("#4A7DB1", "#335F8C");
const OFF: Paint = ("#B3B3B3", "#727272");

/// A shape (`d`) filled, with its edge.
fn shape(d: &str, (fill, edge): Paint) -> String {
    format!(
        "<path d=\"{d}\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\"1\" stroke-linejoin=\"round\"/>"
    )
}

/// An arrow pointing `deg` from up, 13 x 16 px (a slanted one larger, with a shorter shaft: it
/// fills the corners of the box).
fn arrow(deg: i32, paint: Paint) -> String {
    let (fill, edge) = paint;
    let (scale, tail) = if deg % 90 == 0 {
        ("", "15.5")
    } else {
        (" translate(8 8) scale(1.1) translate(-8 -8)", "14")
    };
    format!(
        "<path transform=\"rotate({deg} 8 8){scale}\" d=\"M8 .5L14.5 7.5H11V{tail}H5V7.5H1.5Z\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\"1\" stroke-linejoin=\"miter\"/>"
    )
}

fn circle((fill, edge): Paint) -> String {
    format!(
        "<circle cx=\"8\" cy=\"8\" r=\"7\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\".8\"/>"
    )
}

/// A flag on its pole.
fn flag(paint: Paint) -> String {
    format!(
        "<path d=\"M3.5 1V15.5\" stroke=\"#727272\" stroke-width=\"1.6\"/>{}",
        shape("M4.5 1.5L14 5.75L4.5 10Z", paint)
    )
}

/// A traffic light in its black box.
fn lamp((fill, edge): Paint) -> String {
    format!(
        "<rect x=\".5\" y=\".5\" width=\"15\" height=\"15\" rx=\"3\" fill=\"#505050\" stroke=\"#303030\" stroke-width=\"1\"/><circle cx=\"8\" cy=\"8\" r=\"5\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\".6\"/>"
    )
}

/// ✖, !, ✔ (`kind` 'x', '!', 'v'): in white on a circle, or of the color with its edge.
fn mark(kind: char, paint: Paint, circled: bool) -> String {
    let (fill, edge) = paint;
    if circled {
        let body = match kind {
            'v' => "<path d=\"M4.6 8.4L7 10.7L11.4 5.6\" fill=\"none\" stroke=\"#fff\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"/>".to_string(),
            '!' => "<path d=\"M8 3.6V8.8\" stroke=\"#fff\" stroke-width=\"2\" stroke-linecap=\"round\"/><circle cx=\"8\" cy=\"11.9\" r=\"1.2\" fill=\"#fff\"/>".to_string(),
            _ => "<path d=\"M5.3 5.3L10.7 10.7M10.7 5.3L5.3 10.7\" stroke=\"#fff\" stroke-width=\"2\" stroke-linecap=\"round\"/>".to_string(),
        };
        return format!("{}{body}", circle(paint));
    }
    // A thick line with its edge: the edge drawn under, wider.
    let line = |d: &str, cap: &str| {
        format!(
            "<path d=\"{d}\" fill=\"none\" stroke=\"{edge}\" stroke-width=\"4.4\" stroke-linecap=\"{cap}\" stroke-linejoin=\"miter\"/><path d=\"{d}\" fill=\"none\" stroke=\"{fill}\" stroke-width=\"2.6\" stroke-linecap=\"{cap}\" stroke-linejoin=\"miter\"/>"
        )
    };
    match kind {
        'v' => line("M2.4 8.6L6.2 12.4L13.8 3.8", "butt"),
        '!' => format!(
            "{}{}",
            line("M8 1.2V10.2", "butt"),
            shape("M6.3 11.8H9.7V15.2H6.3Z", paint)
        ),
        _ => line("M2.6 2.6L13.4 13.4M13.4 2.6L2.6 13.4", "butt"),
    }
}

fn bars(on: usize) -> String {
    (0..4)
        .map(|i| {
            let h = 6.0 + i as f64 * 3.0;
            let (fill, edge) = if i < on { BLUE } else { OFF };
            format!(
                "<rect x=\"{}\" y=\"{}\" width=\"3\" height=\"{h}\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\".8\"/>",
                1.0 + i as f64 * 3.6,
                15.4 - h
            )
        })
        .collect()
}

fn quarters(q: usize) -> String {
    let c = "#505050";
    let ring =
        "<circle cx=\"8\" cy=\"8\" r=\"7\" fill=\"#fff\" stroke=\"#545454\" stroke-width=\"1\"/>";
    let fill = match q {
        0 => String::new(),
        4 => format!(
            "<circle cx=\"8\" cy=\"8\" r=\"7\" fill=\"{c}\" stroke=\"#343434\" stroke-width=\"1\"/>"
        ),
        _ => {
            let a = q as f64 * std::f64::consts::FRAC_PI_2;
            let (x, y) = (8.0 + 7.0 * a.sin(), 8.0 - 7.0 * a.cos());
            format!(
                "<path d=\"M8 8V1A7 7 0 {} 1 {x:.2} {y:.2}Z\" fill=\"{c}\"/>",
                if q > 2 { 1 } else { 0 }
            )
        }
    };
    ring.to_string() + &fill
}

/// An empty star, half filled (the left half), full.
fn star(fill: f64) -> String {
    let d = "M8 1.2L10 5.8L15 6.2L11.2 9.4L12.4 14.4L8 11.7L3.6 14.4L4.8 9.4L1 6.2L6 5.8Z";
    let empty = ("#fff", "#727272");
    match fill {
        f if f >= 1.0 => shape(d, YELLOW),
        f if f > 0.0 => {
            // (Each half with its own edge: no line down the middle.)
            let half = |fill: &str, outline: &str, (f, e): Paint| {
                format!(
                    "<path d=\"{fill}\" fill=\"{f}\"/><path d=\"{outline}\" fill=\"none\" stroke=\"{e}\" stroke-width=\"1\" stroke-linejoin=\"round\"/>"
                )
            };
            half(
                "M8 1.2L10 5.8L15 6.2L11.2 9.4L12.4 14.4L8 11.7Z",
                "M8 1.2L10 5.8L15 6.2L11.2 9.4L12.4 14.4L8 11.7",
                empty,
            ) + &half(
                "M8 1.2L8 11.7L3.6 14.4L4.8 9.4L1 6.2L6 5.8Z",
                "M8 1.2L6 5.8L1 6.2L4.8 9.4L3.6 14.4L8 11.7",
                YELLOW,
            )
        }
        _ => shape(d, empty),
    }
}

/// Four boxes, `on` of them blue: those of the bottom row first.
fn boxes(on: usize) -> String {
    [(1.0, 8.5), (8.5, 8.5), (1.0, 1.0), (8.5, 1.0)]
        .iter()
        .enumerate()
        .map(|(i, (x, y))| {
            let (fill, edge) = if i < on { BLUE } else { OFF };
            format!(
                "<rect x=\"{x}\" y=\"{y}\" width=\"6.5\" height=\"6.5\" rx=\"1\" fill=\"{fill}\" stroke=\"{edge}\" stroke-width=\".8\"/>"
            )
        })
        .collect()
}

/// The icon sets and how many icons each has: those of Excel 2007, then those of Excel 2010.
pub const ICON_SETS: &[(&str, usize)] = &[
    ("3Arrows", 3),
    ("3ArrowsGray", 3),
    ("3Flags", 3),
    ("3TrafficLights1", 3),
    ("3TrafficLights2", 3),
    ("3Signs", 3),
    ("3Symbols", 3),
    ("3Symbols2", 3),
    ("4Arrows", 4),
    ("4ArrowsGray", 4),
    ("4RedToBlack", 4),
    ("4Rating", 4),
    ("4TrafficLights", 4),
    ("5Arrows", 5),
    ("5ArrowsGray", 5),
    ("5Rating", 5),
    ("5Quarters", 5),
    ("3Stars", 3),
    ("3Triangles", 3),
    ("5Boxes", 5),
];

/// The icon of a set (its index from the lowest values) as SVG; None for an unknown set.
pub fn icon_svg(set: &str, i: usize) -> Option<String> {
    let pick = |v: &[String]| v.get(i).cloned();
    let body = match set {
        "3Arrows" => pick(&[arrow(180, RED), arrow(90, YELLOW), arrow(0, GREEN)]),
        "3ArrowsGray" => pick(&[arrow(180, GRAY), arrow(90, GRAY), arrow(0, GRAY)]),
        "4Arrows" => pick(&[
            arrow(180, RED),
            arrow(135, YELLOW),
            arrow(45, YELLOW),
            arrow(0, GREEN),
        ]),
        "4ArrowsGray" => pick(&[
            arrow(180, GRAY),
            arrow(135, GRAY),
            arrow(45, GRAY),
            arrow(0, GRAY),
        ]),
        "5Arrows" => pick(&[
            arrow(180, RED),
            arrow(135, YELLOW),
            arrow(90, YELLOW),
            arrow(45, YELLOW),
            arrow(0, GREEN),
        ]),
        "5ArrowsGray" => pick(&[
            arrow(180, GRAY),
            arrow(135, GRAY),
            arrow(90, GRAY),
            arrow(45, GRAY),
            arrow(0, GRAY),
        ]),
        "3Flags" => pick(&[flag(RED), flag(YELLOW), flag(GREEN)]),
        "3TrafficLights1" => pick(&[circle(RED), circle(YELLOW), circle(GREEN)]),
        "3TrafficLights2" => pick(&[lamp(RED), lamp(YELLOW), lamp(GREEN)]),
        "4TrafficLights" => pick(&[circle(DARK), circle(RED), circle(YELLOW), circle(GREEN)]),
        "4RedToBlack" => pick(&[circle(DARK), circle(LIGHT), circle(PINK), circle(RED_2)]),
        "3Signs" => pick(&[
            shape("M8 .9L15.1 8L8 15.1L.9 8Z", RED),
            shape("M8 1.4L15 14.6H1Z", YELLOW),
            circle(GREEN),
        ]),
        "3Symbols" => pick(&[
            mark('x', RED, true),
            mark('!', YELLOW, true),
            mark('v', GREEN, true),
        ]),
        "3Symbols2" => pick(&[
            mark('x', RED_2, false),
            mark('!', YELLOW, false),
            mark('v', GREEN, false),
        ]),
        "4Rating" => pick(&[bars(1), bars(2), bars(3), bars(4)]),
        "5Rating" => pick(&[bars(0), bars(1), bars(2), bars(3), bars(4)]),
        "5Quarters" => pick(&[
            quarters(0),
            quarters(1),
            quarters(2),
            quarters(3),
            quarters(4),
        ]),
        "3Stars" => pick(&[star(0.0), star(0.5), star(1.0)]),
        "3Triangles" => pick(&[
            shape("M.6 4H15.4L8 12.4Z", RED),
            shape("M.6 5.6H15.4V10.4H.6Z", YELLOW),
            shape("M.6 12H15.4L8 3.6Z", GREEN),
        ]),
        "5Boxes" => pick(&[boxes(0), boxes(1), boxes(2), boxes(3), boxes(4)]),
        _ => None,
    }?;
    Some(format!(
        "<svg viewBox=\"0 0 16 16\" width=\"100%\" height=\"100%\">{body}</svg>"
    ))
}
