//! Conditional formats put on cells (ホーム > 条件付き書式), or cleared from them, written as
//! Excel writes them: a `<conditionalFormatting>` of the sheet with its rule (first in
//! priority: the rules there move down one), the format it gives in the styles' `<dxfs>`.
//! Clearing takes the cells out of the rules' ranges (those of Excel 2010 in the sheet's
//! `extLst` too); a rule left with no cells goes.

use gridcraft_core::{CellRef, RangeRef};

use super::Result;
use super::insert::{Sheet, place, tree};
use super::scan::{self, Edits, El};

/// A rule put on cells.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CondRule {
    /// The `type` of the rule: `cellIs`, `expression`, `containsText`, `notContainsText`,
    /// `beginsWith`, `endsWith`, `containsBlanks`, `notContainsBlanks`, `containsErrors`,
    /// `notContainsErrors`, `duplicateValues`, `uniqueValues`, `top10`, `aboveAverage`,
    /// `timePeriod`, `dataBar`, `colorScale`, `iconSet`.
    pub kind: String,
    /// `cellIs`: `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`,
    /// `greaterThanOrEqual`, `lessThanOrEqual`.
    pub operator: String,
    /// `cellIs`: one value (two for `between`, `notBetween`) as typed in Excel's dialog: a
    /// number, a text, or a formula (`=$B$1`); `expression`: the formula.
    pub values: Vec<String>,
    /// The text of `containsText`, `notContainsText`, `beginsWith`, `endsWith`.
    pub text: String,
    /// `top10`: how many (or what percent), and whether the smallest.
    pub rank: u32,
    pub bottom: bool,
    pub percent: bool,
    /// `aboveAverage`: below it, or equal to it too.
    pub below: bool,
    pub equal: bool,
    /// `timePeriod`: `today`, `yesterday`, `tomorrow`, `last7Days`, `thisWeek`, `lastWeek`,
    /// `nextWeek`, `thisMonth`, `lastMonth`, `nextMonth`.
    pub period: String,
    /// `dataBar`: its color; `colorScale`: those of the least, (the middle,) the greatest
    /// (`RRGGBB`).
    pub colors: Vec<String>,
    /// `iconSet`: the set (`3TrafficLights1`, ...), its order reversed, the icons alone.
    pub icons: String,
    pub reverse: bool,
    pub icon_only: bool,
    /// What the rule gives the cells (not for bars, scales and icons).
    pub format: Option<CondFormat>,
    /// "条件を満たす場合は停止".
    pub stop: bool,
}

/// The format a rule gives: font color, fill, bold, italic, underline, strike-through, a border
/// around the cells (`RRGGBB`).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct CondFormat {
    pub color: Option<String>,
    pub fill: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub border: Option<String>,
}

pub(super) const KINDS: &[&str] = &[
    "cellIs",
    "expression",
    "containsText",
    "notContainsText",
    "beginsWith",
    "endsWith",
    "containsBlanks",
    "notContainsBlanks",
    "containsErrors",
    "notContainsErrors",
    "duplicateValues",
    "uniqueValues",
    "top10",
    "aboveAverage",
    "timePeriod",
    "dataBar",
    "colorScale",
    "iconSet",
];

pub(super) const OPERATORS: &[&str] = &[
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
];

pub(super) const PERIODS: &[&str] = &[
    "today",
    "yesterday",
    "tomorrow",
    "last7Days",
    "thisWeek",
    "lastWeek",
    "nextWeek",
    "thisMonth",
    "lastMonth",
    "nextMonth",
];

/// The icon sets of Excel 2007 (those of 2010, the stars, triangles and boxes, are written in
/// the `extLst`), and how many icons each has.
pub(super) const ICON_SETS: &[(&str, usize)] = &[
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
];

impl CondRule {
    /// Whether the rule gives a format (a `dxf`) rather than drawing bars, colors or icons.
    fn formats(&self) -> bool {
        !matches!(self.kind.as_str(), "dataBar" | "colorScale" | "iconSet")
    }

    /// What is wrong with the rule, as the edit says it.
    pub(super) fn check(&self) -> Result<()> {
        if !KINDS.contains(&self.kind.as_str()) {
            return Err(format!(
                "条件付き書式の `type` `{}` は指定できません（{}）",
                self.kind,
                KINDS.join(" / ")
            ));
        }
        match self.kind.as_str() {
            "cellIs" => {
                if !OPERATORS.contains(&self.operator.as_str()) {
                    return Err(format!(
                        "`operator` `{}` は指定できません（{}）",
                        self.operator,
                        OPERATORS.join(" / ")
                    ));
                }
                let two = matches!(self.operator.as_str(), "between" | "notBetween");
                let n = if two { 2 } else { 1 };
                if self.values.len() != n || self.values.iter().any(|v| v.trim().is_empty()) {
                    return Err(format!("`values` には値を {n} つ指定してください"));
                }
            }
            "expression" => {
                if self.values.len() != 1
                    || self.values[0].trim().trim_start_matches('=').is_empty()
                {
                    return Err("`formula` がありません".into());
                }
            }
            "containsText" | "notContainsText" | "beginsWith" | "endsWith"
                if self.text.is_empty() =>
            {
                return Err("`text` がありません".into());
            }
            "top10" => {
                let max = if self.percent { 100 } else { 1000 };
                if !(1..=max).contains(&self.rank) {
                    return Err(format!("`rank` は 1〜{max} で指定してください"));
                }
            }
            "timePeriod" if !PERIODS.contains(&self.period.as_str()) => {
                return Err(format!(
                    "`period` `{}` は指定できません（{}）",
                    self.period,
                    PERIODS.join(" / ")
                ));
            }
            "dataBar" if self.colors.len() != 1 => {
                return Err("データ バーの `colors` は 1 色です".into());
            }
            "colorScale" if !(2..=3).contains(&self.colors.len()) => {
                return Err("カラー スケールの `colors` は 2 色か 3 色です".into());
            }
            "iconSet" if !ICON_SETS.iter().any(|(s, _)| *s == self.icons) => {
                let all: Vec<&str> = ICON_SETS.iter().map(|(s, _)| *s).collect();
                return Err(format!(
                    "アイコン セット `{}` は指定できません（{}）",
                    self.icons,
                    all.join(" / ")
                ));
            }
            _ => {}
        }
        if self.formats() && self.format.is_none() {
            return Err("`format`（ルールが付ける書式）がありません".into());
        }
        Ok(())
    }
}

/// The ranges of a `sqref` (`A1:B2 D5`).
fn ranges(sqref: &str) -> Vec<RangeRef> {
    sqref
        .split_whitespace()
        .filter_map(|r| RangeRef::parse(&r.replace('$', "")))
        .collect()
}

fn sqref(list: &[RangeRef]) -> String {
    list.iter().map(RangeRef::a1).collect::<Vec<_>>().join(" ")
}

/// `a` without the cells of `b`: up to four ranges (above, below, left and right of `b`).
fn subtract(a: RangeRef, b: &RangeRef) -> Vec<RangeRef> {
    let Some(x) = a.intersection(b) else {
        return vec![a];
    };
    let r = |r0: u32, c0: u32, r1: u32, c1: u32| {
        RangeRef::new(CellRef { row: r0, col: c0 }, CellRef { row: r1, col: c1 })
    };
    let mut out = Vec::new();
    if x.start.row > a.start.row {
        out.push(r(a.start.row, a.start.col, x.start.row - 1, a.end.col));
    }
    if x.end.row < a.end.row {
        out.push(r(x.end.row + 1, a.start.col, a.end.row, a.end.col));
    }
    if x.start.col > a.start.col {
        out.push(r(x.start.row, a.start.col, x.end.row, x.start.col - 1));
    }
    if x.end.col < a.end.col {
        out.push(r(x.start.row, x.end.col + 1, x.end.row, a.end.col));
    }
    out
}

/// The ranges of `sqref` without the cells of `clear`; None when they keep all their cells.
fn cleared(sqref: &str, clear: &[RangeRef]) -> Option<Vec<RangeRef>> {
    let list = ranges(sqref);
    if !list.iter().any(|a| clear.iter().any(|b| a.intersects(b))) {
        return None;
    }
    let mut left = list;
    for b in clear {
        left = left.into_iter().flat_map(|a| subtract(a, b)).collect();
    }
    Some(left)
}

/// A value typed into Excel's dialog as a formula of the file: `=B1` → `B1`, a number as it
/// is, a text in quotes.
fn operand(value: &str) -> String {
    let v = value.trim();
    if let Some(f) = v.strip_prefix('=') {
        return f.trim().to_string();
    }
    if v.parse::<f64>().is_ok_and(f64::is_finite) {
        return v.to_string();
    }
    format!("\"{}\"", value.replace('"', "\"\""))
}

/// The formula Excel writes for a rule of text, blanks, errors or dates, of the first cell `a`.
fn formula_of(rule: &CondRule, a: &str) -> Option<String> {
    let t = format!("\"{}\"", rule.text.replace('"', "\"\""));
    Some(match rule.kind.as_str() {
        "containsText" => format!("NOT(ISERROR(SEARCH({t},{a})))"),
        "notContainsText" => format!("ISERROR(SEARCH({t},{a}))"),
        "beginsWith" => format!("LEFT({a},LEN({t}))={t}"),
        "endsWith" => format!("RIGHT({a},LEN({t}))={t}"),
        "containsBlanks" => format!("LEN(TRIM({a}))=0"),
        "notContainsBlanks" => format!("LEN(TRIM({a}))>0"),
        "containsErrors" => format!("ISERROR({a})"),
        "notContainsErrors" => format!("NOT(ISERROR({a}))"),
        "timePeriod" => match rule.period.as_str() {
            "today" => format!("FLOOR({a},1)=TODAY()"),
            "yesterday" => format!("FLOOR({a},1)=TODAY()-1"),
            "tomorrow" => format!("FLOOR({a},1)=TODAY()+1"),
            "last7Days" => format!("AND(TODAY()-FLOOR({a},1)<=6,FLOOR({a},1)<=TODAY())"),
            "thisWeek" => format!(
                "AND(TODAY()-ROUNDDOWN({a},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({a},0)-TODAY()<=7-WEEKDAY(TODAY()))"
            ),
            "lastWeek" => format!(
                "AND(TODAY()-ROUNDDOWN({a},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({a},0)<(WEEKDAY(TODAY())+7))"
            ),
            "nextWeek" => format!(
                "AND(ROUNDDOWN({a},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({a},0)-TODAY()<(15-WEEKDAY(TODAY())))"
            ),
            "thisMonth" => format!("AND(MONTH({a})=MONTH(TODAY()),YEAR({a})=YEAR(TODAY()))"),
            "lastMonth" => format!(
                "AND(MONTH({a})=MONTH(EDATE(TODAY(),0-1)),YEAR({a})=YEAR(EDATE(TODAY(),0-1)))"
            ),
            "nextMonth" => format!(
                "AND(MONTH({a})=MONTH(EDATE(TODAY(),0+1)),YEAR({a})=YEAR(EDATE(TODAY(),0+1)))"
            ),
            _ => return None,
        },
        _ => return None,
    })
}

/// The rule as a `cfRule` of the sheet (prefix `p`), on cells whose first is `first`.
fn rule_xml(p: &str, rule: &CondRule, dxf: Option<u32>, first: CellRef) -> String {
    let a = first.a1();
    let mut attrs = format!(" type=\"{}\"", rule.kind);
    if let Some(id) = dxf {
        attrs.push_str(&format!(" dxfId=\"{id}\""));
    }
    attrs.push_str(" priority=\"1\"");
    if rule.stop {
        attrs.push_str(" stopIfTrue=\"1\"");
    }
    let mut formulas: Vec<String> = Vec::new();
    let mut inner = String::new();
    match rule.kind.as_str() {
        "cellIs" => {
            attrs.push_str(&format!(" operator=\"{}\"", rule.operator));
            formulas = rule.values.iter().map(|v| operand(v)).collect();
        }
        "expression" => {
            formulas.push(
                rule.values[0]
                    .trim()
                    .trim_start_matches('=')
                    .trim()
                    .to_string(),
            );
        }
        "containsText" | "notContainsText" | "beginsWith" | "endsWith" => {
            let op = match rule.kind.as_str() {
                "notContainsText" => "notContains",
                other => other,
            };
            attrs.push_str(&format!(
                " operator=\"{op}\" text=\"{}\"",
                scan::escape(&rule.text)
            ));
        }
        "top10" => {
            if rule.percent {
                attrs.push_str(" percent=\"1\"");
            }
            if rule.bottom {
                attrs.push_str(" bottom=\"1\"");
            }
            attrs.push_str(&format!(" rank=\"{}\"", rule.rank));
        }
        "aboveAverage" => {
            if rule.below {
                attrs.push_str(" aboveAverage=\"0\"");
            }
            if rule.equal {
                attrs.push_str(" equalAverage=\"1\"");
            }
        }
        "timePeriod" => attrs.push_str(&format!(" timePeriod=\"{}\"", rule.period)),
        "dataBar" => {
            inner = format!(
                "<{p}dataBar><{p}cfvo type=\"min\"/><{p}cfvo type=\"max\"/><{p}color rgb=\"FF{}\"/></{p}dataBar>",
                rule.colors[0]
            );
        }
        "colorScale" => {
            let mid = if rule.colors.len() == 3 {
                format!("<{p}cfvo type=\"percentile\" val=\"50\"/>")
            } else {
                String::new()
            };
            let colors: String = rule
                .colors
                .iter()
                .map(|c| format!("<{p}color rgb=\"FF{c}\"/>"))
                .collect();
            inner = format!(
                "<{p}colorScale><{p}cfvo type=\"min\"/>{mid}<{p}cfvo type=\"max\"/>{colors}</{p}colorScale>"
            );
        }
        "iconSet" => {
            let n = ICON_SETS
                .iter()
                .find(|(s, _)| *s == rule.icons)
                .map_or(3, |(_, n)| *n);
            let mut head = format!("<{p}iconSet iconSet=\"{}\"", rule.icons);
            if rule.icon_only {
                head.push_str(" showValue=\"0\"");
            }
            if rule.reverse {
                head.push_str(" reverse=\"1\"");
            }
            let cfvos: String = (0..n)
                .map(|i| format!("<{p}cfvo type=\"percent\" val=\"{}\"/>", i * 100 / n))
                .collect();
            inner = format!("{head}>{cfvos}</{p}iconSet>");
        }
        _ => {}
    }
    if let Some(f) = formula_of(rule, &a) {
        formulas.push(f);
    }
    for f in &formulas {
        inner.push_str(&format!("<{p}formula>{}</{p}formula>", scan::escape(f)));
    }
    if inner.is_empty() {
        format!("<{p}cfRule{attrs}/>")
    } else {
        format!("<{p}cfRule{attrs}>{inner}</{p}cfRule>")
    }
}

/// The format of a rule as a `dxf` of the styles.
fn dxf_xml(f: &CondFormat) -> String {
    let mut font = String::new();
    for (on, tag) in [
        (f.bold, "<b/>"),
        (f.italic, "<i/>"),
        (f.strike, "<strike/>"),
        (f.underline, "<u/>"),
    ] {
        if on {
            font.push_str(tag);
        }
    }
    if let Some(c) = &f.color {
        font.push_str(&format!("<color rgb=\"FF{c}\"/>"));
    }
    let mut out = String::from("<dxf>");
    if !font.is_empty() {
        out.push_str(&format!("<font>{font}</font>"));
    }
    if let Some(c) = &f.fill {
        out.push_str(&format!(
            "<fill><patternFill><bgColor rgb=\"FF{c}\"/></patternFill></fill>"
        ));
    }
    if let Some(c) = &f.border {
        let side = |s: &str| format!("<{s} style=\"thin\"><color rgb=\"FF{c}\"/></{s}>");
        out.push_str(&format!(
            "<border>{}{}{}{}</border>",
            side("left"),
            side("right"),
            side("top"),
            side("bottom")
        ));
    }
    out.push_str("</dxf>");
    out
}

/// The children of a style sheet, in the order the schema has them.
const STYLE_ORDER: [&str; 11] = [
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

/// The styles with the format `dxf` added to their `dxfs`, and its number.
fn with_dxf(doc: &str, dxf: &str) -> Result<(String, u32)> {
    let root = scan::tree(doc, &[]).ok_or("xl/styles.xml を読み取れません")?;
    let p = scan::prefix(root.name);
    let dxf = if p.is_empty() {
        dxf.to_string()
    } else {
        // (Every tag of the format with the prefix of the styles.)
        dxf.replace("</", "\u{0}")
            .replace('<', &format!("<{p}"))
            .replace('\u{0}', &format!("</{p}"))
    };
    let mut edits = Edits::default();
    let id = match root.child("dxfs") {
        Some(list) => {
            let n = list.children.iter().filter(|c| c.local() == "dxf").count();
            if list.open.end == list.end {
                let head = doc[list.open.clone()].trim_end_matches("/>").trim_end();
                edits.replace(list.open.clone(), format!("{head}>{dxf}</{}>", list.name));
            } else {
                edits.insert(list.inner.end, dxf);
            }
            if list.attr(doc, "count").is_some() {
                scan::set_attr(
                    doc,
                    list.open.clone(),
                    "count",
                    &(n + 1).to_string(),
                    &mut edits,
                );
            }
            n
        }
        None => {
            let rank = |n: &str| {
                STYLE_ORDER
                    .iter()
                    .position(|o| *o == n)
                    .unwrap_or(STYLE_ORDER.len())
            };
            let mine = rank("dxfs");
            let at = root
                .children
                .iter()
                .take_while(|c| rank(c.local()) < mine)
                .last()
                .map_or(root.inner.start, |c| c.end);
            edits.insert(at, format!("<{p}dxfs count=\"1\">{dxf}</{p}dxfs>"));
            0
        }
    };
    let out = edits.apply(doc);
    scan::tree(&out, &[]).ok_or("xl/styles.xml を書き換えられませんでした")?;
    Ok((out, id as u32))
}

/// The file with `rule` put on the cells `on` of the sheet `sheet` (first in priority), or,
/// `rule` None, without the rules of those cells.
pub(super) fn conditional(
    bytes: &[u8],
    sheet: &str,
    on: &[RangeRef],
    rule: Option<&CondRule>,
) -> Result<Vec<u8>> {
    let mut s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let p = scan::prefix(root.name).to_string();
    let mut edits = Edits::default();
    let Some(rule) = rule else {
        clear(&sdoc, &root, on, &mut edits)?;
        if edits.is_empty() {
            return Ok(bytes.to_vec());
        }
        let doc = edits.apply(&sdoc);
        return s.finish(Some(doc));
    };
    rule.check()?;
    let first = on.first().ok_or("条件付き書式を付ける範囲がありません")?;
    let dxf = match &rule.format {
        Some(f) if rule.formats() => {
            let styles = s
                .text("xl/styles.xml")?
                .ok_or("ブックに xl/styles.xml がありません")?;
            let (styles, id) = with_dxf(&styles, &dxf_xml(f))?;
            s.put("xl/styles.xml", styles);
            Some(id)
        }
        _ => None,
    };
    // The rules there come after the new one.
    for tag in scan::tags(&sdoc, 0).filter(|t| !t.is_close() && t.local() == "cfRule") {
        let span = tag.start..tag.end;
        if let Some(n) =
            scan::attr(&sdoc, span.clone(), "priority").and_then(|v| v.parse::<i64>().ok())
        {
            scan::set_attr(&sdoc, span, "priority", &(n + 1).to_string(), &mut edits);
        }
    }
    let el = format!(
        "<{p}conditionalFormatting sqref=\"{}\">{}</{p}conditionalFormatting>",
        sqref(on),
        rule_xml(&p, rule, dxf, first.start)
    );
    // After the sheet's other conditional formats.
    let at = root
        .children
        .iter()
        .rfind(|c| c.local() == "conditionalFormatting")
        .map_or_else(|| place(&root, "conditionalFormatting"), |c| c.end);
    edits.insert(at, el);
    let doc = edits.apply(&sdoc);
    tree(&doc, &s.part)?;
    s.finish(Some(doc))
}

/// The edits that take the cells `on` out of the conditional formats of the sheet: those of
/// its children, and those of Excel 2010 in its `extLst`.
fn clear(doc: &str, root: &El, on: &[RangeRef], edits: &mut Edits) -> Result<()> {
    for cf in root
        .children
        .iter()
        .filter(|c| c.local() == "conditionalFormatting")
    {
        let Some(refs) = cf.attr(doc, "sqref") else {
            continue;
        };
        match cleared(&refs, on) {
            None => {}
            Some(left) if left.is_empty() => edits.replace(cf.span(), ""),
            Some(left) => scan::set_attr(doc, cf.open.clone(), "sqref", &sqref(&left), edits),
        }
    }
    let Some(ext_list) = root.child("extLst") else {
        return Ok(());
    };
    let text = &doc[ext_list.span()];
    let base = ext_list.open.start;
    let list = scan::tree(text, &[]).ok_or("シートの extLst を読み取れません")?;
    let shift = |r: std::ops::Range<usize>| r.start + base..r.end + base;
    let mut exts_left = 0;
    for ext in list.children.iter().filter(|c| c.local() == "ext") {
        let Some(cfs) = ext.child("conditionalFormattings") else {
            exts_left += 1;
            continue;
        };
        let mut cfs_left = 0;
        for cf in cfs
            .children
            .iter()
            .filter(|c| c.local() == "conditionalFormatting")
        {
            let Some(refs) = cf.child("sqref") else {
                cfs_left += 1;
                continue;
            };
            match cleared(&refs.text(text), on) {
                None => cfs_left += 1,
                Some(left) if left.is_empty() => edits.replace(shift(cf.span()), ""),
                Some(left) => {
                    cfs_left += 1;
                    edits.replace(shift(refs.inner.clone()), sqref(&left));
                }
            }
        }
        if cfs_left == 0 {
            edits.replace(shift(ext.span()), "");
        } else {
            exts_left += 1;
        }
    }
    if exts_left == 0 {
        edits.replace(ext_list.span(), "");
    }
    Ok(())
}
