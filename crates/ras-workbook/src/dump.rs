//! `ras book dump`: a workbook as text, for people and AI agents working without Excel.
//!
//! Without a sheet it is an outline (the sheets, what is on them, the formulas with those
//! filled down written once, the names); with a sheet it is the cells of a range. The ways
//! large sheets are kept short (a limit on rows and columns with the range that was written,
//! formulas of the same shape as one region) follow what xlflow does for AI agents.

use serde_json::{Value as Json, json};

use crate::book::{Book, CellInfo, RuleInfo, SheetInfo, Val};
use crate::{CellRef, RangeRef};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Format {
    Text,
    Json,
    Csv,
    Markdown,
}

#[derive(Clone, Debug)]
pub struct Options {
    pub format: Format,
    /// Number formats and what the cells show, besides values and formulas.
    pub formats: bool,
    /// The rows and columns written at most (None: all).
    pub max_rows: Option<u32>,
    pub max_cols: Option<u32>,
}

impl Default for Options {
    fn default() -> Self {
        Options { format: Format::Text, formats: false, max_rows: Some(200), max_cols: Some(50) }
    }
}

/// The formula regions of a sheet the outline lists at most.
const MAX_REGIONS: usize = 40;

pub fn type_name(v: &Val) -> &'static str {
    match v {
        Val::Empty => "empty",
        Val::Num(_) => "number",
        Val::Text(_) => "string",
        Val::Bool(_) => "boolean",
        Val::Date(_) => "date",
        Val::Err(_) => "error",
    }
}

fn number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 { format!("{}", n as i64) } else { format!("{n}") }
}

/// A value as it is written in the text form: `"text"`, `12.5`, `TRUE`, `2026-04-01`, `#N/A`.
pub fn literal(book: &mut Book, c: &CellInfo) -> String {
    match &c.value {
        Val::Empty => String::new(),
        Val::Num(n) => number(*n),
        Val::Text(s) => serde_json::to_string(s).unwrap_or_default(),
        Val::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
        Val::Date(d) => book.date_text(*d),
        Val::Err(_) => c.text.clone(),
    }
}

fn json_value(book: &mut Book, c: &CellInfo) -> Json {
    match &c.value {
        Val::Empty => Json::Null,
        Val::Num(n) => json!(n),
        Val::Text(s) => json!(s),
        Val::Bool(b) => json!(b),
        Val::Date(d) => json!(book.date_text(*d)),
        Val::Err(_) => json!(c.text),
    }
}

pub fn visible_name(v: i32) -> &'static str {
    match v {
        0 => "hidden",
        2 => "veryHidden",
        _ => "visible",
    }
}

fn rule_text(r: &RuleInfo, kinds: &[(&str, &str)]) -> String {
    let kind = kinds.iter().find(|(k, _)| *k == r.kind).map(|(_, v)| *v).unwrap_or(r.kind.as_str());
    let mut s = format!("{} {kind}", r.ranges);
    if !r.operator.is_empty() {
        s.push_str(&format!(" {}", r.operator));
    }
    if !r.formulas.is_empty() {
        s.push_str(&format!(" {}", r.formulas.join(" / ")));
    }
    s
}

const VALIDATION_KINDS: &[(&str, &str)] = &[("list", "リスト"), ("whole", "整数"), ("decimal", "小数点数"), ("date", "日付"), ("time", "時刻"), ("textLength", "文字列の長さ"), ("custom", "数式")];
const CONDITIONAL_KINDS: &[(&str, &str)] =
    &[("cellIs", "セルの値"), ("expression", "数式"), ("colorScale", "カラー スケール"), ("dataBar", "データ バー"), ("iconSet", "アイコン セット"), ("top10", "上位/下位"), ("duplicateValues", "重複する値"), ("uniqueValues", "一意の値")];

fn rule_json(r: &RuleInfo) -> Json {
    let mut o = json!({ "ranges": r.ranges, "type": r.kind });
    if !r.operator.is_empty() {
        o["operator"] = json!(r.operator);
    }
    if !r.formulas.is_empty() {
        o["formulas"] = json!(r.formulas);
    }
    o
}

fn sheet_json(s: &SheetInfo) -> Json {
    json!({
        "index": s.index + 1,
        "name": s.name,
        "codeName": s.code_name,
        "visible": visible_name(s.visible),
        "usedRange": s.used_range.map(|r| r.a1()),
        "cellCount": s.cell_count,
        "formulaCount": s.formula_count,
        "mergedCells": s.merged,
        "autoFilter": s.filter,
        "tables": s.tables.iter().map(|t| json!({ "name": t.name, "range": t.range, "columns": t.columns, "totalsRow": t.totals_row })).collect::<Vec<_>>(),
        "validations": s.validations.iter().map(rule_json).collect::<Vec<_>>(),
        "conditionalFormats": s.conditional.iter().map(rule_json).collect::<Vec<_>>(),
    })
}

/// The lines about what is on a sheet besides its cells.
fn sheet_extras(s: &SheetInfo, indent: &str, out: &mut String) {
    if !s.merged.is_empty() {
        out.push_str(&format!("{indent}結合セル: {}\n", s.merged.join(", ")));
    }
    if let Some(f) = &s.filter {
        out.push_str(&format!("{indent}フィルター: {f}\n"));
    }
    for t in &s.tables {
        out.push_str(&format!("{indent}テーブル: {} {}（列: {}{}）\n", t.name, t.range, t.columns.join(", "), if t.totals_row { "、集計行あり" } else { "" }));
    }
    for v in &s.validations {
        out.push_str(&format!("{indent}入力規則: {}\n", rule_text(v, VALIDATION_KINDS)));
    }
    for c in &s.conditional {
        out.push_str(&format!("{indent}条件付き書式: {}\n", rule_text(c, CONDITIONAL_KINDS)));
    }
}

fn sheet_title(s: &SheetInfo) -> String {
    let mut t = s.name.clone();
    let mut notes = Vec::new();
    if let Some(c) = &s.code_name {
        notes.push(format!("コード名 {c}"));
    }
    match s.visible {
        0 => notes.push("非表示".into()),
        2 => notes.push("非表示（VeryHidden）".into()),
        _ => {}
    }
    if !notes.is_empty() {
        t.push_str(&format!("（{}）", notes.join("、")));
    }
    t
}

/// The outline of a workbook: its sheets and names.
pub fn outline(book: &mut Book, file: &str, opts: &Options) -> Result<String, String> {
    let n = book.sheet_count();
    let mut sheets = Vec::new();
    for si in 0..n {
        let info = book.sheet_summary(si)?;
        let regions = book.formula_regions(si)?;
        sheets.push((info, regions));
    }
    let names = book.names();
    let scope = |book: &Book, s: Option<usize>| s.and_then(|i| book.sheet_name(i).map(str::to_string));
    if opts.format == Format::Json {
        let sheets: Vec<Json> = sheets
            .iter()
            .map(|(s, regions)| {
                let mut o = sheet_json(s);
                o["formulas"] = regions.iter().map(|r| json!({ "range": r.range.a1(), "formula": r.formula, "r1c1": r.r1c1, "count": r.count })).collect();
                o
            })
            .collect();
        let names: Vec<Json> = names.iter().map(|(name, s, f)| json!({ "name": name, "scope": scope(book, *s), "refersTo": format!("={f}") })).collect();
        return Ok(format!("{}\n", serde_json::to_string_pretty(&json!({ "file": file, "sheets": sheets, "names": names })).unwrap_or_default()));
    }
    let mut out = String::new();
    let md = opts.format == Format::Markdown;
    out.push_str(&if md { format!("# {file}\n\n") } else { format!("{file}\n") });
    out.push_str(if md { "## シート\n\n" } else { "シート:\n" });
    for (s, regions) in &sheets {
        let used = match s.used_range {
            Some(r) => format!("使用範囲 {}、値のあるセル {} 個（うち数式 {} 個）", r.a1(), s.cell_count, s.formula_count),
            None => "空".to_string(),
        };
        let (head, indent) = if md { (format!("- **{}** {used}\n", sheet_title(s)), "  - ") } else { (format!("  {}. {}  {used}\n", s.index + 1, sheet_title(s)), "     ") };
        out.push_str(&head);
        sheet_extras(s, indent, &mut out);
        for r in regions.iter().take(MAX_REGIONS) {
            let same = if r.count > 1 { format!("（同じ形の数式 {} 個）", r.count) } else { String::new() };
            let f = if md { format!("`{}`", r.formula) } else { r.formula.clone() };
            out.push_str(&format!("{indent}数式: {} {f}{same}\n", r.range.a1()));
        }
        if regions.len() > MAX_REGIONS {
            out.push_str(&format!("{indent}数式: ほか {} 個のまとまり（--sheet で範囲を指定すると見られます）\n", regions.len() - MAX_REGIONS));
        }
    }
    if !names.is_empty() {
        out.push_str(if md { "\n## 名前\n\n" } else { "名前:\n" });
        for (name, s, f) in &names {
            let full = match scope(book, *s) {
                Some(sheet) => format!("{sheet}!{name}"),
                None => name.clone(),
            };
            out.push_str(&if md { format!("- `{full}` = `{f}`\n") } else { format!("  {full} = {f}\n") });
        }
    }
    Ok(out)
}

fn col_name(col: u32) -> String {
    CellRef::new(0, col).a1().trim_end_matches('1').to_string()
}

/// The part of `r` that is written, and whether something was left out.
fn clip(r: RangeRef, opts: &Options) -> (RangeRef, bool) {
    let mut end = r.end;
    if let Some(m) = opts.max_rows {
        end.row = end.row.min(r.start.row.saturating_add(m.saturating_sub(1)));
    }
    if let Some(m) = opts.max_cols {
        end.col = end.col.min(r.start.col.saturating_add(m.saturating_sub(1)));
    }
    (RangeRef::new(r.start, end), end != r.end)
}

fn csv_field(s: &str) -> String {
    if s.contains([',', '"', '\n', '\r']) { format!("\"{}\"", s.replace('"', "\"\"")) } else { s.to_string() }
}

/// The cells of a sheet: of `range`, or of what the sheet uses.
pub fn sheet(book: &mut Book, file: &str, si: usize, range: Option<RangeRef>, opts: &Options) -> Result<String, String> {
    let info = book.sheet_summary(si)?;
    let whole = range.or(info.used_range);
    let Some(whole) = whole else {
        return Ok(match opts.format {
            Format::Json => {
                let mut o = sheet_json(&info);
                o["file"] = json!(file);
                o["cells"] = json!([]);
                format!("{}\n", serde_json::to_string_pretty(&o).unwrap_or_default())
            }
            Format::Csv => String::new(),
            _ => format!("{}: 空のシートです\n", sheet_title(&info)),
        });
    };
    let (r, truncated) = clip(whole, opts);
    let cells = book.cells(si, r)?;
    let notice = format!(
        "{} 行 × {} 列のうち {} を出力しました。--range で範囲を指定するか、--all を付けると続きを見られます",
        whole.height(),
        whole.width(),
        r.a1()
    );
    match opts.format {
        Format::Json => {
            let mut o = sheet_json(&info);
            o["file"] = json!(file);
            o["range"] = json!(whole.a1());
            o["returnedRange"] = json!(r.a1());
            o["truncated"] = json!(truncated);
            let mut list = Vec::with_capacity(cells.len());
            for c in &cells {
                let mut e = json!({ "address": c.at.a1(), "row": c.at.row + 1, "column": c.at.col + 1, "type": type_name(&c.value), "value": json_value(book, c) });
                if let Some(f) = &c.formula {
                    e["formula"] = json!(f);
                }
                if opts.formats {
                    e["numberFormat"] = json!(c.number_format);
                    e["text"] = json!(c.text);
                }
                list.push(e);
            }
            o["cells"] = Json::Array(list);
            Ok(format!("{}\n", serde_json::to_string_pretty(&o).unwrap_or_default()))
        }
        Format::Text => {
            let mut out = format!("{}  {}\n", sheet_title(&info), r.a1());
            sheet_extras(&info, "", &mut out);
            if truncated {
                out.push_str(&format!("（{notice}）\n"));
            }
            for c in &cells {
                let v = literal(book, c);
                let mut line = match &c.formula {
                    Some(f) => format!("{}\t{f}\t→ {v}", c.at.a1()),
                    None => format!("{}\t{v}", c.at.a1()),
                };
                if opts.formats && c.number_format != "General" {
                    line.push_str(&format!("\t[{}] 表示 {}", c.number_format, serde_json::to_string(&c.text).unwrap_or_default()));
                }
                out.push_str(&line);
                out.push('\n');
            }
            if cells.is_empty() {
                out.push_str("（値のあるセルはありません）\n");
            }
            Ok(out)
        }
        Format::Csv | Format::Markdown => {
            // A grid of what the cells show (with `formats`, the formulas before it).
            let (h, w) = (r.height() as usize, r.width() as usize);
            let mut grid = vec![vec![String::new(); w]; h];
            for c in &cells {
                grid[(c.at.row - r.start.row) as usize][(c.at.col - r.start.col) as usize] = match (&c.formula, opts.formats) {
                    (Some(f), true) => format!("{f} → {}", c.text),
                    _ => c.text.clone(),
                };
            }
            let mut out = String::new();
            if opts.format == Format::Csv {
                for row in &grid {
                    out.push_str(&row.iter().map(|s| csv_field(s)).collect::<Vec<_>>().join(","));
                    out.push('\n');
                }
                return Ok(out);
            }
            out.push_str(&format!("## {}  {}\n\n", sheet_title(&info), r.a1()));
            let mut extras = String::new();
            sheet_extras(&info, "- ", &mut extras);
            if truncated {
                extras.push_str(&format!("- {notice}\n"));
            }
            if !extras.is_empty() {
                out.push_str(&extras);
                out.push('\n');
            }
            out.push_str("| |");
            for col in r.start.col..=r.end.col {
                out.push_str(&format!(" {} |", col_name(col)));
            }
            out.push_str("\n|---|");
            out.push_str(&"---|".repeat(w));
            out.push('\n');
            for (i, row) in grid.iter().enumerate() {
                out.push_str(&format!("| {} |", r.start.row as usize + i + 1));
                for s in row {
                    out.push_str(&format!(" {} |", s.replace('|', "\\|").replace('\n', "<br>")));
                }
                out.push('\n');
            }
            Ok(out)
        }
    }
}
