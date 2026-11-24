//! Input rules (データ > データの入力規則) put on cells, or cleared from them, written as Excel
//! writes them: a `<dataValidation>` of the sheet's `<dataValidations>`. Putting a rule on cells
//! takes them out of the rules they had first (as Excel's dialog does); a rule left with no
//! cells goes, and so does a `<dataValidations>` left with none. Excel 2010 writes a rule whose
//! list is on another sheet in the sheet's `extLst` (`x14:dataValidation`): those are read and
//! cleared too.
//!
//! The rules of the sheets as the grid reads them ([`rules_of`]): the file's attributes as they
//! are (umya-spreadsheet reads a missing `operator` as `lessThan`, where it means `between`).

use std::collections::HashMap;

use gridcraft_core::RangeRef;
use serde_json::Value as Json;

use super::insert::{Sheet, place, tree};
use super::scan::{self, Edits, El};
use super::{Pack, Result, WORKBOOK, sheet_parts};

/// An input rule.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct Validation {
    /// `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`; empty: "すべての値"
    /// (a rule of messages alone).
    pub kind: String,
    /// `between` (the default), `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`,
    /// `greaterThanOrEqual`, `lessThanOrEqual`.
    pub operator: String,
    /// As the file has them: without `=` (`10`, `$B$1`, `"a,b,c"`, `46023`).
    pub formula1: String,
    pub formula2: String,
    /// "空白を無視する".
    pub allow_blank: bool,
    /// (A list) "ドロップダウン リストから選択する" off (the file's `showDropDown`).
    pub no_drop_down: bool,
    /// "セルを選択したときに入力時メッセージを表示する", its title and message.
    pub show_input: bool,
    pub prompt_title: String,
    pub prompt: String,
    /// "無効なデータが入力されたらエラー メッセージを表示する", its style (`stop`, `warning`,
    /// `information`), title and message.
    pub show_error: bool,
    pub error_style: String,
    pub error_title: String,
    pub error: String,
    /// 日本語入力 (`noControl`, `on`, `off`, `disabled`, `hiragana`, `fullKatakana`,
    /// `halfKatakana`, `fullAlpha`, `halfAlpha`; empty: `noControl`).
    pub ime: String,
    /// The cells it is on (read from the file; not written from here).
    pub ranges: Vec<RangeRef>,
}

const KINDS: &[&str] = &[
    "",
    "whole",
    "decimal",
    "list",
    "date",
    "time",
    "textLength",
    "custom",
];
const OPERATORS: &[&str] = &[
    "between",
    "notBetween",
    "equal",
    "notEqual",
    "greaterThan",
    "lessThan",
    "greaterThanOrEqual",
    "lessThanOrEqual",
];
const STYLES: &[&str] = &["stop", "warning", "information"];
const IMES: &[&str] = &[
    "noControl",
    "on",
    "off",
    "disabled",
    "hiragana",
    "fullKatakana",
    "halfKatakana",
    "fullAlpha",
    "halfAlpha",
];

impl Validation {
    /// The rule of an edit's JSON (`{ type, operator, formula1, formula2, allowBlank,
    /// inCellDropdown, showInput, promptTitle, prompt, showError, errorStyle, errorTitle,
    /// error, ime }`; the formulas as the file has them, or with `=`).
    pub fn of_json(o: &Json) -> Result<Validation> {
        let text = |k: &str| {
            o.get(k)
                .and_then(Json::as_str)
                .unwrap_or_default()
                .to_string()
        };
        let flag = |k: &str, default: bool| o.get(k).and_then(Json::as_bool).unwrap_or(default);
        let formula = |k: &str| text(k).trim().trim_start_matches('=').trim().to_string();
        let mut kind = text("type");
        if kind == "any" || kind == "none" {
            kind.clear();
        }
        let v = Validation {
            kind,
            operator: text("operator"),
            formula1: formula("formula1"),
            formula2: formula("formula2"),
            allow_blank: flag("allowBlank", true),
            no_drop_down: !flag("inCellDropdown", true),
            show_input: flag("showInput", true),
            prompt_title: text("promptTitle"),
            prompt: text("prompt"),
            show_error: flag("showError", true),
            error_style: text("errorStyle"),
            error_title: text("errorTitle"),
            error: text("error"),
            ime: text("ime"),
            ranges: Vec::new(),
        };
        v.check()?;
        Ok(v)
    }

    /// What is wrong with the rule.
    fn check(&self) -> Result<()> {
        let one_of = |what: &str, v: &str, list: &[&str]| {
            if v.is_empty() || list.contains(&v) {
                Ok(())
            } else {
                Err(format!(
                    "{what} `{v}` は指定できません（{}）",
                    list.iter()
                        .filter(|x| !x.is_empty())
                        .copied()
                        .collect::<Vec<_>>()
                        .join(" / ")
                ))
            }
        };
        one_of("入力規則の種類", &self.kind, KINDS)?;
        one_of("入力規則の `operator`", &self.operator, OPERATORS)?;
        one_of("エラー メッセージのスタイル", &self.error_style, STYLES)?;
        one_of("日本語入力", &self.ime, IMES)?;
        if !self.kind.is_empty() && self.formula1.is_empty() {
            return Err(match self.kind.as_str() {
                "list" => "リストの元の値を入力してください（`formula1`）".into(),
                "custom" => "数式を入力してください（`formula1`）".into(),
                _ => "値を入力してください（`formula1`）".into(),
            });
        }
        if self.two() && self.formula2.is_empty() {
            return Err("最大値を入力してください（`formula2`）".into());
        }
        // (Excel's dialog takes 255 characters of a list's values.)
        if self.kind == "list"
            && self.formula1.starts_with('"')
            && self.formula1.chars().count() > 257
        {
            return Err("リストの元の値は 255 文字までです".into());
        }
        for (what, v, max) in [
            ("入力時メッセージのタイトル", &self.prompt_title, 32),
            ("入力時メッセージ", &self.prompt, 255),
            ("エラー メッセージのタイトル", &self.error_title, 32),
            ("エラー メッセージ", &self.error, 225),
        ] {
            if v.chars().count() > max {
                return Err(format!("{what}は {max} 文字までです"));
            }
        }
        Ok(())
    }

    /// The rule compares with two values (between, not between).
    fn two(&self) -> bool {
        matches!(
            self.kind.as_str(),
            "whole" | "decimal" | "date" | "time" | "textLength"
        ) && matches!(self.operator.as_str(), "" | "between" | "notBetween")
    }

    /// The rule compares the cell's value with values (an operator applies).
    fn compares(&self) -> bool {
        matches!(
            self.kind.as_str(),
            "whole" | "decimal" | "date" | "time" | "textLength"
        )
    }

    /// The rule as `<dataValidation>` (`p`: the prefix of the sheet's elements).
    fn xml(&self, p: &str, refs: &str) -> String {
        let mut a = String::new();
        let mut attr = |k: &str, v: &str| a.push_str(&format!(" {k}=\"{}\"", scan::escape(v)));
        if !self.kind.is_empty() {
            attr("type", &self.kind);
        }
        if !self.error_style.is_empty() && self.error_style != "stop" {
            attr("errorStyle", &self.error_style);
        }
        if !self.ime.is_empty() && self.ime != "noControl" {
            attr("imeMode", &self.ime);
        }
        if self.compares() && !self.operator.is_empty() && self.operator != "between" {
            attr("operator", &self.operator);
        }
        if self.allow_blank {
            attr("allowBlank", "1");
        }
        if self.kind == "list" && self.no_drop_down {
            attr("showDropDown", "1");
        }
        if self.show_input {
            attr("showInputMessage", "1");
        }
        if self.show_error {
            attr("showErrorMessage", "1");
        }
        for (k, v) in [
            ("errorTitle", &self.error_title),
            ("error", &self.error),
            ("promptTitle", &self.prompt_title),
            ("prompt", &self.prompt),
        ] {
            if !v.is_empty() {
                attr(k, v);
            }
        }
        attr("sqref", refs);
        let mut inner = String::new();
        if !self.kind.is_empty() {
            inner.push_str(&format!(
                "<{p}formula1>{}</{p}formula1>",
                scan::escape(&self.formula1)
            ));
            if self.two() {
                inner.push_str(&format!(
                    "<{p}formula2>{}</{p}formula2>",
                    scan::escape(&self.formula2)
                ));
            }
        }
        if inner.is_empty() {
            format!("<{p}dataValidation{a}/>")
        } else {
            format!("<{p}dataValidation{a}>{inner}</{p}dataValidation>")
        }
    }

    /// The rule of a `<dataValidation>` (of the sheet, or of Excel 2010's `x14:dataValidation`,
    /// its formulas in `xm:f` and its cells in `xm:sqref`).
    fn read(doc: &str, el: &El) -> Validation {
        let text = |k: &str| el.attr(doc, k).unwrap_or_default();
        let flag = |k: &str| el.attr(doc, k).is_some_and(|v| v == "1" || v == "true");
        let formula = |k: &str| {
            el.child(k)
                .map(|f| match f.child("f") {
                    Some(inner) => inner.text(doc),
                    None => f.text(doc),
                })
                .unwrap_or_default()
                .trim()
                .to_string()
        };
        let sqref = el
            .attr(doc, "sqref")
            .or_else(|| el.child("sqref").map(|s| s.text(doc)))
            .unwrap_or_default();
        let mut kind = text("type");
        if kind == "none" {
            kind.clear();
        }
        Validation {
            kind,
            operator: text("operator"),
            formula1: formula("formula1"),
            formula2: formula("formula2"),
            allow_blank: flag("allowBlank"),
            no_drop_down: flag("showDropDown"),
            show_input: flag("showInputMessage"),
            prompt_title: text("promptTitle"),
            prompt: text("prompt"),
            show_error: flag("showErrorMessage"),
            error_style: text("errorStyle"),
            error_title: text("errorTitle"),
            error: text("error"),
            ime: text("imeMode"),
            ranges: ranges(&sqref),
        }
    }

    /// The operator as the rule means it (`between` when the file has none).
    pub fn op(&self) -> &str {
        if self.operator.is_empty() {
            "between"
        } else {
            &self.operator
        }
    }

    /// The rule for the grid: `{ ranges, type, operator, formula1, formula2, allowBlank,
    /// inCellDropdown, showInput, promptTitle, prompt, showError, errorStyle, errorTitle,
    /// error, ime }` (the formulas as the file has them).
    pub fn to_json(&self) -> Json {
        serde_json::json!({
            "ranges": self.ranges.iter().map(|r| serde_json::json!([r.start.row, r.start.col, r.end.row, r.end.col])).collect::<Vec<_>>(),
            "type": if self.kind.is_empty() { "any" } else { self.kind.as_str() },
            "operator": self.op(),
            "formula1": self.formula1,
            "formula2": self.formula2,
            "allowBlank": self.allow_blank,
            "inCellDropdown": !self.no_drop_down,
            "showInput": self.show_input,
            "promptTitle": self.prompt_title,
            "prompt": self.prompt,
            "showError": self.show_error,
            "errorStyle": if self.error_style.is_empty() { "stop" } else { self.error_style.as_str() },
            "errorTitle": self.error_title,
            "error": self.error,
            "ime": if self.ime.is_empty() { "noControl" } else { self.ime.as_str() },
        })
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

/// The ranges of `sqref` without the cells of `clear`; None when they keep all their cells.
fn cleared(refs: &str, clear: &[RangeRef]) -> Option<Vec<RangeRef>> {
    let list = ranges(refs);
    if !list.iter().any(|a| clear.iter().any(|b| a.intersects(b))) {
        return None;
    }
    let mut left = list;
    for b in clear {
        left = left
            .into_iter()
            .flat_map(|a| super::conditional::subtract(a, b))
            .collect();
    }
    Some(left)
}

/// The input rules of the sheets, by name: those of the sheet first, then those of Excel 2010.
pub(crate) fn rules_of(bytes: &[u8]) -> Result<HashMap<String, Vec<Validation>>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let mut out = HashMap::new();
    for (name, part) in parts {
        let Some(doc) = pack.find(&part).and_then(|p| pack.text(&p).ok().flatten()) else {
            continue;
        };
        if !doc.contains("dataValidation") {
            continue;
        }
        let list = sheet_rules(&doc);
        if !list.is_empty() {
            out.insert(name, list);
        }
    }
    Ok(out)
}

/// The input rules of a sheet's XML.
pub(crate) fn sheet_rules(doc: &str) -> Vec<Validation> {
    let Some(root) = scan::tree(doc, &["sheetData"]) else {
        return Vec::new();
    };
    let mut list: Vec<Validation> = root
        .children
        .iter()
        .filter(|c| c.local() == "dataValidations")
        .flat_map(|dvs| {
            dvs.children
                .iter()
                .filter(|c| c.local() == "dataValidation")
        })
        .map(|el| Validation::read(doc, el))
        .collect();
    if let Some(ext_list) = root.child("extLst") {
        for ext in ext_list.children.iter().filter(|c| c.local() == "ext") {
            let Some(dvs) = ext.child("dataValidations") else {
                continue;
            };
            list.extend(
                dvs.children
                    .iter()
                    .filter(|c| c.local() == "dataValidation")
                    .map(|el| Validation::read(doc, el)),
            );
        }
    }
    list.retain(|v| !v.ranges.is_empty());
    list
}

/// The file with `rule` put on the cells `on` of the sheet `sheet` (the rules they had taken
/// off them), or, `rule` None, without the rules of those cells.
pub(super) fn validation(
    bytes: &[u8],
    sheet: &str,
    on: &[RangeRef],
    rule: Option<&Validation>,
) -> Result<Vec<u8>> {
    let s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let p = scan::prefix(root.name).to_string();
    let mut edits = Edits::default();
    // The rules of the sheet: those left, and the new one after them.
    let main = root
        .children
        .iter()
        .find(|c| c.local() == "dataValidations");
    let mut left = 0;
    if let Some(dvs) = main {
        for dv in dvs
            .children
            .iter()
            .filter(|c| c.local() == "dataValidation")
        {
            let Some(refs) = dv.attr(&sdoc, "sqref") else {
                left += 1;
                continue;
            };
            match cleared(&refs, on) {
                None => left += 1,
                Some(rest) if rest.is_empty() => edits.replace(dv.span(), ""),
                Some(rest) => {
                    left += 1;
                    scan::set_attr(&sdoc, dv.open.clone(), "sqref", &sqref(&rest), &mut edits);
                }
            }
        }
    }
    clear_x14(&sdoc, &root, on, &mut edits)?;
    let new = rule.map(|r| r.xml(&p, &sqref(on)));
    if let Some(r) = rule {
        r.check()?;
    }
    let count = left + usize::from(new.is_some());
    match main {
        Some(dvs) if count == 0 => edits.replace(dvs.span(), ""),
        Some(dvs) => {
            if let Some(el) = new {
                if dvs.open.end == dvs.end {
                    // (`<dataValidations count="0"/>`.)
                    let open = &sdoc[dvs.open.clone()];
                    let open = format!("{}>", open.trim_end_matches("/>").trim_end());
                    edits.replace(dvs.span(), format!("{open}{el}</{p}dataValidations>"));
                } else {
                    edits.insert(dvs.inner.end, el);
                }
            }
            if dvs.attr(&sdoc, "count").is_some() || count > 0 {
                scan::set_attr(
                    &sdoc,
                    dvs.open.clone(),
                    "count",
                    &count.to_string(),
                    &mut edits,
                );
            }
        }
        None => {
            if let Some(el) = new {
                edits.insert(
                    place(&root, "dataValidations"),
                    format!("<{p}dataValidations count=\"1\">{el}</{p}dataValidations>"),
                );
            }
        }
    }
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    let doc = edits.apply(&sdoc);
    tree(&doc, &s.part)?;
    s.finish(Some(doc))
}

/// The edits that take the cells `on` out of Excel 2010's input rules in the sheet's `extLst`.
fn clear_x14(doc: &str, root: &El, on: &[RangeRef], edits: &mut Edits) -> Result<()> {
    let Some(ext_list) = root.child("extLst") else {
        return Ok(());
    };
    let text = &doc[ext_list.span()];
    let base = ext_list.open.start;
    let list = scan::tree(text, &[]).ok_or("シートの extLst を読み取れません")?;
    let shift = |r: std::ops::Range<usize>| r.start + base..r.end + base;
    let mut exts_left = 0;
    for ext in list.children.iter().filter(|c| c.local() == "ext") {
        let Some(dvs) = ext.child("dataValidations") else {
            exts_left += 1;
            continue;
        };
        let mut dvs_left = 0;
        for dv in dvs
            .children
            .iter()
            .filter(|c| c.local() == "dataValidation")
        {
            let Some(refs) = dv.child("sqref") else {
                dvs_left += 1;
                continue;
            };
            match cleared(&refs.text(text), on) {
                None => dvs_left += 1,
                Some(rest) if rest.is_empty() => edits.replace(shift(dv.span()), ""),
                Some(rest) => {
                    dvs_left += 1;
                    edits.replace(shift(refs.inner.clone()), sqref(&rest));
                }
            }
        }
        if dvs_left == 0 {
            edits.replace(shift(ext.span()), "");
        } else {
            exts_left += 1;
            if let Some(span) = scan::attr_span(text, dvs.open.clone(), "count") {
                edits.replace(shift(span), dvs_left.to_string());
            }
        }
    }
    if exts_left == 0 {
        edits.replace(ext_list.span(), "");
    }
    Ok(())
}
