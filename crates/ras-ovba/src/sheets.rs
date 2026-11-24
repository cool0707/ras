//! Adding, copying, removing and renaming the worksheets of a workbook file,
//! without Excel (RAS language specification §14.16).
//!
//! [`edit_sheets`] changes the package: the sheet parts, the workbook part
//! (`<sheets>`, the indexes that defined names and the active tab hold),
//! the relationships and content types, the formulas that refer to a
//! renamed or removed sheet — and the VBA project, whose document modules
//! follow the sheets. Everything else is copied as it is. Binary workbooks
//! (`.xlsb`) are not edited.

use std::collections::{BTreeMap, HashSet};
use std::fmt;
use std::path::{Path, PathBuf};

use umya_spreadsheet::vba::{self, VbaProject};

use crate::error::OvbaError;
use crate::opc::{self, Package, PackagePlan, SheetKind};
use crate::pack::{self, BookType};
use crate::same_name;
use crate::xml::{self, Edits, TagKind};
use crate::ProjectEdit;

/// Where a new sheet goes among the sheets of the workbook.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Place {
    /// After the last sheet.
    End,
    /// Before the sheet with this code name or tab name.
    Before(String),
    /// After the sheet with this code name or tab name.
    After(String),
}

/// A change to the sheets of a workbook. A sheet is named by its code name
/// or by the name on its tab (the code name first).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetEdit {
    /// A new, empty worksheet.
    Add {
        /// Name on its tab.
        name: String,
        /// Its code name; `None`: the first free `Sheet<n>`.
        code_name: Option<String>,
        /// Where it goes.
        place: Place,
    },
    /// A copy of a worksheet: its cells, formats and settings.
    Copy {
        /// The sheet to copy.
        sheet: String,
        /// Name on the copy's tab; `None`: `<name> (2)`.
        name: Option<String>,
        /// Its code name; `None`: the first free `Sheet<n>`.
        code_name: Option<String>,
        /// Where it goes; `None`: right after the sheet it is a copy of.
        place: Option<Place>,
    },
    /// Removes a worksheet. References to it become `#REF!`.
    Remove {
        /// The sheet to remove.
        sheet: String,
    },
    /// Gives a worksheet another tab name, another code name, or both.
    Rename {
        /// The sheet to rename.
        sheet: String,
        /// New name on its tab.
        name: Option<String>,
        /// New code name.
        code_name: Option<String>,
    },
}

/// Everything [`edit_sheets`] needs to know.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetEditRequest {
    /// The workbook (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`).
    pub book: PathBuf,
    /// What to do.
    pub edit: SheetEdit,
    /// Directory that receives a copy of the file that is about to be
    /// replaced; `None` disables backups.
    pub backup_dir: Option<PathBuf>,
    /// Number of backups of this workbook to keep.
    pub backup_generations: usize,
}

/// A sheet of a workbook.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetInfo {
    /// Name on its tab.
    pub name: String,
    /// Its code name.
    pub code_name: String,
    /// A worksheet (not a chart, dialog or macro sheet).
    pub worksheet: bool,
    /// Shown (not hidden).
    pub visible: bool,
}

/// What [`edit_sheets`] did.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SheetEditReport {
    /// Tab name of the sheet that was added, copied, removed or renamed (after the change).
    pub name: String,
    /// Its code name (after the change).
    pub code_name: String,
    /// The sheet it was copied from, or its names before a rename: `(tab name, code name)`.
    pub before: Option<(String, String)>,
    /// Number of formulas and defined names whose references were rewritten.
    pub rewritten: usize,
}

/// Why a change to the sheets cannot be made.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetError {
    /// No sheet has this code name or tab name.
    NotFound(String),
    /// The sheet is not a worksheet (a chart sheet, …).
    NotAWorksheet(String),
    /// The tab name is not one Excel accepts.
    BadName(String),
    /// Another sheet has this tab name.
    NameTaken(String),
    /// The code name is not a VBA identifier of at most 31 characters.
    BadCodeName(String),
    /// A sheet or a module of the VBA project has this code name.
    CodeNameTaken(String),
    /// The only visible sheet cannot be removed.
    LastVisibleSheet,
    /// The sheet holds what a copy made in the file cannot carry (tables, drawings, …).
    CannotCopy(Vec<String>),
    /// A binary workbook (`.xlsb`).
    Binary,
}

impl fmt::Display for SheetError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SheetError::NotFound(s) => write!(f, "the workbook has no sheet `{s}`"),
            SheetError::NotAWorksheet(s) => write!(f, "`{s}` is not a worksheet"),
            SheetError::BadName(s) => write!(f, "`{s}` is not a valid sheet name"),
            SheetError::NameTaken(s) => write!(f, "a sheet named `{s}` exists already"),
            SheetError::BadCodeName(s) => write!(f, "`{s}` is not a valid code name"),
            SheetError::CodeNameTaken(s) => write!(f, "the code name `{s}` is in use"),
            SheetError::LastVisibleSheet => {
                write!(f, "a workbook needs at least one visible sheet")
            }
            SheetError::CannotCopy(what) => write!(
                f,
                "the sheet cannot be copied in the file: it has {}",
                what.join(", ")
            ),
            SheetError::Binary => {
                write!(f, "the sheets of a binary workbook (.xlsb) are not edited")
            }
        }
    }
}

impl SheetError {
    /// What cannot be done, in Japanese (the editor's and the grid's messages).
    pub fn japanese(&self) -> String {
        match self {
            SheetError::NotFound(s) => format!("ブックに「{s}」というシートがありません"),
            SheetError::NotAWorksheet(s) => format!("「{s}」はワークシートではありません（グラフシートなどは扱えません）"),
            SheetError::BadName(s) => format!("「{s}」はシートの見出しの名前にできません（31 文字まで。`: \\ / ? * [ ]` は使えず、前後を `'` にできません）"),
            SheetError::NameTaken(s) => format!("「{s}」という名前のシートが既にあります"),
            SheetError::BadCodeName(s) => format!("`{s}` はコード名にできません（文字で始まり、文字・数字・`_` だけの 31 文字までの名前にします）"),
            SheetError::CodeNameTaken(s) => format!("`{s}` という名前は、ほかのシートかモジュールが使っています"),
            SheetError::LastVisibleSheet => "表示されているシートがほかに無いため、削除できません".into(),
            SheetError::CannotCopy(what) => format!("このシートは、ブックのファイルの書き換えでは複製できません（{} があります）。Excel でブックを開いた状態で実行してください", what.join("・")),
            SheetError::Binary => "バイナリ形式のブック（.xlsb）のシートは、ファイルの書き換えでは操作できません。Excel でブックを開いた状態で実行してください".into(),
        }
    }
}

fn err(e: SheetError) -> OvbaError {
    OvbaError::Sheet(e)
}

/// Whether Excel accepts `name` as a tab name.
pub fn valid_sheet_name(name: &str) -> bool {
    let len = name.encode_utf16().count();
    (1..=31).contains(&len)
        && !name.contains([':', '\\', '/', '?', '*', '[', ']'])
        && !name.starts_with('\'')
        && !name.ends_with('\'')
}

/// Whether `name` can be the code name of a sheet: a VBA identifier of at most 31 characters.
pub fn valid_code_name(name: &str) -> bool {
    let mut chars = name.chars();
    chars.next().is_some_and(|c| c.is_alphabetic())
        && chars.all(|c| c.is_alphanumeric() || c == '_')
        && name.encode_utf16().count() <= 31
}

/// A `<sheet>` of the workbook part.
#[derive(Clone, Debug)]
struct Listed {
    sheet_id: u32,
    rid: String,
    visible: bool,
    /// Byte range of its tag in the workbook part.
    start: usize,
    end: usize,
}

/// The `<sheet>` tags of the workbook part, in order.
fn listed_sheets(workbook: &str) -> Vec<Listed> {
    let mut out = Vec::new();
    let mut inside = false;
    for tag in xml::tags(workbook) {
        match (tag.local_name(), tag.kind) {
            ("sheets", TagKind::Open) => inside = true,
            ("sheets", TagKind::Close) => inside = false,
            ("sheet", kind) if inside && kind != TagKind::Close => {
                let rid = tag
                    .attributes()
                    .into_iter()
                    .find(|a| a.local_name() == "id" && a.name.contains(':'))
                    .map(|a| a.value())
                    .unwrap_or_default();
                let state = tag.attr("state").unwrap_or_default();
                out.push(Listed {
                    sheet_id: tag
                        .attr("sheetId")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0),
                    rid,
                    visible: state.is_empty() || state == "visible",
                    start: tag.start,
                    end: tag.end,
                });
            }
            _ => {}
        }
    }
    out
}

/// The sheets of a workbook file, in order.
fn sheets_of(layout: &opc::BookLayout, listed: &[Listed]) -> Vec<SheetInfo> {
    layout
        .sheets
        .iter()
        .enumerate()
        .map(|(i, s)| SheetInfo {
            name: s.name.clone(),
            code_name: s.code_name.clone(),
            worksheet: s.kind == SheetKind::Worksheet,
            visible: listed.get(i).is_none_or(|l| l.visible),
        })
        .collect()
}

/// Index of the sheet with the code name, or else the tab name, `wanted`.
fn find_sheet(sheets: &[SheetInfo], wanted: &str) -> Result<usize, OvbaError> {
    sheets
        .iter()
        .position(|s| same_name(&s.code_name, wanted))
        .or_else(|| sheets.iter().position(|s| same_name(&s.name, wanted)))
        .ok_or_else(|| err(SheetError::NotFound(wanted.to_string())))
}

// ---------------------------------------------------------------- references to a sheet

/// How Excel writes a sheet name in front of `!`.
fn quoted(name: &str) -> String {
    let plain = name
        .chars()
        .next()
        .is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_');
    // (`A1`, `R1C1`: a name that reads as a cell needs the quotes.)
    let cell_like = {
        let letters = name.chars().take_while(|c| c.is_ascii_alphabetic()).count();
        let rest = &name[letters..];
        (1..=3).contains(&letters) && !rest.is_empty() && rest.chars().all(|c| c.is_ascii_digit())
            || name.eq_ignore_ascii_case("r")
            || name.eq_ignore_ascii_case("c")
            || name.to_ascii_uppercase().starts_with('R') && rest_is_rc(name)
    };
    if plain && !cell_like {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn rest_is_rc(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
    let rest = &rest[digits..];
    rest.is_empty()
        || rest
            .strip_prefix('C')
            .is_some_and(|r| r.chars().all(|c| c.is_ascii_digit()))
}

fn name_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.' || c == '?' || c == '\\'
}

/// Rewrites the references to the sheet `old` in a formula (or a defined name's value):
/// to the sheet `new`, or — `None` — to `#REF!`. `None` when the formula has none.
fn rewrite_refs(formula: &str, old: &str, new: Option<&str>) -> Option<String> {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut changed = false;
    let same = |a: &str| same_name(a, old);
    // One side of `Sheet1:Sheet3`, or the name itself.
    let replace = |written: &str, out: &mut String, quote: bool| -> bool {
        let (book, sheets) = match written.rfind(']') {
            Some(i) => (&written[..=i], &written[i + 1..]),
            None => ("", written),
        };
        // (Another workbook's sheet of the same name is not this sheet.)
        if !book.is_empty() {
            return false;
        }
        let parts: Vec<&str> = sheets.split(':').collect();
        if !parts.iter().any(|p| same(p)) {
            return false;
        }
        match new {
            None => out.push_str("#REF!"),
            Some(new) => {
                let renamed: Vec<String> = parts
                    .iter()
                    .map(|p| {
                        if same(p) {
                            new.to_string()
                        } else {
                            p.to_string()
                        }
                    })
                    .collect();
                let joined = renamed.join(":");
                if quote || renamed.iter().any(|p| quoted(p) != *p) {
                    out.push_str(&format!("'{}'!", joined.replace('\'', "''")));
                } else {
                    out.push_str(&format!("{joined}!"));
                }
            }
        }
        true
    };
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // A string: as it is (a doubled quote stays inside it).
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                i += 1;
                if chars[i - 1] == '"' {
                    if chars.get(i) == Some(&'"') {
                        out.push('"');
                        i += 1;
                    } else {
                        break;
                    }
                }
            }
            continue;
        }
        if c == '\'' {
            // 'Sheet name'!
            let mut j = i + 1;
            let mut name = String::new();
            let mut closed = false;
            while j < chars.len() {
                if chars[j] == '\'' {
                    if chars.get(j + 1) == Some(&'\'') {
                        name.push('\'');
                        j += 2;
                        continue;
                    }
                    closed = true;
                    j += 1;
                    break;
                }
                name.push(chars[j]);
                j += 1;
            }
            if closed && chars.get(j) == Some(&'!') && replace(&name, &mut out, true) {
                changed = true;
                i = j + 1;
                continue;
            }
            out.extend(&chars[i..j]);
            i = j;
            continue;
        }
        if name_char(c) && (i == 0 || !name_char(chars[i - 1])) {
            // Sheet1!  /  Sheet1:Sheet3!
            let mut j = i;
            while j < chars.len() && (name_char(chars[j]) || chars[j] == ':') {
                j += 1;
            }
            let word: String = chars[i..j].iter().collect();
            // (`[1]Data!A1` is a sheet of another workbook.)
            let external = i > 0 && chars[i - 1] == ']';
            if !external && chars.get(j) == Some(&'!') && replace(&word, &mut out, false) {
                changed = true;
                i = j + 1;
                continue;
            }
            out.push_str(&word);
            i = j;
            continue;
        }
        out.push(c);
        i += 1;
    }
    changed.then_some(out)
}

fn escape_text(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

/// Elements whose text is a formula.
const FORMULA_ELEMENTS: [&str; 5] = ["f", "formula", "formula1", "formula2", "definedName"];

/// Rewrites the references to the sheet `old` in an XML part: in the text of formula
/// elements and in the attributes that name a sheet. Returns the new part and how many
/// places changed; `None` when nothing did.
fn rewrite_part(doc: &str, old: &str, new: Option<&str>) -> Option<(String, usize)> {
    if !doc.to_lowercase().contains(&old.to_lowercase()) {
        return None;
    }
    let mut edits = Edits::default();
    let mut count = 0usize;
    let mut open: Option<usize> = None;
    for tag in xml::tags(doc) {
        if let Some(text_start) = open.take() {
            // The text between the opening tag and the next tag.
            if tag.kind == TagKind::Close
                && FORMULA_ELEMENTS.contains(&tag.local_name())
                && tag.start > text_start
            {
                let text = xml::unescape(&doc[text_start..tag.start]);
                if let Some(rewritten) = rewrite_refs(&text, old, new) {
                    edits.replace(text_start, tag.start, escape_text(&rewritten));
                    count += 1;
                }
            }
        }
        if tag.kind == TagKind::Close {
            continue;
        }
        let local = tag.local_name();
        if tag.kind == TagKind::Open && FORMULA_ELEMENTS.contains(&local) {
            open = Some(tag.end);
        }
        // `<hyperlink location="Sheet2!A1">`, `<worksheetSource sheet="Sheet2">`.
        for attribute in tag.attributes() {
            let rewritten = match (local, attribute.name) {
                ("hyperlink", "location") => rewrite_refs(&attribute.value(), old, new),
                ("worksheetSource", "sheet") if same_name(&attribute.value(), old) => {
                    new.map(str::to_string)
                }
                _ => None,
            };
            if let Some(value) = rewritten {
                edits.replace(
                    attribute.value_start,
                    attribute.value_end,
                    xml::escape_attr(&value),
                );
                count += 1;
            }
        }
    }
    (count > 0).then(|| (edits.apply(doc), count))
}

/// The cells and defined names of a workbook whose formulas refer to the sheet `sheet`
/// (by code name or tab name), other than that sheet's own: `Sheet1!C5`, `名前の定義 合計`.
/// What removing the sheet would turn into `#REF!`.
pub fn sheet_references(book: &Path, sheet: &str) -> Result<Vec<String>, OvbaError> {
    let bytes = std::fs::read(book).map_err(pack::map_open_error)?;
    let mut package = Package::open(bytes)?;
    let layout = opc::analyze(&mut package)?;
    if layout.binary {
        return Err(err(SheetError::Binary));
    }
    let workbook = read_text(&mut package, &layout.workbook_part)?;
    let sheets = sheets_of(&layout, &listed_sheets(&workbook));
    let index = find_sheet(&sheets, sheet)?;
    let name = sheets[index].name.clone();
    let mut out = Vec::new();
    // Defined names.
    let mut defined: Option<(String, usize)> = None;
    for tag in xml::tags(&workbook) {
        if let Some((defined_name, start)) = defined.take() {
            if tag.kind == TagKind::Close
                && tag.local_name() == "definedName"
                && rewrite_refs(&xml::unescape(&workbook[start..tag.start]), &name, None).is_some()
            {
                out.push(format!("名前の定義 {defined_name}"));
            }
        }
        if tag.kind == TagKind::Open && tag.local_name() == "definedName" {
            // (The names of the sheet itself go with it.)
            let local = tag
                .attr("localSheetId")
                .and_then(|v| v.parse::<usize>().ok());
            if local != Some(index) {
                defined = Some((tag.attr("name").unwrap_or_default(), tag.end));
            }
        }
    }
    // The formulas of the other sheets.
    for (i, other) in layout.sheets.iter().enumerate() {
        if i == index || other.part.is_empty() {
            continue;
        }
        let Some((doc, _)) = package_text(&mut package, &other.part)? else {
            continue;
        };
        if !doc.to_lowercase().contains(&name.to_lowercase()) {
            continue;
        }
        let mut cell = String::new();
        let mut open: Option<usize> = None;
        for tag in xml::tags(&doc) {
            if let Some(start) = open.take() {
                if tag.kind == TagKind::Close
                    && rewrite_refs(&xml::unescape(&doc[start..tag.start]), &name, None).is_some()
                {
                    let at = if cell.is_empty() {
                        other.name.clone()
                    } else {
                        format!("{}!{cell}", other.name)
                    };
                    if !out.contains(&at) {
                        out.push(at);
                    }
                }
            }
            if tag.kind == TagKind::Close {
                continue;
            }
            match tag.local_name() {
                "c" => cell = tag.attr("r").unwrap_or_default(),
                "sheetData" => cell.clear(),
                local if tag.kind == TagKind::Open && FORMULA_ELEMENTS.contains(&local) => {
                    open = Some(tag.end)
                }
                _ => {}
            }
        }
    }
    Ok(out)
}

fn package_text(package: &mut Package, part: &str) -> Result<Option<(String, bool)>, OvbaError> {
    match package.read(part)? {
        Some(bytes) => Ok(Some(xml::decode_part(&bytes, part)?)),
        None => Ok(None),
    }
}

fn read_text(package: &mut Package, part: &str) -> Result<String, OvbaError> {
    match package_text(package, part)? {
        Some((text, true)) => Ok(text),
        Some((_, false)) => Err(OvbaError::UnsupportedFormat(format!(
            "package part `{part}` is UTF-16 encoded and cannot be edited"
        ))),
        None => Err(OvbaError::Format(format!(
            "the package has no part `{part}`"
        ))),
    }
}

// ---------------------------------------------------------------- the package

/// Lists the sheets of a workbook file, in order.
pub fn workbook_sheets(book: &Path) -> Result<Vec<SheetInfo>, OvbaError> {
    let bytes = std::fs::read(book).map_err(pack::map_open_error)?;
    let mut package = Package::open(bytes)?;
    let layout = opc::analyze(&mut package)?;
    let listed = if layout.binary {
        Vec::new()
    } else {
        listed_sheets(&read_text(&mut package, &layout.workbook_part)?)
    };
    Ok(sheets_of(&layout, &listed))
}

/// What the edit of the package is made of.
#[derive(Default)]
struct Changes {
    replace: BTreeMap<String, Vec<u8>>,
    remove: HashSet<String>,
    add: Vec<(String, Vec<u8>)>,
}

/// The index shift of what counts sheets by position: `localSheetId` of defined names,
/// `activeTab` / `firstSheet` of the workbook views. `map` gives the new index of an old
/// one, or `None` for a sheet that is gone (its defined names go with it).
fn reindex(
    workbook: &str,
    edits: &mut Edits,
    map: &dyn Fn(usize) -> Option<usize>,
    count_after: usize,
) {
    let mut drop_from: Option<usize> = None;
    for tag in xml::tags(workbook) {
        if let Some(start) = drop_from {
            if tag.kind == TagKind::Close && tag.local_name() == "definedName" {
                edits.replace(start, tag.end, "");
                drop_from = None;
            }
            continue;
        }
        if tag.kind == TagKind::Close {
            continue;
        }
        match tag.local_name() {
            "definedName" => {
                let Some(attribute) = tag.attribute("localSheetId") else {
                    continue;
                };
                let Ok(old) = attribute.value().parse::<usize>() else {
                    continue;
                };
                match map(old) {
                    Some(new) if new != old => {
                        edits.replace(attribute.value_start, attribute.value_end, new.to_string())
                    }
                    Some(_) => {}
                    None if tag.kind == TagKind::Empty => edits.replace(tag.start, tag.end, ""),
                    None => drop_from = Some(tag.start),
                }
            }
            "workbookView" => {
                for key in ["activeTab", "firstSheet"] {
                    let Some(attribute) = tag.attribute(key) else {
                        continue;
                    };
                    let Ok(old) = attribute.value().parse::<usize>() else {
                        continue;
                    };
                    let new = map(old)
                        .unwrap_or(old.min(count_after.saturating_sub(1)))
                        .min(count_after.saturating_sub(1));
                    if new != old {
                        edits.replace(attribute.value_start, attribute.value_end, new.to_string());
                    }
                }
            }
            _ => {}
        }
    }
}

/// The first free `Sheet<n>`.
fn free_code_name(taken: &[String]) -> String {
    (1..)
        .map(|n| format!("Sheet{n}"))
        .find(|c| !taken.iter().any(|t| same_name(t, c)))
        .expect("a free name")
}

/// Sets the `codeName` of a worksheet part, removes what marks it as the selected tab and
/// (for a copy) the references to parts it does not get.
fn sheet_xml_for(doc: &str, code_name: &str, copy: bool) -> String {
    let mut edits = Edits::default();
    let mut has_sheet_pr = false;
    let mut worksheet_end = None;
    for tag in xml::tags(doc) {
        if tag.kind == TagKind::Close {
            continue;
        }
        match tag.local_name() {
            "worksheet" => worksheet_end = Some(tag.end),
            "sheetPr" if !has_sheet_pr => {
                has_sheet_pr = true;
                match tag.attribute("codeName") {
                    Some(a) => {
                        edits.replace(a.value_start, a.value_end, xml::escape_attr(code_name))
                    }
                    None => edits.insert(
                        tag.name_end,
                        format!(" codeName=\"{}\"", xml::escape_attr(code_name)),
                    ),
                }
            }
            "sheetView" if copy => {
                if let Some(a) = tag.attribute("tabSelected") {
                    edits.replace(a.value_start, a.value_end, "0");
                }
            }
            // The printer settings stay with the sheet they were made for.
            "pageSetup" if copy => {
                if let Some(a) = tag
                    .attributes()
                    .into_iter()
                    .find(|a| a.local_name() == "id" && a.name.contains(':'))
                {
                    let start = doc[..a.value_start].rfind(a.name).unwrap_or(a.value_start);
                    edits.replace(start, a.value_end + 1, "");
                }
            }
            _ => {}
        }
    }
    if !has_sheet_pr {
        if let Some(at) = worksheet_end {
            let prefix = xml::tags(doc)
                .find(|t| t.local_name() == "worksheet")
                .map(|t| t.prefix().to_string())
                .unwrap_or_default();
            edits.insert(
                at,
                format!(
                    "<{prefix}sheetPr codeName=\"{}\"/>",
                    xml::escape_attr(code_name)
                ),
            );
        }
    }
    edits.apply(doc)
}

const EMPTY_SHEET: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetPr/><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetData/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>";
const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/// The relationships of a part: `(id, type, resolved target, external)`.
fn relationships(
    package: &mut Package,
    part: &str,
) -> Result<Vec<(String, String, String, bool)>, OvbaError> {
    let Some((doc, _)) = package_text(package, &opc::rels_part_for(part))? else {
        return Ok(Vec::new());
    };
    Ok(xml::tags(&doc)
        .filter(|t| t.kind != TagKind::Close && t.local_name() == "Relationship")
        .map(|t| {
            let external = t
                .attr("TargetMode")
                .is_some_and(|m| m.eq_ignore_ascii_case("External"));
            let target = t.attr("Target").unwrap_or_default();
            (
                t.attr("Id").unwrap_or_default(),
                t.attr("Type").unwrap_or_default(),
                if external {
                    target
                } else {
                    opc::resolve_target(part, &target)
                },
                external,
            )
        })
        .collect())
}

/// Inserts a worksheet part: the part, its `<sheet>`, its relationship and its content type.
#[allow(clippy::too_many_arguments)]
fn insert_sheet(
    package: &mut Package,
    layout: &opc::BookLayout,
    workbook: &str,
    listed: &[Listed],
    changes: &mut Changes,
    sheet_xml: String,
    name: &str,
    at: usize,
) -> Result<(), OvbaError> {
    // The part.
    let dir = layout
        .sheets
        .iter()
        .find(|s| s.kind == SheetKind::Worksheet && !s.part.is_empty())
        .map(|s| opc::part_dir(&s.part).to_string())
        .unwrap_or_else(|| {
            format!("{}/worksheets", opc::part_dir(&layout.workbook_part))
                .trim_start_matches('/')
                .to_string()
        });
    let part = (1..)
        .map(|n| format!("{dir}/sheet{n}.xml"))
        .find(|p| {
            package.find(p).is_none() && !changes.add.iter().any(|(a, _)| a.eq_ignore_ascii_case(p))
        })
        .expect("a free part name");
    changes.add.push((part.clone(), sheet_xml.into_bytes()));
    // The relationship of the workbook to it.
    let rels_doc = read_text(package, &layout.workbook_rels_part)?;
    let existing = relationships(package, &layout.workbook_part)?;
    let rid = (1..)
        .map(|n| format!("rId{n}"))
        .find(|id| !existing.iter().any(|r| r.0 == *id))
        .expect("a free id");
    let rel_type = existing
        .iter()
        .find(|r| r.1.ends_with("/worksheet"))
        .map(|r| r.1.clone())
        .unwrap_or_else(|| REL_WORKSHEET.to_string());
    let workbook_dir = opc::part_dir(&layout.workbook_part);
    let target = match part.strip_prefix(&format!("{workbook_dir}/")) {
        Some(relative) if !workbook_dir.is_empty() => relative.to_string(),
        _ => format!("/{part}"),
    };
    let close = rels_doc
        .rfind("</")
        .ok_or_else(|| OvbaError::Format("the workbook relationships are not closed".into()))?;
    let prefix = xml::tags(&rels_doc)
        .find(|t| t.local_name() == "Relationship")
        .map(|t| t.prefix().to_string())
        .unwrap_or_default();
    let mut rels_edits = Edits::default();
    rels_edits.insert(
        close,
        format!(
            "<{prefix}Relationship Id=\"{rid}\" Type=\"{rel_type}\" Target=\"{}\"/>",
            xml::escape_attr(&target)
        ),
    );
    changes.replace.insert(
        layout.workbook_rels_part.clone(),
        rels_edits.apply(&rels_doc).into_bytes(),
    );
    // The content type.
    let types_part = package
        .find("[Content_Types].xml")
        .map(str::to_string)
        .ok_or_else(|| OvbaError::Format("the package has no [Content_Types].xml".into()))?;
    let types = String::from_utf8(
        changes
            .replace
            .get(&types_part)
            .cloned()
            .unwrap_or(package.read(&types_part)?.unwrap_or_default()),
    )
    .map_err(|_| OvbaError::Format("[Content_Types].xml is not UTF-8".into()))?;
    let content_type = xml::tags(&types)
        .filter(|t| t.local_name() == "Override")
        .find_map(|t| {
            t.attr("ContentType")
                .filter(|c| c.ends_with("worksheet+xml"))
        })
        .unwrap_or_else(|| CT_WORKSHEET.to_string());
    let close = types
        .rfind("</")
        .ok_or_else(|| OvbaError::Format("[Content_Types].xml is not closed".into()))?;
    let prefix = xml::tags(&types)
        .find(|t| t.local_name() == "Override" || t.local_name() == "Default")
        .map(|t| t.prefix().to_string())
        .unwrap_or_default();
    let mut type_edits = Edits::default();
    type_edits.insert(
        close,
        format!("<{prefix}Override PartName=\"/{part}\" ContentType=\"{content_type}\"/>"),
    );
    changes
        .replace
        .insert(types_part, type_edits.apply(&types).into_bytes());
    // The workbook part.
    let sheet_id = listed.iter().map(|l| l.sheet_id).max().unwrap_or(0) + 1;
    let (tag_prefix, rid_attr) = xml::tags(workbook)
        .find(|t| t.local_name() == "sheet" && t.kind != TagKind::Close)
        .map(|t| {
            (
                t.prefix().to_string(),
                t.attributes()
                    .into_iter()
                    .find(|a| a.local_name() == "id" && a.name.contains(':'))
                    .map(|a| a.name.to_string())
                    .unwrap_or_else(|| "r:id".into()),
            )
        })
        .unwrap_or_else(|| (String::new(), "r:id".into()));
    let tag = format!(
        "<{tag_prefix}sheet name=\"{}\" sheetId=\"{sheet_id}\" {rid_attr}=\"{rid}\"/>",
        xml::escape_attr(name)
    );
    let mut edits = Edits::default();
    let position = match listed.get(at) {
        Some(next) => next.start,
        None => listed
            .last()
            .map(|l| l.end)
            .ok_or_else(|| OvbaError::Format("the workbook has no sheets".into()))?,
    };
    edits.insert(position, tag);
    reindex(
        workbook,
        &mut edits,
        &|old| Some(if old >= at { old + 1 } else { old }),
        listed.len() + 1,
    );
    changes.replace.insert(
        layout.workbook_part.clone(),
        edits.apply(workbook).into_bytes(),
    );
    Ok(())
}

/// Rewrites the references to the sheet `old` in every part that may hold formulas.
fn rewrite_references(
    package: &mut Package,
    layout: &opc::BookLayout,
    changes: &mut Changes,
    old: &str,
    new: Option<&str>,
    skip_part: Option<&str>,
) -> Result<usize, OvbaError> {
    let mut count = 0;
    let parts: Vec<String> = package
        .names()
        .into_iter()
        .filter(|n| {
            let lower = n.to_ascii_lowercase();
            lower.ends_with(".xml")
                && (lower.contains("/worksheets/")
                    || lower.contains("/chartsheets/")
                    || lower.contains("/charts/")
                    || lower.contains("/pivotcache/")
                    || same_name(n, &layout.workbook_part))
        })
        .collect();
    for part in parts {
        if skip_part.is_some_and(|s| same_name(s, &part)) || changes.remove.contains(&part) {
            continue;
        }
        let doc = match changes.replace.get(&part) {
            Some(bytes) => String::from_utf8_lossy(bytes).into_owned(),
            None => match package_text(package, &part)? {
                Some((text, true)) => text,
                _ => continue,
            },
        };
        if let Some((rewritten, n)) = rewrite_part(&doc, old, new) {
            changes.replace.insert(part, rewritten.into_bytes());
            count += n;
        }
    }
    Ok(count)
}

/// Brings the VBA project of the package `bytes` in line with its sheets: the document
/// module of a sheet whose code name changed is renamed, those of sheets that are gone are
/// removed, and the new sheets get theirs.
fn follow_in_project(
    bytes: Vec<u8>,
    rename: Option<(String, String)>,
) -> Result<Vec<u8>, OvbaError> {
    let mut package = Package::open(bytes.clone())?;
    let layout = opc::analyze(&mut package)?;
    let Some(part) = layout.vba_part.clone() else {
        return Ok(bytes);
    };
    let bin = package.read(&part)?.unwrap_or_default();
    let loaded = VbaProject::read(&bin)?;
    if loaded.is_protected() {
        return Err(OvbaError::Protected);
    }
    let documents = pack::document_plan(&layout);
    let edit = ProjectEdit {
        rename: rename.into_iter().collect(),
        keep_documents: Some(
            documents
                .known
                .iter()
                .map(|(name, _, _)| name.clone())
                .collect(),
        ),
        ..ProjectEdit::default()
    };
    let built = vba::build_project(Some(&loaded), &documents.ensure, &edit)?;
    let before: Vec<String> = loaded
        .modules()
        .iter()
        .map(|m| m.name().to_string())
        .collect();
    let after: Vec<String> = built.modules.iter().map(|m| m.name.clone()).collect();
    if before == after {
        return Ok(bytes);
    }
    // The project changed: its signature is no longer valid.
    let mut edits = opc::plan_package_edits(
        &mut package,
        &layout,
        &PackagePlan {
            main_content_type: None,
            remove_signatures: true,
            write_code_names: false,
        },
    )?;
    edits.replace.insert(part, built.bytes);
    package.rewrite(&edits.replace, &edits.remove, &edits.add)
}

/// Adds, copies, removes or renames a worksheet of a workbook file (RAS specification
/// §14.16). The new package is written to a temporary file next to the workbook, the
/// previous file is copied to `backup_dir`, and the temporary file is renamed over it.
///
/// Errors: [`OvbaError::Sheet`] (what the request asks cannot be done: nothing is
/// written), [`OvbaError::Locked`] (the workbook is open in Excel), [`OvbaError::Protected`].
pub fn edit_sheets(req: &SheetEditRequest) -> Result<SheetEditReport, OvbaError> {
    if pack::book_type(&req.book)? == BookType::Xlsb {
        return Err(err(SheetError::Binary));
    }
    pack::check_not_locked(&req.book, true)?;
    let bytes = std::fs::read(&req.book).map_err(pack::map_open_error)?;
    let (new_package, report) = edit_sheet_bytes(bytes, &req.edit)?;
    let Some(new_package) = new_package else {
        return Ok(report);
    };
    let mut temp = pack::write_temp_file(&req.book, &new_package)?;
    if let Some(dir) = &req.backup_dir {
        pack::back_up(&req.book, dir, req.backup_generations)?;
    }
    if let Ok(metadata) = std::fs::metadata(&req.book) {
        let _ = std::fs::set_permissions(&temp.path, metadata.permissions());
    }
    pack::check_not_locked(&req.book, true)?;
    std::fs::rename(&temp.path, &req.book).map_err(|e| {
        if pack::is_sharing_violation(&e) || e.kind() == std::io::ErrorKind::PermissionDenied {
            OvbaError::Locked
        } else {
            OvbaError::Io(e)
        }
    })?;
    temp.keep = true;
    Ok(report)
}

/// [`edit_sheets`] on the bytes of a workbook (`.xlsx`, `.xlsm`, ...): the bytes of the new
/// package (None when nothing changed) and what was done.
pub fn edit_sheet_bytes(
    bytes: Vec<u8>,
    edit: &SheetEdit,
) -> Result<(Option<Vec<u8>>, SheetEditReport), OvbaError> {
    let mut package = Package::open(bytes)?;
    let mut layout = opc::analyze(&mut package)?;
    if layout.binary {
        return Err(err(SheetError::Binary));
    }
    // Sheets without a recorded code name have the one their place gives them (`Sheet<n>`):
    // it is written down first, or another sheet coming or going would change it.
    if layout.sheets.iter().any(|s| s.explicit_code_name.is_none())
        || layout.explicit_workbook_code_name.is_none()
    {
        let mut pins = opc::plan_package_edits(
            &mut package,
            &layout,
            &PackagePlan {
                main_content_type: None,
                remove_signatures: false,
                write_code_names: true,
            },
        )?;
        // Only the code names: no VBA project is added to a workbook that has none.
        pins.replace.retain(|part, _| {
            same_name(part, &layout.workbook_part)
                || layout.sheets.iter().any(|s| same_name(&s.part, part))
        });
        if !pins.replace.is_empty() {
            let pinned = package.rewrite(&pins.replace, &HashSet::new(), &[])?;
            package = Package::open(pinned)?;
            layout = opc::analyze(&mut package)?;
        }
    }
    let workbook = read_text(&mut package, &layout.workbook_part)?;
    let listed = listed_sheets(&workbook);
    if listed.len() != layout.sheets.len() {
        return Err(OvbaError::Format(
            "the sheets of the workbook part could not be read".into(),
        ));
    }
    let sheets = sheets_of(&layout, &listed);
    // Code names in use: those of the sheets and the workbook, and the modules of the project.
    let mut code_names: Vec<String> = sheets.iter().map(|s| s.code_name.clone()).collect();
    code_names.push(layout.workbook_code_name.clone());
    if let Some(part) = &layout.vba_part {
        let bin = package.read(part)?.unwrap_or_default();
        code_names.extend(
            VbaProject::read(&bin)?
                .modules()
                .iter()
                .map(|m| m.name().to_string()),
        );
    }
    let check_name = |name: &str, but: Option<usize>| -> Result<(), OvbaError> {
        if !valid_sheet_name(name) {
            return Err(err(SheetError::BadName(name.to_string())));
        }
        if sheets
            .iter()
            .enumerate()
            .any(|(i, s)| Some(i) != but && same_name(&s.name, name))
        {
            return Err(err(SheetError::NameTaken(name.to_string())));
        }
        Ok(())
    };
    let check_code_name = |code_name: &str, but: Option<&str>| -> Result<(), OvbaError> {
        if !valid_code_name(code_name) {
            return Err(err(SheetError::BadCodeName(code_name.to_string())));
        }
        if code_names
            .iter()
            .any(|c| same_name(c, code_name) && !but.is_some_and(|b| same_name(b, c)))
        {
            return Err(err(SheetError::CodeNameTaken(code_name.to_string())));
        }
        Ok(())
    };
    let place_of = |place: &Place| -> Result<usize, OvbaError> {
        Ok(match place {
            Place::End => sheets.len(),
            Place::Before(s) => find_sheet(&sheets, s)?,
            Place::After(s) => find_sheet(&sheets, s)? + 1,
        })
    };
    let worksheet = |wanted: &str| -> Result<usize, OvbaError> {
        let index = find_sheet(&sheets, wanted)?;
        if !sheets[index].worksheet || layout.sheets[index].part.is_empty() {
            return Err(err(SheetError::NotAWorksheet(sheets[index].name.clone())));
        }
        Ok(index)
    };

    let mut changes = Changes::default();
    let mut report = SheetEditReport::default();
    let mut module_rename: Option<(String, String)> = None;
    match edit {
        SheetEdit::Add {
            name,
            code_name,
            place,
        } => {
            check_name(name, None)?;
            let code_name = match code_name {
                Some(c) => {
                    check_code_name(c, None)?;
                    c.clone()
                }
                None => free_code_name(&code_names),
            };
            let at = place_of(place)?;
            insert_sheet(
                &mut package,
                &layout,
                &workbook,
                &listed,
                &mut changes,
                sheet_xml_for(EMPTY_SHEET, &code_name, false),
                name,
                at,
            )?;
            report.name = name.clone();
            report.code_name = code_name;
        }
        SheetEdit::Copy {
            sheet,
            name,
            code_name,
            place,
        } => {
            let index = worksheet(sheet)?;
            let source = &sheets[index];
            let name = match name {
                Some(n) => n.clone(),
                // As Excel names a copy: `<name> (2)`, cut to the 31 characters a name may have.
                None => (2..)
                    .map(|n| {
                        let suffix = format!(" ({n})");
                        let keep: String = source
                            .name
                            .chars()
                            .take(31usize.saturating_sub(suffix.chars().count()))
                            .collect();
                        format!("{keep}{suffix}")
                    })
                    .find(|candidate| !sheets.iter().any(|s| same_name(&s.name, candidate)))
                    .expect("a free name"),
            };
            check_name(&name, None)?;
            let code_name = match code_name {
                Some(c) => {
                    check_code_name(c, None)?;
                    c.clone()
                }
                None => free_code_name(&code_names),
            };
            let at = match place {
                Some(p) => place_of(p)?,
                None => index + 1,
            };
            // What hangs on the sheet's relationships does not come along.
            let part = layout.sheets[index].part.clone();
            let blocking: Vec<String> = relationships(&mut package, &part)?
                .into_iter()
                .filter(|r| !r.1.ends_with("/printerSettings"))
                .map(|r| r.1.rsplit('/').next().unwrap_or_default().to_string())
                .collect();
            if !blocking.is_empty() {
                let mut what = blocking;
                what.sort();
                what.dedup();
                return Err(err(SheetError::CannotCopy(what)));
            }
            let doc = read_text(&mut package, &part)?;
            insert_sheet(
                &mut package,
                &layout,
                &workbook,
                &listed,
                &mut changes,
                sheet_xml_for(&doc, &code_name, true),
                &name,
                at,
            )?;
            report.before = Some((source.name.clone(), source.code_name.clone()));
            report.name = name;
            report.code_name = code_name;
        }
        SheetEdit::Remove { sheet } => {
            let index = worksheet(sheet)?;
            if !sheets
                .iter()
                .enumerate()
                .any(|(i, s)| i != index && s.visible)
            {
                return Err(err(SheetError::LastVisibleSheet));
            }
            let part = layout.sheets[index].part.clone();
            // The part, its relationships, and the parts only it uses.
            let mut gone: Vec<String> = vec![part.clone()];
            let mut pending = vec![part.clone()];
            let all_rels: Vec<String> = package
                .names()
                .into_iter()
                .filter(|n| n.to_ascii_lowercase().ends_with(".rels"))
                .collect();
            while let Some(current) = pending.pop() {
                let rels_part = opc::rels_part_for(&current);
                for (_, _, target, external) in relationships(&mut package, &current)? {
                    if external || gone.iter().any(|g| same_name(g, &target)) {
                        continue;
                    }
                    // Used by a part that stays?
                    let mut shared = false;
                    for other in &all_rels {
                        let owner_gone = gone
                            .iter()
                            .any(|g| same_name(&opc::rels_part_for(g), other));
                        if owner_gone || same_name(other, &rels_part) {
                            continue;
                        }
                        let Some((doc, _)) = package_text(&mut package, other)? else {
                            continue;
                        };
                        let file = target.rsplit('/').next().unwrap_or(&target);
                        if doc.contains(file) {
                            shared = true;
                            break;
                        }
                    }
                    if !shared && package.find(&target).is_some() {
                        gone.push(target.clone());
                        pending.push(target);
                    }
                }
            }
            for g in &gone {
                for name in [g.clone(), opc::rels_part_for(g)] {
                    if let Some(found) = package.find(&name).map(str::to_string) {
                        changes.remove.insert(found);
                    }
                }
            }
            // The calculation chain lists cells by sheet: Excel builds it anew.
            let chain = relationships(&mut package, &layout.workbook_part)?
                .into_iter()
                .find(|r| r.1.ends_with("/calcChain"))
                .map(|r| r.2);
            if let Some(found) = chain
                .as_ref()
                .and_then(|c| package.find(c).map(str::to_string))
            {
                changes.remove.insert(found.clone());
                gone.push(found);
            }
            // The workbook part: the sheet, the indexes, the references.
            let mut edits = Edits::default();
            edits.replace(listed[index].start, listed[index].end, "");
            reindex(
                &workbook,
                &mut edits,
                &|old| {
                    if old == index {
                        None
                    } else {
                        Some(if old > index { old - 1 } else { old })
                    }
                },
                listed.len() - 1,
            );
            changes.replace.insert(
                layout.workbook_part.clone(),
                edits.apply(&workbook).into_bytes(),
            );
            // Its relationships and content types.
            let rels_doc = read_text(&mut package, &layout.workbook_rels_part)?;
            let mut rels_edits = Edits::default();
            for tag in xml::tags(&rels_doc)
                .filter(|t| t.kind != TagKind::Close && t.local_name() == "Relationship")
            {
                let target = opc::resolve_target(
                    &layout.workbook_part,
                    &tag.attr("Target").unwrap_or_default(),
                );
                if tag.attr("Id").as_deref() == Some(listed[index].rid.as_str())
                    || chain.as_ref().is_some_and(|c| same_name(c, &target))
                {
                    rels_edits.replace(tag.start, tag.end, "");
                }
            }
            changes.replace.insert(
                layout.workbook_rels_part.clone(),
                rels_edits.apply(&rels_doc).into_bytes(),
            );
            if let Some(types_part) = package.find("[Content_Types].xml").map(str::to_string) {
                let types =
                    String::from_utf8_lossy(&package.read(&types_part)?.unwrap_or_default())
                        .into_owned();
                let mut type_edits = Edits::default();
                for tag in xml::tags(&types)
                    .filter(|t| t.kind != TagKind::Close && t.local_name() == "Override")
                {
                    let name = tag.attr("PartName").unwrap_or_default();
                    if gone
                        .iter()
                        .any(|g| same_name(g, name.trim_start_matches('/')))
                    {
                        type_edits.replace(tag.start, tag.end, "");
                    }
                }
                changes
                    .replace
                    .insert(types_part, type_edits.apply(&types).into_bytes());
            }
            report.rewritten = rewrite_references(
                &mut package,
                &layout,
                &mut changes,
                &sheets[index].name,
                None,
                Some(&part),
            )?;
            report.name = sheets[index].name.clone();
            report.code_name = sheets[index].code_name.clone();
        }
        SheetEdit::Rename {
            sheet,
            name,
            code_name,
        } => {
            let index = worksheet(sheet)?;
            let old = sheets[index].clone();
            report.before = Some((old.name.clone(), old.code_name.clone()));
            report.name = old.name.clone();
            report.code_name = old.code_name.clone();
            if let Some(code_name) = code_name.as_ref().filter(|c| **c != old.code_name) {
                check_code_name(code_name, Some(&old.code_name))?;
                let part = layout.sheets[index].part.clone();
                let doc = read_text(&mut package, &part)?;
                changes
                    .replace
                    .insert(part, sheet_xml_for(&doc, code_name, false).into_bytes());
                module_rename = Some((old.code_name.clone(), code_name.clone()));
                report.code_name = code_name.clone();
            }
            if let Some(name) = name.as_ref().filter(|n| **n != old.name) {
                check_name(name, Some(index))?;
                let mut edits = Edits::default();
                if let Some(tag) = xml::tags(&workbook).find(|t| t.start == listed[index].start) {
                    if let Some(a) = tag.attribute("name") {
                        edits.replace(a.value_start, a.value_end, xml::escape_attr(name));
                    }
                }
                changes.replace.insert(
                    layout.workbook_part.clone(),
                    edits.apply(&workbook).into_bytes(),
                );
                report.rewritten = rewrite_references(
                    &mut package,
                    &layout,
                    &mut changes,
                    &old.name,
                    Some(name),
                    None,
                )?;
                report.name = name.clone();
            }
        }
    }
    if changes.replace.is_empty() && changes.remove.is_empty() && changes.add.is_empty() {
        return Ok((None, report));
    }
    let new_package = package.rewrite(&changes.replace, &changes.remove, &changes.add)?;
    let new_package = follow_in_project(new_package, module_rename)?;
    // What was written reads back with the sheets as they should be.
    let mut check = Package::open(new_package.clone())
        .map_err(|e| OvbaError::Verify(format!("the new package cannot be opened: {e}")))?;
    let after = opc::analyze(&mut check)
        .map_err(|e| OvbaError::Verify(format!("the new package cannot be analysed: {e}")))?;
    let expected = match edit {
        SheetEdit::Add { .. } | SheetEdit::Copy { .. } => sheets.len() + 1,
        SheetEdit::Remove { .. } => sheets.len() - 1,
        SheetEdit::Rename { .. } => sheets.len(),
    };
    let present = after
        .sheets
        .iter()
        .any(|s| same_name(&s.code_name, &report.code_name) && same_name(&s.name, &report.name));
    if after.sheets.len() != expected || present == matches!(edit, SheetEdit::Remove { .. }) {
        return Err(OvbaError::Verify(
            "the sheets of the new package are not as requested".into(),
        ));
    }
    Ok((Some(new_package), report))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn references_in_formulas() {
        let to = |f: &str, new: Option<&str>| rewrite_refs(f, "Data", new);
        assert_eq!(
            to("SUM(Data!A1:A3)+data!B1", Some("売上")).as_deref(),
            Some("SUM('売上'!A1:A3)+'売上'!B1")
        );
        assert_eq!(
            to("'Data'!A1", Some("Sales")).as_deref(),
            Some("'Sales'!A1")
        );
        assert_eq!(to("Data!A1", Some("Sales")).as_deref(), Some("Sales!A1"));
        assert_eq!(
            to("Data!A1+OtherData!A1+MyData!A1", None).as_deref(),
            Some("#REF!A1+OtherData!A1+MyData!A1")
        );
        assert_eq!(
            to("Sheet1:Data!A1", Some("New Data")).as_deref(),
            Some("'Sheet1:New Data'!A1")
        );
        assert_eq!(to("Sheet1:Data!A1", None).as_deref(), Some("#REF!A1"));
        // Strings, other workbooks and other sheets stay.
        assert_eq!(to("\"Data!A1\"&\"a\"\"b\"", Some("X")), None);
        assert_eq!(to("'[Other.xlsx]Data'!A1+[1]Data!B2", Some("X")), None);
        assert_eq!(to("Data2!A1+A1", Some("X")), None);
        // A name that needs quotes is taken with them.
        assert_eq!(
            rewrite_refs("'My ''Sheet'''!A1+1", "My 'Sheet'", Some("A1")).as_deref(),
            Some("'A1'!A1+1")
        );
        assert_eq!(quoted("Sales_2026"), "Sales_2026");
        assert_eq!(quoted("2026"), "'2026'");
        assert_eq!(quoted("AB12"), "'AB12'");
    }

    #[test]
    fn names_excel_accepts() {
        assert!(valid_sheet_name("売上 (2)") && valid_sheet_name("a"));
        for bad in ["", "a/b", "a:b", "[x]", "'x", "x'", &"x".repeat(32)] {
            assert!(!valid_sheet_name(bad), "{bad}");
        }
        assert!(
            valid_code_name("Sheet1")
                && valid_code_name("売上")
                && !valid_code_name("1st")
                && !valid_code_name("a b")
                && !valid_code_name("")
        );
        assert_eq!(
            free_code_name(&["Sheet1".into(), "sheet2".into(), "ThisWorkbook".into()]),
            "Sheet3"
        );
    }

    #[test]
    fn a_sheet_part_gets_its_code_name() {
        let doc = "<worksheet xmlns=\"x\" xmlns:r=\"r\"><sheetPr codeName=\"Old\"><tabColor rgb=\"FF\"/></sheetPr><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetData/><pageSetup orientation=\"portrait\" r:id=\"rId1\"/></worksheet>";
        let copy = sheet_xml_for(doc, "New", true);
        assert!(
            copy.contains("<sheetPr codeName=\"New\">")
                && copy.contains("tabSelected=\"0\"")
                && copy.contains("<pageSetup orientation=\"portrait\" />"),
            "{copy}"
        );
        assert!(sheet_xml_for(
            "<x:worksheet xmlns:x=\"x\"><x:sheetData/></x:worksheet>",
            "S",
            false
        )
        .starts_with("<x:worksheet xmlns:x=\"x\"><x:sheetPr codeName=\"S\"/>"));
        assert!(sheet_xml_for(EMPTY_SHEET, "S", false).contains("<sheetPr codeName=\"S\"/>"));
    }
}
