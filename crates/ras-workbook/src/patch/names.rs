//! The names of the workbook as "名前の管理" changes them: defined, changed (renamed, made to
//! refer to something else, given a comment) and deleted, in the `definedNames` of
//! `xl/workbook.xml`. The cells are not written: their formulas are worked out again when the
//! file is read (and Excel is told to work them out when it opens it).

use std::collections::{BTreeMap, HashSet};

use serde_json::{Value as Json, json};

use super::scan::{self, Edits};
use super::{Pack, Result, WORKBOOK};

/// A name: `scope` the sheet it belongs to (None: the workbook), `refers` what it refers to
/// (`=Sheet1!$A$1:$B$5`, a formula, a constant).
#[derive(Clone, Debug, PartialEq)]
pub struct NameDef {
    pub name: String,
    pub scope: Option<String>,
    pub refers: String,
    pub comment: String,
}

#[derive(Clone, Debug, PartialEq)]
pub enum NameChange {
    /// A new name (`old` None), or the name `old` (name, scope) changed into `new` (its scope
    /// stays, as Excel's 名前の編集 has it).
    Define {
        old: Option<(String, Option<String>)>,
        new: NameDef,
    },
    Delete {
        name: String,
        scope: Option<String>,
    },
}

/// A name Excel takes: a letter, `_` or `\` first, then letters, digits, `_`, `.` and `\`; not
/// a cell's address (A1, R1C1), nor R or C.
pub(super) fn name_ok(name: &str) -> bool {
    super::insert::table_name_ok(name)
}

/// The names of the file, as the grid lists them (Excel's own, `_xlnm.` and the like, left
/// out): { name, sheet (of a sheet's own), ref (without `=`), comment }.
pub(crate) fn names_of(bytes: &[u8]) -> Result<Vec<Json>> {
    let mut pack = Pack::open(bytes)?;
    let doc = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let root = scan::tree(&doc, &[]).ok_or("xl/workbook.xml を読み取れません")?;
    let sheets = sheet_names(&doc, &root);
    let mut out = Vec::new();
    let Some(names) = root.child("definedNames") else {
        return Ok(out);
    };
    for n in names.children.iter().filter(|c| c.local() == "definedName") {
        let name = n.attr(&doc, "name").unwrap_or_default();
        if name.starts_with("_xlnm.")
            || name.starts_with("_Filter")
            || n.attr(&doc, "hidden").as_deref() == Some("1")
        {
            continue;
        }
        let refers = gridcraft_xlsx_formulas::fmla::from_file(&n.text(&doc));
        let mut j = json!({ "name": name, "ref": refers.trim_start_matches('=') });
        if let Some(s) = scope_of(&doc, n).and_then(|i| sheets.get(i)) {
            j["sheet"] = json!(s);
        }
        if let Some(c) = n.attr(&doc, "comment").filter(|c| !c.is_empty()) {
            j["comment"] = json!(c);
        }
        out.push(j);
    }
    Ok(out)
}

fn sheet_names(doc: &str, root: &scan::El) -> Vec<String> {
    root.child("sheets")
        .map(|s| {
            s.children
                .iter()
                .filter(|c| c.local() == "sheet")
                .map(|c| c.attr(doc, "name").unwrap_or_default())
                .collect()
        })
        .unwrap_or_default()
}

fn scope_of(doc: &str, el: &scan::El) -> Option<usize> {
    el.attr(doc, "localSheetId").and_then(|i| i.parse().ok())
}

/// The file with the names changed as `change` says.
pub(super) fn change(bytes: &[u8], change: &NameChange) -> Result<Vec<u8>> {
    let mut pack = Pack::open(bytes)?;
    let doc = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let root = scan::tree(&doc, &[]).ok_or("xl/workbook.xml を読み取れません")?;
    let p = scan::prefix(root.name);
    let sheets = sheet_names(&doc, &root);
    let index = |scope: &Option<String>| -> Result<Option<usize>> {
        match scope {
            None => Ok(None),
            Some(s) => sheets
                .iter()
                .position(|n| n.to_lowercase() == s.to_lowercase())
                .map(Some)
                .ok_or_else(|| format!("シート `{s}` がありません")),
        }
    };
    let list: Vec<&scan::El> = root
        .child("definedNames")
        .map(|n| {
            n.children
                .iter()
                .filter(|c| c.local() == "definedName")
                .collect()
        })
        .unwrap_or_default();
    let find = |name: &str, scope: Option<usize>| {
        list.iter().copied().find(|n| {
            scope_of(&doc, n) == scope
                && n.attr(&doc, "name")
                    .is_some_and(|x| x.to_lowercase() == name.to_lowercase())
        })
    };
    let mut edits = Edits::default();
    match change {
        NameChange::Delete { name, scope } => {
            let el =
                find(name, index(scope)?).ok_or_else(|| format!("名前 `{name}` がありません"))?;
            match root.child("definedNames") {
                // The last one: the list with it.
                Some(names) if list.len() == 1 => edits.replace(names.span(), ""),
                _ => edits.replace(el.span(), ""),
            }
        }
        NameChange::Define { old, new } => {
            let name = new.name.trim();
            if !name_ok(name) {
                return Err(format!(
                    "名前 `{name}` は使えません（先頭は文字か _ か \\、ほかは文字・数字・_・.・\\ で、セルの番地と同じものは使えません）"
                ));
            }
            let body = new.refers.trim();
            let body = body.strip_prefix('=').unwrap_or(body).trim();
            if body.is_empty() || gridcraft_formula::parse(body).is_err() {
                return Err(format!("参照範囲 `={body}` が正しくありません"));
            }
            let text = gridcraft_xlsx_formulas::fmla::text_to_file(body);
            let was = match old {
                Some((n, s)) => {
                    Some(find(n, index(s)?).ok_or_else(|| format!("名前 `{n}` がありません"))?)
                }
                None => None,
            };
            let scope = match was {
                Some(el) => scope_of(&doc, el),
                None => index(&new.scope)?,
            };
            if let Some(other) = find(name, scope)
                && was.is_none_or(|w| w.open != other.open)
            {
                return Err(format!(
                    "名前 `{name}` は既に存在します。一意の名前を入力してください"
                ));
            }
            if tables_of(&mut pack)?.contains(&name.to_lowercase()) {
                return Err(format!(
                    "名前 `{name}` は既にテーブルの名前に使われています"
                ));
            }
            match was {
                // Written anew, with the attributes it has besides (hidden, function, ...).
                Some(el) => {
                    let mut tag = format!("<{} name=\"{}\"", el.name, scan::escape(name));
                    if !new.comment.is_empty() {
                        tag.push_str(&format!(" comment=\"{}\"", scan::escape(&new.comment)));
                    }
                    for (key, span) in scan::attrs(&doc, el.open.clone()) {
                        if key != "name" && key != "comment" {
                            tag.push_str(&format!(" {key}=\"{}\"", &doc[span]));
                        }
                    }
                    edits.replace(
                        el.span(),
                        format!("{tag}>{}</{}>", scan::escape(&text), el.name),
                    );
                }
                None => {
                    let mut tag = format!("<{p}definedName name=\"{}\"", scan::escape(name));
                    if !new.comment.is_empty() {
                        tag.push_str(&format!(" comment=\"{}\"", scan::escape(&new.comment)));
                    }
                    if let Some(i) = scope {
                        tag.push_str(&format!(" localSheetId=\"{i}\""));
                    }
                    let added = format!("{tag}>{}</{p}definedName>", scan::escape(&text));
                    match root.child("definedNames") {
                        Some(names) if names.end != names.open.end => {
                            edits.insert(names.inner.end, added)
                        }
                        Some(names) => edits.replace(
                            names.span(),
                            format!("<{p}definedNames>{added}</{p}definedNames>"),
                        ),
                        None => {
                            const BEFORE: &[&str] =
                                &["sheets", "functionGroups", "externalReferences"];
                            let last = root
                                .children
                                .iter()
                                .rev()
                                .find(|c| BEFORE.contains(&c.local()))
                                .ok_or("xl/workbook.xml に sheets がありません")?;
                            edits.insert(
                                last.end,
                                format!("<{p}definedNames>{added}</{p}definedNames>"),
                            );
                        }
                    }
                }
            }
        }
    }
    // The cells that use the name: Excel works them out again when it opens the file.
    match root.child("calcPr") {
        Some(calc) => scan::set_attr(&doc, calc.open.clone(), "fullCalcOnLoad", "1", &mut edits),
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
                edits.insert(last.end, format!("<{p}calcPr fullCalcOnLoad=\"1\"/>"));
            }
        }
    }
    let part = pack.find(WORKBOOK).unwrap_or_else(|| WORKBOOK.into());
    let mut replace = BTreeMap::new();
    replace.insert(part, edits.apply(&doc).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// The names of the tables of the workbook (lowercase).
fn tables_of(pack: &mut Pack) -> Result<HashSet<String>> {
    let parts: Vec<String> = pack
        .names()
        .iter()
        .filter(|n| {
            let n = n.to_ascii_lowercase();
            n.starts_with("xl/tables/") && n.ends_with(".xml")
        })
        .cloned()
        .collect();
    let mut out = HashSet::new();
    for part in parts {
        let Some(doc) = pack.text(&part)? else {
            continue;
        };
        if let Some(t) = scan::tags(&doc, 0).find(|t| !t.is_close() && t.local() == "table") {
            for key in ["name", "displayName"] {
                if let Some(n) = scan::attr(&doc, t.start..t.end, key) {
                    out.insert(n.to_lowercase());
                }
            }
        }
    }
    Ok(out)
}
