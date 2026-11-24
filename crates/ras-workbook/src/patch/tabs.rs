//! The sheets of the workbook as their tabs change them: added, copied, removed and renamed
//! (as `ras sheet` does it in the file, [`ras_ovba::edit_sheet_bytes`]), and the color of a tab
//! ("シート見出しの色": the `tabColor` of the sheet's `sheetPr`).

use std::collections::{BTreeMap, HashSet};

use ras_ovba::{OvbaError, SheetEdit};

use super::scan::{self, Edits};
use super::{Pack, Result, WORKBOOK, sheet_parts};

/// The file with the sheets changed as `edit` says.
pub(super) fn edit(bytes: &[u8], edit: &SheetEdit) -> Result<Vec<u8>> {
    match ras_ovba::edit_sheet_bytes(bytes.to_vec(), edit) {
        Ok((Some(out), _)) => Ok(out),
        Ok((None, _)) => Ok(bytes.to_vec()),
        Err(OvbaError::Sheet(e)) => Err(e.japanese()),
        Err(e) => Err(format!("シートを変更できませんでした: {e}")),
    }
}

/// The file with the tab of the sheet `sheet` given the color `color` (`#RRGGBB`; None: none).
pub(super) fn color(bytes: &[u8], sheet: &str, color: Option<&str>) -> Result<Vec<u8>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let part = parts
        .get(sheet)
        .or_else(|| {
            parts
                .iter()
                .find(|(n, _)| n.to_lowercase() == sheet.to_lowercase())
                .map(|(_, p)| p)
        })
        .and_then(|p| pack.find(p))
        .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
    let doc = pack
        .text(&part)?
        .ok_or_else(|| format!("{part} を読めません"))?;
    let root = scan::tree(&doc, &["sheetData", "extLst"])
        .ok_or_else(|| format!("{part} を読み取れません"))?;
    let p = scan::prefix(root.name);
    let tab = color.map(|c| {
        format!(
            "<{p}tabColor rgb=\"FF{}\"/>",
            c.trim_start_matches('#').to_uppercase()
        )
    });
    let mut edits = Edits::default();
    match root.child("sheetPr") {
        // `<sheetPr .../>`: made a start tag with the color in it.
        Some(pr) if pr.open.end == pr.end => {
            if let Some(tab) = tab {
                let head = doc[pr.open.clone()]
                    .trim_end_matches('>')
                    .trim_end_matches('/')
                    .trim_end();
                edits.replace(pr.open.clone(), format!("{head}>{tab}</{}>", pr.name));
            }
        }
        // The first of its children, as the schema orders them.
        Some(pr) => match (pr.child("tabColor"), tab) {
            (Some(old), tab) => edits.replace(old.span(), tab.unwrap_or_default()),
            (None, Some(tab)) => edits.insert(pr.inner.start, tab),
            (None, None) => {}
        },
        // The first child of the worksheet.
        None => {
            if let Some(tab) = tab {
                edits.insert(root.open.end, format!("<{p}sheetPr>{tab}</{p}sheetPr>"));
            }
        }
    }
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut replace = BTreeMap::new();
    replace.insert(part, edits.apply(&doc).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// A color as the edit gives it: `#RRGGBB` (`RRGGBB`), or none.
pub(super) fn color_of(value: Option<&serde_json::Value>) -> Result<Option<String>> {
    match value {
        None | Some(serde_json::Value::Null) => Ok(None),
        Some(serde_json::Value::String(s)) if s.is_empty() => Ok(None),
        Some(serde_json::Value::String(s)) => {
            let hex = s.trim().trim_start_matches('#');
            if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
                Ok(Some(format!("#{}", hex.to_uppercase())))
            } else {
                Err(format!("色 `{s}` は #RRGGBB で指定してください"))
            }
        }
        Some(other) => Err(format!("色 `{other}` は #RRGGBB で指定してください")),
    }
}

/// The `<sheet>` tags of the workbook (`xl/workbook.xml`), in order: their spans and names.
fn sheet_tags(workbook: &str) -> Vec<(std::ops::Range<usize>, String)> {
    scan::tags(workbook, 0)
        .filter(|t| !t.is_close() && t.local() == "sheet")
        .filter_map(|t| {
            let span = t.start..t.end;
            scan::attr(workbook, span.clone(), "name").map(|n| (span, n))
        })
        .collect()
}

/// The place of the sheet named `sheet` among `tags` (its name in any case).
fn sheet_place(tags: &[(std::ops::Range<usize>, String)], sheet: &str) -> Result<usize> {
    tags.iter()
        .position(|(_, n)| n == sheet)
        .or_else(|| {
            tags.iter()
                .position(|(_, n)| n.to_lowercase() == sheet.to_lowercase())
        })
        .ok_or_else(|| format!("シート `{sheet}` がありません"))
}

/// Removes the attribute `name` of the start tag at `tag`.
fn remove_attr(doc: &str, tag: std::ops::Range<usize>, name: &str, edits: &mut Edits) {
    let Some((n, value)) = scan::attrs(doc, tag).into_iter().find(|(n, _)| *n == name) else {
        return;
    };
    // From the space before the name to the closing quote.
    let start = n.as_ptr() as usize - doc.as_ptr() as usize;
    let start = doc[..start].trim_end().len();
    edits.replace(start..value.end + 1, "");
}

/// The state of a sheet (`state` of its `<sheet>`): `0` visible, otherwise hidden.
fn hidden_at(workbook: &str, span: std::ops::Range<usize>) -> bool {
    scan::attr(workbook, span, "state").is_some_and(|s| s != "visible")
}

/// The file with the sheet `sheet` hidden (ホーム > 書式 > シートを表示しない) or shown again
/// ("シートの再表示"). The last visible sheet is not hidden; the active one, hidden, gives its
/// place to the next visible one.
pub(super) fn visible(bytes: &[u8], sheet: &str, show: bool) -> Result<Vec<u8>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let tags = sheet_tags(&workbook);
    let at = sheet_place(&tags, sheet)?;
    let span = tags[at].0.clone();
    let mut edits = Edits::default();
    if show {
        if !hidden_at(&workbook, span.clone()) {
            return Ok(bytes.to_vec());
        }
        remove_attr(&workbook, span, "state", &mut edits);
    } else {
        if hidden_at(&workbook, span.clone()) {
            return Ok(bytes.to_vec());
        }
        let others = tags
            .iter()
            .enumerate()
            .filter(|(i, (s, _))| *i != at && !hidden_at(&workbook, s.clone()))
            .map(|(i, _)| i)
            .collect::<Vec<_>>();
        if others.is_empty() {
            return Err("ブックのすべてのシートを非表示にすることはできません。".into());
        }
        scan::set_attr(&workbook, span, "state", "hidden", &mut edits);
        // The active sheet (and the first tab shown): the next visible one.
        let next = others
            .iter()
            .copied()
            .find(|i| *i > at)
            .unwrap_or(others[others.len() - 1]);
        if let Some(view) =
            scan::tags(&workbook, 0).find(|t| !t.is_close() && t.local() == "workbookView")
        {
            let vspan = view.start..view.end;
            let index = |key: &str| {
                scan::attr(&workbook, vspan.clone(), key)
                    .and_then(|v| v.parse::<usize>().ok())
                    .unwrap_or(0)
            };
            if index("activeTab") == at {
                scan::set_attr(
                    &workbook,
                    vspan.clone(),
                    "activeTab",
                    &next.to_string(),
                    &mut edits,
                );
            }
            if scan::attr(&workbook, vspan.clone(), "firstSheet").is_some()
                && index("firstSheet") == at
            {
                scan::set_attr(
                    &workbook,
                    vspan.clone(),
                    "firstSheet",
                    &next.to_string(),
                    &mut edits,
                );
            }
        }
    }
    let mut replace = BTreeMap::new();
    replace.insert(WORKBOOK.to_string(), edits.apply(&workbook).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// The file with the sheet `sheet` moved before the sheet `before` (None: to the end), as
/// "シートの移動またはコピー" moves it: the order of the `<sheet>`s, with what counts them (the
/// active sheet, the first tab shown, the names of a sheet) following.
pub(super) fn move_to(bytes: &[u8], sheet: &str, before: Option<&str>) -> Result<Vec<u8>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let tags = sheet_tags(&workbook);
    let from = sheet_place(&tags, sheet)?;
    let n = tags.len();
    // The new order: old indexes.
    let mut order: Vec<usize> = (0..n).filter(|i| *i != from).collect();
    let to = match before {
        Some(b) => {
            let b = sheet_place(&tags, b)?;
            if b == from {
                return Ok(bytes.to_vec());
            }
            order.iter().position(|i| *i == b).unwrap_or(order.len())
        }
        None => order.len(),
    };
    order.insert(to, from);
    if order.iter().copied().eq(0..n) {
        return Ok(bytes.to_vec());
    }
    // Old index → new.
    let mut new_of = vec![0; n];
    for (new, old) in order.iter().enumerate() {
        new_of[*old] = new;
    }
    let mut edits = Edits::default();
    for (new, old) in order.iter().enumerate() {
        if new != *old {
            edits.replace(
                tags[new].0.clone(),
                workbook[tags[*old].0.clone()].to_string(),
            );
        }
    }
    for t in scan::tags(&workbook, 0).filter(|t| !t.is_close()) {
        let span = t.start..t.end;
        let keys: &[&str] = match t.local() {
            "workbookView" => &["activeTab", "firstSheet"],
            "definedName" => &["localSheetId"],
            _ => continue,
        };
        for key in keys {
            if let Some(old) =
                scan::attr(&workbook, span.clone(), key).and_then(|v| v.parse::<usize>().ok())
                && old < n
                && new_of[old] != old
            {
                scan::set_attr(
                    &workbook,
                    span.clone(),
                    key,
                    &new_of[old].to_string(),
                    &mut edits,
                );
            }
        }
    }
    let mut replace = BTreeMap::new();
    replace.insert(WORKBOOK.to_string(), edits.apply(&workbook).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

// ---------------------------------------------------------------------------- シートの保護

/// What a protected sheet still lets the user do (the check boxes of "シートの保護"), as the
/// attributes of `sheetProtection` name them.
pub(crate) const ALLOWS: &[&str] = &[
    "selectLockedCells",
    "selectUnlockedCells",
    "formatCells",
    "formatColumns",
    "formatRows",
    "insertColumns",
    "insertRows",
    "insertHyperlinks",
    "deleteColumns",
    "deleteRows",
    "sort",
    "autoFilter",
    "pivotTables",
    "objects",
    "scenarios",
];

/// The protection of a sheet as an edit gives it: on (with the password, empty for none, and
/// what it allows) or off (with the password it was given).
#[derive(Clone, Debug, PartialEq)]
pub enum Protect {
    On {
        password: String,
        allow: Vec<String>,
    },
    Off {
        password: String,
    },
}

/// The file with the sheet `sheet` protected or no longer (ホーム > 書式 > シートの保護), as Excel
/// writes it: the password's SHA-512 hash, salted, 100000 times.
pub(super) fn protect(bytes: &[u8], sheet: &str, change: &Protect) -> Result<Vec<u8>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let part = parts
        .get(sheet)
        .or_else(|| {
            parts
                .iter()
                .find(|(n, _)| n.to_lowercase() == sheet.to_lowercase())
                .map(|(_, p)| p)
        })
        .and_then(|p| pack.find(p))
        .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
    let doc = pack
        .text(&part)?
        .ok_or_else(|| format!("{part} を読めません"))?;
    let root = scan::tree(&doc, &["sheetData", "extLst"])
        .ok_or_else(|| format!("{part} を読み取れません"))?;
    let p = scan::prefix(root.name);
    let old = root.child("sheetProtection");
    let mut edits = Edits::default();
    match change {
        Protect::Off { password } => {
            let Some(old) = old else {
                return Ok(bytes.to_vec());
            };
            if !password_fits(&doc, old.open.clone(), password) {
                return Err("入力したパスワードが間違っています。CapsLock キーの状態を確認し、大文字と小文字が正しく使われていることを確認してください。".into());
            }
            edits.replace(old.span(), "");
        }
        Protect::On { password, allow } => {
            if let Some(bad) = allow.iter().find(|a| !ALLOWS.contains(&a.as_str())) {
                return Err(format!("シートの保護で許可する操作 `{bad}` はありません"));
            }
            let mut xml = format!("<{p}sheetProtection");
            if !password.is_empty() {
                let salt = salt();
                let hash = sha512_spun(&salt, password, 100_000);
                xml.push_str(&format!(
                    " algorithmName=\"SHA-512\" hashValue=\"{}\" saltValue=\"{}\" spinCount=\"100000\"",
                    base64(&hash),
                    base64(&salt)
                ));
            }
            xml.push_str(" sheet=\"1\"");
            let allowed = |k: &str| allow.iter().any(|a| a == k);
            // (The attributes say what is protected; each has its default.)
            for key in ALLOWS {
                let protected = !allowed(key);
                let default = !matches!(
                    *key,
                    "selectLockedCells" | "selectUnlockedCells" | "objects" | "scenarios"
                );
                if protected != default {
                    xml.push_str(&format!(" {key}=\"{}\"", if protected { 1 } else { 0 }));
                }
            }
            xml.push_str("/>");
            match old {
                Some(old) => edits.replace(old.span(), xml),
                None => {
                    // After the cells (and their calculation), as the schema orders them.
                    let Some(last) = root
                        .children
                        .iter()
                        .rev()
                        .find(|c| matches!(c.local(), "sheetData" | "sheetCalcPr"))
                    else {
                        return Err(format!("{part} にセルがありません"));
                    };
                    edits.replace(last.end..last.end, xml);
                }
            }
        }
    }
    let mut replace = BTreeMap::new();
    replace.insert(part, edits.apply(&doc).into_bytes());
    pack.rewrite(&replace, &HashSet::new())
}

/// The sheets of the file that are protected: `{ password, allow: { selectLockedCells, ... } }`
/// (what each still allows), by name.
pub(crate) fn protections(
    bytes: &[u8],
) -> Result<std::collections::HashMap<String, serde_json::Value>> {
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let mut out = std::collections::HashMap::new();
    for (name, part) in parts {
        let Some(doc) = pack.find(&part).and_then(|p| pack.text(&p).ok().flatten()) else {
            continue;
        };
        let Some(tag) =
            scan::tags(&doc, 0).find(|t| !t.is_close() && t.local() == "sheetProtection")
        else {
            continue;
        };
        let span = tag.start..tag.end;
        let flag = |k: &str, default: bool| {
            scan::attr(&doc, span.clone(), k).map_or(default, |v| v == "1" || v == "true")
        };
        if !flag("sheet", false) {
            continue;
        }
        let mut allow = serde_json::Map::new();
        for key in ALLOWS {
            let default = !matches!(
                *key,
                "selectLockedCells" | "selectUnlockedCells" | "objects" | "scenarios"
            );
            allow.insert(
                key.to_string(),
                serde_json::Value::Bool(!flag(key, default)),
            );
        }
        let password = ["hashValue", "password"]
            .iter()
            .any(|k| scan::attr(&doc, span.clone(), k).is_some_and(|v| !v.is_empty()));
        out.insert(
            name,
            serde_json::json!({ "password": password, "allow": allow }),
        );
    }
    Ok(out)
}

/// The password is the one the sheet was protected with (any, when it has none).
fn password_fits(doc: &str, tag: std::ops::Range<usize>, password: &str) -> bool {
    let attr = |k: &str| scan::attr(doc, tag.clone(), k).filter(|v| !v.is_empty());
    if let (Some(hash), Some(salt)) = (attr("hashValue"), attr("saltValue")) {
        let spin = attr("spinCount").and_then(|s| s.parse().ok()).unwrap_or(0);
        let Some(salt) = unbase64(&salt) else {
            return false;
        };
        let algorithm = attr("algorithmName")
            .unwrap_or_default()
            .to_uppercase()
            .replace('-', "");
        let ours = match algorithm.as_str() {
            "SHA512" => sha512_spun(&salt, password, spin),
            "SHA384" => spun::<sha2::Sha384>(&salt, password, spin),
            "SHA256" => spun::<sha2::Sha256>(&salt, password, spin),
            _ => return false,
        };
        return base64(&ours) == hash.trim();
    }
    match attr("password") {
        Some(old) => u16::from_str_radix(old.trim(), 16).is_ok_and(|h| h == legacy_hash(password)),
        None => true,
    }
}

fn sha512_spun(salt: &[u8], password: &str, spin: u32) -> Vec<u8> {
    spun::<sha2::Sha512>(salt, password, spin)
}

/// ECMA-376's hash of a password: of the salt and the password (UTF-16LE), then of the hash and
/// the number of the round (`spin` rounds).
fn spun<D: sha2::Digest>(salt: &[u8], password: &str, spin: u32) -> Vec<u8> {
    let mut h = D::new();
    h.update(salt);
    let utf16: Vec<u8> = password.encode_utf16().flat_map(u16::to_le_bytes).collect();
    h.update(&utf16);
    let mut hash = h.finalize().to_vec();
    for i in 0..spin {
        let mut h = D::new();
        h.update(&hash);
        h.update(i.to_le_bytes());
        hash = h.finalize().to_vec();
    }
    hash
}

/// The 16-bit hash of the passwords of older files (`password`).
fn legacy_hash(password: &str) -> u16 {
    let bytes: Vec<u8> = password.chars().map(|c| c as u32 as u8).collect();
    let mut hash: u16 = 0;
    for b in bytes.iter().rev() {
        hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
        hash ^= u16::from(*b);
    }
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^= bytes.len() as u16;
    hash ^ 0xCE4B
}

/// 16 bytes no two protections share.
fn salt() -> Vec<u8> {
    use std::hash::{BuildHasher, Hasher};
    let mut out = Vec::with_capacity(16);
    let time = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or_default();
    for i in 0..2u8 {
        let mut h = std::collections::hash_map::RandomState::new().build_hasher();
        h.write_u128(time);
        h.write_u8(i);
        out.extend(h.finish().to_le_bytes());
    }
    out
}

fn base64(bytes: &[u8]) -> String {
    use base64::Engine;
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

fn unbase64(text: &str) -> Option<Vec<u8>> {
    use base64::Engine;
    base64::engine::general_purpose::STANDARD
        .decode(text.trim())
        .ok()
}

#[cfg(test)]
mod protect_tests {
    use super::*;

    #[test]
    fn legacy_hash_of_excel() {
        // (What Excel writes for "test".)
        assert_eq!(format!("{:04X}", legacy_hash("test")), "CBEB");
        assert_eq!(legacy_hash(""), 0xCE4B);
    }
}
