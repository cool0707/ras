//! The OPC (zip) package around a workbook: locating the workbook part,
//! the sheets and their code names, the VBA project part and its signature
//! parts, and the edits to `[Content_Types].xml`, relationship parts and
//! code name attributes.

use std::collections::{BTreeMap, HashSet};
use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::biff12;
use crate::error::OvbaError;
use crate::project::same_name;
use crate::xml::{self, Edits, TagKind};
use crate::{CodeNames, SheetCodeName};

pub(crate) const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
pub(crate) const CT_VBA_PROJECT: &str = "application/vnd.ms-office.vbaProject";
pub(crate) const CT_XLSX_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub(crate) const CT_XLSM_MAIN: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
pub(crate) const CT_XLTM_MAIN: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";
pub(crate) const CT_XLAM_MAIN: &str = "application/vnd.ms-excel.addin.macroEnabled.main+xml";
pub(crate) const REL_VBA_PROJECT: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
pub(crate) const DEFAULT_WORKBOOK_CODE_NAME: &str = "ThisWorkbook";

/// Kind of a sheet, taken from its relationship type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SheetKind {
    Worksheet,
    Chartsheet,
    Dialogsheet,
    Macrosheet,
}

/// A sheet of the workbook.
#[derive(Clone, Debug)]
pub(crate) struct SheetInfo {
    pub name: String,
    /// Zip entry name of the sheet part (empty when it could not be resolved).
    pub part: String,
    pub kind: SheetKind,
    /// Code name stored in the sheet part.
    pub explicit_code_name: Option<String>,
    /// Code name in effect: the explicit one, or the default that is assigned
    /// when a project is written.
    pub code_name: String,
}

/// What the package says about the workbook and its VBA project.
#[derive(Clone, Debug)]
pub(crate) struct BookLayout {
    /// Zip entry name of the workbook part (`xl/workbook.xml`).
    pub workbook_part: String,
    /// Zip entry name of its relationship part.
    pub workbook_rels_part: String,
    /// The workbook part is BIFF12 (`.xlsb`).
    pub binary: bool,
    pub explicit_workbook_code_name: Option<String>,
    pub workbook_code_name: String,
    pub sheets: Vec<SheetInfo>,
    /// Zip entry name of the VBA project part, if there is one.
    pub vba_part: Option<String>,
    /// Relationship part of the VBA project part, if present.
    pub vba_rels_part: Option<String>,
    /// Signature parts of the VBA project.
    pub signature_parts: Vec<String>,
}

impl BookLayout {
    pub(crate) fn code_names(&self) -> CodeNames {
        CodeNames {
            workbook: self.workbook_code_name.clone(),
            sheets: self
                .sheets
                .iter()
                .map(|s| SheetCodeName { sheet_name: s.name.clone(), code_name: s.code_name.clone() })
                .collect(),
        }
    }
}

/// A relationship of a relationship part.
#[derive(Clone, Debug)]
struct Relationship {
    id: String,
    rel_type: String,
    target: String,
    external: bool,
}

/// An opened package.
pub(crate) struct Package {
    archive: ZipArchive<Cursor<Vec<u8>>>,
    names: Vec<String>,
}

impl Package {
    pub(crate) fn open(bytes: Vec<u8>) -> Result<Package, OvbaError> {
        if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
            return Err(OvbaError::UnsupportedFormat(
                "the file is a compound file (a legacy .xls workbook or a password-encrypted workbook), not an Open XML package".into(),
            ));
        }
        let archive = ZipArchive::new(Cursor::new(bytes))?;
        let names = (0..archive.len())
            .map(|i| archive.name_for_index(i).unwrap_or("").to_string())
            .collect();
        Ok(Package { archive, names })
    }

    /// Finds the zip entry for a part name (part names are case-insensitive).
    pub(crate) fn find(&self, part: &str) -> Option<&str> {
        let part = part.trim_start_matches('/');
        self.names
            .iter()
            .find(|n| n.as_str() == part)
            .or_else(|| self.names.iter().find(|n| n.eq_ignore_ascii_case(part)))
            .map(|s| s.as_str())
    }

    pub(crate) fn read(&mut self, part: &str) -> Result<Option<Vec<u8>>, OvbaError> {
        let Some(name) = self.find(part).map(|s| s.to_string()) else {
            return Ok(None);
        };
        let mut file = self.archive.by_name(&name)?;
        let mut data = Vec::with_capacity(file.size().min(1 << 26) as usize);
        file.read_to_end(&mut data)?;
        Ok(Some(data))
    }

    fn read_text(&mut self, part: &str) -> Result<Option<(String, bool)>, OvbaError> {
        match self.read(part)? {
            Some(bytes) => Ok(Some(xml::decode_part(&bytes, part)?)),
            None => Ok(None),
        }
    }

    /// Reads the beginning of a part (enough for the first elements of a
    /// sheet) as text.
    fn read_head(&mut self, part: &str, limit: usize) -> Result<Option<String>, OvbaError> {
        let Some(name) = self.find(part).map(|s| s.to_string()) else {
            return Ok(None);
        };
        let file = self.archive.by_name(&name)?;
        let mut data = Vec::with_capacity(limit.min(1 << 16));
        file.take(limit as u64).read_to_end(&mut data)?;
        if data.starts_with(&[0xFF, 0xFE]) || data.starts_with(&[0xFE, 0xFF]) {
            let len = data.len() & !1;
            return Ok(Some(xml::decode_part(&data[..len], part)?.0));
        }
        Ok(Some(String::from_utf8_lossy(&data).into_owned()))
    }

    /// Writes a new package: entries listed in `replace` get new contents
    /// (in place, same position), entries in `remove` are dropped, every
    /// other entry is copied raw (no recompression) and `add` is appended.
    pub(crate) fn rewrite(
        &mut self,
        replace: &BTreeMap<String, Vec<u8>>,
        remove: &HashSet<String>,
        add: &[(String, Vec<u8>)],
    ) -> Result<Vec<u8>, OvbaError> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        writer.set_raw_comment(self.archive.comment().to_vec().into_boxed_slice());
        for index in 0..self.archive.len() {
            let name = self.names[index].clone();
            if remove.contains(&name) {
                continue;
            }
            if let Some(data) = replace.get(&name) {
                let options = {
                    let original = self.archive.by_index_raw(index)?;
                    let mut options = SimpleFileOptions::default()
                        .compression_method(match original.compression() {
                            CompressionMethod::Stored => CompressionMethod::Stored,
                            _ => CompressionMethod::Deflated,
                        });
                    if let Some(time) = original.last_modified() {
                        options = options.last_modified_time(time);
                    }
                    options
                };
                writer.start_file(name.as_str(), options)?;
                writer.write_all(data)?;
            } else {
                let file = self.archive.by_index_raw(index)?;
                writer.raw_copy_file(file)?;
            }
        }
        for (name, data) in add {
            let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            writer.start_file(name.as_str(), options)?;
            writer.write_all(data)?;
        }
        Ok(writer.finish()?.into_inner())
    }
}

/// Directory of a part (`xl/workbook.xml` -> `xl`).
fn part_dir(part: &str) -> &str {
    match part.rfind('/') {
        Some(i) => &part[..i],
        None => "",
    }
}

fn part_file(part: &str) -> &str {
    match part.rfind('/') {
        Some(i) => &part[i + 1..],
        None => part,
    }
}

/// Name of the relationship part that belongs to `part`.
pub(crate) fn rels_part_for(part: &str) -> String {
    let dir = part_dir(part);
    if dir.is_empty() {
        format!("_rels/{}.rels", part_file(part))
    } else {
        format!("{dir}/_rels/{}.rels", part_file(part))
    }
}

/// Resolves a relationship target against the part that owns the
/// relationship part. Returns a zip entry name (no leading slash).
fn resolve_target(source_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    let mut segments: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        part_dir(source_part).split('/').filter(|s| !s.is_empty()).collect()
    };
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

fn parse_relationships(doc: &str) -> Vec<Relationship> {
    xml::tags(doc)
        .filter(|t| t.kind != TagKind::Close && t.local_name() == "Relationship")
        .map(|t| Relationship {
            id: t.attr("Id").unwrap_or_default(),
            rel_type: t.attr("Type").unwrap_or_default(),
            target: t.attr("Target").unwrap_or_default(),
            external: t.attr("TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External")),
        })
        .collect()
}

fn sheet_kind(rel_type: &str) -> Option<SheetKind> {
    let last = rel_type.rsplit('/').next().unwrap_or("");
    match last {
        "worksheet" => Some(SheetKind::Worksheet),
        "chartsheet" => Some(SheetKind::Chartsheet),
        "dialogsheet" => Some(SheetKind::Dialogsheet),
        "xlMacrosheet" | "xlIntlMacrosheet" => Some(SheetKind::Macrosheet),
        _ => None,
    }
}

fn is_signature_type(rel_type: &str) -> bool {
    rel_type.rsplit('/').next().is_some_and(|last| last.starts_with("vbaProjectSignature"))
}

fn is_signature_part_name(name: &str) -> bool {
    let file = part_file(name).to_ascii_lowercase();
    file.starts_with("vbaprojectsignature") && file.ends_with(".bin")
}

/// `codeName` of the `sheetPr` element at the beginning of a sheet part.
fn sheet_code_name_xml(head: &str) -> Option<String> {
    let mut seen_root = false;
    for tag in xml::tags(head) {
        if tag.kind == TagKind::Close {
            continue;
        }
        if !seen_root {
            seen_root = true;
            continue;
        }
        // sheetPr is the first child of the sheet's root element.
        if tag.local_name() == "sheetPr" {
            return tag.attr("codeName").filter(|n| !n.is_empty());
        }
        return None;
    }
    None
}

/// Analyses the package.
pub(crate) fn analyze(package: &mut Package) -> Result<BookLayout, OvbaError> {
    // Main part: the officeDocument relationship of the package.
    let Some((root_rels, _)) = package.read_text("_rels/.rels")? else {
        return Err(OvbaError::Format("the package has no _rels/.rels part".into()));
    };
    let workbook_target = parse_relationships(&root_rels)
        .into_iter()
        .find(|r| !r.external && r.rel_type.ends_with("/officeDocument"))
        .map(|r| resolve_target("", &r.target));
    let Some(workbook_part) = workbook_target.and_then(|t| package.find(&t).map(|s| s.to_string()))
    else {
        return Err(OvbaError::Format("the package has no workbook part".into()));
    };
    let binary = workbook_part.to_ascii_lowercase().ends_with(".bin");
    let workbook_rels_part = match package.find(&rels_part_for(&workbook_part)) {
        Some(name) => name.to_string(),
        None => rels_part_for(&workbook_part),
    };
    let relationships = match package.read_text(&workbook_rels_part)? {
        Some((text, _)) => parse_relationships(&text),
        None => Vec::new(),
    };

    // Sheets (name + relationship id) and the workbook code name.
    let mut explicit_workbook_code_name = None;
    let mut listed: Vec<(String, String)> = Vec::new();
    if binary {
        let data = package.read(&workbook_part)?.unwrap_or_default();
        let (code_name, sheets) = biff12::workbook_info(&data);
        explicit_workbook_code_name = code_name;
        listed = sheets.into_iter().map(|s| (s.name, s.rel_id)).collect();
    } else {
        let (text, _) = package
            .read_text(&workbook_part)?
            .ok_or_else(|| OvbaError::Format("the workbook part is missing".into()))?;
        for tag in xml::tags(&text) {
            if tag.kind == TagKind::Close {
                continue;
            }
            match tag.local_name() {
                "workbookPr" => {
                    explicit_workbook_code_name = tag.attr("codeName").filter(|n| !n.is_empty());
                }
                "sheet" => {
                    let name = tag.attr("name").unwrap_or_default();
                    let rel_id = tag
                        .attributes()
                        .into_iter()
                        .find(|a| a.local_name() == "id" && a.name != "id")
                        .map(|a| a.value())
                        .unwrap_or_default();
                    listed.push((name, rel_id));
                }
                _ => {}
            }
        }
    }

    let mut sheets = Vec::with_capacity(listed.len());
    for (name, rel_id) in listed {
        let relationship = relationships.iter().find(|r| r.id == rel_id && !r.external);
        let kind = relationship.and_then(|r| sheet_kind(&r.rel_type)).unwrap_or(SheetKind::Worksheet);
        let part = relationship
            .map(|r| resolve_target(&workbook_part, &r.target))
            .and_then(|t| package.find(&t).map(|s| s.to_string()))
            .unwrap_or_default();
        let explicit_code_name = if part.is_empty() {
            None
        } else if part.to_ascii_lowercase().ends_with(".bin") {
            match kind {
                SheetKind::Worksheet => {
                    let data = package.read(&part)?.unwrap_or_default();
                    biff12::worksheet_code_name(&data)
                }
                _ => None,
            }
        } else {
            package.read_head(&part, 16 * 1024)?.and_then(|head| sheet_code_name_xml(&head))
        };
        sheets.push(SheetInfo { name, part, kind, explicit_code_name, code_name: String::new() });
    }

    // Effective code names: explicit ones, then defaults for the rest.
    let workbook_code_name = explicit_workbook_code_name
        .clone()
        .unwrap_or_else(|| DEFAULT_WORKBOOK_CODE_NAME.to_string());
    let mut used: Vec<String> = vec![workbook_code_name.clone()];
    used.extend(sheets.iter().filter_map(|s| s.explicit_code_name.clone()));
    for sheet in sheets.iter_mut() {
        if let Some(name) = &sheet.explicit_code_name {
            sheet.code_name = name.clone();
            continue;
        }
        let stem = match sheet.kind {
            SheetKind::Worksheet => "Sheet",
            SheetKind::Chartsheet => "Chart",
            SheetKind::Dialogsheet => "Dialog",
            SheetKind::Macrosheet => "Macro",
        };
        let mut n = 1usize;
        let candidate = loop {
            let candidate = format!("{stem}{n}");
            if !used.iter().any(|u| same_name(u, &candidate)) {
                break candidate;
            }
            n += 1;
        };
        used.push(candidate.clone());
        sheet.code_name = candidate;
    }

    // VBA project part and its signatures.
    let vba_part = relationships
        .iter()
        .find(|r| !r.external && r.rel_type.ends_with("/vbaProject"))
        .map(|r| resolve_target(&workbook_part, &r.target))
        .and_then(|t| package.find(&t).map(|s| s.to_string()));
    let mut vba_rels_part = None;
    let mut signature_parts: Vec<String> = Vec::new();
    if let Some(vba) = &vba_part {
        if let Some(rels_name) = package.find(&rels_part_for(vba)).map(|s| s.to_string()) {
            if let Some((text, _)) = package.read_text(&rels_name)? {
                for relationship in parse_relationships(&text) {
                    if relationship.external || !is_signature_type(&relationship.rel_type) {
                        continue;
                    }
                    let target = resolve_target(vba, &relationship.target);
                    if let Some(name) = package.find(&target) {
                        if !signature_parts.iter().any(|p| p == name) {
                            signature_parts.push(name.to_string());
                        }
                    }
                }
            }
            vba_rels_part = Some(rels_name);
        }
        // Signature parts that are not (or wrongly) referenced.
        let dir = part_dir(vba).to_string();
        for name in &package.names {
            if part_dir(name) == dir && is_signature_part_name(name) && !signature_parts.contains(name) {
                signature_parts.push(name.clone());
            }
        }
    }

    Ok(BookLayout {
        workbook_part,
        workbook_rels_part,
        binary,
        explicit_workbook_code_name,
        workbook_code_name,
        sheets,
        vba_part,
        vba_rels_part,
        signature_parts,
    })
}

fn read_utf8_part(package: &mut Package, part: &str) -> Result<Option<String>, OvbaError> {
    match package.read_text(part)? {
        Some((text, true)) => Ok(Some(text)),
        Some((_, false)) => Err(OvbaError::UnsupportedFormat(format!(
            "package part `{part}` is UTF-16 encoded and cannot be edited"
        ))),
        None => Ok(None),
    }
}

/// The changes a package needs so that it carries the VBA project.
pub(crate) struct PackagePlan {
    /// Target content type of the workbook part (`None`: leave unchanged).
    pub main_content_type: Option<&'static str>,
    /// Remove the signature parts (the project changed).
    pub remove_signatures: bool,
    /// Write missing `codeName` attributes into the workbook and sheet parts.
    pub write_code_names: bool,
}

/// Result of [`plan_package_edits`].
pub(crate) struct PackageEdits {
    pub replace: BTreeMap<String, Vec<u8>>,
    pub remove: HashSet<String>,
    pub add: Vec<(String, Vec<u8>)>,
    /// Zip entry name of the VBA project part after the edit.
    pub vba_part: String,
}

/// Computes the new contents of every XML part that has to change. The
/// VBA project part itself is added to the result by the caller.
pub(crate) fn plan_package_edits(
    package: &mut Package,
    layout: &BookLayout,
    plan: &PackagePlan,
) -> Result<PackageEdits, OvbaError> {
    let mut replace: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut remove: HashSet<String> = HashSet::new();
    let mut add: Vec<(String, Vec<u8>)> = Vec::new();

    let vba_part = match &layout.vba_part {
        Some(part) => part.clone(),
        None => {
            let dir = part_dir(&layout.workbook_part);
            if dir.is_empty() {
                "vbaProject.bin".to_string()
            } else {
                format!("{dir}/vbaProject.bin")
            }
        }
    };
    let removed_signatures: Vec<String> =
        if plan.remove_signatures { layout.signature_parts.clone() } else { Vec::new() };
    remove.extend(removed_signatures.iter().cloned());

    // ---- [Content_Types].xml ------------------------------------------------
    let content_types_name = package
        .find(CONTENT_TYPES_PART)
        .map(|s| s.to_string())
        .ok_or_else(|| OvbaError::Format("the package has no [Content_Types].xml".into()))?;
    let content_types = read_utf8_part(package, &content_types_name)?.unwrap_or_default();
    {
        let mut edits = Edits::default();
        let mut last_default_end = None;
        let mut types_open_end = None;
        let mut types_close_start = None;
        let mut bin_default: Option<String> = None;
        let mut vba_override = false;
        let mut main_override_found = false;
        let workbook_part_name = format!("/{}", layout.workbook_part);
        let vba_part_name = format!("/{vba_part}");
        for tag in xml::tags(&content_types) {
            match (tag.local_name(), tag.kind) {
                ("Types", TagKind::Open) => types_open_end = Some(tag.end),
                ("Types", TagKind::Close) => types_close_start = Some(tag.start),
                ("Default", TagKind::Open | TagKind::Empty) => {
                    last_default_end = Some(tag.end);
                    if tag.attr("Extension").is_some_and(|e| e.eq_ignore_ascii_case("bin")) {
                        bin_default = tag.attr("ContentType");
                    }
                }
                ("Override", TagKind::Open | TagKind::Empty) => {
                    let part_name = tag.attr("PartName").unwrap_or_default();
                    if part_name.eq_ignore_ascii_case(&workbook_part_name) {
                        main_override_found = true;
                        if let (Some(target), Some(attribute)) =
                            (plan.main_content_type, tag.attribute("ContentType"))
                        {
                            if attribute.value() != target {
                                edits.replace(attribute.value_start, attribute.value_end, target);
                            }
                        }
                    } else if part_name.eq_ignore_ascii_case(&vba_part_name) {
                        vba_override = true;
                        if let Some(attribute) = tag.attribute("ContentType") {
                            if attribute.value() != CT_VBA_PROJECT {
                                edits.replace(attribute.value_start, attribute.value_end, CT_VBA_PROJECT);
                            }
                        }
                    } else if removed_signatures
                        .iter()
                        .any(|s| part_name.eq_ignore_ascii_case(&format!("/{s}")))
                    {
                        edits.replace(tag.start, tag.end, "");
                    }
                }
                _ => {}
            }
        }
        let (Some(open_end), Some(close_start)) = (types_open_end, types_close_start) else {
            return Err(OvbaError::Format("[Content_Types].xml has no Types element".into()));
        };
        if !vba_override && bin_default.as_deref() != Some(CT_VBA_PROJECT) {
            if bin_default.is_none() {
                // Like Excel: a Default for the `bin` extension.
                edits.insert(
                    last_default_end.unwrap_or(open_end),
                    format!("<Default Extension=\"bin\" ContentType=\"{CT_VBA_PROJECT}\"/>"),
                );
            } else {
                // `bin` already means something else (printer settings ...).
                edits.insert(
                    close_start,
                    format!(
                        "<Override PartName=\"{}\" ContentType=\"{CT_VBA_PROJECT}\"/>",
                        xml::escape_attr(&vba_part_name)
                    ),
                );
            }
        }
        if !main_override_found {
            if let Some(target) = plan.main_content_type {
                edits.insert(
                    close_start,
                    format!(
                        "<Override PartName=\"{}\" ContentType=\"{target}\"/>",
                        xml::escape_attr(&workbook_part_name)
                    ),
                );
            }
        }
        if !edits.is_empty() {
            replace.insert(content_types_name, edits.apply(&content_types).into_bytes());
        }
    }

    // ---- workbook relationships ----------------------------------------------
    if layout.vba_part.is_none() {
        let target = part_file(&vba_part).to_string();
        match read_utf8_part(package, &layout.workbook_rels_part)? {
            Some(text) => {
                let existing = parse_relationships(&text);
                let mut n = existing.len() + 1;
                let id = loop {
                    let candidate = format!("rId{n}");
                    if !existing.iter().any(|r| r.id == candidate) {
                        break candidate;
                    }
                    n += 1;
                };
                let close = xml::tags(&text)
                    .filter(|t| t.kind == TagKind::Close && t.local_name() == "Relationships")
                    .last()
                    .map(|t| t.start)
                    .ok_or_else(|| {
                        OvbaError::Format(format!(
                            "`{}` has no Relationships element",
                            layout.workbook_rels_part
                        ))
                    })?;
                let mut edits = Edits::default();
                edits.insert(
                    close,
                    format!("<Relationship Id=\"{id}\" Type=\"{REL_VBA_PROJECT}\" Target=\"{target}\"/>"),
                );
                replace.insert(layout.workbook_rels_part.clone(), edits.apply(&text).into_bytes());
            }
            None => {
                let text = format!(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_VBA_PROJECT}\" Target=\"{target}\"/></Relationships>"
                );
                add.push((layout.workbook_rels_part.clone(), text.into_bytes()));
            }
        }
    }

    // ---- relationships of the VBA project part (signatures) ---------------
    if !removed_signatures.is_empty() {
        if let Some(rels_name) = &layout.vba_rels_part {
            if let Some(text) = read_utf8_part(package, rels_name)? {
                let mut edits = Edits::default();
                let mut kept = 0usize;
                for tag in xml::tags(&text) {
                    if tag.kind == TagKind::Close || tag.local_name() != "Relationship" {
                        continue;
                    }
                    if is_signature_type(&tag.attr("Type").unwrap_or_default()) {
                        edits.replace(tag.start, tag.end, "");
                    } else {
                        kept += 1;
                    }
                }
                if kept == 0 {
                    remove.insert(rels_name.clone());
                } else if !edits.is_empty() {
                    replace.insert(rels_name.clone(), edits.apply(&text).into_bytes());
                }
            }
        }
    }

    // ---- code names -------------------------------------------------------------
    if plan.write_code_names && !layout.binary {
        if layout.explicit_workbook_code_name.is_none() {
            if let Some(text) = read_utf8_part(package, &layout.workbook_part)? {
                let code_name = xml::escape_attr(&layout.workbook_code_name);
                let mut edits = Edits::default();
                let mut done = false;
                let mut fallback: Option<(usize, String)> = None;
                for tag in xml::tags(&text) {
                    if tag.kind == TagKind::Close {
                        continue;
                    }
                    match tag.local_name() {
                        "workbookPr" => {
                            match tag.attribute("codeName") {
                                Some(attribute) => {
                                    edits.replace(attribute.value_start, attribute.value_end, code_name.clone())
                                }
                                None => edits.insert(tag.name_end, format!(" codeName=\"{code_name}\"")),
                            }
                            done = true;
                            break;
                        }
                        // workbookPr precedes these in the schema.
                        "workbookProtection" | "bookViews" | "sheets" if fallback.is_none() => {
                            fallback = Some((
                                tag.start,
                                format!("<{}workbookPr codeName=\"{code_name}\"/>", tag.prefix()),
                            ));
                        }
                        _ => {}
                    }
                }
                if !done {
                    if let Some((at, element)) = fallback {
                        edits.insert(at, element);
                    }
                }
                if !edits.is_empty() {
                    replace.insert(layout.workbook_part.clone(), edits.apply(&text).into_bytes());
                }
            }
        }
        for sheet in &layout.sheets {
            if sheet.explicit_code_name.is_some() || sheet.part.is_empty() {
                continue;
            }
            if sheet.part.to_ascii_lowercase().ends_with(".bin") {
                continue;
            }
            let Some(text) = read_utf8_part(package, &sheet.part)? else {
                continue;
            };
            let code_name = xml::escape_attr(&sheet.code_name);
            let mut edits = Edits::default();
            let mut root: Option<(usize, String, bool)> = None;
            for tag in xml::tags(&text) {
                if tag.kind == TagKind::Close {
                    continue;
                }
                match &root {
                    None => {
                        root = Some((tag.end, tag.prefix().to_string(), tag.kind == TagKind::Empty));
                        if tag.kind == TagKind::Empty {
                            break;
                        }
                    }
                    Some((root_end, prefix, _)) => {
                        if tag.local_name() == "sheetPr" {
                            match tag.attribute("codeName") {
                                Some(attribute) => {
                                    edits.replace(attribute.value_start, attribute.value_end, code_name.clone())
                                }
                                None => edits.insert(tag.name_end, format!(" codeName=\"{code_name}\"")),
                            }
                        } else {
                            edits.insert(*root_end, format!("<{prefix}sheetPr codeName=\"{code_name}\"/>"));
                        }
                        break;
                    }
                }
            }
            if !edits.is_empty() {
                replace.insert(sheet.part.clone(), edits.apply(&text).into_bytes());
            }
        }
    }

    Ok(PackageEdits { replace, remove, add, vba_part })
}

/// Reads the workbook code name and the code name of every sheet.
///
/// Code names come from `workbookPr/@codeName` in the workbook part and
/// `sheetPr/@codeName` in each sheet part (for `.xlsb` from the
/// corresponding BIFF12 records). Where a part does not record a code
/// name, the name that [`crate::pack`] assigns when it writes a project is
/// returned: `ThisWorkbook` for the workbook and the first free `Sheet<n>`
/// (`Chart<n>` for chart sheets) for a sheet, in sheet order.
pub fn workbook_code_names(book: &std::path::Path) -> Result<CodeNames, OvbaError> {
    let bytes = std::fs::read(book).map_err(crate::pack::map_open_error)?;
    let mut package = Package::open(bytes)?;
    Ok(analyze(&mut package)?.code_names())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn target_resolution() {
        assert_eq!(resolve_target("xl/workbook.xml", "worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml");
        assert_eq!(resolve_target("xl/workbook.xml", "/xl/vbaProject.bin"), "xl/vbaProject.bin");
        assert_eq!(resolve_target("xl/worksheets/sheet1.xml", "../drawings/d.xml"), "xl/drawings/d.xml");
        assert_eq!(resolve_target("", "xl/workbook.xml"), "xl/workbook.xml");
        assert_eq!(rels_part_for("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
        assert_eq!(rels_part_for("workbook.xml"), "_rels/workbook.xml.rels");
    }

    #[test]
    fn sheet_pr_code_name() {
        let ws = "<?xml version=\"1.0\"?><worksheet xmlns=\"x\"><sheetPr codeName=\"Data\"><tabColor rgb=\"FF\"/></sheetPr><dimension ref=\"A1\"/>";
        assert_eq!(sheet_code_name_xml(ws).as_deref(), Some("Data"));
        assert_eq!(sheet_code_name_xml("<worksheet><dimension ref=\"A1\"/><sheetPr codeName=\"X\"/>"), None);
        assert_eq!(sheet_code_name_xml("<x:chartsheet><x:sheetPr codeName=\"Chart1\"/>").as_deref(), Some("Chart1"));
        assert_eq!(sheet_code_name_xml("<worksheet><sheetPr/>"), None);
        assert_eq!(sheet_code_name_xml(""), None);
    }

    #[test]
    fn signature_names() {
        assert!(is_signature_part_name("xl/vbaProjectSignature.bin"));
        assert!(is_signature_part_name("xl/vbaProjectSignatureAgile.bin"));
        assert!(is_signature_part_name("xl/VBAPROJECTSIGNATUREV3.BIN"));
        assert!(!is_signature_part_name("xl/vbaProject.bin"));
        assert!(is_signature_type("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"));
        assert!(is_signature_type("http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile"));
        assert!(!is_signature_type("http://schemas.microsoft.com/office/2006/relationships/vbaProject"));
    }
}
