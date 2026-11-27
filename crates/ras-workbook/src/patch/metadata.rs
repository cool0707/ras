//! The metadata that marks a formula as a dynamic array (`cm` of a cell, `xl/metadata.xml`):
//! without it Excel reads `<f t="array">` as a legacy array formula (`{=...}`), which does not
//! spill.

use std::cell::Cell;
use std::collections::BTreeMap;

use super::scan::{self, Edits};
use super::{Pack, Result, WORKBOOK};

const PART: &str = "xl/metadata.xml";
const RELS: &str = "xl/_rels/workbook.xml.rels";
const TYPES: &str = "[Content_Types].xml";
const REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";

/// The metadata Excel writes for a workbook with dynamic arrays: one cell metadata, `cm="1"`.
const DOC: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>";

/// The `cm` for the dynamic-array formulas written anew, and whether the file needs the
/// metadata added for it.
pub(super) struct DynamicCm {
    pub(super) cm: Option<String>,
    add: bool,
    used: Cell<bool>,
}

impl DynamicCm {
    /// What the file has: its cell metadata of dynamic arrays, or none (the metadata is then
    /// added when a formula needs it). A file with metadata of other kinds only is left as it
    /// is (the formula is written without `cm`).
    pub(super) fn of(pack: &mut Pack) -> Result<DynamicCm> {
        let part = pack
            .rels(WORKBOOK)?
            .into_iter()
            .find(|r| r.1.ends_with("/sheetMetadata"))
            .and_then(|r| pack.find(&r.2));
        let (cm, add) = match part {
            None if pack.find(PART).is_none() => (Some("1".to_string()), true),
            None => (None, false),
            Some(part) => (pack.text(&part)?.as_deref().and_then(dynamic_index), false),
        };
        Ok(DynamicCm {
            cm,
            add,
            used: Cell::new(false),
        })
    }

    /// The `cm` of a dynamic-array formula of a cell without one of its own.
    pub(super) fn take(&self) -> Option<&str> {
        self.used.set(true);
        self.cm.as_deref()
    }

    /// Adds the metadata to the file (with its relationship and content type) when a formula
    /// was given its `cm`.
    pub(super) fn finish(
        &self,
        pack: &mut Pack,
        replace: &mut BTreeMap<String, Vec<u8>>,
    ) -> Result<()> {
        if !(self.add && self.used.get()) {
            return Ok(());
        }
        let mut text = |part: &str| -> Result<Option<(String, String)>> {
            let Some(name) = pack.find(part) else {
                return Ok(None);
            };
            if let Some(data) = replace.get(&name) {
                return Ok(Some((name, String::from_utf8_lossy(data).into_owned())));
            }
            Ok(pack.text(&name)?.map(|t| (name, t)))
        };
        let (Some((rels_name, rels)), Some((types_name, types))) = (text(RELS)?, text(TYPES)?)
        else {
            return Ok(());
        };
        let used: Vec<String> = scan::tags(&rels, 0)
            .filter(|t| !t.is_close() && t.local() == "Relationship")
            .filter_map(|t| scan::attr(&rels, t.start..t.end, "Id"))
            .collect();
        let id = (1..)
            .map(|n| format!("rId{n}"))
            .find(|id| !used.contains(id))
            .unwrap_or_default();
        let rels_root = scan::tree(&rels, &[]).ok_or("workbook.xml.rels を読み取れません")?;
        let p = scan::prefix(rels_root.name);
        let mut edits = Edits::default();
        edits.insert(
            rels_root.inner.end,
            format!("<{p}Relationship Id=\"{id}\" Type=\"{REL_TYPE}\" Target=\"metadata.xml\"/>"),
        );
        replace.insert(rels_name, edits.apply(&rels).into_bytes());
        let types_root = scan::tree(&types, &[]).ok_or("[Content_Types].xml を読み取れません")?;
        let p = scan::prefix(types_root.name);
        let mut edits = Edits::default();
        edits.insert(
            types_root.inner.end,
            format!("<{p}Override PartName=\"/{PART}\" ContentType=\"{CONTENT_TYPE}\"/>"),
        );
        replace.insert(types_name, edits.apply(&types).into_bytes());
        replace.insert(PART.to_string(), DOC.as_bytes().to_vec());
        Ok(())
    }
}

/// The number (from 1) of the cell metadata of dynamic arrays (`XLDAPR` with `fDynamic="1"`).
fn dynamic_index(doc: &str) -> Option<String> {
    let tags: Vec<scan::Tag> = scan::tags(doc, 0).collect();
    let attr = |t: &scan::Tag, name: &str| scan::attr(doc, t.start..t.end, name);
    // The metadata type (`t`, from 1).
    let kind = tags
        .iter()
        .filter(|t| !t.is_close() && t.local() == "metadataType")
        .position(|t| attr(t, "name").as_deref() == Some("XLDAPR"))?
        + 1;
    // The blocks of the future metadata XLDAPR (`v`, from 0) that are dynamic.
    let (mut in_future, mut in_cells) = (false, false);
    let (mut future, mut block) = (Vec::new(), None::<usize>);
    let mut cell_blocks = 0;
    for t in &tags {
        match (t.local(), t.is_close()) {
            ("futureMetadata", false) => in_future = attr(t, "name").as_deref() == Some("XLDAPR"),
            ("futureMetadata", true) => in_future = false,
            ("cellMetadata", false) => in_cells = true,
            ("cellMetadata", true) => in_cells = false,
            ("bk", false) if in_future => {
                block = Some(block.map_or(0, |b| b + 1));
                future.push(false);
            }
            ("bk", false) if in_cells => cell_blocks += 1,
            ("dynamicArrayProperties", false) if in_future => {
                if let (Some(b), Some("1" | "true")) = (block, attr(t, "fDynamic").as_deref()) {
                    future[b] = true;
                }
            }
            ("rc", false) if in_cells => {
                let t_ = attr(t, "t").and_then(|v| v.parse::<usize>().ok());
                let v = attr(t, "v").and_then(|v| v.parse::<usize>().ok());
                if t_ == Some(kind) && v.is_some_and(|v| future.get(v) == Some(&true)) {
                    return Some(cell_blocks.to_string());
                }
            }
            _ => {}
        }
    }
    None
}

#[cfg(test)]
mod tests {
    #[test]
    fn the_dynamic_array_metadata() {
        assert_eq!(super::dynamic_index(super::DOC).as_deref(), Some("1"));
        let two = super::DOC.replace(
            "<cellMetadata count=\"1\"><bk>",
            "<cellMetadata count=\"2\"><bk><rc t=\"2\" v=\"0\"/></bk><bk>",
        );
        assert_eq!(super::dynamic_index(&two).as_deref(), Some("2"));
    }
}
