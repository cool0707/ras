//! Creation of a minimal, empty workbook package.

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::opc::{CT_XLAM_MAIN, CT_XLSM_MAIN, CT_XLSX_MAIN};
use crate::xml::escape_attr;

/// The flavour of workbook [`create_workbook`] produces.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum BookKind {
    /// Macro-enabled workbook (`.xlsm`).
    Xlsm,
    /// Add-in (`.xlam`).
    Xlam,
    /// Plain workbook (`.xlsx`), cannot hold a VBA project.
    Xlsx,
}

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// Makes sheet names acceptable to Excel: 1 to 31 characters, none of
/// `: \ / ? * [ ]`, no leading or trailing apostrophe, unique ignoring case.
fn sanitize_sheet_names(names: &[&str]) -> Vec<String> {
    let mut result: Vec<String> = Vec::new();
    for (index, raw) in names.iter().enumerate() {
        let mut name: String = raw
            .chars()
            .map(|c| {
                if matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']') || c.is_control() {
                    '_'
                } else {
                    c
                }
            })
            .collect();
        name = name.trim_matches('\'').to_string();
        if name.encode_utf16().count() > 31 {
            let mut units = 0usize;
            name = name
                .chars()
                .take_while(|c| {
                    units += c.len_utf16();
                    units <= 31
                })
                .collect();
        }
        if name.trim().is_empty() {
            name = format!("Sheet{}", index + 1);
        }
        let base = name.clone();
        let mut n = 2usize;
        while result
            .iter()
            .any(|existing| existing.to_lowercase() == name.to_lowercase())
        {
            let suffix = format!(" ({n})");
            let keep: String = base
                .chars()
                .take(31usize.saturating_sub(suffix.chars().count()))
                .collect();
            name = format!("{keep}{suffix}");
            n += 1;
        }
        result.push(name);
    }
    if result.is_empty() {
        result.push("Sheet1".to_string());
    }
    result
}

/// Builds a minimal valid, empty workbook (the bytes of the OPC package)
/// with the given sheet names.
///
/// Sheets get the code names `Sheet1`, `Sheet2`, ... in order, the workbook
/// the code name `ThisWorkbook`. The package has no VBA project; use
/// [`crate::pack`] to add one. Invalid sheet names are sanitised and an
/// empty list yields a single sheet `Sheet1`.
pub fn create_workbook(kind: BookKind, sheet_names: &[&str]) -> Vec<u8> {
    let names = sanitize_sheet_names(sheet_names);
    let main_type = match kind {
        BookKind::Xlsm => CT_XLSM_MAIN,
        BookKind::Xlam => CT_XLAM_MAIN,
        BookKind::Xlsx => CT_XLSX_MAIN,
    };

    let mut content_types = format!(
        "{XML_DECL}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{main_type}\"/>"
    );
    let mut workbook = format!(
        "{XML_DECL}<workbook xmlns=\"{NS_MAIN}\" xmlns:r=\"{NS_REL}\"><workbookPr codeName=\"ThisWorkbook\"/><bookViews><workbookView/></bookViews><sheets>"
    );
    let mut workbook_rels = format!("{XML_DECL}<Relationships xmlns=\"{NS_PKG_REL}\">");
    let mut parts: Vec<(String, String)> = Vec::new();
    for (index, name) in names.iter().enumerate() {
        let n = index + 1;
        content_types.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        ));
        workbook.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{n}\" r:id=\"rId{n}\"/>",
            escape_attr(name)
        ));
        workbook_rels.push_str(&format!(
            "<Relationship Id=\"rId{n}\" Type=\"{NS_REL}/worksheet\" Target=\"worksheets/sheet{n}.xml\"/>"
        ));
        parts.push((
            format!("xl/worksheets/sheet{n}.xml"),
            format!(
                "{XML_DECL}<worksheet xmlns=\"{NS_MAIN}\"><sheetPr codeName=\"Sheet{n}\"/><sheetData/></worksheet>"
            ),
        ));
    }
    let styles_id = names.len() + 1;
    content_types.push_str("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/></Types>");
    workbook.push_str("</sheets></workbook>");
    workbook_rels.push_str(&format!(
        "<Relationship Id=\"rId{styles_id}\" Type=\"{NS_REL}/styles\" Target=\"styles.xml\"/></Relationships>"
    ));
    let root_rels = format!(
        "{XML_DECL}<Relationships xmlns=\"{NS_PKG_REL}\"><Relationship Id=\"rId1\" Type=\"{NS_REL}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
    );
    let styles = format!(
        "{XML_DECL}<styleSheet xmlns=\"{NS_MAIN}\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>"
    );

    let mut entries: Vec<(String, String)> = vec![
        ("[Content_Types].xml".to_string(), content_types),
        ("_rels/.rels".to_string(), root_rels),
        ("xl/workbook.xml".to_string(), workbook),
        ("xl/_rels/workbook.xml.rels".to_string(), workbook_rels),
    ];
    entries.extend(parts);
    entries.push(("xl/styles.xml".to_string(), styles));

    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, text) in &entries {
        // Writing into memory cannot fail; should it ever, an empty result
        // is returned instead of panicking.
        if writer.start_file(name.as_str(), options).is_err()
            || writer.write_all(text.as_bytes()).is_err()
        {
            return Vec::new();
        }
    }
    match writer.finish() {
        Ok(cursor) => cursor.into_inner(),
        Err(_) => Vec::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sheet_names_are_sanitised() {
        assert_eq!(sanitize_sheet_names(&[]), vec!["Sheet1"]);
        assert_eq!(
            sanitize_sheet_names(&["a/b", "A_B", "", "'q'", &"x".repeat(40)]),
            vec![
                "a_b".to_string(),
                "A_B (2)".to_string(),
                "Sheet3".to_string(),
                "q".to_string(),
                "x".repeat(31)
            ]
        );
    }
}
