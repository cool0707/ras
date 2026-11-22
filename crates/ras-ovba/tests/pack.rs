//! `pack` / `unpack` / `workbook_code_names` / `create_workbook` on
//! workbook files.

mod common;

use std::fs;
use std::path::{Path, PathBuf};

use common::*;
use ras_ovba::{
    create_workbook, pack, read_project_bin, unpack, workbook_code_names, BookKind, DocumentCode,
    ModuleKind, OvbaError, PackReport, PackRequest, VbaModule,
};

const CT_VBA: &str = "application/vnd.ms-office.vbaProject";
const CT_XLSM: &str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const CT_XLSX: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_XLTM: &str = "application/vnd.ms-excel.template.macroEnabled.main+xml";
const CT_XLAM: &str = "application/vnd.ms-excel.addin.macroEnabled.main+xml";
const REL_VBA: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

const JP_SOURCE: &str = "Option Explicit\r\n\r\n' 売上を集計する\r\nPublic Function 合計(ByVal 範囲 As Range) As Double\r\n    合計 = WorksheetFunction.Sum(範囲) ' ｶﾅ も OK\r\nEnd Function\r\n";
const CLASS_SOURCE: &str = "Attribute VB_Name = \"Counter\"\r\nAttribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\nOption Explicit\r\nPrivate n As Long\r\nPublic Sub Inc()\r\n    n = n + 1\r\nEnd Sub\r\n";

fn doc_header(name: &str, class_id: &str) -> String {
    format!("Attribute VB_Name = \"{name}\"\r\nAttribute VB_Base = \"0{{{class_id}-0000-0000-C000-000000000046}}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n")
}

fn workbook_header() -> String {
    doc_header("ThisWorkbook", "00020819")
}

fn sheet_header(name: &str) -> String {
    doc_header(name, "00020820")
}

fn new_book(dir: &Path, file: &str, kind: BookKind, sheets: &[&str]) -> PathBuf {
    let path = dir.join(file);
    fs::write(&path, create_workbook(kind, sheets)).unwrap();
    path
}

fn request(book: &Path) -> PackRequest {
    PackRequest {
        book: book.to_path_buf(),
        ..PackRequest::default()
    }
}

fn base_request(book: &Path) -> PackRequest {
    PackRequest {
        modules: vec![standard("集計", JP_SOURCE), class("Counter", CLASS_SOURCE)],
        document_code: vec![
            DocumentCode {
                code_name: "ThisWorkbook".into(),
                code: "Private Sub Workbook_Open()\n    Call 起動\nEnd Sub\n".into(),
            },
            DocumentCode {
                code_name: "Sheet1".into(),
                code: "Private Sub Worksheet_Activate()\r\nEnd Sub".into(),
            },
        ],
        ..request(book)
    }
}

fn sources(book: &Path) -> Vec<VbaModule> {
    unpack(book)
        .unwrap()
        .project
        .expect("workbook has a project")
        .modules
}

fn source_of(book: &Path, name: &str) -> String {
    sources(book)
        .into_iter()
        .find(|m| m.name == name)
        .unwrap_or_else(|| panic!("module {name}"))
        .source
}

fn sorted(mut names: Vec<String>) -> Vec<String> {
    names.sort();
    names
}

fn names(list: &[&str]) -> Vec<String> {
    sorted(list.iter().map(|s| s.to_string()).collect())
}

fn assert_no_changes(report: &PackReport, all: &[&str]) {
    assert!(report.created.is_empty(), "{report:?}");
    assert!(report.updated.is_empty(), "{report:?}");
    assert!(report.removed.is_empty(), "{report:?}");
    assert_eq!(sorted(report.unchanged.clone()), names(all));
}

// ---------------------------------------------------------------------------

#[test]
fn created_workbook_is_a_valid_empty_package() {
    let dir = fresh_dir("create_workbook");
    let path = new_book(&dir, "empty.xlsm", BookKind::Xlsm, &["データ", "Sheet <2>"]);
    assert_eq!(
        zip_names(&path),
        vec![
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
            "xl/styles.xml",
        ]
    );
    assert!(zip_text(&path, "[Content_Types].xml").contains(CT_XLSM));
    assert!(zip_text(&path, "xl/workbook.xml")
        .contains("<sheet name=\"Sheet &lt;2&gt;\" sheetId=\"2\" r:id=\"rId2\"/>"));
    let code_names = workbook_code_names(&path).unwrap();
    assert_eq!(code_names.workbook, "ThisWorkbook");
    let listed: Vec<(&str, &str)> = code_names
        .sheets
        .iter()
        .map(|s| (s.sheet_name.as_str(), s.code_name.as_str()))
        .collect();
    assert_eq!(listed, vec![("データ", "Sheet1"), ("Sheet <2>", "Sheet2")]);

    let unpacked = unpack(&path).unwrap();
    assert!(unpacked.project.is_none());
    assert!(!unpacked.has_signature);
    assert_eq!(unpacked.code_names, code_names);

    let xlsx = new_book(&dir, "plain.xlsx", BookKind::Xlsx, &[]);
    assert!(zip_text(&xlsx, "[Content_Types].xml").contains(CT_XLSX));
    assert_eq!(workbook_code_names(&xlsx).unwrap().sheets.len(), 1);
    let xlam = new_book(&dir, "addin.xlam", BookKind::Xlam, &["A"]);
    assert!(zip_text(&xlam, "[Content_Types].xml").contains(CT_XLAM));
}

#[test]
fn pack_creates_a_project_and_unpack_returns_it() {
    let dir = fresh_dir("new_project");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["データ", "Sheet2"]);
    let before = fs::read(&book).unwrap();

    let report = pack(&base_request(&book)).unwrap();
    assert_eq!(report.output, book);
    assert_eq!(
        sorted(report.created.clone()),
        names(&["ThisWorkbook", "Sheet1", "Sheet2", "集計", "Counter"])
    );
    assert!(report.updated.is_empty() && report.removed.is_empty() && report.unchanged.is_empty());
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);

    // --- unpack equals the input -------------------------------------------
    let unpacked = unpack(&book).unwrap();
    assert!(!unpacked.has_signature);
    let project = unpacked.project.unwrap();
    assert_eq!(project.name, "VBAProject");
    assert_eq!(project.codepage, 932);
    assert!(!project.is_protected);
    let expected = vec![
        module(
            ModuleKind::Document,
            "ThisWorkbook",
            &format!(
                "{}Private Sub Workbook_Open()\r\n    Call 起動\r\nEnd Sub\r\n",
                workbook_header()
            ),
        ),
        module(
            ModuleKind::Document,
            "Sheet1",
            &format!(
                "{}Private Sub Worksheet_Activate()\r\nEnd Sub\r\n",
                sheet_header("Sheet1")
            ),
        ),
        module(ModuleKind::Document, "Sheet2", &sheet_header("Sheet2")),
        standard(
            "集計",
            &format!("Attribute VB_Name = \"集計\"\r\n{JP_SOURCE}"),
        ),
        class("Counter", CLASS_SOURCE),
    ];
    assert_eq!(project.modules, expected);

    // --- the independent `ovba` crate reads the same thing ------------------
    let bin = zip_entry(&book, "xl/vbaProject.bin").unwrap();
    let ours: Vec<(String, String)> = expected
        .iter()
        .map(|m| (m.name.clone(), m.source.clone()))
        .collect();
    assert_eq!(ovba_modules(&bin), ours);
    let project = ovba::open_project(bin.clone()).unwrap();
    assert_eq!(project.information.code_page, 932);
    assert!(project.modules.iter().all(|m| m.text_offset == 0));

    // --- package plumbing ------------------------------------------------------
    let content_types = zip_text(&book, "[Content_Types].xml");
    assert!(content_types.contains(&format!(
        "<Default Extension=\"bin\" ContentType=\"{CT_VBA}\"/>"
    )));
    assert!(content_types.contains(&format!(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT_XLSM}\"/>"
    )));
    let rels = zip_text(&book, "xl/_rels/workbook.xml.rels");
    assert!(rels.contains(&format!(
        "<Relationship Id=\"rId4\" Type=\"{REL_VBA}\" Target=\"vbaProject.bin\"/></Relationships>"
    )));
    assert_eq!(
        zip_names(&book).last().map(String::as_str),
        Some("xl/vbaProject.bin")
    );

    // --- everything else is copied raw, without recompression --------------
    let original = dir.join("original.zip");
    fs::write(&original, &before).unwrap();
    for name in [
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/styles.xml",
    ] {
        assert_eq!(zip_raw(&book, name), zip_raw(&original, name), "{name}");
    }
    // No temporary files are left behind.
    let leftovers: Vec<String> = fs::read_dir(&dir)
        .unwrap()
        .map(|e| e.unwrap().file_name().to_string_lossy().into_owned())
        .filter(|n| n.contains("ras-tmp"))
        .collect();
    assert!(leftovers.is_empty(), "{leftovers:?}");
}

#[test]
fn pack_is_idempotent() {
    let dir = fresh_dir("idempotent");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1", "Sheet2"]);
    let backups = dir.join("backup");
    let req = PackRequest {
        backup_dir: Some(backups.clone()),
        backup_generations: 5,
        ..base_request(&book)
    };
    pack(&req).unwrap();
    let after_first = fs::read(&book).unwrap();
    assert_eq!(fs::read_dir(&backups).unwrap().count(), 1); // the empty workbook

    let report = pack(&req).unwrap();
    assert_no_changes(
        &report,
        &["ThisWorkbook", "Sheet1", "Sheet2", "集計", "Counter"],
    );
    assert!(report.warnings.is_empty());
    // Nothing was rewritten and no further backup was taken.
    assert_eq!(fs::read(&book).unwrap(), after_first);
    assert_eq!(fs::read_dir(&backups).unwrap().count(), 1);

    // Line-ending style and a missing final newline do not count as changes.
    let mut relaxed = req.clone();
    relaxed.modules[0].source = JP_SOURCE.replace("\r\n", "\n").trim_end().to_string();
    let report = pack(&relaxed).unwrap();
    assert_no_changes(
        &report,
        &["ThisWorkbook", "Sheet1", "Sheet2", "集計", "Counter"],
    );
    assert_eq!(fs::read(&book).unwrap(), after_first);
}

#[test]
fn update_and_remove_are_classified() {
    let dir = fresh_dir("update_remove");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let mut first = base_request(&book);
    first
        .modules
        .push(standard("Old", "Sub Legacy()\r\nEnd Sub\r\n"));
    pack(&first).unwrap();

    let second = PackRequest {
        modules: vec![
            standard("集計", &JP_SOURCE.replace("Sum", "Average")),
            class("Counter", CLASS_SOURCE),
            standard("Extra", "Public Const X = 1\r\n"),
        ],
        remove_modules: vec!["Old".into(), "DoesNotExist".into()],
        ..request(&book)
    };
    let report = pack(&second).unwrap();
    assert_eq!(report.created, vec!["Extra".to_string()]);
    assert_eq!(report.updated, vec!["集計".to_string()]);
    assert_eq!(report.removed, vec!["Old".to_string()]);
    assert_eq!(
        sorted(report.unchanged.clone()),
        names(&["ThisWorkbook", "Sheet1", "Counter"])
    );

    let modules = sources(&book);
    let listed: Vec<&str> = modules.iter().map(|m| m.name.as_str()).collect();
    assert_eq!(
        listed,
        vec!["ThisWorkbook", "Sheet1", "集計", "Counter", "Extra"]
    );
    assert!(source_of(&book, "集計").contains("WorksheetFunction.Average(範囲)"));
    assert_eq!(source_of(&book, "Counter"), CLASS_SOURCE);
    // Document code that is not mentioned in a request stays as it is.
    assert!(source_of(&book, "ThisWorkbook").contains("Private Sub Workbook_Open()"));
    // The stream of the removed module is gone.
    let bin = zip_entry(&book, "xl/vbaProject.bin").unwrap();
    assert!(cfb_stream(&bin, "/VBA/Old").is_none());
    assert!(cfb_stream(&bin, "/VBA/Extra").is_some());
}

#[test]
fn document_code_replaces_the_whole_module() {
    let dir = fresh_dir("document_code");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    pack(&base_request(&book)).unwrap();
    // Code written in the VBA editor.
    let hand_written = format!(
        "{}Private Sub 手書き()\r\n    Debug.Print \"before\"\r\nEnd Sub\r\n",
        sheet_header("Sheet1")
    );
    let edit = PackRequest {
        modules: vec![module(ModuleKind::Document, "Sheet1", &hand_written)],
        ..request(&book)
    };
    let report = pack(&edit).unwrap();
    assert_eq!(report.updated, vec!["Sheet1".to_string()]);
    assert_eq!(source_of(&book, "Sheet1"), hand_written);

    // The generated code replaces all of it; the header of the module stays.
    let regen = PackRequest {
        document_code: vec![DocumentCode {
            code_name: "sheet1".into(), // code names are case-insensitive
            code:
                "Private Sub Worksheet_Change(ByVal Target As Range)\r\n    ' 変更\r\nEnd Sub\r\n"
                    .into(),
        }],
        ..request(&book)
    };
    let report = pack(&regen).unwrap();
    assert_eq!(report.updated, vec!["Sheet1".to_string()]);
    assert_eq!(
        source_of(&book, "Sheet1"),
        format!(
            "{}Private Sub Worksheet_Change(ByVal Target As Range)\r\n    ' 変更\r\nEnd Sub\r\n",
            sheet_header("Sheet1")
        )
    );
    // Same code again: unchanged.
    let report = pack(&regen).unwrap();
    assert!(report.updated.is_empty() && report.unchanged.contains(&"Sheet1".to_string()));

    // Empty code: a module without code.
    let clear = PackRequest {
        document_code: vec![DocumentCode {
            code_name: "Sheet1".into(),
            code: String::new(),
        }],
        ..request(&book)
    };
    pack(&clear).unwrap();
    assert_eq!(source_of(&book, "Sheet1"), sheet_header("Sheet1"));
}

#[test]
fn unknown_code_name_is_an_error() {
    let dir = fresh_dir("unknown_code_name");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let before = fs::read(&book).unwrap();
    let req = PackRequest {
        document_code: vec![DocumentCode {
            code_name: "Sheet9".into(),
            code: "' x".into(),
        }],
        ..base_request(&book)
    };
    match pack(&req) {
        Err(OvbaError::UnknownCodeName(name)) => assert_eq!(name, "Sheet9"),
        other => panic!("expected UnknownCodeName, got {other:?}"),
    }
    let req = PackRequest {
        modules: vec![module(ModuleKind::Document, "Chart7", "")],
        ..request(&book)
    };
    assert!(matches!(pack(&req), Err(OvbaError::UnknownCodeName(_))));
    // Nothing was written.
    assert_eq!(fs::read(&book).unwrap(), before);
    assert_eq!(fs::read_dir(&dir).unwrap().count(), 1);
}

#[test]
fn a_workbook_open_in_excel_is_not_touched() {
    let dir = fresh_dir("locked");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let before = fs::read(&book).unwrap();
    // Excel's owner file next to an open workbook.
    let owner = dir.join("~$book.xlsm");
    fs::write(&owner, b"owner").unwrap();
    assert!(matches!(pack(&base_request(&book)), Err(OvbaError::Locked)));
    assert_eq!(fs::read(&book).unwrap(), before);
    assert_eq!(fs::read_dir(&dir).unwrap().count(), 2);

    // The output being open is just as bad.
    let source = new_book(&dir, "source.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let req = PackRequest {
        output: Some(book.clone()),
        ..base_request(&source)
    };
    assert!(matches!(pack(&req), Err(OvbaError::Locked)));
    assert_eq!(fs::read(&book).unwrap(), before);

    fs::remove_file(&owner).unwrap();
    assert!(pack(&base_request(&book)).is_ok());
}

#[test]
fn backups_are_kept_and_pruned() {
    let dir = fresh_dir("backups");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let backups = dir.join("build").join("backup");
    let mut versions: Vec<Vec<u8>> = vec![fs::read(&book).unwrap()];
    for round in 0..5 {
        let req = PackRequest {
            modules: vec![standard("M", &format!("Public Const Round = {round}\r\n"))],
            backup_dir: Some(backups.clone()),
            backup_generations: 2,
            ..request(&book)
        };
        let report = pack(&req).unwrap();
        assert!(report.warnings.is_empty());
        versions.push(fs::read(&book).unwrap());
    }
    let mut files: Vec<PathBuf> = fs::read_dir(&backups)
        .unwrap()
        .map(|e| e.unwrap().path())
        .collect();
    files.sort_by_key(|p| {
        // `<stem>.<stamp>[-n].<ext>`: order by stamp, then by counter.
        let name = p.file_name().unwrap().to_string_lossy().into_owned();
        let middle = name
            .trim_start_matches("book.")
            .trim_end_matches(".xlsm")
            .to_string();
        let (stamp, counter) = match middle.get(15..) {
            Some(tail) if !tail.is_empty() => {
                (middle[..15].to_string(), tail[1..].parse::<u32>().unwrap())
            }
            _ => (middle.clone(), 0),
        };
        (stamp, counter)
    });
    assert_eq!(files.len(), 2, "{files:?}");
    for file in &files {
        let name = file.file_name().unwrap().to_string_lossy().into_owned();
        assert!(
            name.starts_with("book.20") && name.ends_with(".xlsm"),
            "{name}"
        );
    }
    // The two newest backups are the files that the last two packs replaced.
    assert_eq!(fs::read(&files[0]).unwrap(), versions[3]);
    assert_eq!(fs::read(&files[1]).unwrap(), versions[4]);
    assert_eq!(
        source_of(&book, "M"),
        "Attribute VB_Name = \"M\"\r\nPublic Const Round = 4\r\n"
    );

    // generations = 0 disables backups.
    let none = dir.join("none");
    let req = PackRequest {
        modules: vec![standard("M", "' final\r\n")],
        backup_dir: Some(none.clone()),
        backup_generations: 0,
        ..request(&book)
    };
    pack(&req).unwrap();
    assert!(!none.exists());
}

#[test]
fn signed_project_loses_its_signature_with_a_warning() {
    let dir = fresh_dir("signature");
    let book = dir.join("signed.xlsm");
    fs::copy(fixture("rust_xlsxwriter_macro04_signed.xlsm"), &book).unwrap();
    let original = fs::read(&book).unwrap();

    let unpacked = unpack(&book).unwrap();
    assert!(unpacked.has_signature);
    let project = unpacked.project.unwrap();
    let listed: Vec<(&str, ModuleKind)> = project
        .modules
        .iter()
        .map(|m| (m.name.as_str(), m.kind))
        .collect();
    assert_eq!(
        listed,
        vec![
            ("Module1", ModuleKind::Standard),
            ("ThisWorkbook", ModuleKind::Document),
            ("Sheet1", ModuleKind::Document)
        ]
    );
    let module1 = project.modules[0].source.clone();

    // Packing what is already there changes nothing and keeps the signature.
    let same = PackRequest {
        codepage: 1252,
        modules: vec![standard("Module1", &module1)],
        ..request(&book)
    };
    let report = pack(&same).unwrap();
    assert_no_changes(&report, &["Module1", "ThisWorkbook", "Sheet1"]);
    assert!(report.warnings.is_empty());
    assert_eq!(fs::read(&book).unwrap(), original);

    // A real change invalidates the signature: it is removed, with a warning.
    let change = PackRequest {
        codepage: 1252,
        modules: vec![standard(
            "Module1",
            &module1.replace("End Sub", "    Beep\r\nEnd Sub"),
        )],
        ..request(&book)
    };
    let report = pack(&change).unwrap();
    assert_eq!(report.updated, vec!["Module1".to_string()]);
    assert_eq!(report.warnings.len(), 1, "{:?}", report.warnings);
    assert!(
        report.warnings[0].contains("signature")
            && report.warnings[0].contains("xl/vbaProjectSignature.bin")
    );

    let entries = zip_names(&book);
    assert!(
        !entries
            .iter()
            .any(|n| n.to_lowercase().contains("signature")),
        "{entries:?}"
    );
    assert!(!entries.contains(&"xl/_rels/vbaProject.bin.rels".to_string()));
    let content_types = zip_text(&book, "[Content_Types].xml");
    assert!(!content_types.contains("vbaProjectSignature"));
    assert!(content_types.contains(CT_VBA));
    // Same number of entries minus the signature part and its rels part,
    // in the same order.
    let before_names = {
        let copy = dir.join("original.zip");
        fs::write(&copy, &original).unwrap();
        zip_names(&copy)
    };
    let expected: Vec<String> = before_names
        .into_iter()
        .filter(|n| n != "xl/vbaProjectSignature.bin" && n != "xl/_rels/vbaProject.bin.rels")
        .collect();
    assert_eq!(entries, expected);

    let after = unpack(&book).unwrap();
    assert!(!after.has_signature);
    assert!(after.project.unwrap().modules[0]
        .source
        .contains("    Beep\r\nEnd Sub"));

    // The rewritten project no longer carries Excel's p-code or caches.
    let bin = zip_entry(&book, "xl/vbaProject.bin").unwrap();
    assert_eq!(cfb_stream(&bin, "/VBA/_VBA_PROJECT").unwrap().len(), 7);
    assert!(cfb_streams(&bin).iter().all(|(p, _)| !p.contains("__SRP_")));
    assert!(ovba::open_project(bin)
        .unwrap()
        .modules
        .iter()
        .all(|m| m.text_offset == 0));
}

#[test]
fn xlsx_needs_convert_and_is_written_to_a_sibling_xlsm() {
    let dir = fresh_dir("convert");
    let book = new_book(&dir, "plain.xlsx", BookKind::Xlsx, &["Sheet1"]);
    let before = fs::read(&book).unwrap();

    assert!(matches!(
        pack(&base_request(&book)),
        Err(OvbaError::UnsupportedFormat(_))
    ));
    assert_eq!(fs::read_dir(&dir).unwrap().count(), 1);

    let req = PackRequest {
        convert: true,
        ..base_request(&book)
    };
    let report = pack(&req).unwrap();
    let output = dir.join("plain.xlsm");
    assert_eq!(report.output, output);
    assert_eq!(
        fs::read(&book).unwrap(),
        before,
        "the original .xlsx is untouched"
    );
    let content_types = zip_text(&output, "[Content_Types].xml");
    assert!(content_types.contains(&format!(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT_XLSM}\"/>"
    )));
    assert!(!content_types.contains(CT_XLSX));
    assert!(content_types.contains(CT_VBA));
    assert_eq!(sources(&output).len(), 4);

    // An .xlsx output can never hold macros.
    let bad = PackRequest {
        convert: true,
        output: Some(dir.join("out.xlsx")),
        ..base_request(&book)
    };
    assert!(matches!(pack(&bad), Err(OvbaError::UnsupportedFormat(_))));
    // Converting again onto the existing .xlsm: the modules are unchanged
    // relative to the *input*, which has none, so they count as created.
    let again = pack(&req).unwrap();
    assert_eq!(again.created.len(), 4);
}

#[test]
fn unsupported_inputs() {
    let dir = fresh_dir("unsupported");
    let xls = dir.join("legacy.xls");
    fs::write(&xls, b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1 not really").unwrap();
    assert!(matches!(
        pack(&request(&xls)),
        Err(OvbaError::UnsupportedFormat(_))
    ));
    assert!(matches!(unpack(&xls), Err(OvbaError::UnsupportedFormat(_))));
    assert!(matches!(
        unpack(&dir.join("notes.txt")),
        Err(OvbaError::UnsupportedFormat(_))
    ));

    // A password-encrypted workbook is a compound file with an .xlsm name.
    let encrypted = dir.join("encrypted.xlsm");
    fs::write(&encrypted, b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1\0\0\0\0").unwrap();
    assert!(matches!(
        unpack(&encrypted),
        Err(OvbaError::UnsupportedFormat(_))
    ));
    assert!(matches!(
        pack(&request(&encrypted)),
        Err(OvbaError::UnsupportedFormat(_))
    ));
    assert!(matches!(
        workbook_code_names(&encrypted),
        Err(OvbaError::UnsupportedFormat(_))
    ));

    // Not a zip at all / missing file.
    let junk = dir.join("junk.xlsm");
    fs::write(&junk, b"this is not a workbook").unwrap();
    assert!(matches!(
        unpack(&junk),
        Err(OvbaError::Zip(_) | OvbaError::Io(_))
    ));
    assert!(matches!(
        unpack(&dir.join("missing.xlsm")),
        Err(OvbaError::Io(_))
    ));
    assert!(matches!(
        pack(&request(&dir.join("missing.xlsm"))),
        Err(OvbaError::Io(_))
    ));

    // New user forms and invalid code pages.
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let form = PackRequest {
        modules: vec![module(ModuleKind::Form, "UserForm1", "")],
        ..request(&book)
    };
    assert!(matches!(pack(&form), Err(OvbaError::UnsupportedFormat(_))));
    let utf8 = PackRequest {
        codepage: 65001,
        ..request(&book)
    };
    assert!(matches!(pack(&utf8), Err(OvbaError::Encoding(_))));
    let emoji = PackRequest {
        modules: vec![standard("M", "' 🎉")],
        ..request(&book)
    };
    assert!(matches!(pack(&emoji), Err(OvbaError::Encoding(_))));
    assert_eq!(
        fs::read(&book).unwrap(),
        create_workbook(BookKind::Xlsm, &["Sheet1"])
    );
}

#[test]
fn output_path_and_content_type_follow_the_extension() {
    let dir = fresh_dir("output_path");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    let before = fs::read(&book).unwrap();

    let template = dir.join("out").join("template.xltm");
    fs::create_dir_all(template.parent().unwrap()).unwrap();
    let report = pack(&PackRequest {
        output: Some(template.clone()),
        ..base_request(&book)
    })
    .unwrap();
    assert_eq!(report.output, template);
    assert_eq!(
        fs::read(&book).unwrap(),
        before,
        "the input is not modified"
    );
    assert!(zip_text(&template, "[Content_Types].xml").contains(CT_XLTM));
    assert_eq!(sources(&template).len(), 4);

    let addin = dir.join("addin.xlam");
    pack(&PackRequest {
        output: Some(addin.clone()),
        ..base_request(&book)
    })
    .unwrap();
    assert!(zip_text(&addin, "[Content_Types].xml").contains(CT_XLAM));

    // An add-in created as such keeps its type.
    let xlam = new_book(&dir, "native.xlam", BookKind::Xlam, &["Sheet1"]);
    pack(&base_request(&xlam)).unwrap();
    let content_types = zip_text(&xlam, "[Content_Types].xml");
    assert!(content_types.contains(CT_XLAM) && content_types.contains(CT_VBA));

    // XML <-> binary conversion is not offered.
    let binary = PackRequest {
        output: Some(dir.join("book.xlsb")),
        ..base_request(&book)
    };
    assert!(matches!(
        pack(&binary),
        Err(OvbaError::UnsupportedFormat(_))
    ));
}

/// A workbook as third-party libraries write it: no code names anywhere,
/// a `bin` default that is not the VBA project, no workbookPr element.
fn book_without_code_names() -> Vec<u8> {
    let types = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/chartsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/></Types>";
    let root_rels = "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\"/></Relationships>";
    let workbook = "<?xml version=\"1.0\"?>\n<x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:bookViews><x:workbookView/></x:bookViews><x:sheets><x:sheet name=\"First\" sheetId=\"1\" rel:id=\"rId7\"/><x:sheet name=\"Sheet1\" sheetId=\"2\" rel:id=\"rId8\"/><x:sheet name=\"Graph\" sheetId=\"3\" rel:id=\"rId9\"/></x:sheets></x:workbook>";
    let workbook_rels = "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet2.xml\"/><Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet1.xml\"/></Relationships>";
    // sheet1: no sheetPr; sheet2: sheetPr without codeName, already "Sheet1"-like tab name.
    let sheet1 = "<?xml version=\"1.0\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:dimension ref=\"A1\"/><x:sheetData/></x:worksheet>";
    let sheet2 = "<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetPr filterMode=\"1\"><tabColor rgb=\"FFFF0000\"/></sheetPr><sheetData/></worksheet>";
    let chart = "<?xml version=\"1.0\"?><chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetPr/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews></chartsheet>";
    build_zip(&[
        ("[Content_Types].xml", types.as_bytes().to_vec()),
        ("_rels/.rels", root_rels.as_bytes().to_vec()),
        ("xl/workbook.xml", workbook.as_bytes().to_vec()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_rels.as_bytes().to_vec(),
        ),
        ("xl/worksheets/sheet1.xml", sheet1.as_bytes().to_vec()),
        ("xl/worksheets/sheet2.xml", sheet2.as_bytes().to_vec()),
        ("xl/chartsheets/sheet1.xml", chart.as_bytes().to_vec()),
        ("xl/printerSettings/printerSettings1.bin", vec![1, 2, 3, 4]),
    ])
}

#[test]
fn missing_code_names_are_assigned_and_written() {
    let dir = fresh_dir("code_names");
    let book = dir.join("thirdparty.xlsx");
    fs::write(&book, book_without_code_names()).unwrap();

    let code_names = workbook_code_names(&book).unwrap();
    assert_eq!(code_names.workbook, "ThisWorkbook");
    let listed: Vec<(&str, &str)> = code_names
        .sheets
        .iter()
        .map(|s| (s.sheet_name.as_str(), s.code_name.as_str()))
        .collect();
    assert_eq!(
        listed,
        vec![
            ("First", "Sheet1"),
            ("Sheet1", "Sheet2"),
            ("Graph", "Chart1")
        ]
    );

    let req = PackRequest {
        convert: true,
        codepage: 1252,
        modules: vec![standard("Main", "Sub Main()\r\nEnd Sub\r\n")],
        document_code: vec![DocumentCode {
            code_name: "Sheet2".into(),
            code: "' on sheet 2".into(),
        }],
        ..request(&book)
    };
    let report = pack(&req).unwrap();
    let output = dir.join("thirdparty.xlsm");
    assert_eq!(report.output, output);
    assert_eq!(
        sorted(report.created.clone()),
        names(&["ThisWorkbook", "Sheet1", "Sheet2", "Chart1", "Main"])
    );
    assert_eq!(report.warnings.len(), 1);
    assert!(
        report.warnings[0].contains("`First` -> Sheet1")
            && report.warnings[0].contains("`Graph` -> Chart1")
    );

    // The code names are now recorded in the parts, everything else is intact.
    assert_eq!(
        zip_text(&output, "xl/workbook.xml"),
        "<?xml version=\"1.0\"?>\n<x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:workbookPr codeName=\"ThisWorkbook\"/><x:bookViews><x:workbookView/></x:bookViews><x:sheets><x:sheet name=\"First\" sheetId=\"1\" rel:id=\"rId7\"/><x:sheet name=\"Sheet1\" sheetId=\"2\" rel:id=\"rId8\"/><x:sheet name=\"Graph\" sheetId=\"3\" rel:id=\"rId9\"/></x:sheets></x:workbook>"
    );
    assert_eq!(
        zip_text(&output, "xl/worksheets/sheet1.xml"),
        "<?xml version=\"1.0\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetPr codeName=\"Sheet1\"/><x:dimension ref=\"A1\"/><x:sheetData/></x:worksheet>"
    );
    assert!(zip_text(&output, "xl/worksheets/sheet2.xml").contains(
        "<sheetPr codeName=\"Sheet2\" filterMode=\"1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"
    ));
    assert!(zip_text(&output, "xl/chartsheets/sheet1.xml").contains("<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetPr codeName=\"Chart1\"/>"));

    // `bin` already meant printer settings: the project gets an Override.
    let content_types = zip_text(&output, "[Content_Types].xml");
    assert!(content_types.contains("<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>"));
    assert!(content_types.ends_with(&format!(
        "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"{CT_VBA}\"/></Types>"
    )));
    assert!(content_types.contains(&format!(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"{CT_XLSM}\"/>"
    )));
    assert!(zip_text(&output, "xl/_rels/workbook.xml.rels").contains(&format!("<Relationship Id=\"rId4\" Type=\"{REL_VBA}\" Target=\"vbaProject.bin\"/></Relationships>")));

    // Document modules: worksheet and chart sheet classes.
    let modules = sources(&output);
    let listed: Vec<&str> = modules.iter().map(|m| m.name.as_str()).collect();
    assert_eq!(
        listed,
        vec!["ThisWorkbook", "Sheet1", "Sheet2", "Chart1", "Main"]
    );
    assert_eq!(modules[1].source, sheet_header("Sheet1"));
    assert_eq!(
        modules[2].source,
        format!("{}' on sheet 2\r\n", sheet_header("Sheet2"))
    );
    assert_eq!(modules[3].source, doc_header("Chart1", "00020821"));
    assert!(modules[..4].iter().all(|m| m.kind == ModuleKind::Document));

    // After the first pack the names are explicit and a second pack is a no-op.
    let again = PackRequest {
        book: output.clone(),
        convert: false,
        ..req
    };
    let report = pack(&again).unwrap();
    assert_no_changes(
        &report,
        &["ThisWorkbook", "Sheet1", "Sheet2", "Chart1", "Main"],
    );
    assert!(report.warnings.is_empty());
    let listed: Vec<String> = workbook_code_names(&output)
        .unwrap()
        .sheets
        .into_iter()
        .map(|s| s.code_name)
        .collect();
    assert_eq!(listed, vec!["Sheet1", "Sheet2", "Chart1"]);
}

#[test]
fn excel_made_project_inside_a_workbook_is_updated_in_place() {
    // Put a genuine Excel project (with p-code, without matching sheets for
    // two of its document modules) into a workbook.
    let dir = fresh_dir("excel_project");
    let book = dir.join("book.xlsm");
    let empty = create_workbook(BookKind::Xlsm, &["Sheet1", "Sheet2"]);
    let with_project = patch_zip(
        &empty,
        &[
            ("xl/vbaProject.bin", Some(fixture_bytes("xlsxwriter_vbaProject.bin"))),
            (
                "xl/_rels/workbook.xml.rels",
                Some(
                    String::from_utf8(zip_entry_bytes(&empty, "xl/_rels/workbook.xml.rels"))
                        .unwrap()
                        .replace("</Relationships>", &format!("<Relationship Id=\"rId9\" Type=\"{REL_VBA}\" Target=\"vbaProject.bin\"/></Relationships>"))
                        .into_bytes(),
                ),
            ),
        ],
    );
    fs::write(&book, with_project).unwrap();

    let before = sources(&book);
    assert_eq!(before.len(), 5);
    let req = PackRequest {
        codepage: 1252,
        modules: vec![standard(
            "Module1",
            "Sub say_hello()\r\n    MsgBox \"Hello from RAS!\"\r\nEnd Sub\r\n",
        )],
        document_code: vec![DocumentCode {
            code_name: "Sheet2".into(),
            code: "' generated".into(),
        }],
        ..request(&book)
    };
    let report = pack(&req).unwrap();
    assert_eq!(
        sorted(report.updated.clone()),
        names(&["Module1", "Sheet2"])
    );
    assert_eq!(
        sorted(report.unchanged.clone()),
        names(&["ThisWorkbook", "Sheet1", "ThisWorkbook1"])
    );
    assert!(report.created.is_empty() && report.removed.is_empty());

    let after = sources(&book);
    assert_eq!(after[0], before[0]);
    assert_eq!(after[1], before[1]);
    assert_eq!(after[3], before[3]);
    assert_eq!(after[2].source, "Attribute VB_Name = \"Module1\"\r\nSub say_hello()\r\n    MsgBox \"Hello from RAS!\"\r\nEnd Sub\r\n");
    assert_eq!(
        after[4].source,
        format!("{}' generated\r\n", before[4].source)
    );
    // The content type for the part was missing in this hand-made package.
    assert!(zip_text(&book, "[Content_Types].xml").contains(CT_VBA));
    assert_no_changes(
        &pack(&req).unwrap(),
        &[
            "ThisWorkbook",
            "Sheet1",
            "Module1",
            "ThisWorkbook1",
            "Sheet2",
        ],
    );
}

fn zip_entry_bytes(zip: &[u8], name: &str) -> Vec<u8> {
    use std::io::Read;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(zip.to_vec())).unwrap();
    let mut data = Vec::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_end(&mut data)
        .unwrap();
    data
}

// ---------------------------------------------------------------------------
// .xlsb
// ---------------------------------------------------------------------------

fn biff12_record(id: u32, body: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    if id < 0x80 {
        out.push(id as u8);
    } else {
        out.push((id & 0x7F) as u8 | 0x80);
        out.push((id >> 7) as u8);
    }
    let mut len = body.len();
    loop {
        let byte = (len & 0x7F) as u8;
        len >>= 7;
        if len == 0 {
            out.push(byte);
            break;
        }
        out.push(byte | 0x80);
    }
    out.extend_from_slice(body);
    out
}

fn wide(s: &str) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut out = (units.len() as u32).to_le_bytes().to_vec();
    out.extend(units.iter().flat_map(|u| u.to_le_bytes()));
    out
}

/// A skeletal `.xlsb` package: only the records this crate looks at.
fn xlsb(workbook_code_name: &str, sheets: &[(&str, &str)], project: Option<Vec<u8>>) -> Vec<u8> {
    let mut workbook = biff12_record(131, &[]); // BrtBeginBook
    let mut wb_prop = vec![0x20, 0, 1, 0, 0, 0, 0, 0];
    wb_prop.extend(wide(workbook_code_name));
    workbook.extend(biff12_record(153, &wb_prop)); // BrtWbProp
    workbook.extend(biff12_record(143, &[])); // BrtBeginBundleShs
    let mut rels = String::from("<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for (index, (name, code_name)) in sheets.iter().enumerate() {
        let n = index + 1;
        let mut bundle = vec![0, 0, 0, 0];
        bundle.extend((n as u32).to_le_bytes());
        bundle.extend(wide(&format!("rId{n}")));
        bundle.extend(wide(name));
        workbook.extend(biff12_record(156, &bundle)); // BrtBundleSh
        rels.push_str(&format!("<Relationship Id=\"rId{n}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{n}.bin\"/>"));
        let mut sheet = biff12_record(129, &[]); // BrtBeginSheet
        let mut ws_prop = vec![
            0xC9, 0x04, 0x02, 0x00, 0x40, 0, 0, 0, 0, 0, 0, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF,
        ];
        ws_prop.extend(wide(code_name));
        sheet.extend(biff12_record(147, &ws_prop)); // BrtWsProp
        sheet.extend(biff12_record(145, &[])); // BrtBeginSheetData
        sheet.extend(biff12_record(146, &[])); // BrtEndSheetData
        sheet.extend(biff12_record(130, &[])); // BrtEndSheet
        entries.push((format!("xl/worksheets/sheet{n}.bin"), sheet));
    }
    workbook.extend(biff12_record(144, &[])); // BrtEndBundleShs
    workbook.extend(biff12_record(132, &[])); // BrtEndBook
    if project.is_some() {
        rels.push_str(&format!(
            "<Relationship Id=\"rId99\" Type=\"{REL_VBA}\" Target=\"vbaProject.bin\"/>"
        ));
    }
    rels.push_str("</Relationships>");
    let types = "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/worksheets/sheet1.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/></Types>";
    let root_rels = "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/></Relationships>";
    let mut all: Vec<(&str, Vec<u8>)> = vec![
        ("[Content_Types].xml", types.as_bytes().to_vec()),
        ("_rels/.rels", root_rels.as_bytes().to_vec()),
        ("xl/workbook.bin", workbook),
        ("xl/_rels/workbook.bin.rels", rels.into_bytes()),
    ];
    for (name, data) in &entries {
        all.push((name.as_str(), data.clone()));
    }
    if let Some(project) = project {
        all.push(("xl/vbaProject.bin", project));
    }
    build_zip(&all)
}

#[test]
fn xlsb_with_code_names() {
    let dir = fresh_dir("xlsb");
    let book = dir.join("book.xlsb");
    fs::write(
        &book,
        xlsb(
            "ThisWorkbook",
            &[("データ", "Sheet1"), ("Other", "Sheet2")],
            None,
        ),
    )
    .unwrap();

    let code_names = workbook_code_names(&book).unwrap();
    assert_eq!(code_names.workbook, "ThisWorkbook");
    let listed: Vec<(&str, &str)> = code_names
        .sheets
        .iter()
        .map(|s| (s.sheet_name.as_str(), s.code_name.as_str()))
        .collect();
    assert_eq!(listed, vec![("データ", "Sheet1"), ("Other", "Sheet2")]);
    assert!(unpack(&book).unwrap().project.is_none());

    let workbook_before = zip_raw(&book, "xl/workbook.bin");
    let sheet_before = zip_raw(&book, "xl/worksheets/sheet1.bin");
    let report = pack(&base_request(&book)).unwrap();
    assert_eq!(report.created.len(), 5);
    // Binary parts are never rewritten.
    assert_eq!(zip_raw(&book, "xl/workbook.bin"), workbook_before);
    assert_eq!(zip_raw(&book, "xl/worksheets/sheet1.bin"), sheet_before);
    let content_types = zip_text(&book, "[Content_Types].xml");
    assert!(content_types.contains(&format!(
        "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"{CT_VBA}\"/>"
    )));
    assert!(content_types.contains("application/vnd.ms-excel.sheet.binary.macroEnabled.main"));
    assert!(zip_text(&book, "xl/_rels/workbook.bin.rels").contains(REL_VBA));

    let modules = sources(&book);
    let listed: Vec<&str> = modules.iter().map(|m| m.name.as_str()).collect();
    assert_eq!(
        listed,
        vec!["ThisWorkbook", "Sheet1", "Sheet2", "集計", "Counter"]
    );
    assert!(modules[1]
        .source
        .contains("\r\nPrivate Sub Worksheet_Activate()"));
    assert_no_changes(
        &pack(&base_request(&book)).unwrap(),
        &["ThisWorkbook", "Sheet1", "Sheet2", "集計", "Counter"],
    );
}

#[test]
fn xlsb_without_code_names() {
    let dir = fresh_dir("xlsb_no_names");
    // No project and no recorded code names: a project cannot be created.
    let bare = dir.join("bare.xlsb");
    fs::write(&bare, xlsb("", &[("Tabelle1", "")], None)).unwrap();
    assert_eq!(
        workbook_code_names(&bare).unwrap().sheets[0].code_name,
        "Sheet1"
    );
    let req = PackRequest {
        modules: vec![standard("M", "' m")],
        ..request(&bare)
    };
    assert!(matches!(pack(&req), Err(OvbaError::UnsupportedFormat(_))));

    // With an existing project standard and class modules can be written;
    // document code needs the module to exist already.
    let project = {
        let book = new_book(&dir, "donor.xlsm", BookKind::Xlsm, &["Sheet1"]);
        pack(&request(&book)).unwrap();
        zip_entry(&book, "xl/vbaProject.bin").unwrap()
    };
    let book = dir.join("book.xlsb");
    fs::write(
        &book,
        xlsb("", &[("Tabelle1", ""), ("Tabelle2", "")], Some(project)),
    )
    .unwrap();
    let report = pack(&req_with(&book, vec![standard("M", "' m")], vec![])).unwrap();
    assert_eq!(report.created, vec!["M".to_string()]);
    assert_eq!(
        sorted(report.unchanged.clone()),
        names(&["ThisWorkbook", "Sheet1"])
    );
    // Sheet1 exists as a module: its region can be managed.
    let ok = req_with(
        &book,
        vec![],
        vec![DocumentCode {
            code_name: "Sheet1".into(),
            code: "' x".into(),
        }],
    );
    assert_eq!(pack(&ok).unwrap().updated, vec!["Sheet1".to_string()]);
    // Sheet2 would have to be created, which needs a recorded code name.
    let missing = req_with(
        &book,
        vec![],
        vec![DocumentCode {
            code_name: "Sheet2".into(),
            code: "' x".into(),
        }],
    );
    assert!(matches!(
        pack(&missing),
        Err(OvbaError::UnsupportedFormat(_))
    ));
    let unknown = req_with(
        &book,
        vec![],
        vec![DocumentCode {
            code_name: "Sheet3".into(),
            code: "' x".into(),
        }],
    );
    assert!(matches!(pack(&unknown), Err(OvbaError::UnknownCodeName(_))));
}

fn req_with(book: &Path, modules: Vec<VbaModule>, document_code: Vec<DocumentCode>) -> PackRequest {
    PackRequest {
        modules,
        document_code,
        ..request(book)
    }
}

#[test]
fn project_name_and_code_page_changes_are_applied() {
    let dir = fresh_dir("project_name");
    let book = new_book(&dir, "book.xlsm", BookKind::Xlsm, &["Sheet1"]);
    pack(&PackRequest {
        codepage: 1252,
        modules: vec![standard("M", "' ascii only\r\n")],
        ..request(&book)
    })
    .unwrap();
    let project = unpack(&book).unwrap().project.unwrap();
    assert_eq!(
        (project.name.as_str(), project.codepage),
        ("VBAProject", 1252)
    );

    // Same modules, new project name and code page: nothing "updated", but
    // the file is rewritten.
    let req = PackRequest {
        codepage: 932,
        project_name: Some("売上管理".into()),
        modules: vec![standard("M", "' ascii only\r\n")],
        ..request(&book)
    };
    let report = pack(&req).unwrap();
    assert_no_changes(&report, &["ThisWorkbook", "Sheet1", "M"]);
    let project = unpack(&book).unwrap().project.unwrap();
    assert_eq!((project.name.as_str(), project.codepage), ("売上管理", 932));
    let bin = zip_entry(&book, "xl/vbaProject.bin").unwrap();
    assert_eq!(read_project_bin(&bin).unwrap().name, "売上管理");
    let project_stream = cfb_stream(&bin, "/PROJECT").unwrap();
    let (text, _, _) = encoding_rs::SHIFT_JIS.decode(&project_stream);
    assert!(text.contains("Name=\"売上管理\"\r\n"));
}
