//! Adding, copying, removing and renaming sheets in a workbook file (`edit_sheets`).

use std::fs;
use std::path::{Path, PathBuf};

use ras_ovba::{
    create_workbook, edit_sheets, pack, sheet_references, unpack, workbook_sheets, BookKind,
    DocumentCode, ModuleKind, OvbaError, PackRequest, Place, SheetEdit, SheetEditRequest,
    SheetError, VbaModule,
};

fn fresh_dir(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("ras-ovba-sheets-{}-{name}", std::process::id()));
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).unwrap();
    dir
}

/// A macro-enabled workbook with the sheets `Sheet1` and `Data` (code name `Sheet2`), a
/// standard module and code in the module of `Data`.
fn book(dir: &Path) -> PathBuf {
    let book = dir.join("book.xlsm");
    fs::write(&book, create_workbook(BookKind::Xlsm, &["Sheet1", "Data"])).unwrap();
    pack(&PackRequest {
        book: book.clone(),
        modules: vec![VbaModule {
            name: "Main".into(),
            kind: ModuleKind::Standard,
            source: "Sub Hello()\r\nEnd Sub\r\n".into(),
        }],
        document_code: vec![DocumentCode {
            code_name: "Sheet2".into(),
            code: "Private Sub Worksheet_Activate()\r\nEnd Sub\r\n".into(),
        }],
        ..PackRequest::default()
    })
    .unwrap();
    book
}

fn edit(book: &Path, edit: SheetEdit) -> Result<ras_ovba::SheetEditReport, OvbaError> {
    edit_sheets(&SheetEditRequest {
        book: book.to_path_buf(),
        edit,
        backup_dir: None,
        backup_generations: 0,
    })
}

fn sheets(book: &Path) -> Vec<(String, String)> {
    workbook_sheets(book)
        .unwrap()
        .into_iter()
        .map(|s| (s.name, s.code_name))
        .collect()
}

fn modules(book: &Path) -> Vec<(String, ModuleKind)> {
    unpack(book)
        .unwrap()
        .project
        .unwrap()
        .modules
        .into_iter()
        .map(|m| (m.name, m.kind))
        .collect()
}

fn pair(a: &str, b: &str) -> (String, String) {
    (a.to_string(), b.to_string())
}

#[test]
fn a_sheet_is_added_with_its_module() {
    let dir = fresh_dir("add");
    let book = book(&dir);
    let report = edit(
        &book,
        SheetEdit::Add {
            name: "集計".into(),
            code_name: None,
            place: Place::After("Sheet1".into()),
        },
    )
    .unwrap();
    assert_eq!(
        (report.name.as_str(), report.code_name.as_str()),
        ("集計", "Sheet3")
    );
    assert_eq!(
        sheets(&book),
        [
            pair("Sheet1", "Sheet1"),
            pair("集計", "Sheet3"),
            pair("Data", "Sheet2")
        ]
    );
    // The project has a document module for it; the others are as they were.
    let list = modules(&book);
    assert!(
        list.contains(&("Sheet3".to_string(), ModuleKind::Document))
            && list.contains(&("Main".to_string(), ModuleKind::Standard)),
        "{list:?}"
    );
    let project = unpack(&book).unwrap().project.unwrap();
    assert!(project
        .modules
        .iter()
        .find(|m| m.name == "Sheet2")
        .unwrap()
        .source
        .contains("Worksheet_Activate"));
    // At the end, with a code name of its own; a workbook without a project gets none.
    let report = edit(
        &book,
        SheetEdit::Add {
            name: "Notes".into(),
            code_name: Some("shtNotes".into()),
            place: Place::End,
        },
    )
    .unwrap();
    assert_eq!(report.code_name, "shtNotes");
    assert_eq!(sheets(&book).last().unwrap(), &pair("Notes", "shtNotes"));
    let plain = dir.join("plain.xlsx");
    fs::write(&plain, create_workbook(BookKind::Xlsx, &["Sheet1"])).unwrap();
    edit(
        &plain,
        SheetEdit::Add {
            name: "B".into(),
            code_name: None,
            place: Place::Before("Sheet1".into()),
        },
    )
    .unwrap();
    assert_eq!(
        sheets(&plain),
        [pair("B", "Sheet2"), pair("Sheet1", "Sheet1")]
    );
    assert!(unpack(&plain).unwrap().project.is_none());
}

#[test]
fn what_cannot_be_done_changes_nothing() {
    let dir = fresh_dir("refused");
    let book = book(&dir);
    let before = fs::read(&book).unwrap();
    let refused = |e: SheetEdit| match edit(&book, e) {
        Err(OvbaError::Sheet(e)) => e,
        other => panic!("expected a refusal, got {other:?}"),
    };
    assert_eq!(
        refused(SheetEdit::Add {
            name: "data".into(),
            code_name: None,
            place: Place::End
        }),
        SheetError::NameTaken("data".into())
    );
    assert_eq!(
        refused(SheetEdit::Add {
            name: "a/b".into(),
            code_name: None,
            place: Place::End
        }),
        SheetError::BadName("a/b".into())
    );
    assert_eq!(
        refused(SheetEdit::Add {
            name: "New".into(),
            code_name: Some("main".into()),
            place: Place::End
        }),
        SheetError::CodeNameTaken("main".into())
    );
    assert_eq!(
        refused(SheetEdit::Add {
            name: "New".into(),
            code_name: Some("1st".into()),
            place: Place::End
        }),
        SheetError::BadCodeName("1st".into())
    );
    assert_eq!(
        refused(SheetEdit::Add {
            name: "New".into(),
            code_name: None,
            place: Place::After("Nope".into())
        }),
        SheetError::NotFound("Nope".into())
    );
    assert_eq!(
        refused(SheetEdit::Remove {
            sheet: "Nope".into()
        }),
        SheetError::NotFound("Nope".into())
    );
    assert_eq!(
        refused(SheetEdit::Rename {
            sheet: "Sheet1".into(),
            name: Some("Data".into()),
            code_name: None
        }),
        SheetError::NameTaken("Data".into())
    );
    assert_eq!(
        refused(SheetEdit::Rename {
            sheet: "Sheet1".into(),
            name: None,
            code_name: Some("ThisWorkbook".into())
        }),
        SheetError::CodeNameTaken("ThisWorkbook".into())
    );
    assert_eq!(fs::read(&book).unwrap(), before);
    // The last visible sheet stays.
    edit(
        &book,
        SheetEdit::Remove {
            sheet: "Data".into(),
        },
    )
    .unwrap();
    assert_eq!(
        refused(SheetEdit::Remove {
            sheet: "Sheet1".into()
        }),
        SheetError::LastVisibleSheet
    );
    // A workbook open in Excel is not written to.
    fs::write(dir.join("~$book.xlsm"), b"").unwrap();
    assert!(matches!(
        edit(
            &book,
            SheetEdit::Add {
                name: "X".into(),
                code_name: None,
                place: Place::End
            }
        ),
        Err(OvbaError::Locked)
    ));
}

#[test]
fn a_sheet_is_removed_with_its_module() {
    let dir = fresh_dir("remove");
    let book = book(&dir);
    edit(
        &book,
        SheetEdit::Add {
            name: "Third".into(),
            code_name: None,
            place: Place::End,
        },
    )
    .unwrap();
    let report = edit(
        &book,
        SheetEdit::Remove {
            sheet: "Sheet2".into(),
        },
    )
    .unwrap();
    assert_eq!(
        (report.name.as_str(), report.code_name.as_str()),
        ("Data", "Sheet2")
    );
    assert_eq!(
        sheets(&book),
        [pair("Sheet1", "Sheet1"), pair("Third", "Sheet3")]
    );
    let list = modules(&book);
    assert!(
        !list.iter().any(|(name, _)| name == "Sheet2")
            && list.contains(&("Sheet3".to_string(), ModuleKind::Document))
            && list.contains(&("Main".to_string(), ModuleKind::Standard)),
        "{list:?}"
    );
    // A backup of the file before the change is kept.
    let backups = dir.join("backup");
    edit_sheets(&SheetEditRequest {
        book: book.clone(),
        edit: SheetEdit::Remove {
            sheet: "Third".into(),
        },
        backup_dir: Some(backups.clone()),
        backup_generations: 3,
    })
    .unwrap();
    assert_eq!(fs::read_dir(&backups).unwrap().count(), 1);
    assert_eq!(sheets(&book), [pair("Sheet1", "Sheet1")]);
}

#[test]
fn a_sheet_is_renamed_and_its_module_follows() {
    let dir = fresh_dir("rename");
    let book = book(&dir);
    let report = edit(
        &book,
        SheetEdit::Rename {
            sheet: "Data".into(),
            name: Some("売上 2026".into()),
            code_name: Some("Sales".into()),
        },
    )
    .unwrap();
    assert_eq!(report.before, Some(pair("Data", "Sheet2")));
    assert_eq!(
        sheets(&book),
        [pair("Sheet1", "Sheet1"), pair("売上 2026", "Sales")]
    );
    let project = unpack(&book).unwrap().project.unwrap();
    // The module has the new name and its code.
    let module = project
        .modules
        .iter()
        .find(|m| m.name == "Sales")
        .expect("Sales");
    assert_eq!(module.kind, ModuleKind::Document);
    assert!(
        module.source.contains("Attribute VB_Name = \"Sales\"")
            && module.source.contains("Worksheet_Activate"),
        "{}",
        module.source
    );
    assert!(!project.modules.iter().any(|m| m.name == "Sheet2"));
    // The old code name is free again.
    let report = edit(
        &book,
        SheetEdit::Add {
            name: "Next".into(),
            code_name: None,
            place: Place::End,
        },
    )
    .unwrap();
    assert_eq!(report.code_name, "Sheet2");
    // Only the case of the code name: still the same sheet.
    edit(
        &book,
        SheetEdit::Rename {
            sheet: "Sales".into(),
            name: None,
            code_name: Some("SALES".into()),
        },
    )
    .unwrap();
    assert!(sheets(&book).contains(&pair("売上 2026", "SALES")));
}

#[test]
fn a_sheet_is_copied() {
    let dir = fresh_dir("copy");
    let book = book(&dir);
    let report = edit(
        &book,
        SheetEdit::Copy {
            sheet: "Data".into(),
            name: None,
            code_name: None,
            place: None,
        },
    )
    .unwrap();
    assert_eq!(
        (report.name.as_str(), report.code_name.as_str()),
        ("Data (2)", "Sheet3")
    );
    assert_eq!(report.before, Some(pair("Data", "Sheet2")));
    assert_eq!(
        sheets(&book),
        [
            pair("Sheet1", "Sheet1"),
            pair("Data", "Sheet2"),
            pair("Data (2)", "Sheet3")
        ]
    );
    // The copy has a module of its own, without the code of the source (RAS writes it from its file).
    let project = unpack(&book).unwrap().project.unwrap();
    assert!(!project
        .modules
        .iter()
        .find(|m| m.name == "Sheet3")
        .unwrap()
        .source
        .contains("Worksheet_Activate"));
    let report = edit(
        &book,
        SheetEdit::Copy {
            sheet: "Sheet1".into(),
            name: Some("先頭".into()),
            code_name: Some("First".into()),
            place: Some(Place::Before("Sheet1".into())),
        },
    )
    .unwrap();
    assert_eq!(report.name, "先頭");
    assert_eq!(sheets(&book)[0], pair("先頭", "First"));
    assert!(sheet_references(&book, "Data").unwrap().is_empty());
}
