//! Conversion of standard modules and of the hand-written code of document modules
//! (`ras_compiler::vba2ras::convert`, spec §14.7.1): the RAS it writes compiles.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::vba2ras::{convert, public_names, Peer, Target};
use ras_compiler::{analyze, build, BuildOptions, Config, ModuleKind, SourceInput};

/// Compiles `src/<key>.ras` files with the sheet code names `Sheet1` and `Sheet2`.
fn compile(src: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let mut config = Config::default();
    config.sheet_code_names = vec!["Sheet1".into(), "Sheet2".into()];
    let sources: Vec<SourceInput> = src
        .iter()
        .map(|(key, text)| SourceInput { path: PathBuf::from(format!("src/{key}.ras")), rel_path: format!("src/{key}.ras"), key: key.to_string(), kind: ModuleKind::User, text: text.to_string() })
        .collect();
    let prog = analyze(sources, config);
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    let diags: Vec<String> = prog.diagnostics.iter().chain(out.diagnostics.iter()).filter(|d| d.is_error()).map(|d| format!("{}: {}", d.code, d.message)).collect();
    (diags, out.files)
}

fn code(files: &[ras_compiler::OutputFile], name: &str) -> String {
    files.iter().find(|f| f.name == name).map(|f| f.code.replace("\r\n", "\n")).unwrap_or_else(|| panic!("no module {name}"))
}

const MODULE1: &str = r#"Option Explicit

Public Const TAX_RATE As Double = 0.1
Public gCount As Long
Private mName As String

Public Enum Level
    lvLow = 1
    lvHigh
End Enum

' 合計を計算する
Public Function Total(ByVal price As Double, ByVal qty As Long) As Double
    Total = price * qty * (1 + TAX_RATE)
End Function

Sub Report()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    gCount = gCount + 1
    ws.Range("A1").Value = Total(100, 2)
    MsgBox "done: " & gCount
End Sub

Private Sub Helper()
    mName = "x"
End Sub
"#;

const MODULE2: &str = r#"Public Sub Run()
    Module1.Report
    Report
    Debug.Print Total(1, 2), Module1.gCount, TAX_RATE
    Sheet1.ClearMarks
End Sub
"#;

const SHEET1: &str = r#"Option Explicit

Private busy As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    If busy Then Exit Sub
    If Not Intersect(Target, Range("A1:A10")) Is Nothing Then
        Me.Cells(1, 2).Value = Target.Count
        Report
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
    Cells(Target.Row, 3).Value = Name
End Sub

Public Sub ClearMarks()
    Range("C:C").ClearContents
End Sub
"#;

const THIS_WORKBOOK: &str = r#"Private Sub Workbook_Open()
    Worksheets(1).Activate
    Application.StatusBar = Me.Name
End Sub

Private Sub Workbook_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
End Sub
"#;

fn peers() -> Vec<Peer> {
    vec![
        Peer { vba_name: "Module1".into(), key: "Module1".into(), names: public_names(MODULE1, &Target::Standard) },
        Peer { vba_name: "Module2".into(), key: "Module2".into(), names: public_names(MODULE2, &Target::Standard) },
        Peer { vba_name: "Sheet1".into(), key: "Sheet1Events".into(), names: public_names(SHEET1, &Target::Document { code_name: "Sheet1", workbook: false }) },
    ]
}

#[test]
fn public_names_of_a_module() {
    let names: Vec<String> = public_names(MODULE1, &Target::Standard).into_iter().map(|n| n.ras).collect();
    assert_eq!(names, ["TAX_RATE", "gCount", "Level", "Total", "Report"]);
    // Event procedures of a sheet are not public names; a Public procedure is.
    let sheet: Vec<String> = public_names(SHEET1, &Target::Document { code_name: "Sheet1", workbook: false }).into_iter().map(|n| n.ras).collect();
    assert_eq!(sheet, ["ClearMarks"]);
}

#[test]
fn modules_and_a_sheet_convert_to_ras_that_compiles() {
    let peers = peers();
    let m1 = convert(MODULE1, &Target::Standard, &peers);
    let m2 = convert(MODULE2, &Target::Standard, &peers);
    let sheet = convert(SHEET1, &Target::Document { code_name: "Sheet1", workbook: false }, &peers);
    for (name, c) in [("Module1", &m1), ("Module2", &m2), ("Sheet1Events", &sheet)] {
        println!("--- {name}\n{}", c.text);
        assert!(c.todos.is_empty(), "{name}: {:?}", c.todos);
    }
    assert!(m1.text.contains("export const TAX_RATE: double = 0.1;"), "{}", m1.text);
    assert!(m1.text.contains("export let gCount: int;"), "{}", m1.text);
    assert!(m1.text.contains("\nlet mName: string;"), "{}", m1.text);
    assert!(m1.text.contains("export enum Level {"), "{}", m1.text);
    assert!(m1.text.contains("export function Total(price: double, qty: int): double {"), "{}", m1.text);
    assert!(m1.text.contains("export function Report() {"), "{}", m1.text);
    assert!(m1.text.contains("\nfunction Helper() {"), "{}", m1.text);
    // Names of the other modules are imported.
    assert!(m2.text.starts_with("import { Report, TAX_RATE, Total } from \"./Module1\";\nimport * as Module1 from \"./Module1\";\nimport * as Sheet1Events from \"./Sheet1Events\";\n\n"), "{}", m2.text);
    assert!(m2.text.contains("Sheet1Events.ClearMarks();"), "{}", m2.text);
    // The members of the sheet are those of `Sheet1`; the events are handlers.
    assert!(sheet.text.contains("@on(\"Sheet1.Change\")\nfunction sheet1Change(Target: Excel.Range) {"), "{}", sheet.text);
    assert!(sheet.text.contains("@on(\"Sheet1.BeforeDoubleClick\")\nfunction sheet1BeforeDoubleClick(Target: Excel.Range, byref Cancel: boolean) {"), "{}", sheet.text);
    assert!(sheet.text.contains("Sheet1.range(\"A1:A10\")"), "{}", sheet.text);
    assert!(sheet.text.contains("Sheet1.cells(1, 2).value = Target.count;"), "{}", sheet.text);
    assert!(sheet.text.contains("Sheet1.cells(Target.row, 3).value = Sheet1.name;"), "{}", sheet.text);
    assert!(sheet.text.contains("import { Report } from \"./Module1\";"), "{}", sheet.text);

    let (diags, files) = compile(&[("Module1", &m1.text), ("Module2", &m2.text), ("Sheet1Events", &sheet.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let doc = code(&files, "Sheet1");
    println!("{doc}");
    assert!(doc.contains("Private Sub Worksheet_Change(ByVal Target As Range)"), "{doc}");
    assert!(doc.contains("Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)"), "{doc}");
    assert!(code(&files, "Sheet1Events").contains("Public Sub ClearMarks("));
}

#[test]
fn this_workbook_converts_and_unknown_events_are_todos() {
    let wb = convert(THIS_WORKBOOK, &Target::Document { code_name: "ThisWorkbook", workbook: true }, &[]);
    println!("{}\n{:?}", wb.text, wb.todos);
    assert!(wb.text.contains("@on(\"Workbook.Open\")\nfunction workbookOpen() {"), "{}", wb.text);
    assert!(wb.text.contains("thisWorkbook.worksheets(1).activate();"), "{}", wb.text);
    assert!(wb.text.contains("application.statusBar = thisWorkbook.name;"), "{}", wb.text);
    assert_eq!(wb.todos.len(), 1, "{:?}", wb.todos);
    assert!(wb.todos[0].1.contains("対応していないイベント"), "{:?}", wb.todos);
    let lines: Vec<&str> = wb.text.lines().collect();
    assert!(lines[wb.todos[0].0 - 1].starts_with("// TODO(VBA):"));
    let (diags, files) = compile(&[("ThisWorkbookEvents", &wb.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    assert!(code(&files, "ThisWorkbook").contains("Private Sub Workbook_Open()"));
}

#[test]
fn me_in_a_standard_module_is_a_todo() {
    let c = convert("Sub A()\n    Me.Hide\nEnd Sub\n", &Target::Standard, &[]);
    assert_eq!(c.todos.len(), 1, "{c:?}");
}
