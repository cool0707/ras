//! Conversion of standard modules and of the hand-written code of document modules
//! (`ras_compiler::vba2ras::convert`, spec §14.7.1): the RAS it writes compiles.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::vba2ras::{convert, public_names, Peer, Target};
use ras_compiler::{
    analyze, build, object_source, BuildOptions, Config, FormControl, ModuleKind, SourceInput,
};

/// Compiles `src/<key>.ras` files (and `objects/<name>.ras`, for a key `objects/<name>`) with
/// the sheet code names `Sheet1` and `Sheet2`.
fn compile(src: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let mut config = Config::default();
    config.sheet_code_names = vec!["Sheet1".into(), "Sheet2".into()];
    config.sheet_controls = vec![ras_compiler::FormDecl {
        name: "Sheet2".into(),
        controls: controls(),
        resizable: false,
    }];
    let sources: Vec<SourceInput> = src
        .iter()
        .map(|(key, text)| match key.strip_prefix("objects/") {
            Some(name) => object_source(
                PathBuf::from(format!("objects/{name}.ras")),
                name,
                text.to_string(),
            ),
            None => SourceInput {
                path: PathBuf::from(format!("src/{key}.ras")),
                rel_path: format!("src/{key}.ras"),
                key: key.to_string(),
                kind: ModuleKind::User,
                text: text.to_string(),
            },
        })
        .collect();
    let prog = analyze(sources, config);
    let out = build(
        &prog,
        &BuildOptions {
            erl: Some(false),
            ..Default::default()
        },
    );
    let diags: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    (diags, out.files)
}

fn code(files: &[ras_compiler::OutputFile], name: &str) -> String {
    files
        .iter()
        .find(|f| f.name == name)
        .map(|f| f.code.replace("\r\n", "\n"))
        .unwrap_or_else(|| panic!("no module {name}"))
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

const SHEET1_TARGET: Target = Target::Document {
    code_name: "Sheet1",
    workbook: false,
    controls: &[],
};

fn controls() -> Vec<FormControl> {
    vec![
        FormControl {
            name: "CommandButton1".into(),
            control_type: "CommandButton".into(),
        },
        FormControl {
            name: "TextBox1".into(),
            control_type: "TextBox".into(),
        },
    ]
}

fn peers() -> Vec<Peer> {
    vec![
        Peer {
            vba_name: "Module1".into(),
            key: "Module1".into(),
            document: false,
            names: public_names(MODULE1, &Target::Standard),
            form: false,
        },
        Peer {
            vba_name: "Module2".into(),
            key: "Module2".into(),
            document: false,
            names: public_names(MODULE2, &Target::Standard),
            form: false,
        },
        Peer {
            vba_name: "Sheet1".into(),
            key: "Sheet1".into(),
            document: true,
            names: public_names(SHEET1, &SHEET1_TARGET),
            form: false,
        },
    ]
}

#[test]
fn public_names_of_a_module() {
    let names: Vec<String> = public_names(MODULE1, &Target::Standard)
        .into_iter()
        .map(|n| n.ras)
        .collect();
    assert_eq!(names, ["TAX_RATE", "gCount", "Level", "Total", "Report"]);
    // Event procedures of a sheet are not public names; a Public procedure is.
    let sheet: Vec<String> = public_names(SHEET1, &SHEET1_TARGET)
        .into_iter()
        .map(|n| n.ras)
        .collect();
    assert_eq!(sheet, ["ClearMarks"]);
}

/// A line that starts with a non-ASCII character is not mistaken for (or cut as) a `Rem` line.
#[test]
fn a_line_that_starts_with_a_non_ascii_name_converts() {
    let vba = "Sub Report()\n    Dim 長さ As Long\n    長さ = 3\n    Rem 終わり\nEnd Sub\n";
    let c = convert(vba, &Target::Standard, &peers());
    assert!(c.todos.is_empty(), "{:?}", c.todos);
    assert!(c.text.contains("長さ = 3;"), "{}", c.text);
    assert!(c.text.contains("// 終わり"), "{}", c.text);
}

#[test]
fn modules_and_a_sheet_convert_to_ras_that_compiles() {
    let peers = peers();
    let m1 = convert(MODULE1, &Target::Standard, &peers);
    let m2 = convert(MODULE2, &Target::Standard, &peers);
    let sheet = convert(SHEET1, &SHEET1_TARGET, &peers);
    for (name, c) in [("Module1", &m1), ("Module2", &m2), ("Sheet1", &sheet)] {
        println!("--- {name}\n{}", c.text);
        assert!(c.todos.is_empty(), "{name}: {:?}", c.todos);
    }
    assert!(
        m1.text.contains("export const TAX_RATE: double = 0.1;"),
        "{}",
        m1.text
    );
    assert!(m1.text.contains("export let gCount: int;"), "{}", m1.text);
    assert!(m1.text.contains("\nlet mName: string;"), "{}", m1.text);
    assert!(m1.text.contains("export enum Level {"), "{}", m1.text);
    assert!(
        m1.text
            .contains("export function Total(price: double, qty: int): double {"),
        "{}",
        m1.text
    );
    assert!(
        m1.text.contains("export function Report() {"),
        "{}",
        m1.text
    );
    assert!(m1.text.contains("\nfunction Helper() {"), "{}", m1.text);
    // Names of the other modules are imported; the members of a sheet stay its members.
    assert!(
        m2.text
            .starts_with("import { Report, TAX_RATE, Total } from \"src/Module1\";\n\n"),
        "{}",
        m2.text
    );
    assert!(m2.text.contains("Sheet1.ClearMarks();"), "{}", m2.text);
    // The sheet is `this`; the events are handlers of the sheet's file.
    assert!(
        sheet
            .text
            .contains("@on(\"Worksheet.Change\")\nfunction onChange(Target: Excel.Range) {"),
        "{}",
        sheet.text
    );
    assert!(sheet.text.contains("@on(\"Worksheet.BeforeDoubleClick\")\nfunction onBeforeDoubleClick(Target: Excel.Range, byref Cancel: boolean) {"), "{}", sheet.text);
    assert!(
        sheet.text.contains("this.range(\"A1:A10\")"),
        "{}",
        sheet.text
    );
    assert!(
        sheet
            .text
            .contains("this.cells(1, 2).value = Target.count;"),
        "{}",
        sheet.text
    );
    assert!(
        sheet
            .text
            .contains("this.cells(Target.row, 3).value = this.name;"),
        "{}",
        sheet.text
    );
    assert!(
        sheet
            .text
            .contains("import { Report } from \"src/Module1\";"),
        "{}",
        sheet.text
    );
    assert!(
        sheet.text.contains("export function ClearMarks() {"),
        "{}",
        sheet.text
    );

    let (diags, files) = compile(&[
        ("Module1", &m1.text),
        ("Module2", &m2.text),
        ("objects/Sheet1", &sheet.text),
    ]);
    assert!(diags.is_empty(), "{diags:?}");
    let doc = code(&files, "Sheet1");
    println!("{doc}");
    assert!(
        doc.contains("Private Sub Worksheet_Change(ByVal Target As Range)"),
        "{doc}"
    );
    assert!(doc.contains("Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)"), "{doc}");
    // The public procedure stays a member of the sheet: `Sheet1.ClearMarks` as before.
    assert!(doc.contains("Public Sub ClearMarks("), "{doc}");
    assert!(code(&files, "Module2").contains("Sheet1.ClearMarks"));
    assert!(!files.iter().any(|f| f.name == "Sheet1Events"));
}

/// Controls on the sheet (spec §12.2.1) and what a document module cannot make public.
#[test]
fn controls_and_public_declarations_of_a_sheet() {
    let vba = "Public Total As Long\n\nPrivate Sub CommandButton1_Click()\n    TextBox1.Text = \"\"\n    Me.CommandButton1.Caption = Range(\"A1\").Value\n    Total = Total + 1\nEnd Sub\n\nPrivate Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)\n    If KeyCode = 13 Then Me.Activate\nEnd Sub\n\nPrivate Sub TextBox1_LostFocus()\nEnd Sub\n";
    let list = controls();
    let c = convert(
        vba,
        &Target::Document {
            code_name: "Sheet2",
            workbook: false,
            controls: &list,
        },
        &[],
    );
    println!("{}\n{:?}", c.text, c.todos);
    assert!(
        c.text
            .contains("@on(\"CommandButton1.Click\")\nfunction commandButton1Click() {"),
        "{}",
        c.text
    );
    assert!(c.text.contains("this.TextBox1.text = \"\";"), "{}", c.text);
    assert!(
        c.text
            .contains("this.CommandButton1.caption = this.range(\"A1\").value;"),
        "{}",
        c.text
    );
    assert!(c.text.contains("@on(\"TextBox1.KeyDown\")"), "{}", c.text);
    assert!(c.text.contains("@on(\"TextBox1.LostFocus\")"), "{}", c.text);
    // A Public variable cannot be a member of the sheet: a note to see to.
    assert_eq!(c.todos.len(), 1, "{:?}", c.todos);
    assert!(c.todos[0].1.contains("Public 変数"), "{:?}", c.todos);
    assert!(c.text.contains("\nlet Total: int;"), "{}", c.text);
    let (diags, files) = compile(&[("objects/Sheet2", &c.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let doc = code(&files, "Sheet2");
    assert!(
        doc.contains("Private Sub CommandButton1_click()")
            && doc.contains("Private Sub TextBox1_lostFocus()"),
        "{doc}"
    );
}

/// `WithEvents` variables (spec §12.2.2) and controls whose events are not known.
#[test]
fn with_events_and_other_controls() {
    let vba = "Private WithEvents App As Application\n\nPrivate Sub Workbook_Open()\n    Set App = Application\nEnd Sub\n\nPrivate Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n    Debug.Print Target.Address\nEnd Sub\n\nPrivate Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)\n    Cancel = True\nEnd Sub\n";
    let c = convert(
        vba,
        &Target::Document {
            code_name: "ThisWorkbook",
            workbook: true,
            controls: &[],
        },
        &[],
    );
    println!("{}\n{:?}", c.text, c.todos);
    assert!(c.todos.is_empty(), "{:?}", c.todos);
    assert!(c.text.contains("let App: Excel.Application;"), "{}", c.text);
    assert!(
        c.text.contains(
            "@on(\"App.SheetChange\")\nfunction appSheetChange(Sh: object, Target: Excel.Range) {"
        ),
        "{}",
        c.text
    );
    assert!(c.text.contains("@on(\"App.WorkbookBeforeClose\")\nfunction appWorkbookBeforeClose(Wb: Excel.Workbook, byref Cancel: boolean) {"), "{}", c.text);
    let (diags, files) = compile(&[("objects/ThisWorkbook", &c.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let doc = code(&files, "ThisWorkbook");
    assert!(
        doc.contains("Private WithEvents App As Application")
            && doc
                .contains("Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)"),
        "{doc}"
    );

    // A control that is not an MSForms one: the handler keeps the declaration of the procedure.
    let list = vec![FormControl {
        name: "Calendar1".into(),
        control_type: String::new(),
    }];
    let c = convert("Private Sub Calendar1_DateClick(ByVal DateClicked As Date)\n    Range(\"A1\").Value = DateClicked\nEnd Sub\n", &Target::Document { code_name: "Sheet2", workbook: false, controls: &list }, &[]);
    assert!(c.todos.is_empty(), "{:?}", c.todos);
    assert!(
        c.text.contains(
            "@on(\"Calendar1.DateClick\")\nfunction calendar1DateClick(DateClicked: date) {"
        ),
        "{}",
        c.text
    );
}

#[test]
fn this_workbook_converts_and_unknown_events_are_todos() {
    let wb = convert(
        THIS_WORKBOOK,
        &Target::Document {
            code_name: "ThisWorkbook",
            workbook: true,
            controls: &[],
        },
        &[],
    );
    println!("{}\n{:?}", wb.text, wb.todos);
    assert!(
        wb.text
            .contains("@on(\"Workbook.Open\")\nfunction onOpen() {"),
        "{}",
        wb.text
    );
    assert!(
        wb.text.contains("this.worksheets(1).activate();"),
        "{}",
        wb.text
    );
    assert!(
        wb.text.contains("application.statusBar = this.name;"),
        "{}",
        wb.text
    );
    assert_eq!(wb.todos.len(), 1, "{:?}", wb.todos);
    assert!(
        wb.todos[0].1.contains("対応していないイベント"),
        "{:?}",
        wb.todos
    );
    let lines: Vec<&str> = wb.text.lines().collect();
    assert!(lines[wb.todos[0].0 - 1].starts_with("// TODO(VBA):"));
    let (diags, files) = compile(&[("objects/ThisWorkbook", &wb.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    assert!(code(&files, "ThisWorkbook").contains("Private Sub Workbook_Open()"));
}

#[test]
fn me_in_a_standard_module_is_a_todo() {
    let c = convert("Sub A()\n    Me.Hide\nEnd Sub\n", &Target::Standard, &[]);
    assert_eq!(c.todos.len(), 1, "{c:?}");
}

const SCRIPTING: &str = r#"Option Explicit

Public Function ReadNames(ByVal path As String) As Long
    Dim fso As New Scripting.FileSystemObject
    Dim ts As TextStream
    Dim seen As Dictionary
    Dim re As RegExp
    Dim d As Object
    Set seen = New Dictionary
    Set re = New RegExp
    Set d = CreateObject("Scripting.Dictionary")
    re.Pattern = "^\s*$"
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, ForReading)
    Do Until ts.AtEndOfStream
        Dim line As String
        line = ts.ReadLine
        If Not re.Test(line) Then
            If Not seen.Exists(line) Then seen.Add line, True
        End If
    Loop
    ts.Close
    ReadNames = seen.Count + d.Count
End Function
"#;

#[test]
fn scripting_types_convert_and_stay_late_bound() {
    let m = convert(SCRIPTING, &Target::Standard, &[]);
    println!("{}\n{:?}", m.text, m.todos);
    assert!(m.todos.is_empty(), "{:?}", m.todos);
    for want in [
        "new Scripting.FileSystemObject()",
        ": Scripting.TextStream",
        "new Scripting.Dictionary()",
        "new VBScript.RegExp()",
        "fso.openTextFile(path, forReading)",
        "re.test(line)",
    ] {
        assert!(m.text.contains(want), "{want}\n{}", m.text);
    }
    let (diags, files) = compile(&[("Module1", &m.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let vba = code(&files, "Module1");
    assert!(
        vba.contains("CreateObject(\"Scripting.FileSystemObject\")")
            && vba.contains("CreateObject(\"VBScript.RegExp\")"),
        "{vba}"
    );
    assert!(
        vba.contains("fso.OpenTextFile(path_, 1&)") && !vba.contains("As TextStream"),
        "{vba}"
    );
}

const FILE_IO: &str = r#"Option Explicit

Public Sub SaveAndLoad(ByVal path As String)
    Dim f As Integer
    Dim s As String, n As Long, ok As Boolean
    Dim buf() As Byte
    f = FreeFile
    Open path For Output As #f
    Print #f, "name"; 1
    Print #f, "partial";
    Print #f,
    Write #f, "a", 2, True
    Write #f, "b";
    Close #f
    Open path For Input Access Read Shared As #f
    Do While Not EOF(f)
        Line Input #f, s
        Debug.Print s
    Loop
    Seek #f, 1
    Input #f, s, n, ok
    s = Input(3, #f)
    Debug.Print Seek(f), LOF(f), Loc(f)
    Close
    Open path & ".bin" For Binary Lock Write As 2
    Put #2, , n
    Get #2, 1, n
    Lock #2, 1 To 4
    Unlock #2
    Width #2, 80
    Close #1, #2
    Open path & ".dat" For Random As #3 Len = 32
    Put 3, 2, s
    Reset
    Name path As path & ".old"
    Kill path & ".*"
    If Dir(path) = "" Then MkDir path & "_dir"
    Print #3, "a", "b"
    Debug.Print DateDiff("d", #1/2/2020#, #Jan 3, 2020#)
End Sub
"#;

#[test]
fn file_statements_convert_to_the_file_functions() {
    let m = convert(FILE_IO, &Target::Standard, &[]);
    println!("{}\n{:?}", m.text, m.todos);
    // `Print #3, "a", "b"` (tab positions) is left as a TODO; after `(` and `,` a `#…#` with
    // a date in it is still a date literal (left as a TODO as before), not a file number.
    assert_eq!(m.todos.len(), 2, "{:?}", m.todos);
    assert!(m.todos[0].1.contains("Print #"), "{:?}", m.todos);
    assert!(m.todos[1].1.contains("日付リテラル"), "{:?}", m.todos);
    for want in [
        "f = freeFile();",
        "fileOpen(f, path, forOutput);",
        "printLine(f, \"name\", 1);",
        "print(f, \"partial\");",
        "printLine(f);",
        "writeLine(f, \"a\", 2, true);",
        "write(f, \"b\");",
        "fileClose(f);",
        "fileOpen(f, path, forInput, accessRead, lockShared);",
        "while (!eof(f)) {",
        "s = lineInput(f);",
        "seek(f, 1);",
        "input(f, s, n, ok);",
        "s = inputString(f, 3);",
        "fileClose();",
        "fileOpen(2, `${path}.bin`, forBinary, accessDefault, lockWrite);",
        "filePut(2, n);",
        "fileGet(2, n, 1);",
        "lock(2, 1, 4);",
        "unlock(2);",
        "fileWidth(2, 80);",
        "fileClose(1, 2);",
        "fileOpen(3, `${path}.dat`, forRandom, accessDefault, lockDefault, 32);",
        "filePut(3, s, 2);",
        "reset();",
        "rename(path, `${path}.old`);",
        "kill(`${path}.*`);",
    ] {
        assert!(m.text.contains(want), "{want}\n{}", m.text);
    }
    // And back to VBA.
    let (diags, files) = compile(&[("Module1", &m.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let vba = code(&files, "Module1");
    for want in [
        "Open path_ For Output As #f",
        "Print #f, \"name\"; CStr(1&)",
        "Print #f, \"partial\";",
        "Write #f, \"a\", 2&, True",
        "Open path_ For Input Access Read Shared As #f",
        "Line Input #f, ",
        "Input #f, s, n, ok",
        "Input$(3&, #f)",
        "Open path_ & \".bin\" For Binary Lock Write As #2",
        "Get #2, 1&, n",
        "Lock #2, 1& To 4&",
        "Open path_ & \".dat\" For Random As #3 Len = 32&",
        "Put #3, 2&, s",
        "Name path_ As path_ & \".old\"",
    ] {
        assert!(vba.contains(want), "{want}\n{vba}");
    }
}

/// `[A1]` / `[SUM(A1:A3)]` (the shorthand of `Evaluate`) become `` fx`…` `` (spec §2.7);
/// a name written in brackets stays a name.
#[test]
fn bracket_evaluate_becomes_fx() {
    let vba = "Public Sub Report()\n    Dim total As Double\n    total = [SUM(A1:A3)] + [A1] + [$B$2]\n    Range(\"C1\").Value = [Sheet1!A1:B2].Rows.Count\n    Debug.Print [COUNTIF(A:A,\"`x`\")]\n    Debug.Print [total]\nEnd Sub\n";
    let c = convert(vba, &Target::Standard, &[]);
    println!("{}\n{:?}", c.text, c.todos);
    assert!(
        c.text.contains("fx`SUM(A1:A3)` + fx`A1` + fx`$B$2`"),
        "{}",
        c.text
    );
    assert!(c.text.contains("fx`Sheet1!A1:B2`).rows"), "{}", c.text);
    assert!(c.text.contains("debug.print(total);"), "{}", c.text);
    assert!(c.text.contains(r#"fx`COUNTIF(A:A,"\`x\`")`"#), "{}", c.text);
    let (diags, files) = compile(&[("Module1", &c.text)]);
    assert!(diags.is_empty(), "{diags:?}\n{}", c.text);
    let vba = code(&files, "Module1");
    assert!(
        vba.contains(r#"Evaluate("SUM(A1:A3)") + Evaluate("A1") + Evaluate("$B$2")"#),
        "{vba}"
    );
    assert!(vba.contains(r#"Evaluate("COUNTIF(A:A,""`x`"")")"#), "{vba}");
}

const CLASS_MODULE: &str = r#"Option Explicit

Public Name As String
Private Items_ As Object
Private Counts_() As Long

Private Sub Class_Initialize()
    Set Items_ = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get Count() As Long
    Count = Items_.Count
End Property

Public Property Get Item(ByVal Key As String) As Variant
    Item = Items_(Key)
End Property

Public Property Let Item(ByVal Key As String, ByVal Value As Variant)
    Items_(Key) = Value
End Property

Public Function Describe() As String
    Describe = Name & ": " & Suffix()
End Function

Private Function Suffix() As String
    If Name Like "1*" Then Suffix = "!"
End Function
"#;

/// A class module becomes a class: fields, accessors (with parameters), the constructor.
#[test]
fn a_class_module_converts_to_a_class() {
    let target = Target::Class { name: "Basket" };
    let class = Peer {
        vba_name: "Basket".into(),
        key: "Basket".into(),
        document: false,
        names: public_names(CLASS_MODULE, &target),
        form: false,
    };
    assert!(class.names[0].class && class.names[0].members.contains(&"Describe()".to_string()));
    let c = convert(CLASS_MODULE, &target, std::slice::from_ref(&class));
    println!("{}", c.text);
    assert!(c.todos.is_empty(), "{:?}", c.todos);
    for want in [
        "export class Basket {",
        "    Name: string;",
        "    private Items_: object;",
        "    private Counts_: int[] = [];",
        "    constructor() {",
        "    get Count(): int {",
        "    get Item(Key: string): any {",
        "    set Item(Key: string, Value: any) {",
        "        this.Items_[Key] = Value;",
        "    private Suffix(): string {",
        "like(",
    ] {
        assert!(c.text.contains(want), "{want}:\n{}", c.text);
    }
    // A module that uses the class: its type is imported, its members keep their names.
    let user = convert(
        "Sub Fill()\n    Dim b As Basket\n    Set b = New Basket\n    b.Item(\"a\") = 1\n    Debug.Print b.Describe\nEnd Sub\n",
        &Target::Standard,
        std::slice::from_ref(&class),
    );
    println!("{}", user.text);
    assert!(
        user.text.contains("import { Basket } from \"src/Basket\";"),
        "{}",
        user.text
    );
    assert!(user.text.contains("b = new Basket();"), "{}", user.text);
    assert!(user.text.contains("b.Item(\"a\") = 1;"), "{}", user.text);
    let (diags, files) = compile(&[("Basket", &c.text), ("Main", &user.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let class = code(&files, "Basket");
    assert!(
        class.contains("Public Property Let Item(ByVal Key As String, "),
        "{class}"
    );
}

const API_MODULE: &str = r#"Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Enum Level
    lvLow = 1
    lvHigh
End Enum

Public Function Elapsed(ByVal since As Long) As Long
    Elapsed = GetTickCount() - since
End Function

Public Function Pick(ByVal n As Long) As Level
    If n > 1 Then GoTo High
    Pick = lvLow
    Exit Function
High:
    Pick = lvHigh
    GoTo High2
High2:
End Function

Public Sub Back()
Again:
    If Elapsed(0) < 0 Then GoTo Again
    Debug.Print Level.lvHigh, lvLow
End Sub
"#;

/// The take-over keeps as VBA what RAS cannot express (spec §14.7.4): a `Declare`, the
/// procedure that calls it, a procedure with a backward `GoTo`.
#[test]
fn what_ras_cannot_express_stays_vba() {
    let m = ras_compiler::vba2ras::convert_keeping(API_MODULE, &Target::Standard, &[], &[]);
    println!("{}", m.text);
    assert!(m.todos.is_empty(), "{:?}", m.todos);
    assert_eq!(m.raw, ["Elapsed", "Back"]);
    for want in [
        "vba`\nPrivate Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\n`;",
        "declare function GetTickCount(): int;",
        "export declare function Elapsed(since: int): int;",
        "export declare function Back(): void;",
        // The names of the project, as the generated VBA spells them.
        "    Debug.Print Level_lvHigh, Level_lvLow",
        // A forward jump is a block.
        "toHigh: do {",
    ] {
        assert!(m.text.contains(want), "{want}:\n{}", m.text);
    }
    let (diags, files) = compile(&[("Api", &m.text)]);
    assert!(diags.is_empty(), "{diags:?}");
    let api = code(&files, "Api");
    assert!(
        api.contains("Private Declare PtrSafe Function GetTickCount"),
        "{api}"
    );
    assert!(
        api.contains("    If Elapsed(0) < 0 Then GoTo Again"),
        "{api}"
    );
    // Without the fallback (`ras form pull`) they are TODO comments.
    let plain = convert(API_MODULE, &Target::Standard, &[]);
    assert!(
        plain.raw.is_empty() && !plain.todos.is_empty(),
        "{:?}",
        plain.todos
    );
}

/// A class with events of its own (spec §12.2.2), a `WithEvents` field, and the `Not Not a`
/// test of an array that has been dimensioned.
#[test]
fn a_class_with_events_converts() {
    let vba = "Option Explicit\n\nPublic Event Changed(ByVal Index As Long, ByRef Cell As Range)\n\nPrivate WithEvents Sheet_ As Worksheet\nPrivate Cache_() As Variant\n\nPublic Sub Flush()\n    If Not Not Cache_ Then Erase Cache_\n    If (Not Cache_) = -1 Then Debug.Print \"empty\"\nEnd Sub\n\nPrivate Sub Sheet__Change(ByVal Target As Range)\n    RaiseEvent Changed(Target.Row, Target)\nEnd Sub\n";
    let c = convert(vba, &Target::Class { name: "Helper" }, &[]);
    println!("{}", c.text);
    assert!(c.todos.is_empty(), "{:?}", c.todos);
    for want in [
        "    event Changed(Index: int, Cell: Excel.Range);",
        "    private Sheet_: Excel.Worksheet;",
        "        if (isAllocated(this.Cache_)) {",
        "        if (((isAllocated(this.Cache_) ? 1 : -1)) == -1) {",
        "    @on(\"Sheet_.Change\")",
        "    private sheet_Change(Target: Excel.Range) {",
        "        this.Changed(Target.row, Target);",
    ] {
        assert!(c.text.contains(want), "missing `{want}` in\n{}", c.text);
    }
}
