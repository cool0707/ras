//! Form code-behinds (`forms/<Name>.ras`, spec §14.11).

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{analyze, build, form_source, BuildOptions, Config, FormControl, FormDecl, ModuleKind, SourceInput};

fn control(name: &str, ty: &str) -> FormControl {
    FormControl { name: name.into(), control_type: ty.into() }
}

fn compile(form: &str, src: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmInput".into(),
        controls: vec![control("txtName", "TextBox"), control("cboKind", "ComboBox"), control("btnOK", "CommandButton"), control("pages", "MultiPage"), control("page1", "Page")],
    }];
    let mut sources: Vec<SourceInput> = src
        .iter()
        .map(|(key, text)| SourceInput { path: PathBuf::from(format!("src/{key}.ras")), rel_path: format!("src/{key}.ras"), key: key.to_string(), kind: ModuleKind::User, text: text.to_string() })
        .collect();
    sources.push(form_source(PathBuf::from("forms/frmInput.ras"), "frmInput", form.to_string()));
    let prog = analyze(sources, config);
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    let mut diags: Vec<String> = prog.diagnostics.iter().chain(out.diagnostics.iter()).filter(|d| d.module.is_none_or(|m| prog.modules[m].kind == ModuleKind::User)).map(|d| format!("{}: {}", d.code, d.message)).collect();
    diags.sort();
    (diags, out.files)
}

fn form_code(files: &[ras_compiler::OutputFile]) -> String {
    files.iter().find(|f| f.name == "frmInput").map(|f| f.code.replace("\r\n", "\n")).unwrap_or_default()
}

#[test]
fn handlers_become_event_procedures() {
    let (diags, files) = compile(
        r#"
let count = 0;

@on("btnOK.Click")
function ok() {
    count++;
    this.txtName.text = `${count}`;
    unload(this);
}

@on("txtName.Exit")
function leave(cancel: MSForms.ReturnBoolean) {
    if (this.txtName.text == "") cancel.value = true;
}

@on("UserForm.QueryClose")
function closing(byref cancel: short, byref mode: short) {
    if (mode == vbFormControlMenu) cancel = 1;
}

export function selected(): string {
    return this.cboKind.text;
}
"#,
        &[("main", "export function run() {\n    frmInput.show();\n    msgBox(frmInput.selected());\n}\n")],
    );
    assert!(diags.is_empty(), "{diags:?}");
    let code = form_code(&files);
    println!("{code}");
    assert!(code.contains("Private Sub btnOK_Click()"), "{code}");
    assert!(code.contains("Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)"), "{code}");
    assert!(code.contains("Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)"), "{code}");
    assert!(code.contains("Private Sub UserForm_Initialize()"), "{code}");
    assert!(code.contains("Me.txtName.Text = "), "{code}");
    assert!(code.contains("Unload Me"), "{code}");
    assert!(code.contains("Public Function selected() As String"), "{code}");
    assert!(!code.contains("Attribute VB_Name"));
    let main = files.iter().find(|f| f.name == "Main").unwrap();
    println!("{}", main.code);
    assert!(main.code.contains("frmInput.Show"), "{}", main.code);
    assert!(main.code.contains("frmInput.selected"), "{}", main.code);
}

#[test]
fn errors_in_code_behind() {
    let (diags, _) = compile("class Foo {}\nconsole.log(1);\n@macro\nfunction m() {}\nexport let x = 1;\n", &[]);
    let codes: Vec<&str> = diags.iter().map(|d| d.split(':').next().unwrap()).collect();
    assert_eq!(codes, ["E0630", "E0630", "E0630", "E0632"], "{diags:?}");
}

#[test]
fn errors_in_handlers() {
    let (diags, _) = compile(
        r#"
@on("btnNo.Click")
function a() {}
@on("btnOK.Change")
function b() {}
@on("page1.Click")
function c() {}
@on("txtName.Exit")
function d(x: MSForms.ReturnBoolean, y: int) {}
function btnOK_Click() {}
@on("btnOK.Click")
function e() {}
"#,
        &[],
    );
    let codes: Vec<&str> = diags.iter().map(|d| d.split(':').next().unwrap()).collect();
    for code in ["E0613", "E0611", "E0612", "E0614"] {
        assert!(codes.contains(&code), "{code} missing in {diags:?}");
    }
}

#[test]
fn missing_design() {
    let mut config = Config::default();
    config.forms.clear();
    let prog = analyze(vec![form_source(PathBuf::from("forms/frmX.ras"), "frmX", "function f() {}\n".into())], config);
    assert!(prog.diagnostics.iter().any(|d| d.code == "E0631"));
}

const SAMPLE_VBA: &str = r#"Option Explicit

' 入力された件数
Private mCount As Long
Private Const TITLE As String = "入力"

Private Sub UserForm_Initialize()
    Me.Caption = TITLE
    txtName.Text = ""
    With cboKind
        .AddItem "通常"
        .AddItem "お急ぎ"
        .ListIndex = 0
    End With
End Sub

Private Sub btnOK_Click()
    Dim msg As String, i As Long
    If Len(Trim(txtName.Text)) = 0 Then
        MsgBox "名前を入力してください", vbExclamation, TITLE
        txtName.SetFocus
        Exit Sub
    ElseIf cboKind.ListIndex < 0 Then
        Exit Sub
    End If
    For i = 1 To 3
        mCount = mCount + i
    Next i
    msg = "こんにちは、" & txtName.Text & " さん（" & mCount & "）"
    Select Case cboKind.ListIndex
        Case 0: Debug.Print "normal"
        Case 1, 2
            Debug.Print "rush"
        Case Else
    End Select
    Call Report(msg)
    Unload Me
End Sub

Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txtName.Text = "" Then Cancel = True Else txtName.Text = UCase(txtName.Text)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        If MsgBox("閉じますか？", vbYesNo) = vbNo Then Cancel = True
    End If
End Sub

Public Function Total(ByVal factor As Double) As Double
    If factor <= 0 Then
        Total = 0
        Exit Function
    End If
    Total = mCount * factor / 2
End Function

Private Sub Report(ByVal text As String)
    On Error GoTo Failed
    Dim items As Collection
    Set items = New Collection
    items.Add text
    Do While items.Count > 3
        items.Remove 1
    Loop
    Exit Sub
Failed:
    MsgBox "エラー: " & Err.Description
End Sub
"#;

#[test]
fn vba_code_behind_converts_to_ras_that_compiles() {
    let decl = FormDecl {
        name: "frmInput".into(),
        controls: vec![control("txtName", "TextBox"), control("cboKind", "ComboBox"), control("btnOK", "CommandButton"), control("pages", "MultiPage"), control("page1", "Page")],
    };
    let converted = ras_compiler::vba2ras::convert_form(SAMPLE_VBA, &decl);
    println!("{}", converted.text);
    assert!(converted.todos.is_empty(), "{:?}", converted.todos);
    let (diags, files) = compile(&converted.text, &[]);
    assert!(diags.iter().all(|d| !d.starts_with('E')), "{diags:?}");
    let code = form_code(&files);
    println!("{code}");
    for proc in ["Private Sub UserForm_Initialize()", "Private Sub btnOK_Click()", "Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)", "Private Sub UserForm_QueryClose(", "Public Function Total("] {
        assert!(code.contains(proc), "{proc} missing:\n{code}");
    }
}

const HARD_VBA: &str = r#"Private Type Pt
    X As Long
End Type
Private Enum Mode
    mdA = 1
    mdB
End Enum
Dim names() As String
Private Sub btnOK_Click()
    Dim i As Integer, s$, v As Variant, c As Control, ok As Boolean
    On Error Resume Next
    Me.Controls("txtName").Value = "x"
    On Error GoTo 0
    v = Split("a,b,c", ",")
    For i = UBound(v) To 0 Step -1
        If v(i) = "b" Then s = s & v(i): Exit For
    Next
    For Each c In Me.Controls
        c.Visible = (c.Tag <> "hidden") And Not ok
    Next c
    Do
        i = i + 1
    Loop Until i >= 10 Or ok
    s = Format(Now, "yyyy/mm/dd") & _
        " " & Left$(s, 2)
    If s Like "a*" Then MsgBox "like"
    ReDim names(10)
    GoTo Done
Done:
    MsgBox Prompt:="done", Title:="t"
    ok = (i Mod 2 = 0) Xor ok
    i = &HFF And 15
    txtName = "abc"
    If chk Then i = 1
    Debug.Print txtName.Text, cboKind.Text
End Sub
Property Get Current() As Mode
    Current = mdB
End Property
"#;

#[test]
fn unsupported_vba_is_kept_as_todo() {
    let decl = FormDecl {
        name: "frmInput".into(),
        controls: vec![control("txtName", "TextBox"), control("cboKind", "ComboBox"), control("btnOK", "CommandButton"), control("chk", "CheckBox")],
    };
    let converted = ras_compiler::vba2ras::convert_form(HARD_VBA, &decl);
    println!("{}", converted.text);
    println!("{:?}", converted.todos);
    let reasons: Vec<&str> = converted.todos.iter().map(|t| t.1.as_str()).collect();
    for needle in ["ユーザー定義型", "Like", "ReDim", "GoTo", "行ラベル", "Property"] {
        assert!(reasons.iter().any(|r| r.contains(needle)), "{needle} not reported: {reasons:?}");
    }
    // Every TODO line points at the comment.
    let lines: Vec<&str> = converted.text.lines().collect();
    for (line, _) in &converted.todos {
        assert!(lines[line - 1].trim_start().starts_with("// TODO(VBA):"), "{}", lines[line - 1]);
    }
    let mut config = Config::default();
    config.forms = vec![decl];
    let prog = analyze(vec![form_source(PathBuf::from("forms/frmInput.ras"), "frmInput", converted.text.clone())], config);
    let errors: Vec<String> = prog.diagnostics.iter().filter(|d| d.is_error()).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(errors.is_empty(), "{errors:?}");
}

#[test]
fn code_that_ras_generated_converts_back_to_handlers() {
    let source = r#"
let count = 5;

@on("btnOK.Click")
function ok() {
    count++;
    this.txtName.text = `${count}`;
}

@on("txtName.Change", { reentrant: true })
function changed() {
    count = 0;
}
"#;
    let (diags, files) = compile(source, &[]);
    assert!(diags.is_empty(), "{diags:?}");
    // With `Erl` line numbers, as `ras pack` writes the form.
    let mut config = Config::default();
    config.forms = vec![FormDecl { name: "frmInput".into(), controls: vec![control("txtName", "TextBox"), control("btnOK", "CommandButton")] }];
    let prog = analyze(vec![form_source(PathBuf::from("forms/frmInput.ras"), "frmInput", source.to_string())], config.clone());
    let out = build(&prog, &BuildOptions::default());
    let vba = out.files.iter().find(|f| f.name == "frmInput").unwrap().code.clone();
    assert!(vba.lines().any(|l| l.starts_with(|c: char| c.is_ascii_digit())), "{vba}");
    let _ = files;
    let converted = ras_compiler::vba2ras::convert_form(&vba, &config.forms[0]);
    println!("{}", converted.text);
    assert!(converted.todos.is_empty(), "{:?}", converted.todos);
    assert!(converted.text.contains("let count: int = 5;"), "{}", converted.text);
    assert!(converted.text.contains("@on(\"btnOK.Click\")\nfunction ok() {"), "{}", converted.text);
    assert!(converted.text.contains("@on(\"txtName.Change\", { reentrant: true })\nfunction changed() {"), "{}", converted.text);
    assert!(!converted.text.contains("RAS_EventError") && !converted.text.contains("ras_busy"), "{}", converted.text);
    let (diags, _) = compile(&converted.text, &[]);
    assert!(diags.is_empty(), "{diags:?}");
}

#[test]
fn editor_completes_controls_and_form_members() {
    let labels = |line: &str| -> Vec<String> {
        let mut config = Config::default();
        config.forms = vec![FormDecl { name: "frmInput".into(), controls: vec![control("txtName", "TextBox")] }];
        let text = format!("@on(\"UserForm.Initialize\")\nfunction init() {{\n    {line}\n}}\n");
        let prog = analyze(vec![form_source(PathBuf::from("forms/frmInput.ras"), "frmInput", text.clone())], config);
        let m = prog.modules.iter().position(|m| m.form.is_some()).unwrap();
        let offset = text.find(line).unwrap() + line.len();
        ras_compiler::query::completions(&prog, m, offset as u32).into_iter().map(|c| c.label).collect()
    };
    let this = labels("this.");
    for want in ["txtName", "caption", "hide", "controls"] {
        assert!(this.contains(&want.to_string()), "{want} missing: {this:?}");
    }
    let ctl = labels("this.txtName.");
    for want in ["text", "setFocus", "visible"] {
        assert!(ctl.contains(&want.to_string()), "{want} missing: {ctl:?}");
    }
}

const COMMON_VBA: &str = r#"Option Explicit

Private Sub UserForm_Activate()
    Dim ws As Worksheet
    Dim r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lstItems.Clear
    For r = 2 To lastRow
        If ws.Cells(r, 1).Value <> "" Then lstItems.AddItem ws.Cells(r, 1).Value
    Next r
    lblCount.Caption = lstItems.ListCount & " 件"
End Sub

Private Sub lstItems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long, picked As String
    For i = 0 To lstItems.ListCount - 1
        If lstItems.Selected(i) Then
            picked = picked & lstItems.List(i) & vbCrLf
        End If
    Next
    If Len(picked) > 0 Then MsgBox picked, vbInformation, "選択"
    Cancel = True
End Sub

Private Sub txtQty_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub chkAll_Click()
    Dim i As Long
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = chkAll.Value
    Next i
End Sub

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Function Clamp(ByVal v As Double, Optional ByVal lo As Double = 0, Optional ByVal hi As Double = 100) As Double
    If v < lo Then
        Clamp = lo
    ElseIf v > hi Then
        Clamp = hi
    Else
        Clamp = v
    End If
End Function
"#;

#[test]
fn common_form_code_converts() {
    let decl = FormDecl {
        name: "frmList".into(),
        controls: vec![control("lstItems", "ListBox"), control("lblCount", "Label"), control("txtQty", "TextBox"), control("chkAll", "CheckBox"), control("btnClose", "CommandButton")],
    };
    let converted = ras_compiler::vba2ras::convert_form(COMMON_VBA, &decl);
    println!("{}", converted.text);
    println!("{:?}", converted.todos);
    let mut config = Config::default();
    config.forms = vec![decl];
    let prog = analyze(vec![form_source(PathBuf::from("forms/frmList.ras"), "frmList", converted.text.clone())], config);
    let errors: Vec<String> = prog.diagnostics.iter().filter(|d| d.is_error()).map(|d| format!("{}:{}: {}", d.code, prog.modules[d.module.unwrap()].line_col_utf16(d.span.start().into()).0 + 1, d.message)).collect();
    assert!(errors.is_empty(), "{errors:?}");
    assert!(converted.text.contains("if (KeyAscii.value >= asc(\"0\") && KeyAscii.value <= asc(\"9\")) {"), "{}", converted.text);
    assert_eq!(converted.todos.len(), 1, "{:?}", converted.todos);
    assert!(converted.todos[0].1.contains("引数を取るプロパティ"));
}

/// The example of spec §14.11.2 compiles.
#[test]
fn spec_example_compiles() {
    let text = include_str!("../../../docs/RAS_Language_Specification.md");
    let start = text.find("// forms/frmInput.ras\n").unwrap();
    let example = &text[start..start + text[start..].find("```").unwrap()];
    let (diags, files) = compile(example, &[("main", "export function run() {\n    frmInput.show();\n    msgBox(frmInput.enteredName());\n}\n")]);
    assert!(diags.is_empty(), "{diags:?}");
    assert!(form_code(&files).contains("Public Function enteredName() As String"));
}
