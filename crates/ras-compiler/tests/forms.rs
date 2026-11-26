//! Form code-behinds (`forms/<Name>.ras`, spec §14.11).

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{
    analyze, build, form_source, BuildOptions, Config, FormControl, FormDecl, ModuleKind,
    SourceInput,
};

fn control(name: &str, ty: &str) -> FormControl {
    FormControl {
        name: name.into(),
        control_type: ty.into(),
    }
}

fn compile(form: &str, src: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmInput".into(),
        controls: vec![
            control("txtName", "TextBox"),
            control("cboKind", "ComboBox"),
            control("btnOK", "CommandButton"),
            control("btnCancel", "CommandButton"),
            control("pages", "MultiPage"),
            control("page1", "Page"),
        ],
    }];
    let mut sources: Vec<SourceInput> = src
        .iter()
        .map(|(key, text)| SourceInput {
            path: PathBuf::from(format!("src/{key}.ras")),
            rel_path: format!("src/{key}.ras"),
            key: key.to_string(),
            kind: ModuleKind::User,
            text: text.to_string(),
        })
        .collect();
    sources.push(form_source(
        PathBuf::from("forms/frmInput.ras"),
        "frmInput",
        form.to_string(),
    ));
    let prog = analyze(sources, config);
    let out = build(
        &prog,
        &BuildOptions {
            erl: Some(false),
            ..Default::default()
        },
    );
    let mut diags: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .filter(|d| {
            d.module
                .is_none_or(|m| prog.modules[m].kind == ModuleKind::User)
        })
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    diags.sort();
    (diags, out.files)
}

fn form_code(files: &[ras_compiler::OutputFile]) -> String {
    files
        .iter()
        .find(|f| f.name == "frmInput")
        .map(|f| f.code.replace("\r\n", "\n"))
        .unwrap_or_default()
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
        &[(
            "main",
            "export function run() {\n    frmInput.show();\n    msgBox(frmInput.selected());\n}\n",
        )],
    );
    assert!(diags.is_empty(), "{diags:?}");
    let code = form_code(&files);
    println!("{code}");
    assert!(code.contains("Private Sub btnOK_Click()"), "{code}");
    assert!(
        code.contains("Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)"),
        "{code}"
    );
    assert!(
        code.contains(
            "Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)"
        ),
        "{code}"
    );
    assert!(code.contains("Private Sub UserForm_Initialize()"), "{code}");
    assert!(code.contains("Me.txtName.Text = "), "{code}");
    assert!(code.contains("Unload Me"), "{code}");
    assert!(
        code.contains("Public Function selected() As String"),
        "{code}"
    );
    assert!(!code.contains("Attribute VB_Name"));
    // "Run in Excel" of a form calls the runtime's helper through Application.Run.
    let runtime = files
        .iter()
        .find(|f| f.name == "RAS_Runtime")
        .expect("runtime module");
    assert!(
        runtime
            .code
            .contains("Public Sub RAS_ShowForm(ByVal name As String)")
            && runtime.code.contains("VBA.UserForms.Add(name).Show"),
        "{}",
        runtime.code
    );
    let main = files.iter().find(|f| f.name == "Main").unwrap();
    println!("{}", main.code);
    assert!(main.code.contains("frmInput.Show"), "{}", main.code);
    assert!(main.code.contains("frmInput.selected"), "{}", main.code);
}

#[test]
fn errors_in_code_behind() {
    let (diags, _) = compile(
        "class Foo {}\nconsole.log(1);\n@macro\nfunction m() {}\nexport let x = 1;\n",
        &[],
    );
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
    let prog = analyze(
        vec![form_source(
            PathBuf::from("forms/frmX.ras"),
            "frmX",
            "function f() {}\n".into(),
        )],
        config,
    );
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
        controls: vec![
            control("txtName", "TextBox"),
            control("cboKind", "ComboBox"),
            control("btnOK", "CommandButton"),
            control("pages", "MultiPage"),
            control("page1", "Page"),
        ],
    };
    let converted = ras_compiler::vba2ras::convert_form(SAMPLE_VBA, &decl);
    println!("{}", converted.text);
    assert!(converted.todos.is_empty(), "{:?}", converted.todos);
    let (diags, files) = compile(&converted.text, &[]);
    assert!(diags.iter().all(|d| !d.starts_with('E')), "{diags:?}");
    let code = form_code(&files);
    println!("{code}");
    for proc in [
        "Private Sub UserForm_Initialize()",
        "Private Sub btnOK_Click()",
        "Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)",
        "Private Sub UserForm_QueryClose(",
        "Public Function Total(",
    ] {
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
        controls: vec![
            control("txtName", "TextBox"),
            control("cboKind", "ComboBox"),
            control("btnOK", "CommandButton"),
            control("chk", "CheckBox"),
        ],
    };
    let converted = ras_compiler::vba2ras::convert_form(HARD_VBA, &decl);
    println!("{}", converted.text);
    println!("{:?}", converted.todos);
    let reasons: Vec<&str> = converted.todos.iter().map(|t| t.1.as_str()).collect();
    for needle in [
        "ユーザー定義型",
        "Like",
        "ReDim",
        "GoTo",
        "行ラベル",
        "Property",
    ] {
        assert!(
            reasons.iter().any(|r| r.contains(needle)),
            "{needle} not reported: {reasons:?}"
        );
    }
    // Every TODO line points at the comment.
    let lines: Vec<&str> = converted.text.lines().collect();
    for (line, _) in &converted.todos {
        assert!(
            lines[line - 1].trim_start().starts_with("// TODO(VBA):"),
            "{}",
            lines[line - 1]
        );
    }
    let mut config = Config::default();
    config.forms = vec![decl];
    let prog = analyze(
        vec![form_source(
            PathBuf::from("forms/frmInput.ras"),
            "frmInput",
            converted.text.clone(),
        )],
        config,
    );
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
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
    config.forms = vec![FormDecl {
        name: "frmInput".into(),
        controls: vec![
            control("txtName", "TextBox"),
            control("btnOK", "CommandButton"),
        ],
    }];
    let prog = analyze(
        vec![form_source(
            PathBuf::from("forms/frmInput.ras"),
            "frmInput",
            source.to_string(),
        )],
        config.clone(),
    );
    let out = build(&prog, &BuildOptions::default());
    let vba = out
        .files
        .iter()
        .find(|f| f.name == "frmInput")
        .unwrap()
        .code
        .clone();
    assert!(
        vba.lines()
            .any(|l| l.starts_with(|c: char| c.is_ascii_digit())),
        "{vba}"
    );
    let _ = files;
    let converted = ras_compiler::vba2ras::convert_form(&vba, &config.forms[0]);
    println!("{}", converted.text);
    assert!(converted.todos.is_empty(), "{:?}", converted.todos);
    assert!(
        converted.text.contains("let count: int = 5;"),
        "{}",
        converted.text
    );
    assert!(
        converted
            .text
            .contains("@on(\"btnOK.Click\")\nfunction ok() {"),
        "{}",
        converted.text
    );
    assert!(
        converted
            .text
            .contains("@on(\"txtName.Change\", { reentrant: true })\nfunction changed() {"),
        "{}",
        converted.text
    );
    assert!(
        !converted.text.contains("RAS_EventError") && !converted.text.contains("ras_busy"),
        "{}",
        converted.text
    );
    let (diags, _) = compile(&converted.text, &[]);
    assert!(diags.is_empty(), "{diags:?}");
}

#[test]
fn editor_completes_controls_and_form_members() {
    let labels = |line: &str| -> Vec<String> {
        let mut config = Config::default();
        config.forms = vec![FormDecl {
            name: "frmInput".into(),
            controls: vec![control("txtName", "TextBox")],
        }];
        let text = format!("@on(\"UserForm.Initialize\")\nfunction init() {{\n    {line}\n}}\n");
        let prog = analyze(
            vec![form_source(
                PathBuf::from("forms/frmInput.ras"),
                "frmInput",
                text.clone(),
            )],
            config,
        );
        let m = prog.modules.iter().position(|m| m.form.is_some()).unwrap();
        let offset = text.find(line).unwrap() + line.len();
        ras_compiler::query::completions(&prog, m, offset as u32)
            .into_iter()
            .map(|c| c.label)
            .collect()
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
        controls: vec![
            control("lstItems", "ListBox"),
            control("lblCount", "Label"),
            control("txtQty", "TextBox"),
            control("chkAll", "CheckBox"),
            control("btnClose", "CommandButton"),
        ],
    };
    let converted = ras_compiler::vba2ras::convert_form(COMMON_VBA, &decl);
    println!("{}", converted.text);
    println!("{:?}", converted.todos);
    let mut config = Config::default();
    config.forms = vec![decl];
    let prog = analyze(
        vec![form_source(
            PathBuf::from("forms/frmList.ras"),
            "frmList",
            converted.text.clone(),
        )],
        config,
    );
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| {
            format!(
                "{}:{}: {}",
                d.code,
                prog.modules[d.module.unwrap()]
                    .line_col_utf16(d.span.start().into())
                    .0
                    + 1,
                d.message
            )
        })
        .collect();
    assert!(errors.is_empty(), "{errors:?}");
    assert!(
        converted
            .text
            .contains("if (KeyAscii.value >= asc(\"0\") && KeyAscii.value <= asc(\"9\")) {"),
        "{}",
        converted.text
    );
    assert_eq!(converted.todos.len(), 1, "{:?}", converted.todos);
    assert!(converted.todos[0].1.contains("引数を取るプロパティ"));
}

/// The example of spec §14.11.2 compiles.
#[test]
fn spec_example_compiles() {
    // A checkout on Windows may have CR LF line ends.
    let text = include_str!("../../../docs/RAS_Language_Specification.md").replace("\r\n", "\n");
    let start = text.find("// forms/frmInput.ras\n").unwrap();
    let example = &text[start..start + text[start..].find("```").unwrap()];
    let (diags, files) = compile(example, &[("main", "export function run() {\n    frmInput.show();\n    msgBox(frmInput.enteredName());\n}\n")]);
    assert!(diags.is_empty(), "{diags:?}");
    assert!(form_code(&files).contains("Public Function enteredName() As String"));
}

/// A variable of the code-behind whose object's events are handled (spec §12.2.2).
#[test]
fn with_events_in_a_code_behind() {
    let (diags, files) = compile(
        "let sheet: Excel.Worksheet;\n\n@on(\"UserForm.Initialize\")\nfunction init() {\n    sheet = activeSheet as Excel.Worksheet;\n}\n\n@on(\"sheet.Change\")\nfunction changed(target: Excel.Range) {\n    this.txtName.text = target.address;\n}\n",
        &[],
    );
    assert!(diags.iter().all(|d| !d.starts_with('E')), "{diags:#?}");
    let code = form_code(&files);
    assert!(
        code.contains("Private WithEvents sheet As Worksheet"),
        "{code}"
    );
    assert!(
        code.contains("Private Sub sheet_Change(ByVal Target As Range)"),
        "{code}"
    );
    assert!(
        code.contains("RAS_EventError \"frmInput.sheet_Change\""),
        "{code}"
    );
}

fn file(files: &[ras_compiler::OutputFile], name: &str) -> String {
    files
        .iter()
        .find(|f| f.name == name)
        .map(|f| f.code.replace("\r\n", "\n"))
        .unwrap_or_else(|| {
            panic!(
                "no module {name} in {:?}",
                files.iter().map(|f| &f.name).collect::<Vec<_>>()
            )
        })
}

/// `control.on("Click", handler)` (spec §12.2.3): a class with a `WithEvents` variable listens
/// to the control and calls the closure; the form takes its listeners off when it is unloaded.
#[test]
fn handlers_attached_at_run_time() {
    let (diags, files) = compile(
        r#"
let count = 0;
const LIMIT = 20;
function bump() { count++; }
function ok() { this.hide(); }
function keyDown(keyCode: MSForms.ReturnInteger, shift: short) { if (keyCode.value == 13) bump(); }

@on("UserForm.Initialize")
function init() {
    this.btnOK.on("Click", () => { bump(); count += 2; this.caption = `n=${count}/${LIMIT}`; });
    this.btnOK.addEventListener("click", ok);
    this.txtName.on("KeyDown", keyDown);
    for (const c of this.controls) {
        if (c instanceof MSForms.TextBox) {
            const l = c.on("Change", () => { c.backColor = c.text == "" ? 255 : 0; });
            if (count > 3) l.off();
        }
    }
    this.txtName.on("MouseMove", (button, shift, x) => { count = x as int; });
    this.txtName.on("MouseMove", (button) => { count = button; });
}
"#,
        &[],
    );
    assert!(diags.is_empty(), "{diags:?}");
    let code = form_code(&files);
    // The form of the code is the owner of its listeners; they go when it is unloaded.
    assert!(
        code.contains("RAS_Listen New RAS_On_CommandButton_Click, Me.btnOK, ras_t1, 0&, Me"),
        "{code}"
    );
    assert!(
        code.contains("Dim l As Object")
            && code.contains("Set l = RAS_Listen(New RAS_On_TextBox_Change, c, ras_t"),
        "{code}"
    );
    assert!(code.contains("l.off"), "{code}");
    assert!(
        code.contains("Private Sub UserForm_Terminate()") && code.contains("RAS_UnlistenOwner Me"),
        "{code}"
    );
    // What the closures use of the form is public: they are other classes.
    assert!(
        code.contains("Public count As Long")
            && code.contains("Public Sub bump()")
            && code.contains("Public Sub ok()")
            && code.contains("Public Sub keyDown("),
        "{code}"
    );
    assert!(code.contains("Private Const LIMIT As Long = 20&"), "{code}");

    let click = file(&files, "RAS_On_CommandButton_Click");
    assert!(
        click.contains("Private WithEvents ras_target As MSForms.CommandButton"),
        "{click}"
    );
    assert!(
        click.contains("Private Sub ras_target_Click()") && click.contains("ras_handler.Apply\n"),
        "{click}"
    );
    assert!(
        click.contains("Public Sub Off()") && click.contains("RAS_Unlisten Me"),
        "{click}"
    );
    assert!(
        click.contains("RAS_EventError \"CommandButton.Click\""),
        "{click}"
    );
    // One class for an event, whatever number of arguments its handlers take.
    let mouse = file(&files, "RAS_On_TextBox_MouseMove");
    assert!(mouse.contains("Private Sub ras_target_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)"), "{mouse}");
    assert!(
        mouse.contains("Case 1&\n        ras_handler.Apply Button\n")
            && mouse.contains("Case 3&\n        ras_handler.Apply Button, Shift, X\n"),
        "{mouse}"
    );

    // The closures reach the form they were created in, not the default instance.
    let lambdas: Vec<String> = files
        .iter()
        .filter(|f| f.name.starts_with("RAS_Lambda_"))
        .map(|f| f.code.replace("\r\n", "\n"))
        .collect();
    let all = lambdas.join("\n");
    assert!(!all.contains("frmInput."), "{all}");
    assert!(
        all.contains("Me.Captured_Me.bump")
            && all.contains("Me.Captured_Me.count = Me.Captured_Me.count + 2&"),
        "{all}"
    );
    assert!(
        all.contains(
            "Me.Captured_Me.Caption = \"n=\" & CStr(Me.Captured_Me.count) & \"/\" & CStr(20&)"
        ),
        "{all}"
    );
    assert!(
        all.contains("Me.Captured_Me.ok\n") && all.contains("Me.Captured_Me.keyDown a1, a2"),
        "{all}"
    );
    assert!(all.contains("Me.Captured_c.BackColor = "), "{all}");

    let runtime = file(&files, "RAS_Runtime");
    for proc in [
        "Function RAS_Listen(",
        "Sub RAS_Unlisten(",
        "Sub RAS_UnlistenOwner(",
        "Private ras_listeners As Collection",
    ] {
        assert!(runtime.contains(proc), "{proc} missing in the runtime");
    }
}

#[test]
fn errors_of_handlers_attached_at_run_time() {
    let (diags, _) = compile(
        r#"
function two(a: int, b: int) {}
@on("UserForm.Initialize")
function init() {
    const event = "Click";
    this.btnOK.on(event, () => {});                    // E0616: not a literal
    this.btnOK.on("Change", () => {});                 // E0611: no such event
    this.txtName.on("Exit", (cancel) => {});           // E0611: raised by the form
    this.on("Click", () => {});                        // E0616: the form itself
    this.controls.on("NoSuchEvent", () => {});         // E0611
    this.controls.on("KeyDown", (c, keyCode, shift, more) => {});   // E0612
    this.btnOK.on("Click", (a) => {});                 // E0612: too many parameters
    this.btnOK.on("Click", two);                       // E0612
    this.btnOK.on("Click", 1);                         // E0301: not a function
}
"#,
        &[],
    );
    let codes: Vec<&str> = diags.iter().map(|d| d.split(':').next().unwrap()).collect();
    assert_eq!(
        codes,
        ["E0301", "E0611", "E0611", "E0611", "E0612", "E0612", "E0612", "E0616", "E0616"],
        "{diags:?}"
    );
    assert!(
        diags
            .iter()
            .any(|d| d.contains("@on(\"<コントロール名>.Exit\")")),
        "{diags:?}"
    );
}

/// An arrow function in a form's code-behind is another class: `this` and the functions and
/// variables of the code-behind are those of the form that created it.
#[test]
fn closures_of_a_form_reach_their_form() {
    let (diags, files) = compile(
        r#"
let count = 0;
function bump() { count++; }
function each(f: () => void) { f(); }
@on("UserForm.Initialize")
function init() {
    const f = () => { bump(); this.caption = `n=${count}`; };
    each(f);
}
"#,
        &[],
    );
    assert!(diags.iter().all(|d| d.starts_with("I0641")), "{diags:?}");
    let lambda = file(&files, "RAS_Lambda_001");
    assert!(lambda.contains("Public Captured_Me As Object"), "{lambda}");
    assert!(
        lambda.contains("Me.Captured_Me.bump")
            && lambda.contains("Me.Captured_Me.Caption = \"n=\" & CStr(Me.Captured_Me.count)"),
        "{lambda}"
    );
    let code = form_code(&files);
    assert!(code.contains("Set ras_t1.Captured_Me = Me"), "{code}");
    // No listeners: nothing to take off.
    assert!(!code.contains("UserForm_Terminate"), "{code}");
}

/// VBA has no public arrays in an object module: a closure, another class, cannot use one.
#[test]
fn closures_cannot_use_the_arrays_of_a_form() {
    let (diags, _) = compile(
        r#"
let names: string[] = [];
const TITLE = "入力";
@on("UserForm.Initialize")
function init() {
    this.btnOK.on("Click", () => { this.caption = `${TITLE} ${names.length}`; });
}
"#,
        &[],
    );
    let codes: Vec<&str> = diags.iter().map(|d| d.split(':').next().unwrap()).collect();
    assert_eq!(codes, ["E0642"], "{diags:?}");
}

/// The example of spec §12.2.3 (handlers attached at run time).
#[test]
fn spec_example_of_on_compiles() {
    let text = include_str!("../../../docs/RAS_Language_Specification.md").replace("\r\n", "\n");
    let start = text.find("// forms/frmEdit.ras\n").unwrap();
    let example = &text[start..start + text[start..].find("```").unwrap()];
    let (diags, files) = compile(example, &[]);
    assert!(diags.is_empty(), "{diags:?}");
    let code = form_code(&files);
    // What the spec shows of the output.
    assert!(code.contains("    Set ras_t1 = New RAS_Lambda_001\n    Set ras_t1.Captured_Me = Me\n    RAS_Listen New RAS_On_CommandButton_Click, Me.btnOK, ras_t1, 0&, Me\n"), "{code}");
    let listener = file(&files, "RAS_On_CommandButton_Click");
    let shown = &text[text.find("  ' RAS_On_CommandButton_Click.cls\n").unwrap()..];
    let shown: String = shown[..shown.find("  ```").unwrap()]
        .lines()
        .skip(1)
        .map(|l| format!("{}\n", l.strip_prefix("  ").unwrap_or(l)))
        .collect();
    assert!(
        listener.ends_with(&shown),
        "the spec shows:\n{shown}\nthe compiler writes:\n{listener}"
    );
}

/// `on` of the collection of the controls, of an array and of a control whose kind is not
/// known (spec §12.2.3): a generated procedure finds the listener class by the type of each.
#[test]
fn handlers_of_collections_and_arrays() {
    let (diags, files) = compile(
        r#"
let edited = 0;
function pressed(c: MSForms.Control, keyCode: MSForms.ReturnInteger) { edited++; }

@on("UserForm.Initialize")
function init() {
    this.controls.on("Change", (c) => { edited++; c.tag = "edited"; });
    this.controls.addEventListener("KeyDown", pressed);
    const boxes = [this.txtName];
    boxes.on("Change", () => { edited++; });
    for (const c of this.controls) {
        const l = c.on("Click", () => { edited++; });
        if (l != null) l.off();
    }
}
"#,
        &[],
    );
    assert!(diags.is_empty(), "{diags:?}");
    let code = form_code(&files);
    assert!(
        code.contains("RAS_OnEach_Change Me.Controls, ras_t1, 256&, Me"),
        "{code}"
    );
    assert!(
        code.contains("RAS_OnEach_KeyDown Me.Controls, ras_t2, 257&, Me"),
        "{code}"
    );
    assert!(
        code.contains("RAS_OnEach_Change boxes, ras_t3, 0&, Me"),
        "{code}"
    );
    assert!(
        code.contains("Set l = RAS_OnAny_Click(c, ras_t4, 0&, Me)"),
        "{code}"
    );
    let events = file(&files, "RAS_Events");
    println!("{events}");
    assert!(events.contains("Option Private Module"), "{events}");
    assert!(events.contains("Public Function RAS_OnAny_Change(ByVal target As Object, ByVal handler As Object, ByVal argc As Long, ByVal owner As Object) As Object\n    Select Case TypeName(target)\n    Case \"CheckBox\"\n        Set RAS_OnAny_Change = RAS_Listen(New RAS_On_CheckBox_Change, target, handler, argc, owner)\n"), "{events}");
    assert!(events.contains("Case \"TextBox\"\n        Set RAS_OnAny_Change = RAS_Listen(New RAS_On_TextBox_Change, target, handler, argc, owner)\n"), "{events}");
    assert!(events.contains("Public Sub RAS_OnEach_Change(ByRef targets As Variant, ByVal handler As Object, ByVal argc As Long, ByVal owner As Object)") && events.contains("    For Each t In targets\n        RAS_OnAny_Change t, handler, argc, owner\n    Next t\n"), "{events}");
    // A control without the event (a label has no `Change`) is left alone.
    assert!(
        !events.contains("RAS_On_Label_Change") && events.contains("RAS_On_Label_Click"),
        "{events}"
    );
    assert!(!events.contains("RAS_OnEach_Click"), "{events}");
    // The handler of a member gets the member, then the arguments of the event.
    let change = file(&files, "RAS_On_TextBox_Change");
    assert!(
        change.contains("Case 0&\n        ras_handler.Apply\n")
            && change.contains("Case 256&\n        ras_handler.Apply ras_target\n"),
        "{change}"
    );
    let key = file(&files, "RAS_On_TextBox_KeyDown");
    assert!(
        key.contains("    ras_handler.Apply ras_target, KeyCode\n"),
        "{key}"
    );
    let lambdas: String = files
        .iter()
        .filter(|f| f.name.starts_with("RAS_Lambda_"))
        .map(|f| f.code.replace("\r\n", "\n"))
        .collect();
    assert!(
        lambdas.contains("Public Function Apply(ByVal c As Variant) As Variant")
            && lambdas.contains("Me.Captured_Me.pressed a1, a2"),
        "{lambdas}"
    );
}

/// `new frmInput(args)` runs the form's `constructor`; `showModal()` returns what
/// `this.close(value)` passed (spec §14.11.9).
#[test]
fn constructor_arguments_and_modal_results() {
    let (diags, files) = compile(
        r#"
let initial = "";

function constructor(name: string, n: int = 3) {
    initial = name;
    this.txtName.text = name;
}

@on("btnOK.Click")
function ok() {
    this.close(this.txtName.text);
}

@on("btnCancel.Click")
function cancel() {
    this.close();
}
"#,
        &[(
            "main",
            r#"
export function run() {
    const s = new frmInput("x").showModal();
    msgBox(s);
    const f: frmInput = new frmInput("y", 2);
    f.showModal();
    msgBox(frmInput.showModal());
}
"#,
        )],
    );
    assert!(diags.is_empty(), "{diags:#?}");
    let code = form_code(&files);
    for want in [
        "Private RAS_Result As String\nPrivate RAS_Modal As Boolean\n",
        "Friend Sub RAS_Construct(ByVal name_ As String, Optional ByVal n As Long = 3&)\n",
        "Private Sub ok()\n    Me.RAS_Close Me.txtName.Text\nEnd Sub\n",
        "Private Sub cancel()\n    Me.RAS_Close vbNullString\nEnd Sub\n",
        "Friend Function RAS_ShowModal() As String\n    RAS_Modal = True\n    Me.Show vbModal\n    RAS_Modal = False\n    RAS_ShowModal = RAS_Result\n    Unload Me\nEnd Function\n",
        "Friend Sub RAS_Close(ByVal ras_value As String)\n    RAS_Result = ras_value\n    If RAS_Modal Then Me.Hide Else Unload Me\nEnd Sub\n",
        // The close button of the title bar hides the form shown by showModal().
        "Private Sub UserForm_QueryClose(ByRef Cancel As Integer, ByRef CloseMode As Integer)\n    On Error GoTo RAS_OnError\n    If RAS_Modal And Cancel = 0 And CloseMode = vbFormControlMenu Then Cancel = 1: Me.Hide\n",
    ] {
        assert!(code.contains(want), "{want}\n---\n{code}");
    }
    let glue = file(&files, "RAS_Glue");
    assert!(
        glue.contains("Public Function New_frmInput(ByVal name_ As String, ByVal n As Long) As frmInput\n    Set New_frmInput = New frmInput\n    New_frmInput.RAS_Construct name_, n\nEnd Function\n"),
        "{glue}"
    );
    let main = file(&files, "Main");
    for want in [
        "s = RAS_Glue.New_frmInput(\"x\", 3&).RAS_ShowModal()",
        "Dim f As frmInput",
        "Set f = RAS_Glue.New_frmInput(\"y\", 2&)",
        "    f.RAS_ShowModal\n",
        "MsgBox frmInput.RAS_ShowModal()",
    ] {
        assert!(main.contains(want), "{want}\n---\n{main}");
    }
}

/// Without a constructor, `new frmInput()` only makes the instance; a form that is never shown by
/// `showModal()` gets none of its procedures; objects come back with `Set`.
#[test]
fn forms_without_constructor_and_object_results() {
    let (diags, files) = compile(
        r#"
@on("btnOK.Click")
function ok() {
    this.close(this.txtName);
}
"#,
        &[(
            "main",
            r#"
export function run() {
    const r = new frmInput().showModal();
    if (r != null) msgBox(r.text);
}
"#,
        )],
    );
    assert!(diags.is_empty(), "{diags:#?}");
    let code = form_code(&files);
    assert!(
        code.contains("Private RAS_Result As MSForms.TextBox\n"),
        "{code}"
    );
    assert!(code.contains("Set RAS_ShowModal = RAS_Result\n"), "{code}");
    assert!(code.contains("Set RAS_Result = ras_value\n"), "{code}");
    let main = file(&files, "Main");
    assert!(
        main.contains("Set r = RAS_Glue.New_frmInput().RAS_ShowModal()"),
        "{main}"
    );
    let glue = file(&files, "RAS_Glue");
    assert!(
        glue.contains("Public Function New_frmInput() As frmInput\n    Set New_frmInput = New frmInput\nEnd Function\n"),
        "{glue}"
    );

    let (diags, files) = compile(
        "@on(\"btnOK.Click\")\nfunction ok() { this.hide(); }\n",
        &[],
    );
    assert!(diags.is_empty(), "{diags:#?}");
    assert!(
        !form_code(&files).contains("RAS_Modal"),
        "{}",
        form_code(&files)
    );
}

#[test]
fn errors_of_constructors_and_results() {
    let (diags, _) = compile(
        r#"
export function constructor(name: string): int { return 1; }

@on("btnOK.Click")
function ok() {
    this.close(1);
    this.close("a");
}
"#,
        &[(
            "main",
            r#"
export function run() {
    new frmInput();
    frmInput.close(1);
}
"#,
        )],
    );
    let joined = diags.join("\n");
    for code in ["E0633", "E0302", "E0402", "E0635"] {
        assert!(joined.contains(code), "{code}: {joined}");
    }

    // `unload(this)` in a form shown by showModal(): the result is not guaranteed.
    let (diags, _) = compile(
        "@on(\"btnOK.Click\")\nfunction ok() { this.close(this.txtName.text); }\n@on(\"btnCancel.Click\")\nfunction cancel() { unload(this); }\n",
        &[("main", "export function run() { msgBox(frmInput.showModal()); }")],
    );
    assert!(diags.iter().any(|d| d.starts_with("W0631")), "{diags:#?}");
    // A form that is only shown by show() may unload itself.
    let (diags, _) = compile(
        "@on(\"btnCancel.Click\")\nfunction cancel() { unload(this); }\n",
        &[("main", "export function run() { frmInput.show(); }")],
    );
    assert!(diags.is_empty(), "{diags:#?}");

    // No constructor: no arguments.
    let (diags, _) = compile(
        "",
        &[("main", "export function run() { new frmInput(1); }")],
    );
    assert!(diags.iter().any(|d| d.starts_with("E0402")), "{diags:#?}");
    // Nothing passed to close(): showModal() has no value.
    let (diags, _) = compile(
        "@on(\"btnOK.Click\")\nfunction ok() { this.close(); }\n",
        &[(
            "main",
            "export function run() { const x = frmInput.showModal(); }",
        )],
    );
    assert!(!diags.is_empty(), "{diags:#?}");
}

/// `ras form pull` of a form RAS wrote without its `.ras`: the generated procedures go back.
#[test]
fn modal_forms_convert_back() {
    let source = r#"
let initial = "";

function constructor(name: string) {
    initial = name;
}

@on("btnOK.Click")
function ok() {
    this.close(`${initial}:${this.txtName.text}`);
}

@on("UserForm.QueryClose")
function closing(byref cancel: short, closeMode: short) {
    if (this.txtName.text == "") cancel = 1;
}
"#;
    let (diags, files) = compile(
        source,
        &[(
            "main",
            "export function run() { msgBox(new frmInput(\"a\").showModal()); }",
        )],
    );
    assert!(diags.is_empty(), "{diags:?}");
    let vba = form_code(&files);
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmInput".into(),
        controls: vec![
            control("txtName", "TextBox"),
            control("btnOK", "CommandButton"),
        ],
    }];
    let converted = ras_compiler::vba2ras::convert_form(&vba, &config.forms[0]);
    println!("{}", converted.text);
    assert!(converted.todos.is_empty(), "{:?}", converted.todos);
    for want in [
        "function constructor(name_: string) {",
        "    this.close(`${initial}:${this.txtName.text}`);",
        "@on(\"UserForm.QueryClose\")\nfunction closing(",
    ] {
        assert!(
            converted.text.contains(want),
            "{want}\n---\n{}",
            converted.text
        );
    }
    assert!(!converted.text.contains("RAS_"), "{}", converted.text);
    let (diags, _) = compile(
        &converted.text,
        &[(
            "main",
            "export function run() { msgBox(new frmInput(\"a\").showModal()); }",
        )],
    );
    assert!(diags.is_empty(), "{diags:?}");
}

/// A `FormValidator` in a module variable of a form (spec §14.11.12): the `Exit` and `Change` of
/// the form's input controls reach its `commit` / `edit`, after the handlers of the code-behind;
/// the `Exit` of a container reaches `commitActive`.
#[test]
fn validators_receive_the_events_of_input_controls() {
    let (diags, files) = compile(
        r#"
import * as z from "validation";

const form = new z.FormValidator();

function constructor() {
    form.rule(this.txtName, z.string().required("名前を入力してください").max(20));
    form.rule(this.cboKind, z.string().oneOf(["a", "b"]));
}

@on("txtName.Change")
function changed() {
    this.caption = this.txtName.text;
}

@on("btnOK.Click")
function ok() {
    if (!form.validateAll()) return;
    this.close(this.txtName.text);
}
"#,
        &[],
    );
    assert!(diags.is_empty(), "{diags:#?}");
    let code = form_code(&files);
    for want in [
        "Private form As FormValidator\n",
        "Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)\n    On Error GoTo RAS_OnError\n    If Not form Is Nothing Then form.commit Me.txtName\n    Exit Sub\n",
        "        changed\n        ras_busy1 = False\n    End If\n    If Not form Is Nothing Then form.edit Me.txtName\n",
        "Private Sub cboKind_Exit(ByVal Cancel As MSForms.ReturnBoolean)\n",
        "Private Sub cboKind_Change()\n",
        "Private Sub pages_Exit(ByVal Cancel As MSForms.ReturnBoolean)\n    On Error GoTo RAS_OnError\n    If Not form Is Nothing Then form.commitActive Me.pages\n",
    ] {
        assert!(code.contains(want), "{want}\n---\n{code}");
    }
    // Buttons have nothing to validate.
    assert!(!code.contains("btnOK_Exit"), "{code}");
    let lib = file(&files, "FormValidator");
    assert!(
        lib.contains("Public Sub commit(ByVal control As MSForms.Control)"),
        "{lib}"
    );
    assert!(lib.contains("Public Sub commitActive("), "{lib}");
}

/// A form without a `FormValidator` gets no event procedures for its controls.
#[test]
fn forms_without_validator_get_no_validation_events() {
    let (diags, files) = compile(
        r#"
@on("btnOK.Click")
function ok() {
    this.hide();
}
"#,
        &[],
    );
    assert!(diags.is_empty(), "{diags:#?}");
    let code = form_code(&files);
    assert!(!code.contains("txtName_Exit"), "{code}");
    assert!(!files.iter().any(|f| f.name == "FormValidator"));
}

/// The example of spec §14.11.12 compiles, and its form gets the event procedure shown there.
#[test]
fn spec_example_of_validation_compiles() {
    let text = include_str!("../../../docs/RAS_Language_Specification.md").replace("\r\n", "\n");
    let section = &text[text.find("#### 14.11.12").unwrap()..];
    let start = section.find("// forms/frmOrder.ras\n").unwrap();
    let example = &section[start..start + section[start..].find("```").unwrap()];
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmOrder".into(),
        controls: vec![
            control("txtName", "TextBox"),
            control("txtAge", "TextBox"),
            control("txtMail", "TextBox"),
            control("chkAgree", "CheckBox"),
            control("cboKind", "ComboBox"),
            control("fraCorp", "Frame"),
            control("txtCorp", "TextBox"),
            control("btnOK", "CommandButton"),
        ],
    }];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: "export function run() {\n    msgBox(new frmOrder().showModal());\n}\n".into(),
        },
        form_source(
            PathBuf::from("forms/frmOrder.ras"),
            "frmOrder",
            example.to_string(),
        ),
    ];
    let prog = analyze(sources, config);
    let out = build(
        &prog,
        &BuildOptions {
            erl: Some(false),
            ..Default::default()
        },
    );
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "{errors:#?}");
    let code = file(&out.files, "frmOrder");
    let shown = &section[section.find("  ```vb\n").unwrap() + 8..];
    let shown = &shown[..shown.find("  ```").unwrap()];
    let shown: String = shown
        .lines()
        .map(|l| l.strip_prefix("  ").unwrap_or(l))
        .collect::<Vec<_>>()
        .join("\n");
    let (decl, proc) = shown.split_once("\n\n").unwrap();
    assert!(code.contains(decl), "{decl}\n---\n{code}");
    assert!(code.contains(proc.trim_end()), "{proc}\n---\n{code}");
    // The handler of cboKind.Change runs before the validator hears of the change.
    assert!(
        code.contains("        kindChanged\n        ras_busy1 = False\n    End If\n    If Not form Is Nothing Then form.edit Me.cboKind\n"),
        "{code}"
    );
    assert!(
        code.contains("If Not form Is Nothing Then form.commitActive Me.fraCorp"),
        "{code}"
    );
    for module in [
        "RAS_Validation",
        "FormValidator",
        "ValidationSchema",
        "RAS_ValidationField",
        "RAS_ValidationCheck",
    ] {
        assert!(out.files.iter().any(|f| f.name == module), "{module}");
    }
}

/// A `ProgressBar` of a form (spec §14.11.13) is an object of the bundled class, a public member
/// of the form attached to its labels in `UserForm_Initialize`, and operated like a control.
#[test]
fn progress_bars_are_objects_of_the_bundled_class() {
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmUpload".into(),
        controls: vec![
            control("pbUpload", "ProgressBar"),
            control("btnStart", "CommandButton"),
        ],
    }];
    let form = r#"
@on("btnStart.Click")
function start() {
    this.pbUpload.caption = "Upload progress";
    for (let i = 0; i <= 100; i++) {
        this.pbUpload.value = i;
    }
    this.pbUpload.barColor = rgb(0, 0, 0);
    this.pbUpload.showValue = this.pbUpload.value < 100;
}
"#;
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: "export function run() {\n    frmUpload.pbUpload.value = 10;\n    frmUpload.show();\n}\n".into(),
        },
        form_source(PathBuf::from("forms/frmUpload.ras"), "frmUpload", form.into()),
    ];
    let prog = analyze(sources, config);
    let out = build(
        &prog,
        &BuildOptions {
            erl: Some(false),
            ..Default::default()
        },
    );
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .map(|d| {
            format!(
                "{}: {} ({:?})",
                d.code,
                d.message,
                d.module.map(|m| &prog.modules[m].rel_path)
            )
        })
        .collect();
    assert!(errors.is_empty(), "{errors:#?}");
    let code = file(&out.files, "frmUpload");
    for want in [
        "Public pbUpload As ProgressBar\n",
        "Private Sub UserForm_Initialize()\n    On Error GoTo RAS_OnError\n    Set pbUpload = New ProgressBar\n    pbUpload.RAS_Attach Me, \"pbUpload\"\n",
        "Me.pbUpload.value = i",
    ] {
        assert!(code.contains(want), "{want}\n---\n{code}");
    }
    let class = file(&out.files, "ProgressBar");
    assert!(class.contains("Public Sub RAS_Attach("), "{class}");
    assert!(class.contains("Public Property Let value("), "{class}");
    assert!(out.files.iter().any(|f| f.name == "RAS_ProgressBar"));
    let main = file(&out.files, "Main");
    assert!(main.contains("frmUpload.pbUpload.value = 10"), "{main}");
}

/// A project without progress bars gets no class for them.
#[test]
fn no_progress_bar_no_class() {
    let (_, files) = compile("", &[]);
    assert!(!files
        .iter()
        .any(|f| f.name == "ProgressBar" || f.name == "RAS_ProgressBar"));
}

/// A form with a progress bar and no code-behind still gets the code that attaches the bar.
#[test]
fn a_progress_bar_needs_no_code_behind() {
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmWait".into(),
        controls: vec![control("pbWait", "ProgressBar")],
    }];
    let sources = vec![SourceInput {
        path: PathBuf::from("src/main.ras"),
        rel_path: "src/main.ras".into(),
        key: "main".into(),
        kind: ModuleKind::User,
        text: "export function run() {\n    frmWait.pbWait.value = 30;\n    frmWait.show(vbModeless);\n}\n".into(),
    }];
    let prog = analyze(sources, config);
    let out = build(&prog, &BuildOptions::default());
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "{errors:#?}");
    let code = file(&out.files, "frmWait");
    assert!(code.contains("pbWait.RAS_Attach Me, \"pbWait\""), "{code}");
}

/// A `FilePicker` of a form (spec §14.11.14): an object of the bundled class; the click of its
/// button and the exit of its box reach the object, and `@on("<name>.<event>")` receives the
/// events of its box.
#[test]
fn file_pickers_browse_with_their_button() {
    let mut config = Config::default();
    config.forms = vec![FormDecl {
        name: "frmOpen".into(),
        controls: vec![
            control("fpInput", "FilePicker"),
            control("btnOK", "CommandButton"),
        ],
    }];
    let form = r#"
@on("fpInput.Change")
function changed() {
    this.btnOK.enabled = this.fpInput.path != "";
}

@on("btnOK.Click")
function ok() {
    const files: string[] = this.fpInput.paths;
    this.close(files.length);
}

export function pick(): string[] {
    this.fpInput.dialogType = 4;
    this.fpInput.addFilter("Excel", "*.xlsx");
    return this.fpInput.browse();
}
"#;
    let sources = vec![form_source(
        PathBuf::from("forms/frmOpen.ras"),
        "frmOpen",
        form.into(),
    )];
    let prog = analyze(sources, config);
    let out = build(
        &prog,
        &BuildOptions {
            erl: Some(false),
            ..Default::default()
        },
    );
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .chain(out.diagnostics.iter())
        .map(|d| {
            format!(
                "{}: {} ({:?})",
                d.code,
                d.message,
                d.module.map(|m| &prog.modules[m].rel_path)
            )
        })
        .collect();
    assert!(errors.is_empty(), "{errors:#?}");
    let code = file(&out.files, "frmOpen");
    for want in [
        "Public fpInput As FilePicker\n",
        "    Set fpInput = New FilePicker\n    fpInput.RAS_Attach Me, \"fpInput\"\n",
        "Private Sub fpInput_Text_Change()\n",
        "Private Sub fpInput_Button_Click()\n    On Error GoTo RAS_OnError\n    If Not fpInput Is Nothing Then fpInput.RAS_Click\n",
        "Private Sub fpInput_Text_Exit(ByVal Cancel As MSForms.ReturnBoolean)\n    On Error GoTo RAS_OnError\n    If Not fpInput Is Nothing Then fpInput.RAS_Exit\n",
    ] {
        assert!(code.contains(want), "{want}\n---\n{code}");
    }
    let class = file(&out.files, "FilePicker");
    assert!(
        class.contains("Public Function browse() As String()"),
        "{class}"
    );
    assert!(class.contains("Application.FileDialog("), "{class}");
}
