//! The files of sheets and of `ThisWorkbook` (`objects/<CodeName>.ras`, spec §12.2).

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{analyze, build, build_js, object_source, BuildOptions, Config, FormControl, FormDecl, ModuleKind, SourceInput};

fn config() -> Config {
    let mut config = Config::default();
    config.sheet_code_names = vec!["Sheet1".into(), "Sheet2".into()];
    config.sheet_controls_known = true;
    config.sheet_controls = vec![FormDecl {
        name: "Sheet1".into(),
        controls: vec![FormControl { name: "CommandButton1".into(), control_type: "CommandButton".into() }, FormControl { name: "TextBox1".into(), control_type: "TextBox".into() }],
    }];
    config
}

fn sources(objects: &[(&str, &str)], src: &[(&str, &str)]) -> Vec<SourceInput> {
    let mut sources: Vec<SourceInput> = src
        .iter()
        .map(|(key, text)| SourceInput { path: PathBuf::from(format!("src/{key}.ras")), rel_path: format!("src/{key}.ras"), key: key.to_string(), kind: ModuleKind::User, text: text.to_string() })
        .collect();
    for (name, text) in objects {
        sources.push(object_source(PathBuf::from(format!("objects/{name}.ras")), name, text.to_string()));
    }
    sources
}

fn compile(objects: &[(&str, &str)], src: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let prog = analyze(sources(objects, src), config());
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    let mut diags: Vec<String> = prog.diagnostics.iter().chain(out.diagnostics.iter()).filter(|d| d.module.is_none_or(|m| prog.modules[m].kind == ModuleKind::User)).map(|d| format!("{}: {}", d.code, d.message)).collect();
    diags.sort();
    diags.dedup();
    (diags, out.files)
}

fn codes(diags: &[String]) -> Vec<&str> {
    diags.iter().map(|d| d.split(':').next().unwrap_or_default()).collect()
}

fn module(files: &[ras_compiler::OutputFile], name: &str) -> String {
    files.iter().find(|f| f.name == name).map(|f| f.code.replace("\r\n", "\n")).unwrap_or_default()
}

const SHEET1: &str = r#"
let lastEdited = "";

@on("Worksheet.Change")
function onChange(target: Excel.Range) {
    lastEdited = target.address;
    this.range("D1").value = lastEdited;
    helper();
}

@on("Worksheet.BeforeDoubleClick")
function onDoubleClick(target: Excel.Range, byref cancel: boolean) {
    cancel = true;
}

function helper() {
    this.cells(1, 1).value = 1;
}

/** 入力欄を消す。 */
export function clearInput() {
    this.range("A2:C100").clearContents();
}

export function total(n: int): int {
    return n * 2;
}
"#;

#[test]
fn a_sheets_file_is_its_whole_module() {
    let (diags, files) = compile(&[("Sheet1", SHEET1)], &[("main", "export function run() {\n    Sheet1.clearInput();\n    debug.print(Sheet1.total(2));\n    Sheet1.range(\"A1\").value = 1;\n}\n")]);
    assert!(diags.is_empty(), "{diags:#?}");
    let file = files.iter().find(|f| f.name == "Sheet1").expect("Sheet1");
    assert_eq!(file.file_name, "Sheet1.doccls");
    assert_eq!(file.kind, ras_compiler::vba::VModuleKind::Document);
    let code = module(&files, "Sheet1");
    // Generated whole: the hash line and the options of every module, no `Option Private Module`.
    let head: Vec<&str> = code.lines().take(3).collect();
    assert!(head[0].starts_with("' RAS-GENERATED hash="), "{code}");
    assert_eq!(&head[1..], ["Option Explicit", "Option Compare Binary"], "{code}");
    assert!(!code.contains("Option Private Module"), "{code}");
    // Module variables, handlers called from the event procedures of the same module.
    assert!(code.contains("Private lastEdited As String"), "{code}");
    assert!(code.contains("Private Sub Worksheet_Change(ByVal Target As Range)"), "{code}");
    assert!(code.contains("Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)"), "{code}");
    assert!(code.contains("onChange Target"), "{code}");
    assert!(code.contains("onDoubleClick Target, Cancel"), "{code}");
    assert!(code.contains("Private Sub onChange(ByVal target As Range)"), "{code}");
    assert!(code.contains("RAS_EventError \"Sheet1.Change\""), "{code}");
    // `this` is the sheet; functions of the file are called without a qualifier.
    assert!(code.contains("Me.Range(\"D1\").Value"), "{code}");
    assert!(code.contains("Me.Cells(1&, 1&).Value"), "{code}");
    assert!(!code.contains("Sheet1.helper"), "{code}");
    // Exports are public members of the sheet.
    assert!(code.contains("Public Sub clearInput()"), "{code}");
    assert!(code.contains("Public Function total(ByVal n As Long) As Long"), "{code}");
    assert!(code.contains("Private Sub helper()"), "{code}");
    let main = module(&files, "Main");
    assert!(main.contains("Sheet1.clearInput"), "{main}");
    assert!(main.contains("Sheet1.total(2&)"), "{main}");
    // No standard module for the sheet.
    assert!(!files.iter().any(|f| f.name == "Sheet1Events"), "{:?}", files.iter().map(|f| &f.file_name).collect::<Vec<_>>());
}

#[test]
fn the_workbooks_file_and_the_initialisation() {
    let book = r#"
@on("Workbook.Open")
function setup() {
    debug.print(this.name);
}

@on("Workbook.SheetChange")
function onAnySheetChange(sh: object, target: Excel.Range) {
    debug.print(target.address);
}

export function refresh() {
    this.save();
}
"#;
    let (diags, files) = compile(&[("ThisWorkbook", book)], &[("main", "export function run() {\n    thisWorkbook.refresh();\n}\n")]);
    assert!(diags.is_empty(), "{diags:#?}");
    let code = module(&files, "ThisWorkbook");
    assert!(code.contains("Private Sub Workbook_Open()"), "{code}");
    assert!(code.contains("RAS_WorkbookOpen"), "{code}");
    // The handlers of `Open` are called back once the workbook is ready (spec §12.3).
    assert!(code.contains("Public Sub RAS_WorkbookOpened(Optional ByVal ras_dummy As Long = 0&)"), "{code}");
    assert!(code.contains("Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)"), "{code}");
    assert!(code.contains("Private Sub Workbook_Activate()") && code.contains("Private Sub Workbook_WindowActivate(ByVal Wn As Window)"), "{code}");
    assert!(code.contains("Debug.Print Me.Name"), "{code}");
    assert!(code.contains("Public Sub refresh()"), "{code}");
    let glue = module(&files, "RAS_Glue");
    assert!(glue.contains("ThisWorkbook.RAS_WorkbookOpened"), "{glue}");
    assert!(module(&files, "Main").contains("ThisWorkbook.refresh"), "{}", module(&files, "Main"));
}

#[test]
fn this_workbook_gets_the_initialisation_without_a_file() {
    let (diags, files) = compile(&[], &[("main", "let n = compute();\nfunction compute(): int { return 1; }\nexport function run() {\n    debug.print(n);\n}\n")]);
    assert!(diags.is_empty(), "{diags:#?}");
    let code = module(&files, "ThisWorkbook");
    assert!(code.contains("Private Sub Workbook_Open()") && code.contains("RAS_WorkbookOpen"), "{code}");
    assert!(!code.contains("RAS_WorkbookOpened"), "{code}");
}

#[test]
fn module_variables_are_initialised_with_the_others() {
    let (diags, files) = compile(&[("Sheet1", "let labels: string[] = [\"a\", \"b\"];\nexport function labelCount(): int { return labels.length; }\n")], &[]);
    assert!(diags.is_empty(), "{diags:#?}");
    let glue = module(&files, "RAS_Glue");
    assert!(glue.contains("Sheet1.RAS_ModInit"), "{glue}");
}

#[test]
fn controls_on_the_sheet() {
    let sheet = r#"
@on("CommandButton1.Click")
function run() {
    this.CommandButton1.caption = "実行中";
    this.TextBox1.text = "";
}

@on("TextBox1.KeyDown")
function keyDown(keyCode: MSForms.ReturnInteger, shift: short) {
    if (keyCode.value == 13) this.range("B2").value = shift;
}

@on("TextBox1.LostFocus")
function left() {
}
"#;
    let (diags, files) = compile(&[("Sheet1", sheet)], &[("main", "export function run() {\n    Sheet1.CommandButton1.caption = \"x\";\n}\n")]);
    assert!(diags.is_empty(), "{diags:#?}");
    let code = module(&files, "Sheet1");
    assert!(code.contains("Private Sub CommandButton1_Click()"), "{code}");
    assert!(code.contains("Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)"), "{code}");
    assert!(code.contains("Private Sub TextBox1_LostFocus()"), "{code}");
    assert!(code.contains("Me.CommandButton1.Caption = "), "{code}");
    assert!(code.contains("RAS_EventError \"Sheet1.CommandButton1_Click\""), "{code}");
    assert!(module(&files, "Main").contains("Sheet1.CommandButton1.Caption = \"x\""), "{}", module(&files, "Main"));

    // What a form raises is not an event of a control on a sheet; unknown controls and members.
    let (diags, _) = compile(&[("Sheet1", "@on(\"TextBox1.Exit\")\nfunction a() {}\n@on(\"Button9.Click\")\nfunction b() {}\n")], &[]);
    assert_eq!(codes(&diags), ["E0611", "E0613"], "{diags:#?}");
    let (diags, _) = compile(&[("Sheet1", "function c() { this.Nothing1.value = 1; }\n")], &[]);
    assert_eq!(codes(&diags), ["E0202"], "{diags:#?}");
}

#[test]
fn controls_whose_events_are_not_known() {
    let sheet = "@on(\"Calendar1.DateClick\")\nfunction picked(dateClicked: date, byref cancel: boolean) {\n    cancel = true;\n}\n\n@on(\"Calendar1.DateClick\")\nfunction also() {\n}\n";
    // A control that is not an MSForms one: the event procedure is declared as the handler is.
    let mut cfg = config();
    cfg.sheet_controls[0].controls.push(FormControl { name: "Calendar1".into(), control_type: String::new() });
    let prog = analyze(sources(&[("Sheet1", sheet)], &[]), cfg);
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    assert!(prog.diagnostics.iter().chain(out.diagnostics.iter()).all(|d| !d.is_error()), "{:#?}", prog.diagnostics.iter().chain(out.diagnostics.iter()).map(|d| &d.message).collect::<Vec<_>>());
    let code = module(&out.files, "Sheet1");
    assert!(code.contains("Private Sub Calendar1_DateClick(ByVal dateClicked As Date, ByRef cancel As Boolean)"), "{code}");
    assert!(code.contains("picked dateClicked, cancel") && code.contains("        also\n"), "{code}");
    // Without a workbook the controls are not known: any name is taken as a control.
    let prog = analyze(sources(&[("Sheet1", sheet)], &[]), Config::default());
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    assert!(prog.diagnostics.iter().chain(out.diagnostics.iter()).all(|d| !d.is_error()), "{:#?}", prog.diagnostics.iter().chain(out.diagnostics.iter()).map(|d| &d.message).collect::<Vec<_>>());
    assert!(module(&out.files, "Sheet1").contains("Private Sub Calendar1_DateClick(ByVal dateClicked As Date, ByRef cancel As Boolean)"));
}

#[test]
fn what_a_document_module_cannot_hold() {
    let sheet = r#"
export let shared = 1;
export const LIMIT = 10;
export enum Kind { A, B }
struct Point { x: int; y: int; }
class Helper {}
debug.print("top level");
@udf
export function twice(n: int): int { return n * 2; }
@init
function start() {}
function name(): string { return ""; }
let range = 1;
"#;
    let (diags, _) = compile(&[("Sheet1", sheet)], &[]);
    let got = codes(&diags);
    assert_eq!(got.iter().filter(|c| **c == "E0650").count(), 7, "{diags:#?}");
    assert_eq!(got.iter().filter(|c| **c == "E0651").count(), 2, "{diags:#?}");
    assert!(got.iter().all(|c| *c == "E0650" || *c == "E0651"), "{diags:#?}");
}

#[test]
fn where_on_can_be_written() {
    // In a standard module, and the form of v0.9 that names the sheet.
    let (diags, _) = compile(&[], &[("main", "@on(\"Sheet1.Change\")\nfunction a(target: Excel.Range) {}\n@on(\"Workbook.Open\")\nfunction b() {}\n")]);
    assert_eq!(codes(&diags), ["E0610", "E0610"], "{diags:#?}");
    assert!(diags[0].contains("objects/Sheet1.ras") && diags[0].contains("Worksheet.Change"), "{diags:#?}");
    // The workbook's events in a sheet's file and the other way round.
    let (diags, _) = compile(&[("Sheet1", "@on(\"Workbook.Open\")\nfunction a() {}\n"), ("ThisWorkbook", "@on(\"Worksheet.Change\")\nfunction b() {}\n")], &[]);
    assert_eq!(codes(&diags), ["E0610", "E0610"], "{diags:#?}");
    // Unknown events, too many parameters, a function named like the event procedure.
    let (diags, _) = compile(&[("Sheet1", "@on(\"Worksheet.Open\")\nfunction a() {}\n@on(\"Worksheet.Change\")\nfunction b(t: Excel.Range, more: int) {}\n@on(\"Worksheet.Activate\")\nfunction c() {}\nfunction Worksheet_Activate() {}\n")], &[]);
    assert_eq!(codes(&diags), ["E0611", "E0612", "E0614"], "{diags:#?}");
}

#[test]
fn files_and_the_sheets_of_the_workbook() {
    // No sheet of that code name.
    let (diags, _) = compile(&[("Sheet9", "export function a() {}\n")], &[]);
    assert_eq!(codes(&diags), ["E0610"], "{diags:#?}");
    assert!(diags[0].contains("Sheet1, Sheet2"), "{diags:#?}");
    // A standard module named like a sheet or the workbook.
    let (diags, _) = compile(&[], &[("Sheet1", "export function a() {}\n"), ("thisWorkbook", "export function b() {}\n")]);
    assert_eq!(codes(&diags), ["E0102", "E0102"], "{diags:#?}");
    assert!(diags.iter().any(|d| d.contains("objects/Sheet1.ras")) && diags.iter().any(|d| d.contains("objects/ThisWorkbook.ras")), "{diags:#?}");
    // The file name in any case; the module has the code name of the workbook.
    let (diags, files) = compile(&[("sheet1", "export function a() {}\n")], &[]);
    assert!(diags.is_empty(), "{diags:#?}");
    assert!(files.iter().any(|f| f.name == "Sheet1"), "{:?}", files.iter().map(|f| &f.name).collect::<Vec<_>>());
    // Without a workbook, the file declares the sheet.
    let prog = analyze(sources(&[("Data", "export function a() {}\n")], &[("main", "export function run() {\n    Data.a();\n    Data.range(\"A1\").value = 1;\n}\n")]), Config::default());
    assert!(prog.diagnostics.iter().all(|d| !d.is_error()), "{:#?}", prog.diagnostics.iter().map(|d| &d.message).collect::<Vec<_>>());
}

#[test]
fn exported_functions_are_macros() {
    let prog = analyze(sources(&[("Sheet1", SHEET1)], &[]), config());
    let f = prog.fns.iter().find(|f| f.name == "clearInput").unwrap();
    assert!(f.is_macro_entry());
    let out = build(&prog, &BuildOptions { erl: Some(false), ..Default::default() });
    let code = module(&out.files, "Sheet1");
    // The wrapping of a macro entry (spec §12.4).
    let body = &code[code.find("Public Sub clearInput()").unwrap()..];
    assert!(body.contains("RAS_EnsureInit"), "{code}");
}

#[test]
fn debug_build() {
    let prog = analyze(sources(&[("Sheet1", SHEET1), ("ThisWorkbook", "@on(\"Workbook.Open\")\nfunction setup() {}\n")], &[("main", "export function run() {\n    Sheet1.clearInput();\n}\n")]), config());
    let out = build_js(&prog, "");
    assert!(out.diagnostics.iter().all(|d| !d.is_error()), "{:#?}", out.diagnostics.iter().map(|d| &d.message).collect::<Vec<_>>());
    let sheet = &out.files.iter().find(|f| f.rel_path == "@objects/Sheet1.js").expect("@objects/Sheet1.js").text;
    assert!(sheet.contains("$rt.module(\"Sheet1\""), "{sheet}");
    assert!(sheet.contains("{ target: \"Sheet1\", event: \"Change\", reentrant: false, fn: ($a) => onChange($a[0].v) }"), "{sheet}");
    assert!(sheet.contains("{ target: \"Sheet1\", event: \"BeforeDoubleClick\", reentrant: false, fn: ($a) => onDoubleClick($a[0].v, $a[1]) }"), "{sheet}");
    assert!(sheet.contains("$rt.global(\"Sheet1\")"), "{sheet}");
    let book = &out.files.iter().find(|f| f.rel_path == "@objects/ThisWorkbook.js").expect("@objects/ThisWorkbook.js").text;
    assert!(book.contains("{ target: \"Workbook\", event: \"Open\""), "{book}");
    let main = &out.files.iter().find(|f| f.rel_path == "main.js").unwrap().text;
    assert!(main.contains(".clearInput()"), "{main}");
    assert!(!main.contains("$rt.global(\"Sheet1\").clearInput"), "{main}");
}

const BOOK_WITH_APP: &str = r#"
let app: Excel.Application;

@on("Workbook.Open")
function start() {
    app = application;
}

@on("app.SheetChange")
function onAnyChange(sh: object, target: Excel.Range) {
    debug.print(target.address);
}

@on("app.WorkbookBeforeClose")
function onClose(wb: Excel.Workbook, byref cancel: boolean) {
    cancel = wb.name == "x";
}
"#;

#[test]
fn with_events_variables() {
    let (diags, files) = compile(&[("ThisWorkbook", BOOK_WITH_APP)], &[]);
    assert!(diags.is_empty(), "{diags:#?}");
    let code = module(&files, "ThisWorkbook");
    assert!(code.contains("Private WithEvents app As Application"), "{code}");
    assert!(code.contains("Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range)"), "{code}");
    assert!(code.contains("Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, ByRef Cancel As Boolean)"), "{code}");
    assert!(code.contains("onAnyChange Sh, Target") && code.contains("onClose Wb, Cancel"), "{code}");
    assert!(code.contains("RAS_EventError \"ThisWorkbook.app_SheetChange\""), "{code}");
    assert!(code.contains("Set app = Application"), "{code}");
    let prog = analyze(sources(&[("ThisWorkbook", BOOK_WITH_APP)], &[]), config());
    let out = build_js(&prog, "");
    assert!(out.diagnostics.iter().all(|d| !d.is_error()), "{:#?}", out.diagnostics.iter().map(|d| &d.message).collect::<Vec<_>>());
    let js = &out.files.iter().find(|f| f.rel_path == "@objects/ThisWorkbook.js").unwrap().text;
    // The handlers of the variable are registered with the object it holds (spec §12.2.2).
    assert!(js.contains("event: \"Open\""), "{js}");
    assert!(js.contains("$rt.withEvents([{ target: \"app\", event: \"SheetChange\", reentrant: false, source: () => app, fn: ($a) => onAnyChange($a[0].v, $a[1].v) }, { target: \"app\", event: \"WorkbookBeforeClose\", reentrant: false, source: () => app, fn: ($a) => onClose($a[0].v, $a[1]) }]);"), "{js}");
}

#[test]
fn with_events_fields_of_a_class() {
    let watcher = r#"
export class ButtonWatcher {
    private button: MSForms.CommandButton;
    private clicks = 0;

    constructor(button: MSForms.CommandButton) {
        this.button = button;
    }

    @on("button.Click")
    private clicked() {
        this.clicks++;
    }

    @on("button.KeyDown")
    private keyDown(keyCode: MSForms.ReturnInteger) {
        keyCode.value = 0;
    }
}

export function watch(b: MSForms.CommandButton): ButtonWatcher {
    return new ButtonWatcher(b);
}
"#;
    let (diags, files) = compile(&[], &[("watcher", watcher)]);
    assert!(diags.is_empty(), "{diags:#?}");
    let code = module(&files, "ButtonWatcher");
    assert!(code.contains("Private WithEvents button As MSForms.CommandButton"), "{code}");
    assert!(code.contains("Private Sub button_Click()"), "{code}");
    assert!(code.contains("Private Sub button_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)"), "{code}");
    assert!(code.contains("        clicked\n") && code.contains("keyDown KeyCode"), "{code}");
}

#[test]
fn what_with_events_cannot_be() {
    // A standard module cannot listen.
    let (diags, _) = compile(&[], &[("main", "let app: Excel.Application;\n@on(\"app.SheetChange\")\nfunction a(sh: object) {}\n")]);
    assert_eq!(codes(&diags), ["E0610"], "{diags:#?}");
    assert!(diags[0].contains("WithEvents"), "{diags:#?}");
    // The type of the variable; an event its class does not have; a form's event on a variable.
    let sheet = r#"
let anything: object;
let list: Excel.Range[] = [];
let book: Excel.Workbook;
let box1: MSForms.TextBox;

@on("anything.Click")
function a() {}

@on("list.Change")
function b() {}

@on("book.Bogus")
function c() {}

@on("box1.Exit")
function d() {}

@on("book.Open")
function e(more: int) {}
"#;
    let (diags, _) = compile(&[("Sheet1", sheet)], &[]);
    assert_eq!(codes(&diags), ["E0611", "E0611", "E0612", "E0615", "E0615"], "{diags:#?}");
}

/// `on` / `addEventListener` (spec §12.2.3): a control on a sheet, the sheet, the workbook and
/// `Application` are listened to by generated classes; a shape gets the macro's name.
#[test]
fn handlers_attached_at_run_time() {
    let sheet = r#"
let clicks = 0;

@on("Worksheet.Activate")
function activated() {
    this.CommandButton1.on("Click", () => { clicks++; this.range("A1").value = clicks; });
    this.TextBox1.addEventListener("Change", refresh);
    this.on("BeforeDoubleClick", (target) => target.row > 1);
    this.shapes("Button 1").on("Click", refresh);
}

export function refresh() {
    this.calculate();
}
"#;
    let main_src = r#"
export function reload() {}
function quiet(): boolean { return false; }

export function setup(ws: Excel.Worksheet) {
    const shp = ws.shapes.addShape(1, 10, 10, 80, 24);
    shp.on("click", reload);
    shp.addEventListener("Click", Sheet1.refresh);
    application.on("SheetChange", (sh, target) => { console.log(target.address); });
    const l = thisWorkbook.on("BeforeClose", quiet);
    l.off();
}
"#;
    let (diags, files) = compile(&[("Sheet1", sheet)], &[("main", main_src)]);
    assert!(diags.is_empty(), "{diags:?}");
    let code = module(&files, "Sheet1");
    // A sheet has no instances: nothing owns its listeners, they stay until `off`.
    assert!(code.contains("RAS_Listen New RAS_On_CommandButton_Click, Me.CommandButton1, ras_t1, 0&, Nothing"), "{code}");
    assert!(code.contains("RAS_Listen New RAS_On_TextBox_Change, Me.TextBox1, ras_t2, 0&, Nothing"), "{code}");
    assert!(code.contains("RAS_Listen New RAS_On_Worksheet_Befor_"), "{code}");
    assert!(code.contains("Me.Shapes(\"Button 1\").OnAction = RAS_MacroName(\"Sheet1.refresh\")"), "{code}");
    assert!(code.contains("Public clicks As Long"), "{code}");
    let main = module(&files, "Main");
    assert!(main.contains("shp.OnAction = RAS_MacroName(\"Main.reload\")"), "{main}");
    assert!(main.contains("shp.OnAction = RAS_MacroName(\"Sheet1.refresh\")"), "{main}");
    assert!(main.contains("RAS_Listen New RAS_On_Application_SheetChange, Application, ras_t1, 2&, Nothing"), "{main}");
    assert!(main.contains("Set l = RAS_Listen(New RAS_On_Workbook_BeforeClose, ThisWorkbook, ras_t2, 0&, Nothing)"), "{main}");
    assert!(main.contains("Dim l As Object") && main.contains("Public Function quiet() As Boolean"), "{main}");

    // `this` of the sheet's file inside the closure is the sheet; its variables are the sheet's.
    let lambdas: String = files.iter().filter(|f| f.name.starts_with("RAS_Lambda_")).map(|f| f.code.replace("\r\n", "\n")).collect();
    assert!(lambdas.contains("Sheet1.clicks = Sheet1.clicks + 1&") && lambdas.contains("Me.Captured_Me.Range(\"A1\").Value = Sheet1.clicks"), "{lambdas}");
    assert!(lambdas.contains("Sheet1.refresh\n"), "{lambdas}");
    assert!(lambdas.contains("Apply = Main.quiet\n"), "{lambdas}");

    // An event with a `Cancel` argument by reference is cancelled by a handler that returns `false`.
    let cancel = files.iter().find(|f| f.name.starts_with("RAS_On_Worksheet_Befor_")).map(|f| f.code.replace("\r\n", "\n")).unwrap();
    assert!(cancel.contains("Private WithEvents ras_target As Worksheet"), "{cancel}");
    assert!(cancel.contains("Private Sub ras_target_BeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)"), "{cancel}");
    assert!(cancel.contains("If RAS_Cancels(ras_handler.Apply(Target)) Then Cancel = True"), "{cancel}");
    let close = module(&files, "RAS_On_Workbook_BeforeClose");
    assert!(close.contains("If RAS_Cancels(ras_handler.Apply()) Then Cancel = True"), "{close}");
    let runtime = module(&files, "RAS_Runtime");
    assert!(runtime.contains("Function RAS_Cancels(") && runtime.contains("Function RAS_MacroName(") && !runtime.contains("RAS_UnlistenOwner"), "{runtime}");

    // The debug build registers the handlers with the run time.
    let prog = analyze(sources(&[("Sheet1", sheet)], &[("main", main_src)]), config());
    let out = build_js(&prog, "");
    assert!(out.diagnostics.iter().all(|d| !d.is_error()), "{:#?}", out.diagnostics.iter().map(|d| &d.message).collect::<Vec<_>>());
    let js = &out.files.iter().find(|f| f.rel_path == "main.js").unwrap().text;
    assert!(js.contains("$rt.onAction(shp, \"Main.reload\")") && js.contains("$rt.onAction(shp, \"Sheet1.refresh\")"), "{js}");
    assert!(js.contains("\"BeforeClose\", quiet, null, 0)"), "{js}");
    let js = &out.files.iter().find(|f| f.rel_path == "@objects/Sheet1.js").unwrap().text;
    assert!(js.contains("\"BeforeDoubleClick\", ") && js.contains(", null, 1)"), "{js}");
}

#[test]
fn what_on_cannot_attach() {
    let main = r#"
function helper(x: int) {}
export function twoArgs(a: int, b: int) {}
function cancels(target: Excel.Range, byref cancel: boolean) {}

export function setup(ws: Excel.Worksheet, shp: Excel.Shape, r: Excel.Range) {
    r.on("Change", () => {});                        // E0616: a range has no events
    shp.on("Change", setup2);                        // E0611: only Click
    shp.on("Click", () => {});                       // E0617: not the name of a macro
    shp.on("Click", helper);                         // E0617: not a macro
    ws.on("Change", cancels);                        // E0612
    ws.on("BeforeDoubleClick", cancels);             // E0617: byref
    ws.on("Change", (target: string) => {});         // E0301
    ws.on("NoSuchEvent", (a, b) => {});              // E0611
}
export function setup2() {}
"#;
    let (diags, _) = compile(&[], &[("main", main)]);
    assert_eq!(codes(&diags), ["E0301", "E0611", "E0611", "E0612", "E0616", "E0617", "E0617", "E0617"], "{diags:#?}");
}

/// The example of spec §12.2.3 for sheets, the workbook, `Application` and shapes.
#[test]
fn spec_example_of_on_compiles() {
    let text = include_str!("../../../docs/RAS_Language_Specification.md").replace("\r\n", "\n");
    let start = text.find("// src/watch.ras: ").unwrap();
    let example = &text[start..start + text[start..].find("```").unwrap()];
    let (diags, files) = compile(&[("Sheet1", "")], &[("watch", example)]);
    assert!(diags.is_empty(), "{diags:?}");
    let code = module(&files, "Watch");
    assert!(code.contains(".OnAction = RAS_MacroName(\"Watch.reload\")"), "{code}");
    assert!(code.contains("RAS_Listen New RAS_On_Workbook_BeforeClose, ThisWorkbook, "), "{code}");
}
