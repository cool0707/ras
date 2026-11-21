//! End-to-end tests: the code examples of the language specification
//! (docs/RAS_Language_Specification.md) must compile to the VBA shown there.

#![allow(clippy::field_reassign_with_default)]

use ras_compiler::{compile_str, BuildOptions, BuildOutput, Config, ModuleKind, Program, Severity};

fn config() -> Config {
    let mut cfg = Config::default();
    cfg.sheet_code_names = vec!["Sheet1".into()];
    cfg
}

fn compile_with(files: &[(&str, &str)], cfg: Config) -> (Program, BuildOutput) {
    compile_str(files, cfg, &BuildOptions { erl: Some(false), ..Default::default() })
}

fn compile(files: &[(&str, &str)]) -> (Program, BuildOutput) {
    let (prog, out) = compile_with(files, config());
    let errors: Vec<String> = prog.diagnostics.iter().chain(out.diagnostics.iter()).filter(|d| d.is_error()).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    (prog, out)
}

/// Code of a generated module with LF line endings and without the hash line.
fn code(out: &BuildOutput, module: &str) -> String {
    let f = out.files.iter().find(|f| f.name == module).unwrap_or_else(|| panic!("no module {module}; have {:?}", out.files.iter().map(|f| &f.name).collect::<Vec<_>>()));
    f.code.replace("\r\n", "\n").lines().filter(|l| !l.starts_with("' RAS-GENERATED")).map(|l| format!("{l}\n")).collect()
}

/// Error codes reported for user modules.
fn error_codes(files: &[(&str, &str)]) -> Vec<&'static str> {
    let (prog, out) = compile_with(files, config());
    assert!(out.files.is_empty(), "no code must be generated when there are errors");
    prog.diagnostics.iter().filter(|d| d.is_error() && d.module.is_some_and(|m| prog.modules[m].kind == ModuleKind::User)).map(|d| d.code).collect()
}

#[track_caller]
fn assert_contains(haystack: &str, needle: &str) {
    assert!(haystack.contains(needle), "expected to find:\n{needle}\n--- in ---\n{haystack}");
}

#[test]
fn bundled_type_definitions_are_clean() {
    let (prog, _) = compile(&[("empty", "")]);
    let problems: Vec<String> = prog.diagnostics.iter().filter(|d| d.severity <= Severity::Warning).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(problems.is_empty(), "{problems:#?}");
    assert!(prog.namespaces["Excel"].contains_key("Range"));
    assert!(prog.globals.contains_key("xlUp"));
    // Members and globals of the bundled definitions are seen in camelCase (spec §15.3.1).
    assert!(prog.globals.contains_key("msgBox") && !prog.globals.contains_key("MsgBox"));
}

#[test]
fn spec_3_1_byval_byref() {
    let (_, out) = compile(&[("calc", "function calculate(x: double, byref outResult: double): boolean {\n    outResult = x * 2;\n    return true;\n}\n")]);
    assert_contains(
        &code(&out, "Calc"),
        "Private Function calculate(ByVal x As Double, ByRef outResult As Double) As Boolean\n    outResult = x * 2&\n    calculate = True\nEnd Function\n",
    );
    // No runtime is needed for this.
    assert_eq!(out.files.len(), 1);
}

const STRUTIL: &str = r#"
export function isBlank(this: string): boolean {
    return this.trim().length == 0;
}

export function padLeft(this: string, width: int, ch: string = " "): string {
    return this.length >= width ? this : ch.repeat(width - this.length) + this;
}
"#;

const EXCELUTIL: &str = r#"
export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

export function sum(this: double[]): double {
    let t = 0.0;
    for (const v of this) t += v;
    return t;
}
"#;

#[test]
fn spec_3_2_extension_methods() {
    let main = r#"
import { isBlank, padLeft } from "./StrUtil";
import * as xl from "./ExcelUtil";

function demo(name: string, id: string): int {
    let code = "";
    if (!name.isBlank()) {
        code = id.padLeft(6, "0");
    }
    const last = Sheet1.lastRow();
    return last + code.length;
}
"#;
    let mut cfg = config();
    cfg.opt_override = Some(0);
    let (_, out) = compile_with(&[("StrUtil", STRUTIL), ("ExcelUtil", EXCELUTIL), ("main", main)], cfg);
    let m = code(&out, "Main");
    assert_contains(&m, "    If Not StrUtil.isBlank(name_) Then\n        code = StrUtil.padLeft(id, 6&, \"0\")\n    End If\n    last = ExcelUtil.lastRow(Sheet1, 1&)\n");
    let x = code(&out, "ExcelUtil");
    assert_contains(&x, "Public Function lastRow(ByVal this_ As Worksheet, Optional ByVal col As Long = 1&) As Long\n    lastRow = this_.Cells(this_.Rows.Count, col).End(xlUp).Row\nEnd Function\n");
    assert_contains(&x, "Public Function sum(ByRef this_() As Double) As Double\n");
    let s = code(&out, "StrUtil");
    assert_contains(&s, "    isBlank = Len(Trim$(this_)) = 0&\n");
}

#[test]
fn spec_3_3_named_arguments() {
    let src = r#"
function label(a: int, b: int = 2, c: string = "x"): string { return `${a}${b}${c}`; }
function bump(byref n: int, by: int = 1): void { n += by; }
class Box {
    w: int; h: int;
    constructor(w: int, h: int = 1) { this.w = w; this.h = h; }
    resize(dw: int, dh: int = 0): void { this.w += dw; this.h += dh; }
}
export function lastRow(this: Excel.Worksheet, col: int = 1, base: int = 0): int {
    return this.cells(this.rows.count, col).end(xlUp).row + base;
}
export function demo() {
    const s1 = label(a: 123, c: "abc");
    const s2 = label(1, c: "z");
    let k = 0;
    bump(by: 3, n: k);
    const bx = new Box(h: 2, w: 1);
    bx.resize(dh: 5, dw: 2);
    const n = Sheet1.lastRow(base: 5);
    const r = Sheet1.range("A:A").find("x", lookAt: xlWhole, matchCase: true);
    msgBox("hi", title: "t");
}
"#;
    let mut cfg = config();
    cfg.opt_override = Some(0);
    let (_, out) = compile_with(&[("main", src)], cfg);
    let m = code(&out, "Main");
    // Skipped parameters of RAS functions get their defaults; host members get an empty slot.
    assert_contains(&m, "    s1 = label(123&, 2&, \"abc\")\n    s2 = label(1&, 2&, \"z\")\n");
    assert_contains(&m, "    bump k, 3&\n");
    assert_contains(&m, "    Set bx = Main.New_Box(1&, 2&)\n    bx.resize 2&, 5&\n");
    assert_contains(&m, "    n = lastRow(Sheet1, 1&, 5&)\n");
    assert_contains(&m, "    Set r = Sheet1.Range(\"A:A\").Find(\"x\",,, xlWhole,,, True)\n");
    assert_contains(&m, "    MsgBox \"hi\",, \"t\"\n");

    let decl = "function f(a: int, b: int = 2, c: string = \"x\"): string { return c; }\nfunction g(): void {}\n";
    let cases: &[(&str, &str)] = &[
        ("f(a: 1, d: 2);", "E0412"),
        ("f(a: 1, A: 2);", "E0412"),
        ("f(a: 1, 2);", "E0413"),
        ("f(1, a: 2);", "E0413"),
        ("f(c: \"q\");", "E0414"),
        ("g(x: 1);", "E0412"),
        ("const h = (x: int) => x; h(x: 1);", "E0411"),
        ("const o: any = null; o.foo(a: 1);", "E0411"),
        ("const xs = [1]; xs.push(value: 2);", "E0411"),
        ("f(a: \"s\");", "E0301"),
    ];
    for (body, want) in cases {
        let text = format!("{decl}export function t() {{ {body} }}\n");
        let codes = error_codes(&[("main", &text)]);
        assert!(codes.contains(want), "{body}: expected {want}, got {codes:?}");
    }
}

#[test]
fn spec_2_4_generics_are_monomorphized() {
    let src = r#"
export function maxOf<T extends Comparable>(this: T[]): T {
    let m = this[0];
    for (const v of this) if (v > m) m = v;
    return m;
}

class Stack<T> {
    private items: T[] = [];
    push(x: T): void { this.items.push(x); }
    pop(): T { return this.items.pop(); }
}

function demo(ints: int[], names: string[]): string {
    const a = ints.maxOf();
    const b = names.maxOf();
    const s = new Stack<Excel.Range>();
    s.push(Sheet1.cells(1, 1));
    return b + a + s.pop().address;
}
"#;
    let (_, out) = compile(&[("gen", src)]);
    let m = code(&out, "Gen");
    assert_contains(&m, "Public Function maxOf__Long(ByRef this_() As Long) As Long\n");
    assert_contains(&m, "Public Function maxOf__String(ByRef this_() As String) As String\n");
    assert_contains(&m, "    a = maxOf__Long(ints)\n    b = maxOf__String(names_)\n");
    let cls = code(&out, "Stack__Range");
    assert_contains(&cls, "Private items() As Range\n");
    assert_contains(&cls, "Public Sub push(ByVal x As Range)\n");
    assert_contains(&cls, "    Set items(ras_t1) = x\n");
    assert_contains(&cls, "Public Function pop() As Range\n");
    // The class file carries the standard class header.
    let file = out.files.iter().find(|f| f.name == "Stack__Range").unwrap();
    assert!(file.full_text.starts_with("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Stack__Range\"\r\n"));
    assert!(file.stream_text().starts_with("Attribute VB_Name = \"Stack__Range\"\r\n"));
}

#[test]
fn spec_2_4_unused_generic_emits_nothing_and_constraints_are_checked() {
    let (prog, out) = compile(&[("g", "export function first<T>(this: T[]): T { return this[0]; }\n")]);
    assert!(out.files.is_empty());
    let _ = prog;
    let codes = error_codes(&[("g", "function maxOf<T extends Comparable>(a: T, b: T): T { return a > b ? a : b; }\nfunction f(x: object, y: object) { maxOf(x, y); }\n")]);
    assert_eq!(codes, vec!["E0404"]);
}

#[test]
fn spec_4_2_loop_fusion() {
    let (_, out) = compile(&[("fusion", "const base = 10;\nlet arr = [1, 2, 3];\nlet result = arr.map(x => x * base);\n")]);
    let m = code(&out, "Fusion");
    assert_contains(&m, "Public Sub Main()\n    Dim base As Long\n    Dim arr() As Long\n    Dim result() As Long\n    Dim ras_i As Long\n");
    assert_contains(
        &m,
        "    base = 10&\n    ReDim arr(0 To 2)\n    arr(0) = 1&\n    arr(1) = 2&\n    arr(2) = 3&\n    If RAS_Len(arr) > 0& Then\n        ReDim result(0 To UBound(arr))\n        For ras_i = 0 To UBound(arr)\n            result(ras_i) = arr(ras_i) * base\n        Next ras_i\n    Else\n        Erase result\n    End If\n",
    );
    // No closure class for a lambda written directly as the argument.
    assert!(!out.files.iter().any(|f| f.name.starts_with("RAS_Lambda")));
}

#[test]
fn spec_4_1_filter_map_chain_is_one_loop() {
    let src = "function f(a: int[]): int[] {\n    return a.filter(x => x > 0).map(x => x * 2);\n}\n";
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    assert_eq!(m.matches("For ").count(), 1, "{m}");
    assert_contains(&m, "            If a(ras_i) > 0& Then\n                ras_t1(ras_n) = a(ras_i) * 2&\n                ras_n = ras_n + 1&\n            End If\n");
}

#[test]
fn spec_4_3_escaping_lambda_becomes_a_class() {
    let src = r#"
function applyAll(arr: double[], f: (x: double) => double): double[] {
    return arr.map(f);
}
function demo(arr: double[]): double[] {
    const base = 10;
    const f = (x: double) => x * base;
    return applyAll(arr, f);
}
"#;
    let (prog, out) = compile(&[("m", src)]);
    let cls = code(&out, "RAS_Lambda_001");
    assert_contains(&cls, "Public Captured_base As Long\n");
    assert_contains(&cls, "Public Function Apply(ByVal x As Variant) As Variant\n    Apply = x * Me.Captured_base\nEnd Function\n");
    let m = code(&out, "M");
    assert_contains(&m, "    Set ras_t1 = New RAS_Lambda_001\n    ras_t1.Captured_base = base\n");
    assert_contains(&m, "f.Apply(arr(ras_i))");
    // The user is told that this lambda is slow.
    assert!(prog.diagnostics.iter().chain(out.diagnostics.iter()).any(|d| d.code == "I0641"));
}

#[test]
fn spec_5_2_resume_block() {
    let src = "function ignore() {\n    try {\n        let ws = worksheets(\"NonExistent\");\n        ws.activate();\n    } catch {\n        resume;\n    }\n}\n";
    let (_, out) = compile(&[("m", src)]);
    assert_contains(
        &code(&out, "M"),
        "Private Sub ignore()\n    Dim ws As Worksheet\n\n    On Error Resume Next\n    Set ws = Worksheets(\"NonExistent\")\n    ws.Activate\n    On Error GoTo 0\nEnd Sub\n",
    );
    assert_eq!(error_codes(&[("m", "function f() { try { g(); } catch { g(); resume; } }\nfunction g() {}\n")]), vec!["E0350"]);
    assert_eq!(error_codes(&[("m", "function f() { resume; }\n")]), vec!["E0350"]);
}

#[test]
fn spec_5_1_try_catch_finally() {
    let src = r#"
function risky(n: int): int {
    if (n < 0) throw new Error("negative");
    return n * 2;
}
function safe(n: int): int {
    let result = 0;
    try {
        result = risky(n);
    } catch (e) {
        result = -1;
        console.log(e.description);
        throw e;
    } finally {
        status("done");
    }
    return result;
}
"#;
    let (_, out) = compile(&[("m", src)]);
    let m = code(&out, "M");
    assert_contains(&m, "        Err.Raise vbObjectError + 513&, \"M.risky\", \"negative\"\n");
    // One handler per procedure dispatching on the active try block.
    assert_contains(&m, "RAS_Handler:\n    RAS_ErrCapture Err.Number, Err.Source, Err.Description, Erl, \"src/m.ras\"\n    Select Case ras_tryId\n    Case 1&\n        Resume RAS_Catch_2\n");
    // The catch body runs after `Resume`, so the handler is active again; rethrow reaches `finally`.
    assert_contains(&m, "RAS_Catch_2:\n    ras_tryId = 2&\n    Set e = RAS_CaughtError()\n    result = -1&\n");
    assert_contains(&m, "    RAS_Throw e\n");
    assert_contains(&m, "RAS_Finally_1:\n    ras_tryId = 0&\n    RAS_Status \"done\"\n    Select Case ras_exit1\n    Case 1&\n        RAS_Throw ras_pend1\n    End Select\n");
    assert!(out.files.iter().any(|f| f.name == "RAS_Error"));
}

#[test]
fn spec_5_3_return_inside_finally_and_using() {
    let src = r#"
function f(xs: int[]): int {
    for (const x of xs) {
        try {
            if (x == 3) return x;
            if (x == 4) continue;
            if (x == 5) break;
        } finally {
            status("step");
        }
    }
    return 0;
}
export function g() {
    using _ = fastMode({ calc: false });
    status("working");
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    // Every exit path goes through the finally block and continues afterwards.
    assert_contains(&m, "            f = x\n            ras_exit1 = 2&\n            GoTo RAS_Finally_1\n");
    assert_contains(&m, "    Case 2&\n            Exit Function\n");
    assert_contains(&m, "            GoTo RAS_Continue_");
    assert_contains(&m, "            GoTo RAS_Break_");
    // fastMode({ calc: false }) = everything except calculation (63 - 4), restored in a finally.
    assert_contains(&m, "    RAS_Suspend 59&\n");
    assert_contains(&m, "    RAS_Resume 59&\n");
}

#[test]
fn spec_8_4_inline_expansion() {
    let src = r#"
inline function clamp(v: double, lo: double, hi: double): double {
    return v < lo ? lo : (v > hi ? hi : v);
}
export function total(values: double[], n: int): double {
    let acc = 0.0;
    for (let i = 0; i < n; i++) {
        acc += clamp(values[i], 0, 100);
    }
    return acc;
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("util", src)], cfg);
    let m = code(&out, "Util");
    assert_contains(
        &m,
        "    For i = 0 To n - 1&\n        ' @inline clamp (src/util.ras:2)\n        If values(i) < 0& Then\n            ras_inl1_ret = 0&\n        ElseIf values(i) > 100& Then\n            ras_inl1_ret = 100&\n        Else\n            ras_inl1_ret = values(i)\n        End If\n        acc = acc + ras_inl1_ret\n    Next i\n",
    );
    // A non-exported inline function is not emitted at all.
    assert!(!m.contains("Function clamp"));
    // -O0 keeps the call.
    let mut cfg = config();
    cfg.opt_override = Some(0);
    let (_, out) = compile_with(&[("util", src)], cfg);
    assert_contains(&code(&out, "Util"), "acc = acc + clamp(values(i), 0&, 100&)");
    // Recursive inline functions are rejected.
    assert_eq!(error_codes(&[("m", "inline function f(n: int): int { return n <= 0 ? 0 : f(n - 1); }\n")]), vec!["E0420"]);
}

#[test]
fn spec_9_bulk_range_transfer() {
    let src = r#"
export function fill(last: int) {
    const ws = Sheet1;
    for (let r = 2; r <= last; r++) {
        for (let c = 1; c <= 5; c++) {
            ws.cells(r, c).value = r * c;
        }
    }
}
export function sumColumn(last: int): double {
    const ws = Sheet1;
    let sum = 0.0;
    for (let r = 2; r <= last; r++) {
        sum += ws.cells(r, 1).value;
    }
    return sum;
}
export function conditional(last: int) {
    const ws = Sheet1;
    for (let r = 2; r <= last; r++) {
        if (r % 2 == 0) ws.cells(r, 1).value = r;
    }
}
"#;
    let (_, out) = compile(&[("report", src)]);
    let m = code(&out, "Report");
    assert_contains(
        &m,
        "    If last >= 2& Then\n        ReDim ras_buf1(0 To last - 2&, 0 To 4)\n        For r = 2 To last\n            For c = 1 To 5\n                ras_buf1(r - 2&, c - 1&) = r * c\n            Next c\n        Next r\n        ws.Range(ws.Cells(2, 1), ws.Cells(last, 5)).Value = ras_buf1\n    End If\n",
    );
    assert_contains(
        &m,
        "    If last >= 2& Then\n        ras_src1 = RAS_Read2D(ws.Range(ws.Cells(2, 1), ws.Cells(last, 1)))\n        For r = 2 To last\n            sum = sum + ras_src1(r - 2&, 0)\n        Next r\n    End If\n",
    );
    // Conditional writes must not be converted (condition 5): the cell is written directly.
    assert_contains(&m, "            ws.Cells(r, 1&).Value = r\n");
    assert!(out.notes.iter().any(|n| n.1.kind == "bulk" && !n.1.applied));
    // The runtime that rebases the array comes along.
    assert_contains(&code(&out, "RAS_Runtime"), "Public Function RAS_Read2D(ByVal rng As Object) As Variant\n");
}

#[test]
fn spec_10_3_fast_macro_entry() {
    let (_, out) = compile(&[("report", "@fast\nexport function rebuildReport() {\n    status(\"x\");\n}\n")]);
    let m = code(&out, "Report");
    assert_contains(
        &m,
        "Public Sub rebuildReport()\n    If Not RAS_Enter(\"Report.rebuildReport\") Then Exit Sub\n    On Error GoTo RAS_Handler\n    RAS_SuspendAll\n    RAS_Status \"x\"\nRAS_Exit:\n    RAS_ResumeAll\n    RAS_Leave\n    Exit Sub\nRAS_Handler:\n",
    );
    assert_eq!(out.macros, vec!["Report.rebuildReport".to_string()]);
    let rt = code(&out, "RAS_Runtime");
    assert_contains(&rt, "Public Sub RAS_ResetAppState()");
    assert_contains(&rt, "Application.EnableCancelKey = xlErrorHandler");
}

#[test]
fn spec_10_4_doevents_insertion() {
    let src = r#"
export function work(n: int): int {
    let t = 0;
    while (t < n) {
        for (let c = 0; c < 5; c++) t += c;
    }
    return t;
}
@nodoevents
export function quiet(n: int): int {
    let t = 0;
    while (t < n) t++;
    return t;
}
@udf
export function inCell(n: int): int {
    let t = 0;
    while (t < n) t++;
    return t;
}
"#;
    let (_, out) = compile(&[("m", src)]);
    let m = code(&out, "M");
    // One yield point, in the outer loop: the inner loop has a small constant trip count.
    assert_eq!(m.matches("RAS_Yield").count(), 1, "{m}");
    assert_contains(&m, "        Next c\n        ras_y1 = ras_y1 + 1&\n        If (ras_y1 And ras_yieldMask) = 0& Then RAS_Yield\n    Loop\n");
    let rt = code(&out, "RAS_Runtime");
    assert_contains(&rt, "Declare PtrSafe Function GetInputState Lib \"user32\" () As Long");
    assert_contains(&rt, "    If GetInputState() <> 0& Or now_ - ras_lastYield >= RAS_YIELD_MS Then\n");
}

#[test]
fn spec_11_declare_for_32_and_64_bit() {
    let src = r#"
declare module "user32" {
    function GetForegroundWindow(): ptr;
    @alias("GetWindowTextW")
    function GetWindowText(hWnd: ptr, lpString: wstr, nMaxCount: int): int;
}
function title(): int {
    const h = GetForegroundWindow();
    let buf = " ".repeat(256);
    return GetWindowText(h, buf, 256);
}
"#;
    let (_, out) = compile(&[("m", src)]);
    let m = code(&out, "M");
    assert_contains(
        &m,
        "#If VBA7 Then\nPrivate Declare PtrSafe Function GetForegroundWindow Lib \"user32\" () As LongPtr\nPrivate Declare PtrSafe Function GetWindowText Lib \"user32\" Alias \"GetWindowTextW\" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long\n#Else\nPrivate Declare Function GetForegroundWindow Lib \"user32\" () As Long\nPrivate Declare Function GetWindowText Lib \"user32\" Alias \"GetWindowTextW\" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long\n#End If\n",
    );
    assert_contains(&m, "    #If VBA7 Then\n        Dim h As LongPtr\n    #Else\n        Dim h As Long\n    #End If\n");
    assert_contains(&m, "GetWindowText(h, StrPtr(buf), 256&)");
    // target.vba6 = false drops the VBA6 half.
    let mut cfg = config();
    cfg.target.vba6 = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    assert!(!m.contains("#Else"), "{m}");
    assert_contains(&m, "    Dim h As LongPtr\n");
}

#[test]
fn spec_11_standard_win32_module_is_tree_shaken() {
    let (_, out) = compile(&[("m", "import { Sleep } from \"win32\";\nexport function nap() { Sleep(10); }\n")]);
    assert_contains(&code(&out, "M"), "    Win32.Sleep 10&\n");
    let w = code(&out, "Win32");
    assert_contains(&w, "Public Declare PtrSafe Sub Sleep Lib \"kernel32\" (ByVal dwMilliseconds As Long)");
    assert!(!w.contains("GetTickCount"), "{w}");
}

#[test]
fn spec_12_events_and_init() {
    let src = r#"
let cache = new Map<string, int>();

@init
function prepare() { cache.set("a", 1); }

@on("Workbook.Open")
export function setup() { status("ready"); }

@on("Sheet1.Change")
function onSheet1Change(target: Excel.Range) { target.interior.color = 255; }
"#;
    let (_, out) = compile(&[("report", src)]);
    let sheet = code(&out, "Sheet1");
    assert_contains(
        &sheet,
        "Private Sub Worksheet_Change(ByVal Target As Range)\n    Static ras_busy1 As Boolean\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    If Not ras_busy1 Then\n        ras_busy1 = True\n        Report.onSheet1Change Target\n        ras_busy1 = False\n    End If\n    Exit Sub\n",
    );
    let wb = code(&out, "ThisWorkbook");
    assert_contains(&wb, "Private Sub Workbook_Open()\n    RAS_WorkbookOpen\nEnd Sub\n");
    // What `Workbook_Open` left while the workbook was leaving the protected view.
    assert_contains(&wb, "Private Sub Workbook_Activate()\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    RAS_OpenDeferred\n    Exit Sub\n");
    assert_contains(&wb, "Private Sub Workbook_WindowActivate(ByVal Wn As Window)\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    RAS_OpenDeferred\n    Exit Sub\n");
    // No `Application.OnTime`: Excel opens a workbook again to make a call it was asked for.
    assert!(!out.files.iter().any(|f| f.code.contains("OnTime")));
    let glue = code(&out, "RAS_Glue");
    assert_contains(&glue, "Public Sub RAS_InitAll(Optional ByVal ras_dummy As Long = 0&)\n    Report.RAS_ModInit\n    Report.prepare\nEnd Sub\n");
    assert_contains(&glue, "Public Sub RAS_DeferredInit()\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    Report.setup\n");
    assert_contains(&glue, "Public Sub RAS_WorkbookOpen()\n    If RAS_InProtectedView() Then\n        ras_open_deferred = True\n    Else\n        RAS_DeferredInit\n    End If\nEnd Sub\n");
    assert_contains(&glue, "Public Sub RAS_OpenDeferred()\n    If Not ras_open_deferred Then Exit Sub\n    ras_open_deferred = False\n    RAS_DeferredInit\nEnd Sub\n");
    assert_contains(&glue, "    Set w = app.ProtectedViewWindows.Item(ThisWorkbook.Name)\n    RAS_InProtectedView = Not w Is Nothing\nEnd Function\n");
    assert_contains(&glue, "Public Sub Auto_Open()\n    RAS_EnsureInit\nEnd Sub\n");
    assert_contains(&code(&out, "RAS_Runtime"), "    RAS_Glue.RAS_InitAll\n");
    // Document modules are not importable files and carry no hash line.
    let f = out.files.iter().find(|f| f.name == "Sheet1").unwrap();
    assert!(!f.code.contains("RAS-GENERATED"));
    // Unknown sheet code names are an error (spec §14.9.2 item 7).
    assert_eq!(error_codes_build(&[("m", "@on(\"Sheet9.Change\")\nfunction h() {}\n")]), vec!["E0610"]);
}

/// Error codes reported by code generation.
fn error_codes_build(files: &[(&str, &str)]) -> Vec<&'static str> {
    let (_, out) = compile_with(files, config());
    out.diagnostics.iter().filter(|d| d.is_error()).map(|d| d.code).collect()
}

#[test]
fn spec_13_vba_traps() {
    let src = r#"
function logic(a: int[], i: int, s: string, o: object): int {
    let r = 0;
    if (i < a.length && a[i] > 0) r = 1;
    const q = 7 / 2;
    const mod3 = 7 % 3;
    const d = 7.5 % 2;
    const k = (d as int) + s.indexOf("b");
    const label = i > 0 ? "pos" : "neg";
    const sm: short = 300;
    const prod = sm * sm;
    const same = o == null;
    const flag = (i > 0) as int;
    for (let n = 0; n < 3; n++) {
        let fresh: int;
        r += fresh + n;
    }
    return r + q + mod3 + k + label.length + prod + flag + s.substring(1, 3).length + (same ? 1 : 0) + Math.round(d) as int;
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    // && short-circuits (VBA's And does not)
    assert_contains(&m, "    ras_t1 = i < RAS_Len(a)\n    If ras_t1 Then\n        ras_t1 = a(i) > 0&\n    End If\n");
    // integer division and modulo
    assert_contains(&m, "    q = 7& \\ 2&\n    mod3 = 7& Mod 3&\n    d = RAS_FMod(7.5#, 2&)\n");
    // truncation toward zero, 0-based string index
    assert_contains(&m, "    k = CLng(Fix(d)) + (InStr(1&, s, \"b\", vbBinaryCompare) - 1&)\n");
    // ternary never becomes IIf
    assert!(!m.contains("IIf"));
    assert_contains(&m, "    If i > 0& Then\n        label = \"pos\"\n    Else\n        label = \"neg\"\n    End If\n");
    // Integer * Integer is promoted to Long
    assert_contains(&m, "    prod = CLng(sm) * sm\n");
    // objects compare with Is, True converts to 1
    assert_contains(&m, "    same = o Is Nothing\n    flag = -CLng(i > 0&)\n");
    // a declaration inside a loop is re-initialised on every iteration
    assert_contains(&m, "    For n = 0 To 2\n        fresh = 0&\n");
    // substring(1, 3) → Mid$(s, 2, 2); Math.round is not VBA's banker's Round
    assert_contains(&m, "Len(Mid$(s, 2&, 2&))");
    assert_contains(&m, "RAS_Round(d)");
}

#[test]
fn spec_13_6_identifiers_and_13_7_code_page() {
    let src = "/** 絵文字 😀 のコメント */\nexport function show(value: string, count: int): string {\n    const text = \"a😀b\\n\" + value;\n    return text + count;\n}\n";
    let (_, out) = compile(&[("m", src)]);
    let m = code(&out, "M");
    // built-in member names are renamed; characters outside CP932 are escaped
    assert_contains(&m, "Public Function show(ByVal value_ As String, ByVal count_ As Long) As String\n");
    assert_contains(&m, "' 絵文字 \\u{1F600} のコメント\n");
    assert_contains(&m, "    text_ = \"a\" & (ChrW(&HD83D&) & ChrW(&HDE00&)) & \"b\" & vbLf & value_\n");
    assert_contains(&out.files.iter().find(|f| f.name == "M").unwrap().full_text, "Attribute show.VB_Description = \"絵文字  のコメント\"");
    // identifiers that differ only by case collide in VBA
    assert_eq!(error_codes(&[("m", "function foo() {}\nfunction Foo() {}\n")]), vec!["E0430"]);
}

#[test]
fn lint_rules_of_section_6() {
    // 2: const reassignment
    assert_eq!(error_codes(&[("m", "function f() { const a = 1; a = 2; }\n")]), vec!["E0323"]);
    // 3: lambda capture of a reassigned variable / writing an outer variable
    assert_eq!(error_codes(&[("m", "function f(xs: int[]): int { let t = 0; t = 1; return xs.map(x => x + t).length; }\n")]), vec!["E0320"]);
    assert!(error_codes(&[("m", "function f(xs: int[]) { let t = 0; xs.forEach(x => { t += x; }); }\n")]).contains(&"E0322"));
    // 5: literal passed to byref
    assert_eq!(error_codes(&[("m", "function g(byref x: int) { x = 1; }\nfunction f() { g(5); }\n")]), vec!["E0406"]);
    // `number` does not exist
    assert_eq!(error_codes(&[("m", "function f(x: number) {}\n")]), vec!["E0170"]);
    // narrowing needs `as`, constant overflow is a compile error
    assert_eq!(error_codes(&[("m", "function f(d: double): int { return d; }\n")]), vec!["E0301"]);
    assert_eq!(error_codes(&[("m", "const big = 300000 * 200000;\nfunction f() {}\n")]), vec!["E0397"]);
    // 9 / 10: unqualified Range and .Select are reported but do not block the build
    let (prog, out) = compile(&[("m", "export function f() { range(\"A1\").select(); }\n")]);
    let codes: Vec<&str> = prog.diagnostics.iter().map(|d| d.code).collect();
    assert!(codes.contains(&"W0401") && codes.contains(&"I0402"), "{codes:?}");
    assert!(!out.files.is_empty());
}

#[test]
fn spec_13_4_cells_inside_range_are_qualified() {
    let (prog, out) = compile(&[("m", "function f(ws: Excel.Worksheet): Excel.Range {\n    return ws.range(cells(1, 1), cells(2, 2));\n}\n")]);
    assert_contains(&code(&out, "M"), "    Set f = ws.Range(ws.Cells(1&, 1&), ws.Cells(2&, 2&))\n");
    assert!(!prog.diagnostics.iter().any(|d| d.code == "W0401"));
}

#[test]
fn classes_interfaces_and_enums() {
    let src = r#"
enum Color { Red, Green = 5, Blue }
interface Shape { area(): double; }
class Circle implements Shape {
    radius: double = 1.0;
    private static count: int = 0;
    constructor(r: double) { this.radius = r; Circle.count++; }
    area(): double { return Math.PI * this.radius * this.radius; }
    get diameter(): double { return this.radius * 2; }
    static made(): int { return Circle.count; }
}
export function demo(): double {
    const shapes: Shape[] = [new Circle(2.0)];
    let total = 0.0;
    for (const s of shapes) total += s.area();
    const c = new Circle(1.0);
    return total + c.diameter + Circle.made() + Color.Blue;
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("shapes", src)], cfg);
    // `Circle` is a VBA keyword (a syntax error as a type name), so the class becomes `Circle_`.
    let cls = code(&out, "Circle_");
    assert_contains(&cls, "Implements Shape\n");
    assert_contains(&cls, "Public radius As Double\n");
    assert_contains(&cls, "Private Sub Class_Initialize()\n    radius = 1#\nEnd Sub\n");
    assert_contains(&cls, "Public Sub RAS_Init(ByVal r As Double)\n    radius = r\n    Shapes.Circle__count = Shapes.Circle__count + 1&\nEnd Sub\n");
    assert_contains(&cls, "Public Property Get diameter() As Double\n    diameter = radius * 2&\nEnd Property\n");
    assert_contains(&cls, "Private Function Shape_area() As Double\n    Shape_area = Me.area\nEnd Function\n");
    let iface = code(&out, "Shape");
    assert_contains(&iface, "Public Function area() As Double\nEnd Function\n");
    let m = code(&out, "Shapes");
    assert_contains(&m, "Private Enum Color\n    Color_Red = 0\n    Color_Green = 5\n    Color_Blue = 6\nEnd Enum\n");
    assert_contains(&m, "Public Circle__count As Long\n");
    assert_contains(&m, "Public Function New_Circle_(ByVal r As Double) As Circle_\n    Set New_Circle_ = New Circle_\n    New_Circle_.RAS_Init r\nEnd Function\n");
    assert_contains(&m, "Public Function Circle__made() As Long\n    Circle__made = Shapes.Circle__count\nEnd Function\n");
    assert_contains(&m, "    demo = total + c.diameter + Shapes.Circle__made() + 6&\n");
    // Missing interface members are reported.
    assert_eq!(error_codes(&[("m", "interface A { run(): void; }\nclass B implements A {}\n")]), vec!["E0340"]);
}

#[test]
fn push_in_a_loop_uses_capacity_doubling() {
    let src = "function build(n: int): string[] {\n    let out: string[] = [];\n    for (let i = 0; i < n; i++) {\n        out.push(\"item\" + i);\n    }\n    return out;\n}\n";
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    assert_contains(
        &m,
        "    For i = 0 To n - 1&\n        If ras_len_out >= ras_cap_out Then\n            ras_cap_out = ras_cap_out * 2& + 8&\n            ReDim Preserve out(0 To ras_cap_out - 1&)\n        End If\n        out(ras_len_out) = \"item\" & CStr(i)\n        ras_len_out = ras_len_out + 1&\n    Next i\n    If ras_len_out = 0& Then\n        Erase out\n    ElseIf ras_len_out < ras_cap_out Then\n        ReDim Preserve out(0 To ras_len_out - 1&)\n    End If\n",
    );
}

#[test]
fn erl_line_numbers_and_source_map() {
    let src = "export function f(n: int): int {\n    let t = 0;\n    t += n;\n    return t;\n}\n";
    let (_, out) = compile_str(&[("m", src)], config(), &BuildOptions::default());
    let file = out.files.iter().find(|f| f.name == "M").unwrap();
    let text = file.code.replace("\r\n", "\n");
    // The VBA line number is the RAS source line, so `Erl` reports RAS positions (spec §14.8).
    assert_contains(&text, "\n2   t = 0&\n3   t = t + n\n4   f = t\n");
    let line3 = file.vba_line_for("src/m.ras", 3).unwrap();
    assert_eq!(text.lines().nth(line3 as usize - 1).unwrap(), "3   t = t + n");
    assert!(file.rasmap().contains("\"sources\":[\"src/m.ras\"]"));
    // --release: no line numbers.
    let (_, out) = compile_str(&[("m", src)], config(), &BuildOptions { release: true, ..Default::default() });
    assert_contains(&code(&out, "M"), "    t = t + n\n");
}

#[test]
fn drift_hash_round_trips() {
    let (_, out) = compile(&[("m", "export function f(): int { return 1; }\n")]);
    let f = out.files.iter().find(|f| f.name == "M").unwrap();
    assert_eq!(ras_compiler::emit::recorded_hash(&f.code), Some(f.hash));
    assert_eq!(ras_compiler::emit::has_drifted(&f.stream_text()), Some(false));
    assert_eq!(ras_compiler::emit::has_drifted(&f.stream_text().replace("f = 1&", "f = 2&")), Some(true));
}

#[test]
fn spec_2_4_unbounded_generic_recursion_is_an_error() {
    let src = "function grow<T>(x: T, n: int): int {\n    if (n <= 0) return 0;\n    const xs: T[] = [x];\n    return grow(xs, n - 1);\n}\nexport function run(): int { return grow(1, 3); }\n";
    assert!(error_codes_build(&[("m", src)]).contains(&"E0602"));
}

const POINT: &str = "struct Point { x: int; y: int; name?: string; }\n";

#[test]
fn spec_2_5_struct_literals_are_written_in_place() {
    let src = format!(
        "{POINT}function make(x: int, y: int): Point {{\n    return {{ x, y }};\n}}\nexport function run(): int {{\n    let p: Point = {{ x: 1, y: 2, name: \"a\" }};\n    p = {{ x: p.y, y: p.x }};\n    const o = {{ count: 1, label: \"two\" }};\n    const o2 = {{ label: \"x\", count: 2 }};\n    return make(p.x, o.count + o2.count).y;\n}}\n"
    );
    let (_, out) = compile(&[("geo", &src)]);
    let m = code(&out, "Geo");
    assert_contains(&m, "Public Type Point\n    x As Long\n    y As Long\n    name As String\nEnd Type\n");
    // `return { ... }` fills the return value; omitted fields of a fresh value need no reset.
    assert_contains(&m, "Private Function make(ByVal x As Long, ByVal y As Long) As Point\n    make.x = x\n    make.y = y\nEnd Function\n");
    assert_contains(&m, "    p.x = 1&\n    p.y = 2&\n    p.name = \"a\"\n");
    // The right-hand side reads the target: built in a temporary, then copied.
    assert_contains(&m, "    ras_t1.x = p.y\n    ras_t1.y = p.x\n    p = ras_t1\n");
    // Anonymous object types are identified by their fields, whatever the order.
    assert_contains(&m, "Public Type RAS_Obj1\n    count As Long\n    label As String\nEnd Type\n");
    assert!(!m.contains("RAS_Obj2"));
    assert_contains(&m, "    o2.label = \"x\"\n    o2.count = 2&\n");
}

#[test]
fn spec_2_5_tuples_and_destructuring() {
    let src = format!(
        "{POINT}function make(x: int, y: int): Point {{\n    return {{ x, y }};\n}}\nfunction divmod(a: int, b: int): [int, int] {{\n    return [a / b, a % b];\n}}\nexport function run(xs: int[]): int {{\n    const {{ x, y: py }} = make(1, 2);\n    const [q, r] = divmod(17, 5);\n    const [first, , third] = xs;\n    const t = divmod(1, 2);\n    return x + py + q + r + first + third + t[1];\n}}\n"
    );
    let (_, out) = compile(&[("m", &src)]);
    let m = code(&out, "M");
    assert_contains(&m, "Public Type RAS_Tup1\n    item0 As Long\n    item1 As Long\nEnd Type\n");
    assert_contains(&m, "Private Function divmod(ByVal a As Long, ByVal b As Long) As RAS_Tup1\n    divmod.item0 = a \\ b\n    divmod.item1 = a Mod b\nEnd Function\n");
    assert_contains(&m, "    ras_d1 = make(1&, 2&)\n    x = ras_d1.x\n    py = ras_d1.y\n");
    assert_contains(&m, "    ras_d2 = divmod(17&, 5&)\n    q = ras_d2.item0\n    r = ras_d2.item1\n");
    // A plain variable on the right is read directly; holes skip an element.
    assert_contains(&m, "    first = xs(0)\n    third = xs(2)\n");
    assert_contains(&m, "t.item1");
}

#[test]
fn spec_2_5_struct_arrays() {
    let src = format!(
        "{POINT}export function run(n: int): int {{\n    let ps: Point[] = [];\n    for (let i = 0; i < n; i++) {{\n        ps.push({{ x: i, y: i * 2 }});\n    }}\n    let total = 0;\n    for (const p of ps) total += p.x + p.y;\n    for (const {{ x, y }} of ps) total += x * y;\n    for (const p of ps) {{\n        let c = p;\n        c.x = 0;\n        total += c.y;\n    }}\n    return total + ps.length;\n}}\n"
    );
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (prog, out) = compile_with(&[("m", &src)], cfg);
    assert!(!prog.diagnostics.iter().any(|d| d.is_error()), "{:?}", prog.diagnostics);
    let m = code(&out, "M");
    // The literal is built in the array element.
    assert_contains(&m, "        ps(ras_len_ps).x = i\n        ps(ras_len_ps).y = i * 2&\n        ps(ras_len_ps).name = vbNullString\n        ras_len_ps = ras_len_ps + 1&\n");
    // Field reads go to the element: no copy into the loop variable.
    assert_contains(&m, "    For ras_i = 0 To RAS_Len_Point(ps) - 1&\n        total = total + (ps(ras_i).x + ps(ras_i).y)\n    Next ras_i\n");
    assert_contains(&m, "        x = ps(ras_i_2).x\n        y = ps(ras_i_2).y\n        total = total + x * y\n");
    // The variable is used as a whole: it stays a copy.
    assert_contains(&m, "        p = ps(ras_i_3)\n        c = p\n        c.x = 0&\n");
    // An array of structs cannot be passed to the run time as a Variant.
    assert!(!m.contains("RAS_Len("));
    assert_contains(&m, "Public Function RAS_Len_Point(ByRef arr() As Point) As Long\n    On Error Resume Next\n    RAS_Len_Point = UBound(arr) - LBound(arr) + 1&\nEnd Function\n");
}

#[test]
fn spec_2_5_map_of_structs_becomes_a_class() {
    let src = format!(
        "{POINT}class Holder {{\n    private pos: Point;\n    constructor(p: Point) {{ this.pos = p; }}\n    where(): Point {{ return this.pos; }}\n}}\nexport function run(): int {{\n    const pm = new Map<string, Point>();\n    pm.set(\"a\", {{ x: 1, y: 2 }});\n    const h = new Holder(pm.get(\"a\"));\n    const plain = new Map<string, int>();\n    plain.set(\"k\", 1);\n    return h.where().x + pm.size + plain.size;\n}}\n"
    );
    let (_, out) = compile(&[("m", &src)]);
    let m = code(&out, "M");
    assert_contains(&m, "    Dim pm As RAS_StructMap__String_Point\n");
    assert_contains(&m, "    Set pm = RasCollections.New_RAS_StructMap__String_Point()\n    ras_t1.x = 1&\n    ras_t1.y = 2&\n    pm.set_ \"a\", ras_t1\n");
    // Other maps stay dictionaries.
    assert_contains(&m, "    Set plain = CreateObject(\"Scripting.Dictionary\")\n");
    let class = code(&out, "RAS_StructMap__String_Point");
    assert_contains(&class, "Private slots() As Point\n");
    // Members whose signature mentions a `Type` cannot be Public in a class module.
    assert_contains(&class, "Friend Function get_(ByVal key As String) As Point\n");
    assert_contains(&class, "Friend Sub set_(ByVal key As String, ByRef value_ As Point)\n");
    assert_contains(&class, "Public Function has(ByVal key As String) As Boolean\n");
    let holder = code(&out, "Holder");
    assert_contains(&holder, "Private pos As Point\n");
    assert_contains(&holder, "Friend Function where() As Point\n");
    assert_contains(&holder, "Friend Sub RAS_Init(ByRef p As Point)\n");
    // Locals of user code are not renamed because of the members of the bundled class.
    let (_, out) = compile(&[("m", "export function f(values: int[]): int {\n    const keys = values.length;\n    return keys;\n}\n")]);
    assert_contains(&code(&out, "M"), "Public Function f(ByRef values() As Long) As Long\n");
}

#[test]
fn spec_2_5_structs_cannot_leave_typed_code() {
    let errs = |body: &str| error_codes(&[("m", &format!("{POINT}{body}"))]);
    assert_eq!(errs("export function f() {\n    const p: Point = { x: 1 };\n}\n"), vec!["E0392"]);
    assert_eq!(errs("export function f() {\n    const p: Point = { x: 1, y: 2, z: 3 };\n}\n"), vec!["E0202"]);
    assert_eq!(errs("export function f(p: Point) {\n    const a: any = p;\n}\n"), vec!["E0301"]);
    assert_eq!(errs("export function f(p: Point) {\n    console.log(p);\n}\n"), vec!["E0301"]);
    assert_eq!(errs("export function f(p: Point) {\n    const o = { x: 1, y: 2 };\n    const q: Point = o;\n}\n"), vec!["E0301"]);
    assert_eq!(errs("export function f(ps: Point[], p: Point): int {\n    return ps.indexOf(p);\n}\n"), vec!["E0393"]);
    assert_eq!(errs("class C { pos: Point; }\n"), vec!["E0394"]);
    assert_eq!(errs("interface I { where(): Point; }\n"), vec!["E0394"]);
    assert_eq!(errs("export function f(m: Map<string, Point>) {\n    m.get(\"a\").x = 1;\n}\n"), vec!["E0325"]);
    assert_eq!(errs("export function f(): int {\n    const t: [int, string] = [1, \"a\"];\n    return t[2];\n}\n"), vec!["E0389"]);
    assert_eq!(errs("export function f(): int {\n    const { nope } = { a: 1 };\n    return nope;\n}\n"), vec!["E0202"]);
    // A lambda that becomes a class cannot take a struct.
    let src = format!("{POINT}export function f(p: Point): int {{\n    const g = (v: Point) => v.x;\n    return g(p);\n}}\n");
    assert!(error_codes_build(&[("m", &src)]).contains(&"E0395"));
}

#[test]
fn spec_15_3_1_host_names_are_camel_case() {
    let src = r#"
export function touch(ws: Excel.Worksheet): string {
    const range = ws.range("A1");
    range.interior.color = rgb(255, 0, 0);
    range.delete();
    activeSheet.cells(1, 1).value2 = 1;
    debug.print(uBound(split("a,b", ",")), Switch(true, 1));
    return application.name + ws.name;
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (prog, out) = compile_with(&[("m", src)], cfg);
    // The qualified-reference lint (rule 9) still fires for `activeSheet`, and `cells` is not a member of a
    // chart sheet (`activeSheet` is `Excel.Worksheet | Excel.Chart`, spec §2.6.5); nothing complains about the spelling.
    let codes: Vec<&str> = prog.diagnostics.iter().filter(|d| d.severity <= Severity::Warning).map(|d| d.code).collect();
    assert_eq!(codes, vec!["W0401", "W0305"]);
    let m = code(&out, "M");
    // The generated VBA carries the type library's spelling; a local named like a host global is renamed.
    assert_contains(&m, "    Dim range_ As Range\n");
    assert_contains(&m, "    Set range_ = ws.Range(\"A1\")\n    range_.Interior.Color = RGB(255&, 0&, 0&)\n");
    assert_contains(&m, "    range_.Delete");
    // Late-bound members (`activeSheet` is stored in an Object) are spelled like the bundled definitions too.
    assert_contains(&m, "    ActiveSheet.Cells(1&, 1&).Value2 = 1&\n");
    // `switch` is a keyword, so that global keeps its VBA name.
    assert_contains(&m, "    Debug.Print UBound(Split(\"a,b\", \",\")), Switch(True, 1&)\n");
    assert_contains(&m, "    touch = Application.Name & ws.Name\n");

    // The type library's spelling is still understood (VBA is not case-sensitive), with a warning and a fix.
    let old = "export function f(ws: Excel.Worksheet) {\n    ws.Cells(1, 1).Value = Application.Name;\n    MsgBox(\"x\");\n}\n";
    let (prog, out) = compile(&[("m", old)]);
    let warned: Vec<&str> = prog.diagnostics.iter().filter(|d| d.code == "W0304").map(|d| d.message.split('`').nth(3).unwrap()).collect();
    assert_eq!(warned, vec!["cells", "value", "application", "name", "msgBox"]);
    assert_contains(&code(&out, "M"), "    ws.Cells(1&, 1&).Value = Application.Name\n    MsgBox \"x\"\n");
    let fixes: Vec<&ras_compiler::Fix> = prog.fixes.iter().collect();
    assert_eq!(
        ras_compiler::apply_fixes(old, &fixes),
        "export function f(ws: Excel.Worksheet) {\n    ws.cells(1, 1).value = application.name;\n    msgBox(\"x\");\n}\n"
    );

    // Names the project declares itself (sheet code names, `types/*.d.ras`) are used as written.
    let decl = "declare const Summary: Excel.Worksheet;\ndeclare function LegacyMacro(n: int): void;\n";
    let (prog, out) = compile(&[("sheets.d", decl), ("m", "export function f() {\n    LegacyMacro(Summary.index + Sheet1.index);\n}\n")]);
    assert!(prog.diagnostics.iter().all(|d| d.code != "W0304"));
    assert_contains(&code(&out, "M"), "    LegacyMacro Summary.Index + Sheet1.Index\n");
}

#[test]
fn spec_4_4_collections_have_array_methods() {
    let src = r#"
export function visibleNames(): string[] {
    const names = worksheets.filter(ws => ws.visible == xlSheetVisible).map(ws => ws.name);
    return names;
}

export function stamp() {
    worksheets.forEach((ws, i) => { ws.cells(1, 1).value = i; });
}

export function lookup(book: Excel.Workbook): Excel.Worksheet {
    const n = book.worksheets.findIndex(ws => ws.name == "Data");
    const any = workbooks.some(wb => wb.readOnly);
    const total = Sheet1.range("A1:A3").reduce((acc, c) => acc + c.row, 0);
    const all = book.worksheets.toArray();
    return n + total + all.length > 0 && any ? null : worksheets.find(ws => ws.name == "Data");
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (prog, out) = compile_with(&[("m", src)], cfg);
    let problems: Vec<String> = prog.diagnostics.iter().filter(|d| d.severity <= Severity::Warning).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(problems.is_empty(), "{problems:#?}");
    let m = code(&out, "M");
    // `filter` + `map` are one `For Each` loop; the members of `worksheets` are typed as `Worksheet`.
    assert_contains(&m, "    Dim ras_e As Worksheet\n");
    assert_contains(
        &m,
        "    ras_n = 0&\n    If Worksheets.Count > 0& Then\n        ReDim names_(0 To Worksheets.Count - 1&)\n        For Each ras_e In Worksheets\n            If ras_e.Visible = xlSheetVisible Then\n                names_(ras_n) = ras_e.Name\n                ras_n = ras_n + 1&\n            End If\n        Next ras_e\n    End If\n    If ras_n = 0& Then\n        Erase names_\n    ElseIf ras_n <= UBound(names_) Then\n        ReDim Preserve names_(0 To ras_n - 1&)\n    End If\n",
    );
    // The index is counted only when the callback takes it.
    assert_contains(&m, "    ras_i = 0&\n    For Each ras_e In Worksheets\n        ras_e.Cells(1&, 1&).Value = ras_i\n        ras_i = ras_i + 1&\n    Next ras_e\n");
    // A collection that is not a plain name is evaluated once.
    assert_contains(&m, "    Set ras_t1 = book.Worksheets\n    ras_t2 = -1&\n    ras_i = 0&\n    For Each ras_e In ras_t1\n        If ras_e.Name = \"Data\" Then\n            ras_t2 = ras_i\n            Exit For\n        End If\n        ras_i = ras_i + 1&\n    Next ras_e\n    n = ras_t2\n");
    assert_contains(&m, "    ras_t3 = False\n    For Each ras_e_2 In Workbooks\n        If ras_e_2.ReadOnly Then\n            ras_t3 = True\n            Exit For\n        End If\n    Next ras_e_2\n");
    assert_contains(&m, "    Set ras_t4 = Sheet1.Range(\"A1:A3\")\n    ras_t5 = 0&\n    For Each ras_e_3 In ras_t4\n        ras_t5 = ras_t5 + ras_e_3.Row\n    Next ras_e_3\n");
    // `toArray()` copies the members into an array.
    assert_contains(&m, "    If ras_t6.Count > 0& Then\n        ReDim all(0 To ras_t6.Count - 1&)\n        ras_n = 0&\n        For Each ras_e_4 In ras_t6\n            Set all(ras_n) = ras_e_4\n            ras_n = ras_n + 1&\n        Next ras_e_4\n    Else\n        Erase all\n    End If\n");
    // `find` yields `Nothing` when no member matches.
    assert_contains(&m, "        Set ras_t7 = Nothing\n        For Each ras_e_5 In Worksheets\n            If ras_e_5.Name = \"Data\" Then\n                Set ras_t7 = ras_e_5\n                Exit For\n            End If\n        Next ras_e_5\n");

    // A member of the type wins over the array-style method of the same name (`Range.Find`).
    let (_, out) = compile(&[("m", "export function f(r: Excel.Range): Excel.Range {\n    return r.find(\"x\");\n}\n")]);
    assert_contains(&code(&out, "M"), "    Set f = r.Find(\"x\")\n");
    // A host object without a default member and `Count` is not a collection.
    let (prog, _) = compile(&[("m", "export function f(ws: Excel.Worksheet) {\n    ws.toArray();\n}\n")]);
    assert!(prog.diagnostics.iter().any(|d| d.code == "W0301"), "late binding is reported");
}

#[test]
fn spec_4_1_index_after_filter_counts_the_elements_that_passed() {
    let src = r#"
export function tagged(a: int[]): int[] {
    const r = a.filter(x => x > 1).map((x, i) => x * 10 + i);
    return r;
}

export function position(a: int[]): int {
    return a.filter((x, i) => i > 0).filter((x, i) => i != 1).map((x, i) => x + i).findIndex(x => x > 5);
}

export function weighted(a: int[]): int {
    return a.filter(x => x > 1).reduce((acc, x, i) => acc + x * i, 0);
}

export function plain(a: int[]): int {
    return a.filter(x => x > 1).reduce((acc, x) => acc + x, 0);
}

export function sheets(): string[] {
    const r = worksheets.filter(ws => ws.visible == xlSheetVisible).map((ws, i) => `${i}:${ws.name}`);
    return r;
}
"#;
    let mut cfg = config();
    cfg.doevents.enabled = false;
    let (_, out) = compile_with(&[("m", src)], cfg);
    let m = code(&out, "M");
    // The callback of `map` receives the position in the filtered array: the output counter.
    assert_contains(&m, "        For ras_i = 0 To UBound(a)\n            If a(ras_i) > 1& Then\n                r(ras_n) = a(ras_i) * 10& + ras_n\n                ras_n = ras_n + 1&\n            End If\n        Next ras_i\n");
    // Each `filter` counts what it lets through; `findIndex` answers in terms of the last one.
    assert_contains(
        &m,
        "    ras_t1 = -1&\n    ras_k = 0&\n    ras_k_2 = 0&\n    For ras_i = 0 To RAS_Len(a) - 1&\n        If ras_i > 0& Then\n            If ras_k <> 1& Then\n                x = a(ras_i) + ras_k_2\n                If x > 5& Then\n                    ras_t1 = ras_k_2\n                    Exit For\n                End If\n                ras_k_2 = ras_k_2 + 1&\n            End If\n            ras_k = ras_k + 1&\n        End If\n    Next ras_i\n",
    );
    assert_contains(&m, "    ras_k = 0&\n    For ras_i = 0 To RAS_Len(a) - 1&\n        If a(ras_i) > 1& Then\n            ras_t1 = ras_t1 + a(ras_i) * ras_k\n            ras_k = ras_k + 1&\n        End If\n    Next ras_i\n");
    // Nothing is counted when no callback asks for the index.
    assert_contains(&m, "    ras_t1 = 0&\n    For ras_i = 0 To RAS_Len(a) - 1&\n        If a(ras_i) > 1& Then\n            ras_t1 = ras_t1 + a(ras_i)\n        End If\n    Next ras_i\n");
    // The same holds for a collection, which then needs no counter for the source position.
    assert_contains(&m, "        For Each ras_e In Worksheets\n            If ras_e.Visible = xlSheetVisible Then\n                r(ras_n) = CStr(ras_n) & \":\" & ras_e.Name\n                ras_n = ras_n + 1&\n            End If\n        Next ras_e\n");
}

#[test]
fn spec_12_1_files_without_code_are_no_modules() {
    let files = [
        ("Sheet1Events", "// シート Sheet1 のイベント処理を書くファイルです。\n\n/* まだ何もない */\n"),
        ("onlyImports", "import { twice } from \"./util\";\n// twice を使う予定\n"),
        ("util", "export function twice(x: int): int { return x * 2; }\n"),
        ("onlyConst", "const LIMIT = 10;\n"),
    ];
    let (_, out) = compile(&files);
    let names: Vec<&str> = out.files.iter().map(|f| f.name.as_str()).collect();
    assert!(!names.contains(&"Sheet1Events") && !names.contains(&"OnlyImports"), "{names:?}");
    assert!(names.contains(&"Util") && names.contains(&"OnlyConst"), "{names:?}");
    // Nothing refers to the modules left out.
    assert!(!out.files.iter().any(|f| f.code.contains("Sheet1Events") || f.code.contains("OnlyImports")));
}
