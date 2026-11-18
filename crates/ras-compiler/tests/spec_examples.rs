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
    assert!(prog.globals.contains_key("MsgBox"));
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
    return this.Cells(this.Rows.Count, col).End(xlUp).Row;
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
    s.push(Sheet1.Cells(1, 1));
    return b + a + s.pop().Address;
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
    assert_contains(&cls, "Public Function Invoke(ByVal x As Variant) As Variant\n    Invoke = x * Me.Captured_base\nEnd Function\n");
    let m = code(&out, "M");
    assert_contains(&m, "    Set ras_t1 = New RAS_Lambda_001\n    ras_t1.Captured_base = base\n");
    assert_contains(&m, "f.Invoke(arr(ras_i))");
    // The user is told that this lambda is slow.
    assert!(prog.diagnostics.iter().chain(out.diagnostics.iter()).any(|d| d.code == "I0641"));
}

#[test]
fn spec_5_2_resume_block() {
    let src = "function ignore() {\n    try {\n        let ws = Worksheets(\"NonExistent\");\n        ws.Activate();\n    } catch {\n        resume;\n    }\n}\n";
    let (_, out) = compile(&[("m", src)]);
    assert_contains(
        &code(&out, "M"),
        "Private Sub ignore()\n    Dim ws As Object\n\n    On Error Resume Next\n    Set ws = Worksheets(\"NonExistent\")\n    ws.Activate\n    On Error GoTo 0\nEnd Sub\n",
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
            ws.Cells(r, c).Value = r * c;
        }
    }
}
export function sumColumn(last: int): double {
    const ws = Sheet1;
    let sum = 0.0;
    for (let r = 2; r <= last; r++) {
        sum += ws.Cells(r, 1).Value;
    }
    return sum;
}
export function conditional(last: int) {
    const ws = Sheet1;
    for (let r = 2; r <= last; r++) {
        if (r % 2 == 0) ws.Cells(r, 1).Value = r;
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
function onSheet1Change(target: Excel.Range) { target.Interior.Color = 255; }
"#;
    let (_, out) = compile(&[("report", src)]);
    let sheet = code(&out, "Sheet1");
    assert_contains(
        &sheet,
        "Private Sub Worksheet_Change(ByVal Target As Range)\n    Static ras_busy1 As Boolean\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    If Not ras_busy1 Then\n        ras_busy1 = True\n        Report.onSheet1Change Target\n        ras_busy1 = False\n    End If\n    Exit Sub\n",
    );
    let wb = code(&out, "ThisWorkbook");
    assert_contains(&wb, "Private Sub Workbook_Open()\n    Application.OnTime Now, \"'\" & ThisWorkbook.Name & \"'!RAS_DeferredInit\"\nEnd Sub\n");
    assert_contains(&wb, "Private Sub Workbook_Activate()");
    let glue = code(&out, "RAS_Glue");
    assert_contains(&glue, "Public Sub RAS_InitAll(Optional ByVal ras_dummy As Long = 0&)\n    Report.RAS_ModInit\n    Report.prepare\nEnd Sub\n");
    assert_contains(&glue, "Public Sub RAS_DeferredInit()\n    On Error GoTo RAS_OnError\n    RAS_EnsureInit\n    Report.setup\n");
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
    let (prog, out) = compile(&[("m", "export function f() { Range(\"A1\").Select(); }\n")]);
    let codes: Vec<&str> = prog.diagnostics.iter().map(|d| d.code).collect();
    assert!(codes.contains(&"W0401") && codes.contains(&"I0402"), "{codes:?}");
    assert!(!out.files.is_empty());
}

#[test]
fn spec_13_4_cells_inside_range_are_qualified() {
    let (prog, out) = compile(&[("m", "function f(ws: Excel.Worksheet): Excel.Range {\n    return ws.Range(Cells(1, 1), Cells(2, 2));\n}\n")]);
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
    let cls = code(&out, "Circle");
    assert_contains(&cls, "Implements Shape\n");
    assert_contains(&cls, "Public radius As Double\n");
    assert_contains(&cls, "Private Sub Class_Initialize()\n    radius = 1#\nEnd Sub\n");
    assert_contains(&cls, "Public Sub RAS_Init(ByVal r As Double)\n    radius = r\n    Shapes.Circle_count = Shapes.Circle_count + 1&\nEnd Sub\n");
    assert_contains(&cls, "Public Property Get diameter() As Double\n    diameter = radius * 2&\nEnd Property\n");
    assert_contains(&cls, "Private Function Shape_area() As Double\n    Shape_area = Me.area\nEnd Function\n");
    let iface = code(&out, "Shape");
    assert_contains(&iface, "Public Function area() As Double\nEnd Function\n");
    let m = code(&out, "Shapes");
    assert_contains(&m, "Private Enum Color\n    Color_Red = 0\n    Color_Green = 5\n    Color_Blue = 6\nEnd Enum\n");
    assert_contains(&m, "Public Circle_count As Long\n");
    assert_contains(&m, "Public Function New_Circle(ByVal r As Double) As Circle\n    Set New_Circle = New Circle\n    New_Circle.RAS_Init r\nEnd Function\n");
    assert_contains(&m, "Public Function Circle_made() As Long\n    Circle_made = Shapes.Circle_count\nEnd Function\n");
    assert_contains(&m, "    demo = total + c.diameter + Shapes.Circle_made() + 6&\n");
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
