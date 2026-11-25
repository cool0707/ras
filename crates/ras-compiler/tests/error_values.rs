//! Cell error values (`#N/A` ...): the type `cellValue`, its narrowing, the warning W0313
//! (spec §2.6.5) and `Excel.errorToNull` (spec §13.3).

#![allow(clippy::field_reassign_with_default)]

use ras_compiler::{compile_str, BuildOptions, BuildOutput, Config, ModuleKind, Program};

fn config() -> Config {
    let mut cfg = Config::default();
    cfg.doevents.enabled = false;
    cfg.sheet_code_names = vec!["Sheet1".into()];
    cfg
}

fn compile(src: &str) -> (Program, BuildOutput) {
    let (prog, out) = compile_str(
        &[("m", src)],
        config(),
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
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    (prog, out)
}

/// (line, message) of the W0313 warnings, lines counted from 1.
fn warnings(src: &str) -> Vec<(usize, String)> {
    let (prog, _) = compile(src);
    let mut ws: Vec<(usize, String)> = prog
        .diagnostics
        .iter()
        .filter(|d| {
            d.code == "W0313"
                && d.module
                    .is_some_and(|m| prog.modules[m].kind == ModuleKind::User)
        })
        .map(|d| {
            let start = u32::from(d.span.start()) as usize;
            (src[..start].matches('\n').count() + 1, d.message.clone())
        })
        .collect();
    ws.sort();
    ws
}

/// The lines (from 1) with a W0313 warning.
fn warned_lines(src: &str) -> Vec<usize> {
    warnings(src).into_iter().map(|(l, _)| l).collect()
}

/// The lines (from 1) of `src` that end with `// W`.
fn marked_lines(src: &str) -> Vec<usize> {
    src.lines()
        .enumerate()
        .filter(|(_, l)| l.trim_end().ends_with("// W"))
        .map(|(i, _)| i + 1)
        .collect()
}

#[track_caller]
fn assert_marked(src: &str) {
    assert_eq!(warned_lines(src), marked_lines(src), "{:#?}", warnings(src));
}

#[test]
fn unchecked_uses_are_reported() {
    assert_marked(
        r#"
export function f(ws: Excel.Worksheet): void {
    const v = ws.range("A1").value;
    if (v == "x") {}                                // W
    if (ws.range("A2").value > 3) {}                // W
    const n = ws.range("A3").value as int;          // W
    const s: string = ws.range("A4").value;         // W
    msgBox(v);                                      // W
    const total = v + 1;                            // W
    if (ws.range("A5").value) {}                    // W
    const b = !v;                                   // W
    switch (v) {                                    // W
        case 1: break;
    }
    const arr = [1, 2, 3];
    const x = arr[v];                               // W
    const hit = application.vLookup("k", ws.range("A1:B9"), 2, false);
    if (hit == "a") {}                              // W
    const e = cvErr(2042);
    if (e == 1) {}                                  // W
    const t = ws.range("B1").value2;
    if (t < 0) {}                                   // W
    let u: any = 0;
    u = ws.range("B2").value;                       // W
    const list: any[] = [];
    list.push(ws.range("B3").value);                // W
}
"#,
    );
}

#[test]
fn checked_and_harmless_uses_are_not_reported() {
    assert_marked(
        r#"
export function f(ws: Excel.Worksheet): string {
    const v = ws.range("A1").value;
    if (isError(v)) return "";
    if (v == "x") {}
    const w = ws.range("B1").value;
    if (!isError(w) && w > 3) {}
    if (isError(w) || w == "") {}
    if (!isError(w)) { const k = w as string; }
    const y = isError(w) ? 0 : w + 1;
    if (typeof w == "String") { const k2 = w as string; }
    if (typeof w != "Error") { const k3 = w as string; }
    if (isNumeric(w)) { const k4 = w * 2; }
    if (application.worksheetFunction.isError(w)) {} else { const k5 = w * 2; }
    const z = Excel.errorToNull(ws.range("C1").value) ?? "";
    if (z == "a") {}
    if (ws.range("D1").value == null) {}
    console.log(ws.range("E1").value);
    debug.print(ws.range("E1").value);
    console.log(`${ws.range("E2").value}`);
    const kind = typeName(ws.range("E3").value);
    ws.range("F1").value = ws.range("E1").value;
    const keep: cellValue = ws.range("F2").value;
    const raw = ws.range("F3").value as any;
    const m = Excel.match("k", ws.range("A1:A9"));
    if (m == 1) {}
    const a = application.match(ws.range("G1").value, ws.range("A1:A9"), 0);
    const d = application.worksheetFunction.ifError(ws.range("G2").value, 0);
    return kind;
}
"#,
    );
}

#[test]
fn narrowing_ends_where_the_variable_is_assigned() {
    assert_marked(
        r#"
export function f(ws: Excel.Worksheet): void {
    for (let i = 0; i < 3; i++) {
        const cell = ws.cells(i + 1, 1).value;
        if (isError(cell)) continue;
        if (cell == "x") {}
    }
    for (let i = 0; i < 3; i++) {
        const cell = ws.cells(i + 1, 2).value;
        if (cell == "x") {}                         // W
    }
    let v = ws.range("A1").value;
    if (!isError(v)) {
        v = ws.range("A2").value;
        if (v == 1) {}                              // W
    }
}
"#,
    );
}

#[test]
fn bulk_reads_are_cell_values() {
    assert_marked(
        r#"
export function f(ws: Excel.Worksheet): void {
    const data = ws.range("A1:C10").toArray();
    if (data[0][1] > 0) {}                          // W
    for (let i = 0; i < data.length; i++) {
        const cell = data[i][0];
        if (isError(cell)) continue;
        if (cell == "x") {}
    }
    for (const c of ws.range("A1:A9").toColumn()) {
        if (c == "x") {}                            // W
    }
    const clean = Excel.errorToNull(ws.range("A1:C10").toArray());
    if (clean[0][1] > 0) {}
    if (ws.usedTable()[0][0] == "") {}              // W
    const grid = ws.range("A1:C3").value;
    if (grid[0][0] == 1) {}                         // W
}
"#,
    );
}

#[test]
fn the_type_crosses_functions_and_fields() {
    assert_marked(
        r#"
function lookup(key: string) {
    return application.vLookup(key, Sheet1.range("A1:B9"), 2, false);
}

function found(v: cellValue): boolean {
    if (isError(v)) return false;
    return v != "";
}

class PriceRow {
    price: cellValue;
    constructor(r: int) { this.price = Sheet1.cells(r, 2).value; }
}

export function f(): void {
    if (lookup("a") == 1) {}                        // W
    if (found(lookup("b"))) {}
    const line = new PriceRow(2);
    const total = line.price * 2;                    // W
    const p = line.price;
    if (!isError(p)) { const t2 = p * 2; }
}
"#,
    );
}

#[test]
fn types_are_shown_and_narrowed() {
    let src = r#"
function lookup(key: string) {
    return application.vLookup(key, Sheet1.range("A1:B9"), 2, false);
}
export function f(): void {
    const v = Sheet1.range("A1").value;
    if (!isError(v)) {
        const w = v;
    }
    const rows = Excel.errorToNull(Sheet1.range("A1:B2").toArray());
}
"#;
    let (prog, out) = compile(src);
    let f = prog
        .fns
        .iter()
        .find(|d| d.name == "lookup" && prog.modules[d.module].kind == ModuleKind::User)
        .unwrap();
    assert_eq!(prog.show(&f.ret_ty()), "cellValue");
    let g = prog
        .fns
        .iter()
        .find(|d| d.name == "f" && prog.modules[d.module].kind == ModuleKind::User)
        .unwrap();
    let ty = |name: &str| {
        let l = g.locals.iter().find(|l| l.name == name).unwrap();
        prog.show(&l.ty)
    };
    assert_eq!(ty("v"), "cellValue");
    assert_eq!(ty("w"), "any");
    assert_eq!(ty("rows"), "any[][]");
    // In VBA it is a Variant like `any`.
    let m = out
        .files
        .iter()
        .find(|f| f.name == "M")
        .unwrap()
        .code
        .replace("\r\n", "\n");
    assert!(m.contains("    Dim v As Variant\n"), "{m}");
    assert!(
        m.contains("Private Function lookup(ByVal key As String) As Variant\n"),
        "{m}"
    );
}

#[test]
fn message_names_the_value_and_the_fix() {
    let ws =
        warnings("export function f(ws: Excel.Worksheet) { if (ws.range(\"A1\").value == 1) {} }");
    assert_eq!(ws.len(), 1, "{ws:#?}");
    let m = &ws[0].1;
    assert!(m.contains("`ws.range(…).value`"), "{m}");
    assert!(m.contains("`cellValue`"), "{m}");
    assert!(m.contains("比較する"), "{m}");
    assert!(m.contains("isError()"), "{m}");
    assert!(m.contains("Excel.errorToNull()"), "{m}");
}

#[test]
fn test_files_are_not_warned() {
    let src = "export function f(ws: Excel.Worksheet) { if (ws.range(\"A1\").value == 1) {} }";
    let (prog, _) = compile_str(&[("@test/a.test", src)], config(), &BuildOptions::default());
    assert!(
        !prog.diagnostics.iter().any(|d| d.code == "W0313"),
        "{:#?}",
        prog.diagnostics
    );
}

#[test]
fn error_to_null_is_emitted() {
    let (_, out) = compile(
        r#"
export function f(ws: Excel.Worksheet): any {
    const v = Excel.errorToNull(ws.range("A1").value) ?? "";
    const grid = Excel.errorToNull(ws.range("A1:C3").toArray());
    ws.range("D1").value = grid[0][0];
    ws.range("D2").value = v;
    const data = ws.range("A1:C3").toArray();
    const clean = Excel.errorToNull(data);
    ws.range("D3").value = clean[0][0];
    return Excel.errorToNull(data);
}
"#,
    );
    let m = out
        .files
        .iter()
        .find(|f| f.name == "M")
        .expect("module M")
        .code
        .replace("\r\n", "\n");
    // A value: the function, which copies.
    assert!(m.contains("RAS_ErrorToNull(ws.Range(\"A1\").Value)"), "{m}");
    assert!(m.contains("    f = RAS_ErrorToNull(data)\n"), "{m}");
    // Into a variable: the assignment already copies, the errors are cleared there.
    assert!(
        m.contains("    grid = RAS_Read2D(ws.Range(\"A1:C3\"))\n    RAS_ErrorToNullInPlace grid\n"),
        "{m}"
    );
    assert!(
        m.contains("    clean = data\n    RAS_ErrorToNullInPlace clean\n"),
        "{m}"
    );
    let rt = out
        .files
        .iter()
        .find(|f| f.name == "RAS_Runtime")
        .expect("runtime")
        .code
        .replace("\r\n", "\n");
    assert!(
        rt.contains("Public Function RAS_ErrorToNull(ByRef v As Variant) As Variant"),
        "{rt}"
    );
    assert!(
        rt.contains("Public Sub RAS_ErrorToNullInPlace(ByRef v As Variant)"),
        "{rt}"
    );
    // It walks the array with RAS_Len, which comes along.
    assert!(rt.contains("Public Function RAS_Len("), "{rt}");
}
