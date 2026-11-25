//! Union types and narrowing (spec §2.6): typing, the generated VBA and the editor queries.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{
    analyze, compile_str, query, BuildOptions, BuildOutput, Config, ModuleKind, Program,
    SourceInput,
};

fn config() -> Config {
    let mut cfg = Config::default();
    cfg.doevents.enabled = false;
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

/// Code of a generated module with LF line endings and without the hash line.
fn code(out: &BuildOutput, module: &str) -> String {
    let f = out
        .files
        .iter()
        .find(|f| f.name == module)
        .unwrap_or_else(|| panic!("no module {module}"));
    f.code
        .replace("\r\n", "\n")
        .lines()
        .filter(|l| !l.starts_with("' RAS-GENERATED"))
        .map(|l| format!("{l}\n"))
        .collect()
}

/// (code, message) of the errors and warnings reported for the user module.
fn diagnostics(src: &str) -> Vec<(&'static str, String)> {
    let (prog, _) = compile_str(&[("m", src)], config(), &BuildOptions::default());
    prog.diagnostics
        .iter()
        .filter(|d| {
            d.module
                .is_some_and(|m| prog.modules[m].kind == ModuleKind::User)
                && d.code != "W0401"
                && d.code != "I0402"
        })
        .map(|d| (d.code, d.message.clone()))
        .collect()
}

#[track_caller]
fn assert_contains(haystack: &str, needle: &str) {
    assert!(
        haystack.contains(needle),
        "expected to find:\n{needle}\n--- in ---\n{haystack}"
    );
}

#[track_caller]
fn assert_reports(src: &str, code: &str, part: &str) {
    let ds = diagnostics(src);
    assert!(
        ds.iter().any(|(c, m)| *c == code && m.contains(part)),
        "expected {code} containing `{part}`, got {ds:#?}"
    );
}

fn analyze_one(src: &str) -> (Program, usize) {
    let input = SourceInput {
        path: PathBuf::from("src/m.ras"),
        rel_path: "src/m.ras".into(),
        key: "m".into(),
        kind: ModuleKind::User,
        text: src.into(),
    };
    let prog = analyze(vec![input], config());
    let m = prog.modules.iter().position(|md| md.key == "m").unwrap();
    (prog, m)
}

/// Hover text at the first occurrence of `marker` (the offset is just after it).
fn hover_at(src: &str, marker: &str) -> String {
    let (prog, m) = analyze_one(src);
    let offset = src.find(marker).unwrap_or_else(|| panic!("no {marker}")) + marker.len();
    query::hover(&prog, m, offset as u32)
        .map(|h| h.markdown)
        .unwrap_or_default()
}

const SPEC_EXAMPLE: &str = r#"
function describe(s: Excel.Worksheet | Excel.Chart): string {
    s.activate();
    if (s instanceof Excel.Worksheet) {
        return `${s.name}: ${s.usedRange.address()}`;
    }
    return s.chartTitle.text;
}

function cellText(v: string | double): string {
    if (typeof v == "String") return v;
    return v.toFixed(2);
}

export function main() {
    msgBox(describe(activeSheet) + cellText(1.5));
}
"#;

#[test]
fn spec_2_6_example_compiles_to_late_bound_vba() {
    let (_, out) = compile(SPEC_EXAMPLE);
    let m = code(&out, "M");
    assert_contains(
        &m,
        "Private Function describe(ByVal s As Object) As String\n    s.Activate\n    If (TypeOf s Is Worksheet) Then\n        describe = s.Name & \": \" & s.UsedRange.Address()\n        Exit Function\n    End If\n    describe = s.ChartTitle.Text\nEnd Function\n",
    );
    assert_contains(
        &m,
        "Private Function cellText(ByVal v As Variant) As String\n    If TypeName(v) = \"String\" Then\n        cellText = v\n        Exit Function\n    End If\n    cellText = RAS_ToFixed(v, 2&)\nEnd Function\n",
    );
}

#[test]
fn sheets_are_worksheets_or_charts() {
    let src = r#"
export function stamp() {
    const first = thisWorkbook.sheets(1);
    for (const sh of thisWorkbook.sheets) {
        if (sh instanceof Excel.Chart) {
            sh.chartTitle.text = sh.name;
        } else {
            sh.cells(1, 1).value = sh.name;
        }
    }
    const names = thisWorkbook.sheets.map(sh => sh.name);
}
"#;
    // `Sheets` without its type argument is `Sheets<Worksheet | Chart>` (spec §2.6.5).
    assert_eq!(
        hover_at(src, "const first"),
        "```ras\nfirst: Excel.Worksheet | Excel.Chart\n// VBA: Object\n```"
    );
    assert!(
        hover_at(src, "} else {\n            sh.cells(1, 1).value = sh.")
            .contains("Excel.Worksheet")
    );
    let (_, out) = compile(src);
    let m = code(&out, "M");
    assert_contains(&m, "    Dim first As Object\n    Dim sh As Object\n");
    assert_contains(&m, "    For Each sh In ThisWorkbook.Sheets\n        If (TypeOf sh Is Chart) Then\n            sh.ChartTitle.Text = sh.Name\n        Else\n            sh.Cells(1&, 1&).Value = sh.Name\n        End If\n    Next sh\n");
    // `worksheets` keeps its element type.
    let (prog, m) = analyze_one("export function f() { const w = worksheets(1); }");
    let w = prog
        .fns
        .iter()
        .find(|f| f.module == m && f.name == "f")
        .unwrap()
        .locals
        .iter()
        .find(|l| l.name == "w")
        .unwrap()
        .ty
        .clone();
    assert_eq!(prog.show(&w), "Excel.Worksheet");
}

#[test]
fn values_of_a_union_live_in_a_variant() {
    let src = r#"
export function store(): boolean {
    let v: string | double = 1;
    const same = v == "a";
    v = "a";
    const w: string | int = 3;
    let o: Excel.Range | string | null = null;
    o = "x";
    return same;
}
"#;
    let (_, out) = compile(src);
    let m = code(&out, "M");
    assert_contains(&m, "    Dim v As Variant\n    Dim same As Boolean\n    Dim w As Variant\n    Dim o As Variant\n");
    // A value that widens into a member is stored as that member, so that `typeof` sees it.
    assert_contains(&m, "    v = CDbl(1&)\n");
    assert_contains(&m, "    w = 3&\n");
    // Values of different kinds are not equal (spec §2.6.2).
    assert_contains(&m, "    same = RAS_UnionEq(v, \"a\")\n");
    // A union that may hold an object decides at run time.
    assert_contains(&m, "    Set o = Nothing\n    o = \"x\"\n");
    let rt = code(&out, "RAS_Runtime");
    assert_contains(&rt, "Public Function RAS_UnionEq(");
    assert_contains(&rt, "Public Function RAS_ValueKind(");
    // Error values (`#N/A` ...) are a kind of their own, compared by their numbers.
    assert_contains(&rt, "    Case vbError\n        RAS_ValueKind = 5&\n");
    assert_contains(&rt, "        RAS_UnionEq = (CLng(a) = CLng(b))\n");
}

#[test]
fn common_members_of_classes() {
    let src = r#"
class Circle {
    name: string = "circle";
    r: double = 1;
    area(): double { return 3 * this.r * this.r; }
}
class Square {
    name: string = "square";
    s: double = 2;
    area(): double { return this.s * this.s; }
}
type Shape = Circle | Square;

export function total(xs: Shape[]): double {
    let t = 0.0;
    for (const x of xs) {
        t += x.area();
        x.name = "seen";
    }
    return t;
}

export function radius(x: Shape): double {
    return x instanceof Circle ? x.r : 0.0;
}
"#;
    let (_, out) = compile(src);
    let m = code(&out, "M");
    assert_contains(
        &m,
        "Public Function total(ByRef xs() As Object) As Double\n",
    );
    assert_contains(&m, "        t = t + x.area()\n        x.name = \"seen\"\n");
    assert_contains(&m, "    If (TypeOf x Is Circle_) Then\n        radius = x.r\n    Else\n        radius = 0#\n    End If\n");
}

#[test]
fn narrowing() {
    let src = r#"
class A { va: int = 1; }
class B { vb: int = 2; }
class C { vc: int = 3; }

export function guards(x: A | B | C, v: string | int | boolean): int {
    if (x instanceof A || x instanceof B) {
        const ab = x;
    } else {
        const onlyC = x;
    }
    if (typeof v != "String" && typeof v != "Boolean") {
        const i = v;
    }
    if (!(x instanceof C)) {
        return 0;
    }
    const after = x;
    let y: A | B = new A();
    if (y instanceof A) {
        const notNarrowed = y;
        y = new B();
    }
    return after.vc;
}

export function fromAny(v: any): string {
    if (typeof v == "String") {
        const s = v;
        return s;
    }
    const o = activeSheet;
    if (o instanceof Excel.Worksheet) {
        const ws = o;
    }
    return "";
}
"#;
    let ty = |name: &str| {
        let (prog, m) = analyze_one(src);
        let f = prog
            .fns
            .iter()
            .filter(|f| f.module == m)
            .find(|f| f.locals.iter().any(|l| l.name == name))
            .unwrap();
        prog.show(&f.locals.iter().find(|l| l.name == name).unwrap().ty)
    };
    assert_eq!(ty("ab"), "A | B");
    assert_eq!(ty("onlyC"), "C");
    assert_eq!(ty("i"), "int");
    // `if (...) return;` narrows the rest of the block.
    assert_eq!(ty("after"), "C");
    // Assigned in the guarded block: not narrowed.
    assert_eq!(ty("notNarrowed"), "A | B");
    assert_eq!(ty("s"), "string");
    assert_eq!(ty("ws"), "Excel.Worksheet");
    let ds = diagnostics(src);
    assert!(ds.iter().all(|(c, _)| !c.starts_with('E')), "{ds:#?}");
    // The hover of a narrowed variable shows where it is stored.
    assert_contains(
        &hover_at(src, "const onlyC = x"),
        "x: C\n// VBA: Object（宣言は `A | B | C`。ここでは型が絞り込まれています）",
    );
}

#[test]
fn union_errors() {
    let classes = "class A { name: string = \"\"; readonly id: int = 0; go(n: int): int { return n; } }\nclass B { name: string = \"\"; id: int = 0; go(n: int, m: int): int { return n; } }\nclass C { name: int = 0; }\nstruct P { x: int; }\n";
    let with = |body: &str| format!("{classes}{body}");
    assert_reports(
        &with("export function f(x: A | C) { x.go(1); }"),
        "E0202",
        "`C` に `go` はありません",
    );
    assert_reports(
        "export function f(v: string | int) { const n = v.length; }",
        "E0202",
        "ユニオン型 `string | int` のメンバー `length`",
    );
    assert_reports(
        "export function f(v: string | int) { const n = v + 1; }",
        "E0387",
        "`string | int`",
    );
    assert_reports(
        "export function f(v: string | int) { const b = v < 3; }",
        "E0396",
        "大小比較",
    );
    assert_reports(
        &with("export function f(x: A | B) { x.go(1); }"),
        "E0402",
        "`B.go` の引数は 2 個です",
    );
    assert_reports(
        &with("export function f(x: A | B) { x.id = 3; }"),
        "E0323",
        "読み取り専用",
    );
    assert_reports(
        &with("export function f(x: A | C) { x.name = 1; }"),
        "E0325",
        "型ごとに異なる",
    );
    assert_reports(&with("function f(p: P | int) {}"), "E0170", "構造体 `P`");
    assert_reports("function f<T>(p: T | int) {}", "E0170", "型パラメータ `T`");
    assert_reports(
        "function f(p: string | int | null) {}",
        "E0170",
        "`string | int | null`",
    );
    assert_reports(
        "export function f() { const t = true ? 1 : \"a\"; }",
        "E0302",
        "`string | int`",
    );
    assert_reports(
        "export function f() { let w: string | int = 1.5; }",
        "E0301",
        "`double` を `string | int`",
    );
    assert_reports(
        "function g(byref p: string | int) {}\nexport function f(v: string | int) { if (typeof v == \"String\") { g(v); } }",
        "E0415",
        "`v`",
    );
    assert_reports(
        "export function f(v: string | int) { if (typeof v == \"string\") {} }",
        "W0312",
        "\"String\"",
    );
    // A host member that only some member types have is typed by those and late-bound, with a warning;
    // one that none has is late-bound as for a single host type.
    assert_reports(
        "export function f(s: Excel.Worksheet | Excel.Chart) { s.cells(1, 1).value = 1; }",
        "W0305",
        "`Excel.Chart`",
    );
    assert_reports(
        "export function f(s: Excel.Worksheet | Excel.Chart) { s.noSuchMember(); }",
        "W0301",
        "`noSuchMember`",
    );
    assert!(diagnostics("function ok(p: string | Excel.Range | null) {}").is_empty());
}

#[test]
fn type_parameter_defaults() {
    let src = "class Box<T = string | int> { v: T; constructor(v: T) { this.v = v; } }\nexport function f() { const b: Box = new Box<string | int>(1); }\n";
    assert!(
        diagnostics(src).iter().all(|(c, _)| !c.starts_with('E')),
        "{:#?}",
        diagnostics(src)
    );
    assert_reports(
        "function f<T = int>(x: T) {}",
        "E0171",
        "クラスとインターフェイス",
    );
    assert_reports("interface I<T = int, U> {}", "E0171", "`U`");
}

#[test]
fn editor_shows_common_members() {
    let src = "export function f(s: Excel.Worksheet | Excel.Chart) {\n    s.\n}\n";
    let (prog, m) = analyze_one(src);
    let offset = src.find("s.\n").unwrap() + 2;
    let labels: Vec<String> = query::completions(&prog, m, offset as u32)
        .into_iter()
        .map(|c| c.label)
        .collect();
    for want in ["name", "activate", "visible", "protect"] {
        assert!(
            labels.contains(&want.to_string()),
            "{want} missing: {labels:?}"
        );
    }
    for only_one in ["cells", "chartTitle"] {
        assert!(
            !labels.contains(&only_one.to_string()),
            "{only_one} is not common: {labels:?}"
        );
    }
}

#[test]
fn bundled_members_that_return_a_union() {
    let src = r#"
export function members(ws: Excel.Worksheet, ch: Excel.Chart, nm: Excel.Name, shp: Excel.Shape, win: Excel.Window) {
    const active = activeSheet;
    const bookActive = thisWorkbook.activeSheet;
    const next = ws.next;
    const shapeParent = shp.parent;
    const chartParent = ch.parent;
    const nameParent = nm.parent;
    const hit = win.rangeFromPoint(1, 1);
    const status = application.statusBar;
    const saveAs = application.getSaveAsFilename();
    const open = application.getOpenFilename();
    const added = thisWorkbook.sheets.add();
}
"#;
    let (prog, m) = analyze_one(src);
    let f = prog
        .fns
        .iter()
        .find(|f| f.module == m && f.name == "members")
        .unwrap();
    let ty = |name: &str| prog.show(&f.locals.iter().find(|l| l.name == name).unwrap().ty);
    for name in ["active", "bookActive", "next", "shapeParent", "added"] {
        assert_eq!(ty(name), "Excel.Worksheet | Excel.Chart", "{name}");
    }
    assert_eq!(ty("chartParent"), "Excel.Workbook | Excel.ChartObject");
    assert_eq!(ty("nameParent"), "Excel.Workbook | Excel.Worksheet");
    assert_eq!(ty("hit"), "Excel.Range | Excel.Shape");
    assert_eq!(ty("status"), "string | boolean");
    assert_eq!(ty("saveAs"), "string | boolean");
    assert_eq!(ty("open"), "string | boolean | any[]");
    // Members that are not one of a few known types stay as they were.
    let (prog, m) = analyze_one("export function f() { const s = selection; const v = sheet1Value(); }\nfunction sheet1Value(): any { return activeCell.value; }\n");
    let f = prog
        .fns
        .iter()
        .find(|f| f.module == m && f.name == "f")
        .unwrap();
    assert_eq!(
        prog.show(&f.locals.iter().find(|l| l.name == "s").unwrap().ty),
        "object"
    );
}

#[test]
fn code_written_for_object_still_compiles_with_warnings() {
    let src = r#"
export function legacy() {
    const ws: Excel.Worksheet = activeSheet;
    ws.cells(1, 1).value = 1;
    activeSheet.cells(2, 1).value = 2;
    const path: string = application.getSaveAsFilename();
}
"#;
    let ds = diagnostics(src);
    assert!(ds.iter().all(|(c, _)| !c.starts_with('E')), "{ds:#?}");
    assert!(
        ds.iter().any(|(c, m)| *c == "W0306"
            && m.contains("`Excel.Worksheet | Excel.Chart` を `Excel.Worksheet` として扱います")
            && m.contains("`Excel.Chart` のときは実行時エラー")),
        "{ds:#?}"
    );
    assert!(
        ds.iter().any(|(c, m)| *c == "W0306"
            && m.contains("`string | boolean` を `string`")
            && m.contains("暗黙の型変換")),
        "{ds:#?}"
    );
    assert!(
        ds.iter()
            .any(|(c, m)| *c == "W0305" && m.contains("`cells` は `Excel.Chart` にない")),
        "{ds:#?}"
    );
    // The member is typed by the member types that have it.
    assert_contains(&hover_at(src, "activeSheet.cells(2, 1).val"), "value");
    let (_, out) = compile(src);
    let m = code(&out, "M");
    assert_contains(&m, "    Set ws = ActiveSheet\n    ws.Cells(1&, 1&).Value = 1&\n    ActiveSheet.Cells(2&, 1&).Value = 2&\n    path_ = Application.GetSaveAsFilename()\n");

    // What the VBA importer writes for `Set ws = ActiveSheet` (spec §14.7.1) keeps compiling.
    let vba = "Option Explicit\nPublic Sub Test()\n    Dim ws As Worksheet\n    Set ws = ActiveSheet\n    Set ws = ws.Next\n    Application.StatusBar = False\nEnd Sub\n";
    let converted =
        ras_compiler::vba2ras::convert(vba, &ras_compiler::vba2ras::Target::Standard, &[]);
    let ds = diagnostics(&converted.text);
    assert!(
        ds.iter().all(|(c, _)| !c.starts_with('E')),
        "{}\n{ds:#?}",
        converted.text
    );
}

#[test]
fn get_open_filename_arrays_are_zero_based() {
    let src = r#"
export function pick(): int {
    const f = application.getOpenFilename("Excel,*.xlsx", 1, "選択", "開く", true);
    if (typeof f == "Boolean") return 0;
    if (typeof f == "Variant()") {
        for (const p of f) workbooks.open(p);
        return f.length;
    }
    workbooks.open(f);
    return 1;
}
"#;
    let ds = diagnostics(src);
    assert!(ds.is_empty(), "{ds:#?}");
    let (_, out) = compile(src);
    let m = code(&out, "M");
    assert_contains(&m, "    f = RAS_Rebase1(Application.GetOpenFilename(\"Excel,*.xlsx\", 1&, \"選択\", \"開く\", True))\n    If TypeName(f) = \"Boolean\" Then\n");
    assert_contains(&m, "        For ras_i = 0 To RAS_Len(f) - 1&\n");
    assert_contains(
        &code(&out, "RAS_Runtime"),
        "Public Function RAS_Rebase1(ByRef v As Variant) As Variant\n",
    );
}
