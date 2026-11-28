//! Accessors (`get` / `set`, spec §3.5): properties of modules, of sheets, of forms and of
//! classes, with and without parameters, and assignment through the properties with
//! parameters of host and late-bound objects.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{
    analyze, build, form_source, object_source, BuildOptions, Config, FormControl, FormDecl,
    ModuleKind, SourceInput,
};

/// The sources of `files`: `(key, text)` with the key `objects/<CodeName>`, `forms/<Name>` or
/// that of a module of `src/`.
fn sources(files: &[(&str, &str)]) -> Vec<SourceInput> {
    files
        .iter()
        .map(|(key, text)| {
            if let Some(name) = key.strip_prefix("objects/") {
                object_source(
                    PathBuf::from(format!("objects/{name}.ras")),
                    name,
                    text.to_string(),
                )
            } else if let Some(name) = key.strip_prefix("forms/") {
                form_source(
                    PathBuf::from(format!("forms/{name}.ras")),
                    name,
                    text.to_string(),
                )
            } else {
                SourceInput {
                    path: PathBuf::from(format!("src/{key}.ras")),
                    rel_path: format!("src/{key}.ras"),
                    key: key.to_string(),
                    kind: if key.ends_with(".d") {
                        ModuleKind::TypeDef
                    } else {
                        ModuleKind::User
                    },
                    text: text.to_string(),
                }
            }
        })
        .collect()
}

fn config() -> Config {
    let mut config = Config::default();
    config.sheet_code_names = vec!["Sheet1".into()];
    config.sheet_controls_known = true;
    config.forms = vec![FormDecl {
        name: "frmInput".into(),
        controls: vec![
            FormControl {
                name: "txtName".into(),
                control_type: "TextBox".into(),
            },
            FormControl {
                name: "lstItems".into(),
                control_type: "ListBox".into(),
            },
        ],
        resizable: false,
    }];
    config
}

/// Compiles `files`. Returns the diagnostics of the user's modules and the output.
fn compile(files: &[(&str, &str)]) -> (Vec<String>, Vec<ras_compiler::OutputFile>) {
    let prog = analyze(sources(files), config());
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
                .is_none_or(|m| prog.modules[m].kind != ModuleKind::Std)
        })
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    diags.sort();
    diags.dedup();
    (diags, out.files)
}

fn module(files: &[ras_compiler::OutputFile], name: &str) -> String {
    files
        .iter()
        .find(|f| f.name == name)
        .map(|f| f.code.replace("\r\n", "\n"))
        .unwrap_or_default()
}

fn assert_lines(code: &str, wants: &[&str]) {
    for want in wants {
        assert!(code.contains(want), "missing `{want}` in:\n{code}");
    }
}

const COUNTER: &str = r#"
let cache: int;
let label: string = "";
let sheet: Excel.Worksheet;

/** 合計。 */
export get total(): int {
    return cache;
}
/** 合計を設定する。 */
export set total(v: int) {
    cache = v;
}

get secret(): string {
    return "x" + label;
}

export set caption(v: string) {
    label = v;
}

export get target(): Excel.Worksheet {
    return sheet;
}
export set target(ws: Excel.Worksheet) {
    sheet = ws;
}

export function bump(): string {
    total = 3;
    total += 1;
    total++;
    caption = secret;
    return secret + total;
}
"#;

#[test]
fn accessors_of_a_standard_module_are_its_properties() {
    let (diags, files) = compile(&[
        ("counter", COUNTER),
        (
            "main",
            r#"
import { total, caption } from "src/counter";
import * as c from "src/counter";

export function run() {
    total = total + 1;
    total += 2;
    caption = "a";
    Counter.total = 5;
    c.total -= 1;
    Counter.target = Sheet1;
    debug.print(Counter.total, c.total, Counter.target.name);
}
"#,
        ),
    ]);
    assert!(diags.is_empty(), "{diags:#?}");
    let counter = module(&files, "Counter");
    assert_lines(
        &counter,
        &[
            // `export` → Public, otherwise Private; the description goes to the getter.
            "' 合計。\nPublic Property Get total() As Long\nAttribute total.VB_Description = \"合計。\"\n    total = cache\nEnd Property",
            "Public Property Let total(ByVal v As Long)\n    cache = v\nEnd Property",
            "Private Property Get secret() As String\n    secret = \"x\" & label\nEnd Property",
            "Public Property Let caption(ByVal v As String)",
            // An object type: `Property Set`, assigned with `Set`.
            "Public Property Get target() As Worksheet\n    Set target = sheet\nEnd Property",
            "Public Property Set target(ByVal ws As Worksheet)\n    Set sheet = ws\nEnd Property",
            // Used like variables inside the file.
            "    total = 3&\n    total = total + 1&\n    total = total + 1&\n    caption = secret\n    bump = secret & CStr(total)",
        ],
    );
    assert_eq!(counter.matches("VB_Description").count(), 1, "{counter}");
    assert_lines(
        &module(&files, "Main"),
        &[
            "    Counter.total = Counter.total + 1&\n    Counter.total = Counter.total + 2&\n    Counter.caption = \"a\"\n    Counter.total = 5&\n    Counter.total = Counter.total - 1&\n    Set Counter.target = Sheet1\n    Debug.Print Counter.total, Counter.total, Counter.target.Name",
        ],
    );
}

#[test]
fn a_getter_assigns_to_its_own_property_through_the_module() {
    // `total = 0` inside `Property Get total` would set the result.
    let src = "let n: int;\nexport get total(): int {\n    if (n < 0) {\n        total = 0;\n    }\n    return n;\n}\nexport set total(v: int) {\n    n = v;\n}\n";
    let (diags, files) = compile(&[("m", src), ("objects/Sheet1", src)]);
    assert!(diags.is_empty(), "{diags:#?}");
    assert_lines(&module(&files, "M"), &["        M.total = 0&\n"]);
    assert_lines(&module(&files, "Sheet1"), &["        Me.total = 0&\n"]);
}

const SHEET: &str = r#"
let n: int;
let labels: string[] = [];

export get count(): int {
    return n;
}
export set count(v: int) {
    n = v;
}

get hidden(): int {
    return n * 2;
}

export get item(i: int): string {
    return labels[i];
}
export set item(i: int, v: string) {
    labels[i] = v;
}

export function touch() {
    count = hidden + 1;
    this.count += 1;
    item(0) = "a";
    this.item(1) = this.item(0) + "b";
    const f = () => {
        count = 2;
        return item(0);
    };
    f();
}
"#;

const FORM: &str = r#"
let n: int;

export get count(): int {
    return n;
}
export set count(v: int) {
    n = v;
}

export get cell(r: int, c?: int): any {
    return n + r + c;
}
export set cell(r: int, c?: int, v: any) {
    n = r + c;
}

function load() {
    count = 1;
    this.count++;
    cell(1) = 2;
    this.cell(1, 2) = this.cell(3);
}
"#;

#[test]
fn accessors_of_sheets_and_forms_are_their_properties() {
    let (diags, files) = compile(&[
        ("objects/Sheet1", SHEET),
        ("forms/frmInput", FORM),
        (
            "main",
            r#"
function next(): int {
    return 1;
}

export function run() {
    Sheet1.count = Sheet1.count + 1;
    Sheet1.item(next()) += "z";
    frmInput.count = 3;
    const f = new frmInput();
    f.count += 1;
    f.cell(1, 2) = Sheet1.item(0);
    debug.print(f.cell(1), frmInput.cell(1, 2));
}
"#,
        ),
    ]);
    let errors: Vec<&String> = diags.iter().filter(|d| !d.starts_with("I0641")).collect();
    assert!(errors.is_empty(), "{diags:#?}");
    let sheet = module(&files, "Sheet1");
    assert_lines(
        &sheet,
        &[
            "Public Property Get count() As Long",
            "Public Property Let count(ByVal v As Long)",
            "Private Property Get hidden() As Long",
            "Public Property Get item(ByVal i As Long) As String\n    item = labels(i)\nEnd Property",
            "Public Property Let item(ByVal i As Long, ByVal v As String)\n    labels(i) = v\nEnd Property",
            // Bare and through `this`: the module's own members.
            "    count = hidden + 1&\n    count = count + 1&\n    item(0&) = \"a\"\n    item(1&) = item(0&) & \"b\"",
        ],
    );
    // The arrow function is another class: it reaches the sheet by its code name.
    let lambda = files
        .iter()
        .find(|f| f.name.starts_with("RAS_Lambda"))
        .map(|f| f.code.replace("\r\n", "\n"))
        .unwrap_or_default();
    assert_lines(
        &lambda,
        &["    Sheet1.count = 2&\n    Apply = Sheet1.item(0&)"],
    );
    let form = module(&files, "frmInput");
    assert_lines(
        &form,
        &[
            "Public Property Get count() As Long",
            "Public Property Let count(ByVal v As Long)",
            // Optional parameters stay optional before the value of `Property Let`; a `Variant`
            // is assigned with `Let` or `Set`.
            "Public Property Get cell(ByVal r As Long, Optional ByVal c As Long = 0&) As Variant",
            "Public Property Let cell(ByVal r As Long, Optional ByVal c As Long = 0&, ByVal v As Variant)",
            "Public Property Set cell(ByVal r As Long, Optional ByVal c As Long = 0&, ByVal v As Variant)",
            "    count = 1&\n    count = count + 1&\n    cell(1&, 0&) = 2&\n",
            "cell(3&, 0&)",
        ],
    );
    assert_lines(
        &module(&files, "Main"),
        &[
            "    Sheet1.count = Sheet1.count + 1&\n    ras_t1 = next_()\n    Sheet1.item(ras_t1) = Sheet1.item(ras_t1) & \"z\"\n    frmInput.count = 3&",
            "    f.count = f.count + 1&\n",
            "f.cell(1&, 2&) = ",
            "    Debug.Print f.cell(1&, 0&), frmInput.cell(1&, 2&)",
        ],
    );
}

#[test]
fn accessors_of_classes_take_parameters() {
    let (diags, files) = compile(&[(
        "main",
        r#"
class Bag {
    private d = new Map<string, int>();
    private names: string[] = [];
    get item(key: string): int {
        return this.d.get(key);
    }
    set item(key: string, value: int) {
        this.d.set(key, value);
    }
    get name(i: int = 0): string {
        return this.names[i];
    }
    set owner(i: int, ws: Excel.Worksheet) {
        this.names[i] = ws.name;
    }
    static get zero(base: int): int {
        return base * 0;
    }
    bump(key: string) {
        this.item(key) += 1;
    }
}

export function run() {
    const b = new Bag();
    b.item("k") = 1;
    b.item("k") += 2;
    b.item("k")++;
    b.owner(0) = Sheet1;
    debug.print(b.item("k"), b.name(), b.name(1), Bag.zero(3));
}
"#,
    )]);
    assert!(diags.is_empty(), "{diags:#?}");
    assert_lines(
        &module(&files, "Bag"),
        &[
            "Public Property Get item(ByVal key As String) As Long",
            "Public Property Let item(ByVal key As String, ByVal value_ As Long)",
            "Public Property Get name(Optional ByVal i As Long = 0&) As String",
            "Public Property Set owner(ByVal i As Long, ByVal ws As Worksheet)",
            "    item(key) = item(key) + 1&",
        ],
    );
    assert_lines(
        &module(&files, "Main"),
        &[
            "    b.item(\"k\") = 1&\n    b.item(\"k\") = b.item(\"k\") + 2&\n    b.item(\"k\") = b.item(\"k\") + 1&\n    Set b.owner(0&) = Sheet1\n    Debug.Print b.item(\"k\"), b.name(0&), b.name(1&), Main.Bag_zero(3&)",
            "Public Property Get Bag_zero(ByVal base As Long) As Long",
        ],
    );
}

#[test]
fn a_setter_of_any_takes_values_and_objects() {
    // The value decides between `Let` and `Set` (at run time for an `any`): both exist.
    let (diags, files) = compile(&[(
        "main",
        r#"
let slot: any;
export get held(): any {
    return slot;
}
export set held(v: any) {
    slot = v;
}
export function run(a: any) {
    held = 1;
    held = Sheet1;
    held = a;
}
"#,
    )]);
    assert!(diags.is_empty(), "{diags:#?}");
    assert_lines(
        &module(&files, "Main"),
        &[
            "Public Property Get held() As Variant",
            "Public Property Let held(ByVal v As Variant)",
            "Public Property Set held(ByVal v As Variant)",
            "    held = 1&\n    Set held = Sheet1\n",
            "    If IsObject(ras_t1) Then\n        Set held = ras_t1\n    Else\n        held = ras_t1\n    End If",
        ],
    );
}

#[test]
fn properties_with_arguments_of_host_and_late_bound_objects_are_assigned() {
    let (diags, files) = compile(&[
        (
            "forms/frmInput",
            "function load() {\n    this.lstItems.selected(0) = true;\n    this.lstItems.list(0, 1) = \"x\";\n    this.lstItems.list = [1, 2];\n    this.lstItems.selected(1) = !this.lstItems.selected(1);\n}\n",
        ),
        (
            "main",
            r#"
function next(): int {
    return 1;
}

export function run(o: object, a: any) {
    o.prop(1) = 5;
    o.prop(next(), 2) += 1;
    a.obj(1) = Sheet1;
    a(2) = "x";
    const ws = Sheet1;
    ws.cells(1, 2) = "x";
    ws.range("A1") = 3;
    ws.range("A1").offset(1, 0) = null;
    ws.range("A1:B2").cells(1, 1) = a;
    ws.range("A1")(1, 2) = 4;
    ws.range(1, 1, 2, 2) = 0;
    const d = new Scripting.Dictionary();
    d.item("a") = 1;
    d("b") = 2;
    d["c"] = 3;
}

export function add(ws: Excel.Worksheet) {
    ws.cells(1, 2) += 1.5;
    ws.cells(1, 2)++;
}
"#,
        ),
    ]);
    let rest: Vec<&String> = diags.iter().filter(|d| !d.starts_with("W0313")).collect();
    assert!(rest.is_empty(), "{diags:#?}");
    assert_lines(
        &module(&files, "frmInput"),
        &[
            "    Me.lstItems.Selected(0&) = True\n    Me.lstItems.List(0&, 1&) = \"x\"\n",
            "    Me.lstItems.List = ras_t1\n",
            "    Me.lstItems.Selected(1&) = Not Me.lstItems.Selected(1&)",
        ],
    );
    let main = module(&files, "Main");
    assert_lines(
        &main,
        &[
            // Late-bound: the property is found at run time; `Set` follows the value.
            "    o.prop(1&) = 5&\n    ras_t1 = next_()\n",
            "o.prop(ras_t1, 2&) + 1&",
            "    Set a.obj(1&) = Sheet1\n    a(2&) = \"x\"\n",
            // A cell range takes the value through its default member, always with `Let`.
            "    ws.Cells(1&, 2&) = \"x\"\n    ws.Range(\"A1\") = 3&\n    ws.Range(\"A1\").Offset(1&, 0&) = Empty\n    ws.Range(\"A1:B2\").Cells(1&, 1&) = a\n    ws.Range(\"A1\")(1&, 2&) = 4&\n    ws.Range(ws.Cells(1, 1), ws.Cells(2, 2)) = 0&\n",
            "    d.Item(\"a\") = 1&\n    d(\"b\") = 2&\n    d(\"c\") = 3&",
            "    ws.Cells(1&, 2&) = ws.Cells(1&, 2&) + 1.5#\n    ws.Cells(1&, 2&) = ws.Cells(1&, 2&) + 1&",
        ],
    );
}

#[test]
fn errors_of_accessors() {
    let errors = |src: &str| -> Vec<String> {
        let (diags, _) = compile(&[
            ("counter", COUNTER),
            ("objects/Sheet1", SHEET),
            ("main", src),
        ]);
        diags.into_iter().filter(|d| d.starts_with('E')).collect()
    };
    let has = |src: &str, code: &str, text: &str| {
        let e = errors(src);
        assert!(
            e.iter().any(|d| d.starts_with(code) && d.contains(text)),
            "{code} {text}: {e:#?}\n{src}"
        );
    };
    // Read-only and write-only properties.
    has(
        "get ro(): int { return 1; }\nexport function f() { ro = 2; }\n",
        "E0323",
        "`ro` には setter がありません",
    );
    has(
        "get ro(i: int): int { return i; }\nexport function f() { ro(1) = 2; }\n",
        "E0323",
        "`ro` には setter がありません",
    );
    has(
        "let n = 0;\nset wo(v: int) { n = v; }\nexport function f(): int { return wo; }\n",
        "E0326",
        "`wo` には getter がない",
    );
    has(
        "let n = 0;\nset wo(v: int) { n = v; }\nexport function f() { wo += 1; }\n",
        "E0326",
        "getter がない",
    );
    has(
        "let n = 0;\nset wo(i: int, v: int) { n = v; }\nexport function f(): int { return wo(1); }\n",
        "E0326",
        "`wo` には getter がない",
    );
    has(
        "export function f(): string { return Counter.caption; }\n",
        "E0326",
        "`caption` には getter がない",
    );
    // The shape of a setter, and the agreement of the two accessors.
    has("set s() {}\n", "E0155", "値を受け取る引数が必要");
    has(
        "let n = 0;\nset s(v?: int) { n = v; }\n",
        "E0155",
        "省略可能にできません",
    );
    has(
        "let n = 0;\nget p(): int { return n; }\nset p(v: string) { n = v.length; }\n",
        "E0155",
        "getter の型（`int`）と setter の値の型（`string`）",
    );
    has(
        "let n = 0;\nget p(i: int): int { return n + i; }\nset p(key: string, v: int) { n = v; }\n",
        "E0155",
        "引数の数・型",
    );
    has(
        "class C {\n    private n = 0;\n    get p(): int { return this.n; }\n    set p(i: int, v: int) { this.n = v; }\n}\n",
        "E0155",
        "引数の数・型",
    );
    has(
        "let n = 0;\nset p(byref v: int) { n = v; }\n",
        "E0155",
        "`byref`",
    );
    has("@macro\nget p(): int { return 1; }\n", "E0155", "`@macro`");
    // One getter and one setter; no function or variable of the same name.
    has(
        "get p(): int { return 1; }\nget p(): int { return 2; }\n",
        "E0101",
        "`p` は既に宣言されています",
    );
    has(
        "get p(): int { return 1; }\nfunction p(): int { return 2; }\n",
        "E0101",
        "`p` は既に宣言されています",
    );
    // A property with parameters is called; one without is not.
    has(
        "get p(i: int): int { return i; }\nexport function f(): int { return p; }\n",
        "E0382",
        "プロパティ `p` には引数が必要です",
    );
    has(
        "get p(i: int): int { return i; }\nexport function f(): int { return p(); }\n",
        "E0402",
        "`p` の引数の数",
    );
    has(
        "export function f() { Counter.total(1) = 2; }\n",
        "E0403",
        "呼び出せません",
    );
    // The arguments and the value are checked against the setter.
    has(
        "export function f() { Sheet1.item(\"a\") = \"b\"; }\n",
        "E0301",
        "",
    );
    has("export function f() { Sheet1.item(1) = 2; }\n", "E0301", "");
    // Only what is exported is a member of the sheet / of the module.
    has(
        "export function f(): int { return Sheet1.hidden; }\n",
        "E0202",
        "`hidden`",
    );
    has(
        "export function f(): string { return Counter.secret; }\n",
        "E0123",
        "`secret`",
    );
    // A method is no property: its result is not assigned to.
    has(
        "export function f() { Sheet1.range(\"A1\").address(true) = \"x\"; }\n",
        "E0325",
        "`address` は、型定義で引数のあるプロパティ（`@property`）とされていません",
    );
    has(
        "function g(): int { return 1; }\nexport function f() { g() = 2; }\n",
        "E0325",
        "ここには代入できません",
    );
    // A cell range takes a value, not an object (VBA would need `Set`).
    has(
        "export function f() { Sheet1.cells(1, 1) = Sheet1.range(\"B1\"); }\n",
        "E0325",
        "ここにオブジェクト（`Excel.Range`）は代入できません",
    );
    // No accessors in type definitions.
    let (diags, _) = compile(&[("lib.d", "get p(): int;\n")]);
    assert!(
        diags
            .iter()
            .any(|d| d.starts_with("E0155") && d.contains("型定義")),
        "{diags:#?}"
    );
}

#[test]
fn the_editor_treats_accessors_as_properties() {
    let text = "let n = 0;\n/** 件数。 */\nexport get count(): int {\n    return n;\n}\nexport set count(v: int) {\n    n = v;\n}\nexport get item(i: int): int {\n    return n + i;\n}\nexport function f(): int {\n    count = 1;\n    return count + item(2);\n}\n";
    let prog = analyze(sources(&[("m", text)]), Config::default());
    let errors: Vec<&str> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| d.message.as_str())
        .collect();
    assert!(errors.is_empty(), "{errors:#?}");
    let m = prog
        .modules
        .iter()
        .position(|m| m.kind == ModuleKind::User)
        .unwrap();
    let at = |needle: &str| text.find(needle).unwrap() as u32;
    // Hover on a use shows the property and its documentation.
    let hover = ras_compiler::query::hover(&prog, m, at("count + item") + 1).expect("hover");
    assert!(
        hover.markdown.contains("(プロパティ) count: int"),
        "{}",
        hover.markdown
    );
    assert!(hover.markdown.contains("件数。"), "{}", hover.markdown);
    // Go to definition leads to the getter.
    let (dm, span) = ras_compiler::query::definition(&prog, m, at("count = 1") + 1).unwrap();
    assert_eq!((dm, u32::from(span.start())), (m, at("count(): int")));
    // A rename changes the getter, the setter and the uses.
    let edits = ras_compiler::refs::rename(&prog, m, at("count = 1") + 1, "size").unwrap();
    assert_eq!(edits.len(), 4);
    // The other queries walk the new declarations without complaint.
    assert!(!ras_compiler::query::symbols(&prog, m).is_empty());
    let labels: Vec<String> = ras_compiler::query::completions(&prog, m, at("count = 1"))
        .into_iter()
        .map(|c| c.label)
        .collect();
    assert!(labels.contains(&"count".to_string()), "{labels:?}");
    assert!(labels.contains(&"item".to_string()), "{labels:?}");
    let _ = ras_compiler::hints::inlay_hints(&prog, m, 0, text.len() as u32);
    let _ = ras_compiler::hints::semantic_tokens(&prog, m);
    let _ = ras_compiler::query::signature_help(&prog, m, at("item(2)") + 5);
}
