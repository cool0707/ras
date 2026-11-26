//! The JavaScript back end (spec §14.10): the generated code must run under Node with the
//! semantics the VBA build has, and its source maps must lead back to the RAS source.
//!
//! The tests that execute code need `node` on PATH and are skipped without it.

#![allow(clippy::field_reassign_with_default)]

use std::path::{Path, PathBuf};
use std::process::Command;

use ras_compiler::{
    analyze, build_js, form_source, Config, FormDecl, JsOutput, ModuleKind, Program, SourceInput,
};

fn compile(files: &[(&str, &str)]) -> (Program, JsOutput) {
    let mut cfg = Config::default();
    cfg.sheet_code_names = vec!["Sheet1".into()];
    let sources = files
        .iter()
        .map(|(key, text)| SourceInput {
            path: PathBuf::from(format!("src/{key}.ras")),
            rel_path: format!("src/{key}.ras"),
            key: (*key).to_string(),
            kind: ModuleKind::User,
            text: (*text).to_string(),
        })
        .collect();
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    (prog, out)
}

fn has_node() -> bool {
    let found = Command::new("node")
        .arg("--version")
        .output()
        .is_ok_and(|o| o.status.success());
    if !found {
        eprintln!("skipped: node is not installed");
    }
    found
}

/// A project directory with the sources in `src/` and the debug build in `build/debug/`.
struct Project(PathBuf);

impl Project {
    fn new(name: &str, files: &[(&str, &str)], out: &JsOutput) -> Self {
        let root = std::env::temp_dir().join(format!("ras-js-{}-{name}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        for (key, text) in files {
            let p = root.join("src").join(format!("{key}.ras"));
            std::fs::create_dir_all(p.parent().unwrap()).unwrap();
            std::fs::write(p, text).unwrap();
        }
        for f in &out.files {
            let p = root.join("build").join("debug").join(&f.rel_path);
            std::fs::create_dir_all(p.parent().unwrap()).unwrap();
            std::fs::write(p, &f.text).unwrap();
        }
        Project(root)
    }

    /// Runs an entry point; returns (stdout, stderr, success).
    fn run(&self, entry: &str, mock: Option<&Path>) -> (String, String, bool) {
        let mut cmd = Command::new("node");
        cmd.arg("--enable-source-maps")
            .arg(self.0.join("build").join("debug").join("ras-main.js"))
            .arg(entry)
            .current_dir(&self.0);
        cmd.env_remove("RAS_COM_ADDON").env_remove("RAS_MOCK_HOST");
        if let Some(m) = mock {
            cmd.env("RAS_MOCK_HOST", m);
        }
        let o = cmd.output().expect("node");
        (
            String::from_utf8_lossy(&o.stdout).replace("\r\n", "\n"),
            String::from_utf8_lossy(&o.stderr).into_owned(),
            o.status.success(),
        )
    }
}

impl Drop for Project {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("js")
        .join(name)
}

#[test]
fn language_semantics_match_the_vba_build() {
    let files = [
        ("main", include_str!("js/semantics.ras")),
        ("lib", include_str!("js/lib.ras")),
    ];
    let (_, out) = compile(&files);
    assert_eq!(out.entries, vec!["Main.run"]);
    if !has_node() {
        return;
    }
    let p = Project::new("semantics", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        include_str!("js/semantics.out").replace("\r\n", "\n")
    );
}

const REPORT: &str = r#"import { lastRow } from "./sheetutil";
const TAX_RATE = 0.1;

@fast
export function rebuildReport() {
    const ws = Sheet1;
    const last = ws.lastRow();
    let total = 0.0;
    for (let r = 2; r <= last; r++) {
        const price: double = ws.cells(r, 3).value;
        const qty: double = ws.cells(r, 4).value;
        total += price * qty;
    }
    for (let r = 2; r <= last; r++) {
        ws.cells(r, 5).formula = `=ROUND(C${r}*D${r}*${1 + TAX_RATE},0)`;
    }
    status(`合計: ${total.toFixed(0)} 円`);
}

export function countCustomers(): int {
    const names = Sheet1.range("A2:A1000").toColumn();
    const seen = new Map<string, boolean>();
    for (const n of names) {
        const name = `${n}`;
        if (name.trim().length > 0) seen.set(name, true);
    }
    console.log(`customers=${seen.size} sheet=${Sheet1.name}`);
    return seen.size;
}

export function listSheets() {
    // After `filter`, the index counts the members that passed (the hidden sheet is skipped).
    const visible = worksheets.filter(ws => ws.visible == xlSheetVisible).map((ws, i) => `${i}:${ws.name}`);
    worksheets.forEach((ws, i) => { if (i == 2) console.log(i, ws.name); });
    const hidden = worksheets.find(ws => ws.visible != xlSheetVisible);
    // A late-bound member is found whatever its casing, as in VBA.
    const first: any = worksheets.toArray()[0];
    console.log(visible.join(","), hidden.name, worksheets.count, worksheets.some(ws => ws.name == "Summary"), first.name);
}
"#;

const SHEETUTIL: &str = "export function lastRow(this: Excel.Worksheet, col: int = 1): int {\n    return this.cells(this.rows.count, col).end(xlUp).row;\n}\n";

#[test]
fn host_access_goes_through_the_run_time() {
    let files = [("report", REPORT), ("sheetutil", SHEETUTIL)];
    let (_, out) = compile(&files);
    // Macro entries first, then every other function that can be started without arguments.
    assert_eq!(
        out.entries,
        vec![
            "Report.rebuildReport",
            "Report.listSheets",
            "Report.countCustomers"
        ]
    );
    let js = &out
        .files
        .iter()
        .find(|f| f.rel_path == "report.js")
        .unwrap()
        .text;
    assert!(
        js.contains("$rt.cnum($rt.get($rt.call(ws, \"Cells\", [r, 3]), \"Value\"))"),
        "{js}"
    );
    assert!(
        js.contains("$rt.put($rt.call(ws, \"Cells\", [r, 5]), \"Formula\","),
        "{js}"
    );
    assert!(js.contains("$m_sheetutil.lastRow(ws)"), "{js}");
    if !has_node() {
        return;
    }
    let p = Project::new("host", &files, &out);
    let mock = fixture("mock_excel.js");
    let (stdout, stderr, ok) = p.run("Report.rebuildReport", Some(&mock));
    assert!(ok, "{stderr}");
    // Empty cells read as 0 and numeric strings convert, as they do when VBA assigns a Variant to a Double.
    assert_eq!(stdout, "E2==ROUND(C2*D2*1.1,0)\nE5==ROUND(C5*D5*1.1,0)\nScreenUpdating=false StatusBar=合計: 1292 円 StatusBar=false ScreenUpdating=true\n");
    // The entry can be named without its module when that is unambiguous.
    let (stdout, stderr, ok) = p.run("countCustomers", Some(&mock));
    assert!(ok, "{stderr}");
    assert!(stdout.starts_with("customers=2 sheet=Data\n"), "{stdout}");
    // Array-style methods of a host collection (spec §4.4).
    assert!(
        js.contains("Array.from($rt.iter($rt.global(\"Worksheets\"))).filter("),
        "{js}"
    );
    let (stdout, stderr, ok) = p.run("listSheets", Some(&mock));
    assert!(ok, "{stderr}");
    assert!(
        stdout.starts_with("2 Summary\n0:Data,1:Summary Hidden 3 true Data\n"),
        "{stdout}"
    );
    // Without Excel, the first host access fails with an explanation and a RAS source position.
    let (_, stderr, ok) = p.run("Report.rebuildReport", None);
    assert!(!ok);
    assert!(
        stderr.contains("ras_com") && stderr.contains("report.ras:6"),
        "{stderr}"
    );
}

#[test]
fn source_maps_lead_back_to_ras_lines() {
    let src = "function f(z: int): int {\n    return 10 / z;\n}\nexport function run() {\n    try {\n        f(0);\n    } catch (e) {\n        console.log(e.code, e.rasLine, e.rasFile.endsWith(\"m.ras\"));\n    }\n}\n";
    let files = [("m", src)];
    let (_, out) = compile(&files);
    let map: serde_json::Value = serde_json::from_str(
        &out.files
            .iter()
            .find(|f| f.rel_path == "m.js.map")
            .unwrap()
            .text,
    )
    .unwrap();
    assert_eq!(map["version"], 3);
    assert_eq!(map["sources"][0], "../../src/m.ras");
    assert!(map["mappings"].as_str().unwrap().len() > 10);
    let js = &out
        .files
        .iter()
        .find(|f| f.rel_path == "m.js")
        .unwrap()
        .text;
    assert!(js.ends_with("//# sourceMappingURL=m.js.map\n"));
    // The project's files are read from the disk; the standard library's text is in its map.
    assert!(map.get("sourcesContent").is_none());
    let lib = out
        .files
        .iter()
        .find(|f| f.rel_path == "ras_collections.js.map")
        .expect("the standard library's module");
    let lib: serde_json::Value = serde_json::from_str(&lib.text).unwrap();
    assert!(lib["sources"][0]
        .as_str()
        .unwrap()
        .ends_with("<stdlib>/ras_collections.ras"));
    assert!(lib["sourcesContent"][0]
        .as_str()
        .unwrap()
        .contains("export class"));
    assert!(
        !js.contains('\u{1}') && !js.contains('\u{2}'),
        "position markers must be stripped"
    );
    if !has_node() {
        return;
    }
    let p = Project::new("maps", &files, &out);
    let (stdout, stderr, ok) = p.run("run", None);
    assert!(ok, "{stderr}");
    assert_eq!(stdout, "11 2 true\n");
}

#[test]
fn debugger_statements_pause_the_debug_run() {
    let src = "export function run() {\n    let n = 1;\n    debugger;\n    try { debugger } catch { resume; }\n    console.log(n);\n}\n";
    let files = [("m", src)];
    let (_, out) = compile(&files);
    let js = &out
        .files
        .iter()
        .find(|f| f.rel_path == "m.js")
        .unwrap()
        .text;
    assert_eq!(js.matches("debugger;").count(), 2, "{js}");
    if !has_node() {
        return;
    }
    // Without a debugger attached, the statement does nothing.
    let p = Project::new("debugger", &files, &out);
    let (stdout, stderr, ok) = p.run("run", None);
    assert!(ok, "{stderr}");
    assert_eq!(stdout, "1\n");
}

#[test]
fn uncaught_errors_show_only_ras_frames() {
    let src = "function f(z: int): int {\n    return 10 / z;\n}\nfunction g(): int {\n    return f(0);\n}\nexport function run() {\n    g();\n}\n";
    let files = [("m", src)];
    let (_, out) = compile(&files);
    if !has_node() {
        return;
    }
    let p = Project::new("uncaught", &files, &out);
    let (_, stderr, ok) = p.run("run", None);
    assert!(!ok);
    let frames: Vec<&str> = stderr
        .lines()
        .filter(|l| l.trim_start().starts_with("at "))
        .collect();
    assert_eq!(
        frames,
        [
            "    at f (src/m.ras:2:5)",
            "    at g (src/m.ras:5:5)",
            "    at run (src/m.ras:8:5) {"
        ],
        "{stderr}"
    );
    assert!(
        stderr.contains("rasFile: 'src/m.ras'") && stderr.contains("rasLine: 2"),
        "{stderr}"
    );
    assert!(
        !stderr.contains("ras-runtime.js") && !stderr.contains("node:internal"),
        "{stderr}"
    );
}

#[test]
fn modules_in_subdirectories_and_import_cycles() {
    let a = "import { pong } from \"./util/b\";\nexport let hits = 0;\nexport function ping(n: int): int {\n    hits++;\n    return n <= 0 ? 0 : pong(n - 1) + 1;\n}\nexport function run() {\n    console.log(ping(5), hits);\n}\n";
    let b = "import { ping } from \"../a\";\nexport function pong(n: int): int {\n    return n <= 0 ? 0 : ping(n - 1) + 1;\n}\n";
    let files = [("a", a), ("util/b", b)];
    let (_, out) = compile(&files);
    let map: serde_json::Value = serde_json::from_str(
        &out.files
            .iter()
            .find(|f| f.rel_path == "util/b.js.map")
            .unwrap()
            .text,
    )
    .unwrap();
    assert_eq!(map["sources"][0], "../../../src/util/b.ras");
    if !has_node() {
        return;
    }
    let p = Project::new("cycle", &files, &out);
    let (stdout, stderr, ok) = p.run("A.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(stdout, "5 3\n");
}

#[test]
fn unsupported_constructs_are_reported() {
    let src = "declare module \"user32\" {\n    function GetForegroundWindow(): ptr;\n}\nexport function run() {\n    const h = GetForegroundWindow();\n}\n";
    let (_, out) = compile(&[("m", src)]);
    let codes: Vec<&str> = out.diagnostics.iter().map(|d| d.code).collect();
    assert_eq!(codes, vec!["W1410"]);
}

const FORM_MAIN: &str = r#"export function run() {
    frmInput.show();
    console.log(frmInput.summary());
    unload(frmInput);
    frmInput.show(vbModeless);
    console.log(`end ${frmInput.caption}`);
}
"#;

/// A form's code-behind runs in the debug build; the form is drawn by a host (the VS Code
/// extension), here a test double that plays a user (spec §14.10.7).
#[test]
fn forms_run_with_a_host() {
    let spec = include_str!("js/forms/frmInput.form.json");
    let code = include_str!("js/forms/frmInput.ras");
    let mut cfg = Config::default();
    cfg.forms = vec![FormDecl::from_form_json(spec).unwrap()];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: FORM_MAIN.into(),
        },
        form_source(
            PathBuf::from("forms/frmInput.ras"),
            "frmInput",
            code.to_string(),
        ),
    ];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    assert_eq!(out.entries, vec!["Main.run", "frmInput.show"]);
    let js = &out
        .files
        .iter()
        .find(|f| f.rel_path == "@forms/frmInput.js")
        .expect("the form's module")
        .text;
    assert!(js.contains("$rt.form(\"frmInput\", {"), "{js}");
    assert!(js.contains("$rt.put($me, \"Caption\","), "{js}");
    assert!(
        js.contains("fn: ($a) => closing($a[0], $a[1].v)"),
        "byref parameters get the cell: {js}"
    );
    // Handlers attached at run time (spec §12.2.3): the form is their owner.
    assert!(
        js.contains("$rt.listen($rt.get($me, \"chkAgree\"), \"Click\", "),
        "{js}"
    );
    assert!(
        js.contains(
            "firstSpin = $rt.listen($rt.get($me, \"spnQty\"), \"SpinUp\", spunOnce, $me, -1)"
        ),
        "{js}"
    );
    let manifest = &out
        .files
        .iter()
        .find(|f| f.rel_path == "ras-debug.json")
        .unwrap()
        .text;
    assert!(
        manifest.contains("\"forms\": [\n    \"frmInput\"\n  ]"),
        "{manifest}"
    );
    if !has_node() {
        return;
    }
    let root = std::env::temp_dir().join(format!("ras-js-{}-forms", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    let write = |rel: &str, text: &str| {
        let p = root.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    };
    write("src/main.ras", FORM_MAIN);
    write("forms/frmInput.ras", code);
    write("forms/frmInput.form.json", spec);
    for f in &out.files {
        write(&format!("build/debug/{}", f.rel_path), &f.text);
    }
    let p = Project(root);
    let (stdout, stderr, ok) = p.run("Main.run", Some(&fixture("mock_forms.js")));
    assert!(ok, "{stderr}");
    let expected = "\
show frmInput modal=true \"入力 1\" txtName:TextBox cboKind:ComboBox chkAgree:CheckBox chkMaybe:CheckBox fraSize:Frame fraSize/optS:OptionButton fraSize/optM:OptionButton spnQty:SpinButton lstStaff:ListBox btnOK:CommandButton btnBoom:CommandButton
draw txtName= seq=undefined
focus txtName
draw txtName=NG seq=1
focus txtName
draw txtName=NG seq=1
focus txtName
draw txtName=YAMADA seq=2
focus txtName
focus cboKind
agree false 1
focus chkAgree
maybe=Null
focus chkMaybe
maybe=True
focus chkMaybe
maybe=False
focus chkMaybe
optS=false optM=true
focus optM
up 2
first spin
spin spnQty
focus spnQty
up 3
spin spnQty
focus spnQty
up 3
spin spnQty
error frmInput.btnBoom_Click: 実行時エラー 9: インデックスが有効範囲にありません。 (3)（forms/frmInput.ras:56）
clicked btnBoom
focus btnBoom
QueryClose mode=0
focus btnBoom
hide
clicked btnOK
YAMADA/予約:1/3/false/3/1/002:佐藤:山田/入力用:120
QueryClose mode=1
Terminate
show frmInput modal=false \"入力 1\" txtName:TextBox cboKind:ComboBox chkAgree:CheckBox chkMaybe:CheckBox fraSize:Frame fraSize/optS:OptionButton fraSize/optM:OptionButton spnQty:SpinButton lstStaff:ListBox btnOK:CommandButton btnBoom:CommandButton
draw txtName= seq=undefined
focus txtName
end 入力 1
QueryClose mode=0
focus txtName
Terminate
B2=false
";
    assert_eq!(stdout, expected, "{stderr}");
    assert!(
        stderr.contains("[ras] frmInput.btnBoom_Click: 実行時エラー 9"),
        "{stderr}"
    );
    // Without a host (outside VS Code), showing a form explains where it can be shown.
    let (_, stderr, ok) = p.run("frmInput.show", None);
    assert!(!ok);
    assert!(stderr.contains("VS Code"), "{stderr}");
}

const UNIONS: &str = r#"
class Circle {
    r: double = 1;
    area(): double { return 3 * this.r * this.r; }
    label(): string { return "circle"; }
}
class Square {
    s: double = 2;
    area(): double { return this.s * this.s; }
    label(): string { return "square"; }
}

function describe(x: Circle | Square): string {
    if (x instanceof Circle) {
        return `${x.label()} r=${x.r} a=${x.area()}`;
    }
    return `${x.label()} s=${x.s} a=${x.area()}`;
}

function kind(v: string | int | boolean): string {
    if (typeof v == "String") return `s:${v.length}`;
    if (typeof v == "Long") return `i:${v + 1}`;
    return `b:${v}`;
}

export function run() {
    const shapes: (Circle | Square)[] = [new Circle(), new Square()];
    for (const s of shapes) console.log(describe(s), s.area());
    let v: string | int | boolean = "abc";
    console.log(kind(v), v == "abc", v == 3);
    v = 7;
    console.log(kind(v), v == 7, v == "7");
    v = true;
    console.log(kind(v));
    const w: string | double = 4;
    console.log(typeof w);
}
"#;

#[test]
fn union_types_narrow_like_the_vba_build() {
    let files = [("main", UNIONS)];
    let (_, out) = compile(&files);
    let js = &out
        .files
        .iter()
        .find(|f| f.rel_path == "main.js")
        .unwrap()
        .text;
    // Values of different kinds are not equal; a number's `TypeName` follows the static type when it can.
    assert!(js.contains("$rt.unionEq(v, \"abc\")"), "{js}");
    assert!(js.contains("$rt.typeName(w, \"Double\")"), "{js}");
    if !has_node() {
        return;
    }
    let p = Project::new("unions", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        "circle r=1 a=3 3\nsquare s=2 a=4 4\ns:3 true false\ni:8 true false\nb:true\nDouble\n"
    );
}

const SCRIPTING: &str = r#"
export function files() {
    const fso = new Scripting.FileSystemObject();
    if (fso.folderExists("work")) fso.deleteFolder("work");
    const work = fso.createFolder("work");
    const out = fso.createTextFile(fso.buildPath(work.path, "名簿.txt"));
    out.writeLine("佐藤,41");
    out.write("鈴木,");
    out.writeLine("29");
    out.close();
    const f = fso.getFile("work\\名簿.txt");
    console.log(`${f.name} ${f.size} ${fso.getBaseName(f.path)} ${fso.getExtensionName(f.path)} ${f.parentFolder.name} ${typeof f}`);
    const ts = fso.openTextFile(f.path, forReading);
    while (!ts.atEndOfStream) {
        const at = ts.line;
        console.log(`${at}:${ts.readLine()}`);
    }
    ts.close();
    const more = f.openAsTextStream(forAppending);
    more.writeLine("高橋,35");
    more.close();
    console.log(`${fso.openTextFile(f.path).readAll().length}`);
    // Unicode: UTF-16 with a byte order mark.
    const u = fso.createTextFile("work/u.txt", true, true);
    u.write("あ");
    u.close();
    console.log(`${fso.getFile("work/u.txt").size} ${fso.openTextFile("work/u.txt", forReading, false, tristateTrue).readAll()}`);
    // Copies with wildcards, the collections of a folder, moves and deletes.
    work.subFolders.add("copy");
    fso.copyFile("work\\*.txt", "work\\copy\\");
    let names = "";
    for (const file of fso.getFolder("work/copy").files) names += `${file.name};`;
    fso.moveFile("work/copy/u.txt", "work/v.txt");
    fso.deleteFile("work/copy/*.txt");
    console.log(`${names} ${work.files.count} ${work.subFolders.count} ${work.subFolders("copy").files.count} ${fso.fileExists("work/v.txt")}`);
    console.log(`${fso.getParentFolderName("C:\\a\\b.txt")} ${fso.getFileName("C:\\a\\b.txt")} ${fso.getBaseName("a/b.tar.gz")} ${fso.getDriveName("C:\\a")}`);
    try {
        fso.getFile("work/none.txt");
    } catch (e) {
        console.log(`${e.code}`);
    }
    try {
        fso.createFolder("work");
    } catch (e) {
        console.log(`${e.code}`);
    }
    work.delete();
    console.log(`${fso.folderExists("work")}`);
}

export function dictionary() {
    const d = createObject("Scripting.Dictionary");
    d.add("b", 2);
    d.add("a", 1);
    d["c"] = 3;
    d["a"] = 10;
    console.log(`${d.count} ${d.exists("a")} ${d.exists("A")} ${d("a")} ${d.item("b")} ${d.keys().join(",")} ${d.items().join(",")}`);
    let keys = "";
    for (const k of d) keys += `${k}`;
    try {
        d.add("a", 0);
    } catch (e) {
        console.log(`${keys} ${e.code}`);
    }
    d.remove("b");
    const missing = d("zz");
    console.log(`${d.count} ${typeof missing}`);
    const t = new Scripting.Dictionary();
    t.compareMode = vbTextCompare;
    t.add("Key", 1);
    console.log(`${t.exists("KEY")} ${typeof t}`);
}

export function regex() {
    const re = new VBScript.RegExp();
    re.pattern = "(\\d{4})/(\\d{1,2})";
    re.global = true;
    const text = "2026/3 と 2025/12";
    const found = re.execute(text);
    console.log(`${re.test(text)} ${found.count} ${found(1).value} ${found(1).firstIndex} ${found(0).subMatches(1)} ${found.item(0).length}`);
    let years = "";
    for (const m of found) years += `${m.subMatches(0)};`;
    console.log(`${years} ${re.replace(text, "$2月")}`);
    re.global = false;
    console.log(`${re.replace(text, "$1年")} ${re.execute(text).count}`);
    re.pattern = "^abc";
    re.ignoreCase = true;
    console.log(`${re.test("ABCdef")}`);
    re.pattern = "(";
    try {
        re.test("x");
    } catch (e) {
        console.log(`${e.code}`);
    }
    try {
        createObject("ADODB.Connection");
    } catch (e) {
        console.log(`${e.code}`);
    }
}
"#;

#[test]
fn scripting_objects_run_without_com() {
    let files = [("tools", SCRIPTING)];
    let (prog, out) = compile(&files);
    // The VBA build creates them by their ProgID, late-bound: no reference to their libraries.
    let vba = ras_compiler::build(&prog, &ras_compiler::BuildOptions::default());
    let bas = &vba
        .files
        .iter()
        .find(|f| f.file_name == "Tools.bas")
        .expect("Tools.bas")
        .code;
    for want in [
        "Dim fso As Object",
        "Set fso = CreateObject(\"Scripting.FileSystemObject\")",
        "Set ts = fso.OpenTextFile(f.Path, 1&)",
        "Set d = CreateObject(\"Scripting.Dictionary\")",
        "Set re = CreateObject(\"VBScript.RegExp\")",
        "re.Global = True",
    ] {
        assert!(bas.contains(want), "{want}\n{bas}");
    }
    if !has_node() {
        return;
    }
    let p = Project::new("scripting", &files, &out);
    let (stdout, stderr, ok) = p.run("Tools.files", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout.lines().collect::<Vec<_>>(),
        vec![
            "名簿.txt 18 名簿 txt work File",
            "1:佐藤,41",
            "2:鈴木,29",
            "21",
            "4 あ",
            "u.txt;名簿.txt; 3 1 0 true",
            "C:\\a b.txt b.tar C:",
            "53",
            "58",
            "false",
        ]
    );
    let (stdout, stderr, ok) = p.run("Tools.dictionary", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout.lines().collect::<Vec<_>>(),
        vec![
            "3 true false 10 2 b,a,c 2,10,3",
            "bac 457",
            "3 Empty",
            "true Dictionary"
        ]
    );
    let (stdout, stderr, ok) = p.run("Tools.regex", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout.lines().collect::<Vec<_>>(),
        vec![
            "true 2 2025/12 9 3 6",
            "2026;2025; 3月 と 12月",
            "2026年 と 2025/12 1",
            "true",
            "5020",
            "429"
        ]
    );
}

/// A function of a module named as one of VBA's would hide it for the whole project:
/// `RAS_Runtime` calls `Left$`, and Excel then stops at a compile error (seen with Excel 2021).
#[test]
fn a_function_named_as_one_of_vba_is_renamed_in_the_vba_build() {
    let source = "let trim = 0;\n\n@macro\nfunction left() {\n    trim += 1;\n}\n\nexport function format(): int {\n    left();\n    return trim;\n}\n";
    let (prog, _) = compile(&[("main", source)]);
    let vba = ras_compiler::build(&prog, &ras_compiler::BuildOptions::default());
    let all: String = vba
        .files
        .iter()
        .map(|f| f.code.as_str())
        .collect::<Vec<_>>()
        .join("\n");
    let main = &vba
        .files
        .iter()
        .find(|f| f.file_name == "Main.bas")
        .expect("Main.bas")
        .code;
    for want in ["Sub left_()", "Function format_() As Long", "trim_ As Long"] {
        assert!(all.contains(want), "{want}\n{all}");
    }
    for bad in ["Sub left()", "Function format()", " trim As Long"] {
        assert!(!main.contains(bad), "{bad}\n{main}");
    }
}

const FILE_IO: &str = r#"
class Rec { n: int = 0; }
export function run() {
    if (dir("work", vbDirectory) != "") {
        kill("work\\*");
        rmDir("work");
    }
    mkDir("work");
    const path = "work\\out.txt";
    const f = freeFile();
    fileOpen(f, path, forOutput);
    printLine(f, "名前", 1, 2.5, true);
    print(f, "no newline");
    printLine(f);
    writeLine(f, "x,y", 3, false, -1.5);
    write(f, "y");
    writeLine(f, 7);
    fileClose(f);
    console.log(`len=${fileLen(path)} free=${freeFile()}`);
    fileOpen(1, path, forInput, accessRead, lockShared);
    for (let i = 0; i < 2; i++) console.log(`[${lineInput(1)}]`);
    let s = "";
    let n = 0;
    let b = true;
    let d = 0.0;
    const r = new Rec();
    input(1, s, n, b, d);
    console.log(`${s}|${n}|${b}|${d}`);
    input(1, s, r.n);
    console.log(`${s}|${r.n}|${eof(1)}|${seek(1)}`);
    seek(1, 1);
    console.log(inputString(1, 3));
    fileClose(1);
    fileOpen(2, "work/b.bin", forBinary);
    filePut(2, 42, 1);
    filePut(2, "あいう");
    const arr: int[] = [7, 8];
    filePut(2, arr);
    let back = 0;
    fileGet(2, back, 1);
    let txt = "xxx";
    fileGet(2, txt);
    const arr2: int[] = [0, 0];
    fileGet(2, arr2);
    console.log(`${back} ${txt} ${arr2.join(",")} lof=${lof(2)} loc=${loc(2)} mode=${fileAttr(2, 1) == forBinary} eof=${eof(2)}`);
    fileGet(2, back);
    console.log(`eof=${eof(2)}`);
    fileClose();
    fileOpen(3, "work/r.dat", forRandom, accessDefault, lockDefault, 16);
    filePut(3, "abc", 2);
    let rs = "";
    fileGet(3, rs, 2);
    console.log(`${rs} ${loc(3)} ${lof(3)}`);
    reset();
    rename("work/b.bin", "work/c.bin");
    let names = "";
    let e = dir("work\\*.*");
    while (e != "") {
        names += `${e};`;
        e = dir();
    }
    console.log(names);
    try {
        fileOpen(9, "work/none.txt", forInput);
    } catch (err) {
        console.log(`${err.code}`);
    }
    try {
        printLine(5, "x");
    } catch (err) {
        console.log(`${err.code}`);
    }
    fileOpen(4, "work/c.bin", forBinary, accessRead);
    try {
        filePut(4, 1);
    } catch (err) {
        console.log(`${err.code}`);
    }
    try {
        kill("work/c.bin");
    } catch (err) {
        console.log(`${err.code}`);
    }
    fileClose(4);
    kill("work\\*.*");
    rmDir("work");
    console.log(`${dir("work", vbDirectory) == ""}`);
}
"#;

#[test]
fn files_opened_by_number_run_without_excel() {
    let files = [("Main", FILE_IO)];
    let (_prog, out) = compile(&files);
    if !has_node() {
        return;
    }
    let p = Project::new("fileio", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout.lines().collect::<Vec<_>>(),
        vec![
            "len=55 free=1",
            "[名前12.5true]",
            "[no newline]",
            "x,y|3|false|-1.5",
            "y|7|true|56",
            "名前1",
            "42 あいう 7,8 lof=18 loc=18 mode=true eof=false",
            "eof=true",
            "abc 2 21",
            "c.bin;out.txt;r.dat;",
            "53",
            "52",
            "75",
            "75",
            "true",
        ]
    );
}

const ASK_MAIN: &str = r#"export function run() {
    const color = new frmAsk("色", "青").showModal();
    console.log(`color: ${color}`);
    const a = new frmAsk("A");
    const b = new frmAsk("B");
    console.log(`${a.describe()} ${b.describe()}`);
    console.log(`cancelled: [${a.showModal()}]`);
    console.log(`closed: [${b.showModal()}]`);
    frmAsk.txtAnswer.text = "既定";
    console.log(`default: ${frmAsk.showModal()}`);
}
"#;

/// `new frmAsk(...)` runs the form's `constructor`; instances keep their own variables;
/// `showModal()` returns what `this.close(value)` passed (spec §14.11.9).
#[test]
fn forms_take_arguments_and_return_values() {
    let spec = include_str!("js/forms/frmAsk.form.json");
    let code = include_str!("js/forms/frmAsk.ras");
    let mut cfg = Config::default();
    cfg.forms = vec![FormDecl::from_form_json(spec).unwrap()];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: ASK_MAIN.into(),
        },
        form_source(
            PathBuf::from("forms/frmAsk.ras"),
            "frmAsk",
            code.to_string(),
        ),
    ];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    if !has_node() {
        return;
    }
    let root = std::env::temp_dir().join(format!("ras-js-{}-ask", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    let write = |rel: &str, text: &str| {
        let p = root.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    };
    write("src/main.ras", ASK_MAIN);
    write("forms/frmAsk.ras", code);
    write("forms/frmAsk.form.json", spec);
    for f in &out.files {
        write(&format!("build/debug/{}", f.rel_path), &f.text);
    }
    let p = Project(root);
    let (stdout, stderr, ok) = p.run("Main.run", Some(&fixture("mock_ask.js")));
    assert!(ok, "{stderr}");
    let expected = "\
show frmAsk#1 \"色\" modal=true
asking 色 (1)
hide #1
bye 色
color: 色=赤
A/0 B/0
show frmAsk#2 \"A\" modal=true
asking A (1)
hide #2
bye A
cancelled: []
show frmAsk#3 \"B\" modal=true
asking B (1)
hide #3
bye B
closed: []
show frmAsk#4 \"質問\" modal=true
asking  (1)
hide #4
bye 
default: =既定
";
    assert_eq!(stdout, expected, "{stderr}");
}

const ORDER_MAIN: &str =
    "export function run() {\n    console.log(new frmOrder().showModal());\n}\n";

/// The validation of a form (spec §14.11.12) in the debug run: a rule checks its control when
/// the focus leaves it, the error goes to the label `<control>Error` (made under the control when
/// the form has none, transparent), a committed control is checked again as it changes, rules
/// change with what was entered (`txtCorp` shows its error in the label the form passes as
/// `errorLabel`), and `validateAll` shows every error and focuses the first.
#[test]
fn forms_validate_their_input() {
    let spec = include_str!("js/forms/frmOrder.form.json");
    let code = include_str!("js/forms/frmOrder.ras");
    let mut cfg = Config::default();
    cfg.forms = vec![FormDecl::from_form_json(spec).unwrap()];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: ORDER_MAIN.into(),
        },
        form_source(
            PathBuf::from("forms/frmOrder.ras"),
            "frmOrder",
            code.to_string(),
        ),
    ];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    // The library is no entry point of the project.
    assert_eq!(out.entries, vec!["Main.run", "frmOrder.show"]);
    if !has_node() {
        return;
    }
    let root = std::env::temp_dir().join(format!("ras-js-{}-order", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    let write = |rel: &str, text: &str| {
        let p = root.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    };
    write("src/main.ras", ORDER_MAIN);
    write("forms/frmOrder.ras", code);
    write("forms/frmOrder.form.json", spec);
    for f in &out.files {
        write(&format!("build/debug/{}", f.rel_path), &f.text);
    }
    let p = Project(root);
    let (stdout, stderr, ok) = p.run("Main.run", Some(&fixture("mock_validation.js")));
    assert!(ok, "{stderr}");
    let expected = "\
show txtName,txtAge,txtAgeError,cboKind,fraCorp,txtCorp,lblCorpMessage,chkAgree,btnOK
  focus txtName
> {\"t\":\"focus\",\"control\":\"txtName\"}
> {\"t\":\"focus\",\"control\":\"txtAge\"}
  txtNameError \"名前を入力してください\" visible=true back=0 fore=204 at 12,31 150x15.5
  focus txtAge
> {\"t\":\"input\",\"control\":\"txtAge\",\"text\":\"abc\",\"seq\":1}
> {\"t\":\"focus\",\"control\":\"cboKind\"}
  txtAgeError \"整数で入力してください\" visible=true back=1 fore=2147483666 at 170,48 120x18
  focus cboKind
> {\"t\":\"focus\",\"control\":\"txtAge\"}
  cboKindError \"入力してください\" visible=true back=0 fore=204 at 12,103 150x15.5
  focus txtAge
> {\"t\":\"input\",\"control\":\"txtAge\",\"text\":\"30\",\"seq\":2}
  txtAgeError \"\" visible=false back=1 fore=2147483666 at 170,48 120x18
  focus txtAge
> {\"t\":\"select\",\"control\":\"cboKind\",\"index\":1}
  cboKindError \"\" visible=false back=0 fore=204 at 12,103 150x15.5
  focus cboKind
> {\"t\":\"click\",\"control\":\"btnOK\"}
  chkAgreeError \"同意が必要です\" visible=true back=0 fore=204 at 12,199 150x15.5
  fraCorp/lblCorpMessage \"法人名を入力してください\" visible=true back=1 fore=2147483666 at 6,30 180x14
  focus txtName
> {\"t\":\"input\",\"control\":\"txtName\",\"text\":\"山田\",\"seq\":3}
  txtNameError \"\" visible=false back=0 fore=204 at 12,31 150x15.5
  focus txtName
> {\"t\":\"click\",\"control\":\"chkAgree\"}
  chkAgreeError \"\" visible=false back=0 fore=204 at 12,199 150x15.5
  focus chkAgree
> {\"t\":\"input\",\"control\":\"txtCorp\",\"text\":\"ABC\",\"seq\":4}
  fraCorp/lblCorpMessage \"\" visible=false back=1 fore=2147483666 at 6,30 180x14
  focus txtCorp
> {\"t\":\"click\",\"control\":\"btnOK\"}
hide 
山田
";
    assert_eq!(stdout, expected, "{stderr}");
}

const SCHEMAS: &str = r#"import * as z from "validation";

export function run() {
    const limit = 3;
    const cases: [z.ValidationSchema, string][] = [
        [z.string().required(), "  "],
        [z.string(), ""],
        [z.string().max(3), "abcd"],
        [z.string().min(2, "2 文字以上"), "a"],
        [z.string().length(3), "ab"],
        [z.string().regex("^[0-9]+$"), "12a"],
        [z.string().email(), "a@b.jp"],
        [z.string().email(), "a@b"],
        [z.string().oneOf(["A", "B"]), "C"],
        [z.string().refine(v => v.length < limit, "短く"), "abc"],
        [z.int(), "1.5"],
        [z.int().min(1.0).max(10.0), "11"],
        [z.int("整数で"), "x"],
        [z.number().positive(), "0"],
        [z.number().nonnegative(), "-0.5"],
        [z.number(), " 2.5 "],
        [z.date(), "2026/13/40"],
        [z.date().min(cDbl(dateSerial(2026, 1, 1))), "2025/12/31"],
        [z.boolean().required(), "false"],
        [z.boolean(), "false"],
        [z.string().required().noTrim().min(2), " a"],
    ];
    for (const [s, v] of cases) {
        console.log(`[${v}] ${s.check(v)}`);
    }
}
"#;

/// The rules of `ValidationSchema` (spec §14.11.12).
#[test]
fn validation_schemas_check_values() {
    let files = [("main", SCHEMAS)];
    let (_, out) = compile(&files);
    if !has_node() {
        return;
    }
    let p = Project::new("schemas", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        "\
[  ] 入力してください
[] 
[abcd] 3 文字以内で入力してください
[a] 2 文字以上
[ab] 3 文字で入力してください
[12a] 形式が正しくありません
[a@b.jp] 
[a@b] メールアドレスの形式で入力してください
[C] 一覧から選択してください
[abc] 短く
[1.5] 整数で入力してください
[11] 10 以下で入力してください
[x] 整数で
[0] 0 より大きい値を入力してください
[-0.5] 0 以上で入力してください
[ 2.5 ] 
[2026/13/40] 日付で入力してください
[2025/12/31] 2026/01/01 以上で入力してください
[false] チェックしてください
[false] 
[ a] 
",
        "{stderr}"
    );
}

const DATES: &str = r#"export function run() {
    // The end of the month is kept, not carried over (unlike JavaScript's setMonth).
    console.log(dateAdd("m", 1, dateSerial(2026, 1, 31)), dateAdd("m", 1, dateSerial(2024, 1, 31)), dateAdd("m", -1, dateSerial(2026, 3, 31)));
    console.log(dateAdd("m", 13, dateSerial(2026, 1, 31)), dateAdd("yyyy", 1, dateSerial(2024, 2, 29)), dateAdd("q", 1, dateSerial(2026, 11, 30)));
    console.log(dateAdd("d", 1, dateSerial(2026, 12, 31)), dateAdd("d", -1, dateSerial(2026, 1, 1)), dateAdd("h", 25, dateSerial(2026, 12, 31) + timeSerial(23, 0, 0)));
    console.log(dateAdd("m", 1, dateSerial(2026, 1, 31) + timeSerial(10, 30, 0)));
    // DateSerial carries months and days over, and reads two-digit years as VBA does.
    console.log(dateSerial(2026, 14, 0), dateSerial(2026, 1, -1), dateSerial(2026, 0, 1), dateSerial(26, 1, 1), dateSerial(30, 1, 1));
    // Before 1899-12-30 the time is a fraction of the same sign as the day.
    const early = dateAdd("h", -18, dateSerial(1899, 12, 30));
    console.log(cDbl(early), early, day(early), hour(early), dateDiff("d", early, dateSerial(1899, 12, 30)));
    console.log(dateSerial(1899, 12, 29) + timeSerial(6, 0, 0));
    // Out of 100–9999: error 5.
    try { console.log(dateAdd("d", 1, dateSerial(9999, 12, 31))); } catch (e) { console.log(`error ${e.code}`); }
    try { console.log(dateAdd("yyyy", -2000, dateSerial(2026, 1, 1))); } catch (e) { console.log(`error ${e.code}`); }
}
"#;

/// `DateAdd` / `DateSerial` / `Date` values compute as VBA does.
#[test]
fn dates_compute_as_in_vba() {
    let files = [("main", DATES)];
    let (_, out) = compile(&files);
    if !has_node() {
        return;
    }
    let p = Project::new("dates", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        "\
2026/02/28 2024/02/29 2026/02/28
2027/02/28 2025/02/28 2027/02/28
2027/01/01 2025/12/31 2027/01/02
2026/02/28 10:30:00
2027/01/31 2025/12/30 2025/12/01 2026/01/01 1930/01/01
-1.25 1899/12/29 6:00:00 29 6 1
18:00:00
error 5
error 5
",
        "{stderr}"
    );
}

/// `rule(..., errorLabel: label)` (spec §14.11.12): the error moves to the label passed, the one
/// shown before is cleared, and calling `rule` again without a label keeps the last one.
#[test]
fn validators_show_errors_in_the_labels_passed() {
    let spec = include_str!("js/forms/frmLabels.form.json");
    let code = include_str!("js/forms/frmLabels.ras");
    let main = "export function main() {\n    new frmLabels().showModal();\n}\n";
    let mut cfg = Config::default();
    cfg.forms = vec![FormDecl::from_form_json(spec).unwrap()];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: main.into(),
        },
        form_source(
            PathBuf::from("forms/frmLabels.ras"),
            "frmLabels",
            code.to_string(),
        ),
    ];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    if !has_node() {
        return;
    }
    let root = std::env::temp_dir().join(format!("ras-js-{}-labels", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    let write = |rel: &str, text: &str| {
        let p = root.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    };
    write("src/main.ras", main);
    write("forms/frmLabels.ras", code);
    write("forms/frmLabels.form.json", spec);
    for f in &out.files {
        write(&format!("build/debug/{}", f.rel_path), &f.text);
    }
    let p = Project(root);
    let (stdout, stderr, ok) = p.run("Main.main", Some(&fixture("mock_labels.js")));
    assert!(ok, "{stderr}");
    assert_eq!(
        stdout,
        "\
txtAError \"必須\" visible=true
auto=[]false one=[必須]true
one=[]false two=[必須]true
two=[必須]true (kept)
",
        "{stderr}"
    );
}

const UPLOAD_MAIN: &str = "export function run() {\n    frmUpload.show();\n    console.log(frmUpload.pbUpload.value);\n}\n";

/// A progress bar of a form (spec §14.11.13) in the debug run: its labels are drawn as the form
/// is written into the workbook, and setting its properties moves them.
#[test]
fn progress_bars_move_their_labels() {
    let spec = include_str!("js/forms/frmUpload.form.json");
    let code = include_str!("js/forms/frmUpload.ras");
    let mut cfg = Config::default();
    cfg.forms = vec![FormDecl::from_form_json(spec).unwrap()];
    let sources = vec![
        SourceInput {
            path: PathBuf::from("src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: UPLOAD_MAIN.into(),
        },
        form_source(
            PathBuf::from("forms/frmUpload.ras"),
            "frmUpload",
            code.to_string(),
        ),
    ];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    assert_eq!(out.entries, vec!["Main.run", "frmUpload.show"]);
    if !has_node() {
        return;
    }
    let root = std::env::temp_dir().join(format!("ras-js-{}-upload", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    let write = |rel: &str, text: &str| {
        let p = root.join(rel);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    };
    write("src/main.ras", UPLOAD_MAIN);
    write("forms/frmUpload.ras", code);
    write("forms/frmUpload.form.json", spec);
    for f in &out.files {
        write(&format!("build/debug/{}", f.rel_path), &f.text);
    }
    let p = Project(root);
    let (stdout, stderr, ok) = p.run("Main.run", Some(&fixture("mock_progress.js")));
    assert!(ok, "{stderr}");
    // After the form is closed, `frmUpload.pbUpload` loads it again (as VBA does): 56 as designed.
    let expected = "\
Upload progress 56 56%
show pbUpload_Track,pbUpload_Bar,pbUpload_Caption,pbUpload_Value,btnStart
  pbUpload_Track \"\" visible=true back=15066597 at 12,31.5 240x4.5
  pbUpload_Bar \"\" visible=true back=2147483661 at 12,31.5 134.25x4.5
  pbUpload_Caption \"Upload progress\" visible=true back=2147483663 at 12,12 204x16.5
  pbUpload_Value \"56%\" visible=true back=2147483663 at 216,12 36x16.5
> {\"t\":\"click\",\"control\":\"btnStart\"}
25% 0.25
  pbUpload_Caption \"Uploading\" visible=true back=2147483663 at 12,12 120x16.5
  pbUpload_Bar \"\" visible=true back=2147483661 at 12,31.5 120x4.5
  pbUpload_Value \"100%\" visible=false back=2147483663 at 132,12 0.75x16.5
  pbUpload_Track \"\" visible=true back=15066597 at 12,31.5 120x4.5
Upload progress 56 56%
56
";
    assert_eq!(stdout, expected, "{stderr}");
}
