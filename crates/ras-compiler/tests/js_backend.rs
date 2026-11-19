//! The JavaScript back end (spec §14.10): the generated code must run under Node with the
//! semantics the VBA build has, and its source maps must lead back to the RAS source.
//!
//! The tests that execute code need `node` on PATH and are skipped without it.

#![allow(clippy::field_reassign_with_default)]

use std::path::{Path, PathBuf};
use std::process::Command;

use ras_compiler::{analyze, build_js, Config, JsOutput, ModuleKind, Program, SourceInput};

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
    let errors: Vec<String> = prog.diagnostics.iter().filter(|d| d.is_error()).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    assert!(!out.files.is_empty(), "no output: {:?}", out.diagnostics);
    (prog, out)
}

fn has_node() -> bool {
    let found = Command::new("node").arg("--version").output().is_ok_and(|o| o.status.success());
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
        cmd.arg("--enable-source-maps").arg(self.0.join("build").join("debug").join("ras-main.js")).arg(entry).current_dir(&self.0);
        cmd.env_remove("RAS_COM_ADDON").env_remove("RAS_MOCK_HOST");
        if let Some(m) = mock {
            cmd.env("RAS_MOCK_HOST", m);
        }
        let o = cmd.output().expect("node");
        (String::from_utf8_lossy(&o.stdout).replace("\r\n", "\n"), String::from_utf8_lossy(&o.stderr).into_owned(), o.status.success())
    }
}

impl Drop for Project {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests").join("js").join(name)
}

#[test]
fn language_semantics_match_the_vba_build() {
    let files = [("main", include_str!("js/semantics.ras")), ("lib", include_str!("js/lib.ras"))];
    let (_, out) = compile(&files);
    assert_eq!(out.entries, vec!["Main.run"]);
    if !has_node() {
        return;
    }
    let p = Project::new("semantics", &files, &out);
    let (stdout, stderr, ok) = p.run("Main.run", None);
    assert!(ok, "{stderr}");
    assert_eq!(stdout, include_str!("js/semantics.out").replace("\r\n", "\n"));
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
    assert_eq!(out.entries, vec!["Report.rebuildReport", "Report.listSheets", "Report.countCustomers"]);
    let js = &out.files.iter().find(|f| f.rel_path == "report.js").unwrap().text;
    assert!(js.contains("$rt.cnum($rt.get($rt.call(ws, \"Cells\", [r, 3]), \"Value\"))"), "{js}");
    assert!(js.contains("$rt.put($rt.call(ws, \"Cells\", [r, 5]), \"Formula\","), "{js}");
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
    assert!(js.contains("Array.from($rt.iter($rt.global(\"Worksheets\"))).filter("), "{js}");
    let (stdout, stderr, ok) = p.run("listSheets", Some(&mock));
    assert!(ok, "{stderr}");
    assert!(stdout.starts_with("2 Summary\n0:Data,1:Summary Hidden 3 true Data\n"), "{stdout}");
    // Without Excel, the first host access fails with an explanation and a RAS source position.
    let (_, stderr, ok) = p.run("Report.rebuildReport", None);
    assert!(!ok);
    assert!(stderr.contains("ras_com") && stderr.contains("report.ras:6"), "{stderr}");
}

#[test]
fn source_maps_lead_back_to_ras_lines() {
    let src = "function f(z: int): int {\n    return 10 / z;\n}\nexport function run() {\n    try {\n        f(0);\n    } catch (e) {\n        console.log(e.code, e.rasLine, e.rasFile.endsWith(\"m.ras\"));\n    }\n}\n";
    let files = [("m", src)];
    let (_, out) = compile(&files);
    let map: serde_json::Value = serde_json::from_str(&out.files.iter().find(|f| f.rel_path == "m.js.map").unwrap().text).unwrap();
    assert_eq!(map["version"], 3);
    assert_eq!(map["sources"][0], "../../src/m.ras");
    assert!(map["mappings"].as_str().unwrap().len() > 10);
    let js = &out.files.iter().find(|f| f.rel_path == "m.js").unwrap().text;
    assert!(js.ends_with("//# sourceMappingURL=m.js.map\n"));
    assert!(!js.contains('\u{1}') && !js.contains('\u{2}'), "position markers must be stripped");
    if !has_node() {
        return;
    }
    let p = Project::new("maps", &files, &out);
    let (stdout, stderr, ok) = p.run("run", None);
    assert!(ok, "{stderr}");
    assert_eq!(stdout, "11 2 true\n");
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
    let frames: Vec<&str> = stderr.lines().filter(|l| l.trim_start().starts_with("at ")).collect();
    assert_eq!(frames, ["    at f (src/m.ras:2:5)", "    at g (src/m.ras:5:5)", "    at run (src/m.ras:8:5) {"], "{stderr}");
    assert!(stderr.contains("rasFile: 'src/m.ras'") && stderr.contains("rasLine: 2"), "{stderr}");
    assert!(!stderr.contains("ras-runtime.js") && !stderr.contains("node:internal"), "{stderr}");
}

#[test]
fn modules_in_subdirectories_and_import_cycles() {
    let a = "import { pong } from \"./util/b\";\nexport let hits = 0;\nexport function ping(n: int): int {\n    hits++;\n    return n <= 0 ? 0 : pong(n - 1) + 1;\n}\nexport function run() {\n    console.log(ping(5), hits);\n}\n";
    let b = "import { ping } from \"../a\";\nexport function pong(n: int): int {\n    return n <= 0 ? 0 : ping(n - 1) + 1;\n}\n";
    let files = [("a", a), ("util/b", b)];
    let (_, out) = compile(&files);
    let map: serde_json::Value = serde_json::from_str(&out.files.iter().find(|f| f.rel_path == "util/b.js.map").unwrap().text).unwrap();
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
