//! The stand-in for Excel of the debug run time (spec §14.10.8), end to end: RAS compiled to
//! JavaScript, run under Node with this crate's addon on a workbook, without Excel.
//!
//! Needs `node` on PATH; skipped without it.

use std::path::{Path, PathBuf};
use std::process::Command;

use ras_workbook::RangeRef;
use ras_workbook::book::{Book, Read, Val};
use ras_compiler::{analyze, build_js, Config, ModuleKind, SourceInput};

fn has_node() -> bool {
    let found = Command::new("node").arg("--version").output().is_ok_and(|o| o.status.success());
    if !found {
        eprintln!("skipped: node is not installed");
    }
    found
}

/// This crate's addon, as Node loads it (a `.node` file).
fn addon(dir: &Path) -> PathBuf {
    let exe = std::env::current_exe().unwrap();
    let deps = exe.parent().unwrap();
    let name = if cfg!(windows) {
        "ras_book.dll"
    } else if cfg!(target_os = "macos") {
        "libras_book.dylib"
    } else {
        "libras_book.so"
    };
    let lib = [deps.join(name), deps.parent().unwrap().join(name)].into_iter().find(|p| p.exists()).unwrap_or_else(|| panic!("{name} was not built next to {}", exe.display()));
    let target = dir.join("ras_book.node");
    std::fs::copy(lib, &target).unwrap();
    target
}

struct Project(PathBuf);

impl Drop for Project {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// A project with one module, its workbook `Book.xlsx` (as `book` leaves it), and its debug build.
fn project(name: &str, source: &str, book: impl FnOnce(&mut Book)) -> Project {
    let root = std::env::temp_dir().join(format!("ras-book-{}-{name}", std::process::id()));
    let _ = std::fs::remove_dir_all(&root);
    std::fs::create_dir_all(root.join("src")).unwrap();
    std::fs::write(root.join("src/main.ras"), source).unwrap();
    let mut b = Book::new_book();
    book(&mut b);
    b.save(&root.join("Book.xlsx")).unwrap();

    let mut cfg = Config::default();
    cfg.project.workbook = Some("Book.xlsx".into());
    cfg.sheet_code_names = vec!["Sheet1".into()];
    let sources = vec![SourceInput { path: root.join("src/main.ras"), rel_path: "src/main.ras".into(), key: "main".into(), kind: ModuleKind::User, text: source.into() }];
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog.diagnostics.iter().filter(|d| d.is_error()).map(|d| format!("{}: {}", d.code, d.message)).collect();
    assert!(errors.is_empty(), "unexpected errors: {errors:#?}");
    let out = build_js(&prog, "../../");
    for f in &out.files {
        let p = root.join("build/debug").join(&f.rel_path);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, &f.text).unwrap();
    }
    Project(root)
}

impl Project {
    fn run(&self, entry: &str) -> (String, String, bool) {
        let addon = addon(&self.0);
        let o = Command::new("node")
            .arg("--enable-source-maps")
            .arg(self.0.join("build/debug/ras-main.js"))
            .arg(entry)
            .current_dir(&self.0)
            .env_remove("RAS_COM_ADDON")
            .env_remove("RAS_MOCK_HOST")
            .env("RAS_BOOK_ADDON", addon)
            .output()
            .expect("node");
        (String::from_utf8_lossy(&o.stdout).replace("\r\n", "\n"), String::from_utf8_lossy(&o.stderr).into_owned(), o.status.success())
    }

    fn saved(&self) -> Book {
        Book::open(&self.0.join("build/debug/workbooks/Book.xlsx")).expect("the changed workbook is saved")
    }
}

fn rng(a1: &str) -> RangeRef {
    RangeRef::parse(a1).unwrap()
}

fn lines(stdout: &str) -> Vec<&str> {
    stdout.lines().filter(|l| !l.starts_with("[ras]")).collect()
}

const REPORT: &str = r#"
export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

export function rebuild() {
    const ws = Sheet1;
    const last = ws.lastRow();
    let total = 0.0;
    for (let r = 2; r <= last; r++) {
        const price: double = ws.cells(r, 3).value;
        const qty: double = ws.cells(r, 4).value;
        total += price * qty;
    }
    for (let r = 2; r <= last; r++) {
        ws.cells(r, 5).formula = `=ROUND(C${r}*D${r}*1.1,0)`;
    }
    ws.range("E1").value = "税込";
    ws.range("F1").formula = "=SUM(E:E)";
    console.log(`last=${last} total=${total} sum=${ws.range("F1").value}`);
    // Formulas follow what the program writes next.
    ws.cells(2, 4).value = 10;
    console.log(`E2=${ws.range("E2").value} F1=${ws.range("F1").value} formula=${ws.range("E2").formula}`);
    console.log(`used=${ws.usedRange.address} region=${ws.range("B2").currentRegion.address(false, false)}`);
    console.log(`match=${application.worksheetFunction.match("B社", ws.range("A:A"), 0)} count=${application.worksheetFunction.countIf(ws.range("A2:A4"), "*社")}`);
    const hit = ws.range("A1:A10").find("C社");
    console.log(`find=${hit.address} next=${ws.cells(1, 1).end(xlDown).address}`);
    let names = "";
    for (const c of ws.range("A2:A4")) names += `${c.value};`;
    console.log(`names=${names} text=${ws.range("C3").text}`);
}

export function dates() {
    const ws = Sheet1;
    ws.range("G1").value = "2026/3/4";
    const v = ws.range("G1").value;
    console.log(`${typeName(v)} ${v} ${ws.range("G1").numberFormat} ${ws.range("G1").value2}`);
    ws.range("G2").value = dateSerial(2026, 1, 31);
    ws.range("G3").formula = "=G2+1";
    console.log(`${ws.range("G3").text} ${typeName(ws.range("G2").value)}`);
}

export function sheets() {
    const added = thisWorkbook.worksheets.add(undefined, Sheet1);
    added.name = "集計";
    added.range("A1").formula = "=SUM(Sheet1!C2:C4)";
    Sheet1.name = "売上";
    console.log(`${thisWorkbook.worksheets.count} ${added.index} ${added.range("A1").formula} ${added.range("A1").value}`);
    for (const s of thisWorkbook.worksheets) console.log(`- ${s.name} (${s.codeName})`);
}

export function failing() {
    try {
        Sheet1.range("A1").sort();
    } catch (e) {
        console.log(`${e.code}: ${e.description}`);
    }
    try {
        application.worksheetFunction.vLookup("X社", Sheet1.range("A2:D4"), 2, false);
    } catch (e) {
        console.log(`${e.code}: ${e.description}`);
    }
    console.log(`${typeName(application.vLookup("X社", Sheet1.range("A2:D4"), 2, false))}`);
}
"#;

fn sales(b: &mut Book) {
    let t = |s: &str| Val::Text(s.into());
    let n = Val::Num;
    b.set_values(
        0,
        rng("A1:D4"),
        &[
            vec![t("顧客"), t("担当"), t("単価"), t("数量")],
            vec![t("A社"), t("佐藤"), n(100.0), n(2.0)],
            vec![t("B社"), t("鈴木"), n(250.5), n(4.0)],
            vec![t("C社"), t("佐藤"), n(30.0), n(3.0)],
        ],
    )
    .unwrap();
    b.set_number_format(0, rng("C2:C4"), "#,##0.00").unwrap();
}

#[test]
fn a_program_runs_on_the_workbook_without_excel() {
    if !has_node() {
        return;
    }
    let p = project("report", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.rebuild");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "last=4 total=1292 sum=1421",
            "E2=1100 F1=2301 formula==ROUND(C2*D2*1.1,0)",
            "used=$A$1:$F$4 region=A1:F4",
            "match=3 count=3",
            "find=$A$4 next=$A$4",
            "names=A社;B社;C社; text=250.50",
        ]
    );
    assert!(stdout.contains("build/debug/workbooks/Book.xlsx"), "{stdout}");
    // The workbook itself is not written; the copy has what the program did.
    let mut original = Book::open(&p.0.join("Book.xlsx")).unwrap();
    assert_eq!(original.get(0, rng("E2"), Read::Value).unwrap(), vec![vec![Val::Empty]]);
    let mut saved = p.saved();
    assert_eq!(saved.get(0, rng("E2:F2"), Read::Formula).unwrap(), vec![vec![Val::Text("=ROUND(C2*D2*1.1,0)".into()), Val::Text("".into())]]);
    assert_eq!(saved.get(0, rng("F1"), Read::Value).unwrap(), vec![vec![Val::Num(2301.0)]]);
}

#[test]
fn dates_are_dates() {
    if !has_node() {
        return;
    }
    let p = project("dates", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.dates");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(lines(&stdout), vec!["Date 2026/03/04 yyyy/m/d 46085", "2026/2/1 Date"]);
}

#[test]
fn sheets_are_added_and_renamed() {
    if !has_node() {
        return;
    }
    let p = project("sheets", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.sheets");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(lines(&stdout), vec!["2 2 =SUM(売上!C2:C4) 380.5", "- 売上 (Sheet1)", "- 集計 (Sheet2)"]);
}

#[test]
fn what_needs_excel_says_so() {
    if !has_node() {
        return;
    }
    let p = project("failing", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.failing");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "445: 並べ替え（Range.Sort） は Excel が必要です。Excel のないデバッグ実行（ブックの代役）では使えません",
            "1004: WorksheetFunction クラスの VLookup プロパティを取得できません。",
            "Error",
        ]
    );
}
