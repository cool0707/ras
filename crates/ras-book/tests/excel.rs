//! The stand-in for Excel of the debug run time (spec §14.10.8), end to end: RAS compiled to
//! JavaScript, run under Node with this crate's addon on a workbook, without Excel.
//!
//! Needs `node` on PATH; skipped without it.

use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::OnceLock;

use ras_compiler::{Config, ModuleKind, SourceInput, analyze, build_js};
use ras_workbook::book::{Book, Read, Val};
use ras_workbook::{CellRef, RangeRef};

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

/// This crate's addon, built once for the test run. `cargo test` does not build the cdylib of a
/// package whose integration tests cannot link it, so the test builds it with Cargo, into the
/// same target directory and profile as itself.
fn built_addon() -> &'static Path {
    static BUILT: OnceLock<PathBuf> = OnceLock::new();
    BUILT.get_or_init(|| {
        let exe = std::env::current_exe().unwrap();
        // <target>/<profile>/deps/excel-<hash>
        let profile_dir = exe
            .parent()
            .and_then(Path::parent)
            .expect("the test runs from <target>/<profile>/deps");
        let target_dir = profile_dir
            .parent()
            .expect("the profile directory is in the target directory");
        let mut cmd = Command::new(env!("CARGO"));
        cmd.args(["build", "-p", "ras-book", "--target-dir"])
            .arg(target_dir)
            .current_dir(env!("CARGO_MANIFEST_DIR"));
        if profile_dir.file_name().is_some_and(|n| n == "release") {
            cmd.arg("--release");
        }
        let out = cmd.output().expect("cargo");
        assert!(
            out.status.success(),
            "cargo build -p ras-book failed:\n{}",
            String::from_utf8_lossy(&out.stderr)
        );
        let name = if cfg!(windows) {
            "ras_book.dll"
        } else if cfg!(target_os = "macos") {
            "libras_book.dylib"
        } else {
            "libras_book.so"
        };
        let lib = profile_dir.join(name);
        assert!(lib.exists(), "{} was not built", lib.display());
        lib
    })
}

/// The addon as Node loads it (a `.node` file), in the project directory.
fn addon(dir: &Path) -> PathBuf {
    let target = dir.join("ras_book.node");
    std::fs::copy(built_addon(), &target).unwrap();
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
    project_with(name, source, &[], book)
}

/// A project of `src/main.ras` and the files of `objects/` (name, text).
fn project_with(
    name: &str,
    source: &str,
    objects: &[(&str, &str)],
    book: impl FnOnce(&mut Book),
) -> Project {
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
    let mut sources = vec![SourceInput {
        path: root.join("src/main.ras"),
        rel_path: "src/main.ras".into(),
        key: "main".into(),
        kind: ModuleKind::User,
        text: source.into(),
    }];
    for (object, text) in objects {
        let path = root.join(format!("objects/{object}.ras"));
        std::fs::create_dir_all(path.parent().unwrap()).unwrap();
        std::fs::write(&path, text).unwrap();
        sources.push(ras_compiler::object_source(path, object, text.to_string()));
    }
    let prog = analyze(sources, cfg);
    let errors: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| format!("{}: {}", d.code, d.message))
        .collect();
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
        (
            String::from_utf8_lossy(&o.stdout).replace("\r\n", "\n"),
            String::from_utf8_lossy(&o.stderr).into_owned(),
            o.status.success(),
        )
    }

    fn saved(&self) -> Book {
        Book::open(&self.0.join("build/debug/workbooks/Book.xlsx"))
            .expect("the changed workbook is saved")
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
        application.run("手書きのマクロ");
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
    assert!(
        stdout.contains("build/debug/workbooks/Book.xlsx"),
        "{stdout}"
    );
    // The workbook itself is not written; the copy has what the program did.
    let mut original = Book::open(&p.0.join("Book.xlsx")).unwrap();
    assert_eq!(
        original.get(0, rng("E2"), Read::Value).unwrap(),
        vec![vec![Val::Empty]]
    );
    let mut saved = p.saved();
    assert_eq!(
        saved.get(0, rng("E2:F2"), Read::Formula).unwrap(),
        vec![vec![
            Val::Text("=ROUND(C2*D2*1.1,0)".into()),
            Val::Text("".into())
        ]]
    );
    assert_eq!(
        saved.get(0, rng("F1"), Read::Value).unwrap(),
        vec![vec![Val::Num(2301.0)]]
    );
}

#[test]
fn dates_are_dates() {
    if !has_node() {
        return;
    }
    let p = project("dates", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.dates");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec!["Date 2026/03/04 m/d/yyyy 46085", "2026/2/1 Date"]
    );
}

#[test]
fn sheets_are_added_and_renamed() {
    if !has_node() {
        return;
    }
    let p = project("sheets", REPORT, sales);
    let (stdout, stderr, ok) = p.run("Main.sheets");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "2 2 =SUM(売上!C2:C4) 380.5",
            "- 売上 (Sheet1)",
            "- 集計 (Sheet2)"
        ]
    );
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
            "445: 手書きの VBA（手書きのマクロ）は、Excel のないデバッグ実行では実行できません",
            "1004: WorksheetFunction クラスの VLookup プロパティを取得できません。",
            "Error",
        ]
    );
}

/// What `sales` has, read back as lines of `a|b|c|d`.
const HELPERS: &str = r#"
function dump(r: Excel.Range): string {
    let out = "";
    let rows = 0;
    for (const row of r.rows) {
        let line = "";
        let cells = 0;
        for (const c of row.cells) {
            if (cells > 0) line += "|";
            line += c.text;
            cells++;
        }
        if (rows > 0) out += " / ";
        out += line;
        rows++;
    }
    return out;
}
"#;

const EDITS: &str = r#"
export function insertAndDelete() {
    const ws = Sheet1;
    ws.range("F1").formula = "=SUM(C2:C4)";
    ws.range("A2:D2").font.bold = true;
    ws.rows(3).insert();
    console.log(`${dump(ws.range("A1:D5"))}`);
    console.log(`${ws.range("F1").formula} ${ws.range("F1").value} bold=${ws.range("A3").font.bold},${ws.range("A4").font.bold}`);
    ws.range("B2").insert(xlShiftToRight);
    console.log(`${dump(ws.range("A2:E2"))}`);
    ws.range("B2").delete(xlShiftToLeft);
    ws.columns("B").delete();
    ws.range("A3").entireRow.delete();
    console.log(`${dump(ws.range("A1:C4"))} ${ws.range("E1").formula}`);
    // Rows deleted from the bottom up, as macros do.
    for (let r = ws.cells(ws.rows.count, 1).end(xlUp).row; r >= 2; r--) {
        if (ws.cells(r, 3).value < 4) ws.rows(r).delete();
    }
    console.log(`${dump(ws.usedRange)}`);
}

export function clipboard() {
    const ws = Sheet1;
    ws.range("E2").formula = "=C2*D2";
    ws.range("E2").interior.color = rgb(255, 0, 0);
    ws.range("E2").copy();
    console.log(`mode=${application.cutCopyMode}`);
    ws.range("E3:E4").pasteSpecial(xlPasteFormulas);
    console.log(`${ws.range("E4").formula} ${ws.range("E4").value} color=${ws.range("E4").interior.colorIndex} ${ws.range("E2").interior.color}`);
    ws.range("E2:E4").copy();
    ws.range("G2").pasteSpecial(xlPasteValues);
    ws.range("G6").pasteSpecial(paste: xlPasteValues, transpose: true);
    application.cutCopyMode = 0;
    console.log(`${ws.range("G3").formula} ${ws.range("G2").interior.colorIndex} ${dump(ws.range("G6:I6"))} ${ws.range("H6").formula} mode=${application.cutCopyMode}`);
    // Cut and paste: the references follow the cells.
    ws.range("H1").formula = "=C2+1";
    ws.range("C2:C4").cut();
    ws.paste(ws.range("K2"));
    console.log(`${ws.range("H1").formula} ${ws.range("E2").formula} ${ws.range("C2").value == null} ${ws.range("K3").text} sel=${selection.address(false, false)}`);
    // Writing a value leaves the copy, clearing cells ends it, as in Excel.
    ws.range("A2").copy();
    ws.range("J9").value = 1;
    console.log(`mode=${application.cutCopyMode}`);
    ws.range("J9").clearContents();
    try {
        ws.range("A9").pasteSpecial(xlPasteValues);
    } catch (e) {
        console.log(`${e.code} mode=${application.cutCopyMode}`);
    }
    // Insert copied cells.
    ws.rows(2).copy();
    ws.rows(4).insert();
    console.log(`${dump(ws.range("A2:B5"))}`);
}

export function sortAndFilter() {
    const ws = Sheet1;
    ws.range("A1:D4").sort(ws.range("C1"), xlDescending, header: xlYes);
    console.log(`${dump(ws.range("A1:A4"))}`);
    ws.sort.sortFields.clear();
    ws.sort.sortFields.add(ws.range("B2:B4"), order: xlAscending);
    ws.sort.sortFields.add(ws.range("A2:A4"), order: xlDescending);
    ws.sort.setRange(ws.range("A1:D4"));
    ws.sort.header = xlYes;
    ws.sort.apply();
    console.log(`${dump(ws.range("A1:B4"))}`);

    ws.range("F1").formula = "=SUBTOTAL(109,D2:D4)";
    ws.range("A1").autoFilter(2, "佐藤");
    console.log(`filter=${ws.autoFilterMode},${ws.filterMode} ${ws.autoFilter.range.address} hidden=${ws.rows(2).hidden},${ws.rows(3).hidden} sub=${ws.range("F1").value}`);
    const shown = ws.range("A2:D4").specialCells(xlCellTypeVisible);
    console.log(`${shown.address(false, false)} areas=${shown.areas.count} count=${shown.count}`);
    shown.copy(ws.range("A10"));
    console.log(`${dump(ws.range("A10:D11"))}`);
    ws.range("A1").autoFilter(4, ">=3", xlAnd, "<10");
    console.log(`${ws.range("A2:A4").specialCells(xlCellTypeVisible).address(false, false)}`);
    ws.showAllData();
    console.log(`filter=${ws.autoFilterMode},${ws.filterMode} ${ws.rows("2:4").hidden}`);
    ws.range("A1").autoFilter(1, ["A社", "C社"], xlFilterValues);
    console.log(`${ws.range("A2:A4").specialCells(xlCellTypeVisible).address(false, false)}`);
    ws.autoFilterMode = false;
    console.log(`filter=${ws.autoFilterMode} ${ws.rows("2:4").hidden}`);
    try {
        ws.range("H1:H3").specialCells(xlCellTypeConstants);
    } catch (e) {
        console.log(`${e.code}: ${e.description}`);
    }
    const u = union(ws.range("A1"), ws.range("C1:C2"), ws.range("A2"));
    u.font.italic = true;
    console.log(`${u.address(false, false)} ${u.count} ${ws.range("C2").font.italic} ${ws.range("B1").font.italic} ${intersect(u, ws.range("A1:B9")).address(false, false)}`);
    ws.range("A1:D4").removeDuplicates(2, xlYes);
    console.log(`${dump(ws.range("A1:B4"))}`);
}

export function formats() {
    const ws = Sheet1;
    const head = ws.range("A1:D1");
    head.font.bold = true;
    head.font.color = rgb(255, 255, 255);
    head.interior.color = rgb(0, 112, 192);
    head.horizontalAlignment = xlCenter;
    head.borders(xlEdgeBottom).lineStyle = xlContinuous;
    head.borders(xlEdgeBottom).weight = xlMedium;
    ws.range("A2:D4").borders.lineStyle = xlContinuous;
    ws.columns("A").columnWidth = 20;
    ws.rows(1).rowHeight = 24;
    ws.range("F1:G1").merge();
    ws.columns("D").hidden = true;
    console.log(`${head.font.bold} ${head.interior.color} ${ws.range("A1:A2").font.bold == null} ${ws.range("B3").borders(xlEdgeTop).lineStyle} ${ws.range("G1").mergeArea.address} ${ws.range("A1").columnWidth}`);
    thisWorkbook.save();
}

export function manual() {
    const ws = Sheet1;
    ws.range("F1").formula = "=SUM(D2:D4)";
    application.calculation = xlCalculationManual;
    ws.range("D2").value = 100;
    console.log(`${ws.range("F1").value}`);
    application.calculate();
    console.log(`${ws.range("F1").value}`);
    using _ = fastMode();
    ws.range("D2").value = 200;
    console.log(`${ws.range("F1").value} ${application.calculation == xlCalculationManual}`);
}
"#;

fn run_with(name: &str, source: &str, entry: &str) -> (Project, Vec<String>) {
    let p = project(name, &format!("{HELPERS}{source}"), sales);
    let (stdout, stderr, ok) = p.run(entry);
    assert!(ok, "{stderr}\n{stdout}");
    let out = lines(&stdout).into_iter().map(str::to_string).collect();
    (p, out)
}

#[test]
fn cells_are_inserted_and_deleted() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("edits", EDITS, "Main.insertAndDelete");
    assert_eq!(
        out,
        vec![
            "顧客|担当|単価|数量 / A社|佐藤|100.00|2 / ||| / B社|鈴木|250.50|4 / C社|佐藤|30.00|3",
            "=SUM(C2:C5) 380.5 bold=true,false",
            "A社||佐藤|100.00|2",
            "顧客|単価|数量 / A社|100.00|2 / B社|250.50|4 / C社|30.00|3 =SUM(B2:B4)",
            "顧客|単価|数量||250.50 / B社|250.50|4||",
        ]
    );
}

#[test]
fn cells_are_cut_copied_and_pasted() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("clipboard", EDITS, "Main.clipboard");
    assert_eq!(
        out,
        vec![
            "mode=1",
            "=C4*D4 90 color=-4142 255",
            "1002 -4142 200|1002|90 1002 mode=0",
            "=K2+1 =K2*D2 true 250.50 sel=G6:I6",
            "mode=1",
            "1004 mode=0",
            "A社|佐藤 / B社|鈴木 / A社|佐藤 / C社|佐藤",
        ]
    );
}

#[test]
fn ranges_are_sorted_and_filtered() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("sort", EDITS, "Main.sortAndFilter");
    assert_eq!(
        out,
        vec![
            "顧客 / B社 / A社 / C社",
            "顧客|担当 / C社|佐藤 / A社|佐藤 / B社|鈴木",
            "filter=true,true $A$1:$D$4 hidden=false,false sub=5",
            "A2:D3 areas=1 count=8",
            "C社|佐藤|30.00|3 / A社|佐藤|100.00|2",
            "A2",
            "filter=true,false false",
            "A2:A3",
            "filter=false false",
            "1004: 該当するセルが見つかりません。",
            "C1:C2,A1:A2 4 true false A1:A2",
            "顧客|担当 / C社|佐藤 / B社|鈴木 / |",
        ]
    );
}

#[test]
fn formats_are_saved_with_the_workbook() {
    if !has_node() {
        return;
    }
    let (p, out) = run_with("formats", EDITS, "Main.formats");
    assert_eq!(out, vec!["true 12611584 true 1 $F$1:$G$1 20"]);
    let mut saved = p.saved();
    let v = |b: &mut Book, a1: &str, prop: &str| b.format_of(0, rng(a1), prop).unwrap();
    assert_eq!(v(&mut saved, "A1:D1", "font.bold"), Some(Val::Bool(true)));
    assert_eq!(
        v(&mut saved, "A1:D1", "font.color"),
        Some(Val::Num(16777215.0))
    );
    assert_eq!(
        v(&mut saved, "B1", "interior.color"),
        Some(Val::Num(12611584.0))
    );
    assert_eq!(
        v(&mut saved, "C1", "horizontalalignment"),
        Some(Val::Num(-4108.0))
    );
    assert_eq!(
        v(&mut saved, "A1:D1", "borders.9.weight"),
        Some(Val::Num(-4138.0))
    );
    assert_eq!(
        v(&mut saved, "A2:D4", "borders.all.linestyle"),
        Some(Val::Num(1.0))
    );
    assert_eq!(v(&mut saved, "A9", "columnwidth"), Some(Val::Num(20.0)));
    assert_eq!(v(&mut saved, "A1", "rowheight"), Some(Val::Num(24.0)));
    assert_eq!(v(&mut saved, "F1:G1", "mergecells"), Some(Val::Bool(true)));
    assert_eq!(saved.hidden(0, true, 3, 3), Some(true));
    // The number format set before is still there.
    assert_eq!(
        saved.get(0, rng("C3"), Read::Text).unwrap(),
        vec![vec![Val::Text("250.50".into())]]
    );
}

#[test]
fn manual_calculation_stops_formulas() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("manual", EDITS, "Main.manual");
    assert_eq!(out, vec!["9", "107", "107 true"]);
}

/// `objects/Sheet1.ras`.
const EVENTS_SHEET1: &str = r#"
@on("Worksheet.Change")
function changed(target: Excel.Range) {
    console.log(`Change ${target.address(false, false)}`);
    // Written by the handler itself: not entered again.
    if (target.column == 1) this.cells(target.row, 8).value = "seen";
}

@on("Worksheet.SelectionChange")
function selected(target: Excel.Range) {
    console.log(`SelectionChange ${target.address(false, false)}`);
}

@on("Worksheet.Calculate")
function recalculated() {
    console.log(`Calculate F1=${this.range("F1").value}`);
}

@on("Worksheet.Deactivate")
function left() {
    console.log("Deactivate");
}

@on("Worksheet.Activate")
function entered() {
    console.log("Activate");
}

@on("Worksheet.BeforeDelete")
function deleting() {
    throw new Error("消せません");
}
"#;

/// `objects/ThisWorkbook.ras`.
const EVENTS_BOOK: &str = r#"
@on("Workbook.SheetChange")
function anySheet(sh: object, target: Excel.Range) {
    console.log(`SheetChange ${sh.name}!${target.address(false, false)}`);
}

@on("Workbook.NewSheet")
function added(sh: object) {
    console.log(`NewSheet ${sh.name}`);
}

@on("Workbook.BeforeSave")
function saving(saveAsUI: boolean, byref cancel: boolean) {
    console.log(`BeforeSave ${saveAsUI}`);
    cancel = true;
}
"#;

const EVENTS: &str = r#"
export function run() {
    const ws = Sheet1;
    ws.range("F1").formula = "=SUM(D2:D4)";
    ws.range("A5").value = "D社";
    ws.range("D2").value = 5;
    ws.range("B2").select();
    using _ = fastMode();
    ws.range("A6").value = "E社";
}

export function sheets() {
    const added = thisWorkbook.worksheets.add(after: Sheet1);
    added.range("A1").value = 1;
    Sheet1.activate();
    thisWorkbook.save();
    console.log(`saved=${thisWorkbook.saved}`);
    Sheet1.delete();
    console.log(`left=${thisWorkbook.worksheets.count}`);
}
"#;

#[test]
fn events_are_raised_for_what_the_program_does() {
    if !has_node() {
        return;
    }
    let p = project_with(
        "events",
        EVENTS,
        &[("Sheet1", EVENTS_SHEET1), ("ThisWorkbook", EVENTS_BOOK)],
        sales,
    );
    let (stdout, stderr, ok) = p.run("Main.run");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "Calculate F1=9",
            "Change F1",
            "SheetChange Sheet1!F1",
            "Change A5",
            "SheetChange Sheet1!H5",
            "SheetChange Sheet1!A5",
            "Calculate F1=12",
            "Change D2",
            "SheetChange Sheet1!D2",
            "SelectionChange B2",
        ]
    );
    let (stdout, stderr, ok) = p.run("Main.sheets");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "NewSheet Sheet2",
            "Deactivate",
            "SheetChange Sheet2!A1",
            "Activate",
            "BeforeSave false",
            "saved=false",
            "Deactivate",
            "left=1"
        ]
    );
    // The error of a handler is reported, and the program goes on.
    assert!(
        stderr.contains("Sheet1.BeforeDelete: 実行時エラー") && stderr.contains("消せません"),
        "{stderr}"
    );
}

/// `objects/ThisWorkbook.ras`: variables whose objects are listened to (spec §12.2.2).
const SOURCES_BOOK: &str = r#"
let app: Excel.Application;
let sheet: Excel.Worksheet;

export function listen() {
    app = application;
    sheet = Sheet1;
}

export function stop() {
    app = null;
    sheet = null;
}

@on("app.SheetChange")
function anyChange(sh: object, target: Excel.Range) {
    console.log(`app SheetChange ${sh.name}!${target.address(false, false)}`);
}

@on("app.WorkbookBeforeSave")
function saving(wb: Excel.Workbook, saveAsUI: boolean, byref cancel: boolean) {
    console.log(`app WorkbookBeforeSave ${wb === thisWorkbook} ${saveAsUI}`);
    cancel = true;
}

@on("sheet.Change")
function changed(target: Excel.Range) {
    console.log(`sheet Change ${target.address(false, false)}`);
}
"#;

const SOURCES: &str = r#"
class BookWatcher {
    private book: Excel.Workbook;
    private seen = 0;

    constructor() {
        this.book = thisWorkbook;
    }

    @on("book.SheetChange")
    private changed(sh: object, target: Excel.Range) {
        this.seen++;
        console.log(`class SheetChange ${target.address(false, false)} #${this.seen}`);
    }
}

let watcher: BookWatcher;

export function listen() {
    Sheet1.range("A6").value = "nobody listens yet";
    thisWorkbook.listen();
    watcher = new BookWatcher();
    Sheet1.range("A7").value = "x";
    thisWorkbook.save();
    console.log(`saved=${thisWorkbook.saved}`);
    thisWorkbook.stop();
    Sheet1.range("A8").value = "y";
}
"#;

/// The stand-in raises the events of the objects that variables hold (`WithEvents`).
#[test]
fn events_reach_the_variables_that_hold_the_object() {
    if !has_node() {
        return;
    }
    let p = project_with("sources", SOURCES, &[("ThisWorkbook", SOURCES_BOOK)], sales);
    let (stdout, stderr, ok) = p.run("Main.listen");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            // The sheet's own event, then the workbook's: to the variable that holds the
            // sheet, to the one that holds `Application`, to the field of the class.
            "sheet Change A7",
            "app SheetChange Sheet1!A7",
            "class SheetChange A7 #1",
            // The workbook's event reaches `Application` with the workbook in front.
            "app WorkbookBeforeSave true false",
            "saved=false",
            // Variables that hold nothing any more are not called.
            "class SheetChange A8 #2",
        ]
    );
}

/// Handlers attached at run time (`on`, spec §12.2.3).
const LISTENERS: &str = r#"
let changes = 0;

function saving(saveAsUI: boolean): boolean {
    console.log(`BeforeSave ${saveAsUI}`);
    return false;
}

export function listen() {
    const sheet = Sheet1.on("Change", (target) => {
        changes++;
        console.log(`sheet Change ${target.address(false, false)} #${changes}`);
    });
    application.addEventListener("SheetChange", (sh, target) => console.log(`app SheetChange ${sh.name}!${target.address(false, false)}`));
    const closing = application.on("WorkbookBeforeSave", (wb) => console.log(`app WorkbookBeforeSave ${wb === thisWorkbook}`));
    // Every sheet of a collection: the handler is told which one.
    thisWorkbook.worksheets.on("Change", (ws, target) => console.log(`each ${ws.name}!${target.address(false, false)}`));
    // A handler that returns `false` cancels the event.
    const save = thisWorkbook.on("BeforeSave", saving);
    Sheet1.range("A7").value = "x";
    thisWorkbook.save();
    console.log(`saved=${thisWorkbook.saved}`);
    sheet.off();
    sheet.off();
    save.off();
    closing.off();
    Sheet1.range("A8").value = "y";
    thisWorkbook.save();
    console.log(`saved=${thisWorkbook.saved} changes=${changes}`);
}
"#;

/// The stand-in raises the events for the handlers that `on` attached, until `off`.
#[test]
fn events_reach_the_handlers_attached_at_run_time() {
    if !has_node() {
        return;
    }
    let p = project_with("listeners", LISTENERS, &[], sales);
    let (stdout, stderr, ok) = p.run("Main.listen");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            "sheet Change A7 #1",
            "each Sheet1!A7",
            "app SheetChange Sheet1!A7",
            "BeforeSave false",
            "app WorkbookBeforeSave true",
            "saved=false",
            // Taken off: the handlers of the collection and of `Application` are left.
            "each Sheet1!A8",
            "app SheetChange Sheet1!A8",
            "saved=true changes=1",
        ]
    );
}

const MORE: &str = r#"
export function more() {
    const ws = Sheet1;
    // Several areas by address; deleting rows of several areas.
    const picked = ws.range("A2:A2,A4:A4");
    console.log(`${picked.address(false, false)} ${picked.areas.count} ${picked.entireRow.address(false, false)}`);
    picked.entireRow.delete();
    console.log(`${dump(ws.range("A1:B3"))}`);
    // Fill down: formulas move, formats come along.
    ws.range("E2").formula = "=C2*2";
    ws.range("E2").font.bold = true;
    ws.range("A3:B4").value = [["D社", "田中"], ["E社", "山田"]];
    ws.range("C3:D4").value = [[10, 1], [20, 2]];
    ws.range("E2:E4").fillDown();
    console.log(`${ws.range("E4").formula} ${ws.range("E4").value} ${ws.range("E4").font.bold}`);
    // Sorting with the headings guessed, from one cell.
    ws.range("A1").sort(key1: ws.range("C1"), order1: xlDescending, header: xlGuess);
    console.log(`${dump(ws.range("A1:A4"))} ${ws.range("E2").formula}`);
    // Cut cells inserted elsewhere: the rows change places.
    ws.rows(4).cut();
    ws.rows(2).insert();
    console.log(`${dump(ws.range("A1:A4"))} mode=${application.cutCopyMode}`);
    // An advanced filter: in place, then copied out.
    ws.range("G1:G2").value = [["数量"], [">1"]];
    ws.range("A1:D4").advancedFilter(xlFilterInPlace, ws.range("G1:G2"));
    console.log(`${ws.filterMode} ${ws.range("A2:A4").specialCells(xlCellTypeVisible).address(false, false)}`);
    ws.showAllData();
    ws.range("A1:D4").advancedFilter(xlFilterCopy, ws.range("G1:G2"), ws.range("I1"), true);
    console.log(`${dump(ws.range("I1:L3"))}`);
    // Borders as a collection, and around a range.
    ws.range("A1:B2").borderAround(xlContinuous, xlThick);
    console.log(`${ws.range("A1").borders(xlEdgeLeft).weight == xlThick} ${ws.range("B1").borders.item(xlEdgeLeft).lineStyle == xlLineStyleNone}`);
    // A copy of the sheet, and another workbook.
    Sheet1.copy(after: Sheet1);
    const copy = thisWorkbook.worksheets(2);
    console.log(`${copy.name} ${copy.range("A2").value} ${copy.range("E2").formula} ${activeSheet.name}`);
    const other = workbooks.add();
    ws.range("A1:E3").copy(other.worksheets(1).range("B2"));
    const pasted = other.worksheets(1);
    console.log(`${pasted.range("B4").value} ${pasted.range("F3").formula} ${pasted.range("D4").numberFormat}`);
    other.close(false);
    console.log(`${workbooks.count} ${activeWorkbook.name}`);
}
"#;

#[test]
fn the_rest_of_the_object_model() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("more", MORE, "Main.more");
    assert_eq!(
        out,
        vec![
            "A2,A4 2 2:2,4:4",
            "顧客|担当 / B社|鈴木 / |",
            "=C4*2 40 true",
            "顧客 / B社 / E社 / D社 =C2*2",
            "顧客 / D社 / B社 / E社 mode=0",
            "true A3:A4",
            "顧客|担当|単価|数量 / B社|鈴木|250.50|4 / E社|山田|20|2",
            "true true",
            "Sheet1 (2) D社 =C2*2 Sheet1 (2)",
            "B社 =D3*2 #,##0.00",
            "1 Book.xlsx",
        ]
    );
}

/// `objects/Sheet1.ras`: the links of the sheet followed.
const LINKS_SHEET1: &str = r#"
@on("Worksheet.FollowHyperlink")
function followed(target: Excel.Hyperlink) {
    console.log(`FollowHyperlink ${target.range.address(false, false)} ${target.address}|${target.subAddress}`);
}

@on("Worksheet.SelectionChange")
function selected(target: Excel.Range) {
    console.log(`SelectionChange ${target.address(false, false)}`);
}

@on("Worksheet.Change")
function changed(target: Excel.Range) {
    console.log(`Change ${target.address(false, false)}`);
}
"#;

const LINKS_BOOK: &str = r#"
@on("Workbook.SheetFollowHyperlink")
function anyLink(sh: object, target: Excel.Hyperlink) {
    console.log(`SheetFollowHyperlink ${sh.name} ${target.textToDisplay}`);
}
"#;

const LINKS: &str = r##"
export function follow() {
    const ws = Sheet1;
    // The link the file has (A1), and one added.
    ws.hyperlinks.add(ws.range("A6"), "", "Sheet1!D3", "D3 へ", "合計へ");
    console.log(`count ${ws.hyperlinks.count} ${ws.range("A1:A5").hyperlinks.count}`);
    for (const h of ws.hyperlinks) {
        console.log(`${h.range.address(false, false)} ${h.address}|${h.subAddress}|${h.screenTip}|${h.textToDisplay}`);
    }
    ws.range("A6").hyperlinks(1).follow();
    ws.hyperlinks(1).follow();
    thisWorkbook.followHyperlink("#Sheet1!B2");
    using _ = fastMode();
    ws.hyperlinks(2).follow();
}

export function edit() {
    const ws = Sheet1;
    const h = ws.hyperlinks(1);
    h.address = "https://example.com/b";
    h.screenTip = "ヒント";
    ws.range("B6").hyperlinks.delete();
    ws.hyperlinks.add(ws.range("C6"), "https://example.com/c");
    ws.range("C6").clearHyperlinks();
    ws.rows(1).insert();
    console.log(`after ${ws.hyperlinks.count} ${ws.hyperlinks(1).range.address(false, false)} ${ws.hyperlinks(1).screenTip}`);
    try {
        console.log(ws.range("A1").hyperlinks(1).address);
    } catch (e) {
        console.log(`error ${e.code}`);
    }
}
"##;

#[test]
fn hyperlinks_are_followed_with_their_event() {
    if !has_node() {
        return;
    }
    let p = project_with(
        "links",
        LINKS,
        &[("Sheet1", LINKS_SHEET1), ("ThisWorkbook", LINKS_BOOK)],
        |b| {
            sales(b);
            b.set_hyperlink(0, CellRef::new(0, 0), "https://example.com/a", "", "")
                .unwrap();
        },
    );
    let (stdout, stderr, ok) = p.run("Main.follow");
    assert!(ok, "{stderr}\n{stdout}");
    assert_eq!(
        lines(&stdout),
        vec![
            // The text to display is written into the cell.
            "Change A6",
            "count 2 1",
            "A1 https://example.com/a|||顧客",
            "A6 |Sheet1!D3|D3 へ|合計へ",
            // A place in the workbook is selected first, then the event.
            "SelectionChange D3",
            "FollowHyperlink A6 |Sheet1!D3",
            "SheetFollowHyperlink Sheet1 合計へ",
            // An address is said (a note), not opened.
            "FollowHyperlink A1 https://example.com/a|",
            "SheetFollowHyperlink Sheet1 顧客",
            // `Workbook.FollowHyperlink` follows no cell's link: no event.
            "SelectionChange B2",
            // Events off: the link is followed, nothing is raised.
        ]
    );
    assert!(
        stdout.contains(
            "[ras] リンク先 https://example.com/a を開きます（Excel のないデバッグ実行では開きません）"
        ),
        "{stdout}"
    );
    // The links are saved with the workbook.
    let saved = p.saved();
    let links = saved.hyperlinks(0, rng("A1:Z100")).unwrap();
    assert_eq!(
        links
            .iter()
            .map(|l| (
                l.at.a1(),
                l.address.clone(),
                l.sub_address.clone(),
                l.tip.clone()
            ))
            .collect::<Vec<_>>(),
        [
            (
                "A1".to_string(),
                "https://example.com/a".to_string(),
                String::new(),
                String::new()
            ),
            (
                "A6".to_string(),
                String::new(),
                "Sheet1!D3".to_string(),
                "D3 へ".to_string()
            )
        ]
    );

    let (stdout, stderr, ok) = p.run("Main.edit");
    assert!(ok, "{stderr}\n{stdout}");
    // The empty cell of a link added shows its address; the link of A1 changed and moved to A2
    // with its row; B6 and C6 had none left.
    assert_eq!(
        lines(&stdout),
        vec!["Change C6", "Change 1:1", "after 1 A2 ヒント", "error 9"]
    );
    let links = p.saved().hyperlinks(0, rng("A1:Z100")).unwrap();
    assert_eq!(links.len(), 1);
    assert_eq!(
        (
            links[0].at.a1(),
            links[0].address.as_str(),
            links[0].tip.as_str()
        ),
        ("A2".to_string(), "https://example.com/b", "ヒント")
    );
}

const EXTEND: &str = r#"
export function extend() {
    const ws = Sheet1;
    const r = ws.range("B3:C3");
    console.log(`${r.extend(1, 1, 1, 1).address(false, false)} ${r.extend(down: 2).address(false, false)} ${r.extend(1).address(false, false)} ${r.extend(left: 1, right: 2).address(false, false)}`);
    // Negative counts shrink it; `undefined` skips a side.
    console.log(`${ws.range("A1:D4").extend(-1, -1, -1, -1).address(false, false)} ${r.extend(undefined, 1).address(false, false)} ${r.extend().address(false, false)}`);
    // The cells from (top, left) to (bottom, right) by numbers.
    console.log(`${ws.range(2, 4, 8, 9).address(false, false)} ${dump(ws.range(1, 1, 2, 2))}`);
    // The headings above the data, with the cells read back.
    console.log(`${dump(ws.range("A2:B2").extend(up: 1))}`);
    try {
        console.log(ws.range("A1").extend(1).address(false, false));
    } catch (e) {
        console.log(`error ${e.code}`);
    }
}
"#;

#[test]
fn ranges_are_extended_on_each_side() {
    if !has_node() {
        return;
    }
    let (_p, out) = run_with("extend", EXTEND, "Main.extend");
    assert_eq!(
        out,
        vec![
            "A2:D4 B3:C5 B2:C3 A3:E3",
            "B2:C3 B3:C4 B3:C3",
            "D2:I8 顧客|担当 / A社|佐藤",
            "顧客|担当 / A社|佐藤",
            "error 1004",
        ]
    );
}
