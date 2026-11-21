//! Working without Excel from the command line: `ras book dump` / `diff`, `ras run` and
//! `ras test` (the roadmap's P0). `ras run` and `ras test` need `node` on PATH and are skipped
//! without it; the `ras_book` addon is built with Cargo for the test run.

use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::OnceLock;

use ras_workbook::book::{Book, Val};
use ras_workbook::RangeRef;

fn has_node() -> bool {
    let found = Command::new("node").arg("--version").output().is_ok_and(|o| o.status.success());
    if !found {
        eprintln!("skipped: node is not installed");
    }
    found
}

/// The `ras_book` addon as Node loads it. `cargo test` does not build the library of another
/// package for an integration test, so it is built here, into the same target directory.
fn addon() -> &'static Path {
    static BUILT: OnceLock<PathBuf> = OnceLock::new();
    BUILT.get_or_init(|| {
        let exe = PathBuf::from(env!("CARGO_BIN_EXE_ras"));
        let profile_dir = exe.parent().expect("the binary is in <target>/<profile>");
        let target_dir = profile_dir.parent().expect("the profile directory is in the target directory");
        let mut cmd = Command::new(env!("CARGO"));
        cmd.args(["build", "-p", "ras-book", "--target-dir"]).arg(target_dir).current_dir(env!("CARGO_MANIFEST_DIR"));
        if profile_dir.file_name().is_some_and(|n| n == "release") {
            cmd.arg("--release");
        }
        let out = cmd.output().expect("cargo");
        assert!(out.status.success(), "cargo build -p ras-book failed:\n{}", String::from_utf8_lossy(&out.stderr));
        let name = if cfg!(windows) {
            "ras_book.dll"
        } else if cfg!(target_os = "macos") {
            "libras_book.dylib"
        } else {
            "libras_book.so"
        };
        let node = std::env::temp_dir().join(format!("ras-offline-test-{}.node", std::process::id()));
        std::fs::copy(profile_dir.join(name), &node).expect("the addon was built");
        node
    })
}

/// Runs `ras`; returns the exit code, stdout and stderr.
fn ras(dir: &Path, args: &[&str]) -> (i32, String, String) {
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_ras"));
    cmd.args(args).current_dir(dir).env_remove("RAS_NODE");
    if args.first().is_some_and(|a| *a == "run" || *a == "test") {
        cmd.env("RAS_BOOK_ADDON", addon());
    }
    let out = cmd.output().expect("run ras");
    (out.status.code().unwrap_or(-1), String::from_utf8_lossy(&out.stdout).into_owned(), String::from_utf8_lossy(&out.stderr).into_owned())
}

fn temp_dir(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("ras-offline-test-{name}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

fn at(a1: &str) -> RangeRef {
    RangeRef::parse(a1).unwrap()
}

fn text(s: &str) -> Val {
    Val::Text(s.into())
}

/// A workbook with a header, three rows of data and a formula in each.
fn sales() -> Book {
    let mut b = Book::new_book();
    b.set_values(0, at("A1:C1"), &[vec![text("商品"), text("数量"), text("金額")]]).unwrap();
    b.set_values(0, at("A2:B4"), &[vec![text("りんご"), Val::Num(3.0)], vec![text("みかん"), Val::Num(5.0)], vec![text("ぶどう"), Val::Num(2.0)]]).unwrap();
    b.set_formulas(0, at("C2:C4"), &[vec![text("=B2*100")]], false).unwrap();
    b.set_formulas(0, at("C5"), &[vec![text("=SUM(C2:C4)")]], false).unwrap();
    b.add_name("合計", None, "Sheet1!$C$5").unwrap();
    b
}

/// A project with one workbook `Book.xlsx` and the given files.
fn project(name: &str, book: &mut Book, files: &[(&str, &str)]) -> PathBuf {
    let dir = temp_dir(name);
    std::fs::write(dir.join("ras.toml"), "[project]\nname = \"Sales\"\nworkbook = \"Book.xlsx\"\n").unwrap();
    book.save(&dir.join("Book.xlsx")).unwrap();
    for (path, text) in files {
        let p = dir.join(path);
        std::fs::create_dir_all(p.parent().unwrap()).unwrap();
        std::fs::write(p, text).unwrap();
    }
    dir
}

const REPORT: &str = r#"export function rebuildReport() {
    const ws = Sheet1;
    const last = ws.lastRow();
    for (let r = 2; r <= last; r++) {
        ws.cells(r, 4).value = ws.cells(r, 2).value * 2;
    }
    debug.print("rows", last - 1);
}

export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

export function ask(): string {
    if (msgBox("続けますか", vbYesNo) != vbYes) {
        return "";
    }
    const month = inputBox("月", "入力", "2026/1");
    Sheet1.range("F1").value = "'" + month;
    return month;
}

export function askAndPrint() {
    debug.print("answer", ask());
}

export function outOfRange() {
    const a: int[] = [1, 2];
    Sheet1.range("F2").value = "before";
    debug.print(a[5]);
}

@on("Sheet1.Change")
function changed(target: Excel.Range) {
    if (target.address == "$H$1") {
        throw new Error("H1 は書き換えられません");
    }
}
"#;

#[test]
fn book_dump_writes_the_outline_and_the_cells() {
    let dir = temp_dir("dump");
    sales().save(&dir.join("Book.xlsx")).unwrap();

    let (code, out, err) = ras(&dir, &["book", "dump", "Book.xlsx"]);
    assert_eq!(code, 0, "{err}");
    assert!(out.contains("1. Sheet1（コード名 Sheet1）  使用範囲 A1:C5、値のあるセル 13 個（うち数式 4 個）"), "{out}");
    // Formulas filled down are written once.
    assert!(out.contains("数式: C2:C4 =B2*100（同じ形の数式 3 個）"), "{out}");
    assert!(out.contains("数式: C5 =SUM(C2:C4)"), "{out}");
    assert!(out.contains("合計 = Sheet1!$C$5"), "{out}");

    let (code, out, err) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "Sheet1", "--range", "A1:C2"]);
    assert_eq!(code, 0, "{err}");
    assert!(out.contains("A1\t\"商品\"\n"), "{out}");
    assert!(out.contains("B2\t3\n"), "{out}");
    assert!(out.contains("C2\t=B2*100\t→ 300\n"), "{out}");
    assert!(!out.contains("A3"), "{out}");

    // A sheet by its number, as CSV and as Markdown.
    let (_, out, _) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "1", "--range", "A1:C2", "--format", "csv"]);
    assert_eq!(out, "商品,数量,金額\nりんご,3,300\n");
    let (_, out, _) = ras(&dir, &["book", "dump", "Book.xlsx", "--range", "Sheet1!A1:B2", "--format", "markdown"]);
    assert!(out.contains("| | A | B |\n|---|---|---|\n| 1 | 商品 | 数量 |\n| 2 | りんご | 3 |\n"), "{out}");

    let (_, out, _) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "Sheet1", "--format", "json", "--formats"]);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["usedRange"], "A1:C5");
    let c2 = v["cells"].as_array().unwrap().iter().find(|c| c["address"] == "C2").unwrap();
    assert_eq!((c2["formula"].as_str(), c2["value"].as_f64(), c2["type"].as_str(), c2["text"].as_str()), (Some("=B2*100"), Some(300.0), Some("number"), Some("300")));

    // More than the limit: what was written is told, with how to see the rest.
    let (_, out, _) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "Sheet1", "--max-rows", "2"]);
    assert!(out.contains("5 行 × 3 列のうち A1:C2 を出力しました"), "{out}");
    assert!(!out.contains("A3\t"), "{out}");

    let (code, _, err) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "なし"]);
    assert_eq!(code, 2);
    assert!(err.contains("シート `なし` が見つかりません（シート: Sheet1）"), "{err}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn book_diff_tells_cells_and_structure() {
    let dir = temp_dir("diff");
    sales().save(&dir.join("a.xlsx")).unwrap();

    let (code, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "a.xlsx"]);
    assert_eq!((code, out.as_str()), (0, "差分はありません\n"));

    // A value, a formula, a number format, a new sheet and a name.
    let mut b = sales();
    b.set_values(0, at("B3"), &[vec![Val::Num(7.0)]]).unwrap();
    b.set_formulas(0, at("C4"), &[vec![text("=B4*200")]], false).unwrap();
    b.set_number_format(0, at("B2"), "0.00").unwrap();
    let added = b.add_sheet(1, Some("集計")).unwrap();
    b.set_values(added, at("A1"), &[vec![text("計")]]).unwrap();
    b.add_name("税率", None, "0.1").unwrap();
    b.save(&dir.join("b.xlsx")).unwrap();
    let (code, out, err) = ras(&dir, &["book", "diff", "a.xlsx", "b.xlsx"]);
    assert_eq!(code, 1, "{err}");
    assert!(out.contains("+ シート「集計」を追加"), "{out}");
    assert!(out.contains("  B2\t3 → 3\t表示形式 General → 0.00"), "{out}");
    assert!(out.contains("  B3\t5 → 7"), "{out}");
    assert!(out.contains("  C3\t=B3*100（500） → =B3*100（700）"), "{out}");
    assert!(out.contains("  C4\t=B4*100（200） → =B4*200（400）"), "{out}");
    assert!(out.contains("集計:\n  A1\t（空） → \"計\""), "{out}");
    assert!(out.contains("+ 名前 税率 =0.1"), "{out}");
    let (_, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "b.xlsx", "--format", "json"]);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["structure"][0], serde_json::json!({ "kind": "sheetAdded", "sheet": "集計" }));
    let b3 = v["cells"].as_array().unwrap().iter().find(|c| c["address"] == "B3").unwrap();
    assert_eq!((b3["before"]["value"].as_str(), b3["after"]["value"].as_str()), (Some("5"), Some("7")));

    // A row inserted: told as that, not as every cell below it changed.
    let mut c = Book::new_book();
    c.set_values(0, at("A1:C1"), &[vec![text("商品"), text("数量"), text("金額")]]).unwrap();
    c.set_values(0, at("A2:B5"), &[vec![text("りんご"), Val::Num(3.0)], vec![text("もも"), Val::Num(1.0)], vec![text("みかん"), Val::Num(5.0)], vec![text("ぶどう"), Val::Num(2.0)]]).unwrap();
    c.set_formulas(0, at("C2:C5"), &[vec![text("=B2*100")]], false).unwrap();
    c.set_formulas(0, at("C6"), &[vec![text("=SUM(C2:C5)")]], false).unwrap();
    c.add_name("合計", None, "Sheet1!$C$6").unwrap();
    c.save(&dir.join("c.xlsx")).unwrap();
    let (code, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "c.xlsx"]);
    assert_eq!(code, 1);
    assert!(out.contains("~ Sheet1: 行 3 を挿入（推定）"), "{out}");
    assert!(!out.contains("  A4\t"), "the rows below only moved: {out}");
    assert!(out.contains("~ 名前 合計 =Sheet1!$C$5 → =Sheet1!$C$6"), "{out}");

    // A sheet with the same cells under another name was renamed.
    let mut d = sales();
    d.rename_sheet(0, "売上").unwrap();
    d.save(&dir.join("d.xlsx")).unwrap();
    let (_, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "d.xlsx"]);
    assert!(out.contains("~ シート「Sheet1」の名前を「売上」に変更"), "{out}");
    assert!(!out.contains("  A1\t"), "{out}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn run_reports_output_dialogs_errors_and_changed_cells() {
    if !has_node() {
        return;
    }
    let dir = project("run", &mut sales(), &[("src/report.ras", REPORT)]);

    let (code, out, err) = ras(&dir, &["run", "--list"]);
    assert_eq!(code, 0, "{err}");
    assert_eq!(out, "Report.rebuildReport\nReport.askAndPrint\nReport.outOfRange\nReport.ask\n", "{out}");

    let (code, out, err) = ras(&dir, &["run", "Report.rebuildReport"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(out.contains("rows 3\n"), "{out}");
    assert!(out.contains("変更されたセル（Book.xlsx → build/debug/workbooks/Book.xlsx）:"), "{out}");
    assert!(out.contains("  D2\t（空） → 6"), "{out}");
    assert!(out.contains("  D4\t（空） → 4"), "{out}");
    // The project's workbook is not written.
    let (code, _, _) = ras(&dir, &["book", "diff", "Book.xlsx", "build/debug/workbooks/Book.xlsx"]);
    assert_eq!(code, 1);
    assert!(Book::open(&dir.join("Book.xlsx")).unwrap().cells(0, at("D1:D9")).unwrap().is_empty());

    // The answers of the dialogs come from the command line, in order.
    let (code, out, err) = ras(&dir, &["run", "askAndPrint", "--msgbox", "vbYes", "--inputbox", "2026/4"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(out.contains("[MsgBox] 続けますか → vbYes\n"), "{out}");
    assert!(out.contains("answer 2026/4\n"), "{out}");
    assert!(out.contains("  F1\t（空） → \"2026/4\""), "{out}");

    // Without them, the defaults (and that is told).
    let (code, out, _) = ras(&dir, &["run", "askAndPrint", "--json"]);
    assert_eq!(code, 0, "{out}");
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["ok"], true);
    assert_eq!(v["printed"], serde_json::json!(["answer "]));
    assert_eq!(v["dialogs"][0]["kind"], "msgBox");
    assert_eq!((v["dialogs"][0]["answer"].as_i64(), v["dialogs"][0]["answered"].as_bool()), (Some(1), Some(false)));
    assert_eq!(v["workbooks"], serde_json::json!([]));

    // A run-time error: its number, description and place; what was written before it.
    let (code, out, err) = ras(&dir, &["run", "outOfRange"]);
    assert_eq!(code, 1, "{out}{err}");
    assert!(err.contains("実行時エラー 9: インデックスが有効範囲にありません。"), "{err}");
    assert!(err.contains("（src/report.ras:30）"), "{err}");
    assert!(out.contains("  F2\t（空） → \"before\""), "{out}");
    let (code, out, _) = ras(&dir, &["run", "outOfRange", "--json"]);
    assert_eq!(code, 1);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!((v["ok"].as_bool(), v["error"]["code"].as_i64(), v["error"]["file"].as_str(), v["error"]["line"].as_i64()), (Some(false), Some(9), Some("src/report.ras"), Some(30)));
    assert_eq!(v["workbooks"][0]["diff"]["cells"][0]["address"], "F2");

    let (code, _, err) = ras(&dir, &["run", "nothing"]);
    assert_eq!(code, 1, "{err}");
    assert!(err.contains("実行する関数 `nothing` が見つかりません"), "{err}");
    let (code, _, err) = ras(&dir, &["run"]);
    assert_eq!(code, 2);
    assert!(err.contains("実行する関数を指定してください（候補: Report.rebuildReport"), "{err}");
    let _ = std::fs::remove_dir_all(&dir);
}

const REPORT_TEST: &str = r#"import { rebuildReport, lastRow, ask } from "../src/report";
import { twice } from "./helper";

let runs = 0;

describe("lastRow", () => {
    test("データの最終行を返す", () => {
        expect(Sheet1.lastRow()).toBe(4);
    });

    test.each([[1, 4], [3, 5], [9, 1]])("列 %i の最終行は %i", (col: int, want: int) => {
        expect(Sheet1.lastRow(col)).toBe(want);
    });
});

describe("rebuildReport", () => {
    let last: int;

    beforeAll(() => {
        Sheet1.range("G1").value = "prepared";
        last = Sheet1.lastRow();
    });

    test("D 列に数量の 2 倍を書く", () => {
        rebuildReport();
        runs++;
        expect(Sheet1.cells(2, 4).value).toBe(6);
        expect(Sheet1.cells(last, 4).value).toBe(twice(2));
        expect(Sheet1.range("D2:D4").value).toEqual([[6], [10], [4]]);
        expect(printed()).toEqual(["rows 3"]);
    });

    test("前のテストの変更は残らず、beforeAll の準備は残る", () => {
        expect(Sheet1.cells(2, 4).value).toBeEmpty();
        expect(Sheet1.range("G1").value).toBe("prepared");
        // Module variables stay, as the variables of the file do.
        expect(runs).toBe(1);
    });

    test("失敗するテスト", () => {
        rebuildReport();
        expect(Sheet1.cells(2, 4).value).toBe(7);
    });

    test.skip("あとで直す", () => {
        expect(1).toBe(2);
    });
    test.todo("空のシート");
});

describe("ダイアログ", () => {
    test("置き換えた応答を返す", () => {
        const box = mockMsgBox((prompt, buttons) => vbYes);
        const input = mockInputBox(() => "2026/5");
        expect(ask()).toBe("2026/5");
        expect(box.calls.length).toBe(1);
        expect(box.calls[0].prompt).toContain("続けます");
        expect(input.calls[0].defaultValue).toBe("2026/1");
        expect(Sheet1.range("F1").value).toBe("2026/5");
    });

    test("置き換えは次のテストに残らない", () => {
        expect(ask()).toBe("");
    });
});
"#;

const MORE_TEST: &str = r#"struct Point { x: int; y: int; }

function origin(): Point {
    return { x: 0, y: 0 };
}

describe("マッチャ", () => {
    test("値", () => {
        expect(1 + 1).toBe(2);
        expect(1 + 1).not.toBe(3);
        expect(0.1 + 0.2).toBeCloseTo(0.3);
        expect(3).toBeGreaterThan(2);
        expect(3).toBeLessThanOrEqual(3);
        expect("abc").toContain("b");
        expect([1, 2, 3]).toContain(2);
        expect([1, 2, 3]).toHaveLength(3);
        expect("abc-123").toMatch("^[a-z]+-\\d+$");
        expect(origin()).toEqual({ x: 0, y: 0 });
        expect(origin()).not.toEqual({ x: 0, y: 1 });
        expect(Sheet1.range("Z99").value).toBeEmpty();
        expect(1).toBeTruthy();
        expect(0).toBeFalsy();
    });

    test("エラー", () => {
        const a: int[] = [1];
        expect(() => a[3]).toThrow();
        expect(() => a[3]).toThrow(9);
        expect(() => a[0]).not.toThrow();
        expect(() => { throw new Error("だめです"); }).toThrow("だめ");
    });

    test("try の中で失敗したマッチャも失敗になる", () => {
        try {
            expect(1).toBe(2);
        } catch (e) {
            debug.print("caught");
        }
    });
});

describe("置き換え", () => {
    test("手書き VBA と COM オブジェクト", () => {
        const run = mockRun("Legacy", (a: int, b: int) => a + b);
        expect(application.run("Legacy", 2, 3)).toBe(5);
        expect(run.calls.length).toBe(1);
        const made = mockCreateObject("ADODB.Connection", () => new Scripting.Dictionary());
        const conn = createObject("ADODB.Connection");
        expect(made.calls.length).toBe(1);
        expect(conn == null).toBeFalsy();
    });

    test("置き換えていない手書き VBA はエラー", () => {
        expect(() => application.run("Legacy")).toThrow(445);
    });
});

describe("ブックとファイル", () => {
    test("useWorkbook はブックの中身を置き換える", () => {
        const wb = useWorkbook("test/fixtures/other.xlsx");
        expect(wb.name).toBe("Book.xlsx");
        expect(Sheet1.range("A1").value).toBe("fixture");
        expect(Sheet1.name).toBe("データ");
    });

    test("次のテストでは元に戻る", () => {
        expect(Sheet1.range("A1").value).toBe("商品");
        expect(workbooks.count).toBe(1);
    });

    test("tempDir は空のフォルダー", () => {
        const fso = new Scripting.FileSystemObject();
        const dir = tempDir();
        expect(fso.folderExists(dir)).toBeTruthy();
        fso.createTextFile(fso.buildPath(dir, "a.txt")).close();
        expect(fso.getFolder(dir).files.count).toBe(1);
    });

    test("イベント ハンドラのエラーはテストを失敗にする", () => {
        Sheet1.range("H1").value = 1;
    });
});

describe("フック", () => {
    let order = "";
    beforeEach(() => { order = order + "a"; });
    afterEach(() => { order = order + "z"; });

    describe("内側", () => {
        beforeEach(() => { order = order + "b"; });
        afterEach(() => { order = order + "y"; });
        test("外側から順に実行する", () => {
            expect(order).toBe("ab");
        });
        test("afterEach は内側から順に実行する", () => {
            expect(order).toBe("abyzab");
        });
    });

    describe("beforeAll の失敗", () => {
        beforeAll(() => {
            throw new Error("準備できません");
        });
        test("実行されない", () => {
            debug.print("not reached");
        });
    });
});
"#;

/// The results of a run with `--reporter json`, by full name.
fn results(out: &str) -> Vec<serde_json::Value> {
    out.lines().filter_map(|l| serde_json::from_str::<serde_json::Value>(l).ok()).collect()
}

fn result<'a>(events: &'a [serde_json::Value], name: &str) -> &'a serde_json::Value {
    events
        .iter()
        .find(|e| e["type"] == "test" && e["name"].as_array().is_some_and(|n| n.iter().filter_map(|x| x.as_str()).collect::<Vec<_>>().join(" > ") == name))
        .unwrap_or_else(|| panic!("no result for `{name}` in {events:#?}"))
}

fn test_project(name: &str) -> PathBuf {
    let dir = project(
        name,
        &mut sales(),
        &[
            ("src/report.ras", REPORT),
            ("test/helper.ras", "export function twice(n: int): int {\n    return n * 2;\n}\n"),
            ("test/report.test.ras", REPORT_TEST),
            ("test/more.test.ras", MORE_TEST),
        ],
    );
    let mut other = Book::new_book();
    other.rename_sheet(0, "データ").unwrap();
    other.set_values(0, at("A1"), &[vec![text("fixture")]]).unwrap();
    std::fs::create_dir_all(dir.join("test/fixtures")).unwrap();
    other.save(&dir.join("test/fixtures/other.xlsx")).unwrap();
    dir
}

#[test]
fn test_lists_runs_and_reports() {
    if !has_node() {
        return;
    }
    let dir = test_project("test");

    // The list comes from the syntax: nothing runs.
    let (code, out, err) = ras(&dir, &["test", "--list", "test/report.test.ras"]);
    assert_eq!(code, 0, "{err}");
    assert!(out.starts_with("test/report.test.ras\n  lastRow\n    データの最終行を返す\n    列 1 の最終行は 4\n    列 3 の最終行は 5\n    列 9 の最終行は 1\n  rebuildReport\n"), "{out}");
    assert!(out.contains("    あとで直す（skip）\n    空のシート（todo）\n"), "{out}");
    let (_, out, _) = ras(&dir, &["test", "--list", "--reporter", "json"]);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["files"][0]["file"], "test/more.test.ras");
    assert_eq!(v["files"][1]["tests"][0], serde_json::json!({ "type": "describe", "name": "lastRow", "line": 6, "column": 1, "children": [
        { "type": "test", "name": "データの最終行を返す", "line": 7, "column": 5 },
        { "type": "test", "name": "列 1 の最終行は 4", "line": 11, "column": 5 },
        { "type": "test", "name": "列 3 の最終行は 5", "line": 11, "column": 5 },
        { "type": "test", "name": "列 9 の最終行は 1", "line": 11, "column": 5 },
    ] }));

    let (code, out, err) = ras(&dir, &["test", "test/report.test.ras"]);
    assert_eq!(code, 1, "{out}{err}");
    assert!(out.contains("  ✓ lastRow > データの最終行を返す"), "{out}");
    assert!(out.contains("  ✓ lastRow > 列 3 の最終行は 5"), "{out}");
    assert!(out.contains("  ✓ rebuildReport > 前のテストの変更は残らず、beforeAll の準備は残る"), "{out}");
    assert!(out.contains("  ✗ rebuildReport > 失敗するテスト"), "{out}");
    assert!(out.contains("      toBe: 期待値 7、実際の値 6\n      at test/report.test.ras:42:9\n      debug.print:\n        rows 3\n      ブック: build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx"), "{out}");
    assert!(out.contains("  - rebuildReport > あとで直す（スキップ）"), "{out}");
    assert!(out.contains("  □ rebuildReport > 空のシート（未実装）"), "{out}");
    assert!(out.contains("  ✓ ダイアログ > 置き換えた応答を返す"), "{out}");
    // A dialog that is not replaced gets its default answer, and that is told.
    assert!(out.contains("  ✓ ダイアログ > 置き換えは次のテストに残らない"), "{out}");
    assert!(out.contains("警告: msgBox（\"続けますか\"）は置き換えられていません"), "{out}");
    assert!(out.contains("テスト 11 件: 成功 8、失敗 1、スキップ 1、未実装 1"), "{out}");
    // The workbook of the failed test holds what the test wrote; the project's is untouched.
    let mut failed = Book::open(&dir.join("build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx")).unwrap();
    assert_eq!(failed.get(0, at("D2"), ras_workbook::book::Read::Value).unwrap(), vec![vec![Val::Num(6.0)]]);
    assert!(Book::open(&dir.join("Book.xlsx")).unwrap().cells(0, at("D1:G9")).unwrap().is_empty());
    // Tests are no part of what is built into the workbook.
    let (code, out, err) = ras(&dir, &["build"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(dir.join("build/Report.bas").is_file() && !dir.join("build/T_test_helper.bas").exists());
    let (code, out, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{out}{err}");

    // By name; as JSON lines.
    let (code, out, err) = ras(&dir, &["test", "-t", "失敗する", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}{err}");
    let events = results(&out);
    assert_eq!(events.iter().filter(|e| e["type"] == "test").count(), 1, "{out}");
    let failed = result(&events, "rebuildReport > 失敗するテスト");
    assert_eq!(failed["file"], "test/report.test.ras");
    assert_eq!((failed["status"].as_str(), failed["line"].as_i64()), (Some("failed"), Some(40)));
    assert_eq!(failed["failure"], serde_json::json!({
        "kind": "expect", "matcher": "toBe", "expected": "7", "actual": "6", "message": "期待値 7、実際の値 6",
        "location": { "file": "test/report.test.ras", "line": 42, "column": 9 },
    }));
    assert_eq!(failed["printed"], serde_json::json!(["rows 3"]));
    assert_eq!(failed["workbook"], "build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx");
    let summary = events.iter().find(|e| e["type"] == "summary").unwrap();
    assert_eq!((summary["tests"].as_i64(), summary["failed"].as_i64(), summary["passed"].as_i64()), (Some(1), Some(1), Some(0)));

    // --ci: a dialog that is not replaced fails its test.
    let (code, out, _) = ras(&dir, &["test", "test/report.test.ras", "-t", "ダイアログ", "--ci", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}");
    let events = results(&out);
    assert_eq!(result(&events, "ダイアログ > 置き換えた応答を返す")["status"], "passed");
    let unanswered = result(&events, "ダイアログ > 置き換えは次のテストに残らない");
    assert_eq!((unanswered["status"].as_str(), unanswered["failure"]["kind"].as_str()), (Some("failed"), Some("error")));

    // JUnit XML.
    let (code, out, _) = ras(&dir, &["test", "test/report.test.ras", "-t", "rebuildReport", "--reporter", "junit"]);
    assert_eq!(code, 1);
    assert!(out.contains("<testsuite name=\"test/report.test.ras\" tests=\"5\" failures=\"1\" skipped=\"2\""), "{out}");
    assert!(out.contains("<failure message=\"toBe: 期待値 7、実際の値 6\">"), "{out}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_matchers_replacements_hooks_and_the_workbook() {
    if !has_node() {
        return;
    }
    let dir = test_project("test-more");
    let (code, out, err) = ras(&dir, &["test", "test/more.test.ras", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}{err}");
    let events = results(&out);
    for name in [
        "マッチャ > 値",
        "マッチャ > エラー",
        "置き換え > 手書き VBA と COM オブジェクト",
        "置き換え > 置き換えていない手書き VBA はエラー",
        "ブックとファイル > useWorkbook はブックの中身を置き換える",
        "ブックとファイル > 次のテストでは元に戻る",
        "ブックとファイル > tempDir は空のフォルダー",
        "フック > 内側 > 外側から順に実行する",
        "フック > 内側 > afterEach は内側から順に実行する",
    ] {
        let r = result(&events, name);
        assert_eq!(r["status"], "passed", "{name}: {r:#}");
    }
    // A failed matcher is not caught by the program's own `try`.
    let caught = result(&events, "マッチャ > try の中で失敗したマッチャも失敗になる");
    assert_eq!((caught["status"].as_str(), caught["failure"]["matcher"].as_str()), (Some("failed"), Some("toBe")));
    assert_eq!(caught["printed"], serde_json::json!(["caught"]));
    // An error an event handler does not catch fails the test (spec §14.12.5).
    let event = result(&events, "ブックとファイル > イベント ハンドラのエラーはテストを失敗にする");
    assert_eq!(event["status"], "failed");
    assert!(event["failure"]["description"].as_str().unwrap().contains("Sheet1.Change"), "{event:#}");
    // A failing `beforeAll` fails its tests without running them.
    let hook = result(&events, "フック > beforeAll の失敗 > 実行されない");
    assert_eq!((hook["status"].as_str(), hook["failure"]["kind"].as_str(), hook["failure"]["hook"].as_str()), (Some("failed"), Some("hook"), Some("beforeAll")));
    assert_eq!(hook["printed"], serde_json::json!([]));
    let summary = events.iter().find(|e| e["type"] == "summary").unwrap();
    assert_eq!((summary["passed"].as_i64(), summary["failed"].as_i64()), (Some(9), Some(3)), "{out}");

    // `event-errors = "ignore"`: the error is told, the test passes.
    std::fs::write(dir.join("ras.toml"), "[project]\nname = \"Sales\"\nworkbook = \"Book.xlsx\"\n\n[test]\nevent-errors = \"ignore\"\n").unwrap();
    let (_, out, _) = ras(&dir, &["test", "test/more.test.ras", "-t", "イベント", "--reporter", "json"]);
    let events = results(&out);
    let event = result(&events, "ブックとファイル > イベント ハンドラのエラーはテストを失敗にする");
    assert_eq!(event["status"], "passed", "{event:#}");
    assert!(event["warnings"][0].as_str().unwrap().contains("H1 は書き換えられません"), "{event:#}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_time_outs_only_and_what_test_files_may_hold() {
    if !has_node() {
        return;
    }
    let slow = "test(\"速い\", () => {\n    expect(1).toBe(1);\n});\ntest(\"終わらない\", () => {\n    let i = 0;\n    while (true) {\n        i = (i + 1) % 10;\n    }\n}, { timeout: 300 });\ntest(\"その後\", () => {\n    expect(1).toBe(1);\n});\n";
    let only = "test(\"a\", () => {\n    expect(1).toBe(2);\n});\ntest.only(\"b\", () => {\n    expect(1).toBe(1);\n});\ndescribe.skip(\"s\", () => {\n    test(\"c\", () => {\n        expect(1).toBe(2);\n    });\n});\n";
    let dir = project("test-limits", &mut sales(), &[("src/report.ras", REPORT), ("test/slow.test.ras", slow), ("test/only.test.ras", only)]);

    // The file is given up at the time-out: the tests after it fail too (spec §14.12.4).
    let (code, out, err) = ras(&dir, &["test", "test/slow.test.ras", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}{err}");
    let events = results(&out);
    assert_eq!(result(&events, "速い")["status"], "passed");
    let slow = result(&events, "終わらない");
    assert_eq!((slow["status"].as_str(), slow["failure"]["kind"].as_str()), (Some("failed"), Some("timeout")), "{slow:#}");
    assert_eq!(result(&events, "その後")["status"], "failed");

    // `only` runs alone; with --ci it is an error.
    let (code, out, err) = ras(&dir, &["test", "test/only.test.ras", "--reporter", "json"]);
    assert_eq!(code, 0, "{out}{err}");
    let events = results(&out);
    assert_eq!((result(&events, "a")["status"].as_str(), result(&events, "b")["status"].as_str(), result(&events, "s > c")["status"].as_str()), (Some("skipped"), Some("passed"), Some("skipped")));
    let (code, _, err) = ras(&dir, &["test", "test/only.test.ras", "--ci"]);
    assert_eq!(code, 2);
    assert!(err.contains("`only` の付いたテストがあります"), "{err}");

    // --bail stops at the first failure.
    std::fs::write(dir.join("test/only.test.ras"), only.replace("test.only", "test")).unwrap();
    let (code, out, _) = ras(&dir, &["test", "--bail", "1", "--reporter", "json"]);
    assert_eq!(code, 1);
    let events = results(&out);
    assert_eq!((result(&events, "a")["status"].as_str(), result(&events, "b")["status"].as_str()), (Some("failed"), Some("skipped")), "{out}");
    assert!(!events.iter().any(|e| e["type"] == "file" && e["file"] == "test/slow.test.ras"), "{out}");

    // What a test file may hold, and who may import whom.
    std::fs::write(
        dir.join("test/only.test.ras"),
        "const name = \"x\";\nSheet1.cells(1, 1).value = 1;\ndescribe(name, () => {});\ndescribe(\"d\", () => {\n    for (let i = 0; i < 2; i++) {}\n});\n@macro\nfunction m() {}\n",
    )
    .unwrap();
    std::fs::write(dir.join("test/helper.ras"), "export function twice(n: int): int {\n    return n * 2;\n}\n").unwrap();
    std::fs::write(dir.join("src/bad.ras"), "import { twice } from \"../test/helper\";\n").unwrap();
    let (code, _, err) = ras(&dir, &["test"]);
    assert_eq!(code, 2, "{err}");
    for expected in ["error[E0701]", "error[E0702]", "test/only.test.ras:2:1", "test/only.test.ras:5:5", "error[E0703]", "test/only.test.ras:3:10", "error[E0704]", "test/only.test.ras:7:1"] {
        assert!(err.contains(expected), "{expected} is missing in:\n{err}");
    }
    // `ras check` reports the tests' problems too.
    let (code, _, err) = ras(&dir, &["check"]);
    assert_eq!(code, 1);
    assert!(err.contains("error[E0702]") && err.contains("error[E0701]"), "{err}");
    let _ = std::fs::remove_dir_all(&dir);
}
