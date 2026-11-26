//! Working without Excel from the command line: `ras book dump` / `diff`, `ras run` and
//! `ras test` (the roadmap's P0). `ras run` and `ras test` need `node` on PATH and are skipped
//! without it; the `ras_book` addon is built with Cargo for the test run.

use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::OnceLock;

use ras_workbook::book::{Book, Val};
use ras_workbook::RangeRef;

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

/// The `ras_book` addon as Node loads it. `cargo test` does not build the library of another
/// package for an integration test, so it is built here, into the same target directory.
fn addon() -> &'static Path {
    static BUILT: OnceLock<PathBuf> = OnceLock::new();
    BUILT.get_or_init(|| {
        let exe = PathBuf::from(env!("CARGO_BIN_EXE_ras"));
        let profile_dir = exe.parent().expect("the binary is in <target>/<profile>");
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
        let node =
            std::env::temp_dir().join(format!("ras-offline-test-{}.node", std::process::id()));
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
    (
        out.status.code().unwrap_or(-1),
        String::from_utf8_lossy(&out.stdout).into_owned(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
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
    b.set_values(
        0,
        at("A1:C1"),
        &[vec![text("商品"), text("数量"), text("金額")]],
    )
    .unwrap();
    b.set_values(
        0,
        at("A2:B4"),
        &[
            vec![text("りんご"), Val::Num(3.0)],
            vec![text("みかん"), Val::Num(5.0)],
            vec![text("ぶどう"), Val::Num(2.0)],
        ],
    )
    .unwrap();
    b.set_formulas(0, at("C2:C4"), &[vec![text("=B2*100")]], false)
        .unwrap();
    b.set_formulas(0, at("C5"), &[vec![text("=SUM(C2:C4)")]], false)
        .unwrap();
    b.add_name("合計", None, "Sheet1!$C$5").unwrap();
    b
}

/// A project with one workbook `Book.xlsx` and the given files.
fn project(name: &str, book: &mut Book, files: &[(&str, &str)]) -> PathBuf {
    let dir = temp_dir(name);
    std::fs::write(
        dir.join("ras.toml"),
        "[project]\nname = \"Sales\"\nworkbook = \"Book.xlsx\"\n",
    )
    .unwrap();
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
"#;

/// `objects/Sheet1.ras`: the module of the sheet.
const SHEET1: &str = r#"@on("Worksheet.Change")
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
    assert!(
        out.contains(
            "1. Sheet1（コード名 Sheet1）  使用範囲 A1:C5、値のあるセル 13 個（うち数式 4 個）"
        ),
        "{out}"
    );
    // Formulas filled down are written once.
    assert!(
        out.contains("数式: C2:C4 =B2*100（同じ形の数式 3 個）"),
        "{out}"
    );
    assert!(out.contains("数式: C5 =SUM(C2:C4)"), "{out}");
    assert!(out.contains("合計 = Sheet1!$C$5"), "{out}");

    let (code, out, err) = ras(
        &dir,
        &[
            "book",
            "dump",
            "Book.xlsx",
            "--sheet",
            "Sheet1",
            "--range",
            "A1:C2",
        ],
    );
    assert_eq!(code, 0, "{err}");
    assert!(out.contains("A1\t\"商品\"\n"), "{out}");
    assert!(out.contains("B2\t3\n"), "{out}");
    assert!(out.contains("C2\t=B2*100\t→ 300\n"), "{out}");
    assert!(!out.contains("A3"), "{out}");

    // A sheet by its number, as CSV and as Markdown.
    let (_, out, _) = ras(
        &dir,
        &[
            "book",
            "dump",
            "Book.xlsx",
            "--sheet",
            "1",
            "--range",
            "A1:C2",
            "--format",
            "csv",
        ],
    );
    assert_eq!(out, "商品,数量,金額\nりんご,3,300\n");
    let (_, out, _) = ras(
        &dir,
        &[
            "book",
            "dump",
            "Book.xlsx",
            "--range",
            "Sheet1!A1:B2",
            "--format",
            "markdown",
        ],
    );
    assert!(
        out.contains("| | A | B |\n|---|---|---|\n| 1 | 商品 | 数量 |\n| 2 | りんご | 3 |\n"),
        "{out}"
    );

    let (_, out, _) = ras(
        &dir,
        &[
            "book",
            "dump",
            "Book.xlsx",
            "--sheet",
            "Sheet1",
            "--format",
            "json",
            "--formats",
        ],
    );
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["usedRange"], "A1:C5");
    let c2 = v["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["address"] == "C2")
        .unwrap();
    assert_eq!(
        (
            c2["formula"].as_str(),
            c2["value"].as_f64(),
            c2["type"].as_str(),
            c2["text"].as_str()
        ),
        (Some("=B2*100"), Some(300.0), Some("number"), Some("300"))
    );

    // More than the limit: what was written is told, with how to see the rest.
    let (_, out, _) = ras(
        &dir,
        &[
            "book",
            "dump",
            "Book.xlsx",
            "--sheet",
            "Sheet1",
            "--max-rows",
            "2",
        ],
    );
    assert!(
        out.contains("5 行 × 3 列のうち A1:C2 を出力しました"),
        "{out}"
    );
    assert!(!out.contains("A3\t"), "{out}");

    let (code, _, err) = ras(&dir, &["book", "dump", "Book.xlsx", "--sheet", "なし"]);
    assert_eq!(code, 2);
    assert!(
        err.contains("シート `なし` が見つかりません（シート: Sheet1）"),
        "{err}"
    );
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
    b.set_formulas(0, at("C4"), &[vec![text("=B4*200")]], false)
        .unwrap();
    b.set_number_format(0, at("B2"), "0.00").unwrap();
    let added = b.add_sheet(1, Some("集計")).unwrap();
    b.set_values(added, at("A1"), &[vec![text("計")]]).unwrap();
    b.add_name("税率", None, "0.1").unwrap();
    b.save(&dir.join("b.xlsx")).unwrap();
    let (code, out, err) = ras(&dir, &["book", "diff", "a.xlsx", "b.xlsx"]);
    assert_eq!(code, 1, "{err}");
    assert!(out.contains("+ シート「集計」を追加"), "{out}");
    assert!(
        out.contains("  B2\t3 → 3\t表示形式 General → 0.00"),
        "{out}"
    );
    assert!(out.contains("  B3\t5 → 7"), "{out}");
    assert!(
        out.contains("  C3\t=B3*100（500） → =B3*100（700）"),
        "{out}"
    );
    assert!(
        out.contains("  C4\t=B4*100（200） → =B4*200（400）"),
        "{out}"
    );
    assert!(out.contains("集計:\n  A1\t（空） → \"計\""), "{out}");
    assert!(out.contains("+ 名前 税率 =0.1"), "{out}");
    let (_, out, _) = ras(
        &dir,
        &["book", "diff", "a.xlsx", "b.xlsx", "--format", "json"],
    );
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(
        v["structure"][0],
        serde_json::json!({ "kind": "sheetAdded", "sheet": "集計" })
    );
    let b3 = v["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["address"] == "B3")
        .unwrap();
    assert_eq!(
        (
            b3["before"]["value"].as_str(),
            b3["after"]["value"].as_str()
        ),
        (Some("5"), Some("7"))
    );

    // A row inserted: told as that, not as every cell below it changed.
    let mut c = Book::new_book();
    c.set_values(
        0,
        at("A1:C1"),
        &[vec![text("商品"), text("数量"), text("金額")]],
    )
    .unwrap();
    c.set_values(
        0,
        at("A2:B5"),
        &[
            vec![text("りんご"), Val::Num(3.0)],
            vec![text("もも"), Val::Num(1.0)],
            vec![text("みかん"), Val::Num(5.0)],
            vec![text("ぶどう"), Val::Num(2.0)],
        ],
    )
    .unwrap();
    c.set_formulas(0, at("C2:C5"), &[vec![text("=B2*100")]], false)
        .unwrap();
    c.set_formulas(0, at("C6"), &[vec![text("=SUM(C2:C5)")]], false)
        .unwrap();
    c.add_name("合計", None, "Sheet1!$C$6").unwrap();
    c.save(&dir.join("c.xlsx")).unwrap();
    let (code, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "c.xlsx"]);
    assert_eq!(code, 1);
    assert!(out.contains("~ Sheet1: 行 3 を挿入（推定）"), "{out}");
    assert!(!out.contains("  A4\t"), "the rows below only moved: {out}");
    assert!(
        out.contains("~ 名前 合計 =Sheet1!$C$5 → =Sheet1!$C$6"),
        "{out}"
    );

    // A sheet with the same cells under another name was renamed.
    let mut d = sales();
    d.rename_sheet(0, "売上").unwrap();
    d.save(&dir.join("d.xlsx")).unwrap();
    let (_, out, _) = ras(&dir, &["book", "diff", "a.xlsx", "d.xlsx"]);
    assert!(
        out.contains("~ シート「Sheet1」の名前を「売上」に変更"),
        "{out}"
    );
    assert!(!out.contains("  A1\t"), "{out}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn run_reports_output_dialogs_errors_and_changed_cells() {
    if !has_node() {
        return;
    }
    let dir = project(
        "run",
        &mut sales(),
        &[("src/report.ras", REPORT), ("objects/Sheet1.ras", SHEET1)],
    );

    let (code, out, err) = ras(&dir, &["run", "--list"]);
    assert_eq!(code, 0, "{err}");
    assert_eq!(
        out, "Report.rebuildReport\nReport.askAndPrint\nReport.outOfRange\nReport.ask\n",
        "{out}"
    );

    let (code, out, err) = ras(&dir, &["run", "Report.rebuildReport"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(out.contains("rows 3\n"), "{out}");
    assert!(
        out.contains("変更されたセル（Book.xlsx → build/debug/workbooks/Book.xlsx）:"),
        "{out}"
    );
    assert!(out.contains("  D2\t（空） → 6"), "{out}");
    assert!(out.contains("  D4\t（空） → 4"), "{out}");
    // The project's workbook is not written.
    let (code, _, _) = ras(
        &dir,
        &[
            "book",
            "diff",
            "Book.xlsx",
            "build/debug/workbooks/Book.xlsx",
        ],
    );
    assert_eq!(code, 1);
    assert!(Book::open(&dir.join("Book.xlsx"))
        .unwrap()
        .cells(0, at("D1:D9"))
        .unwrap()
        .is_empty());

    // The answers of the dialogs come from the command line, in order.
    let (code, out, err) = ras(
        &dir,
        &[
            "run",
            "askAndPrint",
            "--msgbox",
            "vbYes",
            "--inputbox",
            "2026/4",
        ],
    );
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
    assert_eq!(
        (
            v["dialogs"][0]["answer"].as_i64(),
            v["dialogs"][0]["answered"].as_bool()
        ),
        (Some(1), Some(false))
    );
    assert_eq!(v["workbooks"], serde_json::json!([]));

    // A run-time error: its number, description and place; what was written before it.
    let (code, out, err) = ras(&dir, &["run", "outOfRange"]);
    assert_eq!(code, 1, "{out}{err}");
    assert!(
        err.contains("実行時エラー 9: インデックスが有効範囲にありません。"),
        "{err}"
    );
    assert!(err.contains("（src/report.ras:30）"), "{err}");
    assert!(out.contains("  F2\t（空） → \"before\""), "{out}");
    let (code, out, _) = ras(&dir, &["run", "outOfRange", "--json"]);
    assert_eq!(code, 1);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(
        (
            v["ok"].as_bool(),
            v["error"]["code"].as_i64(),
            v["error"]["file"].as_str(),
            v["error"]["line"].as_i64()
        ),
        (Some(false), Some(9), Some("src/report.ras"), Some(30))
    );
    assert_eq!(v["workbooks"][0]["diff"]["cells"][0]["address"], "F2");

    let (code, _, err) = ras(&dir, &["run", "nothing"]);
    assert_eq!(code, 1, "{err}");
    assert!(
        err.contains("実行する関数 `nothing` が見つかりません"),
        "{err}"
    );
    let (code, _, err) = ras(&dir, &["run"]);
    assert_eq!(code, 2);
    assert!(
        err.contains("実行する関数を指定してください（候補: Report.rebuildReport"),
        "{err}"
    );
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
    out.lines()
        .filter_map(|l| serde_json::from_str::<serde_json::Value>(l).ok())
        .collect()
}

fn result<'a>(events: &'a [serde_json::Value], name: &str) -> &'a serde_json::Value {
    events
        .iter()
        .find(|e| {
            e["type"] == "test"
                && e["name"].as_array().is_some_and(|n| {
                    n.iter()
                        .filter_map(|x| x.as_str())
                        .collect::<Vec<_>>()
                        .join(" > ")
                        == name
                })
        })
        .unwrap_or_else(|| panic!("no result for `{name}` in {events:#?}"))
}

fn test_project(name: &str) -> PathBuf {
    let dir = project(
        name,
        &mut sales(),
        &[
            ("src/report.ras", REPORT),
            ("objects/Sheet1.ras", SHEET1),
            (
                "test/helper.ras",
                "export function twice(n: int): int {\n    return n * 2;\n}\n",
            ),
            ("test/report.test.ras", REPORT_TEST),
            ("test/more.test.ras", MORE_TEST),
        ],
    );
    let mut other = Book::new_book();
    other.rename_sheet(0, "データ").unwrap();
    other
        .set_values(0, at("A1"), &[vec![text("fixture")]])
        .unwrap();
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
    assert!(
        out.contains("    あとで直す（skip）\n    空のシート（todo）\n"),
        "{out}"
    );
    let (_, out, _) = ras(&dir, &["test", "--list", "--reporter", "json"]);
    let v: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(v["files"][0]["file"], "test/more.test.ras");
    assert_eq!(
        v["files"][1]["tests"][0],
        serde_json::json!({ "type": "describe", "name": "lastRow", "line": 6, "column": 1, "children": [
        { "type": "test", "name": "データの最終行を返す", "line": 7, "column": 5 },
        { "type": "test", "name": "列 1 の最終行は 4", "line": 11, "column": 5 },
        { "type": "test", "name": "列 3 の最終行は 5", "line": 11, "column": 5 },
        { "type": "test", "name": "列 9 の最終行は 1", "line": 11, "column": 5 },
    ] })
    );

    let (code, out, err) = ras(&dir, &["test", "test/report.test.ras"]);
    assert_eq!(code, 1, "{out}{err}");
    assert!(out.contains("  ✓ lastRow > データの最終行を返す"), "{out}");
    assert!(out.contains("  ✓ lastRow > 列 3 の最終行は 5"), "{out}");
    assert!(
        out.contains("  ✓ rebuildReport > 前のテストの変更は残らず、beforeAll の準備は残る"),
        "{out}"
    );
    assert!(out.contains("  ✗ rebuildReport > 失敗するテスト"), "{out}");
    assert!(out.contains("      toBe: 期待値 7、実際の値 6\n      at test/report.test.ras:42:9\n      debug.print:\n        rows 3\n      ブック: build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx"), "{out}");
    assert!(
        out.contains("  - rebuildReport > あとで直す（スキップ）"),
        "{out}"
    );
    assert!(
        out.contains("  □ rebuildReport > 空のシート（未実装）"),
        "{out}"
    );
    assert!(
        out.contains("  ✓ ダイアログ > 置き換えた応答を返す"),
        "{out}"
    );
    // A dialog that is not replaced gets its default answer, and that is told.
    assert!(
        out.contains("  ✓ ダイアログ > 置き換えは次のテストに残らない"),
        "{out}"
    );
    assert!(
        out.contains("警告: msgBox（\"続けますか\"）は置き換えられていません"),
        "{out}"
    );
    assert!(
        out.contains("テスト 11 件: 成功 8、失敗 1、スキップ 1、未実装 1"),
        "{out}"
    );
    // The workbook of the failed test holds what the test wrote; the project's is untouched.
    let mut failed = Book::open(
        &dir.join("build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx"),
    )
    .unwrap();
    assert_eq!(
        failed
            .get(0, at("D2"), ras_workbook::book::Read::Value)
            .unwrap(),
        vec![vec![Val::Num(6.0)]]
    );
    assert!(Book::open(&dir.join("Book.xlsx"))
        .unwrap()
        .cells(0, at("D1:G9"))
        .unwrap()
        .is_empty());
    // Tests are no part of what is built into the workbook.
    let (code, out, err) = ras(&dir, &["build"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(
        dir.join("build/Report.bas").is_file() && !dir.join("build/T_test_helper.bas").exists()
    );
    let (code, out, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{out}{err}");

    // By name; as JSON lines.
    let (code, out, err) = ras(&dir, &["test", "-t", "失敗する", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}{err}");
    let events = results(&out);
    assert_eq!(
        events.iter().filter(|e| e["type"] == "test").count(),
        1,
        "{out}"
    );
    let failed = result(&events, "rebuildReport > 失敗するテスト");
    assert_eq!(failed["file"], "test/report.test.ras");
    assert_eq!(
        (failed["status"].as_str(), failed["line"].as_i64()),
        (Some("failed"), Some(40))
    );
    assert_eq!(
        failed["failure"],
        serde_json::json!({
            "kind": "expect", "matcher": "toBe", "expected": "7", "actual": "6", "message": "期待値 7、実際の値 6",
            "location": { "file": "test/report.test.ras", "line": 42, "column": 9 },
        })
    );
    assert_eq!(failed["printed"], serde_json::json!(["rows 3"]));
    assert_eq!(
        failed["workbook"],
        "build/test/workbooks/report.test/rebuildReport _ 失敗するテスト.xlsx"
    );
    let summary = events.iter().find(|e| e["type"] == "summary").unwrap();
    assert_eq!(
        (
            summary["tests"].as_i64(),
            summary["failed"].as_i64(),
            summary["passed"].as_i64()
        ),
        (Some(1), Some(1), Some(0))
    );

    // --ci: a dialog that is not replaced fails its test.
    let (code, out, _) = ras(
        &dir,
        &[
            "test",
            "test/report.test.ras",
            "-t",
            "ダイアログ",
            "--ci",
            "--reporter",
            "json",
        ],
    );
    assert_eq!(code, 1, "{out}");
    let events = results(&out);
    assert_eq!(
        result(&events, "ダイアログ > 置き換えた応答を返す")["status"],
        "passed"
    );
    let unanswered = result(&events, "ダイアログ > 置き換えは次のテストに残らない");
    assert_eq!(
        (
            unanswered["status"].as_str(),
            unanswered["failure"]["kind"].as_str()
        ),
        (Some("failed"), Some("error"))
    );

    // JUnit XML.
    let (code, out, _) = ras(
        &dir,
        &[
            "test",
            "test/report.test.ras",
            "-t",
            "rebuildReport",
            "--reporter",
            "junit",
        ],
    );
    assert_eq!(code, 1);
    assert!(
        out.contains(
            "<testsuite name=\"test/report.test.ras\" tests=\"5\" failures=\"1\" skipped=\"2\""
        ),
        "{out}"
    );
    assert!(
        out.contains("<failure message=\"toBe: 期待値 7、実際の値 6\">"),
        "{out}"
    );
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
    assert_eq!(
        (
            caught["status"].as_str(),
            caught["failure"]["matcher"].as_str()
        ),
        (Some("failed"), Some("toBe"))
    );
    assert_eq!(caught["printed"], serde_json::json!(["caught"]));
    // An error an event handler does not catch fails the test (spec §14.12.5).
    let event = result(
        &events,
        "ブックとファイル > イベント ハンドラのエラーはテストを失敗にする",
    );
    assert_eq!(event["status"], "failed");
    assert!(
        event["failure"]["description"]
            .as_str()
            .unwrap()
            .contains("Sheet1.Change"),
        "{event:#}"
    );
    // A failing `beforeAll` fails its tests without running them.
    let hook = result(&events, "フック > beforeAll の失敗 > 実行されない");
    assert_eq!(
        (
            hook["status"].as_str(),
            hook["failure"]["kind"].as_str(),
            hook["failure"]["hook"].as_str()
        ),
        (Some("failed"), Some("hook"), Some("beforeAll"))
    );
    assert_eq!(hook["printed"], serde_json::json!([]));
    let summary = events.iter().find(|e| e["type"] == "summary").unwrap();
    assert_eq!(
        (summary["passed"].as_i64(), summary["failed"].as_i64()),
        (Some(9), Some(3)),
        "{out}"
    );

    // `event-errors = "ignore"`: the error is told, the test passes.
    std::fs::write(dir.join("ras.toml"), "[project]\nname = \"Sales\"\nworkbook = \"Book.xlsx\"\n\n[test]\nevent-errors = \"ignore\"\n").unwrap();
    let (_, out, _) = ras(
        &dir,
        &[
            "test",
            "test/more.test.ras",
            "-t",
            "イベント",
            "--reporter",
            "json",
        ],
    );
    let events = results(&out);
    let event = result(
        &events,
        "ブックとファイル > イベント ハンドラのエラーはテストを失敗にする",
    );
    assert_eq!(event["status"], "passed", "{event:#}");
    assert!(
        event["warnings"][0]
            .as_str()
            .unwrap()
            .contains("H1 は書き換えられません"),
        "{event:#}"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_time_outs_only_and_what_test_files_may_hold() {
    if !has_node() {
        return;
    }
    let slow = "test(\"速い\", () => {\n    expect(1).toBe(1);\n});\ntest(\"終わらない\", () => {\n    let i = 0;\n    while (true) {\n        i = (i + 1) % 10;\n    }\n}, { timeout: 300 });\ntest(\"その後\", () => {\n    expect(1).toBe(1);\n});\n";
    let only = "test(\"a\", () => {\n    expect(1).toBe(2);\n});\ntest.only(\"b\", () => {\n    expect(1).toBe(1);\n});\ndescribe.skip(\"s\", () => {\n    test(\"c\", () => {\n        expect(1).toBe(2);\n    });\n});\n";
    let dir = project(
        "test-limits",
        &mut sales(),
        &[
            ("src/report.ras", REPORT),
            ("objects/Sheet1.ras", SHEET1),
            ("test/slow.test.ras", slow),
            ("test/only.test.ras", only),
        ],
    );

    // The file is given up at the time-out: the tests after it fail too (spec §14.12.4).
    let (code, out, err) = ras(&dir, &["test", "test/slow.test.ras", "--reporter", "json"]);
    assert_eq!(code, 1, "{out}{err}");
    let events = results(&out);
    assert_eq!(result(&events, "速い")["status"], "passed");
    let slow = result(&events, "終わらない");
    assert_eq!(
        (slow["status"].as_str(), slow["failure"]["kind"].as_str()),
        (Some("failed"), Some("timeout")),
        "{slow:#}"
    );
    assert_eq!(result(&events, "その後")["status"], "failed");

    // `only` runs alone; with --ci it is an error.
    let (code, out, err) = ras(&dir, &["test", "test/only.test.ras", "--reporter", "json"]);
    assert_eq!(code, 0, "{out}{err}");
    let events = results(&out);
    assert_eq!(
        (
            result(&events, "a")["status"].as_str(),
            result(&events, "b")["status"].as_str(),
            result(&events, "s > c")["status"].as_str()
        ),
        (Some("skipped"), Some("passed"), Some("skipped"))
    );
    let (code, _, err) = ras(&dir, &["test", "test/only.test.ras", "--ci"]);
    assert_eq!(code, 2);
    assert!(err.contains("`only` の付いたテストがあります"), "{err}");

    // --bail stops at the first failure.
    std::fs::write(
        dir.join("test/only.test.ras"),
        only.replace("test.only", "test"),
    )
    .unwrap();
    let (code, out, _) = ras(&dir, &["test", "--bail", "1", "--reporter", "json"]);
    assert_eq!(code, 1);
    let events = results(&out);
    assert_eq!(
        (
            result(&events, "a")["status"].as_str(),
            result(&events, "b")["status"].as_str()
        ),
        (Some("failed"), Some("skipped")),
        "{out}"
    );
    assert!(
        !events
            .iter()
            .any(|e| e["type"] == "file" && e["file"] == "test/slow.test.ras"),
        "{out}"
    );

    // What a test file may hold, and who may import whom.
    std::fs::write(
        dir.join("test/only.test.ras"),
        "const name = \"x\";\nSheet1.cells(1, 1).value = 1;\ndescribe(name, () => {});\ndescribe(\"d\", () => {\n    for (let i = 0; i < 2; i++) {}\n});\n@macro\nfunction m() {}\n",
    )
    .unwrap();
    std::fs::write(
        dir.join("test/helper.ras"),
        "export function twice(n: int): int {\n    return n * 2;\n}\n",
    )
    .unwrap();
    std::fs::write(
        dir.join("src/bad.ras"),
        "import { twice } from \"../test/helper\";\n",
    )
    .unwrap();
    let (code, _, err) = ras(&dir, &["test"]);
    assert_eq!(code, 2, "{err}");
    for expected in [
        "error[E0701]",
        "error[E0702]",
        "test/only.test.ras:2:1",
        "test/only.test.ras:5:5",
        "error[E0703]",
        "test/only.test.ras:3:10",
        "error[E0704]",
        "test/only.test.ras:7:1",
    ] {
        assert!(err.contains(expected), "{expected} is missing in:\n{err}");
    }
    // `ras check` reports the tests' problems too.
    let (code, _, err) = ras(&dir, &["check"]);
    assert_eq!(code, 1);
    assert!(
        err.contains("error[E0702]") && err.contains("error[E0701]"),
        "{err}"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn run_replaces_hand_written_vba_and_com_objects() {
    if !has_node() {
        return;
    }
    let main = r#"export function legacyTotal() {
    const total = application.run("LegacyTotal", 2, 3);
    debug.print("total", total);
    const conn = createObject("ADODB.Connection");
    conn.open("dsn");
    debug.print("state", conn.state);
}
"#;
    let mocks = r#"export function total(a: int, b: int): int {
    return a + b;
}

export class FakeConnection {
    state: int = 0;
    open(source: string) {
        this.state = 1;
    }
}
"#;
    let dir = project(
        "replace",
        &mut sales(),
        &[("src/main.ras", main), ("src/mocks.ras", mocks)],
    );

    // Without replacements both need Excel.
    let (code, _, err) = ras(&dir, &["run", "legacyTotal"]);
    assert_eq!(code, 1);
    assert!(err.contains("実行時エラー 445"), "{err}");

    let (code, out, err) = ras(
        &dir,
        &[
            "run",
            "legacyTotal",
            "--replace-run",
            "LegacyTotal=Mocks.total",
            "--replace-object",
            "ADODB.Connection=Mocks.FakeConnection",
        ],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert!(
        out.contains("total 5\n") && out.contains("state 1\n"),
        "{out}"
    );

    let (code, _, err) = ras(
        &dir,
        &[
            "run",
            "legacyTotal",
            "--replace-run",
            "LegacyTotal=Mocks.nothing",
        ],
    );
    assert_eq!(code, 1);
    assert!(
        err.contains("置き換え `Mocks.nothing` が見つかりません"),
        "{err}"
    );
    let (code, _, err) = ras(
        &dir,
        &["run", "legacyTotal", "--replace-run", "LegacyTotal"],
    );
    assert_eq!(code, 2);
    assert!(err.contains("`名前=モジュール名.関数名` の形"), "{err}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn diagnostics_as_json_and_explained() {
    let dir = project("json", &mut sales(), &[("src/main.ras", "export function f() {\n    const ws = Sheet1;\n    ws.Cells(1, 1).value = \"a\" - 1;\n}\n")]);
    let (code, out, err) = ras(&dir, &["check", "--message-format", "json"]);
    assert_eq!(code, 1, "{err}");
    let lines: Vec<serde_json::Value> = out
        .lines()
        .map(|l| serde_json::from_str(l).unwrap())
        .collect();
    assert_eq!(lines.len(), 2, "{out}");
    assert_eq!(
        lines[0],
        serde_json::json!({
            "type": "diagnostic", "code": "W0304", "severity": "warning", "file": "src/main.ras",
            "message": "`Cells` は RAS では `cells` と書きます（Excel / VBA のメンバーは camelCase で参照します）",
            "range": { "start": { "line": 3, "column": 8 }, "end": { "line": 3, "column": 13 } },
            "fixes": [{ "file": "src/main.ras", "newText": "cells", "range": { "start": { "line": 3, "column": 8 }, "end": { "line": 3, "column": 13 } } }],
            "explain": "ras explain W0304",
        })
    );
    assert_eq!(
        (
            lines[1]["code"].as_str(),
            lines[1]["severity"].as_str(),
            lines[1]["range"]["start"]["line"].as_i64()
        ),
        (Some("E0387"), Some("error"), Some(3))
    );
    // The other commands that compile write them the same way.
    let (code, out, _) = ras(&dir, &["build", "--message-format", "json"]);
    assert_eq!(code, 1);
    assert!(
        out.lines().any(|l| l.contains("\"code\":\"E0387\"")),
        "{out}"
    );
    // People are pointed to the explanations.
    let (_, _, err) = ras(&dir, &["check"]);
    assert!(
        err.contains("error[E0387]") && err.contains("ras explain E0387"),
        "{err}"
    );

    let (code, out, _) = ras(&dir, &["explain", "e0387"]);
    assert_eq!(code, 0);
    assert!(out.starts_with("E0387 — 数値が必要（エラー）\n"), "{out}");
    let (code, out, _) = ras(&dir, &["explain"]);
    assert_eq!(code, 0);
    assert!(out.contains("W0304  VBA の綴りの名前\n"), "{out}");
    let (code, _, err) = ras(&dir, &["explain", "E9999"]);
    assert_eq!(code, 2);
    assert!(
        err.contains("`E9999` という診断のコードはありません"),
        "{err}"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

/// `stdlib/standin.json` is what `tools/standin/members.js` lists from the stand-in itself.
#[test]
fn stand_in_members_are_up_to_date() {
    if !has_node() {
        return;
    }
    let dir = project(
        "standin-list",
        &mut sales(),
        &[("src/main.ras", "export function f() {\n}\n")],
    );
    let (code, _, err) = ras(&dir, &["build", "--target", "js"]);
    assert_eq!(code, 0, "{err}");
    let repo = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let out = Command::new("node")
        .arg(repo.join("tools/standin/members.js"))
        .arg(dir.join("build/debug"))
        .env("RAS_BOOK_ADDON", addon())
        .output()
        .expect("node");
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    let listed = String::from_utf8_lossy(&out.stdout).replace("\r\n", "\n");
    let file = repo.join("crates/ras-compiler/stdlib/standin.json");
    if std::env::var_os("RAS_UPDATE_STANDIN").is_some() {
        std::fs::write(&file, &listed).unwrap();
    }
    let committed = std::fs::read_to_string(&file)
        .unwrap()
        .replace("\r\n", "\n");
    assert!(committed == listed, "crates/ras-compiler/stdlib/standin.json is out of date: run this test with RAS_UPDATE_STANDIN=1");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn check_warns_about_what_the_stand_in_lacks() {
    let source = r#"export function f() {
    const ws = Sheet1;
    ws.range("A1").value = application.max(1, 2) + worksheets.count;
    ws.range("A1").font.bold = true;
    ws.range("A1").addComment("x");
    ws.shapes.count;
    const file = application.getOpenFilename();
    application.run("Legacy");
    const conn = createObject("ADODB.Connection");
    const dict = new Scripting.Dictionary();
    debug.print(format(now(), "yyyy"), file, dict.count);
}
"#;
    let dir = project("standin", &mut sales(), &[("src/main.ras", source)]);
    // Not by default: the program is VBA for Excel.
    let (code, _, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{err}");
    assert!(!err.contains("W051"), "{err}");

    let (code, out, _) = ras(
        &dir,
        &["check", "--target", "stand-in", "--message-format", "json"],
    );
    assert_eq!(code, 0, "warnings only: {out}");
    let found: Vec<(String, i64)> = out
        .lines()
        .filter_map(|l| serde_json::from_str::<serde_json::Value>(l).ok())
        .filter(|d| d["code"].as_str().is_some_and(|c| c.starts_with("W051")))
        .map(|d| {
            (
                d["code"].as_str().unwrap().to_string(),
                d["range"]["start"]["line"].as_i64().unwrap(),
            )
        })
        .collect();
    assert_eq!(
        found,
        vec![
            ("W0510".to_string(), 5),
            ("W0510".to_string(), 6),
            ("W0511".to_string(), 7),
            ("W0511".to_string(), 8),
            ("W0512".to_string(), 9)
        ],
        "{out}"
    );

    // Or always, from ras.toml.
    std::fs::write(
        dir.join("ras.toml"),
        "[project]\nname = \"Sales\"\nworkbook = \"Book.xlsx\"\n\n[lint]\nstand-in = true\n",
    )
    .unwrap();
    let (_, _, err) = ras(&dir, &["check"]);
    assert!(
        err.contains("warning[W0510]: `addComment` はブックの代役の `Range` にありません"),
        "{err}"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

/// `window.DATA` of a page of `ras book view`.
fn view_data(page: &str) -> serde_json::Value {
    let start = page.find("window.DATA = ").expect("the data of the page") + "window.DATA = ".len();
    let end = page[start..].find(";</script>").unwrap() + start;
    serde_json::from_str(&page[start..end].replace("<\\/", "</")).unwrap()
}

#[test]
fn book_view_draws_the_sheets_and_marks_changes() {
    let dir = temp_dir("view");
    let mut a = sales();
    a.set_format(0, at("A1:C1"), "font.bold", &Val::Bool(true))
        .unwrap();
    a.set_format(
        0,
        at("A1:C1"),
        "interior.color",
        &Val::Num(0x41_7C_10 as f64),
    )
    .unwrap();
    a.set_number_format(0, at("C2:C5"), "#,##0").unwrap();
    a.merge_cells(0, at("E1:F2"), false).unwrap();
    a.save(&dir.join("a.xlsx")).unwrap();
    a.set_values(0, at("B3"), &[vec![Val::Num(70.0)]]).unwrap();
    a.save(&dir.join("b.xlsx")).unwrap();

    let (code, out, err) = ras(
        &dir,
        &["book", "view", "b.xlsx", "--diff", "a.xlsx", "--stdout"],
    );
    assert_eq!(code, 0, "{err}");
    assert!(
        out.starts_with("<!doctype html>") && out.contains("<div id=\"grid\">"),
        "{}",
        &out[..200.min(out.len())]
    );
    let data = view_data(&out);
    assert_eq!(data["file"], "b.xlsx");
    let sheet = &data["sheets"][0];
    assert_eq!(
        (
            sheet["name"].as_str(),
            sheet["rows"].as_i64(),
            sheet["cols"].as_i64()
        ),
        (Some("Sheet1"), Some(5), Some(6)),
        "{sheet:#}"
    );
    let cell = |a1: (i64, i64)| {
        sheet["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|c| c["r"] == a1.0 && c["c"] == a1.1)
            .cloned()
            .unwrap()
    };
    let style = |c: &serde_json::Value| {
        sheet["styles"][c["s"].as_u64().unwrap() as usize]
            .as_str()
            .unwrap()
            .to_string()
    };
    // A1: bold on a green fill.
    let a1 = cell((0, 0));
    assert_eq!(a1["t"], "商品");
    assert!(
        style(&a1).contains("font-weight:bold") && style(&a1).contains("background:#107C41"),
        "{}",
        style(&a1)
    );
    // C3: what the cell shows, its formula, and numbers to the right.
    let c3 = cell((2, 2));
    assert_eq!(
        (c3["t"].as_str(), c3["f"].as_str(), c3["k"].as_str()),
        (Some("7,000"), Some("=B3*100"), Some("number"))
    );
    assert!(style(&c3).contains("text-align:right"), "{}", style(&c3));
    assert_eq!(sheet["merges"], serde_json::json!([[0, 4, 1, 5]]));
    // The grid goes on past the cells; no drawing, no objects.
    assert_eq!(
        (
            sheet["maxRows"].as_u64(),
            sheet["maxCols"].as_u64(),
            sheet["gridLines"].as_bool()
        ),
        (Some(1_048_576), Some(16_384), Some(true))
    );
    assert_eq!(sheet["objects"], serde_json::json!([]));
    // What changed since a.xlsx, for the marks and the list.
    assert_eq!(data["diff"]["before"], "a.xlsx");
    let changed: Vec<&str> = data["diff"]["changes"]["Sheet1"]
        .as_array()
        .unwrap()
        .iter()
        .map(|c| c["a"].as_str().unwrap())
        .collect();
    assert_eq!(changed, vec!["B3", "C3", "C5"]);
    assert_eq!(
        data["diff"]["changes"]["Sheet1"][0],
        serde_json::json!({ "a": "B3", "before": "5", "after": "70" })
    );

    // The drawings of another workbook (the one a run read).
    let (code, _, err) = ras(
        &dir,
        &["book", "view", "b.xlsx", "--drawings", "a.xlsx", "--stdout"],
    );
    assert_eq!(code, 0, "{err}");

    // Into a file; the limit on rows is told.
    let (code, out, err) = ras(&dir, &["book", "view", "b.xlsx", "-o", "out/b.html"]);
    assert_eq!(code, 0, "{err}");
    assert!(out.contains("b.html を作成しました"), "{out}");
    let page = std::fs::read_to_string(dir.join("out/b.html")).unwrap();
    assert_eq!(view_data(&page)["diff"], serde_json::Value::Null);
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn run_view_and_the_snapshot_of_a_stopped_run() {
    if !has_node() {
        return;
    }
    let dir = project(
        "run-view",
        &mut sales(),
        &[("src/report.ras", REPORT), ("objects/Sheet1.ras", SHEET1)],
    );
    let (code, out, err) = ras(&dir, &["run", "rebuildReport", "--view"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(
        out.contains("グリッド表示: build/view/Book.html")
            || out.contains("グリッド表示: build\\view\\Book.html"),
        "{out}"
    );
    let data = view_data(&std::fs::read_to_string(dir.join("build/view/Book.html")).unwrap());
    assert_eq!(
        data["diff"]["changes"]["Sheet1"].as_array().unwrap().len(),
        3,
        "{:#}",
        data["diff"]
    );

    // The editor asks the stopped program for its workbooks (`$rasBookSnapshot`), then draws them.
    let script = dir.join("snapshot.js");
    std::fs::write(
        &script,
        r#"const path = require("path");
const dir = path.join(__dirname, "build", "debug");
const $rt = require(path.join(dir, "ras-runtime.js"));
$rt.configure({ manifest: { workbook: "Book.xlsx" }, root: __dirname, book: process.env.RAS_BOOK_ADDON, saveOnExit: false });
$rt.hooks.note = () => {};
$rt.hooks.print = () => {};
require(path.join(dir, "report.js"));
const out = path.join(__dirname, "snap");
console.log($rasBookSnapshot(out));
$rt.run("Report.rebuildReport");
console.log($rasBookSnapshot(out));
"#,
    )
    .unwrap();
    let run = Command::new("node")
        .arg(&script)
        .env("RAS_BOOK_ADDON", addon())
        .output()
        .expect("node");
    assert!(
        run.status.success(),
        "{}",
        String::from_utf8_lossy(&run.stderr)
    );
    // Nothing before the program used a workbook; then the project's workbook.
    assert_eq!(
        String::from_utf8_lossy(&run.stdout).replace("\r\n", "\n"),
        "0\n1\n"
    );
    let list: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(dir.join("snap/snapshot.json")).unwrap())
            .unwrap();
    assert_eq!(
        (
            list["books"][0]["name"].as_str(),
            list["books"][0]["file"].as_str(),
            list["books"][0]["project"].as_bool()
        ),
        (Some("Book.xlsx"), Some("book-0.xlsx"), Some(true))
    );
    let (code, page, err) = ras(
        &dir,
        &[
            "book",
            "view",
            "snap/book-0.xlsx",
            "--diff",
            "Book.xlsx",
            "--title",
            "Book.xlsx",
            "--stdout",
        ],
    );
    assert_eq!(code, 0, "{err}");
    let data = view_data(&page);
    assert_eq!(
        (data["file"].as_str(), data["diff"]["before"].as_str()),
        (Some("Book.xlsx"), Some("前回"))
    );
    assert_eq!(
        data["diff"]["changes"]["Sheet1"][0],
        serde_json::json!({ "a": "D2", "before": "（空）", "after": "6" })
    );
    let _ = std::fs::remove_dir_all(&dir);
}

/// The run mode of the editor's sheet panel (`ras-main.js --serve`): what the user does on the
/// page raises the program's events on the stand-in, and each request is answered on a line of
/// its own, after what the handlers printed.
#[test]
fn run_mode_of_the_sheet_panel_raises_the_events_of_what_the_user_does() {
    if !has_node() {
        return;
    }
    let dir = project("serve", &mut sales(), &[]);
    let ops = r#"[{"op":"activeX","sheet":"Sheet1","control":"commandButton","at":"F2","name":"CommandButton1","text":"押す"},{"op":"control","sheet":"Sheet1","control":"button","at":"F6","text":"ボタン"}]"#;
    let (code, _, err) = ras(&dir, &["book", "edit", "Book.xlsx", "--op", ops]);
    assert_eq!(code, 0, "{err}");
    // The form control's name, as the workbook has it ("ボタン 2").
    let (_, page, _) = ras(&dir, &["book", "view", "Book.xlsx", "--stdout"]);
    let data = view_data(&page);
    let button = data["sheets"][0]["objects"]
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["activex"].is_null())
        .and_then(|o| o["name"].as_str())
        .unwrap()
        .to_string();
    std::fs::create_dir_all(dir.join("objects")).unwrap();
    std::fs::write(
        dir.join("objects/Sheet1.ras"),
        format!(
            r#"@on("Worksheet.Change")
function changed(target: Excel.Range) {{
    debug.print("Change", target.address(false, false));
    if (target.column == 1) this.cells(target.row, 8).value = "seen";
}}

@on("Worksheet.SelectionChange")
function selected(target: Excel.Range) {{
    debug.print("SelectionChange", target.address(false, false));
}}

@on("Worksheet.Activate")
function entered() {{
    debug.print("Activate");
}}

@on("Worksheet.BeforeDoubleClick")
function doubleClicked(target: Excel.Range, byref cancel: boolean) {{
    debug.print("BeforeDoubleClick", target.address(false, false));
    cancel = true;
}}

@on("CommandButton1.Click")
function pushed() {{
    this.range("E1").value = "clicked";
}}

@on("CommandButton1.MouseDown")
function pressed(button: int, shift: int, x: float, y: float) {{
    debug.print("MouseDown", button, shift, x, y);
}}

@on("{button}.Click")
function buttonClicked() {{
    this.range("E2").value = "button";
}}
"#
        ),
    )
    .unwrap();
    let (code, out, err) = ras(&dir, &["build", "--target", "js"]);
    assert_eq!(code, 0, "{out}{err}");
    let requests = [
        r#"{"id":1,"op":"start"}"#.to_string(),
        r#"{"id":2,"op":"select","sheet":"Sheet1","range":"A6","active":"A6"}"#.to_string(),
        r#"{"id":3,"op":"enter","sheet":"Sheet1","ranges":["A6"],"text":"E社"}"#.to_string(),
        r#"{"id":4,"op":"select","sheet":"Sheet1","range":"A6","active":"A6"}"#.to_string(),
        r#"{"id":5,"op":"doubleClick","sheet":"Sheet1","range":"C3"}"#.to_string(),
        r#"{"id":6,"op":"control","sheet":"Sheet1","object":"CommandButton1","events":[{"event":"MouseDown","args":[1,0,3,4]},{"event":"Click"}]}"#.to_string(),
        format!(r#"{{"id":7,"op":"macro","sheet":"Sheet1","object":"{button}","macro":""}}"#),
        r#"{"id":8,"op":"macro","sheet":"Sheet1","object":"","macro":"Sheet1.nothing"}"#.to_string(),
    ];
    let mut child = Command::new("node")
        .arg(dir.join("build/debug/ras-main.js"))
        .arg("--serve")
        .arg(dir.join("snap"))
        .current_dir(&dir)
        .env("RAS_BOOK_ADDON", addon())
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .expect("node");
    {
        use std::io::Write;
        let mut stdin = child.stdin.take().unwrap();
        stdin
            .write_all(format!("{}\n", requests.join("\n")).as_bytes())
            .unwrap();
    }
    let run = child.wait_with_output().unwrap();
    assert!(
        run.status.success(),
        "{}",
        String::from_utf8_lossy(&run.stderr)
    );
    let stdout = String::from_utf8_lossy(&run.stdout).replace("\r\n", "\n");
    let mut printed = Vec::new();
    let mut replies = Vec::new();
    for line in stdout.lines() {
        match line.strip_prefix("\u{1e}RAS ") {
            Some(json) => replies.push(serde_json::from_str::<serde_json::Value>(json).unwrap()),
            // (Not what the run time notes: `[ras] Excel を使わずに実行します…`.)
            None if !line.starts_with("[ras]") => printed.push(line.to_string()),
            None => {}
        }
    }
    // In Excel's order; the same selection again raises nothing.
    assert_eq!(
        printed,
        [
            "SelectionChange A6",
            "Change A6",
            "BeforeDoubleClick C3",
            "MouseDown 1 0 3 4",
            // (What the handlers of the buttons wrote.)
            "Change E1",
            "Change E2",
        ],
        "{stdout}"
    );
    assert_eq!(replies.len(), 8, "{stdout}");
    assert!(replies[..7].iter().all(|r| r["ok"] == true), "{stdout}");
    assert_eq!(replies[4]["cancel"], true);
    assert_eq!(replies[5]["handled"], true);
    assert_eq!(replies[1]["where"]["cell"], "A6");
    // Each change of the workbook is a new revision.
    let rev = |i: usize| replies[i]["rev"].as_i64().unwrap();
    assert!(
        rev(2) > rev(1) && rev(5) > rev(4) && rev(6) > rev(5),
        "{stdout}"
    );
    // A macro that is nowhere: the request fails, the program goes on.
    assert_eq!(replies[7]["ok"], false);
    // The workbook as the program left it, for the panel to draw.
    let (code, dump, err) = ras(
        &dir,
        &["book", "dump", "snap/book-0.xlsx", "--sheet", "Sheet1"],
    );
    assert_eq!(code, 0, "{err}");
    for text in ["E社", "seen", "clicked", "button"] {
        assert!(dump.contains(text), "{text}: {dump}");
    }
    let _ = std::fs::remove_dir_all(&dir);
}

/// The examples of the cheat sheet are checked as they are written (docs/cheatsheet.md).
#[test]
fn cheat_sheet_examples_pass_the_check() {
    // (A checkout on Windows may have CRLF line ends.)
    let sheet = std::fs::read_to_string(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../docs/cheatsheet.md"),
    )
    .unwrap()
    .replace("\r\n", "\n");
    let mut files: Vec<(String, String)> = Vec::new();
    for (i, block) in sheet.split("```typescript\n").skip(1).enumerate() {
        let code = block.split("```").next().unwrap().to_string();
        // `// objects/Sheet1.ras…`: the file of a sheet or of the workbook.
        let object = code
            .strip_prefix("// objects/")
            .and_then(|rest| rest.split_once(".ras"))
            .map(|(name, _)| format!("objects/{name}.ras"));
        // `// forms/frmInput.ras…`: a form's code-behind, with a design of the controls it uses.
        let form = code
            .strip_prefix("// forms/")
            .and_then(|rest| rest.split_once(".ras"))
            .map(|(name, _)| name.to_string());
        if let Some(name) = &form {
            let spec = serde_json::json!({
                "schemaVersion": 1, "kind": "ras.userform", "basis": "designer",
                "form": { "name": name, "caption": name, "width": 240, "height": 140 },
                "controls": [
                    { "id": "n", "name": "txtName", "type": "TextBox", "left": 12, "top": 12 },
                    { "id": "o", "name": "btnOK", "type": "CommandButton", "left": 12, "top": 48 },
                    { "id": "c", "name": "btnCancel", "type": "CommandButton", "left": 96, "top": 48 },
                    { "id": "p", "name": "pbWork", "type": "ProgressBar", "left": 12, "top": 80, "width": 216 },
                    { "id": "f", "name": "fpBook", "type": "FilePicker", "left": 12, "top": 110, "width": 216 }
                ]
            });
            files.push((format!("forms/{name}.form.json"), spec.to_string()));
        }
        let path = if code.starts_with("// test/") {
            format!("test/block{i}.test.ras")
        } else if let Some(name) = form {
            format!("forms/{name}.ras")
        } else {
            object.unwrap_or_else(|| format!("src/block{i}.ras"))
        };
        files.push((path, code));
    }
    assert!(files.len() >= 6, "the cheat sheet has code examples");
    let refs: Vec<(&str, &str)> = files
        .iter()
        .map(|(p, c)| (p.as_str(), c.as_str()))
        .collect();
    let dir = project("cheatsheet", &mut sales(), &refs);
    let (code, out, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(!err.contains("warning["), "{err}");
    // `ras cheatsheet` prints it, and `ras init` leaves the guide for AI agents.
    let (code, out, _) = ras(&dir, &["cheatsheet"]);
    assert_eq!(code, 0);
    assert!(
        out.starts_with("# RAS 早見表"),
        "{}",
        &out[..40.min(out.len())]
    );
    let fresh = temp_dir("init-agents");
    let (code, _, err) = ras(&fresh, &["init", "."]);
    assert_eq!(code, 0, "{err}");
    assert!(std::fs::read_to_string(fresh.join("AGENTS.md"))
        .unwrap()
        .contains("ras book dump"));
    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_dir_all(&fresh);
}

/// `ras sheet`: the sheets of the workbook file and their files in `objects/` change together
/// (spec §14.16). The workbook is read back as Excel's stand-in reads it.
#[test]
fn sheets_are_added_copied_renamed_and_removed() {
    use ras_workbook::book::Read;
    let mut book = sales();
    book.add_sheet(1, Some("Data")).unwrap();
    book.set_values(
        1,
        at("A1:A3"),
        &[
            vec![Val::Num(1.0)],
            vec![Val::Num(2.0)],
            vec![Val::Num(3.0)],
        ],
    )
    .unwrap();
    book.set_formulas(0, at("E1"), &[vec![text("=SUM(Data!A1:A3)")]], false)
        .unwrap();
    book.add_name("データ", None, "Data!$A$1:$A$3").unwrap();
    let main = "export function total(): double {\n    return Sheet2.range(\"A1\").value as double + Sheet1.range(\"E1\").value as double;\n}\n";
    let dir = project(
        "sheets",
        &mut book,
        &[
            ("src/main.ras", main),
            (
                "objects/Sheet2.ras",
                "export function first(): any {\n    return this.range(\"A1\").value;\n}\n",
            ),
        ],
    );
    let book_file = dir.join("Book.xlsx");
    let names = || {
        let b = Book::open(&book_file).unwrap();
        (0..b.sheet_count())
            .map(|i| {
                format!(
                    "{}={}",
                    b.sheet_name(i).unwrap_or_default(),
                    b.code_name(i).unwrap_or_default()
                )
            })
            .collect::<Vec<_>>()
    };
    let formula = |cell: &str| match &Book::open(&book_file)
        .unwrap()
        .get(0, at(cell), Read::Formula)
        .unwrap()[0][0]
    {
        Val::Text(t) => t.clone(),
        other => format!("{other:?}"),
    };
    assert_eq!(names(), ["Sheet1=Sheet1", "Data=Sheet2"]);

    let (code, out, err) = ras(&dir, &["sheet", "list"]);
    assert_eq!(code, 0, "{err}");
    assert!(
        out.contains("Sheet1（コード名 Sheet1）")
            && out.contains("Data（コード名 Sheet2）  objects/Sheet2.ras")
            && out.contains("ファイルなし"),
        "{out}"
    );

    // Add: the sheet and its file.
    let (code, out, err) = ras(&dir, &["sheet", "add", "集計", "--after", "Sheet1"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(
        out.contains("シート「集計」（コード名 Sheet3）を追加しました")
            && out.contains("作成: objects/Sheet3.ras"),
        "{out}"
    );
    assert_eq!(names(), ["Sheet1=Sheet1", "集計=Sheet3", "Data=Sheet2"]);
    assert!(std::fs::read_to_string(dir.join("objects/Sheet3.ras"))
        .unwrap()
        .contains("シート「集計」（コード名 Sheet3）"));
    assert_eq!(formula("E1"), "=SUM(Data!A1:A3)");
    let (code, _, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{err}");

    // Rename: the tab name in the formulas, the code name in the sources and the file name.
    let (code, out, err) = ras(
        &dir,
        &[
            "sheet",
            "rename",
            "Data",
            "--name",
            "売上 2026",
            "--code-name",
            "Sales",
            "--json",
        ],
    );
    assert_eq!(code, 0, "{out}{err}");
    let result: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(result["codeName"], "Sales");
    assert_eq!(result["before"]["codeName"], "Sheet2");
    assert_eq!(result["renamed"][0]["to"], "objects/Sales.ras");
    assert_eq!(result["references"], 1);
    assert!(result["formulas"].as_u64().unwrap() >= 2, "{result}");
    assert_eq!(names(), ["Sheet1=Sheet1", "集計=Sheet3", "売上 2026=Sales"]);
    assert_eq!(formula("E1"), "=SUM('売上 2026'!A1:A3)");
    assert!(
        Book::open(&book_file)
            .unwrap()
            .names()
            .iter()
            .any(|(name, _, refers)| name == "データ" && refers.contains("売上 2026")),
        "{:?}",
        Book::open(&book_file).unwrap().names()
    );
    assert!(std::fs::read_to_string(dir.join("src/main.ras"))
        .unwrap()
        .contains("Sales.range(\"A1\")"));
    assert!(dir.join("objects/Sales.ras").is_file() && !dir.join("objects/Sheet2.ras").exists());
    let (code, _, err) = ras(&dir, &["check"]);
    assert_eq!(code, 0, "{err}");
    // The value the formula computes is still there for the stand-in.
    let (code, out, err) = ras(&dir, &["run", "Main.total"]);
    assert_eq!(code, 0, "{out}{err}");

    // Copy: the cells and the file.
    let (code, out, err) = ras(&dir, &["sheet", "copy", "Sales"]);
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(
        names(),
        [
            "Sheet1=Sheet1",
            "集計=Sheet3",
            "売上 2026=Sales",
            "売上 2026 (2)=Sheet2"
        ]
    );
    assert_eq!(
        std::fs::read_to_string(dir.join("objects/Sheet2.ras")).unwrap(),
        std::fs::read_to_string(dir.join("objects/Sales.ras")).unwrap()
    );
    assert_eq!(
        Book::open(&book_file)
            .unwrap()
            .get(3, at("A1:A3"), Read::Value)
            .unwrap(),
        vec![
            vec![Val::Num(1.0)],
            vec![Val::Num(2.0)],
            vec![Val::Num(3.0)]
        ]
    );

    // Remove: refused while something refers to the sheet.
    let (code, out, err) = ras(&dir, &["sheet", "remove", "Sales"]);
    assert_eq!(code, 1, "{out}{err}");
    assert!(
        err.contains("objects/Sales.ras にコードがあります")
            && err.contains("src/main.ras:2")
            && err.contains("Sheet1!E1")
            && err.contains("名前の定義 データ")
            && err.contains("--force"),
        "{err}"
    );
    assert_eq!(names().len(), 4);
    assert!(dir.join("objects/Sales.ras").is_file());
    let (code, out, err) = ras(&dir, &["sheet", "remove", "Sales", "--force", "--json"]);
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(
        names(),
        ["Sheet1=Sheet1", "集計=Sheet3", "売上 2026 (2)=Sheet2"]
    );
    assert!(!dir.join("objects/Sales.ras").exists());
    // The formula and the defined name that referred to the sheet were rewritten to `#REF!`
    // (as Excel does; the stand-in cannot read such a formula back).
    let result: serde_json::Value = serde_json::from_str(&out).unwrap();
    assert_eq!(
        (result["removed"][0].as_str(), result["formulas"].as_u64()),
        (Some("objects/Sales.ras"), Some(2)),
        "{result}"
    );
    assert!(!formula("E1").contains("売上"), "{}", formula("E1"));
    // A sheet without code that nothing refers to goes without `--force`.
    let (code, out, err) = ras(&dir, &["sheet", "remove", "集計"]);
    assert_eq!(code, 0, "{out}{err}");
    assert!(!dir.join("objects/Sheet3.ras").exists());

    // What cannot be done says why, and changes nothing.
    for (args, why) in [
        (&["sheet", "add", "Sheet1"][..], "既にあります"),
        (&["sheet", "add", "a/b"][..], "見出しの名前にできません"),
        (
            &["sheet", "add", "New", "--code-name", "main"][..],
            "ほかのシートかモジュールが使っています",
        ),
        (
            &["sheet", "rename", "Nope", "--name", "x"][..],
            "というシートがありません",
        ),
        (
            &["sheet", "rename", "Sheet1"][..],
            "新しい名前を指定してください",
        ),
    ] {
        let (code, out, err) = ras(&dir, args);
        assert!(code != 0 && err.contains(why), "{args:?}: {out}{err}");
    }
    assert_eq!(names(), ["Sheet1=Sheet1", "売上 2026 (2)=Sheet2"]);
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn book_edit_writes_the_cells_into_the_file() {
    use ras_workbook::book::Read;
    let mut book = sales();
    let dir = project("edit", &mut book, &[("src/main.ras", REPORT)]);
    let file = dir.join("Book.xlsx");
    let cells = |range: &str, what: Read| -> Vec<Vec<Val>> {
        Book::open(&file).unwrap().get(0, at(range), what).unwrap()
    };

    // The project's workbook: cells typed into, the formulas that read them calculated.
    let (code, out, err) = ras(
        &dir,
        &["book", "edit", "--set", "B2=10", "--set", "D2==C2/2"],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert!(out.contains("Book.xlsx を書き換えました"), "{out}");
    assert_eq!(
        cells("B2:D2", Read::Value),
        [[Val::Num(10.0), Val::Num(1000.0), Val::Num(500.0)]]
    );
    assert_eq!(cells("C5", Read::Value), [[Val::Num(1700.0)]]);
    // The file as it was is kept in build/backup/.
    let kept = std::fs::read_dir(dir.join("build/backup")).unwrap().count();
    assert_eq!(kept, 1);

    // Rows inserted and deleted, cells cleared; edits as JSON, from the standard input too.
    let (code, out, err) = ras(
        &dir,
        &[
            "book",
            "edit",
            "Book.xlsx",
            "--insert",
            "3:3",
            "--clear",
            "D2",
        ],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(
        cells("A2:C4", Read::Formula),
        [
            [text("りんご"), text("10"), text("=B2*100")],
            [text(""), text(""), text("")],
            [text("みかん"), text("5"), text("=B4*100")],
        ]
    );
    assert_eq!(cells("C6", Read::Formula), [[text("=SUM(C2:C5)")]]);
    let (code, out, err) = ras(
        &dir,
        &[
            "book",
            "edit",
            "--sheet",
            "Sheet1",
            "--op",
            r#"[{"op":"delete","range":"3:3"},{"op":"copy","from":"A2:C2","to":"A8"}]"#,
        ],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(
        cells("A8:C8", Read::Formula),
        [[text("りんご"), text("10"), text("=B8*100")]]
    );
    assert_eq!(
        Book::open(&file).unwrap().names()[0].2,
        "Sheet1!$C$5",
        "the name is where it was again"
    );

    // --view: the page of the workbook as it is now, for the editor's panel.
    let (code, out, err) = ras(
        &dir,
        &["book", "edit", "--set", "A9=x", "--view", "--editable"],
    );
    assert_eq!(code, 0, "{err}");
    let data = view_data(&out);
    assert_eq!(data["editable"], true);
    assert!(
        data["sheets"][0]["cells"]
            .as_array()
            .unwrap()
            .iter()
            .any(|c| c["r"] == 8 && c["c"] == 0 && c["t"] == "x"),
        "{}",
        data["sheets"][0]["cells"]
    );
    // (`ras book view` makes a page that is only read, unless asked.)
    let (_, out, _) = ras(&dir, &["book", "view", "--stdout"]);
    assert!(view_data(&out).get("editable").is_none());

    // Formats, sizes and the filter, and what the page of the grid is told of them.
    let (code, out, err) = ras(
        &dir,
        &[
            "book",
            "edit",
            "--view",
            "--op",
            r##"[{"op":"format","range":"A1:C1","set":{"bold":true,"fill":"#DDEEFF"}},{"op":"size","range":"A:A","px":160},{"op":"filter","range":"A1"},{"op":"filterBy","column":"A","values":["りんご"]}]"##,
        ],
    );
    assert_eq!(code, 0, "{out}{err}");
    let sheet = &view_data(&out)["sheets"][0];
    assert_eq!(sheet["colWidths"][0], 160.0);
    assert_eq!(sheet["filter"], "A1:C5");
    assert_eq!(
        sheet["filterColumns"]["0"],
        serde_json::json!({ "values": ["りんご"], "blanks": false })
    );
    // (みかん and ぶどう are hidden.)
    assert_eq!(sheet["rowHeights"][2], 0.0);
    let mut edited = Book::open(&file).unwrap();
    assert_eq!(
        edited.format_of(0, at("A1:C1"), "font.bold").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(edited.hidden(0, false, 2, 3), Some(true));
    let (code, out, err) = ras(
        &dir,
        &["book", "edit", "--op", r#"{"op":"filter","off":true}"#],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(
        Book::open(&file).unwrap().hidden(0, false, 2, 3),
        Some(false)
    );

    // -o: another file; the workbook stays.
    let before = std::fs::read(&file).unwrap();
    let (code, out, err) = ras(
        &dir,
        &["book", "edit", "--set", "A1=見出し", "-o", "build/out.xlsx"],
    );
    assert_eq!(code, 0, "{out}{err}");
    assert_eq!(std::fs::read(&file).unwrap(), before);
    assert_eq!(
        Book::open(&dir.join("build/out.xlsx"))
            .unwrap()
            .get(0, at("A1"), Read::Value)
            .unwrap(),
        [[text("見出し")]]
    );

    // What cannot be done says why (exit code 1), and the file is not written.
    for (args, why) in [
        (&["book", "edit"][..], "編集が指定されていません"),
        (&["book", "edit", "--set", "B2"][..], "`=` で区切って"),
        (
            &["book", "edit", "--sheet", "Nope", "--set", "A1=x"][..],
            "シート `Nope` がありません",
        ),
        (
            &["book", "edit", "--op", "{\"op\":\"x\"}"][..],
            "編集 `x` はありません",
        ),
    ] {
        let (code, out, err) = ras(&dir, args);
        assert!(code != 0 && err.contains(why), "{args:?}: {out}{err}");
    }
    // A workbook open in Excel (its owner file is there) is not written to.
    std::fs::write(dir.join("~$Book.xlsx"), "").unwrap();
    let (code, _, err) = ras(&dir, &["book", "edit", "--set", "A1=x"]);
    assert!(code == 1 && err.contains("Excel で開かれている"), "{err}");
    assert_eq!(std::fs::read(&file).unwrap(), before);
    let _ = std::fs::remove_dir_all(&dir);
}

/// `application.fileDialog(…)` of the stand-in for Excel is answered as `getOpenFilename` is
/// (spec §14.11.14): by a test's `mockGetOpenFilename`, with several paths or a cancel.
#[test]
fn the_stand_in_answers_file_dialogs_like_get_open_filename() {
    if !has_node() {
        return;
    }
    const SRC: &str = r#"export function pick(): string[] {
    const fd = application.fileDialog(Office.MsoFileDialogType.msoFileDialogFilePicker);
    fd.allowMultiSelect = true;
    fd.filters.clear();
    fd.filters.add("Excel", "*.xlsx");
    const out: string[] = [];
    if (fd.show() != -1) return out;
    for (let i = 1; i <= fd.selectedItems.count; i++) {
        out.push(fd.selectedItems.item(i));
    }
    return out;
}
"#;
    const TEST: &str = r#"import { pick } from "../src/pick";

describe("FileDialog", () => {
    test("選んだファイル", () => {
        const m = mockGetOpenFilename(() => ["C:\\a.xlsx", "C:\\b.xlsx"]);
        expect(pick()).toEqual(["C:\\a.xlsx", "C:\\b.xlsx"]);
        expect(m.calls.length).toBe(1);
    });
    test("キャンセル", () => {
        mockGetOpenFilename(() => false);
        expect(pick()).toHaveLength(0);
    });
});
"#;
    let dir = project(
        "filedialog",
        &mut sales(),
        &[("src/pick.ras", SRC), ("test/pick.test.ras", TEST)],
    );
    let (code, out, err) = ras(&dir, &["test"]);
    assert_eq!(code, 0, "{out}{err}");
    let _ = std::fs::remove_dir_all(&dir);
}
