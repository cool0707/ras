//! End-to-end test of the `ras` binary: init → check → build → pack → unpack.

use std::path::Path;
use std::process::Command;

fn ras(dir: &Path, args: &[&str]) -> (bool, String) {
    let out = Command::new(env!("CARGO_BIN_EXE_ras")).args(args).current_dir(dir).output().expect("run ras");
    let text = format!("{}{}", String::from_utf8_lossy(&out.stdout), String::from_utf8_lossy(&out.stderr));
    (out.status.success(), text)
}

fn temp_dir(name: &str) -> std::path::PathBuf {
    let dir = std::env::temp_dir().join(format!("ras-cli-test-{name}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

#[test]
fn init_build_pack_unpack_round_trip() {
    let dir = temp_dir("roundtrip");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    assert!(dir.join("ras.toml").is_file() && dir.join("src/main.ras").is_file() && dir.join("Book.xlsm").is_file());

    std::fs::write(
        dir.join("src/report.ras"),
        "/** 日本語の説明。 */\n@fast\nexport function fill() {\n    const ws = Sheet1;\n    for (let r = 1; r <= 10; r++) {\n        ws.Cells(r, 1).Value = r * 2;\n    }\n}\n\n@on(\"Sheet1.Change\")\nfunction changed(target: Excel.Range) {\n    status(\"変更: \" + target.Address);\n}\n",
    )
    .unwrap();

    let (ok, out) = ras(&dir, &["check"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["build"]);
    assert!(ok, "{out}");
    // Module files are written in the target code page (CP932), not UTF-8.
    let bytes = std::fs::read(dir.join("build/Report.bas")).unwrap();
    let (text, _, bad) = encoding_rs::SHIFT_JIS.decode(&bytes);
    assert!(!bad);
    assert!(text.contains("Attribute VB_Name = \"Report\"\r\n"), "{text}");
    assert!(text.contains("' 日本語の説明。"), "{text}");
    assert!(text.contains("ws.Range(ws.Cells(1, 1), ws.Cells(10, 1)).Value = ras_buf1"), "{text}");
    assert!(std::str::from_utf8(&bytes).is_err(), "file must not be UTF-8");
    assert!(dir.join("build/Report.rasmap").is_file());
    assert!(dir.join("build/Sheet1.doccls").is_file());

    // The sheet code names come from the workbook: an unknown sheet is rejected.
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(out.contains("作成"), "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(out.contains("作成 0 / 更新 0 / 削除 0"), "second pack must be a no-op: {out}");
    // A relative `--project`: the sheet code names are still read from the workbook.
    let (ok, out) = ras(dir.parent().unwrap(), &["--project", dir.file_name().unwrap().to_str().unwrap(), "pack"]);
    assert!(ok && out.contains("作成 0 / 更新 0 / 削除 0"), "{out}");

    // Read the workbook back with the offline reader.
    let book = ras_ovba::unpack(&dir.join("Book.xlsm")).unwrap();
    let project = book.project.expect("VBA project");
    assert_eq!(project.codepage, 932);
    let names: Vec<&str> = project.modules.iter().map(|m| m.name.as_str()).collect();
    for expected in ["ThisWorkbook", "Sheet1", "Main", "Report", "RAS_Runtime", "RAS_Glue"] {
        assert!(names.contains(&expected), "{expected} missing from {names:?}");
    }
    let report = project.modules.iter().find(|m| m.name == "Report").unwrap();
    assert!(report.source.contains("' 日本語の説明。"));
    let sheet = project.modules.iter().find(|m| m.name == "Sheet1").unwrap();
    assert!(sheet.source.contains("'--- RAS BEGIN ---") && sheet.source.contains("Private Sub Worksheet_Change(ByVal Target As Range)"), "{}", sheet.source);

    // Removing a source file removes its module; unpack reports only RAS-managed modules.
    std::fs::remove_file(dir.join("src/main.ras")).unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok && out.contains("削除 1"), "{out}");
    let (ok, out) = ras(&dir, &["unpack"]);
    assert!(ok && out.contains("ドリフト 0 個"), "{out}");
    assert!(!dir.join("build/Main.bas").exists(), "stale build output must be removed");

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn errors_are_reported_with_source_position() {
    let dir = temp_dir("errors");
    assert!(ras(&dir, &["init", "."]).0);
    std::fs::write(dir.join("src/main.ras"), "export function f(): int {\n    const a = 1;\n    a = 2;\n    return a;\n}\n").unwrap();
    let (ok, out) = ras(&dir, &["build"]);
    assert!(!ok);
    assert!(out.contains("error[E0323]") && out.contains("src/main.ras:3:5"), "{out}");
    assert!(!dir.join("build").exists(), "nothing is written when there are errors");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn fmt_check_and_unimplemented_commands() {
    let dir = temp_dir("fmt");
    assert!(ras(&dir, &["init", "."]).0);
    std::fs::write(dir.join("src/main.ras"), "export function f() {\nMsgBox(\"x\");\n}\n").unwrap();
    let (ok, _) = ras(&dir, &["fmt", "--check"]);
    assert!(!ok);
    assert!(ras(&dir, &["fmt"]).0);
    assert_eq!(std::fs::read_to_string(dir.join("src/main.ras")).unwrap(), "export function f() {\n    MsgBox(\"x\");\n}\n");
    assert!(ras(&dir, &["fmt", "--check"]).0);
    // Commands that need live sync say so instead of pretending to work.
    let (ok, out) = ras(&dir, &["sync"]);
    assert!(!ok && out.contains("未実装"), "{out}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn bundled_example_project_builds() {
    let example = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../examples/sales-report");
    let out_dir = temp_dir("example");
    // Work on a copy so that the repository stays clean.
    for sub in ["src", "types"] {
        std::fs::create_dir_all(out_dir.join(sub)).unwrap();
        for entry in std::fs::read_dir(example.join(sub)).unwrap() {
            let entry = entry.unwrap();
            std::fs::copy(entry.path(), out_dir.join(sub).join(entry.file_name())).unwrap();
        }
    }
    std::fs::copy(example.join("ras.toml"), out_dir.join("ras.toml")).unwrap();
    let (ok, out) = ras(&out_dir, &["build"]);
    assert!(ok, "{out}");
    let bytes = std::fs::read(out_dir.join("build/Report.bas")).unwrap();
    let (text, _, _) = encoding_rs::SHIFT_JIS.decode(&bytes);
    // read loop → one Range read; formula loop → one Range write; the inline helper is expanded
    assert!(text.contains("RAS_Read2D(ws.Range(ws.Cells(2, 3), ws.Cells(last, 4)))"), "{text}");
    assert!(text.contains(".Formula = ras_buf"), "{text}");
    assert!(text.contains("' @inline nonNegative"), "{text}");
    assert!(out_dir.join("build/Sheet1.doccls").is_file());
    let _ = std::fs::remove_dir_all(&out_dir);
}
