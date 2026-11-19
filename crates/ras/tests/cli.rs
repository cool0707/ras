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
        "/** 日本語の説明。 */\n@fast\nexport function fill() {\n    const ws = Sheet1;\n    for (let r = 1; r <= 10; r++) {\n        ws.cells(r, 1).value = r * 2;\n    }\n}\n\n@on(\"Sheet1.Change\")\nfunction changed(target: Excel.Range) {\n    status(\"変更: \" + target.address);\n}\n",
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

/// What only the VBE would report, or what would lose hand-written code, is refused before anything is written.
#[test]
fn pack_refuses_to_clobber_hand_written_code() {
    let dir = temp_dir("handwritten");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    let book = dir.join("Book.xlsm");
    // A workbook that already has a hand-written module `Main` and a hand-written `Workbook_Open`.
    ras_ovba::pack(&ras_ovba::PackRequest {
        book: book.clone(),
        codepage: 932,
        modules: vec![
            ras_ovba::VbaModule { name: "MAIN".into(), kind: ras_ovba::ModuleKind::Standard, source: "Option Explicit\r\nSub Legacy()\r\nEnd Sub\r\n".into() },
            ras_ovba::VbaModule { name: "ThisWorkbook".into(), kind: ras_ovba::ModuleKind::Document, source: "Private Sub Workbook_Open()\r\nEnd Sub\r\n".into() },
        ],
        ..Default::default()
    })
    .unwrap();
    let before = std::fs::read(&book).unwrap();

    let (ok, out) = ras(&dir, &["pack"]);
    assert!(!ok && out.contains("同名のモジュール") && out.contains("MAIN") && out.contains("--overwrite"), "{out}");
    assert_eq!(std::fs::read(&book).unwrap(), before, "nothing may be written");

    // The generated `Workbook_Open` (from `@init`) would be a second definition next to the hand-written one.
    std::fs::rename(dir.join("src/main.ras"), dir.join("src/app.ras")).unwrap();
    std::fs::write(dir.join("src/setup.ras"), "@init\nfunction prepare() {\n    status(\"ready\");\n}\n").unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(!ok && out.contains("ThisWorkbook") && out.contains("Workbook_Open") && out.contains("名前が重なります"), "{out}");
    assert_eq!(std::fs::read(&book).unwrap(), before, "nothing may be written");

    // A leftover owner file is named in the message.
    std::fs::write(dir.join("~$Book.xlsm"), b"").unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(!ok && out.contains("Excel で開かれている") && out.contains("~$Book.xlsm") && out.contains("削除"), "{out}");
    std::fs::remove_file(dir.join("~$Book.xlsm")).unwrap();

    // With --overwrite the module is replaced (the previous workbook is kept in build/backup/).
    std::fs::remove_file(dir.join("src/setup.ras")).unwrap();
    std::fs::rename(dir.join("src/app.ras"), dir.join("src/main.ras")).unwrap();
    let (ok, out) = ras(&dir, &["pack", "--overwrite"]);
    assert!(ok, "{out}");
    assert!(std::fs::read_dir(dir.join("build/backup")).unwrap().count() >= 1);
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
    std::fs::write(dir.join("src/main.ras"), "export function f() {\nmsgBox(\"x\");\n}\n").unwrap();
    let (ok, _) = ras(&dir, &["fmt", "--check"]);
    assert!(!ok);
    assert!(ras(&dir, &["fmt"]).0);
    assert_eq!(std::fs::read_to_string(dir.join("src/main.ras")).unwrap(), "export function f() {\n    msgBox(\"x\");\n}\n");
    assert!(ras(&dir, &["fmt", "--check"]).0);
    // Commands that need live sync say so instead of pretending to work.
    let (ok, out) = ras(&dir, &["sync", "--pull"]);
    assert!(!ok && out.contains("未実装"), "{out}");
    // Live sync needs the workbook open in Excel (and Windows): without it the command says what to do instead.
    let (ok, out) = ras(&dir, &["sync"]);
    assert!(!ok && out.contains("project.workbook"), "{out}");
    let book = temp_dir("sync-closed");
    assert!(ras(&book, &["init", ".", "--workbook", "Book.xlsm"]).0);
    let (ok, out) = ras(&book, &["sync"]);
    assert!(!ok && out.contains("ras pack"), "{out}");
    let _ = std::fs::remove_dir_all(&book);
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

#[test]
fn check_fix_rewrites_vba_spellings() {
    let dir = temp_dir("fix");
    assert!(ras(&dir, &["init", "."]).0);
    // Written the way VBA spells it, in a file with a byte-order mark.
    let old = "\u{feff}export function f(ws: Excel.Worksheet) {\n    ws.Cells(1, 1).Value = Application.Name;\n    MsgBox(\"済\");\n}\n";
    std::fs::write(dir.join("src/main.ras"), old).unwrap();
    let (ok, out) = ras(&dir, &["check"]);
    assert!(ok && out.contains("warning[W0304]") && out.contains("`cells`"), "{out}");
    assert_eq!(std::fs::read_to_string(dir.join("src/main.ras")).unwrap(), old, "`check` alone changes nothing");
    let (ok, out) = ras(&dir, &["check", "--fix"]);
    assert!(ok && out.contains("5 箇所を修正しました") && !out.contains("W0304"), "{out}");
    assert_eq!(
        std::fs::read_to_string(dir.join("src/main.ras")).unwrap(),
        "\u{feff}export function f(ws: Excel.Worksheet) {\n    ws.cells(1, 1).value = application.name;\n    msgBox(\"済\");\n}\n"
    );
    // The generated VBA is the same either way.
    assert!(ras(&dir, &["build"]).0);
    // (the module is written in the project's code page, not UTF-8)
    let text = String::from_utf8_lossy(&std::fs::read(dir.join("build/Main.bas")).unwrap()).to_string();
    assert!(text.contains("ws.Cells(1&, 1&).Value = Application.Name") && text.contains("MsgBox \""), "{text}");
}

/// `ras form new` / `ras form check` and writing forms with `ras pack` (spec §14.11).
#[test]
fn forms_are_created_checked_and_packed() {
    let dir = temp_dir("forms");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["form", "new", "frmInput"]);
    assert!(ok, "{out}");
    let spec_path = dir.join("forms/frmInput.form.json");
    let spec: serde_json::Value = serde_json::from_str(&std::fs::read_to_string(&spec_path).unwrap()).unwrap();
    assert_eq!(spec["kind"], "ras.userform");
    assert_eq!(spec["form"]["width"], 240.0);
    assert_eq!(std::fs::read_to_string(dir.join("forms/frmInput.bas")).unwrap(), "");
    let (ok, out) = ras(&dir, &["form", "new", "frmInput"]);
    assert!(!ok && out.contains("既に存在します"), "{out}");
    let (ok, out) = ras(&dir, &["form", "new", "1bad"]);
    assert!(!ok, "{out}");

    // Validation errors point into the file.
    let mut broken = spec.clone();
    broken["controls"] = serde_json::json!([{ "id": "a", "name": "txtA", "type": "TextBox", "colour": 1 }]);
    std::fs::write(&spec_path, serde_json::to_string_pretty(&broken).unwrap()).unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(!ok && out.contains("error[UFV001]") && out.contains("forms/frmInput.form.json:"), "{out}");
    let (ok, out) = ras(&dir, &["check"]);
    assert!(!ok && out.contains("UFV001"), "{out}");
    let (ok, out) = ras(&dir, &["build"]);
    assert!(!ok && out.contains("UFV001"), "{out}");

    let mut good = spec.clone();
    good["form"]["caption"] = "入力".into();
    good["controls"] = serde_json::json!([
        { "id": "name", "name": "txtName", "type": "TextBox", "left": 12, "top": 12 },
        { "id": "ok", "name": "btnOK", "type": "CommandButton", "caption": "OK", "left": 12, "top": 40 }
    ]);
    std::fs::write(&spec_path, serde_json::to_string_pretty(&good).unwrap()).unwrap();
    std::fs::write(dir.join("forms/frmInput.bas"), "Private Sub btnOK_Click()\n    Unload Me\nEnd Sub\n").unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok && out.contains("作成 0 / 更新 0 / 削除 0"), "second pack must be a no-op: {out}");

    let book = ras_ovba::unpack(&dir.join("Book.xlsm")).unwrap();
    let project = book.project.expect("VBA project");
    let form = project.modules.iter().find(|m| m.name == "frmInput").expect("form module");
    assert_eq!(form.kind, ras_ovba::ModuleKind::Form);
    assert!(form.source.contains("' RAS-FORM forms/frmInput.form.json\r\nPrivate Sub btnOK_Click()\r\n"), "{}", form.source);

    // Changing the design updates the form; unpack leaves generated forms alone.
    good["form"]["width"] = 300.into();
    std::fs::write(&spec_path, serde_json::to_string_pretty(&good).unwrap()).unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok && out.contains("更新 1"), "{out}");
    let (ok, out) = ras(&dir, &["unpack"]);
    assert!(ok, "{out}");
    assert!(!dir.join("vba/frmInput.frm").exists());

    // A form of the same name made in the VBE is not replaced without --overwrite.
    let other = temp_dir("forms-handmade");
    let (ok, out) = ras(&other, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    let mut hand_made = ras_form::FormSpec::new_form("frmInput");
    hand_made.form.caption = Some("VBE".into());
    let designer = ras_form::compile(&hand_made, 932).unwrap();
    ras_ovba::pack(&ras_ovba::PackRequest {
        book: other.join("Book.xlsm"),
        codepage: 932,
        forms: vec![ras_ovba::UserForm { name: "frmInput".into(), code: "Sub Old()\r\nEnd Sub\r\n".into(), streams: designer.streams, storages: designer.storages }],
        ..Default::default()
    })
    .unwrap();
    std::fs::create_dir_all(other.join("forms")).unwrap();
    std::fs::copy(&spec_path, other.join("forms/frmInput.form.json")).unwrap();
    let (ok, out) = ras(&other, &["pack"]);
    assert!(!ok && out.contains("frmInput") && out.contains("--overwrite"), "{out}");
    let (ok, out) = ras(&other, &["pack", "--overwrite"]);
    assert!(ok, "{out}");

    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_dir_all(&other);
}
