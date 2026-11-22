//! End-to-end test of the `ras` binary: init → check → build → pack → unpack.

use std::path::Path;
use std::process::Command;

fn ras(dir: &Path, args: &[&str]) -> (bool, String) {
    let out = Command::new(env!("CARGO_BIN_EXE_ras"))
        .args(args)
        .current_dir(dir)
        .output()
        .expect("run ras");
    let text = format!(
        "{}{}",
        String::from_utf8_lossy(&out.stdout),
        String::from_utf8_lossy(&out.stderr)
    );
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
    assert!(
        dir.join("ras.toml").is_file()
            && dir.join("src/main.ras").is_file()
            && dir.join("Book.xlsm").is_file()
    );

    std::fs::write(
        dir.join("src/report.ras"),
        "/** 日本語の説明。 */\n@fast\nexport function fill() {\n    const ws = Sheet1;\n    for (let r = 1; r <= 10; r++) {\n        ws.cells(r, 1).value = r * 2;\n    }\n}\n",
    )
    .unwrap();
    // The module of the sheet (spec §12.2).
    std::fs::create_dir_all(dir.join("objects")).unwrap();
    std::fs::write(dir.join("objects/Sheet1.ras"), "@on(\"Worksheet.Change\")\nfunction changed(target: Excel.Range) {\n    status(\"変更: \" + target.address);\n}\n").unwrap();

    let (ok, out) = ras(&dir, &["check"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["build"]);
    assert!(ok, "{out}");
    // Module files are written in the target code page (CP932), not UTF-8.
    let bytes = std::fs::read(dir.join("build/Report.bas")).unwrap();
    let (text, _, bad) = encoding_rs::SHIFT_JIS.decode(&bytes);
    assert!(!bad);
    assert!(
        text.contains("Attribute VB_Name = \"Report\"\r\n"),
        "{text}"
    );
    assert!(text.contains("' 日本語の説明。"), "{text}");
    assert!(
        text.contains("ws.Range(ws.Cells(1, 1), ws.Cells(10, 1)).Value = ras_buf1"),
        "{text}"
    );
    assert!(
        std::str::from_utf8(&bytes).is_err(),
        "file must not be UTF-8"
    );
    assert!(dir.join("build/Report.rasmap").is_file());
    assert!(dir.join("build/Sheet1.doccls").is_file());

    // The sheet code names come from the workbook: an unknown sheet is rejected.
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(out.contains("作成"), "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(
        out.contains("作成 0 / 更新 0 / 削除 0"),
        "second pack must be a no-op: {out}"
    );
    // A relative `--project`: the sheet code names are still read from the workbook.
    let (ok, out) = ras(
        dir.parent().unwrap(),
        &[
            "--project",
            dir.file_name().unwrap().to_str().unwrap(),
            "pack",
        ],
    );
    assert!(ok && out.contains("作成 0 / 更新 0 / 削除 0"), "{out}");

    // Read the workbook back with the offline reader.
    let book = ras_ovba::unpack(&dir.join("Book.xlsm")).unwrap();
    let project = book.project.expect("VBA project");
    assert_eq!(project.codepage, 932);
    let names: Vec<&str> = project.modules.iter().map(|m| m.name.as_str()).collect();
    for expected in [
        "ThisWorkbook",
        "Sheet1",
        "Main",
        "Report",
        "RAS_Runtime",
        "RAS_Glue",
    ] {
        assert!(
            names.contains(&expected),
            "{expected} missing from {names:?}"
        );
    }
    let report = project.modules.iter().find(|m| m.name == "Report").unwrap();
    assert!(report.source.contains("' 日本語の説明。"));
    let sheet = project.modules.iter().find(|m| m.name == "Sheet1").unwrap();
    assert!(
        sheet.source.contains("' RAS-GENERATED hash=")
            && sheet.source.contains("Option Explicit")
            && sheet
                .source
                .contains("Private Sub Worksheet_Change(ByVal Target As Range)"),
        "{}",
        sheet.source
    );
    assert!(
        !sheet.source.contains("'--- RAS BEGIN ---"),
        "{}",
        sheet.source
    );

    // Removing a source file removes its module; unpack reports only RAS-managed modules.
    std::fs::remove_file(dir.join("src/main.ras")).unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok && out.contains("削除 1"), "{out}");
    let (ok, out) = ras(&dir, &["unpack"]);
    assert!(ok && out.contains("ドリフト 0 個"), "{out}");
    assert!(
        !dir.join("build/Main.bas").exists(),
        "stale build output must be removed"
    );

    let _ = std::fs::remove_dir_all(&dir);
}

/// What only the VBE would report, or what would lose hand-written code, is refused before anything is written.
#[test]
fn pack_refuses_to_clobber_hand_written_code() {
    let dir = temp_dir("handwritten");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    // Without the take-over of hand-written code (see `hand_written_vba_is_taken_over`).
    let mut toml = std::fs::read_to_string(dir.join("ras.toml")).unwrap();
    toml.push_str("\n[sync]\npull-unmanaged = false\n");
    std::fs::write(dir.join("ras.toml"), toml).unwrap();
    let book = dir.join("Book.xlsm");
    // A workbook that already has a hand-written module `Main` and a hand-written `Workbook_Open`.
    ras_ovba::pack(&ras_ovba::PackRequest {
        book: book.clone(),
        codepage: 932,
        modules: vec![
            ras_ovba::VbaModule {
                name: "MAIN".into(),
                kind: ras_ovba::ModuleKind::Standard,
                source: "Option Explicit\r\nSub Legacy()\r\nEnd Sub\r\n".into(),
            },
            ras_ovba::VbaModule {
                name: "ThisWorkbook".into(),
                kind: ras_ovba::ModuleKind::Document,
                source: "Private Sub Workbook_Open()\r\nEnd Sub\r\n".into(),
            },
        ],
        ..Default::default()
    })
    .unwrap();
    let before = std::fs::read(&book).unwrap();

    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        !ok && out.contains("同名のモジュール")
            && out.contains("MAIN")
            && out.contains("--overwrite"),
        "{out}"
    );
    assert_eq!(
        std::fs::read(&book).unwrap(),
        before,
        "nothing may be written"
    );

    // The generated `ThisWorkbook` (the calls of `@init`) would replace the hand-written `Workbook_Open`.
    std::fs::rename(dir.join("src/main.ras"), dir.join("src/app.ras")).unwrap();
    std::fs::write(
        dir.join("src/setup.ras"),
        "@init\nfunction prepare() {\n    status(\"ready\");\n}\n",
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        !ok && out.contains("ThisWorkbook")
            && out.contains("同名のモジュール")
            && out.contains("--overwrite"),
        "{out}"
    );
    assert_eq!(
        std::fs::read(&book).unwrap(),
        before,
        "nothing may be written"
    );

    // A leftover owner file is named in the message.
    std::fs::write(dir.join("~$Book.xlsm"), b"").unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        !ok && out.contains("Excel で開かれている")
            && out.contains("~$Book.xlsm")
            && out.contains("削除"),
        "{out}"
    );
    std::fs::remove_file(dir.join("~$Book.xlsm")).unwrap();

    // With --overwrite the module is replaced (the previous workbook is kept in build/backup/).
    std::fs::remove_file(dir.join("src/setup.ras")).unwrap();
    std::fs::rename(dir.join("src/app.ras"), dir.join("src/main.ras")).unwrap();
    let (ok, out) = ras(&dir, &["pack", "--overwrite"]);
    assert!(ok, "{out}");
    assert!(std::fs::read_dir(dir.join("build/backup")).unwrap().count() >= 1);
    let _ = std::fs::remove_dir_all(&dir);
}

/// The hand-written VBA of a workbook — standard modules, the code of sheets and of
/// `ThisWorkbook`, forms made in the VBE — is converted to RAS and taken over (spec §14.7.1).
#[test]
fn hand_written_vba_is_taken_over() {
    let dir = temp_dir("takeover");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    std::fs::remove_file(dir.join("src/main.ras")).unwrap();
    let book = dir.join("Book.xlsm");
    let mut spec = ras_form::FormSpec::new_form("frmHello");
    spec.form.caption = Some("こんにちは".into());
    let designer = ras_form::compile(&spec, 932).unwrap();
    ras_ovba::pack(&ras_ovba::PackRequest {
        book: book.clone(),
        codepage: 932,
        modules: vec![
            ras_ovba::VbaModule {
                name: "mod_Util".into(),
                kind: ras_ovba::ModuleKind::Standard,
                source: "Attribute VB_Name = \"mod_Util\"\r\nOption Explicit\r\nPublic Function Twice(ByVal n As Long) As Long\r\n    Twice = n * 2\r\nEnd Function\r\n".into(),
            },
            ras_ovba::VbaModule {
                name: "Module1".into(),
                kind: ras_ovba::ModuleKind::Standard,
                source: "Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n    MsgBox mod_Util.Twice(21)\r\n    frmHello.Show\r\nEnd Sub\r\nSub Jump()\r\n    GoTo Done\r\nDone:\r\nEnd Sub\r\n".into(),
            },
            ras_ovba::VbaModule {
                name: "Sheet1".into(),
                kind: ras_ovba::ModuleKind::Document,
                source: "Option Explicit\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n    Range(\"B1\").Value = Twice(Target.Value)\r\nEnd Sub\r\n".into(),
            },
        ],
        forms: vec![ras_ovba::UserForm { name: "frmHello".into(), code: "Private Sub UserForm_Click()\r\n    Me.Hide\r\nEnd Sub\r\n".into(), streams: designer.streams, storages: designer.storages }],
        ..Default::default()
    })
    .unwrap();

    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    for file in [
        "src/mod_Util.ras",
        "src/Module1.ras",
        "objects/Sheet1.ras",
        "forms/frmHello.form.json",
        "forms/frmHello.ras",
        "vba/mod_Util.bas",
        "vba/Module1.bas",
        "vba/Sheet1.cls",
        "vba/frmHello.frm",
    ] {
        assert!(dir.join(file).is_file(), "{file} missing: {out}");
    }
    assert!(
        out.contains("src/Module1.ras")
            && out.contains("TODO(VBA)")
            && out.contains("`Module1` は置き換えずに"),
        "{out}"
    );
    let read = |file: &str| std::fs::read_to_string(dir.join(file)).unwrap();
    // Taken over: the marker line became a comment, the hand-written code is gone from the workbook.
    for file in [
        "src/mod_Util.ras",
        "objects/Sheet1.ras",
        "forms/frmHello.ras",
    ] {
        assert!(
            !read(file).contains("RAS-IMPORTED") && read(file).starts_with("// ブックの VBA（"),
            "{file}: {}",
            read(file)
        );
    }
    assert!(
        read("objects/Sheet1.ras").contains("@on(\"Worksheet.Change\")"),
        "{}",
        read("objects/Sheet1.ras")
    );
    assert!(
        read("objects/Sheet1.ras").contains("this.range(\"B1\").value"),
        "{}",
        read("objects/Sheet1.ras")
    );
    assert!(
        read("objects/Sheet1.ras").contains("import { Twice } from \"../mod_Util\";"),
        "{}",
        read("objects/Sheet1.ras")
    );
    // Pending: `Module1` has a TODO, the workbook keeps it.
    let module1 = read("src/Module1.ras");
    assert!(
        module1.starts_with("// RAS-IMPORTED Module1 ") && module1.contains("// TODO(VBA):"),
        "{module1}"
    );
    assert!(
        module1.contains("import * as ModUtil from \"./mod_Util\";")
            && module1.contains("ModUtil.Twice(21)"),
        "{module1}"
    );

    let source_of = |name: &str| {
        let project = ras_ovba::unpack(&book).unwrap().project.unwrap();
        project
            .modules
            .iter()
            .find(|m| m.name == name)
            .map(|m| m.source.clone())
    };
    assert!(source_of("mod_Util").is_none(), "replaced by ModUtil");
    assert!(source_of("ModUtil").unwrap().contains("' RAS-GENERATED"));
    assert!(
        source_of("Module1").unwrap().contains("GoTo Done"),
        "pending: kept as it was"
    );
    let sheet = source_of("Sheet1").unwrap();
    // The module of the sheet is generated whole from `objects/Sheet1.ras`.
    assert!(
        sheet.contains("' RAS-GENERATED hash=")
            && sheet.contains("Me.Range(\"B1\")")
            && !sheet.contains("'--- RAS BEGIN ---")
            && sheet.matches("Option Explicit").count() == 1,
        "{sheet}"
    );
    let form = ras_ovba::unpack_user_forms(&book)
        .unwrap()
        .into_iter()
        .find(|f| f.name == "frmHello")
        .unwrap();
    assert!(ras_ovba::is_generated_form(&form.code), "{}", form.code);

    // The TODO is fixed: the next pack replaces `Module1`.
    let fixed: String = module1
        .lines()
        .filter(|l| !l.trim_start().starts_with("//"))
        .map(|l| format!("{l}\n"))
        .collect();
    std::fs::write(
        dir.join("src/Module1.ras"),
        format!("{}\n{fixed}", module1.lines().next().unwrap()),
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(
        source_of("Module1").unwrap().contains("' RAS-GENERATED"),
        "{out}"
    );
    assert!(!read("src/Module1.ras").contains("RAS-IMPORTED"));
    // Nothing more to take over.
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        ok && !out.contains("取り込みました") && out.contains("作成 0 / 更新 0 / 削除 0"),
        "{out}"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

/// Code added in the VBE to modules RAS generated is taken over into their files (spec §14.7.3).
#[test]
fn code_added_in_the_vbe_is_taken_over() {
    let dir = temp_dir("added");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    std::fs::create_dir_all(dir.join("objects")).unwrap();
    let sheet_ras = "@on(\"Worksheet.Change\")\nfunction changed(target: Excel.Range) {\n    this.range(\"D1\").value = target.address;\n}\n";
    std::fs::write(dir.join("objects/Sheet1.ras"), sheet_ras).unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    // What was written is kept, to compare with later.
    assert!(
        dir.join("build/sync/Sheet1.vba").is_file() && dir.join("build/sync/Main.vba").is_file()
    );
    let book = dir.join("Book.xlsm");
    let source_of = |name: &str| {
        ras_ovba::unpack(&book)
            .unwrap()
            .project
            .unwrap()
            .modules
            .into_iter()
            .find(|m| m.name == name)
            .unwrap()
            .source
    };
    let read = |file: &str| std::fs::read_to_string(dir.join(file)).unwrap();
    // As the VBE would: a procedure and a variable in the sheet's module, a procedure in `Main`.
    let edit = |name: &str, kind: ras_ovba::ModuleKind, change: &dyn Fn(String) -> String| {
        ras_ovba::pack(&ras_ovba::PackRequest {
            book: book.clone(),
            codepage: 932,
            modules: vec![ras_ovba::VbaModule {
                name: name.into(),
                kind,
                source: change(source_of(name)),
            }],
            ..Default::default()
        })
        .unwrap();
    };
    edit("Sheet1", ras_ovba::ModuleKind::Document, &|s| {
        format!("{}\r\nPrivate Sub Worksheet_Activate()\r\n    Range(\"A1\").Value = Counter\r\nEnd Sub\r\n", s.replace("Option Compare Binary\r\n", "Option Compare Binary\r\nPrivate Counter As Long\r\n"))
    });
    edit("Main", ras_ovba::ModuleKind::Standard, &|s| {
        format!("{s}\r\nPublic Sub Extra()\r\n    MsgBox \"extra\"\r\nEnd Sub\r\n")
    });
    assert!(source_of("Sheet1").contains("Private Counter As Long"));

    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(
        out.contains("VBE で足されたコード")
            && out.contains("objects/Sheet1.ras")
            && out.contains("Worksheet_Activate")
            && out.contains("src/main.ras")
            && out.contains("Extra"),
        "{out}"
    );
    let sheet = read("objects/Sheet1.ras");
    assert!(
        sheet.starts_with(sheet_ras)
            && sheet.contains("// VBE で追加されたコードを取り込みました（"),
        "{sheet}"
    );
    assert!(
        sheet.contains("let Counter: int;")
            && sheet.contains("@on(\"Worksheet.Activate\")")
            && sheet.contains("this.range(\"A1\").value = Counter;"),
        "{sheet}"
    );
    assert!(
        read("src/main.ras").contains("export function Extra() {"),
        "{}",
        read("src/main.ras")
    );
    assert!(read("vba/Sheet1.cls").contains("Private Sub Worksheet_Activate()"));
    // The modules are written again from their files: no drift is left.
    assert!(
        ras_compiler::emit::has_drifted(&source_of("Sheet1")) == Some(false)
            && source_of("Sheet1").contains("Private Sub Worksheet_Activate()"),
        "{}",
        source_of("Sheet1")
    );
    assert_eq!(
        ras_compiler::emit::has_drifted(&source_of("Main")),
        Some(false)
    );
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        ok && !out.contains("取り込みました") && out.contains("作成 0 / 更新 0 / 削除 0"),
        "{out}"
    );

    // What does not compile as RAS is not taken over: the module stays as drift.
    let before = read("src/main.ras");
    edit("Main", ras_ovba::ModuleKind::Standard, &|s| {
        format!("{s}\r\nPublic Sub Broken()\r\n    NoSuchProcedure 1\r\nEnd Sub\r\n")
    });
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        !ok && out.contains("コンパイル エラーになるため取り込みませんでした")
            && out.contains("ドリフト"),
        "{out}"
    );
    assert_eq!(read("src/main.ras"), before);

    // A procedure RAS wrote was changed: drift; what was added next to it is taken over all the same.
    let (ok, out) = ras(&dir, &["pack", "--overwrite"]);
    assert!(ok, "{out}");
    edit("Sheet1", ras_ovba::ModuleKind::Document, &|s| {
        format!(
            "{}\r\nPrivate Sub Worksheet_Deactivate()\r\nEnd Sub\r\n",
            s.replace("changed Target", "changed Target: Beep")
        )
    });
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        !ok && out.contains("ドリフト")
            && out.contains("Sheet1")
            && out.contains("Worksheet_Deactivate"),
        "{out}"
    );
    assert!(read("objects/Sheet1.ras").contains("@on(\"Worksheet.Deactivate\")"));
    assert!(source_of("Sheet1").contains("Beep"), "left as it is");
    let _ = std::fs::remove_dir_all(&dir);
}

/// `ras sheet` in a workbook with a VBA project: the document modules follow the sheets, and
/// `ras pack` writes them from the files of `objects/` (spec §14.16).
#[test]
fn sheets_and_their_modules() {
    let dir = temp_dir("sheet-modules");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    let book = dir.join("Book.xlsm");
    let modules = || {
        ras_ovba::unpack(&book)
            .unwrap()
            .project
            .map(|p| {
                p.modules
                    .into_iter()
                    .map(|m| (m.name, m.source))
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default()
    };
    let source = |name: &str| {
        modules()
            .into_iter()
            .find(|(n, _)| n == name)
            .map(|(_, s)| s)
    };
    assert!(ras(&dir, &["pack"]).0);

    let (ok, out) = ras(&dir, &["sheet", "add", "集計"]);
    assert!(ok && out.contains("objects/Sheet2.ras"), "{out}");
    assert!(source("Sheet2").is_some(), "the new sheet has its module");
    std::fs::write(
        dir.join("objects/Sheet2.ras"),
        "@on(\"Worksheet.Activate\")\nfunction shown() {\n    this.range(\"A1\").value = 1;\n}\n",
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(source("Sheet2")
        .unwrap()
        .contains("Private Sub Worksheet_Activate()"));

    // The code name: the module, the file and what was written follow.
    let (ok, out) = ras(&dir, &["sheet", "rename", "集計", "--code-name", "Summary"]);
    assert!(
        ok && out.contains("objects/Sheet2.ras → objects/Summary.ras"),
        "{out}"
    );
    assert!(
        source("Sheet2").is_none()
            && source("Summary")
                .unwrap()
                .contains("Private Sub Worksheet_Activate()"),
        "{:?}",
        modules().iter().map(|m| &m.0).collect::<Vec<_>>()
    );
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(
        source("Summary")
            .unwrap()
            .contains("RAS_EventError \"Summary.Activate\""),
        "{}",
        source("Summary").unwrap()
    );

    // A copy gets the file, and its module the code at the next pack.
    let (ok, out) = ras(
        &dir,
        &[
            "sheet",
            "copy",
            "Summary",
            "集計の写し",
            "--code-name",
            "Summary2",
        ],
    );
    assert!(ok, "{out}");
    assert!(ras(&dir, &["pack"]).0);
    assert!(source("Summary2")
        .unwrap()
        .contains("Private Sub Worksheet_Activate()"));

    let (ok, out) = ras(&dir, &["sheet", "remove", "Summary", "--force"]);
    assert!(ok, "{out}");
    assert!(
        source("Summary").is_none() && source("Summary2").is_some() && source("Main").is_some()
    );
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok && out.contains("作成 0 / 更新 0 / 削除 0"), "{out}");
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn errors_are_reported_with_source_position() {
    let dir = temp_dir("errors");
    assert!(ras(&dir, &["init", "."]).0);
    std::fs::write(
        dir.join("src/main.ras"),
        "export function f(): int {\n    const a = 1;\n    a = 2;\n    return a;\n}\n",
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["build"]);
    assert!(!ok);
    assert!(
        out.contains("error[E0323]") && out.contains("src/main.ras:3:5"),
        "{out}"
    );
    assert!(
        !dir.join("build").exists(),
        "nothing is written when there are errors"
    );
    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn fmt_check_and_unimplemented_commands() {
    let dir = temp_dir("fmt");
    assert!(ras(&dir, &["init", "."]).0);
    std::fs::write(
        dir.join("src/main.ras"),
        "export function f() {\nmsgBox(\"x\");\n}\n",
    )
    .unwrap();
    let (ok, _) = ras(&dir, &["fmt", "--check"]);
    assert!(!ok);
    assert!(ras(&dir, &["fmt"]).0);
    assert_eq!(
        std::fs::read_to_string(dir.join("src/main.ras")).unwrap(),
        "export function f() {\n    msgBox(\"x\");\n}\n"
    );
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
    for sub in ["src", "types", "objects"] {
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
    assert!(
        text.contains("RAS_Read2D(ws.Range(ws.Cells(2, 3), ws.Cells(last, 4)))"),
        "{text}"
    );
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
    assert!(
        ok && out.contains("warning[W0304]") && out.contains("`cells`"),
        "{out}"
    );
    assert_eq!(
        std::fs::read_to_string(dir.join("src/main.ras")).unwrap(),
        old,
        "`check` alone changes nothing"
    );
    let (ok, out) = ras(&dir, &["check", "--fix"]);
    assert!(
        ok && out.contains("5 箇所を修正しました") && !out.contains("W0304"),
        "{out}"
    );
    assert_eq!(
        std::fs::read_to_string(dir.join("src/main.ras")).unwrap(),
        "\u{feff}export function f(ws: Excel.Worksheet) {\n    ws.cells(1, 1).value = application.name;\n    msgBox(\"済\");\n}\n"
    );
    // The generated VBA is the same either way.
    assert!(ras(&dir, &["build"]).0);
    // (the module is written in the project's code page, not UTF-8)
    let text =
        String::from_utf8_lossy(&std::fs::read(dir.join("build/Main.bas")).unwrap()).to_string();
    assert!(
        text.contains("ws.Cells(1&, 1&).Value = Application.Name") && text.contains("MsgBox \""),
        "{text}"
    );
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
    let spec: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(&spec_path).unwrap()).unwrap();
    assert_eq!(spec["kind"], "ras.userform");
    assert_eq!(spec["form"]["width"], 240.0);
    // The code-behind is RAS: a new one only holds comments.
    let template = std::fs::read_to_string(dir.join("forms/frmInput.ras")).unwrap();
    assert!(template.lines().all(|l| l.starts_with("//")), "{template}");
    let (ok, out) = ras(&dir, &["form", "new", "frmInput"]);
    assert!(!ok && out.contains("既に存在します"), "{out}");
    let (ok, out) = ras(&dir, &["form", "new", "1bad"]);
    assert!(!ok, "{out}");

    // Validation errors point into the file.
    let mut broken = spec.clone();
    broken["controls"] =
        serde_json::json!([{ "id": "a", "name": "txtA", "type": "TextBox", "colour": 1 }]);
    std::fs::write(&spec_path, serde_json::to_string_pretty(&broken).unwrap()).unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(
        !ok && out.contains("error[UFV001]") && out.contains("forms/frmInput.form.json:"),
        "{out}"
    );
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
    // Errors in the code-behind are compile errors of the project.
    std::fs::write(
        dir.join("forms/frmInput.ras"),
        "@on(\"btnOK.Click\")\nfunction ok() {\n    this.txtNone.text = \"\";\n}\n",
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(!ok && out.contains("forms/frmInput.ras:3:"), "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(!ok && out.contains("forms/frmInput.ras:3:"), "{out}");
    std::fs::write(
        dir.join("forms/frmInput.ras"),
        "@on(\"btnOK.Click\")\nfunction ok() {\n    unload(this);\n}\n",
    )
    .unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["check"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        ok && out.contains("作成 0 / 更新 0 / 削除 0"),
        "second pack must be a no-op: {out}"
    );

    let book = ras_ovba::unpack(&dir.join("Book.xlsm")).unwrap();
    let project = book.project.expect("VBA project");
    let form = project
        .modules
        .iter()
        .find(|m| m.name == "frmInput")
        .expect("form module");
    assert_eq!(form.kind, ras_ovba::ModuleKind::Form);
    let code: Vec<&str> = form
        .source
        .split("\r\n")
        .skip_while(|l| l.starts_with("Attribute "))
        .collect();
    assert!(
        code[0].starts_with("' RAS-FORM forms/frmInput.form.json design=")
            && code[0].contains(" code="),
        "{}",
        form.source
    );
    assert!(
        form.source.contains("Private Sub btnOK_Click()\r\n"),
        "{}",
        form.source
    );
    assert!(form.source.contains("Unload Me\r\n"), "{}", form.source);
    assert!(
        !dir.join("build/frmInput.frmcode").exists(),
        "forms are not build products"
    );

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
        forms: vec![ras_ovba::UserForm {
            name: "frmInput".into(),
            code: "Sub Old()\r\nEnd Sub\r\n".into(),
            streams: designer.streams,
            storages: designer.storages,
        }],
        ..Default::default()
    })
    .unwrap();
    std::fs::create_dir_all(other.join("forms")).unwrap();
    std::fs::copy(&spec_path, other.join("forms/frmInput.form.json")).unwrap();
    let (ok, out) = ras(&other, &["pack"]);
    assert!(
        !ok && out.contains("frmInput") && out.contains("--overwrite"),
        "{out}"
    );
    let (ok, out) = ras(&other, &["pack", "--overwrite"]);
    assert!(ok, "{out}");

    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_dir_all(&other);
}

/// A 2×2 bitmap, 24 bits per pixel.
fn tiny_bmp() -> Vec<u8> {
    let mut data = vec![0u8; 54 + 16];
    data[..2].copy_from_slice(b"BM");
    data[2..6].copy_from_slice(&70u32.to_le_bytes());
    data[10..14].copy_from_slice(&54u32.to_le_bytes());
    data[14..18].copy_from_slice(&40u32.to_le_bytes());
    data[18..22].copy_from_slice(&2i32.to_le_bytes());
    data[22..26].copy_from_slice(&2i32.to_le_bytes());
    data[26..28].copy_from_slice(&1u16.to_le_bytes());
    data[28..30].copy_from_slice(&24u16.to_le_bytes());
    data[54..].copy_from_slice(&[0, 0, 255, 0, 255, 0, 0, 0, 255, 0, 0, 255, 255, 255, 0, 0]);
    data
}

/// Pictures and list items are written by `ras pack`, and `ras form pull` takes a form back
/// out of the workbook (spec §14.11).
#[test]
fn form_pictures_lists_and_pull() {
    let dir = temp_dir("forms-pull");
    let (ok, out) = ras(&dir, &["init", ".", "--workbook", "Book.xlsm"]);
    assert!(ok, "{out}");
    std::fs::create_dir_all(dir.join("forms/assets")).unwrap();
    std::fs::write(dir.join("forms/assets/logo.bmp"), tiny_bmp()).unwrap();
    let spec = serde_json::json!({
        "schemaVersion": 1, "kind": "ras.userform", "basis": "designer",
        "form": { "name": "frmOrder", "caption": "注文", "width": 240, "height": 180 },
        "controls": [
            { "id": "logo", "name": "imgLogo", "type": "Image", "left": 6, "top": 6, "width": 32, "height": 32, "picture": { "path": "forms/assets/logo.bmp" }, "pictureSizeMode": 3 },
            { "id": "kind", "name": "cboKind", "type": "ComboBox", "left": 48, "top": 6, "list": ["通常", "お急ぎ"], "selectedIndex": 0 },
            { "id": "items", "name": "lstItems", "type": "ListBox", "left": 6, "top": 48, "list": ["A", "B", "C"] },
            { "id": "note", "name": "txtNote", "type": "TextBox", "left": 6, "top": 130, "tag": "note", "controlTipText": "備考" }
        ]
    });
    let spec_path = dir.join("forms/frmOrder.form.json");
    std::fs::write(&spec_path, serde_json::to_string_pretty(&spec).unwrap()).unwrap();
    let hand =
        "@on(\"UserForm.Initialize\")\nfunction init() {\n    this.txtNote.text = \"\";\n}\n";
    std::fs::write(dir.join("forms/frmOrder.ras"), hand).unwrap();
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(ok, "{out}");
    assert!(
        !out.contains("UFV013"),
        "list and tag are supported fields: {out}"
    );

    // The picture is in the designer storage, the list items in the code.
    let book = dir.join("Book.xlsm");
    let stored = ras_ovba::unpack_user_forms(&book).unwrap();
    let form = stored
        .iter()
        .find(|f| f.name == "frmOrder")
        .expect("the form");
    let bmp = tiny_bmp();
    assert!(
        form.streams["o"]
            .windows(bmp.len())
            .any(|w| w == bmp.as_slice()),
        "picture not stored"
    );
    assert!(form.code.contains("Private Sub RAS_FormInit()\n    With Me.cboKind\n        .AddItem \"通常\"\n        .AddItem \"お急ぎ\"\n        .ListIndex = 0\n    End With\n"), "{}", form.code);
    assert!(
        form.code
            .contains("Private Sub UserForm_Initialize()\n    RAS_FormInit '--- RAS ---\n"),
        "{}",
        form.code
    );
    assert!(
        form.code.contains("Me.txtNote.Text = \"\""),
        "{}",
        form.code
    );

    // A selection has to name an item; a broken picture is an error that points into the file.
    let mut bad = spec.clone();
    bad["controls"][1]["selectedIndex"] = 5.into();
    std::fs::write(&spec_path, serde_json::to_string_pretty(&bad).unwrap()).unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(!ok && out.contains("error[UFV020]"), "{out}");
    let mut bad = spec.clone();
    bad["controls"][0]["picture"]["path"] = "forms/assets/missing.bmp".into();
    std::fs::write(&spec_path, serde_json::to_string_pretty(&bad).unwrap()).unwrap();
    let (ok, out) = ras(&dir, &["form", "check"]);
    assert!(
        !ok && out.contains("missing.bmp") && out.contains("forms/frmOrder.form.json:"),
        "{out}"
    );
    std::fs::write(&spec_path, serde_json::to_string_pretty(&spec).unwrap()).unwrap();

    // Pull into an empty project: design, code-behind and picture come back.
    let other = temp_dir("forms-pull-target");
    let (ok, out) = ras(&other, &["init", "."]);
    assert!(ok, "{out}");
    let book_arg = book.to_string_lossy().to_string();
    let (ok, out) = ras(&other, &["form", "pull", "--book", &book_arg]);
    assert!(
        ok && out.contains("frmOrder") && out.contains("1 個を forms/ に取り出しました"),
        "{out}"
    );
    // The code comes back as RAS: the generated event procedure becomes the handler again.
    let pulled_code = std::fs::read_to_string(other.join("forms/frmOrder.ras")).unwrap();
    assert!(
        pulled_code.contains(
            "@on(\"UserForm.Initialize\")\nfunction init() {\n    this.txtNote.text = \"\";\n}\n"
        ),
        "{pulled_code}"
    );
    assert_eq!(
        std::fs::read(other.join("forms/assets/frmOrder.imgLogo.bmp")).unwrap(),
        bmp
    );
    let pulled: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string(other.join("forms/frmOrder.form.json")).unwrap(),
    )
    .unwrap();
    assert_eq!(pulled["form"]["caption"], "注文");
    assert_eq!(
        pulled["controls"][0]["picture"]["path"],
        "forms/assets/frmOrder.imgLogo.bmp"
    );
    assert_eq!(pulled["controls"][0]["pictureSizeMode"], 3);
    assert_eq!(pulled["controls"][3]["tag"], "note");
    assert_eq!(pulled["controls"][3]["controlTipText"], "備考");
    assert!(
        pulled["controls"][1].get("list").is_none(),
        "list items are not stored in the workbook"
    );
    let (ok, out) = ras(&other, &["form", "check"]);
    assert!(ok, "the pulled form is a valid form: {out}");

    // Existing files stay unless --force is given; unknown names are reported.
    let (ok, out) = ras(&other, &["form", "pull", "frmOrder", "--book", &book_arg]);
    assert!(
        ok && out.contains("既にあるため") && out.contains("0 個を forms/ に取り出しました"),
        "{out}"
    );
    let (ok, out) = ras(&other, &["form", "pull", "frmNone", "--book", &book_arg]);
    assert!(
        !ok && out.contains("frmNone") && out.contains("frmOrder"),
        "{out}"
    );

    // Pulling over the project the form came from keeps the list items.
    let (ok, out) = ras(&dir, &["form", "pull", "--force"]);
    assert!(ok, "{out}");
    let again: serde_json::Value =
        serde_json::from_str(&std::fs::read_to_string(&spec_path).unwrap()).unwrap();
    assert_eq!(
        again["controls"][1]["list"],
        serde_json::json!(["通常", "お急ぎ"])
    );
    assert_eq!(again["controls"][1]["selectedIndex"], 0);
    assert_eq!(
        again["controls"][0]["picture"]["path"],
        "forms/assets/logo.bmp"
    );
    // The code-behind the form was compiled from stays as it is.
    assert!(out.contains("既存の forms/frmOrder.ras"), "{out}");
    assert_eq!(
        std::fs::read_to_string(dir.join("forms/frmOrder.ras")).unwrap(),
        hand
    );
    let (ok, out) = ras(&dir, &["pack"]);
    assert!(
        ok && out.contains("作成 0 / 更新 0 / 削除 0"),
        "pulling a generated form back changes nothing: {out}"
    );

    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_dir_all(&other);
}
