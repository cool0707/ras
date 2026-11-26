//! User forms in `forms/` (spec §14.11): `forms/<Name>.form.json` (the
//! designer, a FormSpec) and `forms/<Name>.ras` (the code-behind, compiled by
//! the RAS compiler). The work is done by `ras_form::workspace`; this module
//! prints.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use ras_form::workspace::{self, Built, PullOutcome};

/// Loads, validates and generates the forms of `forms/`. Returns their number and the result.
/// `compiled`: the code the compiler generated for the `.ras` code-behinds (form name → code);
/// empty when only the designs are checked.
pub fn build(
    root: &Path,
    codepage: u16,
    compiled: &HashMap<String, String>,
) -> Result<(usize, Built), String> {
    let loaded = workspace::load(root)?;
    let built = workspace::build(&loaded, root, codepage, compiled);
    Ok((loaded.len(), built))
}

/// The code of the forms among the compiler's output files (form name → code).
pub fn compiled_code(files: &[ras_compiler::OutputFile]) -> HashMap<String, String> {
    files
        .iter()
        .filter(|f| f.kind == ras_compiler::vba::VModuleKind::Form)
        .map(|f| (f.name.clone(), f.code.clone()))
        .collect()
}

/// Prints the problems of the forms; returns the number of errors.
pub fn report(built: &Built) -> usize {
    for problem in &built.problems {
        eprintln!("{problem}\n");
    }
    built.errors()
}

/// `ras form new <Name>`.
pub fn new_form(root: &Path, name: &str) -> Result<(PathBuf, PathBuf), String> {
    if !workspace::valid_form_name(name) {
        return Err(format!("`{name}` はフォーム名に使えません（英字などで始まり、英数字と _ からなる 31 文字以内の名前にしてください）"));
    }
    let dir = workspace::forms_dir(root);
    let spec_path = dir.join(format!("{name}.form.json"));
    let code_path = dir.join(format!("{name}.ras"));
    for path in [&spec_path, &code_path] {
        if path.exists() {
            return Err(format!("{} は既に存在します", path.display()));
        }
    }
    if let Ok(existing) = workspace::load(root) {
        if let Some(other) = existing.iter().find(|f| f.name.eq_ignore_ascii_case(name)) {
            return Err(format!(
                "{} があります（フォーム名は大文字小文字を区別しません）",
                other.rel_path
            ));
        }
    }
    std::fs::create_dir_all(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    std::fs::write(&spec_path, ras_form::FormSpec::new_form(name).to_json())
        .map_err(|e| format!("{}: {e}", spec_path.display()))?;
    std::fs::write(&code_path, new_code_behind(name))
        .map_err(|e| format!("{}: {e}", code_path.display()))?;
    Ok((spec_path, code_path))
}

/// The code-behind of a new form: comments that show how handlers are written.
pub fn new_code_behind(name: &str) -> String {
    format!(
        "// フォーム {name} のコードビハインド。`this` がフォームで、コントロールは `this.<名前>` で参照します。\n\
         // イベントは @on(\"<コントロール名>.<イベント>\") を付けた関数で受け取ります（フォーム自身は UserForm）。\n\
         //\n\
         // @on(\"UserForm.Initialize\")\n\
         // function initialize() {{\n\
         // }}\n\
         //\n\
         // new {name}(...) の引数は constructor で受け取り、this.close(値) の値が showModal() の戻り値になります。\n\
         // function constructor(text: string) {{\n\
         // }}\n"
    )
}

fn relative(root: &Path, path: &Path) -> String {
    path.strip_prefix(root)
        .unwrap_or(path)
        .to_string_lossy()
        .replace('\\', "/")
}

/// `ras form pull`: writes `forms/` files for forms of the workbook (spec §14.11.6).
pub fn pull(root: &Path, book: &Path, names: &[String], force: bool) -> Result<ExitCode, String> {
    let at = |e: ras_ovba::OvbaError| format!("{}: {e}", book.display());
    let Some(project) = ras_ovba::unpack(book).map_err(at)?.project else {
        println!("{} に VBA プロジェクトはありません", book.display());
        return Ok(ExitCode::SUCCESS);
    };
    let in_book = ras_ovba::unpack_user_forms(book).map_err(at)?;
    let mut selected = Vec::new();
    if names.is_empty() {
        // The bundled form of the dialogs built in code (spec §14.11.15) is no form of `forms/`.
        selected.extend(
            in_book
                .iter()
                .filter(|f| !f.name.eq_ignore_ascii_case(ras_compiler::DIALOG_FORM)),
        );
    }
    for name in names {
        match in_book.iter().find(|f| f.name.eq_ignore_ascii_case(name)) {
            Some(form) => selected.push(form),
            None => {
                let known: Vec<&str> = in_book.iter().map(|f| f.name.as_str()).collect();
                return Err(format!(
                    "{} に `{name}` というユーザーフォームはありません（ブックのフォーム: {}）",
                    book.display(),
                    if known.is_empty() {
                        "なし".to_string()
                    } else {
                        known.join(", ")
                    }
                ));
            }
        }
    }
    if selected.is_empty() {
        println!("{} にユーザーフォームはありません", book.display());
        return Ok(ExitCode::SUCCESS);
    }
    let owner = book
        .file_name()
        .and_then(|n| n.to_str())
        .map(|n| book.with_file_name(format!("~${n}")));
    if owner.is_some_and(|o| o.exists()) {
        eprintln!(
            "warning: {} は Excel で開かれています。最後に保存した時点のフォームを取り出します",
            book.display()
        );
    }

    let (mut written, mut skipped, mut failed, mut hand_made) = (0, 0, 0, Vec::new());
    for form in selected {
        match workspace::pull(root, form, project.codepage, force) {
            Ok(PullOutcome::Written(pulled)) => {
                written += 1;
                let pictures = if pulled.pictures.is_empty() {
                    String::new()
                } else {
                    format!("、画像 {} 個", pulled.pictures.len())
                };
                if pulled.kept_code {
                    println!("{}: {} を書き出しました{pictures}（コードビハインドは既存の {} を使います）", form.name, relative(root, &pulled.spec_path), relative(root, &pulled.code_path));
                } else {
                    println!(
                        "{}: {} と {} を書き出しました{pictures}",
                        form.name,
                        relative(root, &pulled.spec_path),
                        relative(root, &pulled.code_path)
                    );
                }
                for warning in &pulled.warnings {
                    let place = if warning.control.is_empty() {
                        form.name.clone()
                    } else {
                        format!("{}.{}", form.name, warning.control)
                    };
                    eprintln!("warning: {place}: {}", warning.message);
                }
                for (line, message) in &pulled.todos {
                    eprintln!(
                        "warning: {}:{line}: VBA から RAS に変換できませんでした: {message}",
                        relative(root, &pulled.code_path)
                    );
                }
                if !pulled.todos.is_empty() {
                    eprintln!("  変換できなかったコードは `// TODO(VBA):` のコメントとして残しています。RAS で書き直してください");
                }
                if !ras_ovba::is_generated_form(&form.code) {
                    hand_made.push(form.name.clone());
                }
            }
            Ok(PullOutcome::Exists(files)) => {
                skipped += 1;
                let files: Vec<String> = files.iter().map(|f| relative(root, f)).collect();
                eprintln!("{}: {} が既にあるため書き出していません（置き換えるには --force を付けてください）", form.name, files.join(" と "));
            }
            Err(message) => {
                failed += 1;
                eprintln!("error: {message}");
            }
        }
    }
    println!("ユーザーフォーム {written} 個を forms/ に取り出しました（既存のためスキップ {skipped} 個、失敗 {failed} 個）");
    if !hand_made.is_empty() {
        println!(
            "次のフォームは VBE で作られたものです: {}\n  forms/ の内容でブックのフォームを置き換えるには、`ras pack --overwrite`（開いたままなら `ras sync --overwrite`）を実行してください",
            hand_made.join(", ")
        );
    }
    Ok(if failed > 0 {
        ExitCode::from(1)
    } else {
        ExitCode::SUCCESS
    })
}
