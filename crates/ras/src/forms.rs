//! User forms in `forms/` (spec §14.11): `forms/<Name>.form.json` (the
//! designer, a FormSpec) and `forms/<Name>.bas` (the code-behind).

use std::path::{Path, PathBuf};

use ras_form::Severity;

/// First code line of a form written by `ras pack`; marks the form as
/// generated from `forms/` (and therefore safe to replace).
pub const MARKER: &str = "' RAS-FORM ";

/// One form found in `forms/`.
pub struct LoadedForm {
    /// Name taken from the file name (`<Name>.form.json`).
    pub name: String,
    /// Path of the `.form.json` file relative to the project root, with `/`.
    pub rel_path: String,
    pub text: String,
    pub parsed: ras_form::Parsed,
    /// Path of the code-behind file (may not exist).
    pub code_path: PathBuf,
}

pub fn forms_dir(root: &Path) -> PathBuf {
    root.join("forms")
}

/// Reads every `forms/*.form.json`, sorted by name.
pub fn load(root: &Path) -> Result<Vec<LoadedForm>, String> {
    let dir = forms_dir(root);
    let Ok(entries) = std::fs::read_dir(&dir) else {
        return Ok(Vec::new());
    };
    let mut paths: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.is_file() && p.to_string_lossy().ends_with(".form.json"))
        .collect();
    paths.sort();
    let mut forms = Vec::new();
    for path in paths {
        let file_name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        let name = file_name.trim_end_matches(".form.json").to_string();
        let text =
            std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
        let text = text
            .strip_prefix('\u{feff}')
            .map(str::to_string)
            .unwrap_or(text);
        let mut parsed = ras_form::parse(&text);
        if let Some(spec) = &parsed.spec {
            if spec.form.name != name {
                parsed.issues.push(ras_form::Issue {
                    code: "UFV003".into(),
                    severity: Severity::Error,
                    message: format!("form.name {:?} がファイル名 {file_name} と一致しません。ファイル名は <form.name>.form.json にしてください。", spec.form.name),
                    field: "form.name".into(),
                });
            }
        }
        forms.push(LoadedForm {
            code_path: dir.join(format!("{name}.bas")),
            name,
            rel_path: format!("forms/{file_name}"),
            text,
            parsed,
        });
    }
    Ok(forms)
}

fn print_issue(form: &LoadedForm, severity: &str, code: &str, message: &str, field: &str) {
    eprintln!("{severity}[{code}]: {message}");
    match ras_form::locate(&form.text, field) {
        Some((line, column)) => eprintln!("  --> {}:{line}:{column}", form.rel_path),
        None => eprintln!("  --> {}", form.rel_path),
    }
    eprintln!();
}

/// Prints the validation issues of the forms; returns the number of errors.
pub fn report(forms: &[LoadedForm]) -> usize {
    let mut errors = 0;
    for form in forms {
        for issue in &form.parsed.issues {
            let severity = match issue.severity {
                Severity::Error => {
                    errors += 1;
                    "error"
                }
                Severity::Warning => "warning",
            };
            print_issue(form, severity, &issue.code, &issue.message, &issue.field);
        }
    }
    errors
}

/// Code-behind text of a form (empty when the file does not exist).
fn read_code(form: &LoadedForm) -> Result<String, String> {
    match std::fs::read(&form.code_path) {
        Ok(bytes) => {
            let text = String::from_utf8(bytes).map_err(|_| {
                format!(
                    "{}: UTF-8 のテキストとして読めません",
                    form.code_path.display()
                )
            })?;
            let text = text
                .strip_prefix('\u{feff}')
                .unwrap_or(&text)
                .replace("\r\n", "\n");
            if text.lines().any(|l| {
                l.trim_start().starts_with("Attribute VB_")
                    || l.trim_start().starts_with("VERSION ")
            }) {
                return Err(format!(
                    "{}: コードビハインドには Attribute 行や VERSION 行を書かないでください（コードペインの内容だけを書きます）",
                    form.code_path.display()
                ));
            }
            Ok(text)
        }
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(String::new()),
        Err(e) => Err(format!("{}: {e}", form.code_path.display())),
    }
}

/// Generates the designer storage of every form. Prints problems and
/// returns `None` when there were errors.
pub fn compile(
    forms: &[LoadedForm],
    codepage: u16,
) -> Result<Option<Vec<ras_ovba::UserForm>>, String> {
    let mut out = Vec::new();
    let mut errors = 0;
    for form in forms {
        let Some(spec) = &form.parsed.spec else {
            errors += 1;
            continue;
        };
        if form.parsed.has_errors() {
            errors += 1;
            continue;
        }
        match ras_form::compile(spec, codepage) {
            Ok(designer) => {
                let code = read_code(form)?;
                let code = format!("{MARKER}{}\n{code}", form.rel_path);
                out.push(ras_ovba::UserForm {
                    name: designer.name,
                    code,
                    streams: designer.streams,
                    storages: designer.storages,
                });
            }
            Err(e) => {
                errors += 1;
                print_issue(form, "error", e.code, &e.to_string(), &e.property);
            }
        }
    }
    Ok((errors == 0).then_some(out))
}

/// Whether a module's source was written by `ras pack` from `forms/`.
pub fn is_generated(source: &str) -> bool {
    source
        .replace("\r\n", "\n")
        .lines()
        .find(|l| {
            !l.starts_with("Attribute ") && !l.starts_with("VERSION ") && !l.trim().is_empty()
        })
        .is_some_and(|l| l.starts_with(MARKER))
}

/// `ras form check`: validates and generates every form. Returns the number of errors.
pub fn check(root: &Path, codepage: u16) -> Result<usize, String> {
    let forms = load(root)?;
    let mut errors = report(&forms);
    if errors == 0 && compile(&forms, codepage)?.is_none() {
        errors += 1;
    }
    Ok(errors)
}

/// Whether `name` can name a form (a VBA identifier of at most 31 characters).
fn valid_name(name: &str) -> bool {
    let mut chars = name.chars();
    chars.next().is_some_and(char::is_alphabetic)
        && chars.all(|c| c.is_alphanumeric() || c == '_')
        && name.encode_utf16().count() <= 31
}

/// `ras form new <Name>`.
pub fn new_form(root: &Path, name: &str) -> Result<(PathBuf, PathBuf), String> {
    if !valid_name(name) {
        return Err(format!("`{name}` はフォーム名に使えません（英字などで始まり、英数字と _ からなる 31 文字以内の名前にしてください）"));
    }
    let dir = forms_dir(root);
    let spec_path = dir.join(format!("{name}.form.json"));
    let code_path = dir.join(format!("{name}.bas"));
    for path in [&spec_path, &code_path] {
        if path.exists() {
            return Err(format!("{} は既に存在します", path.display()));
        }
    }
    if let Ok(existing) = load(root) {
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
    std::fs::write(&code_path, "").map_err(|e| format!("{}: {e}", code_path.display()))?;
    Ok((spec_path, code_path))
}
