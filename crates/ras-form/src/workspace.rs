//! The forms of a project: `forms/<Name>.form.json` (the design) and
//! `forms/<Name>.bas` (the code-behind), RAS language specification §14.11.
//!
//! * [`load`] and [`build`] turn the files into what `ras pack` and the
//!   live sync write into a workbook.
//! * [`pull`] goes the other way: it writes the files for a form found in a
//!   workbook.

use std::path::{Path, PathBuf};

use ras_ovba::UserForm;
use serde_json::json;

use crate::snapshot::SnapshotWarning;
use crate::{codebehind, Issue, Parsed, Severity};

/// One form found in `forms/`.
pub struct LoadedForm {
    /// Name taken from the file name (`<Name>.form.json`).
    pub name: String,
    /// Path of the `.form.json` file relative to the project root, with `/`.
    pub rel_path: String,
    /// Text of the `.form.json` file.
    pub text: String,
    /// Result of parsing and validating `text`.
    pub parsed: Parsed,
    /// Path of the code-behind file (which may not exist).
    pub code_path: PathBuf,
}

/// The directory that holds the forms of the project at `root`.
pub fn forms_dir(root: &Path) -> PathBuf {
    root.join("forms")
}

/// Reads every `forms/*.form.json`, sorted by name. A form whose file name
/// does not match its `form.name` gets an error issue.
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
        let mut parsed = crate::parse(&text);
        if let Some(spec) = &parsed.spec {
            if spec.form.name != name {
                parsed.issues.push(Issue {
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

/// Something wrong with a form, pointing into its `.form.json` file.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Problem {
    /// The file, relative to the project root.
    pub rel_path: String,
    /// Line and column (1-based) when the place is known.
    pub position: Option<(usize, usize)>,
    /// Error or warning.
    pub severity: Severity,
    /// `UFVnnn`, `userform_generation_...` or `userform_code`.
    pub code: String,
    /// Message in Japanese.
    pub message: String,
}

impl std::fmt::Display for Problem {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let severity = match self.severity {
            Severity::Error => "error",
            Severity::Warning => "warning",
        };
        writeln!(f, "{severity}[{}]: {}", self.code, self.message)?;
        match self.position {
            Some((line, column)) => write!(f, "  --> {}:{line}:{column}", self.rel_path),
            None => write!(f, "  --> {}", self.rel_path),
        }
    }
}

fn problem(form: &LoadedForm, severity: Severity, code: &str, message: &str, field: &str) -> Problem {
    Problem {
        rel_path: form.rel_path.clone(),
        position: crate::locate(&form.text, field),
        severity,
        code: code.to_string(),
        message: message.to_string(),
    }
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

/// What [`build`] made of the forms of a project.
pub struct Built {
    /// The forms that could be generated, ready to be written into a
    /// workbook (code with its marker line, designer storage).
    pub forms: Vec<UserForm>,
    /// Validation issues and generation failures of all forms.
    pub problems: Vec<Problem>,
}

impl Built {
    /// Number of problems that are errors.
    pub fn errors(&self) -> usize {
        self.problems
            .iter()
            .filter(|p| p.severity == Severity::Error)
            .count()
    }
}

/// Validates every form and generates its designer storage and code module.
/// `root` is the project root (picture paths are relative to it).
pub fn build(forms: &[LoadedForm], root: &Path, codepage: u16) -> Built {
    let mut built = Built {
        forms: Vec::new(),
        problems: Vec::new(),
    };
    for form in forms {
        for issue in &form.parsed.issues {
            built.problems.push(problem(
                form,
                issue.severity,
                &issue.code,
                &issue.message,
                &issue.field,
            ));
        }
        let Some(spec) = form.parsed.spec.as_ref().filter(|_| !form.parsed.has_errors()) else {
            continue;
        };
        let designer = match crate::compile_in(spec, codepage, root) {
            Ok(designer) => designer,
            Err(e) => {
                built
                    .problems
                    .push(problem(form, Severity::Error, e.code, &e.to_string(), &e.property));
                continue;
            }
        };
        let code = read_code(form).and_then(|code| {
            codebehind::compose(&form.rel_path, &designer, spec, &code, codepage)
        });
        match code {
            Ok(code) => built.forms.push(UserForm {
                name: designer.name,
                code,
                streams: designer.streams,
                storages: designer.storages,
            }),
            Err(message) => built
                .problems
                .push(problem(form, Severity::Error, "userform_code", &message, "")),
        }
    }
    built
}

/// The files written for one form by [`pull`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Pulled {
    /// The `.form.json` file.
    pub spec_path: PathBuf,
    /// The code-behind file.
    pub code_path: PathBuf,
    /// Picture files, relative to the project root.
    pub pictures: Vec<String>,
    /// What the FormSpec could not carry over.
    pub warnings: Vec<SnapshotWarning>,
}

/// Result of [`pull`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PullOutcome {
    /// The files were written.
    Written(Pulled),
    /// Nothing was written because these files exist already.
    Exists(Vec<PathBuf>),
}

fn is_list_control(control_type: &str) -> bool {
    let t = control_type.trim();
    t.eq_ignore_ascii_case("ComboBox") || t.eq_ignore_ascii_case("ListBox")
}

/// Writes `forms/<Name>.form.json`, `forms/<Name>.bas` and the pictures
/// (`forms/assets/<Name>.<Control>.bmp|jpg`) for a form taken from a
/// workbook (see `ras_ovba::unpack_user_forms`).
///
/// Existing files are only replaced with `force`. What the designer
/// storage does not hold — the `list` / `selectedIndex` of ComboBox and
/// ListBox controls — is kept from an existing `.form.json`. `codepage` is
/// the code page of the project the form comes from.
pub fn pull(
    root: &Path,
    form: &UserForm,
    codepage: u16,
    force: bool,
) -> Result<PullOutcome, String> {
    let name = &form.name;
    if name.is_empty() || name.contains(['/', '\\', ':', '\0']) || name.starts_with('.') {
        return Err(format!("フォーム名 {name:?} はファイル名に使えません"));
    }
    let dir = forms_dir(root);
    let spec_path = dir.join(format!("{name}.form.json"));
    let code_path = dir.join(format!("{name}.bas"));
    if !force {
        let existing: Vec<PathBuf> = [&spec_path, &code_path]
            .into_iter()
            .filter(|p| p.exists())
            .cloned()
            .collect();
        if !existing.is_empty() {
            return Ok(PullOutcome::Exists(existing));
        }
    }
    if form.streams.is_empty() {
        return Err(format!(
            "フォーム `{name}` のデザイナストレージがブックにありません"
        ));
    }
    let snapshot = crate::snapshot(name, &form.streams, &form.storages, codepage)
        .map_err(|e| format!("フォーム `{name}`: {e}"))?;
    let mut spec = snapshot.spec;
    spec.warnings = None;

    // What only the previous `.form.json` knows.
    let previous = std::fs::read_to_string(&spec_path)
        .ok()
        .and_then(|text| crate::parse(&text).spec);
    let previous_control = |name: &str| {
        previous
            .as_ref()
            .and_then(|p| p.controls.iter().find(|c| c.name.eq_ignore_ascii_case(name)))
    };
    for control in spec.controls.iter_mut().filter(|c| is_list_control(&c.control_type)) {
        if let Some(old) = previous_control(&control.name).filter(|o| is_list_control(&o.control_type)) {
            control.list = old.list.clone();
            control.selected_index = old.selected_index;
        }
    }

    let mut pictures = Vec::new();
    for found in &snapshot.pictures {
        let control = &mut spec.controls[found.control_index];
        let data = &found.picture.data;
        // Keep the file the previous spec points to when it is this picture.
        let kept = previous_control(&control.name)
            .and_then(|old| old.picture.as_ref()?.get("path")?.as_str().map(str::to_string))
            .filter(|path| {
                crate::picture::resolve(root, path)
                    .ok()
                    .and_then(|file| std::fs::read(file).ok())
                    .is_some_and(|existing| existing == *data)
            });
        let rel = match kept {
            Some(path) => path,
            None => {
                let rel = format!(
                    "forms/assets/{name}.{}.{}",
                    control.name,
                    found.picture.format.extension()
                );
                let file = root.join(&rel);
                let same = std::fs::read(&file).is_ok_and(|existing| existing == *data);
                if file.exists() && !same && !force {
                    return Err(format!(
                        "{} は既に存在し、内容がブックの画像と異なります（置き換えるには --force を付けてください）",
                        file.display()
                    ));
                }
                if !same {
                    if let Some(parent) = file.parent() {
                        std::fs::create_dir_all(parent)
                            .map_err(|e| format!("{}: {e}", parent.display()))?;
                    }
                    std::fs::write(&file, data).map_err(|e| format!("{}: {e}", file.display()))?;
                }
                rel
            }
        };
        control.picture = Some(json!({ "path": rel }));
        pictures.push(rel);
    }

    std::fs::create_dir_all(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    std::fs::write(&spec_path, spec.to_json())
        .map_err(|e| format!("{}: {e}", spec_path.display()))?;
    std::fs::write(&code_path, codebehind::code_behind(&form.code))
        .map_err(|e| format!("{}: {e}", code_path.display()))?;
    Ok(PullOutcome::Written(Pulled {
        spec_path,
        code_path,
        pictures,
        warnings: snapshot.warnings,
    }))
}

/// `form.name`-style check used for new forms: a VBA identifier of at most
/// 31 characters.
pub fn valid_form_name(name: &str) -> bool {
    let mut chars = name.chars();
    chars.next().is_some_and(char::is_alphabetic)
        && chars.all(|c| c.is_alphanumeric() || c == '_')
        && name.encode_utf16().count() <= 31
}
