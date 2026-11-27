//! The forms of a project: `forms/<Name>.form.json` (the design) and
//! `forms/<Name>.ras` (the code-behind, written in RAS and compiled by
//! `ras-compiler` into the form's code module), RAS language specification
//! §14.11.
//!
//! * [`load`] and [`build`] turn the files into what `ras pack` and the
//!   live sync write into a workbook.
//! * [`pull`] goes the other way: it writes the files for a form found in a
//!   workbook.

use std::collections::HashMap;
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
    /// Path of the code-behind `forms/<Name>.ras` (which may not exist).
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
        let text =
            std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
        forms.push(loaded(&dir, &file_name, text));
    }
    Ok(forms)
}

/// A form of `forms/` from the text of its file `file_name` (`<Name>.form.json`), parsed and
/// validated; `dir` is the directory of the forms (where its code-behind is).
pub fn loaded(dir: &Path, file_name: &str, text: String) -> LoadedForm {
    let name = file_name.trim_end_matches(".form.json").to_string();
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
    LoadedForm {
        code_path: dir.join(format!("{name}.ras")),
        name,
        rel_path: format!("forms/{file_name}"),
        text,
        parsed,
    }
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

fn problem(
    form: &LoadedForm,
    severity: Severity,
    code: &str,
    message: &str,
    field: &str,
) -> Problem {
    Problem {
        rel_path: form.rel_path.clone(),
        position: crate::locate(&form.text, field),
        severity,
        code: code.to_string(),
        message: message.to_string(),
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
///
/// `compiled` holds the code the RAS compiler generated for the `.ras`
/// code-behinds (form name → code, the `OutputFile::code` of the form's
/// module). A `.ras` form missing from it gets no code: enough to check the
/// design (`ras check`), not to write the form.
pub fn build(
    forms: &[LoadedForm],
    root: &Path,
    codepage: u16,
    compiled: &HashMap<String, String>,
) -> Built {
    let mut built = Built {
        forms: Vec::new(),
        problems: Vec::new(),
    };
    // The bundled form of the dialogs built in code (spec §14.11.15), when the compiler made its
    // code (the project uses the dialogs).
    let bundled = bundled_dialog(forms, compiled);
    for form in forms.iter().chain(bundled.iter()) {
        for issue in &form.parsed.issues {
            built.problems.push(problem(
                form,
                issue.severity,
                &issue.code,
                &issue.message,
                &issue.field,
            ));
        }
        let Some(spec) = form
            .parsed
            .spec
            .as_ref()
            .filter(|_| !form.parsed.has_errors())
        else {
            continue;
        };
        let designer = match crate::compile_in(spec, codepage, root) {
            Ok(designer) => designer,
            Err(e) => {
                built.problems.push(problem(
                    form,
                    Severity::Error,
                    e.code,
                    &e.to_string(),
                    &e.property,
                ));
                continue;
            }
        };
        let code = compiled
            .iter()
            .find(|(name, _)| name.eq_ignore_ascii_case(&form.name))
            .map(|(_, code)| code.replace("\r\n", "\n"))
            .unwrap_or_default();
        let code = codebehind::compose(&form.rel_path, &designer, spec, &code, codepage);
        match code {
            Ok(code) => built.forms.push(UserForm {
                name: designer.name,
                code,
                streams: designer.streams,
                storages: designer.storages,
            }),
            Err(message) => built.problems.push(problem(
                form,
                Severity::Error,
                "userform_code",
                &message,
                "",
            )),
        }
    }
    built
}

/// The bundled form `RAS_Dialog` (spec §14.11.15) when `compiled` has its code and `forms` has
/// no form of that name.
fn bundled_dialog(forms: &[LoadedForm], compiled: &HashMap<String, String>) -> Option<LoadedForm> {
    let name = ras_compiler::DIALOG_FORM;
    if !compiled.keys().any(|k| k.eq_ignore_ascii_case(name))
        || forms.iter().any(|f| f.name.eq_ignore_ascii_case(name))
    {
        return None;
    }
    let mut form = loaded(
        Path::new("forms"),
        &format!("{name}.form.json"),
        ras_compiler::dialog_form_spec(),
    );
    form.rel_path = ras_compiler::DIALOG_FORM_PATH.replace(".ras", ".form.json");
    Some(form)
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
    /// VBA the conversion to RAS could not translate: (line in the `.ras` file, message).
    /// The original code is kept there as a comment.
    pub todos: Vec<(usize, String)>,
    /// The existing `.ras` code-behind was kept: the form was written from it.
    pub kept_code: bool,
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

/// Writes `forms/<Name>.form.json`, `forms/<Name>.ras` and the pictures
/// (`forms/assets/<Name>.<Control>.bmp|jpg`) for a form taken from a
/// workbook (see `ras_ovba::unpack_user_forms`).
///
/// The VBA code of the form is converted to RAS
/// (`ras_compiler::vba2ras::convert_form`); what could not be converted is
/// kept as comments and reported in [`Pulled::todos`]. A form
/// that was written from `forms/` keeps an existing `.ras` file: its code was
/// compiled from it ([`Pulled::kept_code`]).
///
/// Existing files are only replaced with `force`. What the designer
/// storage does not hold — the `list` / `selectedIndex` of ComboBox and
/// ListBox controls, the `anchor` of controls, `resizable` and the size
/// limits of the form — is kept from an existing `.form.json`. `codepage` is
/// the code page of the project the form comes from.
pub fn pull(
    root: &Path,
    form: &UserForm,
    codepage: u16,
    force: bool,
) -> Result<PullOutcome, String> {
    pull_with(root, form, codepage, force, &[], None)
}

/// [`pull`] for the sync's take-over of a workbook (spec §14.7.1): names of
/// the other converted modules (`peers`) are imported from them, and the
/// code-behind starts with `header` (lines, the import marker).
pub fn pull_with(
    root: &Path,
    form: &UserForm,
    codepage: u16,
    force: bool,
    peers: &[ras_compiler::vba2ras::Peer],
    header: Option<&str>,
) -> Result<PullOutcome, String> {
    let name = &form.name;
    if name.is_empty() || name.contains(['/', '\\', ':', '\0']) || name.starts_with('.') {
        return Err(format!("フォーム名 {name:?} はファイル名に使えません"));
    }
    let dir = forms_dir(root);
    let spec_path = dir.join(format!("{name}.form.json"));
    let code_path = dir.join(format!("{name}.ras"));
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
        previous.as_ref().and_then(|p| {
            p.controls
                .iter()
                .find(|c| c.name.eq_ignore_ascii_case(name))
        })
    };
    for control in spec
        .controls
        .iter_mut()
        .filter(|c| is_list_control(&c.control_type))
    {
        // A list that comes from cells (RowSource) has no items of its own.
        if control
            .scalar("rowSource")
            .and_then(serde_json::Value::as_str)
            .is_some_and(|s| !s.is_empty())
        {
            continue;
        }
        if let Some(old) =
            previous_control(&control.name).filter(|o| is_list_control(&o.control_type))
        {
            control.list = old.list.clone();
            control.selected_index = old.selected_index;
        }
    }

    // Nor is the size of the form at run time (spec §14.11.17).
    if let Some(old) = previous.as_ref().map(|p| &p.form) {
        for field in crate::sizing::FIELDS {
            if let Some(value) = old.scalar(field) {
                spec.form.scalars.insert(field.into(), value.clone());
            }
        }
    }

    // The anchors are not in the designer storage either (RAS_FormResize applies them).
    for control in spec.controls.iter_mut() {
        if let Some(anchor) = previous_control(&control.name)
            .and_then(|old| old.scalar("anchor"))
            .filter(|_| crate::anchor::takes_anchor(&control.control_type))
        {
            control.scalars.insert("anchor".into(), anchor.clone());
        }
    }

    let mut pictures = Vec::new();
    for found in &snapshot.pictures {
        let data = &found.picture.data;
        let icon = found.property == "mouseIcon";
        // The previous action of the same property, and the color a converted picture of it
        // was blended onto.
        let (previous, owner) = match found.control_index {
            Some(index) => {
                let old = previous_control(&spec.controls[index].name);
                let action = old.and_then(|o| {
                    if icon {
                        o.mouse_icon.clone()
                    } else {
                        o.picture.clone()
                    }
                });
                (
                    action.zip(Some(old.and_then(|o| o.back_color.clone()))),
                    spec.controls[index].name.clone(),
                )
            }
            None => {
                let old = previous.as_ref().map(|p| &p.form);
                let action = old.and_then(|o| {
                    if icon {
                        o.mouse_icon.clone()
                    } else {
                        o.picture.clone()
                    }
                });
                (
                    action.zip(Some(old.and_then(|o| o.back_color.clone()))),
                    String::new(),
                )
            }
        };
        // Keep the file the previous spec points to when the form holds
        // what that file is stored as (the file itself, or its conversion).
        let kept = previous.and_then(|(action, back_color)| {
            let path = action.get("path")?.as_str()?.to_string();
            let stored = if icon {
                crate::picture::load_icon(root, &path).ok()?
            } else {
                crate::picture::load_with(
                    root,
                    &path,
                    crate::picture::background_of(back_color.as_ref()),
                )
                .ok()?
            };
            (stored.data == *data).then_some(path)
        });
        let rel = match kept {
            Some(path) => path,
            None => {
                // forms/assets/<Form>[.<Control>][.mouseIcon].<ext>
                let mut stem = name.to_string();
                if !owner.is_empty() {
                    stem = format!("{stem}.{owner}");
                }
                if icon {
                    stem.push_str(".mouseIcon");
                }
                let rel = format!("forms/assets/{stem}.{}", found.picture.format.extension());
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
        let action = Some(json!({ "path": rel }));
        match (found.control_index, icon) {
            (Some(index), false) => spec.controls[index].picture = action,
            (Some(index), true) => spec.controls[index].mouse_icon = action,
            (None, false) => spec.form.picture = action,
            (None, true) => spec.form.mouse_icon = action,
        }
        pictures.push(rel);
    }

    // The labels of a progress bar are the progress bar again.
    crate::progress::collapse(&mut spec);
    crate::picker::collapse(&mut spec);
    crate::placeholder::collapse(&mut spec);
    let decl = ras_compiler::FormDecl {
        name: spec.form.name.clone(),
        controls: spec
            .controls
            .iter()
            .map(|c| ras_compiler::FormControl {
                name: c.name.clone(),
                control_type: c.control_type.trim().to_string(),
            })
            .collect(),
        resizable: crate::sizing::resizable(&spec.form),
    };
    // A form written from `forms/` was compiled from the `.ras` file: that file is the source
    // of its code, the VBA in the workbook only what it compiled to.
    let keep_code = ras_ovba::is_generated_form(&form.code) && code_path.exists();
    let mut converted = if keep_code {
        ras_compiler::vba2ras::Converted::default()
    } else {
        ras_compiler::vba2ras::convert(
            &codebehind::code_behind(&form.code),
            &ras_compiler::vba2ras::Target::Form(&decl),
            peers,
        )
    };
    if let Some(header) = header.filter(|_| !keep_code) {
        let shift = header.lines().count();
        converted.text = format!("{header}{}", converted.text);
        for todo in &mut converted.todos {
            todo.0 += shift;
        }
    }

    std::fs::create_dir_all(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    std::fs::write(&spec_path, spec.to_json())
        .map_err(|e| format!("{}: {e}", spec_path.display()))?;
    if !keep_code {
        std::fs::write(&code_path, &converted.text)
            .map_err(|e| format!("{}: {e}", code_path.display()))?;
    }
    Ok(PullOutcome::Written(Pulled {
        spec_path,
        code_path,
        pictures,
        warnings: snapshot.warnings,
        todos: converted.todos,
        kept_code: keep_code,
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
