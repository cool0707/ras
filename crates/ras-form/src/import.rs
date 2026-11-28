//! Taking over the hand-written VBA of a workbook when it is synced (RAS
//! language specification §14.7.1).
//!
//! [`prepare`] looks at the components of the workbook before the project is
//! built. What RAS does not manage yet and the project has no file for — a
//! standard module, the hand-written code of a sheet or of `ThisWorkbook`, a
//! user form made in the VBE — is converted to RAS and written into the
//! project:
//!
//! * a standard module `Module1` → `src/Module1.ras`,
//! * the hand-written code of the sheet `Sheet1` → `objects/Sheet1.ras`, of the
//!   workbook → `objects/ThisWorkbook.ras` (spec §12.2: the file is the whole
//!   module; event procedures become `@on("Worksheet.Change")` handlers),
//! * a form `frmInput` → `forms/frmInput.form.json` and `forms/frmInput.ras`
//!   (as `ras form pull` does).
//!
//! The original VBA is kept in `vba/`. Each converted RAS file starts with
//! the marker line `// RAS-IMPORTED <name>`: as long as it is there, the sync
//! may replace the hand-written component `<name>` of the workbook with the
//! code generated from the file ([`Prepared::adopt`]) — unless the file
//! still has lines the conversion left as `// TODO(VBA):`; the workbook then
//! keeps its component and the sync leaves it alone ([`Prepared::pending`])
//! until they are fixed. [`finish`] turns the marker into a plain comment once
//! the workbook has the generated code.

use std::path::{Path, PathBuf};

use ras_compiler::collect::vba_module_name;
use ras_compiler::vba2ras::{self, Peer, Target};
use ras_ovba::{ModuleKind, UserForm};

use crate::workspace::{self, PullOutcome};

/// Start of the first line of a RAS file converted from the workbook.
pub const IMPORT_MARKER: &str = "// RAS-IMPORTED ";

/// What a VBA conversion could not translate, kept as a comment starting with this.
const TODO_MARKER: &str = "// TODO(VBA):";

/// A component of the workbook's VBA project.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BookModule {
    /// Its name (for a document module, the code name of the sheet or workbook).
    pub name: String,
    /// Standard, class, form or document module.
    pub kind: ModuleKind,
    /// Its code (with or without the `Attribute` header).
    pub code: String,
}

/// The workbook a sync is about to write to.
pub struct Book<'a> {
    /// Its components.
    pub modules: &'a [BookModule],
    /// The code name of the workbook (`ThisWorkbook`); the other document modules are sheets.
    pub workbook_code_name: &'a str,
    /// The ActiveX controls on the sheets, by the sheet's code name (spec §12.2.1).
    pub sheet_controls: &'a [ras_compiler::FormDecl],
    /// The designer of a form made in the VBE, by name (exported from the open workbook,
    /// or read from the workbook file).
    pub form: &'a mut dyn FnMut(&str) -> Result<UserForm, String>,
    /// The code RAS wrote into the module of this name the last time (spec §14.7.3), if known.
    pub written: &'a mut dyn FnMut(&str) -> Option<String>,
}

/// Code added in the VBE to a module RAS generated, taken over into the module's file
/// (spec §14.7.3).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Appended {
    /// Name of the module in the workbook.
    pub name: String,
    /// The RAS file the code was appended to, relative to the project root.
    pub path: String,
    /// What was added: the names of the procedures, and how many declarations.
    pub names: Vec<String>,
    /// The file as it was before, to put it back when the result does not compile.
    original: Option<String>,
}

/// What [`prepare`] takes over from a component of the workbook.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ImportKind {
    /// A standard module.
    Standard,
    /// The hand-written code of a sheet, or of the workbook (`ThisWorkbook`).
    Document {
        /// The workbook's.
        workbook: bool,
    },
    /// A user form.
    Form,
    /// A class module: a `class` of `src/<Name>.ras`.
    Class,
}

/// A component converted and written into the project by [`prepare`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Imported {
    /// Name of the component in the workbook.
    pub name: String,
    /// What it is.
    pub kind: ImportKind,
    /// The RAS file written (the code-behind for a form), relative to the project root.
    pub path: String,
    /// The other files written (the form's design and pictures, the copy in `vba/`).
    pub other_files: Vec<String>,
    /// Lines the conversion left as `// TODO(VBA):`: (1-based line in `path`, message).
    pub todos: Vec<(usize, String)>,
    /// What the form's design could not carry over.
    pub warnings: Vec<String>,
    /// The procedures kept as VBA (`` vba`…` `` blocks, spec §14.7.4).
    pub kept: Vec<String>,
}

/// What [`keep_procedures`] needs to convert a module again.
#[derive(Clone, Debug, PartialEq, Eq)]
struct Reconvert {
    /// The RAS file, relative to the project root.
    path: String,
    name: String,
    kind: ImportKind,
    code: String,
    controls: Vec<ras_compiler::FormControl>,
    /// The form the code is the code-behind of.
    form: Option<ras_compiler::FormDecl>,
    /// The procedures of the VBA to keep as they are.
    keep: Vec<String>,
    /// The procedures converted: (first line, last line of the file, 1-based; name in the VBA).
    procs: Vec<(usize, usize, String)>,
}

/// A component of the workbook that has a converted file with `// TODO(VBA):` lines left:
/// the sync leaves it as it is and does not write what the file generates.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Pending {
    /// Name of the component in the workbook.
    pub name: String,
    /// Name of the module generated from the file (`ModMain` for `mod_Main`); the same as
    /// `name` for a form, a sheet and the workbook.
    pub generated: String,
    /// What it is.
    pub kind: ImportKind,
    /// The converted file, relative to the project root.
    pub path: String,
    /// Number of `// TODO(VBA):` lines in it.
    pub todos: usize,
}

/// Result of [`prepare`].
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Prepared {
    /// Components converted and written into the project now.
    pub imported: Vec<Imported>,
    /// Hand-written components of the workbook the sync may replace (or, for a standard
    /// module whose RAS name differs, remove), with what it is.
    pub adopt: Vec<(String, ImportKind)>,
    /// Components the sync must leave alone.
    pub pending: Vec<Pending>,
    /// Components that were not taken over, with the reason.
    pub skipped: Vec<String>,
    /// Files [`sheet_files`] created for sheets (and `ThisWorkbook`) that had none, relative
    /// to the project root, with what they are for.
    pub created: Vec<(String, String)>,
    /// Code added in the VBE that was appended to the file of its module (spec §14.7.3).
    pub appended: Vec<Appended>,
    /// Files whose marker line becomes a plain comment once the workbook has the generated code.
    marked: Vec<PathBuf>,
    /// The modules converted now, and the names they make public to each other.
    reconvert: Vec<Reconvert>,
    peers: Vec<Peer>,
}

impl Prepared {
    /// Whether the sync may replace the hand-written component `name`.
    pub fn adopts(&self, name: &str) -> bool {
        self.adopt.iter().any(|(n, _)| n.eq_ignore_ascii_case(name))
    }

    /// Lines that tell the user what was taken over and what waits for them.
    pub fn messages(&self) -> Vec<String> {
        let mut out = Vec::new();
        for item in &self.imported {
            let what = match item.kind {
                ImportKind::Standard => "標準モジュール",
                ImportKind::Document { workbook: false } => "シートモジュール",
                ImportKind::Document { workbook: true } => "ブックのモジュール",
                ImportKind::Form => "ユーザーフォーム",
                ImportKind::Class => "クラスモジュール",
            };
            out.push(format!(
                "ブックの{what} `{}` を {} に取り込みました",
                item.name, item.path
            ));
            if !item.kept.is_empty() {
                out.push(format!(
                    "  RAS に変換できず、VBA のまま残したプロシージャ（`vba` ブロック）: {}",
                    item.kept.join(", ")
                ));
            }
            for warning in &item.warnings {
                out.push(format!("  {warning}"));
            }
            for (line, message) in &item.todos {
                out.push(format!(
                    "  {}:{line}: VBA から RAS に変換できませんでした: {message}",
                    item.path
                ));
            }
        }
        for p in &self.pending {
            out.push(format!(
                "{} には変換できなかった行（`// TODO(VBA):`）が {} 行残っているため、ブックの `{}` は置き換えずにそのまま残しています。直して `// TODO(VBA):` の行を消すと、次の同期で置き換えます",
                p.path, p.todos, p.name
            ));
        }
        for a in &self.appended {
            out.push(format!(
                "ブックの `{}` に VBE で足されたコード（{}）を {} に取り込みました",
                a.name,
                a.names.join("、"),
                a.path
            ));
        }
        out.extend(self.skipped.iter().cloned());
        for (path, what) in &self.created {
            out.push(format!(
                "{what}のモジュールのファイル {path} を作成しました"
            ));
        }
        out
    }
}

/// The first line of a converted file: the marker, and what it means.
fn header(name: &str) -> String {
    format!(
        "{IMPORT_MARKER}{name} — ブックの VBA から変換したファイルです。この行がある間は、同期のときにブックの手書きの `{name}` をこのファイルから生成したコードで置き換えます（`// TODO(VBA):` の行が残っている間は置き換えません）。\n\n"
    )
}

/// The name in the marker line of a RAS file, if it has one.
fn marker_of(text: &str) -> Option<String> {
    let first = text
        .strip_prefix('\u{feff}')
        .unwrap_or(text)
        .lines()
        .next()?;
    let rest = first.strip_prefix(IMPORT_MARKER)?;
    rest.split_whitespace().next().map(str::to_string)
}

fn todo_count(text: &str) -> usize {
    text.lines()
        .filter(|l| l.trim_start().starts_with(TODO_MARKER))
        .count()
}

/// Whether code holds anything besides `Attribute` lines, `Option` statements, comments and blank lines.
fn has_code(code: &str) -> bool {
    code.lines().map(str::trim).any(|l| {
        !l.is_empty()
            && !l.starts_with('\'')
            && !l.starts_with("Attribute ")
            && !l.starts_with("VERSION ")
            && !l.to_ascii_lowercase().starts_with("option ")
            && !l.to_ascii_lowercase().starts_with("rem ")
    })
}

/// The code pane text: without the `Attribute` header, with LF line ends.
fn code_pane(code: &str) -> String {
    code.replace("\r\n", "\n")
        .lines()
        .skip_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION "))
        .map(|l| format!("{l}\n"))
        .collect()
}

/// The hand-written code of a document module: its code pane without the RAS-managed region
/// that RAS wrote into document modules before a file was the whole module.
pub fn hand_written(code: &str) -> String {
    code_pane(&ras_ovba::merge_managed_region(
        &code.replace("\r\n", "\n"),
        "",
    ))
}

/// Whether a RAS source holds anything besides comments, blank lines and imports: a file
/// without code generates nothing (spec §12.1), as the files [`sheet_files`] creates.
pub fn ras_has_code(text: &str) -> bool {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let mut rest = text;
    let mut code = String::new();
    // Comments go; what a string holds does not matter here, a string being code itself.
    while !rest.is_empty() {
        if let Some(r) = rest.strip_prefix("//") {
            rest = r.find('\n').map(|i| &r[i..]).unwrap_or("");
        } else if let Some(r) = rest.strip_prefix("/*") {
            rest = r.find("*/").map(|i| &r[i + 2..]).unwrap_or("");
        } else {
            let mut chars = rest.chars();
            let c = chars.next().unwrap_or(' ');
            if c == '"' || c == '\'' || c == '`' {
                // The module name of an import; any other string is code.
                let statement = code
                    .rsplit([';', '\n'])
                    .next()
                    .unwrap_or_default()
                    .trim_start();
                if !statement.starts_with("import") {
                    return true;
                }
                rest = rest.find('\n').map(|i| &rest[i..]).unwrap_or("");
                continue;
            }
            code.push(c);
            rest = chars.as_str();
        }
    }
    code.split([';', '\n'])
        .map(str::trim)
        .any(|stmt| !stmt.is_empty() && !stmt.starts_with("import"))
}

/// The file of a sheet (by its code name) or of the workbook (spec §12.2).
fn object_path(root: &Path, code_name: &str, workbook: bool) -> PathBuf {
    ras_compiler::object_code_path(root, if workbook { "ThisWorkbook" } else { code_name })
}

/// A project module found on disk.
struct ProjectFile {
    /// Its VBA name.
    vba_name: String,
    path: PathBuf,
    /// The workbook component its marker names.
    marker: Option<String>,
    todos: usize,
}

fn relative(root: &Path, path: &Path) -> String {
    path.strip_prefix(root)
        .unwrap_or(path)
        .to_string_lossy()
        .replace('\\', "/")
}

fn walk(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for e in entries.flatten() {
        let path = e.path();
        if path.is_dir() {
            walk(&path, out);
        } else if path.extension().is_some_and(|x| x == "ras")
            && !path.to_string_lossy().ends_with(".d.ras")
        {
            out.push(path);
        }
    }
}

/// The modules of `src/`, the files of `objects/` and the code-behinds of `forms/`.
fn project_files(root: &Path) -> Vec<ProjectFile> {
    let mut paths = Vec::new();
    walk(&root.join("src"), &mut paths);
    paths.sort();
    let mut out = Vec::new();
    for path in paths {
        let key = relative(&root.join("src"), &path);
        let key = key.strip_suffix(".ras").unwrap_or(&key).to_string();
        let text = std::fs::read_to_string(&path).unwrap_or_default();
        out.push(ProjectFile {
            vba_name: vba_module_name(&key),
            marker: marker_of(&text),
            todos: todo_count(&text),
            path,
        });
    }
    if let Ok(forms) = workspace::load(root) {
        for form in forms {
            let text = std::fs::read_to_string(&form.code_path).unwrap_or_default();
            out.push(ProjectFile {
                vba_name: form.name.clone(),
                marker: marker_of(&text),
                todos: todo_count(&text),
                path: form.code_path,
            });
        }
    }
    for path in object_files(root) {
        let text = std::fs::read_to_string(&path).unwrap_or_default();
        let name = path
            .file_stem()
            .map(|s| s.to_string_lossy().to_string())
            .unwrap_or_default();
        out.push(ProjectFile {
            vba_name: name,
            marker: marker_of(&text),
            todos: todo_count(&text),
            path,
        });
    }
    out
}

/// The files of `objects/` (spec §12.2), sorted.
fn object_files(root: &Path) -> Vec<PathBuf> {
    let Ok(entries) = std::fs::read_dir(root.join(ras_compiler::OBJECTS_DIR)) else {
        return Vec::new();
    };
    let mut paths: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| {
            p.is_file()
                && p.extension().is_some_and(|x| x == "ras")
                && !p.to_string_lossy().ends_with(".d.ras")
        })
        .collect();
    paths.sort();
    paths
}

/// A component to convert.
struct Candidate<'b> {
    module: &'b BookModule,
    kind: ImportKind,
    /// Its code (for a document module, the hand-written part).
    code: String,
    /// Key of the RAS module in `src/` (not used for forms).
    key: String,
}

/// Converts what the sync takes over from the workbook and writes it into the project at
/// `root` (see the module documentation). `codepage` is the project's `target.codepage`.
pub fn prepare(root: &Path, book: &mut Book, codepage: u16) -> Result<Prepared, String> {
    let mut prepared = Prepared::default();
    let files = project_files(root);
    let taken = |name: &str| {
        files.iter().any(|f| f.vba_name.eq_ignore_ascii_case(name))
            || book
                .modules
                .iter()
                .any(|m| m.name.eq_ignore_ascii_case(name))
    };
    let marked = |name: &str| {
        files.iter().find(|f| {
            f.marker
                .as_deref()
                .is_some_and(|m| m.eq_ignore_ascii_case(name))
        })
    };

    let mut candidates: Vec<Candidate> = Vec::new();
    for module in book.modules {
        let (kind, code) = match module.kind {
            ModuleKind::Standard => {
                // RAS-generated (also the temporary module of the sync), or nothing in it.
                if ras_compiler::emit::has_drifted(&module.code).is_some()
                    || module.name.starts_with("RAS_")
                    || !has_code(&module.code)
                {
                    continue;
                }
                (ImportKind::Standard, code_pane(&module.code))
            }
            ModuleKind::Document => {
                // RAS-generated (the file is the whole module, spec §12.2), or nothing in it.
                let code = hand_written(&module.code);
                if ras_compiler::emit::has_drifted(&code).is_some() || !has_code(&code) {
                    continue;
                }
                (
                    ImportKind::Document {
                        workbook: module.name.eq_ignore_ascii_case(book.workbook_code_name),
                    },
                    code,
                )
            }
            ModuleKind::Form => {
                if ras_ovba::is_generated_form(&module.code) {
                    continue;
                }
                (ImportKind::Form, code_pane(&module.code))
            }
            ModuleKind::Class => {
                // RAS-generated (a `class` of the project), or nothing in it.
                if ras_compiler::emit::has_drifted(&module.code).is_some()
                    || module.name.starts_with("RAS_")
                    || !has_code(&module.code)
                {
                    continue;
                }
                (ImportKind::Class, code_pane(&module.code))
            }
            #[allow(unreachable_patterns)]
            _ => continue,
        };
        // Converted earlier: the file decides.
        if let Some(file) = marked(&module.name) {
            let generated = file.vba_name.clone();
            if file.todos > 0 {
                prepared.pending.push(Pending {
                    name: module.name.clone(),
                    generated,
                    kind,
                    path: relative(root, &file.path),
                    todos: file.todos,
                });
            } else {
                prepared.adopt.push((module.name.clone(), kind));
                prepared.marked.push(file.path.clone());
            }
            continue;
        }
        match kind {
            ImportKind::Class => {
                // A file or a module of that name is the project's.
                let key = module.name.clone();
                if files.iter().any(|f| {
                    f.vba_name.eq_ignore_ascii_case(&module.name)
                        || relative(&root.join("src"), &f.path)
                            .eq_ignore_ascii_case(&format!("{key}.ras"))
                }) {
                    continue;
                }
                candidates.push(Candidate {
                    module,
                    kind,
                    code,
                    key,
                });
            }
            ImportKind::Standard => {
                // A module of the same name (or file) is the project's: the usual rules of the sync apply.
                let key = module.name.clone();
                if files.iter().any(|f| {
                    f.vba_name.eq_ignore_ascii_case(&module.name)
                        || relative(&root.join("src"), &f.path)
                            .eq_ignore_ascii_case(&format!("{key}.ras"))
                }) {
                    continue;
                }
                let ras_name = vba_module_name(&key);
                if !ras_name.eq_ignore_ascii_case(&module.name) && taken(&ras_name) {
                    prepared.skipped.push(format!("ブックの標準モジュール `{}` は、RAS のモジュール名 `{ras_name}` が既に使われているため取り込みませんでした", module.name));
                    continue;
                }
                candidates.push(Candidate {
                    module,
                    kind,
                    code,
                    key,
                });
            }
            ImportKind::Document { workbook } => {
                // The file of the sheet. One with code is the project's: the usual rules of
                // the sync apply (the hand-written code is not replaced silently). One without
                // code — as the sync creates them (spec §14.7.2) — is replaced.
                let path = object_path(root, &module.name, workbook);
                if std::fs::read_to_string(&path).is_ok_and(|text| ras_has_code(&text)) {
                    continue;
                }
                let key = if workbook {
                    "ThisWorkbook".to_string()
                } else {
                    module.name.clone()
                };
                candidates.push(Candidate {
                    module,
                    kind,
                    code,
                    key,
                });
            }
            ImportKind::Form => {
                let dir = workspace::forms_dir(root);
                let exists = ["form.json", "ras"]
                    .iter()
                    .any(|ext| dir.join(format!("{}.{ext}", module.name)).exists())
                    || files
                        .iter()
                        .any(|f| f.vba_name.eq_ignore_ascii_case(&module.name));
                if exists {
                    continue;
                }
                candidates.push(Candidate {
                    module,
                    kind,
                    code,
                    key: String::new(),
                });
            }
        }
    }

    let sheet_controls = book.sheet_controls;
    let controls_of = |code_name: &str| {
        sheet_controls
            .iter()
            .find(|d| d.name.to_lowercase() == code_name.to_lowercase())
            .map(|d| &d.controls[..])
            .unwrap_or(&[])
    };
    // Names the converted modules make public to each other.
    let peers: Vec<Peer> = candidates
        .iter()
        .map(|c| {
            let target = match c.kind {
                ImportKind::Standard => Target::Standard,
                ImportKind::Document { workbook } => Target::Document {
                    code_name: &c.module.name,
                    workbook,
                    controls: controls_of(&c.module.name),
                },
                // (The public procedures of a form: its event procedures are private.)
                ImportKind::Form => Target::Standard,
                ImportKind::Class => Target::Class {
                    name: &c.module.name,
                },
            };
            Peer {
                vba_name: c.module.name.clone(),
                key: c.key.clone(),
                document: matches!(c.kind, ImportKind::Document { .. }),
                names: vba2ras::public_names(&c.code, &target),
                form: c.kind == ImportKind::Form,
            }
        })
        .collect();

    let vba_dir = root.join("vba");
    for c in &candidates {
        let name = &c.module.name;
        // The original, as `ras unpack` writes it.
        let ext = match c.kind {
            ImportKind::Standard => "bas",
            ImportKind::Document { .. } | ImportKind::Class => "cls",
            ImportKind::Form => "frm",
        };
        std::fs::create_dir_all(&vba_dir).map_err(|e| format!("{}: {e}", vba_dir.display()))?;
        let copy = vba_dir.join(format!("{name}.{ext}"));
        std::fs::write(&copy, &c.code).map_err(|e| format!("{}: {e}", copy.display()))?;
        let copy = relative(root, &copy);

        let header = header(name);
        let imported = match c.kind {
            ImportKind::Form => {
                let mut form = match (book.form)(name) {
                    Ok(form) => form,
                    Err(e) => {
                        prepared.skipped.push(format!(
                            "ブックのユーザーフォーム `{name}` を取り込めませんでした: {e}"
                        ));
                        continue;
                    }
                };
                form.code = c.code.clone();
                let pulled =
                    match workspace::pull_with(root, &form, codepage, false, &peers, Some(&header))
                    {
                        Ok(PullOutcome::Written(pulled)) => pulled,
                        Ok(PullOutcome::Exists(_)) => continue,
                        Err(e) => {
                            prepared.skipped.push(format!(
                                "ブックのユーザーフォーム `{name}` を取り込めませんでした: {e}"
                            ));
                            continue;
                        }
                    };
                if !pulled.kept_code {
                    prepared.reconvert.push(Reconvert {
                        path: relative(root, &pulled.code_path),
                        name: name.clone(),
                        kind: c.kind,
                        code: pulled.code.clone(),
                        controls: Vec::new(),
                        form: Some(pulled.decl.clone()),
                        keep: Vec::new(),
                        procs: pulled.procs.clone(),
                    });
                }
                let mut other_files = vec![relative(root, &pulled.spec_path)];
                other_files.extend(pulled.pictures.iter().cloned());
                other_files.push(copy);
                let warnings = pulled
                    .warnings
                    .iter()
                    .map(|w| {
                        if w.control.is_empty() {
                            w.message.clone()
                        } else {
                            format!("{}: {}", w.control, w.message)
                        }
                    })
                    .collect();
                Imported {
                    name: name.clone(),
                    kind: c.kind,
                    path: relative(root, &pulled.code_path),
                    other_files,
                    todos: pulled.todos,
                    warnings,
                    kept: pulled.raw,
                }
            }
            ImportKind::Standard | ImportKind::Document { .. } | ImportKind::Class => {
                let target = match c.kind {
                    ImportKind::Document { workbook } => Target::Document {
                        code_name: name,
                        workbook,
                        controls: controls_of(name),
                    },
                    ImportKind::Class => Target::Class { name },
                    _ => Target::Standard,
                };
                let converted = vba2ras::convert_keeping(&c.code, &target, &peers, &[]);
                let shift = header.lines().count();
                let path = match c.kind {
                    ImportKind::Document { workbook } => object_path(root, name, workbook),
                    _ => root.join("src").join(format!("{}.ras", c.key)),
                };
                if let Some(parent) = path.parent() {
                    std::fs::create_dir_all(parent)
                        .map_err(|e| format!("{}: {e}", parent.display()))?;
                }
                std::fs::write(&path, format!("{header}{}", converted.text))
                    .map_err(|e| format!("{}: {e}", path.display()))?;
                let todos = converted
                    .todos
                    .into_iter()
                    .map(|(line, message)| (line + shift, message))
                    .collect();
                prepared.reconvert.push(Reconvert {
                    path: relative(root, &path),
                    name: name.clone(),
                    kind: c.kind,
                    code: c.code.clone(),
                    controls: controls_of(name).to_vec(),
                    form: None,
                    keep: Vec::new(),
                    procs: converted
                        .procs
                        .iter()
                        .map(|(from, to, name)| (from + shift, to + shift, name.clone()))
                        .collect(),
                });
                Imported {
                    name: name.clone(),
                    kind: c.kind,
                    path: relative(root, &path),
                    other_files: vec![copy],
                    todos,
                    warnings: Vec::new(),
                    kept: converted.raw,
                }
            }
        };
        let generated = match c.kind {
            ImportKind::Form | ImportKind::Document { .. } | ImportKind::Class => name.clone(),
            ImportKind::Standard => vba_module_name(&c.key),
        };
        if imported.todos.is_empty() {
            prepared.adopt.push((name.clone(), c.kind));
            prepared.marked.push(root.join(&imported.path));
        } else {
            prepared.pending.push(Pending {
                name: name.clone(),
                generated,
                kind: c.kind,
                path: imported.path.clone(),
                todos: imported.todos.len(),
            });
        }
        prepared.imported.push(imported);
    }
    prepared.peers = peers.clone();

    // What was added in the VBE to the modules RAS generated (spec §14.7.3).
    for module in book.modules {
        let (kind, pane) = match module.kind {
            ModuleKind::Standard if !module.name.starts_with("RAS_") => {
                (ImportKind::Standard, code_pane(&module.code))
            }
            ModuleKind::Document => (
                ImportKind::Document {
                    workbook: module.name.eq_ignore_ascii_case(book.workbook_code_name),
                },
                hand_written(&module.code),
            ),
            _ => continue,
        };
        if ras_compiler::emit::has_drifted(&pane) != Some(true) {
            continue;
        }
        // The file the module is generated from. (A sheet without a file: a copy, `sheet_files`.)
        let path = match kind {
            ImportKind::Document { workbook } => object_path(root, &module.name, workbook),
            _ => match files.iter().find(|f| {
                f.vba_name.eq_ignore_ascii_case(&module.name)
                    && f.path.starts_with(root.join("src"))
            }) {
                Some(f) => f.path.clone(),
                None => continue,
            },
        };
        if !path.is_file() {
            continue;
        }
        let Some(written) = (book.written)(&module.name) else {
            continue;
        };
        let diff = crate::vbdiff::diff(&written, &pane);
        if !diff.has_additions() {
            continue;
        }
        let target = match kind {
            ImportKind::Document { workbook } => Target::Document {
                code_name: &module.name,
                workbook,
                controls: controls_of(&module.name),
            },
            _ => Target::Standard,
        };
        let todos = append_added(
            root,
            &path,
            &module.name,
            &diff,
            &target,
            matches!(kind, ImportKind::Document { .. }),
            &mut prepared,
        )?;
        if todos > 0 {
            prepared.pending.push(Pending {
                name: module.name.clone(),
                generated: module.name.clone(),
                kind,
                path: relative(root, &path),
                todos,
            });
        } else if !diff.has_drift() {
            // Nothing RAS wrote was changed: the module is written again from its file.
            prepared.adopt.push((module.name.clone(), kind));
        }
    }
    Ok(prepared)
}

/// Start of the comment line above code that was added in the VBE and taken over.
const ADDED_MARKER: &str = "// VBE で追加されたコードを取り込みました";

/// Appends the code a [`Diff`](crate::vbdiff::Diff) found added in the VBE to the RAS file
/// `path` of its module `name`, converted to RAS, unless the file has it already (the comment
/// line above it carries a hash of the added code). The added VBA is kept in `vba/`. Returns
/// the number of `// TODO(VBA):` lines the file has afterwards.
fn append_added(
    root: &Path,
    path: &Path,
    name: &str,
    diff: &crate::vbdiff::Diff,
    target: &Target,
    document: bool,
    prepared: &mut Prepared,
) -> Result<usize, String> {
    let text = std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
    let marker = format!("{ADDED_MARKER}（{}）", diff.added_hash());
    if text.contains(&marker) {
        return Ok(todo_count(&text));
    }
    let added = diff.added_code();
    let vba_dir = root.join("vba");
    std::fs::create_dir_all(&vba_dir).map_err(|e| format!("{}: {e}", vba_dir.display()))?;
    let copy = vba_dir.join(format!("{name}.{}", if document { "cls" } else { "bas" }));
    std::fs::write(&copy, &added).map_err(|e| format!("{}: {e}", copy.display()))?;
    let converted = vba2ras::convert(&added, target, &[]);
    let mut new_text = text.clone();
    if !new_text.is_empty() && !new_text.ends_with('\n') {
        new_text.push('\n');
    }
    if !new_text.is_empty() {
        new_text.push('\n');
    }
    new_text.push_str(&format!(
        "{marker}: {}\n{}",
        diff.added_names().join("、"),
        converted.text
    ));
    std::fs::write(path, &new_text).map_err(|e| format!("{}: {e}", path.display()))?;
    prepared.appended.push(Appended {
        name: name.to_string(),
        path: relative(root, path),
        names: diff.added_names(),
        original: Some(text),
    });
    Ok(todo_count(&new_text))
}

/// Takes back what [`prepare`] and [`sheet_files`] appended to the files `failed` (paths
/// relative to `root`): the result did not compile (a name that exists already, …). The
/// module stays as it is in the workbook, as drift (spec §14.7.3).
pub fn revert_appended(
    root: &Path,
    prepared: &mut Prepared,
    failed: &[String],
) -> Result<(), String> {
    let mut kept = Vec::new();
    for a in std::mem::take(&mut prepared.appended) {
        if !failed.iter().any(|f| f.eq_ignore_ascii_case(&a.path)) {
            kept.push(a);
            continue;
        }
        let path = root.join(&a.path);
        match &a.original {
            Some(original) => {
                std::fs::write(&path, original).map_err(|e| format!("{}: {e}", path.display()))?
            }
            // A file made for a copied sheet: it goes.
            None => {
                let _ = std::fs::remove_file(&path);
                prepared
                    .created
                    .retain(|(p, _)| !p.eq_ignore_ascii_case(&a.path));
            }
        }
        prepared
            .adopt
            .retain(|(n, _)| !n.eq_ignore_ascii_case(&a.name));
        prepared
            .pending
            .retain(|p| !p.name.eq_ignore_ascii_case(&a.name));
        prepared.skipped.push(format!("ブックの `{}` に VBE で足されたコード（{}）は、RAS に変換するとコンパイル エラーになるため取り込みませんでした（ドリフトとして残ります）", a.name, a.names.join("、")));
    }
    prepared.appended = kept;
    Ok(())
}

/// After a sync that replaced the components [`Prepared::adopt`] names: turns the marker
/// line of their files into a plain comment, so that they are ordinary RAS modules. The
/// line stays (as a comment), so the line numbers — and the generated code — do not change.
pub fn finish(prepared: &Prepared) -> Result<(), String> {
    for path in &prepared.marked {
        let Ok(text) = std::fs::read_to_string(path) else {
            continue;
        };
        let body = text.strip_prefix('\u{feff}').unwrap_or(&text);
        let Some(first) = body.lines().next() else {
            continue;
        };
        let Some(name) = marker_of(first) else {
            continue;
        };
        let bom = if text.starts_with('\u{feff}') {
            "\u{feff}"
        } else {
            ""
        };
        let comment = format!("// ブックの VBA（{name}）から変換したファイルです。元の VBA は vba/ に保存しています。");
        std::fs::write(path, format!("{bom}{comment}{}", &body[first.len()..]))
            .map_err(|e| format!("{}: {e}", path.display()))?;
    }
    Ok(())
}

/// The procedure of the VBA that the line `line` (1-based) of the converted file `path`
/// (relative to the project root) was converted from, if it is in one.
pub fn procedure_at(prepared: &Prepared, path: &str, line: usize) -> Option<String> {
    let module = prepared
        .reconvert
        .iter()
        .find(|r| r.path.eq_ignore_ascii_case(path))?;
    module
        .procs
        .iter()
        .find(|(from, to, _)| (*from..=*to).contains(&line))
        .map(|(_, _, name)| name.clone())
}

/// Converts the module of the file `path` again with the procedures `names` kept as VBA
/// (`` vba`…` `` blocks, spec §14.7.4): what the take-over does when their conversion does not
/// compile. Returns whether the file changed (a procedure that cannot be kept — of a class, a
/// `Property` — changes nothing).
pub fn keep_procedures(
    root: &Path,
    prepared: &mut Prepared,
    path: &str,
    names: &[String],
) -> Result<bool, String> {
    let peers = prepared.peers.clone();
    let Some(module) = prepared
        .reconvert
        .iter_mut()
        .find(|r| r.path.eq_ignore_ascii_case(path))
    else {
        return Ok(false);
    };
    for name in names {
        if !module.keep.iter().any(|k| k.eq_ignore_ascii_case(name)) {
            module.keep.push(name.clone());
        }
    }
    let target = match module.kind {
        ImportKind::Document { workbook } => Target::Document {
            code_name: &module.name,
            workbook,
            controls: &module.controls,
        },
        ImportKind::Class => Target::Class { name: &module.name },
        ImportKind::Form => match &module.form {
            Some(form) => Target::Form(form),
            None => return Ok(false),
        },
        _ => Target::Standard,
    };
    let converted = vba2ras::convert_keeping(&module.code, &target, &peers, &module.keep);
    let header = header(&module.name);
    let shift = header.lines().count();
    let file = root.join(&module.path);
    let text = format!("{header}{}", converted.text);
    let before = std::fs::read_to_string(&file).unwrap_or_default();
    let changed = converted.raw.len()
        != prepared
            .imported
            .iter()
            .find(|i| i.path.eq_ignore_ascii_case(path))
            .map_or(0, |i| i.kept.len());
    if !changed {
        return Ok(false);
    }
    let _ = before;
    std::fs::write(&file, text).map_err(|e| format!("{}: {e}", file.display()))?;
    module.procs = converted
        .procs
        .iter()
        .map(|(from, to, name)| (from + shift, to + shift, name.clone()))
        .collect();
    if let Some(item) = prepared
        .imported
        .iter_mut()
        .find(|i| i.path.eq_ignore_ascii_case(path))
    {
        item.kept = converted.raw;
        item.todos = converted
            .todos
            .into_iter()
            .map(|(line, message)| (line + shift, message))
            .collect();
    }
    Ok(true)
}

/// A worksheet of the workbook, for [`sheet_files`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Sheet {
    /// Its code name (`Sheet1`).
    pub code_name: String,
    /// The name on its tab.
    pub name: String,
}

/// The text of a file [`sheet_files`] creates for the worksheet `sheet`, or the workbook when `None`.
pub fn sheet_file_text(sheet: Option<&Sheet>, workbook_code_name: &str) -> String {
    match sheet {
        Some(s) => format!(
            "// シート「{}」（コード名 {}）のモジュールです（同期のときに作成しました）。\n// このファイルに書いたコードが、ブックのシート {} のモジュールになります（仕様 §12.2）。\n// イベントは @on(\"Worksheet.<イベント名>\") を付けた関数で受け取ります。例:\n//\n//   @on(\"Worksheet.Change\")\n//   function onChange(target: Excel.Range) {{\n//       // セルの値が変わったとき。このシートは this で表します: this.range(\"A1\")\n//   }}\n",
            s.name, s.code_name, s.code_name
        ),
        None => format!(
            "// ブック（{workbook_code_name}）のモジュールです（同期のときに作成しました）。\n// このファイルに書いたコードが、ブックのモジュール {workbook_code_name} になります（仕様 §12.2）。\n// イベントは @on(\"Workbook.<イベント名>\") を付けた関数で受け取ります。例:\n//\n//   @on(\"Workbook.Open\")\n//   function onOpen() {{\n//       // ブックを開いたとき。このブックは this で表します: this.name\n//   }}\n"
        ),
    }
}

/// Creates `objects/<CodeName>.ras` for each worksheet of `sheets`, and
/// `objects/ThisWorkbook.ras` for the workbook, that has no file yet (spec §14.7.2), and
/// records them in `prepared.created` (relative to `root`, with what they are for).
///
/// * Nothing is created while `objects/` holds a file no sheet answers to: a sheet whose code
///   name was changed, whose file is to be renamed.
/// * A sheet whose module holds code RAS generated is a copy of another sheet (Excel copies
///   the code with the sheet): it gets a copy of the file of the sheet it was copied from —
///   the sheet whose written code its module holds. Code added in the VBE since is appended
///   (spec §14.7.3). A copy whose source is not found gets no file; its module is left alone.
/// * A file without code does not change the module (spec §12.2), so the other files cost
///   nothing until something is written in them.
pub fn sheet_files(
    root: &Path,
    sheets: &[Sheet],
    book: &mut Book,
    prepared: &mut Prepared,
) -> Result<(), String> {
    let workbook_code_name = book.workbook_code_name;
    let same = |a: &str, b: &str| a.to_lowercase() == b.to_lowercase();
    let existing: Vec<String> = object_files(root)
        .iter()
        .filter_map(|p| p.file_stem().map(|s| s.to_string_lossy().to_string()))
        .collect();
    let orphan = existing.iter().any(|name| {
        !same(name, "ThisWorkbook") && !sheets.iter().any(|s| same(&s.code_name, name))
    });
    let mut targets: Vec<Option<&Sheet>> = sheets
        .iter()
        .filter(|s| !s.code_name.is_empty())
        .map(Some)
        .collect();
    if !workbook_code_name.is_empty() {
        targets.push(None);
    }
    for sheet in targets {
        let (file, what) = match sheet {
            Some(s) => (
                s.code_name.as_str(),
                format!("シート「{}」（{}）", s.name, s.code_name),
            ),
            None => ("ThisWorkbook", format!("ブック（{workbook_code_name}）")),
        };
        if existing.iter().any(|name| same(name, file)) || (orphan && sheet.is_some()) {
            continue;
        }
        let path = ras_compiler::object_code_path(root, file);
        if let Some(dir) = path.parent() {
            std::fs::create_dir_all(dir).map_err(|e| format!("{}: {e}", dir.display()))?;
        }
        // The module of the sheet holds generated code: a copy of another sheet.
        let module = sheet.and_then(|s| {
            book.modules
                .iter()
                .find(|m| m.kind == ModuleKind::Document && same(&m.name, &s.code_name))
        });
        let pane = module.map(|m| hand_written(&m.code)).unwrap_or_default();
        if let (Some(s), true) = (sheet, ras_compiler::emit::has_drifted(&pane).is_some()) {
            // The sheet whose written code the module holds; one it holds unchanged first.
            let mut source: Option<(&Sheet, crate::vbdiff::Diff)> = None;
            for other in sheets.iter().filter(|o| {
                !same(&o.code_name, &s.code_name)
                    && existing.iter().any(|name| same(name, &o.code_name))
            }) {
                let Some(written) = (book.written)(&other.code_name) else {
                    continue;
                };
                let diff = crate::vbdiff::diff(&written, &pane);
                if diff.has_drift() || crate::vbdiff::split(&written).procedures.is_empty() {
                    continue;
                }
                let exact = !diff.has_additions();
                if source.is_none() || exact {
                    source = Some((other, diff));
                }
                if exact {
                    break;
                }
            }
            let Some((from, diff)) = source else {
                prepared.skipped.push(format!("{what}のモジュールには RAS が生成したコードがありますが、どのシートのファイルから生成されたものか分からないため、ファイルを作成していません（モジュールはそのまま残しています）"));
                continue;
            };
            let from_path = ras_compiler::object_code_path(
                root,
                &existing
                    .iter()
                    .find(|name| same(name, &from.code_name))
                    .cloned()
                    .unwrap_or_else(|| from.code_name.clone()),
            );
            let text = std::fs::read_to_string(&from_path)
                .map_err(|e| format!("{}: {e}", from_path.display()))?;
            std::fs::write(&path, &text).map_err(|e| format!("{}: {e}", path.display()))?;
            prepared.created.push((
                relative(root, &path),
                format!(
                    "{what}（シート「{}」の複製。{} を写しました）",
                    from.name,
                    relative(root, &from_path)
                ),
            ));
            if diff.has_additions() {
                let controls = book
                    .sheet_controls
                    .iter()
                    .find(|d| same(&d.name, &s.code_name))
                    .map(|d| &d.controls[..])
                    .unwrap_or(&[]);
                let target = Target::Document {
                    code_name: &s.code_name,
                    workbook: false,
                    controls,
                };
                let todos =
                    append_added(root, &path, &s.code_name, &diff, &target, true, prepared)?;
                // (Taking it back removes the copy: the file did not exist before.)
                if let Some(a) = prepared.appended.last_mut() {
                    a.original = None;
                }
                let kind = ImportKind::Document { workbook: false };
                if todos > 0 {
                    prepared.pending.push(Pending {
                        name: s.code_name.clone(),
                        generated: s.code_name.clone(),
                        kind,
                        path: relative(root, &path),
                        todos,
                    });
                } else {
                    prepared.adopt.push((s.code_name.clone(), kind));
                }
            }
            continue;
        }
        std::fs::write(&path, sheet_file_text(sheet, workbook_code_name))
            .map_err(|e| format!("{}: {e}", path.display()))?;
        prepared.created.push((relative(root, &path), what));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn project() -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "ras-sheet-files-{}-{:?}",
            std::process::id(),
            std::thread::current().id()
        ));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        dir
    }

    #[test]
    fn sheets_without_a_file_get_one() {
        let root = project();
        std::fs::create_dir_all(root.join("objects")).unwrap();
        std::fs::write(root.join("objects/Sheet3.ras"), "// mine\n").unwrap();
        let sheets: Vec<Sheet> = [("Sheet1", "売上"), ("Sheet2", "集計"), ("Sheet3", "メモ")]
            .iter()
            .map(|(c, n)| Sheet {
                code_name: c.to_string(),
                name: n.to_string(),
            })
            .collect();
        let created = create(&root, &sheets, &[], &mut |_: &str| None).created;
        let paths: Vec<&str> = created.iter().map(|(p, _)| p.as_str()).collect();
        // Sheet3 has its file.
        assert_eq!(
            paths,
            [
                "objects/Sheet1.ras",
                "objects/Sheet2.ras",
                "objects/ThisWorkbook.ras"
            ]
        );
        let text = std::fs::read_to_string(root.join("objects/Sheet1.ras")).unwrap();
        assert!(
            text.contains("シート「売上」（コード名 Sheet1）")
                && text.contains("@on(\"Worksheet.Change\")"),
            "{text}"
        );
        assert_eq!(
            std::fs::read_to_string(root.join("objects/Sheet3.ras")).unwrap(),
            "// mine\n"
        );
        // Nothing more the next time.
        assert!(create(&root, &sheets, &[], &mut |_: &str| None)
            .created
            .is_empty());
        // A file no sheet answers to (a code name that was changed): no file for the sheets.
        std::fs::remove_file(root.join("objects/Sheet2.ras")).unwrap();
        std::fs::write(root.join("objects/Old.ras"), "export function a() {}\n").unwrap();
        assert!(create(&root, &sheets, &[], &mut |_: &str| None)
            .created
            .is_empty());
        let _ = std::fs::remove_dir_all(&root);
    }

    /// Runs [`sheet_files`] for a workbook with the document modules `modules`.
    fn create(
        root: &Path,
        sheets: &[Sheet],
        modules: &[BookModule],
        written: &mut dyn FnMut(&str) -> Option<String>,
    ) -> Prepared {
        let mut prepared = Prepared::default();
        let mut form = |_: &str| Err(String::new());
        let mut book = Book {
            modules,
            workbook_code_name: "ThisWorkbook",
            sheet_controls: &[],
            form: &mut form,
            written,
        };
        sheet_files(root, sheets, &mut book, &mut prepared).unwrap();
        prepared
    }

    const SHEET1_RAS: &str = "@on(\"Worksheet.Change\")\nfunction onChange(target: Excel.Range) {\n    this.range(\"D1\").value = target.address;\n}\n";

    /// The code RAS generates for `objects/Sheet1.ras` holding `SHEET1_RAS`.
    fn generated() -> String {
        let cfg = ras_compiler::Config {
            sheet_code_names: vec!["Sheet1".into(), "Sheet2".into()],
            ..Default::default()
        };
        let (_, out) = ras_compiler::compile_str(
            &[("objects/Sheet1", SHEET1_RAS)],
            cfg,
            &ras_compiler::BuildOptions {
                no_attributes: true,
                ..Default::default()
            },
        );
        out.files
            .iter()
            .find(|f| f.name == "Sheet1")
            .expect("Sheet1")
            .code
            .clone()
    }

    #[test]
    fn a_copied_sheet_gets_a_copy_of_the_file() {
        let root = project();
        std::fs::create_dir_all(root.join("objects")).unwrap();
        std::fs::write(root.join("objects/Sheet1.ras"), SHEET1_RAS).unwrap();
        std::fs::write(root.join("objects/ThisWorkbook.ras"), "// nothing\n").unwrap();
        let code = generated();
        let sheets: Vec<Sheet> = [
            ("Sheet1", "売上"),
            ("Sheet2", "売上 (2)"),
            ("Sheet3", "新規"),
        ]
        .iter()
        .map(|(c, n)| Sheet {
            code_name: c.to_string(),
            name: n.to_string(),
        })
        .collect();
        // Sheet2 was copied from Sheet1 in Excel, and a procedure was added to it in the VBE.
        let copy = format!("{code}\r\nPrivate Sub Worksheet_Activate()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub\r\n");
        let modules = [
            BookModule {
                name: "Sheet1".into(),
                kind: ModuleKind::Document,
                code: code.clone(),
            },
            BookModule {
                name: "Sheet2".into(),
                kind: ModuleKind::Document,
                code: copy,
            },
            BookModule {
                name: "Sheet3".into(),
                kind: ModuleKind::Document,
                code: String::new(),
            },
        ];
        let mut written = |name: &str| (name == "Sheet1").then(|| code.clone());
        let prepared = create(&root, &sheets, &modules, &mut written);
        let paths: Vec<&str> = prepared.created.iter().map(|(p, _)| p.as_str()).collect();
        assert_eq!(paths, ["objects/Sheet2.ras", "objects/Sheet3.ras"]);
        assert!(
            prepared.created[0].1.contains("複製")
                && prepared.created[0].1.contains("objects/Sheet1.ras"),
            "{:?}",
            prepared.created
        );
        let text = std::fs::read_to_string(root.join("objects/Sheet2.ras")).unwrap();
        // The file of the source, and what was added to the copy.
        assert!(text.starts_with(SHEET1_RAS), "{text}");
        assert!(
            text.contains("// VBE で追加されたコードを取り込みました（")
                && text.contains("@on(\"Worksheet.Activate\")")
                && text.contains("this.range(\"A1\").value = 1;"),
            "{text}"
        );
        assert_eq!(prepared.appended.len(), 1, "{:?}", prepared.appended);
        assert!(prepared.adopts("Sheet2"));
        // The new sheet gets the plain file.
        assert!(!ras_has_code(
            &std::fs::read_to_string(root.join("objects/Sheet3.ras")).unwrap()
        ));

        // A copy whose source is not known gets no file.
        std::fs::remove_file(root.join("objects/Sheet2.ras")).unwrap();
        let prepared = create(&root, &sheets, &modules, &mut |_: &str| None);
        assert!(
            prepared.created.is_empty() && prepared.skipped.len() == 1,
            "{prepared:?}"
        );
        assert!(!root.join("objects/Sheet2.ras").exists());
        let _ = std::fs::remove_dir_all(&root);
    }

    #[test]
    fn the_files_build_into_nothing() {
        let sheet = Sheet {
            code_name: "Sheet1".into(),
            name: "Sheet1".into(),
        };
        for (key, text) in [
            (
                "objects/Sheet1",
                sheet_file_text(Some(&sheet), "ThisWorkbook"),
            ),
            (
                "objects/ThisWorkbook",
                sheet_file_text(None, "ThisWorkbook"),
            ),
        ] {
            assert!(!ras_has_code(&text), "{text}");
            let cfg = ras_compiler::Config {
                sheet_code_names: vec!["Sheet1".into()],
                ..Default::default()
            };
            let (prog, out) =
                ras_compiler::compile_str(&[(key, text.as_str())], cfg, &Default::default());
            assert!(
                !prog.diagnostics.iter().any(|d| d.is_error()),
                "{:?}",
                prog.diagnostics
            );
            assert!(
                out.files.iter().all(|f| f.name.starts_with("RAS_")),
                "{:?}",
                out.files.iter().map(|f| &f.name).collect::<Vec<_>>()
            );
        }
    }

    #[test]
    fn what_counts_as_code() {
        assert!(!ras_has_code(
            "// a\n/* b\n c */\n\nimport { x } from \"../m\";\n"
        ));
        assert!(ras_has_code("// a\nlet n = 1;\n"));
        assert!(ras_has_code("export function f() {}\n"));
        assert!(ras_has_code("@on(\"Worksheet.Change\")\nfunction f() {}\n"));
    }
}
