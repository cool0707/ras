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
//! * the hand-written code of the sheet `Sheet1` → `src/Sheet1Events.ras`
//!   (its event procedures become `@on("Sheet1.Change")` handlers),
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
    /// The designer of a form made in the VBE, by name (exported from the open workbook,
    /// or read from the workbook file).
    pub form: &'a mut dyn FnMut(&str) -> Result<UserForm, String>,
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
}

/// A component of the workbook that has a converted file with `// TODO(VBA):` lines left:
/// the sync leaves it as it is and does not write what the file generates.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Pending {
    /// Name of the component in the workbook.
    pub name: String,
    /// Name of the module generated from the file (`Sheet1Events` for `Sheet1`); the
    /// same as `name` for a form.
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
    /// Files whose marker line becomes a plain comment once the workbook has the generated code.
    marked: Vec<PathBuf>,
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
            };
            out.push(format!("ブックの{what} `{}` を {} に取り込みました", item.name, item.path));
            for warning in &item.warnings {
                out.push(format!("  {warning}"));
            }
            for (line, message) in &item.todos {
                out.push(format!("  {}:{line}: VBA から RAS に変換できませんでした: {message}", item.path));
            }
        }
        for p in &self.pending {
            out.push(format!(
                "{} には変換できなかった行（`// TODO(VBA):`）が {} 行残っているため、ブックの `{}` は置き換えずにそのまま残しています。直して `// TODO(VBA):` の行を消すと、次の同期で置き換えます",
                p.path, p.todos, p.name
            ));
        }
        out.extend(self.skipped.iter().cloned());
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
    let first = text.strip_prefix('\u{feff}').unwrap_or(text).lines().next()?;
    let rest = first.strip_prefix(IMPORT_MARKER)?;
    rest.split_whitespace().next().map(str::to_string)
}

fn todo_count(text: &str) -> usize {
    text.lines().filter(|l| l.trim_start().starts_with(TODO_MARKER)).count()
}

/// Whether code holds anything besides `Attribute` lines, `Option` statements, comments and blank lines.
fn has_code(code: &str) -> bool {
    code.lines().map(str::trim).any(|l| {
        !l.is_empty() && !l.starts_with('\'') && !l.starts_with("Attribute ") && !l.starts_with("VERSION ") && !l.to_ascii_lowercase().starts_with("option ") && !l.to_ascii_lowercase().starts_with("rem ")
    })
}

/// The code pane text: without the `Attribute` header, with LF line ends.
fn code_pane(code: &str) -> String {
    code.replace("\r\n", "\n").lines().skip_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION ")).map(|l| format!("{l}\n")).collect()
}

/// The hand-written part of a document module (without the RAS-managed region).
pub fn hand_written(code: &str) -> String {
    code_pane(&ras_ovba::merge_managed_region(&code.replace("\r\n", "\n"), ""))
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
    path.strip_prefix(root).unwrap_or(path).to_string_lossy().replace('\\', "/")
}

fn walk(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else { return };
    for e in entries.flatten() {
        let path = e.path();
        if path.is_dir() {
            walk(&path, out);
        } else if path.extension().is_some_and(|x| x == "ras") && !path.to_string_lossy().ends_with(".d.ras") {
            out.push(path);
        }
    }
}

/// The modules of `src/` and the code-behinds of `forms/`.
fn project_files(root: &Path) -> Vec<ProjectFile> {
    let mut paths = Vec::new();
    walk(&root.join("src"), &mut paths);
    paths.sort();
    let mut out = Vec::new();
    for path in paths {
        let key = relative(&root.join("src"), &path);
        let key = key.strip_suffix(".ras").unwrap_or(&key).to_string();
        let text = std::fs::read_to_string(&path).unwrap_or_default();
        out.push(ProjectFile { vba_name: vba_module_name(&key), marker: marker_of(&text), todos: todo_count(&text), path });
    }
    if let Ok(forms) = workspace::load(root) {
        for form in forms {
            let text = std::fs::read_to_string(&form.code_path).unwrap_or_default();
            out.push(ProjectFile { vba_name: form.name.clone(), marker: marker_of(&text), todos: todo_count(&text), path: form.code_path });
        }
    }
    out
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
pub fn prepare(root: &Path, book: Book, codepage: u16) -> Result<Prepared, String> {
    let mut prepared = Prepared::default();
    let files = project_files(root);
    let taken = |name: &str| files.iter().any(|f| f.vba_name.eq_ignore_ascii_case(name)) || book.modules.iter().any(|m| m.name.eq_ignore_ascii_case(name));
    let marked = |name: &str| files.iter().find(|f| f.marker.as_deref().is_some_and(|m| m.eq_ignore_ascii_case(name)));

    let mut candidates: Vec<Candidate> = Vec::new();
    for module in book.modules {
        let (kind, code) = match module.kind {
            ModuleKind::Standard => {
                // RAS-generated (also the temporary module of the sync), or nothing in it.
                if ras_compiler::emit::has_drifted(&module.code).is_some() || module.name.starts_with("RAS_") || !has_code(&module.code) {
                    continue;
                }
                (ImportKind::Standard, code_pane(&module.code))
            }
            ModuleKind::Document => {
                let code = hand_written(&module.code);
                if !has_code(&code) {
                    continue;
                }
                (ImportKind::Document { workbook: module.name.eq_ignore_ascii_case(book.workbook_code_name) }, code)
            }
            ModuleKind::Form => {
                if ras_ovba::is_generated_form(&module.code) {
                    continue;
                }
                (ImportKind::Form, code_pane(&module.code))
            }
            _ => continue,
        };
        // Converted earlier: the file decides.
        if let Some(file) = marked(&module.name) {
            let generated = file.vba_name.clone();
            if file.todos > 0 {
                prepared.pending.push(Pending { name: module.name.clone(), generated, kind, path: relative(root, &file.path), todos: file.todos });
            } else {
                prepared.adopt.push((module.name.clone(), kind));
                prepared.marked.push(file.path.clone());
            }
            continue;
        }
        match kind {
            ImportKind::Standard => {
                // A module of the same name (or file) is the project's: the usual rules of the sync apply.
                let key = module.name.clone();
                if files.iter().any(|f| f.vba_name.eq_ignore_ascii_case(&module.name) || relative(&root.join("src"), &f.path).eq_ignore_ascii_case(&format!("{key}.ras"))) {
                    continue;
                }
                let ras_name = vba_module_name(&key);
                if !ras_name.eq_ignore_ascii_case(&module.name) && taken(&ras_name) {
                    prepared.skipped.push(format!("ブックの標準モジュール `{}` は、RAS のモジュール名 `{ras_name}` が既に使われているため取り込みませんでした", module.name));
                    continue;
                }
                candidates.push(Candidate { module, kind, code, key });
            }
            ImportKind::Document { .. } => {
                let Some(key) = (1..100).map(|n| if n == 1 { format!("{}Events", module.name) } else { format!("{}Events{n}", module.name) }).find(|k| !taken(&vba_module_name(k)) && !root.join("src").join(format!("{k}.ras")).exists()) else {
                    continue;
                };
                candidates.push(Candidate { module, kind, code, key });
            }
            ImportKind::Form => {
                let dir = workspace::forms_dir(root);
                let exists = ["form.json", "ras", "bas"].iter().any(|ext| dir.join(format!("{}.{ext}", module.name)).exists()) || files.iter().any(|f| f.vba_name.eq_ignore_ascii_case(&module.name));
                if exists {
                    continue;
                }
                candidates.push(Candidate { module, kind, code, key: String::new() });
            }
        }
    }

    // Names the converted modules make public to each other.
    let peers: Vec<Peer> = candidates
        .iter()
        .filter_map(|c| {
            let target = match c.kind {
                ImportKind::Standard => Target::Standard,
                ImportKind::Document { workbook } => Target::Document { code_name: &c.module.name, workbook },
                ImportKind::Form => return None,
            };
            Some(Peer { vba_name: c.module.name.clone(), key: c.key.clone(), names: vba2ras::public_names(&c.code, &target) })
        })
        .collect();

    let vba_dir = root.join("vba");
    for c in &candidates {
        let name = &c.module.name;
        // The original, as `ras unpack` writes it.
        let ext = match c.kind {
            ImportKind::Standard => "bas",
            ImportKind::Document { .. } => "cls",
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
                        prepared.skipped.push(format!("ブックのユーザーフォーム `{name}` を取り込めませんでした: {e}"));
                        continue;
                    }
                };
                form.code = c.code.clone();
                let pulled = match workspace::pull_with(root, &form, codepage, false, &peers, Some(&header)) {
                    Ok(PullOutcome::Written(pulled)) => pulled,
                    Ok(PullOutcome::Exists(_)) => continue,
                    Err(e) => {
                        prepared.skipped.push(format!("ブックのユーザーフォーム `{name}` を取り込めませんでした: {e}"));
                        continue;
                    }
                };
                let mut other_files = vec![relative(root, &pulled.spec_path)];
                other_files.extend(pulled.pictures.iter().cloned());
                other_files.push(copy);
                let warnings = pulled.warnings.iter().map(|w| if w.control.is_empty() { w.message.clone() } else { format!("{}: {}", w.control, w.message) }).collect();
                Imported { name: name.clone(), kind: c.kind, path: relative(root, &pulled.code_path), other_files, todos: pulled.todos, warnings }
            }
            ImportKind::Standard | ImportKind::Document { .. } => {
                let target = match c.kind {
                    ImportKind::Document { workbook } => Target::Document { code_name: name, workbook },
                    _ => Target::Standard,
                };
                let converted = vba2ras::convert(&c.code, &target, &peers);
                let shift = header.lines().count();
                let path = root.join("src").join(format!("{}.ras", c.key));
                if let Some(parent) = path.parent() {
                    std::fs::create_dir_all(parent).map_err(|e| format!("{}: {e}", parent.display()))?;
                }
                std::fs::write(&path, format!("{header}{}", converted.text)).map_err(|e| format!("{}: {e}", path.display()))?;
                let todos = converted.todos.into_iter().map(|(line, message)| (line + shift, message)).collect();
                Imported { name: name.clone(), kind: c.kind, path: relative(root, &path), other_files: vec![copy], todos, warnings: Vec::new() }
            }
        };
        let generated = match c.kind {
            ImportKind::Form => name.clone(),
            _ => vba_module_name(&c.key),
        };
        if imported.todos.is_empty() {
            prepared.adopt.push((name.clone(), c.kind));
            prepared.marked.push(root.join(&imported.path));
        } else {
            prepared.pending.push(Pending { name: name.clone(), generated, kind: c.kind, path: imported.path.clone(), todos: imported.todos.len() });
        }
        prepared.imported.push(imported);
    }
    Ok(prepared)
}

/// After a sync that replaced the components [`Prepared::adopt`] names: turns the marker
/// line of their files into a plain comment, so that they are ordinary RAS modules. The
/// line stays (as a comment), so the line numbers — and the generated code — do not change.
pub fn finish(prepared: &Prepared) -> Result<(), String> {
    for path in &prepared.marked {
        let Ok(text) = std::fs::read_to_string(path) else { continue };
        let body = text.strip_prefix('\u{feff}').unwrap_or(&text);
        let Some(first) = body.lines().next() else { continue };
        let Some(name) = marker_of(first) else { continue };
        let bom = if text.starts_with('\u{feff}') { "\u{feff}" } else { "" };
        let comment = format!("// ブックの VBA（{name}）から変換したファイルです。元の VBA は vba/ に保存しています。");
        std::fs::write(path, format!("{bom}{comment}{}", &body[first.len()..])).map_err(|e| format!("{}: {e}", path.display()))?;
    }
    Ok(())
}
