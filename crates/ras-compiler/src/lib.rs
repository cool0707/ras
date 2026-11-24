//! RAS compiler: declaration collection, type checking, lowering to VBA.
//!
//! ```text
//! sources ──collect──▶ Program ──check──▶ typed Program ──lower──▶ VBA modules ──emit──▶ files
//! ```

#![allow(
    clippy::too_many_arguments,
    clippy::type_complexity,
    clippy::wrong_self_convention,
    clippy::should_implement_trait,
    clippy::field_reassign_with_default,
    clippy::needless_range_loop
)]

pub mod astutil;
pub mod check;
pub mod collect;
pub mod config;
pub mod diag;
pub mod emit;
pub mod formula;
pub mod hints;
pub mod intrinsics;
pub mod js;
mod lower;
pub mod model;
pub mod names;
pub mod query;
pub mod refs;
pub mod runtime;
pub mod standin;
pub mod testing;
pub mod userform;
pub mod vba;
pub mod vba2ras;

use std::path::{Path, PathBuf};

pub use collect::SourceInput;
pub use config::{Config, FormControl, FormDecl};
pub use diag::{apply_fixes, Diagnostic, Fix, Severity};
pub use emit::OutputFile;
pub use js::{build_js, JsFile, JsOutput};
pub use model::{ModuleId, ModuleKind, OptNote, Program};

/// Type definitions and library modules bundled with the compiler (spec §15.3).
const STDLIB: &[(&str, ModuleKind, &str)] = &[
    (
        "core.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/core.d.ras"),
    ),
    (
        "vba.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/vba.d.ras"),
    ),
    (
        "excel.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/excel.d.ras"),
    ),
    (
        "office.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/office.d.ras"),
    ),
    (
        "msforms.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/msforms.d.ras"),
    ),
    (
        "scripting.d",
        ModuleKind::TypeDef,
        include_str!("../stdlib/scripting.d.ras"),
    ),
    (
        collect::TEST_DEFS_MODULE,
        ModuleKind::TypeDef,
        include_str!("../stdlib/test.d.ras"),
    ),
    (
        "win32",
        ModuleKind::Std,
        include_str!("../stdlib/win32.ras"),
    ),
    (
        collect::STRUCT_MAP_MODULE,
        ModuleKind::Std,
        include_str!("../stdlib/ras_collections.ras"),
    ),
];

fn stdlib_inputs() -> Vec<SourceInput> {
    STDLIB
        .iter()
        .map(|(key, kind, text)| SourceInput {
            path: PathBuf::from(format!("<stdlib>/{key}.ras")),
            rel_path: format!("<stdlib>/{key}.ras"),
            key: (*key).to_string(),
            kind: kind.clone(),
            text: (*text).to_string(),
        })
        .collect()
}

/// Parses, collects and type-checks a project. Never fails: problems are in `Program::diagnostics`.
pub fn analyze(sources: Vec<SourceInput>, config: Config) -> Program {
    let mut inputs = stdlib_inputs();
    inputs.extend(sources);
    let mut prog = collect::collect(inputs, config);
    check::check(&mut prog);
    testing::check(&mut prog);
    prog
}

#[derive(Debug, Clone, Default)]
pub struct BuildOptions {
    /// `--release`: optimisation level 2 unless overridden, no `Erl` line numbers.
    pub release: bool,
    /// Override of `[build] erl`.
    pub erl: Option<bool>,
    /// Omit `Attribute` lines (code destined for `CodeModule.AddFromString`, spec §16.1.3).
    pub no_attributes: bool,
}

pub struct BuildOutput {
    pub files: Vec<OutputFile>,
    /// Diagnostics produced while generating code (in addition to `Program::diagnostics`).
    pub diagnostics: Vec<Diagnostic>,
    pub notes: Vec<(ModuleId, OptNote)>,
    /// Macro entry points (`Module.proc`).
    pub macros: Vec<String>,
}

/// Generates VBA for an analysed program. Returns no files when the program has errors.
pub fn build(prog: &Program, opts: &BuildOptions) -> BuildOutput {
    if prog.diagnostics.iter().any(|d| d.is_error()) {
        return BuildOutput {
            files: Vec::new(),
            diagnostics: Vec::new(),
            notes: Vec::new(),
            macros: Vec::new(),
        };
    }
    let lowered = lower::lower(prog);
    let erl = opts.erl.unwrap_or(prog.config.build.erl && !opts.release);
    let (mut files, mut diagnostics) = emit::finalize(
        prog,
        &lowered,
        &emit::EmitOptions {
            erl,
            attributes: !opts.no_attributes,
        },
    );
    diagnostics.extend(lowered.diagnostics.iter().cloned());
    if diagnostics.iter().any(|d| d.is_error()) {
        files.clear();
    }
    BuildOutput {
        files,
        diagnostics,
        notes: lowered.notes,
        macros: lowered.macros,
    }
}

/// Reads the project as [`load_project`] does, with its tests (spec §14.12.2): the test files
/// (`[test] include`, `test/**/*.test.ras`) and the helper modules of `test/`. They are
/// type-checked and run by `ras test`, and are no part of what is built into the workbook.
pub fn load_project_with_tests(root: &Path) -> Result<(Config, Vec<SourceInput>), String> {
    let (mut config, mut sources) = load_project(root)?;
    let mut files = Vec::new();
    walk_tests(root, root, &mut files).map_err(|e| format!("{}: {e}", root.display()))?;
    files.sort();
    for path in files {
        let rel = path
            .strip_prefix(root)
            .unwrap_or(&path)
            .to_string_lossy()
            .replace('\\', "/");
        let is_test = config.test.is_test_file(&rel);
        if !is_test && !rel.starts_with("test/") {
            continue;
        }
        if rel.ends_with(".d.ras") {
            continue;
        }
        let text =
            std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
        let text = text
            .strip_prefix('\u{feff}')
            .map(str::to_string)
            .unwrap_or(text);
        if is_test {
            config.test_files.push(rel.clone());
        }
        sources.push(SourceInput {
            key: format!(
                "{}{}",
                collect::TEST_KEY_PREFIX,
                rel.strip_suffix(".ras").unwrap_or(&rel)
            ),
            rel_path: rel,
            kind: ModuleKind::User,
            text,
            path,
        });
    }
    Ok((config, sources))
}

/// The `.ras` files of the project outside what holds no sources (`build/`, hidden folders,
/// `node_modules/`).
fn walk_tests(root: &Path, dir: &Path, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(e) => return Err(e),
    };
    for entry in entries {
        let path = entry?.path();
        let name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        if path.is_dir() {
            if dir == root && matches!(name.as_str(), "build" | "node_modules" | "vba" | "types")
                || name.starts_with('.')
            {
                continue;
            }
            walk_tests(root, &path, out)?;
        } else if name.ends_with(".ras") {
            out.push(path);
        }
    }
    Ok(())
}

/// Reads `ras.toml` and every source file of the project rooted at `root` (spec §15.1).
/// Files matching `[test] include` are tests, wherever they are: they are left out.
pub fn load_project(root: &Path) -> Result<(Config, Vec<SourceInput>), String> {
    let toml_path = root.join("ras.toml");
    let config = match std::fs::read_to_string(&toml_path) {
        Ok(text) => {
            Config::from_toml(&text).map_err(|e| format!("{}: {e}", toml_path.display()))?
        }
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => {
            return Err(format!(
                "{} が見つかりません。`ras init` でプロジェクトを作成してください",
                toml_path.display()
            ));
        }
        Err(e) => return Err(format!("{}: {e}", toml_path.display())),
    };
    let mut sources = Vec::new();
    for (dir, kind) in [("src", ModuleKind::User), ("types", ModuleKind::TypeDef)] {
        let base = root.join(dir);
        let mut files = Vec::new();
        walk(&base, &mut files).map_err(|e| format!("{}: {e}", base.display()))?;
        files.sort();
        for path in files {
            let rel = path
                .strip_prefix(&base)
                .unwrap_or(&path)
                .to_string_lossy()
                .replace('\\', "/");
            if config.test.is_test_file(&format!("{dir}/{rel}")) {
                continue;
            }
            let is_typedef = rel.ends_with(".d.ras");
            let key = rel.strip_suffix(".ras").unwrap_or(&rel).to_string();
            let text =
                std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
            // A leading byte-order mark is not part of the source.
            let text = text
                .strip_prefix('\u{feff}')
                .map(str::to_string)
                .unwrap_or(text);
            sources.push(SourceInput {
                rel_path: format!("{dir}/{rel}"),
                key,
                kind: if is_typedef {
                    ModuleKind::TypeDef
                } else {
                    kind.clone()
                },
                text,
                path,
            });
        }
    }
    let (form_sources, forms) = load_forms(root)?;
    sources.extend(form_sources);
    sources.extend(load_objects(root)?);
    let mut config = config;
    config.forms = forms;
    Ok((config, sources))
}

/// Directory of the files of the sheets and of `ThisWorkbook` (spec §12.2): the VBE's
/// "Microsoft Excel Objects".
pub const OBJECTS_DIR: &str = "objects";

/// Path of the file of a sheet (by its code name) or of `ThisWorkbook`: `objects/<CodeName>.ras`.
pub fn object_code_path(root: &Path, code_name: &str) -> PathBuf {
    root.join(OBJECTS_DIR).join(format!("{code_name}.ras"))
}

/// Reads the files of the sheets and of `ThisWorkbook` of the project at `root`
/// (`objects/<CodeName>.ras`, spec §12.2). Sub-directories are not read.
pub fn load_objects(root: &Path) -> Result<Vec<SourceInput>, String> {
    let dir = root.join(OBJECTS_DIR);
    let Ok(entries) = std::fs::read_dir(&dir) else {
        return Ok(Vec::new());
    };
    let mut paths: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.is_file())
        .collect();
    paths.sort();
    let mut sources = Vec::new();
    for path in paths {
        let file_name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        let Some(stem) = file_name.strip_suffix(".ras") else {
            continue;
        };
        if stem.ends_with(".d") {
            continue;
        }
        let text =
            std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
        let text = text
            .strip_prefix('\u{feff}')
            .map(str::to_string)
            .unwrap_or(text);
        sources.push(object_source(path.clone(), stem, text));
    }
    Ok(sources)
}

/// The source input of the file `objects/<name>.ras`.
pub fn object_source(path: PathBuf, name: &str, text: String) -> SourceInput {
    SourceInput {
        path,
        rel_path: format!("{OBJECTS_DIR}/{name}.ras"),
        key: format!("{}{name}", collect::OBJECT_KEY_PREFIX),
        kind: ModuleKind::User,
        text,
    }
}

/// Path of a form's RAS code-behind: `forms/<Name>.ras` (spec §14.11.1).
pub fn form_code_path(root: &Path, name: &str) -> PathBuf {
    root.join("forms").join(format!("{name}.ras"))
}

/// Reads the RAS code-behinds of the forms of the project at `root` (`forms/<Name>.ras`) and
/// the declarations (name and controls) of every `forms/<Name>.form.json`.
pub fn load_forms(root: &Path) -> Result<(Vec<SourceInput>, Vec<FormDecl>), String> {
    let dir = root.join("forms");
    let Ok(entries) = std::fs::read_dir(&dir) else {
        return Ok((Vec::new(), Vec::new()));
    };
    let mut paths: Vec<PathBuf> = entries
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.is_file())
        .collect();
    paths.sort();
    let mut sources = Vec::new();
    let mut forms = Vec::new();
    for path in paths {
        let file_name = path
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        if let Some(stem) = file_name.strip_suffix(".form.json") {
            let Ok(text) = std::fs::read_to_string(&path) else {
                continue;
            };
            if let Some(mut decl) = FormDecl::from_form_json(&text) {
                // The file name is the form's name (a mismatch is reported by the form checks).
                decl.name = stem.to_string();
                forms.push(decl);
            }
        } else if let Some(stem) = file_name.strip_suffix(".ras") {
            if stem.contains('.') {
                continue;
            }
            let text =
                std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
            let text = text
                .strip_prefix('\u{feff}')
                .map(str::to_string)
                .unwrap_or(text);
            sources.push(form_source(path.clone(), stem, text));
        }
    }
    Ok((sources, forms))
}

/// The source input of the code-behind `forms/<name>.ras`.
pub fn form_source(path: PathBuf, name: &str, text: String) -> SourceInput {
    SourceInput {
        path,
        rel_path: format!("forms/{name}.ras"),
        key: format!("{}{name}", collect::FORM_KEY_PREFIX),
        kind: ModuleKind::User,
        text,
    }
}

fn walk(dir: &Path, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => return Ok(()),
        Err(e) => return Err(e),
    };
    for entry in entries {
        let path = entry?.path();
        if path.is_dir() {
            walk(&path, out)?;
        } else if path.extension().is_some_and(|e| e == "ras") {
            out.push(path);
        }
    }
    Ok(())
}

/// `file:line:col: severity[code]: message` lines for every diagnostic, sorted by position.
pub fn render_diagnostics(prog: &Program, diags: &[Diagnostic]) -> String {
    let mut items: Vec<&Diagnostic> = diags.iter().collect();
    items.sort_by_key(|d| (d.module, u32::from(d.span.start()), d.severity));
    let mut out = String::new();
    for d in items {
        let sev = match d.severity {
            Severity::Error => "error",
            Severity::Warning => "warning",
            Severity::Info => "info",
            Severity::Hint => "hint",
        };
        match d.module {
            Some(m) => {
                let md = &prog.modules[m];
                let (line, col) = md.line_col_utf16(u32::from(d.span.start()));
                out.push_str(&format!(
                    "{}:{}:{}: {sev}[{}]: {}\n",
                    md.rel_path,
                    line + 1,
                    col + 1,
                    d.code,
                    d.message
                ));
            }
            None => out.push_str(&format!("{sev}[{}]: {}\n", d.code, d.message)),
        }
    }
    out
}

/// Compiles in-memory sources (tests, playground). `files` are `(import key, text)` pairs.
pub fn compile_str(
    files: &[(&str, &str)],
    config: Config,
    opts: &BuildOptions,
) -> (Program, BuildOutput) {
    let sources = files
        .iter()
        .map(|(key, text)| match key.strip_prefix("objects/") {
            // `objects/Sheet1`: the file of a sheet or of `ThisWorkbook` (spec §12.2).
            Some(name) => object_source(
                PathBuf::from(format!("objects/{name}.ras")),
                name,
                (*text).to_string(),
            ),
            None => SourceInput {
                path: PathBuf::from(format!("src/{key}.ras")),
                rel_path: format!("src/{key}.ras"),
                key: (*key).to_string(),
                kind: if key.ends_with(".d") {
                    ModuleKind::TypeDef
                } else {
                    ModuleKind::User
                },
                text: (*text).to_string(),
            },
        })
        .collect();
    let prog = analyze(sources, config);
    let out = build(&prog, opts);
    (prog, out)
}
