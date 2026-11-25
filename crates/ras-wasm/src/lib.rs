//! The compiler and the stand-in for Excel as WebAssembly: what the playground (`web/`) runs in
//! the browser, without the `ras` program, Node or Excel.
//!
//! - [`compile`]: `ras check` / `ras build` / `ras build --target js` of a project held in memory.
//! - [`format`]: `ras fmt` of one file; [`explain`]: `ras explain <code>`; [`encode_vba`]: the
//!   bytes of a `.bas` / `.cls` as `ras build` writes them.
//! - [`Book`]: the workbook of the stand-in for Excel, as the `ras_book` addon gives it to Node
//!   (`crates/ras-book`); its methods are those of `ras_workbook::js`.
//! - [`pack_book`]: `ras pack` into the bytes of a workbook (with the user forms of `forms/`);
//!   [`new_form_spec`]: the design of a new form (`ras form new`); [`sheet_controls`]: the
//!   ActiveX controls of the sheets of a workbook, for the compiler.
//! - [`diff_text`], [`view_page`]: what `ras run` prints of the cells a program changed, and the
//!   page of `ras book view`; [`edit_book`]: `ras book edit`, for the edits made on that page.
//!
//! Everything crosses as JSON text, numbers and bytes, so that the JavaScript side needs nothing
//! but what `wasm-bindgen` generates.

use std::path::{Path, PathBuf};

use ras_compiler::vba::VModuleKind;
use ras_compiler::{BuildOptions, Config, ModuleKind, Severity, SourceInput};
use ras_workbook::{book, diff, js, view};
use serde_json::{Value as Json, json};
use wasm_bindgen::prelude::*;

/// The directory the debug build is written to, relative to the project (as `ras run` does).
const DEBUG_DIR: &str = "build/debug";

fn source_of(path: &str, text: &str) -> Option<SourceInput> {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text).to_string();
    if let Some(rel) = path.strip_prefix("objects/") {
        let name = rel.strip_suffix(".ras")?;
        return Some(ras_compiler::object_source(PathBuf::from(path), name, text));
    }
    if let Some(rel) = path.strip_prefix("forms/") {
        // The code-behind of a form (`forms/<Name>.ras`); its design is read by `read_input`.
        let name = rel
            .strip_suffix(".ras")
            .filter(|n| !n.contains(['.', '/']))?;
        return Some(ras_compiler::form_source(PathBuf::from(path), name, text));
    }
    let (dir, rel) = path.split_once('/')?;
    let kind = match dir {
        "src" => ModuleKind::User,
        "types" => ModuleKind::TypeDef,
        _ => return None,
    };
    let key = rel.strip_suffix(".ras")?.to_string();
    Some(SourceInput {
        path: PathBuf::from(path),
        rel_path: path.to_string(),
        kind: if rel.ends_with(".d.ras") {
            ModuleKind::TypeDef
        } else {
            kind
        },
        key,
        text,
    })
}

/// A test file or a helper of `test/`, as `load_project_with_tests` reads them (spec §14.12.2).
fn test_source(path: &str, text: &str) -> Option<SourceInput> {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text).to_string();
    Some(SourceInput {
        key: format!(
            "{}{}",
            ras_compiler::collect::TEST_KEY_PREFIX,
            path.strip_suffix(".ras")?
        ),
        rel_path: path.to_string(),
        path: PathBuf::from(path),
        kind: ModuleKind::User,
        text,
    })
}

fn severity(s: Severity) -> &'static str {
    match s {
        Severity::Error => "error",
        Severity::Warning => "warning",
        Severity::Info => "info",
        Severity::Hint => "hint",
    }
}

/// Compiles a project held in memory. `input` is JSON:
///
/// ```json
/// { "files": [{ "path": "src/main.ras", "text": "..." }],
///   "toml": "[project]\nname = \"Playground\"",
///   "sheets": ["Sheet1"],
///   "sheetControls": [{ "name": "Sheet1", "controls": [{ "name": "CommandButton1", "type": "CommandButton" }] }],
///   "release": false,
///   "js": true }
/// ```
///
/// Paths are those of the project: `src/`, `types/` (`.d.ras`), `objects/` (the sheets and
/// `ThisWorkbook`), `forms/` (`<Name>.form.json` and its code-behind `<Name>.ras`) and `test/`.
/// `sheets` are the code names of the sheets of the workbook, `sheetControls` the ActiveX
/// controls on them (as [`sheet_controls`] reads them; without it they are unknown).
///
/// Returns JSON: `{ diagnostics: [{ path, line, col, endLine, endCol, severity, code, message }],
/// vba: [{ name, fileName, text, file }], js: [{ path, text }] | null, entries, macros, tests, error }`,
/// lines and columns from 0 (columns in UTF-16 code units). `js` is the debug build, laid out as
/// `build/debug` (the paths are relative to it).
#[wasm_bindgen]
pub fn compile(input: &str) -> String {
    compile_json(input).to_string()
}

/// The project of an input of [`compile`]: its configuration, its sources and its tests.
struct Input {
    json: Json,
    config: Config,
    sources: Vec<SourceInput>,
    tests: Vec<SourceInput>,
    /// The designs of the forms (`forms/<Name>.form.json`).
    forms: Vec<ras_form::workspace::LoadedForm>,
}

fn read_input(input: &str) -> Result<Input, String> {
    let input: Json =
        serde_json::from_str(input).map_err(|e| format!("入力を読み取れません: {e}"))?;
    let mut config = Config::from_toml(input["toml"].as_str().unwrap_or_default())
        .map_err(|e| format!("ras.toml: {e}"))?;
    if let Some(sheets) = input["sheets"].as_array() {
        config.sheet_code_names = sheets
            .iter()
            .filter_map(|s| s.as_str().map(str::to_string))
            .collect();
    }
    if let Some(sheets) = input["sheetControls"].as_array() {
        config.sheet_controls = sheets
            .iter()
            .map(|s| ras_compiler::FormDecl {
                name: s["name"].as_str().unwrap_or_default().to_string(),
                controls: s["controls"]
                    .as_array()
                    .into_iter()
                    .flatten()
                    .map(|c| ras_compiler::FormControl {
                        name: c["name"].as_str().unwrap_or_default().to_string(),
                        control_type: c["type"].as_str().unwrap_or_default().to_string(),
                    })
                    .collect(),
            })
            .filter(|s| !s.controls.is_empty())
            .collect();
        config.sheet_controls_known = true;
    }
    // The VBA is built without the tests (`ras build`); they are checked and run (`ras check`,
    // `ras test`) with the rest.
    let mut sources = Vec::new();
    let mut tests = Vec::new();
    let mut forms = Vec::new();
    for f in input["files"].as_array().into_iter().flatten() {
        let (Some(path), Some(text)) = (f["path"].as_str(), f["text"].as_str()) else {
            continue;
        };
        if let Some(file_name) = path
            .strip_prefix("forms/")
            .filter(|n| n.ends_with(".form.json") && !n.contains('/'))
        {
            let form = ras_form::workspace::loaded(Path::new("forms"), file_name, text.to_string());
            if let Some(mut decl) = ras_compiler::FormDecl::from_form_json(&form.text) {
                // The file name is the form's name (a mismatch is reported by the form checks).
                decl.name = form.name.clone();
                config.forms.push(decl);
            }
            forms.push(form);
            continue;
        }
        let is_test = config.test.is_test_file(path);
        if is_test || path.starts_with("test/") {
            if path.ends_with(".d.ras") {
                continue;
            }
            if is_test {
                config.test_files.push(path.to_string());
            }
            tests.extend(test_source(path, text));
        } else {
            sources.extend(source_of(path, text));
        }
    }
    Ok(Input {
        json: input,
        config,
        sources,
        tests,
        forms,
    })
}

fn compile_json(input: &str) -> Json {
    let Input {
        json: input,
        config,
        mut sources,
        tests,
        forms,
    } = match read_input(input) {
        Ok(i) => i,
        Err(e) => return json!({ "error": e }),
    };
    let opts = BuildOptions {
        release: input["release"].as_bool().unwrap_or(false),
        ..BuildOptions::default()
    };
    let build_prog = (!tests.is_empty()).then(|| {
        let mut c = config.clone();
        c.test_files.clear();
        ras_compiler::analyze(sources.clone(), c)
    });
    let has_tests = !tests.is_empty();
    sources.extend(tests);
    let prog = ras_compiler::analyze(sources, config);
    let out = ras_compiler::build(build_prog.as_ref().unwrap_or(&prog), &opts);
    let js = (input["js"].as_bool().unwrap_or(false)
        && !prog.diagnostics.iter().any(|d| d.is_error()))
    .then(|| ras_compiler::build_js(&prog, "../../"));
    // The diagnostics of the build are those of the program it built.
    let built = build_prog.as_ref().unwrap_or(&prog);
    let mut diags: Vec<(&ras_compiler::Program, &ras_compiler::Diagnostic)> = prog
        .diagnostics
        .iter()
        .map(|d| (&prog, d))
        .chain(out.diagnostics.iter().map(|d| (built, d)))
        .chain(js.iter().flat_map(|j| &j.diagnostics).map(|d| (&prog, d)))
        .collect();
    let at = |p: &ras_compiler::Program, d: &ras_compiler::Diagnostic| {
        d.module
            .map(|m| (p.modules[m].rel_path.clone(), u32::from(d.span.start())))
    };
    diags.sort_by_key(|(p, d)| (at(p, d), d.severity));
    diags.dedup_by(|a, b| {
        a.1.code == b.1.code && a.1.message == b.1.message && at(a.0, a.1) == at(b.0, b.1)
    });
    let codepage = built.config.target.codepage;
    let mut diagnostics: Vec<Json> = diags
        .iter()
        .map(|(p, d)| {
            let Some(m) = d.module else {
                return json!({ "path": null, "line": 0, "col": 0, "endLine": 0, "endCol": 0,
                    "severity": severity(d.severity), "code": d.code, "message": d.message });
            };
            let md = &p.modules[m];
            let (line, col) = md.line_col_utf16(u32::from(d.span.start()));
            let (end_line, end_col) = md.line_col_utf16(u32::from(d.span.end()));
            json!({ "path": md.rel_path, "line": line, "col": col, "endLine": end_line, "endCol": end_col,
                "severity": severity(d.severity), "code": d.code, "message": d.message })
        })
        .collect();
    // The forms' designs (`ras check` checks them with the code compiled for them).
    diagnostics.extend(
        build_forms(&forms, codepage, &out.files)
            .problems
            .iter()
            .map(form_problem),
    );
    let vba: Vec<Json> = out
        .files
        .iter()
        .map(|f| json!({ "name": f.name, "fileName": f.file_name, "text": f.code.replace("\r\n", "\n"), "file": f.full_text }))
        .collect();
    let js_files = js.as_ref().map(|j| {
        j.files
            .iter()
            .map(|f| json!({ "path": f.rel_path, "text": f.text }))
            .collect::<Vec<_>>()
    });
    json!({
        "diagnostics": diagnostics,
        "vba": vba,
        "js": js_files,
        "debugDir": DEBUG_DIR,
        "entries": js.as_ref().map(|j| j.entries.clone()).unwrap_or_default(),
        "macros": out.macros,
        "tests": has_tests,
        "codepage": codepage,
        "error": null,
    })
}

/// The forms validated and generated with the code compiled for their code-behinds.
fn build_forms(
    forms: &[ras_form::workspace::LoadedForm],
    codepage: u16,
    files: &[ras_compiler::OutputFile],
) -> ras_form::workspace::Built {
    let compiled: std::collections::HashMap<String, String> = files
        .iter()
        .filter(|f| f.kind == VModuleKind::Form)
        .map(|f| (f.name.clone(), f.code.clone()))
        .collect();
    // Pictures are files of the project, which the playground does not have.
    ras_form::workspace::build(forms, Path::new("/"), codepage, &compiled)
}

/// A problem of a form as a diagnostic of [`compile`] (its position from 1 to from 0).
fn form_problem(p: &ras_form::workspace::Problem) -> Json {
    let (line, col) = p
        .position
        .map_or((0, 0), |(l, c)| (l.saturating_sub(1), c.saturating_sub(1)));
    let severity = match p.severity {
        ras_form::Severity::Error => "error",
        ras_form::Severity::Warning => "warning",
    };
    json!({ "path": p.rel_path, "line": line, "col": col, "endLine": line, "endCol": col,
        "severity": severity, "code": p.code, "message": p.message })
}

// The entries of `ras explain` (the `ras` program's own module, shared with it; formatted as
// part of the `ras` crate, so not here).
#[path = "../../ras/src/explain.rs"]
#[allow(dead_code)]
#[rustfmt::skip]
mod explain;

/// `ras explain <code>`: what a diagnostic means and how to fix it, or `undefined`.
#[wasm_bindgen]
pub fn explain(code: &str) -> Option<String> {
    explain::find(code).map(explain::text)
}

/// `ras fmt` of one file: the formatted text, or `undefined` when it does not parse.
#[wasm_bindgen]
pub fn format(text: &str) -> Option<String> {
    ras_syntax::format(text)
}

/// The bytes of a VBA file as `ras build` writes it: in the code page of `[target] codepage`
/// (932: Shift_JIS), which the VBE reads an imported `.bas` / `.cls` in.
#[wasm_bindgen(js_name = encodeVba)]
pub fn encode_vba(text: &str, codepage: u16) -> Vec<u8> {
    let enc = match codepage {
        932 => encoding_rs::SHIFT_JIS,
        936 => encoding_rs::GBK,
        949 => encoding_rs::EUC_KR,
        950 => encoding_rs::BIG5,
        874 => encoding_rs::WINDOWS_874,
        1250 => encoding_rs::WINDOWS_1250,
        1251 => encoding_rs::WINDOWS_1251,
        1253 => encoding_rs::WINDOWS_1253,
        1254 => encoding_rs::WINDOWS_1254,
        1255 => encoding_rs::WINDOWS_1255,
        1256 => encoding_rs::WINDOWS_1256,
        1257 => encoding_rs::WINDOWS_1257,
        1258 => encoding_rs::WINDOWS_1258,
        _ => encoding_rs::WINDOWS_1252,
    };
    enc.encode(text).0.into_owned()
}

/// A workbook with the project's VBA written into it ([`pack_book`]).
#[wasm_bindgen]
pub struct PackedBook {
    bytes: Vec<u8>,
    name: String,
    report: String,
}

#[wasm_bindgen]
impl PackedBook {
    #[wasm_bindgen(getter)]
    pub fn bytes(&self) -> Vec<u8> {
        self.bytes.clone()
    }
    /// The file name: `.xlsm` for a `.xlsx`.
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> String {
        self.name.clone()
    }
    /// JSON: `{ created, updated, removed, unchanged, warnings, replaced }` (`replaced`: modules
    /// of the workbook that RAS had not generated, replaced by the project's of the same name).
    #[wasm_bindgen(getter)]
    pub fn report(&self) -> String {
        self.report.clone()
    }
}

fn to_ovba_kind(k: &VModuleKind) -> ras_ovba::ModuleKind {
    match k {
        VModuleKind::Standard => ras_ovba::ModuleKind::Standard,
        VModuleKind::Class => ras_ovba::ModuleKind::Class,
        VModuleKind::Document => ras_ovba::ModuleKind::Document,
        VModuleKind::Form => ras_ovba::ModuleKind::Form,
    }
}

/// `ras pack --overwrite` in memory: writes the VBA of the project (`input` as for [`compile`])
/// into the workbook `bytes` named `name`, and returns the new workbook (a `.xlsx` becomes a
/// `.xlsm`). Modules that RAS generated earlier and the project no longer has are removed;
/// modules of the same name are replaced. The user forms of the project are written with their
/// code (forms with pictures cannot be: the playground has no picture files).
#[wasm_bindgen(js_name = packBook)]
pub fn pack_book(input: &str, bytes: &[u8], name: &str) -> Result<PackedBook, JsError> {
    let mut i = read_input(input).map_err(err)?;
    let existing = ras_ovba::unpack_bytes(bytes.to_vec()).map_err(|e| err(e.to_string()))?;
    i.config.sheet_code_names = existing
        .code_names
        .sheets
        .iter()
        .map(|s| s.code_name.clone())
        .collect();
    i.config.test_files.clear();
    if let Ok(Some(sheets)) = ras_ovba::workbook_sheet_controls_bytes(bytes.to_vec()) {
        i.config.sheet_controls = sheet_decls(sheets);
        i.config.sheet_controls_known = true;
    }
    let codepage = i.config.target.codepage;
    let prog = ras_compiler::analyze(i.sources, i.config);
    let out = ras_compiler::build(&prog, &BuildOptions::default());
    let user_forms = build_forms(&i.forms, codepage, &out.files);
    if user_forms.errors() > 0 {
        let text: Vec<String> = user_forms
            .problems
            .iter()
            .filter(|p| p.severity == ras_form::Severity::Error)
            .map(ToString::to_string)
            .collect();
        return Err(err(format!(
            "ユーザーフォームにエラーがあるため書き込めません:\n{}",
            text.join("\n")
        )));
    }
    let errors: Vec<&ras_compiler::Diagnostic> = prog
        .diagnostics
        .iter()
        .chain(&out.diagnostics)
        .filter(|d| d.is_error())
        .collect();
    if !errors.is_empty() {
        return Err(err(format!(
            "エラーがあるため書き込めません:\n{}",
            ras_compiler::render_diagnostics(&prog, &prog.diagnostics)
        )));
    }
    let generated: Vec<&str> = out
        .files
        .iter()
        .filter(|f| f.kind != VModuleKind::Form && f.kind != VModuleKind::Document)
        .map(|f| f.name.as_str())
        .collect();
    let modules: Vec<&ras_ovba::VbaModule> =
        existing.project.iter().flat_map(|p| &p.modules).collect();
    let mut remove = Vec::new();
    let mut replaced = Vec::new();
    for m in &modules {
        if !matches!(
            m.kind,
            ras_ovba::ModuleKind::Standard | ras_ovba::ModuleKind::Class
        ) {
            continue;
        }
        let ours = ras_compiler::emit::has_drifted(&m.source).is_some();
        let kept = generated.iter().any(|g| g.eq_ignore_ascii_case(&m.name));
        if ours && !kept {
            remove.push(m.name.clone());
        } else if !ours && kept {
            replaced.push(m.name.clone());
        }
    }
    let workbook_name = existing.code_names.workbook.clone();
    let mut document_code = Vec::new();
    let mut module_list = Vec::new();
    for f in &out.files {
        match f.kind {
            VModuleKind::Form => {}
            VModuleKind::Document => document_code.push(ras_ovba::DocumentCode {
                code_name: if f.name.eq_ignore_ascii_case("ThisWorkbook") {
                    workbook_name.clone()
                } else {
                    f.name.clone()
                },
                code: f.code.clone(),
            }),
            _ => module_list.push(ras_ovba::VbaModule {
                name: f.name.clone(),
                kind: to_ovba_kind(&f.kind),
                source: f.stream_text(),
            }),
        }
    }
    let book = PathBuf::from(name);
    let output = match book.extension().and_then(|e| e.to_str()) {
        Some(e) if e.eq_ignore_ascii_case("xlsx") => book.with_extension("xlsm"),
        Some(e) if e.eq_ignore_ascii_case("xltx") => book.with_extension("xltm"),
        _ => book.clone(),
    };
    let req = ras_ovba::PackRequest {
        book,
        output: Some(output.clone()),
        convert: true,
        codepage,
        project_name: None,
        modules: module_list,
        document_code,
        remove_modules: remove,
        rename_modules: Vec::new(),
        prune_documents: false,
        forms: user_forms.forms,
        backup_dir: None,
        backup_generations: 0,
    };
    let (bytes, rep) =
        ras_ovba::pack_bytes(&req, bytes.to_vec()).map_err(|e| err(e.to_string()))?;
    Ok(PackedBook {
        bytes,
        name: output.to_string_lossy().into_owned(),
        report: json!({
            "created": rep.created,
            "updated": rep.updated,
            "removed": rep.removed,
            "unchanged": rep.unchanged,
            "warnings": rep.warnings,
            "replaced": replaced,
        })
        .to_string(),
    })
}

/// The ActiveX controls of the sheets as the compiler takes them (one entry per sheet with any).
fn sheet_decls(sheets: Vec<ras_ovba::SheetControls>) -> Vec<ras_compiler::FormDecl> {
    sheets
        .into_iter()
        .filter(|s| !s.controls.is_empty())
        .map(|s| ras_compiler::FormDecl {
            name: s.code_name,
            controls: s
                .controls
                .into_iter()
                .map(|c| ras_compiler::FormControl {
                    name: c.name,
                    control_type: c.class,
                })
                .collect(),
        })
        .collect()
}

/// The ActiveX controls on the sheets of a workbook (spec §12.2.1), as `sheetControls` of the
/// input of [`compile`]: `[{ name: <code name>, controls: [{ name, type }] }]`, or `null` when
/// they cannot be read (an `.xlsb`, bytes that are not a workbook).
#[wasm_bindgen(js_name = sheetControls)]
pub fn sheet_controls(bytes: &[u8]) -> String {
    match ras_ovba::workbook_sheet_controls_bytes(bytes.to_vec()) {
        Ok(Some(sheets)) => Json::Array(
            sheet_decls(sheets)
                .into_iter()
                .map(|s| {
                    json!({ "name": s.name, "controls": s.controls.iter()
                        .map(|c| json!({ "name": c.name, "type": c.control_type }))
                        .collect::<Vec<_>>() })
                })
                .collect(),
        )
        .to_string(),
        _ => "null".to_string(),
    }
}

/// A new form `name` (`ras form new`): the text of its `forms/<name>.form.json`.
#[wasm_bindgen(js_name = newFormSpec)]
pub fn new_form_spec(name: &str) -> String {
    ras_form::FormSpec::new_form(name).to_json()
}

fn err(e: String) -> JsError {
    JsError::new(&e)
}

/// The workbook of the stand-in for Excel: the `Book` of the `ras_book` addon, for the browser.
/// Files are bytes here (`fromBytes`, `toBytes`, `saveBytes`); the playground keeps them.
#[wasm_bindgen]
pub struct Book {
    inner: book::Book,
    saved_path: String,
}

#[wasm_bindgen]
impl Book {
    #[wasm_bindgen(js_name = fromBytes)]
    pub fn from_bytes(data: &[u8]) -> Result<Book, JsError> {
        Ok(Book {
            inner: book::Book::open_bytes(data).map_err(err)?,
            saved_path: String::new(),
        })
    }

    pub fn create() -> Book {
        Book {
            inner: book::Book::new_book(),
            saved_path: String::new(),
        }
    }

    #[wasm_bindgen(js_name = toBytes)]
    pub fn to_bytes(&mut self) -> Result<Vec<u8>, JsError> {
        self.inner.to_bytes().map_err(err)
    }

    /// The bytes `save(path)` writes; `savedPath` is then the path they go to (`.xlsx` for a
    /// workbook without macros saved as `.xlsm`).
    #[wasm_bindgen(js_name = saveBytes)]
    pub fn save_bytes(&mut self, path: &str) -> Result<Vec<u8>, JsError> {
        let (written, bytes) = self.inner.save_bytes(Path::new(path)).map_err(err)?;
        self.saved_path = written.to_string_lossy().into_owned();
        Ok(bytes)
    }

    #[wasm_bindgen(getter, js_name = savedPath)]
    pub fn saved_path(&self) -> String {
        self.saved_path.clone()
    }

    /// A method of the addon's `Book` (`ras_workbook::js::METHODS`): `args` is the JSON array of
    /// its arguments; returns the JSON of its result.
    pub fn call(&mut self, method: &str, args: &str) -> Result<String, JsError> {
        let args: Vec<Json> = serde_json::from_str(args).map_err(|e| err(e.to_string()))?;
        js::call(&mut self.inner, method, &args)
            .map(|v| v.to_string())
            .map_err(err)
    }

    #[wasm_bindgen(js_name = isFunction)]
    pub fn is_function(name: &str) -> bool {
        book::Book::is_function(name)
    }

    /// The names of the methods `call` takes, as a JSON array.
    pub fn methods() -> String {
        json!(js::METHODS).to_string()
    }
}

fn open(bytes: &[u8]) -> Result<book::Book, JsError> {
    book::Book::open_bytes(bytes).map_err(err)
}

/// The cells and structure that differ between two workbooks, as `ras run` prints them.
#[wasm_bindgen(js_name = diffText)]
pub fn diff_text(before: &[u8], after: &[u8], max_cells: usize) -> Result<String, JsError> {
    let d = diff::diff(&mut open(before)?, &mut open(after)?).map_err(err)?;
    Ok(if d.is_empty() {
        String::new()
    } else {
        diff::text(&d, max_cells)
    })
}

/// The page of `ras book view` for a workbook: with `before`, what differs from it is marked.
/// `editable`: the page of `ras book view --editable`, whose edits are posted to the page that
/// holds it (its parent window), which makes them with [`edit_book`] and draws the page again.
#[wasm_bindgen(js_name = viewPage)]
pub fn view_page(
    bytes: &[u8],
    name: &str,
    before: Option<Vec<u8>>,
    before_name: &str,
    editable: bool,
) -> Result<String, JsError> {
    // (The page that edits the workbook: RAND, NOW and the like as the file has them, as
    // `ras book view --editable` shows them.)
    let mut b = book::Book::open_bytes_as(bytes, !editable).map_err(err)?;
    let changes = match before {
        Some(old) => Some(diff::diff(&mut open(&old)?, &mut b).map_err(err)?),
        None => None,
    };
    let mut data = view::data(
        &mut b,
        name,
        Some(bytes),
        Some((book::VIEW_ROWS, book::VIEW_COLS)),
        changes.as_ref().map(|d| (before_name, d)),
    )
    .map_err(err)?;
    if editable {
        data["editable"] = true.into();
        data["host"] = "parent".into();
        // For the completion of formulas typed into the cells.
        data["functions"] = view::functions();
        // For the menu and the dialogs of conditional formats.
        data["iconSets"] = view::icon_sets();
    }
    Ok(view::page(&data))
}

/// `ras book edit`: the bytes of the workbook with the edits the page posted (JSON: one edit or
/// an array of them; those that name no sheet are of the first sheet).
#[wasm_bindgen(js_name = editBook)]
pub fn edit_book(bytes: &[u8], ops: &str) -> Result<Vec<u8>, JsError> {
    let first = open(bytes)?.sheet_name(0).map(str::to_string);
    let json: Json =
        serde_json::from_str(ops).map_err(|e| err(format!("編集の JSON を読み取れません: {e}")))?;
    let ops = ras_workbook::patch::parse_ops(&json, first.as_deref()).map_err(err)?;
    ras_workbook::patch::apply(bytes, &ops).map_err(err)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compiles_a_project_in_memory() {
        let out: Json = serde_json::from_str(&compile(
            &json!({
                "files": [{ "path": "src/main.ras", "text": "export function hello() {\n    debug.print(\"hi\");\n}\n" }],
                "toml": "[project]\nname = \"Playground\"\n",
                "sheets": ["Sheet1"],
                "js": true,
            })
            .to_string(),
        ))
        .unwrap();
        assert_eq!(out["error"], Json::Null);
        let errors: Vec<&Json> = out["diagnostics"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|d| d["severity"] == "error")
            .collect();
        assert!(errors.is_empty(), "{errors:?}");
        let vba = out["vba"].as_array().unwrap();
        assert!(
            vba.iter()
                .any(|f| f["text"].as_str().unwrap().contains("Public Sub hello()")),
            "{vba:?}"
        );
        let js = out["js"].as_array().unwrap();
        assert!(js.iter().any(|f| f["path"] == "ras-main.js"));
        assert!(
            out["entries"]
                .as_array()
                .unwrap()
                .contains(&json!("Main.hello"))
        );
    }

    #[test]
    fn packs_the_vba_into_a_workbook() {
        let xlsx = book::Book::new_book().to_bytes().unwrap();
        let input = json!({
            "files": [
                { "path": "src/main.ras", "text": "export function hello() {\n    msgBox(\"こんにちは\");\n}\n" },
                { "path": "objects/Sheet1.ras", "text": "@on(\"Worksheet.Change\")\nfunction onChange(target: Excel.Range) {\n    debug.print(target.address);\n}\n" },
                { "path": "test/main.test.ras", "text": "test(\"t\", () => {\n    expect(1).toBe(1);\n});\n" },
            ],
            "toml": "[project]\nname = \"P\"\n",
        })
        .to_string();
        let packed = pack_book(&input, &xlsx, "Book.xlsx").unwrap_or_else(|_| panic!("pack"));
        assert_eq!(packed.name(), "Book.xlsm");
        let report: Json = serde_json::from_str(&packed.report()).unwrap();
        assert!(
            report["created"]
                .as_array()
                .unwrap()
                .contains(&json!("Main")),
            "{report}"
        );
        let unpacked = ras_ovba::unpack_bytes(packed.bytes()).unwrap();
        let project = unpacked.project.expect("a VBA project");
        let main = project
            .modules
            .iter()
            .find(|m| m.name == "Main")
            .expect("Main");
        assert!(
            main.source.contains("Public Sub hello()"),
            "{}",
            main.source
        );
        assert!(main.source.contains("こんにちは"));
        let sheet = project
            .modules
            .iter()
            .find(|m| m.name == "Sheet1")
            .expect("Sheet1");
        assert!(
            sheet.source.contains("Worksheet_Change"),
            "{}",
            sheet.source
        );
        // The tests are not part of the workbook.
        assert!(!project.modules.iter().any(|m| m.name.starts_with("T_")));

        // Packed again without `hello`: the module RAS generated goes.
        let input = json!({
            "files": [{ "path": "src/other.ras", "text": "export function f() {\n}\n" }],
            "toml": "[project]\nname = \"P\"\n",
        })
        .to_string();
        let again =
            pack_book(&input, &packed.bytes(), "Book.xlsm").unwrap_or_else(|_| panic!("pack"));
        let report: Json = serde_json::from_str(&again.report()).unwrap();
        assert!(
            report["removed"]
                .as_array()
                .unwrap()
                .contains(&json!("Main")),
            "{report}"
        );
    }

    #[test]
    fn reports_errors_with_positions() {
        let out: Json = serde_json::from_str(&compile(
            &json!({
                "files": [{ "path": "src/main.ras", "text": "export function f() {\n    let x: int = \"a\";\n}\n" }],
                "toml": "[project]\nname = \"P\"\n",
            })
            .to_string(),
        ))
        .unwrap();
        let d = out["diagnostics"]
            .as_array()
            .unwrap()
            .iter()
            .find(|d| d["severity"] == "error")
            .cloned()
            .expect("an error");
        assert_eq!(d["path"], "src/main.ras");
        assert_eq!(d["line"], 1);
        assert!(out["vba"].as_array().unwrap().is_empty());
    }
}
