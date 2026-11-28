//! Live sync (spec §14.1–14.8): writes the generated VBA into a workbook that is open in Excel.
//!
//! * [`plan`] — what a sync does to a VBA project: the rules, tested without Excel.
//! * `excel` — the open workbook, reached over COM (Windows only).
//!
//! The code is written through the VBE object model (`CodeModule`) from this process; a user
//! form whose design changed is imported from a generated `.frm` / `.frx` pair. The
//! shared-memory transport and the RASBridge add-in of spec §14.4 are an optimisation of the
//! same rules and are not part of this crate.

use std::fmt;
use std::path::{Path, PathBuf};

pub mod plan;
pub mod sheets;

#[cfg(windows)]
mod disp;
#[cfg(windows)]
mod excel;

pub use plan::{Forms, Report, SyncOptions};
use ras_compiler::config::SyncCfg;
use ras_compiler::emit::OutputFile;
pub use ras_form::import::Prepared;

/// The options of a sync of the project at `root` that follows the take-over `prepared`
/// (spec §14.7.1).
pub fn options(root: &Path, prepared: &Prepared, overwrite: bool) -> SyncOptions {
    SyncOptions {
        overwrite,
        adopt: prepared
            .adopt
            .iter()
            .map(|(name, _)| name.clone())
            .collect(),
        pending: prepared
            .pending
            .iter()
            .map(|p| (p.name.clone(), p.generated.clone()))
            .collect(),
        written: written_modules(root),
    }
}

/// Where the code a sync wrote into each module is kept (`build/sync/<Module>.vba`): what a
/// module is compared with to find the code added to it in the VBE (spec §14.7.3).
pub fn written_dir(root: &Path) -> PathBuf {
    root.join("build").join("sync")
}

/// The modules a sync of this project has written code into.
pub fn written_modules(root: &Path) -> Vec<String> {
    let Ok(entries) = std::fs::read_dir(written_dir(root)) else {
        return Vec::new();
    };
    entries
        .flatten()
        .filter_map(|e| {
            e.file_name()
                .to_string_lossy()
                .strip_suffix(".vba")
                .map(str::to_string)
        })
        .collect()
}

/// Keeps the code a sync (or `ras pack`) wrote: that of the standard and document modules
/// among `files`, but for those in `not_written` (drifted or pending modules, left as they
/// were). Copies of modules that are not generated any more are removed.
pub fn save_written(
    root: &Path,
    files: &[OutputFile],
    not_written: &[String],
) -> Result<(), String> {
    use ras_compiler::vba::VModuleKind;
    let dir = written_dir(root);
    std::fs::create_dir_all(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    let same = |a: &str, b: &str| a.to_lowercase() == b.to_lowercase();
    let kept = |name: &str| not_written.iter().any(|n| same(n, name));
    let modules: Vec<&OutputFile> = files
        .iter()
        .filter(|f| matches!(f.kind, VModuleKind::Standard | VModuleKind::Document))
        .collect();
    for name in written_modules(root) {
        if !modules.iter().any(|f| same(&f.name, &name)) && !kept(&name) {
            let _ = std::fs::remove_file(dir.join(format!("{name}.vba")));
        }
    }
    for f in modules.iter().filter(|f| !kept(&f.name)) {
        let path = dir.join(format!("{}.vba", f.name));
        if std::fs::read_to_string(&path).is_ok_and(|old| old == f.code) {
            continue;
        }
        std::fs::write(&path, &f.code).map_err(|e| format!("{}: {e}", path.display()))?;
    }
    Ok(())
}

/// The code RAS wrote into the modules of the workbook the last time, for
/// [`ras_form::import::Book::written`]: the copy [`save_written`] kept, or — where there is
/// none (a sync from another computer, `build/` removed) — what the project generates now.
pub struct Written {
    root: PathBuf,
    /// The modules the project generates now (built when a copy is missing).
    generated: Option<Vec<OutputFile>>,
}

impl Written {
    /// For the project at `root`.
    pub fn new(root: &Path) -> Written {
        Written {
            root: root.to_path_buf(),
            generated: None,
        }
    }

    /// The code written into the module `name`.
    pub fn code(&mut self, name: &str) -> Option<String> {
        let same = |a: &str, b: &str| a.to_lowercase() == b.to_lowercase();
        if let Some(saved) = written_modules(&self.root)
            .into_iter()
            .find(|n| same(n, name))
        {
            if let Ok(text) =
                std::fs::read_to_string(written_dir(&self.root).join(format!("{saved}.vba")))
            {
                return Some(text);
            }
        }
        if self.generated.is_none() {
            self.generated = Some(generate(&self.root).unwrap_or_default());
        }
        self.generated
            .as_ref()?
            .iter()
            .find(|f| same(&f.name, name))
            .map(|f| f.code.clone())
    }
}

/// The modules the project at `root` generates as it is now; `None` when it does not compile.
fn generate(root: &Path) -> Option<Vec<OutputFile>> {
    let prog = analyze(root)?;
    if prog.diagnostics.iter().any(|d| d.is_error()) {
        return None;
    }
    let out = ras_compiler::build(
        &prog,
        &ras_compiler::BuildOptions {
            no_attributes: true,
            ..Default::default()
        },
    );
    if out.diagnostics.iter().any(|d| d.is_error()) {
        return None;
    }
    Some(out.files)
}

/// The project at `root`, analysed with what its workbook says.
pub(crate) fn analyze(root: &Path) -> Option<ras_compiler::Program> {
    let (mut config, sources) = ras_compiler::load_project(root).ok()?;
    if let Some(book) = config
        .project
        .workbook
        .as_ref()
        .map(|w| root.join(w))
        .filter(|b| b.is_file())
    {
        let _ = read_workbook(&mut config, root, &book);
    }
    Some(ras_compiler::analyze(sources, config))
}

/// After a take-over that appended code added in the VBE to RAS files (spec §14.7.3): a file
/// that does not compile with what was appended gets its text back, and the module stays as
/// drift.
pub fn settle_appended(root: &Path, prepared: &mut Prepared) -> Result<(), String> {
    if prepared.appended.is_empty() {
        return Ok(());
    }
    let Some(prog) = analyze(root) else {
        return Ok(());
    };
    let failed: Vec<String> = prog
        .diagnostics
        .iter()
        .filter(|d| d.is_error())
        .filter_map(|d| d.module)
        .map(|m| prog.modules[m].rel_path.clone())
        .collect();
    ras_form::import::revert_appended(root, prepared, &failed)
}

/// After a take-over: corrects, in the files it converted from VBA, what the checker knows
/// how to correct so that the code does what the VBA did (`Program::repairs`: a number used as
/// a condition, the default member of a `Range`, a member written with another casing …), and
/// applies the mechanical corrections of `ras check --fix`. The types are only known once every
/// converted file is in the project, which is why this follows the conversion. `book` is the
/// workbook the project is checked against (the project's own when `None`).
pub fn repair_imported(
    root: &Path,
    prepared: &mut Prepared,
    book: Option<&Path>,
) -> Result<(), String> {
    let converted: Vec<String> = prepared
        .imported
        .iter()
        .map(|i| i.path.replace('\\', "/").to_lowercase())
        .collect();
    if converted.is_empty() {
        return Ok(());
    }
    let analyze = || -> Option<ras_compiler::Program> {
        let (mut config, sources) = ras_compiler::load_project(root).ok()?;
        let book = book
            .map(Path::to_path_buf)
            .or_else(|| config.project.workbook.as_ref().map(|w| root.join(w)))
            .filter(|b| b.is_file());
        if let Some(book) = book {
            let _ = read_workbook(&mut config, root, &book);
        }
        Some(ras_compiler::analyze(sources, config))
    };
    let is_converted =
        |rel_path: &str| converted.contains(&rel_path.replace('\\', "/").to_lowercase());
    // A procedure whose conversion still does not compile stays VBA (spec §14.7.4); the
    // code that calls it is checked again.
    for _ in 0..6 {
        // A correction can uncover the next one (`!n` → `~n` → `~n != 0`).
        let mut prog = None;
        for _ in 0..8 {
            let Some(p) = analyze() else {
                return Ok(());
            };
            let mut changed = false;
            for md in p.modules.iter().filter(|md| is_converted(&md.rel_path)) {
                let fixes: Vec<&ras_compiler::Fix> = p
                    .repairs
                    .iter()
                    .chain(p.fixes.iter())
                    .filter(|f| f.module == md.id)
                    .collect();
                if fixes.is_empty() {
                    continue;
                }
                let fixed = ras_compiler::apply_fixes(&md.text, &fixes);
                if fixed != *md.text {
                    std::fs::write(&md.path, fixed)
                        .map_err(|e| format!("{}: {e}", md.path.display()))?;
                    changed = true;
                }
            }
            prog = Some(p);
            if !changed {
                break;
            }
        }
        let Some(prog) = prog else {
            return Ok(());
        };
        // The procedures the remaining errors are in.
        let mut failed: Vec<(String, Vec<String>)> = Vec::new();
        for d in prog.diagnostics.iter().filter(|d| d.is_error()) {
            let Some(md) = d
                .module
                .map(|m| &prog.modules[m])
                .filter(|md| is_converted(&md.rel_path))
            else {
                continue;
            };
            let offset = (u32::from(d.span.start()) as usize).min(md.text.len());
            let line = md.text.as_bytes()[..offset]
                .iter()
                .filter(|b| **b == b'\n')
                .count()
                + 1;
            let path = md.rel_path.replace('\\', "/");
            let Some(name) = ras_form::import::procedure_at(prepared, &path, line) else {
                continue;
            };
            if std::env::var_os("RAS_IMPORT_TRACE").is_some() {
                eprintln!("[keep] {path}:{line} {name}: {} {}", d.code, d.message);
                eprintln!(
                    "       {}",
                    md.text.lines().nth(line - 1).unwrap_or("").trim()
                );
            }
            match failed.iter_mut().find(|(p, _)| *p == path) {
                Some((_, names)) if !names.contains(&name) => names.push(name),
                Some(_) => {}
                None => failed.push((path, vec![name])),
            }
        }
        let mut kept = false;
        for (path, names) in &failed {
            kept |= ras_form::import::keep_procedures(root, prepared, path, names)?;
        }
        if !kept {
            break;
        }
    }
    Ok(())
}
/// The file the live sync keeps the code names of the open workbook's sheets in: a sheet
/// added in Excel is not in the workbook file until it is saved (spec §12.2).
fn live_code_names_path(root: &Path) -> PathBuf {
    written_dir(root).join("code-names.txt")
}

/// Records the code names of the sheets of the open workbook (see [`read_workbook`]).
pub fn save_live_code_names(root: &Path, names: &[String]) {
    let path = live_code_names_path(root);
    if std::fs::create_dir_all(written_dir(root)).is_ok() {
        let text: String = names.iter().map(|n| format!("{n}\n")).collect();
        if std::fs::read_to_string(&path).is_ok_and(|old| old == text) {
            // Kept newer than the workbook file all the same.
            let _ = std::fs::File::options()
                .append(true)
                .open(&path)
                .and_then(|f| f.set_modified(std::time::SystemTime::now()));
            return;
        }
        let _ = std::fs::write(&path, text);
    }
}

/// The ActiveX controls on the sheets of the workbook file `book` as the compiler takes them
/// (spec §12.2.1): one entry per sheet that has any. `None` when they cannot be read (an
/// `.xlsb`, a file that is not a workbook).
pub fn sheet_controls(book: &Path) -> Option<Vec<ras_compiler::FormDecl>> {
    let sheets = ras_ovba::workbook_sheet_controls(book).ok()??;
    Some(
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
                resizable: false,
            })
            .collect(),
    )
}

/// Fills in what the compiler takes from the workbook file `book` of the project at `root`:
/// the code names of its sheets (spec §12.2) and the controls on them (spec §12.2.1). The
/// code names the live sync saw in the open workbook after the file was saved last are taken
/// instead of the file's (a sheet added in Excel and not saved yet).
pub fn read_workbook(
    config: &mut ras_compiler::Config,
    root: &Path,
    book: &Path,
) -> Result<(), String> {
    let names = ras_ovba::workbook_code_names(book).map_err(|e| e.to_string())?;
    config.sheet_code_names = names.sheets.into_iter().map(|s| s.code_name).collect();
    let live = live_code_names_path(root);
    let modified = |p: &Path| std::fs::metadata(p).and_then(|m| m.modified()).ok();
    if let (Some(seen), Some(saved)) = (modified(&live), modified(book)) {
        if seen > saved {
            if let Ok(text) = std::fs::read_to_string(&live) {
                let names: Vec<String> = text
                    .lines()
                    .map(str::trim)
                    .filter(|l| !l.is_empty())
                    .map(str::to_string)
                    .collect();
                if !names.is_empty() {
                    config.sheet_code_names = names;
                }
            }
        }
    }
    if let Some(controls) = sheet_controls(book) {
        config.sheet_controls = controls;
        config.sheet_controls_known = true;
    }
    Ok(())
}

/// A component of the project as [`ras_form::import`] takes it; `None` for what it does not take.
pub fn book_module(c: &plan::Component) -> Option<ras_form::import::BookModule> {
    let kind = match c.kind {
        plan::Kind::Standard => ras_ovba::ModuleKind::Standard,
        plan::Kind::Class => ras_ovba::ModuleKind::Class,
        plan::Kind::Form => ras_ovba::ModuleKind::Form,
        plan::Kind::Document => ras_ovba::ModuleKind::Document,
        plan::Kind::Other => return None,
    };
    Some(ras_form::import::BookModule {
        name: c.name.clone(),
        kind,
        code: c.code.clone(),
    })
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SyncError {
    /// Live sync needs Excel, which only exists on Windows.
    Unsupported,
    /// No running Excel has the workbook open.
    NotOpen(PathBuf),
    /// Excel cannot take the change right now (cell being edited, dialog open, macro running). Worth retrying.
    Busy(String),
    Failed(String),
}

impl fmt::Display for SyncError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SyncError::Unsupported => {
                write!(f, "Excel とのライブ同期は Windows でのみ利用できます")
            }
            SyncError::NotOpen(book) => write!(f, "{} は Excel で開かれていません", book.display()),
            SyncError::Busy(why) => write!(
                f,
                "Excel が同期を受け付けられません（{why}）。Excel 側の操作が終わると同期します"
            ),
            SyncError::Failed(message) => write!(f, "{message}"),
        }
    }
}

impl std::error::Error for SyncError {}

/// The connection to one workbook. All calls must come from the thread that created it (COM apartment).
pub struct Session {
    book: PathBuf,
    #[cfg(windows)]
    live: Option<excel::Book>,
}

impl Session {
    /// `book` is the full path of the workbook (`project.workbook` of `ras.toml`, resolved).
    pub fn new(book: &Path) -> Session {
        let book = std::path::absolute(book).unwrap_or_else(|_| book.to_path_buf());
        Session {
            book,
            #[cfg(windows)]
            live: None,
        }
    }

    pub fn book(&self) -> &Path {
        &self.book
    }
}

#[cfg(windows)]
impl Session {
    /// Finds the workbook in the running Excel instances (spec §14.3), unless it is already connected.
    pub fn connect(&mut self) -> Result<(), SyncError> {
        if self.live.as_ref().is_some_and(|b| b.is_alive()) {
            return Ok(());
        }
        self.live = excel::find(&self.book)?;
        match self.live {
            Some(_) => Ok(()),
            None => Err(SyncError::NotOpen(self.book.clone())),
        }
    }

    /// Writes the generated modules and the user forms into the open workbook.
    pub fn apply(
        &mut self,
        files: &[OutputFile],
        forms: Forms,
        opts: &SyncOptions,
    ) -> Result<Report, SyncError> {
        self.connect()?;
        let mut project = self.live.as_ref().expect("connected").project()?;
        plan::apply(&mut project, files, forms, opts)
    }

    /// Takes over the hand-written VBA of the open workbook — standard modules, the code of
    /// sheets and of `ThisWorkbook`, forms made in the VBE — into the project at `root`
    /// (spec §14.7.1) when `[sync] pull-unmanaged` is on, then creates a RAS file for each
    /// sheet (and `ThisWorkbook`) that has none when `[sync] sheet-files` is on (spec §14.7.2).
    /// `codepage` is the project's `target.codepage`.
    pub fn take_over(
        &mut self,
        root: &Path,
        codepage: u16,
        sync: &SyncCfg,
    ) -> Result<Prepared, SyncError> {
        use plan::VbProject;
        self.connect()?;
        let book = self.live.as_ref().expect("connected");
        let code_name = book.code_name()?;
        let mut prepared = Prepared::default();
        if sync.pull_unmanaged {
            let mut project = book.project()?;
            let modules: Vec<ras_form::import::BookModule> = project
                .components()?
                .iter()
                .filter_map(book_module)
                .collect();
            // The controls on the sheets, as the workbook was saved last.
            let sheet_controls: Vec<ras_compiler::FormDecl> =
                sheet_controls(&self.book).unwrap_or_default();
            let mut failure = None;
            let mut form = |name: &str| {
                project.export_form(name, codepage).map_err(|e| {
                    let message = e.to_string();
                    if matches!(e, SyncError::Busy(_)) {
                        failure = Some(e);
                    }
                    message
                })
            };
            let mut written = Written::new(root);
            let mut written = |name: &str| written.code(name);
            let mut taken = ras_form::import::Book {
                modules: &modules,
                workbook_code_name: &code_name,
                sheet_controls: &sheet_controls,
                form: &mut form,
                written: &mut written,
            };
            prepared =
                ras_form::import::prepare(root, &mut taken, codepage).map_err(SyncError::Failed)?;
            if let Some(busy) = failure {
                return Err(busy);
            }
            repair_imported(root, &mut prepared, Some(&self.book)).map_err(SyncError::Failed)?;
        }
        let sheets = book.sheets()?;
        // What the compiler takes for the sheets of the workbook until it is saved.
        save_live_code_names(
            root,
            &sheets
                .iter()
                .map(|s| s.code_name.clone())
                .collect::<Vec<_>>(),
        );
        if sync.sheet_files {
            // A copied sheet gets a copy of its source's file: the modules tell which it is.
            let modules: Vec<ras_form::import::BookModule> = if sync.pull_unmanaged {
                book.project()?
                    .components()?
                    .iter()
                    .filter_map(book_module)
                    .collect()
            } else {
                Vec::new()
            };
            let sheet_controls: Vec<ras_compiler::FormDecl> =
                sheet_controls(&self.book).unwrap_or_default();
            let mut form = |_: &str| Err(String::new());
            let mut written = Written::new(root);
            let mut written = |name: &str| written.code(name);
            let mut taken = ras_form::import::Book {
                modules: &modules,
                workbook_code_name: &code_name,
                sheet_controls: &sheet_controls,
                form: &mut form,
                written: &mut written,
            };
            ras_form::import::sheet_files(root, &sheets, &mut taken, &mut prepared)
                .map_err(SyncError::Failed)?;
        }
        settle_appended(root, &mut prepared).map_err(SyncError::Failed)?;
        Ok(prepared)
    }

    /// The worksheets of the open workbook (code name and tab name), in workbook order.
    pub fn sheets(&mut self) -> Result<Vec<ras_form::import::Sheet>, SyncError> {
        self.connect()?;
        self.live.as_ref().expect("connected").sheets()
    }

    /// Adds, copies, removes or renames a worksheet of the open workbook (spec §14.16).
    pub fn edit_sheet(
        &mut self,
        edit: &ras_ovba::SheetEdit,
    ) -> Result<ras_ovba::SheetEditReport, SyncError> {
        self.connect()?;
        self.live.as_ref().expect("connected").edit_sheet(edit)
    }

    /// Runs the macro `Module.proc` in the workbook and waits for it to end.
    pub fn run_macro(&mut self, name: &str) -> Result<(), SyncError> {
        self.connect()?;
        self.live.as_ref().expect("connected").run(name)
    }

    /// Shows the user form `name` in Excel (with Excel in front) and waits until a modal form is closed.
    pub fn show_form(&mut self, name: &str) -> Result<(), SyncError> {
        self.connect()?;
        self.live.as_ref().expect("connected").show_form(name)
    }

    /// Opens the workbook in Excel when it is not open yet, and brings Excel to the front.
    pub fn open(&mut self) -> Result<(), SyncError> {
        if self.connect().is_err() {
            self.live = Some(excel::open(&self.book)?);
        }
        self.live.as_ref().expect("connected").activate();
        Ok(())
    }
}

#[cfg(not(windows))]
impl Session {
    pub fn connect(&mut self) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn apply(
        &mut self,
        _files: &[OutputFile],
        _forms: Forms,
        _opts: &SyncOptions,
    ) -> Result<Report, SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn open(&mut self) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn sheets(&mut self) -> Result<Vec<ras_form::import::Sheet>, SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn edit_sheet(
        &mut self,
        _edit: &ras_ovba::SheetEdit,
    ) -> Result<ras_ovba::SheetEditReport, SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn run_macro(&mut self, _name: &str) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn show_form(&mut self, _name: &str) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn take_over(
        &mut self,
        _root: &Path,
        _codepage: u16,
        _sync: &SyncCfg,
    ) -> Result<Prepared, SyncError> {
        Err(SyncError::Unsupported)
    }
}
