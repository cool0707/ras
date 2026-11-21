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

#[cfg(windows)]
mod disp;
#[cfg(windows)]
mod excel;

use ras_compiler::config::SyncCfg;
use ras_compiler::emit::OutputFile;
pub use plan::{Forms, Report, SyncOptions};
pub use ras_form::import::Prepared;

/// The options of a sync that follows the take-over `prepared` (spec §14.7.1).
pub fn options(prepared: &Prepared, overwrite: bool) -> SyncOptions {
    SyncOptions {
        overwrite,
        adopt: prepared.adopt.iter().map(|(name, _)| name.clone()).collect(),
        pending: prepared.pending.iter().map(|p| (p.name.clone(), p.generated.clone())).collect(),
    }
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
    Some(ras_form::import::BookModule { name: c.name.clone(), kind, code: c.code.clone() })
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
            SyncError::Unsupported => write!(f, "Excel とのライブ同期は Windows でのみ利用できます"),
            SyncError::NotOpen(book) => write!(f, "{} は Excel で開かれていません", book.display()),
            SyncError::Busy(why) => write!(f, "Excel が同期を受け付けられません（{why}）。Excel 側の操作が終わると同期します"),
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
    pub fn apply(&mut self, files: &[OutputFile], forms: Forms, opts: &SyncOptions) -> Result<Report, SyncError> {
        self.connect()?;
        let mut project = self.live.as_ref().expect("connected").project()?;
        plan::apply(&mut project, files, forms, opts)
    }

    /// Takes over the hand-written VBA of the open workbook — standard modules, the code of
    /// sheets and of `ThisWorkbook`, forms made in the VBE — into the project at `root`
    /// (spec §14.7.1) when `[sync] pull-unmanaged` is on, then creates a RAS file for each
    /// sheet (and `ThisWorkbook`) that has none when `[sync] sheet-files` is on (spec §14.7.2).
    /// `codepage` is the project's `target.codepage`.
    pub fn take_over(&mut self, root: &Path, codepage: u16, sync: &SyncCfg) -> Result<Prepared, SyncError> {
        use plan::VbProject;
        self.connect()?;
        let book = self.live.as_ref().expect("connected");
        let code_name = book.code_name()?;
        let mut prepared = Prepared::default();
        if sync.pull_unmanaged {
            let mut project = book.project()?;
            let modules: Vec<ras_form::import::BookModule> = project.components()?.iter().filter_map(book_module).collect();
            // The controls on the sheets are read from the saved workbook by the compiler's driver;
            // the take-over of an open workbook converts their event procedures without them.
            let sheet_controls: Vec<ras_compiler::FormDecl> = Vec::new();
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
            prepared = ras_form::import::prepare(root, ras_form::import::Book { modules: &modules, workbook_code_name: &code_name, sheet_controls: &sheet_controls, form: &mut form }, codepage).map_err(SyncError::Failed)?;
            if let Some(busy) = failure {
                return Err(busy);
            }
        }
        if sync.sheet_files {
            let sheets = book.sheets()?;
            prepared.created = ras_form::import::sheet_files(root, &sheets, &code_name).map_err(SyncError::Failed)?;
        }
        Ok(prepared)
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

    pub fn apply(&mut self, _files: &[OutputFile], _forms: Forms, _opts: &SyncOptions) -> Result<Report, SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn open(&mut self) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn run_macro(&mut self, _name: &str) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn show_form(&mut self, _name: &str) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn take_over(&mut self, _root: &Path, _codepage: u16, _sync: &SyncCfg) -> Result<Prepared, SyncError> {
        Err(SyncError::Unsupported)
    }
}
