//! Live sync (spec §14.1–14.8): writes the generated VBA into a workbook that is open in Excel.
//!
//! * [`plan`] — what a sync does to a VBA project: the rules, tested without Excel.
//! * `excel` — the open workbook, reached over COM (Windows only).
//!
//! The code is written through the VBE object model (`CodeModule`) from this process. The
//! shared-memory transport and the RASBridge add-in of spec §14.4 are an optimisation of the
//! same rules and are not part of this crate.

use std::fmt;
use std::path::{Path, PathBuf};

pub mod plan;

#[cfg(windows)]
mod disp;
#[cfg(windows)]
mod excel;

use ras_compiler::emit::OutputFile;
pub use plan::{Report, SyncOptions};

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

    /// Writes the generated modules into the open workbook.
    pub fn apply(&mut self, files: &[OutputFile], opts: &SyncOptions) -> Result<Report, SyncError> {
        self.connect()?;
        let mut project = self.live.as_ref().expect("connected").project()?;
        plan::apply(&mut project, files, opts)
    }

    /// Runs the macro `Module.proc` in the workbook and waits for it to end.
    pub fn run_macro(&mut self, name: &str) -> Result<(), SyncError> {
        self.connect()?;
        self.live.as_ref().expect("connected").run(name)
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

    pub fn apply(&mut self, _files: &[OutputFile], _opts: &SyncOptions) -> Result<Report, SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn open(&mut self) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }

    pub fn run_macro(&mut self, _name: &str) -> Result<(), SyncError> {
        Err(SyncError::Unsupported)
    }
}
