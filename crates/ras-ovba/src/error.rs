//! Error type shared by the whole crate.

use std::fmt;

/// Everything that can go wrong while reading or writing a VBA project.
#[derive(Debug)]
#[non_exhaustive]
pub enum OvbaError {
    /// File-system level failure.
    Io(std::io::Error),
    /// The OPC (zip) container could not be read or written.
    Zip(zip::result::ZipError),
    /// The compound file (`vbaProject.bin`) could not be read or written.
    Cfb(std::io::Error),
    /// The data does not follow \[MS-OVBA\] / the OPC conventions.
    Format(String),
    /// The VBA project is password protected / locked for viewing
    /// (`DPB` / `CMG` in the `PROJECT` stream). Nothing is read or written.
    Protected,
    /// The workbook is open in Excel (owner file `~$name` exists, or the
    /// file is share-locked). Nothing is written.
    Locked,
    /// Document code was supplied for a code name that does not exist in the
    /// workbook. Sheets are never created by this crate.
    UnknownCodeName(String),
    /// The file type or the requested operation is outside the supported
    /// scope (`.xls`, new user forms, ...).
    UnsupportedFormat(String),
    /// The file that was just written did not read back as expected.
    Verify(String),
    /// Text cannot be represented in the project code page, or the code page
    /// itself is not supported.
    Encoding(String),
    /// The request itself is inconsistent (bad module name, removal of a
    /// document module, ...).
    InvalidRequest(String),
    /// The generated code for a document module defines a procedure that
    /// the hand-written part of that module defines as well. The VBE would
    /// refuse the module with "Ambiguous name detected" when the workbook
    /// is opened. Nothing is written.
    DuplicateProcedure {
        /// Code name of the document module.
        module: String,
        /// Name of the procedure, as written in the hand-written part.
        procedure: String,
    },
}

impl fmt::Display for OvbaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OvbaError::Io(e) => write!(f, "I/O error: {e}"),
            OvbaError::Zip(e) => write!(f, "zip container error: {e}"),
            OvbaError::Cfb(e) => write!(f, "compound file error: {e}"),
            OvbaError::Format(m) => write!(f, "invalid VBA project data: {m}"),
            OvbaError::Protected => {
                write!(f, "the VBA project is password protected (locked for viewing)")
            }
            OvbaError::Locked => {
                write!(f, "the workbook is open in Excel or locked by another process")
            }
            OvbaError::UnknownCodeName(n) => {
                write!(f, "the workbook has no sheet or workbook with code name `{n}`")
            }
            OvbaError::UnsupportedFormat(m) => write!(f, "unsupported: {m}"),
            OvbaError::Verify(m) => write!(f, "verification after writing failed: {m}"),
            OvbaError::Encoding(m) => write!(f, "code page error: {m}"),
            OvbaError::InvalidRequest(m) => write!(f, "invalid request: {m}"),
            OvbaError::DuplicateProcedure { module, procedure } => write!(
                f,
                "the document module `{module}` already has a hand-written procedure `{procedure}`; the generated code defines it too (VBA: ambiguous name)"
            ),
        }
    }
}

impl std::error::Error for OvbaError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            OvbaError::Io(e) | OvbaError::Cfb(e) => Some(e),
            OvbaError::Zip(e) => Some(e),
            _ => None,
        }
    }
}

impl From<std::io::Error> for OvbaError {
    fn from(e: std::io::Error) -> Self {
        OvbaError::Io(e)
    }
}

impl From<zip::result::ZipError> for OvbaError {
    fn from(e: zip::result::ZipError) -> Self {
        match e {
            zip::result::ZipError::Io(io) => OvbaError::Io(io),
            other => OvbaError::Zip(other),
        }
    }
}

pub(crate) fn format_err<T>(msg: impl Into<String>) -> Result<T, OvbaError> {
    Err(OvbaError::Format(msg.into()))
}
