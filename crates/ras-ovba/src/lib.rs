//! Offline reading and writing of the VBA project inside Excel workbooks —
//! no Excel required (RAS language specification §14.9 "offline pack /
//! unpack", §17.3).
//!
//! A macro-enabled workbook (`.xlsm`, `.xltm`, `.xlam`, `.xlsb`) is an OPC
//! (zip) package whose part `xl/vbaProject.bin` is a compound file
//! (\[MS-CFB\]) laid out as described by \[MS-OVBA\]. The project itself —
//! reading it (including projects saved by Excel, with p-code in front of
//! the source), writing it without p-code while keeping references, the
//! project id, protection values and designer storages byte for byte, the
//! module source text and the user forms — is
//! [`umya_spreadsheet::vba`], which xljs uses as well; this crate re-exports
//! its types and functions. What this crate adds is the package around it:
//!
//! * [`unpack`] / [`pack`]: the project of a workbook file, with the
//!   relationship, the content types and the code names of the sheets, and
//!   the code of document modules as RAS generates it
//!   ([`replace_document_code`], the RAS-managed region of a module);
//! * the sheets of a workbook ([`edit_sheets`], [`workbook_sheets`]) and new
//!   workbooks ([`create_workbook`]);
//! * the user forms of a workbook ([`unpack_user_forms`]).
//!
//! The package is edited in place: XML parts are changed by splicing and
//! every unrelated zip entry is copied raw, without recompression (it is not
//! saved through umya-spreadsheet).
//!
//! # Safety net of [`pack`]
//!
//! The package is written to a temporary file next to the target, read back
//! and verified module by module, and only then renamed over the target. The
//! previous file is first copied to the backup directory. A workbook that is
//! open in Excel (owner file `~$name` present, or the file is locked) is
//! refused with [`OvbaError::Locked`]; a password-protected project is refused
//! with [`OvbaError::Protected`]; signature parts are removed with a warning
//! because the signature would no longer be valid.
//!
//! # Scope
//!
//! Re-signing is not implemented, `.xls` (BIFF8) is not supported, and binary
//! (`.xlsb`) workbook and sheet parts are only read, never rewritten.

#![forbid(unsafe_code)]

mod biff12;
mod error;
mod forms;
mod opc;
mod pack;
mod sheets;
mod source;
mod workbook;
mod xml;

pub use error::OvbaError;
pub use forms::{
    export_user_form, import_user_form, is_generated_form, read_user_forms, unpack_user_forms,
    RAS_FORM_MARKER,
};
pub use opc::{workbook_code_names, workbook_sheet_controls, workbook_sheet_controls_bytes};
pub use pack::{pack, pack_bytes, replace_workbook, unpack, unpack_bytes};
pub use sheets::{
    edit_sheets, sheet_references, valid_code_name, valid_sheet_name, workbook_sheets, Place,
    SheetEdit, SheetEditReport, SheetEditRequest, SheetError, SheetInfo,
};
pub use source::{
    extract_managed_region, merge_managed_region, procedure_names, replace_document_code,
    strip_hand_written, RAS_BEGIN_MARKER, RAS_END_MARKER,
};
/// The MS-OVBA compression ([`umya_spreadsheet::vba::compression`]).
pub use umya_spreadsheet::vba::compression;
pub use umya_spreadsheet::vba::{
    DocumentKind, DocumentSpec, FormFiles, ModuleKind, ProjectEdit, UserForm, VbaModule,
    VbaProjectInfo,
};
pub use workbook::{create_workbook, BookKind};

use std::path::PathBuf;

use umya_spreadsheet::vba;

/// VBA identifiers (and the names of the parts of a package) compare
/// case-insensitively.
pub(crate) fn same_name(a: &str, b: &str) -> bool {
    a == b || a.to_lowercase() == b.to_lowercase()
}

/// Reads a `vbaProject.bin` ([`umya_spreadsheet::vba::read_project_bin`]).
///
/// For a protected project the function still succeeds and reports
/// `is_protected = true`; [`unpack`] and [`pack`] turn that into
/// [`OvbaError::Protected`].
pub fn read_project_bin(bin: &[u8]) -> Result<VbaProjectInfo, OvbaError> {
    Ok(vba::read_project_bin(bin)?)
}

/// Creates a new `vbaProject.bin` or rewrites an existing one
/// ([`umya_spreadsheet::vba::write_project_bin`], see
/// [`umya_spreadsheet::vba::build_project`] for what is written and kept).
pub fn write_project_bin(
    existing: Option<&[u8]>,
    documents: &[DocumentSpec],
    edit: &ProjectEdit,
) -> Result<Vec<u8>, OvbaError> {
    Ok(vba::write_project_bin(existing, documents, edit)?)
}

/// Encodes `text` in a Windows code page (as used for module source and
/// designer properties). A character the code page cannot represent is an
/// [`OvbaError::Encoding`] error.
pub fn encode_text(codepage: u16, text: &str) -> Result<Vec<u8>, OvbaError> {
    Ok(vba::encode_text(codepage, text)?)
}

/// Decodes text stored in a Windows code page (module source, the designer
/// properties of a user form). Bytes that are not valid in the code page
/// become U+FFFD.
pub fn decode_text(codepage: u16, bytes: &[u8]) -> Result<String, OvbaError> {
    Ok(vba::decode_text(codepage, bytes)?)
}

/// The code names of a workbook and its sheets.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CodeNames {
    /// Code name of the workbook (normally `ThisWorkbook`).
    pub workbook: String,
    /// One entry per sheet, in workbook order.
    pub sheets: Vec<SheetCodeName>,
}

/// Code name of one sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetCodeName {
    /// Name on the sheet tab.
    pub sheet_name: String,
    /// VBA code name (name of the sheet's document module).
    pub code_name: String,
}

/// The ActiveX controls on a worksheet (see [`workbook_sheet_controls`]).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetControls {
    /// VBA code name of the sheet.
    pub code_name: String,
    pub controls: Vec<SheetControl>,
}

/// An ActiveX control on a worksheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetControl {
    /// Its name: a member of the sheet's module (`CommandButton1`).
    pub name: String,
    /// Its MSForms class (`CommandButton`, `TextBox`, …); empty for another control.
    pub class: String,
}

/// Result of [`unpack`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UnpackedBook {
    /// The VBA project; `None` when the workbook has none.
    pub project: Option<VbaProjectInfo>,
    /// Code names of the workbook and its sheets.
    pub code_names: CodeNames,
    /// The VBA project is digitally signed.
    pub has_signature: bool,
}

/// RAS-generated code for one document module.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocumentCode {
    /// Code name of the workbook or sheet the code belongs to.
    pub code_name: String,
    /// The whole code of that module, as RAS generates it from the module's
    /// file (RAS specification §12.2): it replaces everything after the
    /// module's `Attribute` header. Empty code leaves the module without code.
    pub code: String,
}

/// Everything [`pack`] needs to know.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PackRequest {
    /// existing .xlsm/.xltm/.xlam/.xlsb, or .xlsx when convert = true
    pub book: PathBuf,
    /// None = replace `book` atomically (for .xlsx+convert: sibling .xlsm, original untouched)
    pub output: Option<PathBuf>,
    /// Allow a `.xlsx` (or `.xltx`) input; the result is written as a
    /// macro-enabled workbook and the input is left untouched.
    pub convert: bool,
    /// Target code page of the project (RAS `target.codepage`).
    pub codepage: u16,
    /// Project name to set; `None` keeps the current name.
    pub project_name: Option<String>,
    /// Standard/Class modules to create or replace
    pub modules: Vec<VbaModule>,
    /// The code of document modules (the workbook's, a sheet's), each
    /// replaced as a whole.
    pub document_code: Vec<DocumentCode>,
    /// Standard or class modules to delete.
    pub remove_modules: Vec<String>,
    /// Modules to rename, `(old, new)`, before anything else: the document module of a
    /// sheet whose code name was changed.
    pub rename_modules: Vec<(String, String)>,
    /// Delete the document modules of sheets the workbook no longer has.
    pub prune_documents: bool,
    /// User forms to create or replace (see [`UserForm`]).
    pub forms: Vec<UserForm>,
    /// Directory that receives a copy of the file that is about to be
    /// replaced; `None` disables backups.
    pub backup_dir: Option<PathBuf>,
    /// Number of backups of this workbook to keep (oldest are pruned);
    /// `0` disables backups.
    pub backup_generations: usize,
}

impl Default for PackRequest {
    /// An empty request with code page 932 and backups disabled.
    fn default() -> Self {
        PackRequest {
            book: PathBuf::new(),
            output: None,
            convert: false,
            codepage: 932,
            project_name: None,
            modules: Vec::new(),
            document_code: Vec::new(),
            remove_modules: Vec::new(),
            rename_modules: Vec::new(),
            prune_documents: false,
            forms: Vec::new(),
            backup_dir: None,
            backup_generations: 0,
        }
    }
}

/// What [`pack`] did.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PackReport {
    /// The file that now holds the result.
    pub output: PathBuf,
    /// Modules that did not exist before.
    pub created: Vec<String>,
    /// Modules whose source changed.
    pub updated: Vec<String>,
    /// Modules that were deleted.
    pub removed: Vec<String>,
    /// Modules whose source is the same as before.
    pub unchanged: Vec<String>,
    /// Things the user should know (signature removed, code names assigned,
    /// suspected VBA stomping, ...).
    pub warnings: Vec<String>,
}
