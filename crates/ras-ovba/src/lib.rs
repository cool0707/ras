//! Offline reading and writing of the VBA project inside Excel workbooks —
//! no Excel required (RAS language specification §14.9 "offline pack /
//! unpack", §17.3).
//!
//! A macro-enabled workbook (`.xlsm`, `.xltm`, `.xlam`, `.xlsb`) is an OPC
//! (zip) package whose part `xl/vbaProject.bin` is a compound file
//! (\[MS-CFB\]) laid out as described by \[MS-OVBA\]. This crate
//!
//! * reads such a project ([`unpack`], [`read_project_bin`]) — including
//!   projects saved by Excel, where every module stream starts with compiled
//!   p-code and the source follows at `MODULEOFFSET`;
//! * writes it ([`pack`], [`write_project_bin`]) **without p-code**: module
//!   streams contain only the compressed source, `_VBA_PROJECT` is reduced to
//!   its 7-byte version-independent header with version `0xFFFF`, and the
//!   `__SRP_*` caches are dropped, so Office recompiles from source on the
//!   next open regardless of its build or bitness;
//! * keeps everything it does not have to touch: project references, the
//!   project id and protection values, user forms (designer storages) and
//!   unrelated zip entries (copied raw, without recompression).
//!
//! # Module source text
//!
//! [`VbaModule::source`] is the complete text of a module stream, including
//! the leading `Attribute VB_...` lines that the VBA editor hides. When a
//! module is written
//!
//! * line terminators are normalised to CR LF,
//! * `Attribute VB_Name` is made to match the module name,
//! * a standard module without `Attribute VB_Name` gets one; a class module
//!   that lacks part of the standard header (`VB_Name`, `VB_GlobalNameSpace`,
//!   `VB_Creatable`, `VB_PredeclaredId`, `VB_Exposed`) gets the missing lines
//!   with their default value `False`,
//! * a new document module gets the document header (`VB_Base` of the
//!   workbook or worksheet class, `VB_PredeclaredId = True`, ...),
//! * the text is encoded in the project code page; a character the code page
//!   cannot represent is an error ([`OvbaError::Encoding`]), never a `?`.
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
//! User forms can be created, or their designer storage replaced, from a
//! complete designer storage generated elsewhere ([`UserForm`], e.g. by the
//! `ras-form` crate). Existing designer storages are handed out as they are
//! ([`read_user_forms`], [`unpack_user_forms`]) and can be written as the
//! `.frm` / `.frx` pair the VBA editor imports ([`export_user_form`]); they
//! are never decoded or patched here. Re-signing is not implemented, `.xls` (BIFF8) is not supported,
//! and binary (`.xlsb`) workbook and sheet parts are only read, never
//! rewritten.

#![forbid(unsafe_code)]

mod biff12;
mod cfbio;
mod codepage;
pub mod compression;
mod crypto;
mod dir;
mod error;
mod forms;
mod opc;
mod pack;
mod project;
mod source;
mod stomp;
mod vba;
mod workbook;
mod xml;

pub use error::OvbaError;
pub use forms::{export_user_form, is_generated_form, read_user_forms, unpack_user_forms, FormFiles, RAS_FORM_MARKER};
pub use opc::workbook_code_names;
pub use pack::{pack, unpack};
pub use source::{extract_managed_region, merge_managed_region, procedure_names, RAS_BEGIN_MARKER, RAS_END_MARKER};
pub use vba::{read_project_bin, write_project_bin};
pub use workbook::{create_workbook, BookKind};

/// Encodes `text` in a Windows code page (as used for module source and
/// designer properties). A character the code page cannot represent is an
/// [`OvbaError::Encoding`] error.
pub fn encode_text(codepage: u16, text: &str) -> Result<Vec<u8>, OvbaError> {
    codepage::encode(codepage, text, "text")
}

/// Decodes text stored in a Windows code page (module source, the designer
/// properties of a user form). Bytes that are not valid in the code page
/// become U+FFFD.
pub fn decode_text(codepage: u16, bytes: &[u8]) -> Result<String, OvbaError> {
    codepage::decode(codepage, bytes)
}

use std::path::PathBuf;

/// The four kinds of VBA modules.
#[derive(Clone, Copy, PartialEq, Eq, Debug, Hash)]
pub enum ModuleKind {
    /// Standard (procedural) module, `.bas`.
    Standard,
    /// Class module, `.cls`.
    Class,
    /// Document module: `ThisWorkbook` or the module behind a sheet.
    Document,
    /// Code module of a user form (designer module), `.frm`.
    Form,
}

/// A VBA module and its source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VbaModule {
    /// Module name (a VBA identifier, at most 31 characters).
    pub name: String,
    /// What kind of module this is.
    pub kind: ModuleKind,
    /// Full module text as stored in the module stream, decoded to UTF-8, including leading `Attribute VB_...` lines.
    pub source: String,
}

/// What [`read_project_bin`] found in a `vbaProject.bin`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VbaProjectInfo {
    /// Project name (`PROJECTNAME`).
    pub name: String,
    /// Windows code page of all MBCS text in the project.
    pub codepage: u16,
    /// The modules, in the order of the `dir` stream.
    pub modules: Vec<VbaModule>,
    /// The project is protected: its `DPB` value carries a password, its
    /// `CMG` value says it is locked (by the user, the host or the editor),
    /// or one of the two cannot be decrypted (Office calls such a project
    /// "unviewable"). [`pack`] and [`unpack`] refuse protected projects.
    pub is_protected: bool,
    /// Modules whose stored source does not match their compiled p-code
    /// ("VBA stomping", RAS specification §14.9.4). Best effort: only a
    /// mismatch between the p-code line table and the number of source
    /// lines is detected; empty when nothing is known or nothing is wrong.
    /// The source text is what is returned in any case.
    pub stomped_modules: Vec<String>,
}

/// A user form to create, or whose code and designer storage to replace.
///
/// The designer storage is written as given (it is not validated beyond its
/// paths); the code module gets the standard form header. A new form gets a
/// fresh `VB_Base` component identity, an existing form keeps its header.
/// Writing a form adds the reference to the Microsoft Forms 2.0 Object
/// Library (`MSForms`) when the project does not have it yet.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct UserForm {
    /// Name of the form (module name and name of the designer storage).
    pub name: String,
    /// Code pane text of the form, without `Attribute` lines.
    pub code: String,
    /// Designer streams, keyed by their path relative to the form storage
    /// (`"f"`, `"o"`, `"\u{1}CompObj"`, `"\u{3}VBFrame"`, `"i01/f"`, ...).
    pub streams: std::collections::BTreeMap<String, Vec<u8>>,
    /// Storages below (and including, as `""`) the form storage with their
    /// CLSID in wire byte order.
    pub storages: std::collections::BTreeMap<String, [u8; 16]>,
}

/// Changes to apply with [`write_project_bin`].
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ProjectEdit {
    /// Modules to create or replace (standard, class, document; the code
    /// of an *existing* form may be replaced too). `source` is the full
    /// module text including `Attribute` lines.
    pub upsert: Vec<VbaModule>,
    /// Names of standard or class modules to delete (never documents or
    /// forms). Names that do not exist are ignored.
    pub remove: Vec<String>,
    /// Code page to store the project in. `None` keeps the existing one (a
    /// new project uses 932). Changing it re-encodes every module.
    pub codepage: Option<u16>,
    /// New project name; `None` keeps the existing one (`VBAProject` for a
    /// new project).
    pub project_name: Option<String>,
    /// User forms to create or replace (code and designer storage).
    pub forms: Vec<UserForm>,
}

/// A document module that must exist in the project.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DocumentSpec {
    /// Code name of the workbook or sheet.
    pub code_name: String,
    /// `true` for the workbook module (`ThisWorkbook`), `false` for a
    /// worksheet module.
    pub is_workbook: bool,
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
    /// The code to put between the RAS markers of that module (see
    /// [`merge_managed_region`]). Empty code removes the region.
    pub managed_code: String,
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
    /// merged into the RAS-managed region of that document module
    pub document_code: Vec<DocumentCode>,
    /// Standard or class modules to delete.
    pub remove_modules: Vec<String>,
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
