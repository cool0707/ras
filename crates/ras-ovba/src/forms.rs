//! User forms: the forms of a workbook file, the marker of the forms RAS
//! writes, and the project-level functions of [`umya_spreadsheet::vba`]
//! (handing out the designer storages, the `.frm` / `.frx` pair of the VBA
//! editor) with this crate's error type.

use std::path::Path;

use umya_spreadsheet::vba::{self, VbaProject};

use crate::error::OvbaError;
use crate::opc::{self, Package};
use crate::{FormFiles, UserForm};

/// Start of the first code line of a form that RAS wrote from `forms/`
/// (`' RAS-FORM forms/<Name>.form.json ...`). A form without it was made in
/// the VBA editor.
pub const RAS_FORM_MARKER: &str = "' RAS-FORM ";

/// Whether the code of a form module (with or without its `Attribute`
/// header) starts with [`RAS_FORM_MARKER`].
pub fn is_generated_form(source: &str) -> bool {
    source
        .replace("\r\n", "\n")
        .lines()
        .find(|l| {
            !l.starts_with("Attribute ") && !l.starts_with("VERSION ") && !l.trim().is_empty()
        })
        .is_some_and(|l| l.starts_with(RAS_FORM_MARKER))
}

/// The user forms of a `vbaProject.bin`: for every form module its code
/// (without the `Attribute` header) and its designer storage, in the layout
/// of [`UserForm`] ([`umya_spreadsheet::vba::read_user_forms`]).
pub fn read_user_forms(bin: &[u8]) -> Result<Vec<UserForm>, OvbaError> {
    Ok(vba::read_user_forms(bin)?)
}

/// The user forms of a workbook (see [`read_user_forms`]); empty when the
/// workbook has no VBA project. A protected project is refused with
/// [`OvbaError::Protected`].
pub fn unpack_user_forms(book: &Path) -> Result<Vec<UserForm>, OvbaError> {
    let bytes = std::fs::read(book).map_err(crate::pack::map_open_error)?;
    let mut package = Package::open(bytes)?;
    let layout = opc::analyze(&mut package)?;
    let Some(part) = &layout.vba_part else {
        return Ok(Vec::new());
    };
    let bin = package.read(part)?.unwrap_or_default();
    let project = VbaProject::read(&bin)?;
    if project.is_protected() {
        return Err(OvbaError::Protected);
    }
    Ok(project.user_forms())
}

/// Writes `form` as a `.frm` / `.frx` pair for `VBComponents.Import`
/// ([`umya_spreadsheet::vba::export_user_form`]). `codepage` is the code page
/// the `VBFrame` stream is stored in.
pub fn export_user_form(form: &UserForm, codepage: u16) -> Result<FormFiles, OvbaError> {
    Ok(vba::export_user_form(form, codepage)?)
}

/// Reads a user form from the `.frm` / `.frx` pair the VBA editor writes
/// (`VBComponent.Export`) ([`umya_spreadsheet::vba::import_user_form`]). `frm`
/// is the text of the `.frm` file, already decoded from the ANSI code page it
/// was written in; the designer block goes to the `VBFrame` stream in
/// `codepage`.
pub fn import_user_form(frm: &str, frx: &[u8], codepage: u16) -> Result<UserForm, OvbaError> {
    Ok(vba::import_user_form(frm, frx, codepage)?)
}
