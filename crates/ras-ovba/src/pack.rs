//! `pack` / `unpack`: the workbook-level operations (RAS specification
//! §14.9.2 – §14.9.4).

use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use crate::codepage;
use crate::crypto;
use crate::error::OvbaError;
use crate::opc::{self, BookLayout, Package, PackagePlan, SheetKind};
use crate::project::same_name;
use crate::source;
use crate::vba::{self, DocSpec, LoadedProject};
use crate::{
    ModuleKind, PackReport, PackRequest, ProjectEdit, UnpackedBook, VbaModule,
};

/// Workbook file types, by extension.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum BookType {
    Xlsm,
    Xltm,
    Xlam,
    Xlsb,
    Xlsx,
    Xltx,
}

impl BookType {
    fn can_hold_macros(self) -> bool {
        !matches!(self, BookType::Xlsx | BookType::Xltx)
    }
}

fn book_type(path: &Path) -> Result<BookType, OvbaError> {
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase())
        .unwrap_or_default();
    match extension.as_str() {
        "xlsm" => Ok(BookType::Xlsm),
        "xltm" => Ok(BookType::Xltm),
        "xlam" => Ok(BookType::Xlam),
        "xlsb" => Ok(BookType::Xlsb),
        "xlsx" => Ok(BookType::Xlsx),
        "xltx" => Ok(BookType::Xltx),
        "xls" | "xlt" | "xla" => Err(OvbaError::UnsupportedFormat(format!(
            "`{}`: legacy binary workbooks (.{extension}, BIFF8) are not supported",
            path.display()
        ))),
        _ => Err(OvbaError::UnsupportedFormat(format!(
            "`{}` is not an Excel workbook (.xlsm, .xltm, .xlam, .xlsb or .xlsx)",
            path.display()
        ))),
    }
}

/// A sharing or lock violation reported by the operating system.
fn is_sharing_violation(error: &io::Error) -> bool {
    #[cfg(windows)]
    {
        // ERROR_SHARING_VIOLATION (32), ERROR_LOCK_VIOLATION (33)
        matches!(error.raw_os_error(), Some(32) | Some(33))
    }
    #[cfg(not(windows))]
    {
        let _ = error;
        false
    }
}

/// Maps a failure to open a workbook for reading: a sharing violation means
/// another process (Excel) holds the file.
pub(crate) fn map_open_error(error: io::Error) -> OvbaError {
    if is_sharing_violation(&error) {
        OvbaError::Locked
    } else {
        OvbaError::Io(error)
    }
}

/// Path of the owner file Excel creates next to an open workbook
/// (`~$Book.xlsm` for `Book.xlsm`).
fn owner_file(path: &Path) -> Option<PathBuf> {
    let name = path.file_name()?.to_str()?;
    Some(path.with_file_name(format!("~${name}")))
}

/// Fails with [`OvbaError::Locked`] when `path` looks open in Excel: its
/// owner file exists or, with `for_write`, the existing file cannot be
/// opened for writing.
fn check_not_locked(path: &Path, for_write: bool) -> Result<(), OvbaError> {
    if let Some(owner) = owner_file(path) {
        if owner.exists() {
            return Err(OvbaError::Locked);
        }
    }
    if for_write && path.exists() {
        match fs::OpenOptions::new().read(true).write(true).open(path) {
            Ok(_) => {}
            Err(e) if is_sharing_violation(&e) || e.kind() == io::ErrorKind::PermissionDenied => {
                return Err(OvbaError::Locked);
            }
            Err(e) => return Err(OvbaError::Io(e)),
        }
    }
    Ok(())
}

fn same_file(a: &Path, b: &Path) -> bool {
    if a == b {
        return true;
    }
    match (fs::canonicalize(a), fs::canonicalize(b)) {
        (Ok(a), Ok(b)) => a == b,
        _ => false,
    }
}

/// Reads the VBA project of a workbook without Excel.
///
/// Works for `.xlsm`, `.xltm`, `.xlam`, `.xlsb` and (trivially, they have no
/// project) `.xlsx` / `.xltx`. Returns
///
/// * [`OvbaError::Protected`] when the project is password protected,
/// * [`OvbaError::UnsupportedFormat`] for `.xls` and for workbooks that are
///   encrypted as a whole,
/// * [`OvbaError::Locked`] when the operating system refuses to open the
///   file because another process holds it.
///
/// Suspected VBA stomping is reported in
/// [`VbaProjectInfo::stomped_modules`](crate::VbaProjectInfo::stomped_modules);
/// the source text is returned in any case.
pub fn unpack(book: &Path) -> Result<UnpackedBook, OvbaError> {
    book_type(book)?;
    let bytes = fs::read(book).map_err(map_open_error)?;
    let mut package = Package::open(bytes)?;
    let layout = opc::analyze(&mut package)?;
    let project = match &layout.vba_part {
        Some(part) => {
            let bin = package.read(part)?.unwrap_or_default();
            let info = vba::read_project_bin(&bin)?;
            if info.is_protected {
                return Err(OvbaError::Protected);
            }
            Some(info)
        }
        None => None,
    };
    Ok(UnpackedBook {
        project,
        code_names: layout.code_names(),
        has_signature: !layout.signature_parts.is_empty(),
    })
}

/// Removes the temporary file unless the write completed.
struct TempFile {
    path: PathBuf,
    keep: bool,
}

impl Drop for TempFile {
    fn drop(&mut self) {
        if !self.keep {
            let _ = fs::remove_file(&self.path);
        }
    }
}

/// `YYYYMMDD-HHMMSS` (UTC) for a point in time.
fn timestamp(time: SystemTime) -> String {
    let secs = time.duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0);
    let days = (secs / 86_400) as i64;
    let rem = secs % 86_400;
    // Civil-from-days (proleptic Gregorian calendar).
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + if month <= 2 { 1 } else { 0 };
    format!(
        "{year:04}{month:02}{day:02}-{:02}{:02}{:02}",
        rem / 3_600,
        rem % 3_600 / 60,
        rem % 60
    )
}

/// Splits a backup file name `<stem>.<YYYYMMDD-HHMMSS>[-<n>].<ext>` of the
/// workbook `stem`/`ext` into a sortable key.
fn backup_key(file_name: &str, stem: &str, extension: &str) -> Option<(String, u32)> {
    let rest = file_name.strip_prefix(stem)?.strip_prefix('.')?;
    let middle = rest.strip_suffix(extension)?.strip_suffix('.')?;
    let (stamp, counter) = match middle.get(15..) {
        Some("") => (middle, 0),
        Some(tail) => (&middle[..15], tail.strip_prefix('-')?.parse::<u32>().ok()?),
        None => return None,
    };
    let bytes = stamp.as_bytes();
    let well_formed = bytes.len() == 15
        && bytes[8] == b'-'
        && bytes.iter().enumerate().all(|(i, b)| i == 8 || b.is_ascii_digit());
    well_formed.then(|| (stamp.to_string(), counter))
}

/// Sort key (`timestamp`, `counter`) and path of a backup file.
type BackupEntry = ((String, u32), PathBuf);

/// The backups of the workbook `stem`.`extension` in `dir`, oldest first.
fn list_backups(dir: &Path, stem: &str, extension: &str) -> Result<Vec<BackupEntry>, OvbaError> {
    let mut backups = Vec::new();
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let name = entry.file_name();
        if let Some(key) = name.to_str().and_then(|n| backup_key(n, stem, extension)) {
            backups.push((key, entry.path()));
        }
    }
    backups.sort();
    Ok(backups)
}

/// Copies `target` into `dir` under a timestamped name and keeps only the
/// newest `generations` backups of that workbook.
fn back_up(target: &Path, dir: &Path, generations: usize) -> Result<Option<PathBuf>, OvbaError> {
    if generations == 0 || !target.exists() {
        return Ok(None);
    }
    let stem = target.file_stem().and_then(|s| s.to_str()).unwrap_or("workbook").to_string();
    let extension = target.extension().and_then(|s| s.to_str()).unwrap_or("bak").to_string();
    fs::create_dir_all(dir)?;

    // The new backup must sort after every existing one, also when several
    // are taken within one second or the clock was set back.
    let mut stamp = timestamp(SystemTime::now());
    let mut counter = 0u32;
    if let Some(((newest_stamp, newest_counter), _)) = list_backups(dir, &stem, &extension)?.last() {
        if *newest_stamp >= stamp {
            stamp = newest_stamp.clone();
            counter = newest_counter + 1;
        }
    }
    let backup = loop {
        let name = if counter == 0 {
            format!("{stem}.{stamp}.{extension}")
        } else {
            format!("{stem}.{stamp}-{counter}.{extension}")
        };
        let candidate = dir.join(name);
        if !candidate.exists() {
            break candidate;
        }
        counter += 1;
    };
    fs::copy(target, &backup)?;

    let backups = list_backups(dir, &stem, &extension)?;
    let excess = backups.len().saturating_sub(generations);
    for (_, path) in backups.into_iter().take(excess) {
        let _ = fs::remove_file(path);
    }
    Ok(Some(backup))
}

fn write_temp_file(target: &Path, bytes: &[u8]) -> Result<TempFile, OvbaError> {
    let dir = match target.parent() {
        Some(parent) if !parent.as_os_str().is_empty() => parent.to_path_buf(),
        _ => PathBuf::from("."),
    };
    let name = target.file_name().and_then(|n| n.to_str()).unwrap_or("workbook");
    for _ in 0..16 {
        let path = dir.join(format!(
            ".{name}.{}-{:08x}.ras-tmp",
            std::process::id(),
            crypto::random_u64() as u32
        ));
        match fs::OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(mut file) => {
                let temp = TempFile { path, keep: false };
                file.write_all(bytes)?;
                file.sync_all()?;
                return Ok(temp);
            }
            Err(e) if e.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(e) => return Err(OvbaError::Io(e)),
        }
    }
    Err(OvbaError::Io(io::Error::new(
        io::ErrorKind::AlreadyExists,
        "could not create a temporary file next to the workbook",
    )))
}

/// The document modules the project must contain and the code names that
/// may receive document code.
struct DocumentPlan {
    /// Modules to create when missing.
    ensure: Vec<DocSpec>,
    /// `(code name, what it is, creatable)` of the workbook and every sheet.
    known: Vec<(String, &'static str, Option<&'static str>)>,
}

fn document_plan(layout: &BookLayout) -> DocumentPlan {
    let mut plan = DocumentPlan { ensure: Vec::new(), known: Vec::new() };
    // In a binary workbook a missing code name cannot be written into the
    // part, so only recorded code names may get a module created.
    let creatable = |explicit: bool| !layout.binary || explicit;
    let workbook_base =
        creatable(layout.explicit_workbook_code_name.is_some()).then_some(source::VB_BASE_WORKBOOK);
    plan.known.push((layout.workbook_code_name.clone(), "the workbook", workbook_base));
    for sheet in &layout.sheets {
        let (what, base) = match sheet.kind {
            SheetKind::Worksheet => ("a worksheet", Some(source::VB_BASE_WORKSHEET)),
            SheetKind::Chartsheet => ("a chart sheet", Some(source::VB_BASE_CHART)),
            SheetKind::Dialogsheet => ("a dialog sheet", None),
            SheetKind::Macrosheet => ("a macro sheet", None),
        };
        let base = base.filter(|_| creatable(sheet.explicit_code_name.is_some()));
        plan.known.push((sheet.code_name.clone(), what, base));
    }
    for (name, _, base) in &plan.known {
        if let Some(base) = base {
            plan.ensure.push(DocSpec { code_name: name.clone(), base });
        }
    }
    plan
}

/// Writes modules into a workbook without Excel (RAS specification
/// §14.9.2 / §14.9.3).
///
/// * `modules` are created or replaced as given (standard and class
///   modules; a `Document` module is accepted when its name is a code name
///   of the workbook, a `Form` module when the form already exists).
/// * `document_code` is merged into the RAS-managed region of the document
///   module with that code name ([`merge_managed_region`](crate::merge_managed_region));
///   hand-written code outside the markers is preserved, except in the
///   modules `replace_hand_written` names. A code name the workbook does
///   not have is an [`OvbaError::UnknownCodeName`].
/// * `remove_modules` are deleted.
/// * A workbook without a VBA project gets one (document modules for the
///   workbook and every sheet included), together with the relationship,
///   the content types and — where the sheet parts do not record them —
///   `codeName` attributes. A `.xlsx` input needs `convert` and is written
///   to a sibling `.xlsm` (the original is not modified).
///
/// The new package is written to a temporary file in the target directory,
/// read back and verified, the previous target is copied to `backup_dir`
/// (pruned to `backup_generations`), and the temporary file is renamed over
/// the target. When the target already holds exactly the requested modules
/// nothing is written at all; the report then lists every module as
/// `unchanged`.
///
/// Errors: [`OvbaError::Locked`] (workbook open in Excel),
/// [`OvbaError::Protected`] (password-protected project),
/// [`OvbaError::UnsupportedFormat`] (`.xls`, `.xlsx` without `convert`, new
/// user forms, ...), [`OvbaError::Encoding`] (text not representable in the
/// code page), [`OvbaError::Verify`] (the written file did not read back).
pub fn pack(req: &PackRequest) -> Result<PackReport, OvbaError> {
    // ---- what goes where ---------------------------------------------------
    let input_type = book_type(&req.book)?;
    if !input_type.can_hold_macros() && !req.convert {
        return Err(OvbaError::UnsupportedFormat(format!(
            "`{}` cannot hold a VBA project; set `convert` to write a macro-enabled copy",
            req.book.display()
        )));
    }
    let output = match &req.output {
        Some(path) => path.clone(),
        None => match input_type {
            BookType::Xlsx => req.book.with_extension("xlsm"),
            BookType::Xltx => req.book.with_extension("xltm"),
            _ => req.book.clone(),
        },
    };
    let output_type = book_type(&output)?;
    if !output_type.can_hold_macros() {
        return Err(OvbaError::UnsupportedFormat(format!(
            "the output `{}` must be a macro-enabled workbook (.xlsm, .xltm, .xlam or .xlsb)",
            output.display()
        )));
    }
    codepage::writable_encoding(req.codepage)?;
    let in_place = same_file(&req.book, &output);

    // ---- never touch a workbook that is open in Excel ---------------------
    check_not_locked(&req.book, false)?;
    check_not_locked(&output, true)?;

    let bytes = fs::read(&req.book).map_err(map_open_error)?;
    let mut package = Package::open(bytes)?;
    let layout = opc::analyze(&mut package)?;
    if layout.binary != (output_type == BookType::Xlsb) {
        return Err(OvbaError::UnsupportedFormat(
            "converting between binary (.xlsb) and XML workbook formats is not supported".into(),
        ));
    }

    let loaded = match &layout.vba_part {
        Some(part) => {
            let bin = package.read(part)?.unwrap_or_default();
            Some((LoadedProject::read(&bin)?, bin))
        }
        None => None,
    };
    let loaded_project = loaded.as_ref().map(|(project, _)| project);
    if loaded_project.is_some_and(|p| p.protection.is_protected()) {
        return Err(OvbaError::Protected);
    }
    let existing_module = |name: &str| {
        loaded_project.and_then(|p| p.modules.iter().find(|m| same_name(&m.record.name, name)))
    };

    // ---- document modules ---------------------------------------------------
    let documents = document_plan(&layout);
    if layout.binary && loaded_project.is_none() && documents.ensure.len() != documents.known.len() {
        return Err(OvbaError::UnsupportedFormat(
            "this .xlsb workbook has no VBA project and does not record code names for all sheets; a new project cannot be created offline (save it once from Excel with a macro, or use .xlsm)".into(),
        ));
    }
    let mut warnings: Vec<String> = Vec::new();
    let mut upsert: Vec<VbaModule> = Vec::with_capacity(req.modules.len() + req.document_code.len());
    for module in &req.modules {
        if module.kind == ModuleKind::Document
            && !documents.known.iter().any(|(name, _, _)| same_name(name, &module.name))
        {
            return Err(OvbaError::UnknownCodeName(module.name.clone()));
        }
        upsert.push(module.clone());
    }
    // Hand-written code that goes from a document module the RAS code does not write to.
    let mut document_code = req.document_code.clone();
    for code_name in &req.replace_hand_written {
        if !document_code.iter().any(|c| same_name(&c.code_name, code_name)) {
            document_code.push(crate::DocumentCode { code_name: code_name.clone(), managed_code: String::new() });
        }
    }
    for (index, code) in document_code.iter().enumerate() {
        let Some((code_name, what, base)) =
            documents.known.iter().find(|(name, _, _)| same_name(name, &code.code_name))
        else {
            return Err(OvbaError::UnknownCodeName(code.code_name.clone()));
        };
        if document_code[..index].iter().any(|c| same_name(&c.code_name, &code.code_name))
            || req.modules.iter().any(|m| same_name(&m.name, &code.code_name))
        {
            return Err(OvbaError::InvalidRequest(format!(
                "code for `{}` is given more than once",
                code.code_name
            )));
        }
        let current = match existing_module(code_name) {
            Some(module) if module.kind == ModuleKind::Document => module.source.clone(),
            Some(module) => {
                return Err(OvbaError::InvalidRequest(format!(
                    "`{code_name}` is the code name of {what}, but the project has a {:?} module with that name",
                    module.kind
                )));
            }
            None => match base {
                Some(base) => source::default_document_source(code_name, base),
                None => {
                    return Err(OvbaError::UnsupportedFormat(format!(
                        "the document module `{code_name}` of {what} does not exist yet and cannot be created offline"
                    )));
                }
            },
        };
        let current = if req.replace_hand_written.iter().any(|n| same_name(n, code_name)) {
            source::strip_hand_written(&current)
        } else {
            current
        };
        // What the generated code defines must not exist in the hand-written
        // part as well: the VBE reports that only when the workbook is opened
        // ("Ambiguous name detected"), and the whole project stops compiling.
        let hand_written = source::procedure_names(&source::merge_managed_region(&current, ""));
        for generated in source::procedure_names(&code.managed_code) {
            if let Some(clash) = hand_written.iter().find(|name| same_name(name, &generated)) {
                return Err(OvbaError::DuplicateProcedure {
                    module: code_name.clone(),
                    procedure: clash.clone(),
                });
            }
        }
        upsert.push(VbaModule {
            name: code_name.clone(),
            kind: ModuleKind::Document,
            source: source::merge_managed_region(&current, &code.managed_code),
        });
    }

    // ---- build the project -----------------------------------------------------
    let edit = ProjectEdit {
        upsert,
        remove: req.remove_modules.clone(),
        codepage: Some(req.codepage),
        project_name: req.project_name.clone(),
        forms: req.forms.clone(),
    };
    let built = vba::build_project(loaded_project, &documents.ensure, &edit)?;

    let mut report = PackReport { output: output.clone(), ..PackReport::default() };
    let old_modules: Vec<VbaModule> = loaded_project.map(|p| p.info().modules).unwrap_or_default();
    for module in &built.modules {
        match old_modules.iter().find(|m| same_name(&m.name, &module.name)) {
            None => report.created.push(module.name.clone()),
            Some(old) => {
                if old.kind == module.kind
                    && old.name == module.name
                    && source::sources_equivalent(&old.source, &module.source)
                    && !built.changed_designers.iter().any(|n| same_name(n, &module.name))
                {
                    report.unchanged.push(module.name.clone());
                } else {
                    report.updated.push(module.name.clone());
                }
            }
        }
    }
    for old in &old_modules {
        if !built.modules.iter().any(|m| same_name(&m.name, &old.name)) {
            report.removed.push(old.name.clone());
        }
    }
    let project_changed = match loaded_project {
        None => true,
        Some(project) => {
            !report.created.is_empty()
                || !report.updated.is_empty()
                || !report.removed.is_empty()
                || project.dir.codepage != req.codepage
                || req.project_name.as_ref().is_some_and(|name| *name != project.dir.name)
        }
    };
    if let Some(project) = loaded_project {
        for name in &project.stomped {
            warnings.push(format!(
                "module `{name}`: the stored source does not match the compiled p-code (possible VBA stomping); the source text is used"
            ));
        }
    }
    if project_changed {
        for name in &built.padded {
            warnings.push(format!(
                "module `{name}`: blank lines were appended so that the source can be stored losslessly"
            ));
        }
    }

    // ---- package ------------------------------------------------------------------
    let plan = PackagePlan {
        main_content_type: match output_type {
            BookType::Xlsm => Some(opc::CT_XLSM_MAIN),
            BookType::Xltm => Some(opc::CT_XLTM_MAIN),
            BookType::Xlam => Some(opc::CT_XLAM_MAIN),
            _ => None,
        },
        remove_signatures: project_changed,
        write_code_names: project_changed,
    };
    let mut edits = opc::plan_package_edits(&mut package, &layout, &plan)?;
    let package_changed = project_changed
        || !edits.replace.is_empty()
        || !edits.remove.is_empty()
        || !edits.add.is_empty();
    if !package_changed && in_place {
        report.warnings = warnings;
        return Ok(report);
    }

    if project_changed {
        if !layout.signature_parts.is_empty() {
            warnings.push(format!(
                "the VBA project was digitally signed; the signature is no longer valid and was removed ({})",
                layout.signature_parts.join(", ")
            ));
        }
        if !layout.binary && loaded_project.is_none() {
            let assigned: Vec<String> = layout
                .sheets
                .iter()
                .filter(|s| s.explicit_code_name.is_none())
                .map(|s| format!("`{}` -> {}", s.name, s.code_name))
                .collect();
            if !assigned.is_empty() {
                warnings.push(format!(
                    "sheets without a code name were given one: {}",
                    assigned.join(", ")
                ));
            }
        }
    }
    let expected_bin: Vec<u8> = match (&loaded, project_changed) {
        (Some((_, original)), false) => original.clone(),
        _ => built.bytes.clone(),
    };
    let expected_modules: Vec<VbaModule> =
        if project_changed { built.modules.clone() } else { old_modules.clone() };
    if project_changed {
        if layout.vba_part.is_some() {
            edits.replace.insert(edits.vba_part.clone(), built.bytes);
        } else {
            edits.add.push((edits.vba_part.clone(), built.bytes));
        }
    }
    let new_package = package.rewrite(&edits.replace, &edits.remove, &edits.add)?;

    // ---- write, verify, back up, replace ---------------------------------------
    let mut temp = write_temp_file(&output, &new_package)?;
    verify_written_package(&temp.path, &expected_bin, &expected_modules, req.codepage, project_changed)?;
    if let Some(dir) = &req.backup_dir {
        back_up(&output, dir, req.backup_generations)?;
    }
    // The replacement keeps the permissions of the file it replaces.
    if let Ok(metadata) = fs::metadata(&output) {
        let _ = fs::set_permissions(&temp.path, metadata.permissions());
    }
    // Re-check right before the switch: Excel may have opened the file since.
    check_not_locked(&output, true)?;
    fs::rename(&temp.path, &output).map_err(|e| {
        if is_sharing_violation(&e) || e.kind() == io::ErrorKind::PermissionDenied {
            OvbaError::Locked
        } else {
            OvbaError::Io(e)
        }
    })?;
    temp.keep = true;

    report.warnings = warnings;
    Ok(report)
}

/// Re-opens the package that was just written and checks that it carries
/// exactly the intended project.
fn verify_written_package(
    path: &Path,
    expected_bin: &[u8],
    expected_modules: &[VbaModule],
    codepage: u16,
    project_changed: bool,
) -> Result<(), OvbaError> {
    let fail = |what: String| OvbaError::Verify(what);
    let bytes = fs::read(path)?;
    let mut package =
        Package::open(bytes).map_err(|e| fail(format!("the new package cannot be opened: {e}")))?;
    let layout = opc::analyze(&mut package)
        .map_err(|e| fail(format!("the new package cannot be analysed: {e}")))?;
    let Some(part) = &layout.vba_part else {
        return Err(fail("the new package has no VBA project part".into()));
    };
    let bin = package
        .read(part)
        .map_err(|e| fail(format!("the VBA project part cannot be read: {e}")))?
        .unwrap_or_default();
    if bin != expected_bin {
        return Err(fail("the VBA project part differs from what was written".into()));
    }
    let info = vba::read_project_bin(&bin)
        .map_err(|e| fail(format!("the VBA project cannot be read back: {e}")))?;
    vba::verify_modules(expected_modules, &info.modules)?;
    if project_changed {
        if info.codepage != codepage {
            return Err(fail("the project code page did not read back as written".into()));
        }
        if !layout.signature_parts.is_empty() {
            return Err(fail("signature parts are still present".into()));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn timestamps() {
        let at = |secs: u64| timestamp(UNIX_EPOCH + std::time::Duration::from_secs(secs));
        assert_eq!(at(0), "19700101-000000");
        assert_eq!(at(951_782_400), "20000229-000000");
        assert_eq!(at(1_794_927_301), "20261117-145501");
        assert_eq!(at(4_102_444_799), "20991231-235959");
    }

    #[test]
    fn backup_names() {
        assert_eq!(
            backup_key("Book.20261118-145501.xlsm", "Book", "xlsm"),
            Some(("20261118-145501".to_string(), 0))
        );
        assert_eq!(
            backup_key("Book.20261118-145501-12.xlsm", "Book", "xlsm"),
            Some(("20261118-145501".to_string(), 12))
        );
        assert_eq!(backup_key("Book.xlsm", "Book", "xlsm"), None);
        assert_eq!(backup_key("Book2.20261118-145501.xlsm", "Book", "xlsm"), None);
        assert_eq!(backup_key("Book.20261118-145501.xlsb", "Book", "xlsm"), None);
        assert_eq!(backup_key("Book.2026111x-145501.xlsm", "Book", "xlsm"), None);
        assert_eq!(backup_key("Book.20261118-145501x.xlsm", "Book", "xlsm"), None);
    }

    fn scratch_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "ras-ovba-{name}-{}-{:08x}",
            std::process::id(),
            crypto::random_u64() as u32
        ));
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    /// Turns the project of a packed workbook into a password-protected one
    /// by replacing the CMG / DPB values of its PROJECT stream.
    fn protect(book: &Path) {
        let mut package = Package::open(fs::read(book).unwrap()).unwrap();
        let bin = package.read("xl/vbaProject.bin").unwrap().unwrap();
        let mut tree = crate::cfbio::read_tree(&bin).unwrap();
        let node = tree.nodes.iter_mut().find(|n| n.path == ["PROJECT"]).unwrap();
        let text = String::from_utf8(node.data.clone().unwrap()).unwrap();
        // A locked project stores the all-zero id and a 29-byte password
        // hash structure ([MS-OVBA] 2.3.1.2, 2.4.4.1).
        let id = "{00000000-0000-0000-0000-000000000000}";
        let mut hash = vec![0xFFu8];
        hash.extend([0xFF, 0xFF, 0xFF]);
        hash.extend([0x11; 4]);
        hash.extend([0x22; 20]);
        hash.push(0);
        assert_eq!(hash.len(), 29);
        let mut out = String::new();
        for line in crate::project::split_lines(&text) {
            let line = if line.starts_with("ID=") {
                format!("ID=\"{id}\"")
            } else if line.starts_with("CMG=") {
                format!("CMG=\"{}\"", crypto::encrypt(&[1, 0, 0, 0], id, 0x3C, &[1, 2, 3]))
            } else if line.starts_with("DPB=") {
                format!("DPB=\"{}\"", crypto::encrypt(&hash, id, 0x5A, &[4, 5, 6]))
            } else {
                line.to_string()
            };
            out.push_str(&line);
            out.push_str("\r\n");
        }
        node.data = Some(out.into_bytes());
        let protected = crate::cfbio::write_tree(&tree).unwrap();
        let mut replace = std::collections::BTreeMap::new();
        replace.insert("xl/vbaProject.bin".to_string(), protected);
        let bytes = package.rewrite(&replace, &Default::default(), &[]).unwrap();
        fs::write(book, bytes).unwrap();
    }

    #[test]
    fn protected_projects_are_refused() {
        let dir = scratch_dir("protected");
        let book = dir.join("book.xlsm");
        fs::write(&book, crate::create_workbook(crate::BookKind::Xlsm, &["Sheet1"])).unwrap();
        let request = PackRequest {
            book: book.clone(),
            modules: vec![VbaModule {
                name: "Secret".into(),
                kind: ModuleKind::Standard,
                source: "Sub S()\r\nEnd Sub\r\n".into(),
            }],
            ..PackRequest::default()
        };
        pack(&request).unwrap();
        assert_eq!(unpack(&book).unwrap().project.unwrap().modules.len(), 3);

        protect(&book);
        let locked = fs::read(&book).unwrap();
        assert!(matches!(unpack(&book), Err(OvbaError::Protected)));
        assert!(matches!(pack(&request), Err(OvbaError::Protected)));
        let other = PackRequest { output: Some(dir.join("copy.xlsm")), ..request.clone() };
        assert!(matches!(pack(&other), Err(OvbaError::Protected)));
        assert_eq!(fs::read(&book).unwrap(), locked, "nothing was written");
        assert!(!dir.join("copy.xlsm").exists());

        // The low-level functions report / refuse it as well.
        let mut package = Package::open(locked).unwrap();
        let bin = package.read("xl/vbaProject.bin").unwrap().unwrap();
        let info = vba::read_project_bin(&bin).unwrap();
        assert!(info.is_protected);
        assert!(matches!(
            vba::write_project_bin(Some(&bin), &[], &ProjectEdit::default()),
            Err(OvbaError::Protected)
        ));
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn generated_event_handler_must_not_duplicate_a_hand_written_one() {
        let dir = scratch_dir("dup");
        let book = dir.join("book.xlsm");
        fs::write(&book, crate::create_workbook(crate::BookKind::Xlsm, &["Sheet1"])).unwrap();
        let hand = "Private Sub workbook_open()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n";
        pack(&PackRequest {
            book: book.clone(),
            modules: vec![VbaModule { name: "ThisWorkbook".into(), kind: ModuleKind::Document, source: hand.into() }],
            ..PackRequest::default()
        })
        .unwrap();
        let before = fs::read(&book).unwrap();
        let generated = |code: &str| PackRequest {
            book: book.clone(),
            document_code: vec![crate::DocumentCode { code_name: "ThisWorkbook".into(), managed_code: code.into() }],
            ..PackRequest::default()
        };
        match pack(&generated("Private Sub Workbook_Open()\r\nEnd Sub\r\n")) {
            Err(OvbaError::DuplicateProcedure { module, procedure }) => {
                assert_eq!((module.as_str(), procedure.as_str()), ("ThisWorkbook", "workbook_open"));
            }
            other => panic!("expected a duplicate procedure, got {other:?}"),
        }
        assert_eq!(fs::read(&book).unwrap(), before, "nothing was written");
        // Another event is fine, and replacing the generated region is not a clash with itself.
        let activate = "Private Sub Workbook_Activate()\r\nEnd Sub\r\n";
        pack(&generated(activate)).unwrap();
        let report = pack(&generated(activate)).unwrap();
        assert_eq!(report.unchanged, ["ThisWorkbook", "Sheet1"]);
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn locked_output_is_detected_before_anything_is_written() {
        let dir = scratch_dir("lock");
        let book = dir.join("Book.xlsm");
        fs::write(&book, crate::create_workbook(crate::BookKind::Xlsm, &["Sheet1"])).unwrap();
        assert!(check_not_locked(&book, true).is_ok());
        fs::write(dir.join("~$Book.xlsm"), b"").unwrap();
        assert!(matches!(check_not_locked(&book, false), Err(OvbaError::Locked)));
        assert!(matches!(check_not_locked(&book, true), Err(OvbaError::Locked)));
        // A workbook that does not exist yet is not locked.
        assert!(check_not_locked(&dir.join("new.xlsm"), true).is_ok());
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn backups_sort_after_existing_ones() {
        let dir = scratch_dir("backup");
        let target = dir.join("b.xlsm");
        let backups = dir.join("bk");
        fs::create_dir_all(&backups).unwrap();
        // A backup "from the future" (clock set back since).
        fs::write(backups.join("b.29991231-235959.xlsm"), b"future").unwrap();
        fs::write(backups.join("unrelated.txt"), b"keep me").unwrap();
        fs::write(&target, b"one").unwrap();
        let first = back_up(&target, &backups, 2).unwrap().unwrap();
        assert_eq!(first.file_name().unwrap(), "b.29991231-235959-1.xlsm");
        fs::write(&target, b"two").unwrap();
        let second = back_up(&target, &backups, 2).unwrap().unwrap();
        assert_eq!(second.file_name().unwrap(), "b.29991231-235959-2.xlsm");
        assert!(!backups.join("b.29991231-235959.xlsm").exists(), "oldest pruned");
        assert_eq!(fs::read(&first).unwrap(), b"one");
        assert_eq!(fs::read(&second).unwrap(), b"two");
        assert!(backups.join("unrelated.txt").exists());
        assert!(back_up(&target, &backups, 0).unwrap().is_none());
        assert!(back_up(&dir.join("missing.xlsm"), &backups, 3).unwrap().is_none());
        let _ = fs::remove_dir_all(&dir);
    }

    #[test]
    fn book_types() {
        assert_eq!(book_type(Path::new("a/b.XLSM")).unwrap(), BookType::Xlsm);
        assert_eq!(book_type(Path::new("b.xlsb")).unwrap(), BookType::Xlsb);
        assert!(matches!(book_type(Path::new("b.xls")), Err(OvbaError::UnsupportedFormat(_))));
        assert!(matches!(book_type(Path::new("b.docm")), Err(OvbaError::UnsupportedFormat(_))));
        assert!(matches!(book_type(Path::new("noext")), Err(OvbaError::UnsupportedFormat(_))));
        assert_eq!(owner_file(Path::new("dir/Book.xlsm")).unwrap(), Path::new("dir/~$Book.xlsm"));
    }
}
