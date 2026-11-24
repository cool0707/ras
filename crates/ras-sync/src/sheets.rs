//! Adding, copying, removing and renaming the sheets of the project's workbook together with
//! their files in `objects/` (spec §14.16).
//!
//! A sheet and its file (`objects/<CodeName>.ras`, spec §12.2) go together: [`apply`] changes
//! the workbook — the open one in Excel, or the file when it is closed — and then the files of
//! the project, and takes the workbook back when the files cannot follow.

use std::path::{Path, PathBuf};

use ras_compiler::refs::{self, Target};
use ras_ovba::{SheetEdit, SheetError};

use crate::{Session, SyncError};

/// A sheet of the workbook and its file.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Row {
    /// Name on its tab.
    pub name: String,
    pub code_name: String,
    /// A worksheet (not a chart sheet, …).
    pub worksheet: bool,
    pub visible: bool,
    /// Its file, relative to the project root, when it has one.
    pub file: Option<String>,
}

/// What [`apply`] did.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Outcome {
    /// `add`, `copy`, `remove` or `rename`.
    pub action: &'static str,
    /// Tab name and code name of the sheet after the change (of the removed sheet).
    pub name: String,
    pub code_name: String,
    /// Its names before a rename, or the sheet it is a copy of: `(tab name, code name)`.
    pub before: Option<(String, String)>,
    /// The open workbook in Excel was changed (it is not saved), not the file.
    pub live: bool,
    /// Files created, removed and renamed (`from`, `to`), relative to the project root.
    pub created: Vec<String>,
    pub removed: Vec<String>,
    pub renamed: Vec<(String, String)>,
    /// Number of references to the code name rewritten in the sources.
    pub references: usize,
    /// Number of formulas and defined names of the workbook file rewritten.
    pub formulas: usize,
    pub warnings: Vec<String>,
}

/// Why [`apply`] did nothing.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Failure {
    /// Removing the sheet would lose something: the reasons. `force` removes it all the same.
    Refused(Vec<String>),
    /// It cannot be done.
    Error(String),
}

impl From<String> for Failure {
    fn from(message: String) -> Failure {
        Failure::Error(message)
    }
}

fn same(a: &str, b: &str) -> bool {
    a.to_lowercase() == b.to_lowercase()
}

fn relative(root: &Path, path: &Path) -> String {
    path.strip_prefix(root)
        .unwrap_or(path)
        .to_string_lossy()
        .replace('\\', "/")
}

/// The workbook of the project at `root`.
fn workbook(root: &Path) -> Result<(ras_compiler::Config, PathBuf), String> {
    let (config, _) = ras_compiler::load_project(root)?;
    let Some(book) = config.project.workbook.as_ref().map(|w| root.join(w)) else {
        return Err(
            "対象のブックが指定されていません。ras.toml の project.workbook で指定してください"
                .into(),
        );
    };
    if !book.is_file() {
        return Err(format!("{} が見つかりません", book.display()));
    }
    Ok((config, book))
}

/// The workbook as it is reached: open in Excel, or the file.
enum Reach {
    Live(Session),
    File,
}

fn reach(book: &Path) -> Result<Reach, String> {
    let mut session = Session::new(book);
    match session.connect() {
        Ok(()) => Ok(Reach::Live(session)),
        Err(SyncError::NotOpen(_) | SyncError::Unsupported) => {
            let owner = book
                .file_name()
                .and_then(|n| n.to_str())
                .map(|n| book.with_file_name(format!("~${n}")));
            if owner.is_some_and(|o| o.exists()) {
                return Err(format!("{} は Excel で開かれていますが、Excel に接続できません（Excel で開いたブックのシートの操作は Windows でのみ行えます）。ブックを閉じてから実行してください", book.display()));
            }
            Ok(Reach::File)
        }
        Err(e) => Err(e.to_string()),
    }
}

/// The file of a sheet in `objects/`, whatever the case of its name.
fn object_file(root: &Path, code_name: &str) -> Option<PathBuf> {
    let entries = std::fs::read_dir(root.join(ras_compiler::OBJECTS_DIR)).ok()?;
    entries.flatten().map(|e| e.path()).find(|p| {
        p.is_file()
            && p.extension().is_some_and(|x| x == "ras")
            && p.file_stem()
                .is_some_and(|s| same(&s.to_string_lossy(), code_name))
    })
}

fn current_sheets(reach: &mut Reach, book: &Path) -> Result<Vec<Row>, String> {
    match reach {
        Reach::Live(session) => Ok(session
            .sheets()
            .map_err(|e| e.to_string())?
            .into_iter()
            .map(|s| Row {
                name: s.name,
                code_name: s.code_name,
                worksheet: true,
                visible: true,
                file: None,
            })
            .collect()),
        Reach::File => Ok(ras_ovba::workbook_sheets(book)
            .map_err(|e| format!("{}: {e}", book.display()))?
            .into_iter()
            .map(|s| Row {
                name: s.name,
                code_name: s.code_name,
                worksheet: s.worksheet,
                visible: s.visible,
                file: None,
            })
            .collect()),
    }
}

/// The sheets of the project's workbook with their files, and whether the open workbook in
/// Excel was asked (not the file).
pub fn list(root: &Path) -> Result<(Vec<Row>, bool), String> {
    let (_, book) = workbook(root)?;
    let mut reach = reach(&book)?;
    let mut rows = current_sheets(&mut reach, &book)?;
    for row in &mut rows {
        row.file = object_file(root, &row.code_name).map(|p| relative(root, &p));
    }
    Ok((rows, matches!(reach, Reach::Live(_))))
}

fn find<'a>(sheets: &'a [Row], wanted: &str) -> Result<&'a Row, String> {
    sheets.iter().find(|s| same(&s.code_name, wanted)).or_else(|| sheets.iter().find(|s| same(&s.name, wanted))).ok_or_else(|| format!("ブックに「{wanted}」というシートがありません（シートはコード名か見出しの名前で指定します。`ras sheet list` で確かめられます）"))
}

/// The message for what [`ras_ovba::edit_sheets`] refused.
fn refusal(e: &SheetError) -> String {
    e.japanese()
}

/// The names of the modules of the project (other than the files of `objects/`).
fn module_names(root: &Path) -> Vec<String> {
    let Ok((_, sources)) = ras_compiler::load_project(root) else {
        return Vec::new();
    };
    sources
        .iter()
        .filter(|s| s.rel_path.starts_with("src/") && !s.rel_path.ends_with(".d.ras"))
        .map(|s| ras_compiler::collect::vba_module_name(&s.key))
        .chain(sources.iter().filter_map(|s| {
            s.rel_path
                .strip_prefix("forms/")
                .and_then(|f| f.strip_suffix(".ras"))
                .map(str::to_string)
        }))
        .collect()
}

/// `file:line` of every reference to the sheet's code name in the sources, with what to put
/// there to rename it: `(file relative to root, byte range)`.
fn code_name_references(
    root: &Path,
    code_name: &str,
) -> Vec<(String, u32, std::ops::Range<usize>)> {
    let Some(prog) = crate::analyze(root) else {
        return Vec::new();
    };
    let Some(ras_compiler::model::Sym::Var(v)) = prog
        .globals
        .get(code_name)
        .or_else(|| {
            prog.globals
                .iter()
                .find(|(k, _)| same(k, code_name))
                .map(|(_, v)| v)
        })
        .cloned()
    else {
        return Vec::new();
    };
    refs::references_of(&prog, &[Target::Global(v)])
        .into_iter()
        .filter(|r| prog.modules[r.module].kind == ras_compiler::ModuleKind::User && !r.decl)
        .map(|r| {
            let md = &prog.modules[r.module];
            (
                md.rel_path.clone(),
                md.line_of(u32::from(r.span.start())),
                usize::from(r.span.start())..usize::from(r.span.end()),
            )
        })
        .collect()
}

/// Lines of the sources that name the sheet by the name on its tab (`worksheets("売上")`).
fn tab_name_mentions(root: &Path, name: &str) -> Vec<String> {
    let Ok((_, sources)) = ras_compiler::load_project(root) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for source in sources.iter().filter(|s| !s.rel_path.ends_with(".d.ras")) {
        for (i, line) in source.text.lines().enumerate() {
            if ['"', '\'', '`']
                .iter()
                .any(|q| line.contains(&format!("{q}{name}{q}")))
            {
                out.push(format!("{}:{}", source.rel_path, i + 1));
            }
        }
    }
    out
}

/// The edit that takes `edit` back.
fn inverse(edit: &SheetEdit, report: &ras_ovba::SheetEditReport) -> Option<SheetEdit> {
    match edit {
        SheetEdit::Add { .. } | SheetEdit::Copy { .. } => Some(SheetEdit::Remove {
            sheet: report.code_name.clone(),
        }),
        SheetEdit::Rename { .. } => {
            report
                .before
                .as_ref()
                .map(|(name, code_name)| SheetEdit::Rename {
                    sheet: report.code_name.clone(),
                    name: Some(name.clone()),
                    code_name: Some(code_name.clone()),
                })
        }
        SheetEdit::Remove { .. } => None,
    }
}

fn edit_book(
    reach: &mut Reach,
    root: &Path,
    config: &ras_compiler::Config,
    book: &Path,
    edit: &SheetEdit,
) -> Result<ras_ovba::SheetEditReport, String> {
    match reach {
        Reach::Live(session) => session.edit_sheet(edit).map_err(|e| e.to_string()),
        Reach::File => {
            let request = ras_ovba::SheetEditRequest {
                book: book.to_path_buf(),
                edit: edit.clone(),
                backup_dir: Some(root.join("build").join("backup")),
                backup_generations: config.pack.backup_generations,
            };
            ras_ovba::edit_sheets(&request).map_err(|e| match e {
                ras_ovba::OvbaError::Sheet(e) => refusal(&e),
                ras_ovba::OvbaError::Locked => format!(
                    "{} は Excel で開かれているため、ファイルを書き換えられません",
                    book.display()
                ),
                other => format!("{}: {other}", book.display()),
            })
        }
    }
}

/// Adds, copies, removes or renames a sheet of the workbook of the project at `root`, and
/// its file with it (spec §14.16). `force`: remove a sheet although something refers to it.
pub fn apply(root: &Path, edit: &SheetEdit, force: bool) -> Result<Outcome, Failure> {
    let (config, book) = workbook(root)?;
    let mut reach = reach(&book)?;
    let sheets = current_sheets(&mut reach, &book)?;
    let live = matches!(reach, Reach::Live(_));
    let taken_code_names = || -> Vec<String> {
        sheets
            .iter()
            .map(|s| s.code_name.clone())
            .chain(["ThisWorkbook".to_string()])
            .chain(module_names(root))
            .collect()
    };
    let check_name = |name: &str, but: Option<&Row>| -> Result<(), String> {
        if !ras_ovba::valid_sheet_name(name) {
            return Err(refusal(&SheetError::BadName(name.to_string())));
        }
        if sheets
            .iter()
            .any(|s| same(&s.name, name) && !but.is_some_and(|b| same(&b.code_name, &s.code_name)))
        {
            return Err(refusal(&SheetError::NameTaken(name.to_string())));
        }
        Ok(())
    };
    let check_code_name = |code_name: &str, but: Option<&Row>| -> Result<(), String> {
        if !ras_ovba::valid_code_name(code_name) {
            return Err(refusal(&SheetError::BadCodeName(code_name.to_string())));
        }
        if taken_code_names()
            .iter()
            .any(|c| same(c, code_name) && !but.is_some_and(|b| same(&b.code_name, c)))
        {
            return Err(refusal(&SheetError::CodeNameTaken(code_name.to_string())));
        }
        if !but.is_some_and(|b| same(&b.code_name, code_name))
            && object_file(root, code_name).is_some()
        {
            return Err(format!("objects/{code_name}.ras が既にあります（どのシートにも対応していないファイルです）。ファイルを消すか名前を変えてから実行してください"));
        }
        Ok(())
    };
    let place_exists = |place: &ras_ovba::Place| -> Result<(), String> {
        match place {
            ras_ovba::Place::End => Ok(()),
            ras_ovba::Place::Before(s) | ras_ovba::Place::After(s) => find(&sheets, s).map(|_| ()),
        }
    };

    let mut outcome = Outcome {
        live,
        ..Outcome::default()
    };
    // What must be known before the workbook changes.
    let mut source_edits: Vec<(String, u32, std::ops::Range<usize>)> = Vec::new();
    let mut aside: Option<(PathBuf, PathBuf)> = None;
    match edit {
        SheetEdit::Add {
            name,
            code_name,
            place,
        } => {
            outcome.action = "add";
            check_name(name, None)?;
            if let Some(c) = code_name {
                check_code_name(c, None)?;
            }
            place_exists(place)?;
        }
        SheetEdit::Copy {
            sheet,
            name,
            code_name,
            place,
        } => {
            outcome.action = "copy";
            find(&sheets, sheet)?;
            if let Some(n) = name {
                check_name(n, None)?;
            }
            if let Some(c) = code_name {
                check_code_name(c, None)?;
            }
            if let Some(p) = place {
                place_exists(p)?;
            }
        }
        SheetEdit::Remove { sheet } => {
            outcome.action = "remove";
            let row = find(&sheets, sheet)?.clone();
            if !sheets
                .iter()
                .any(|s| s.visible && !same(&s.code_name, &row.code_name))
            {
                return Err(Failure::Error(refusal(&SheetError::LastVisibleSheet)));
            }
            let file = object_file(root, &row.code_name);
            let mut reasons = Vec::new();
            if file
                .as_ref()
                .and_then(|f| std::fs::read_to_string(f).ok())
                .is_some_and(|text| ras_form::import::ras_has_code(&text))
            {
                reasons.push(format!(
                    "objects/{}.ras にコードがあります（ファイルも削除されます）",
                    row.code_name
                ));
            }
            let own = file.as_ref().map(|f| relative(root, f));
            let in_sources: Vec<String> = code_name_references(root, &row.code_name)
                .into_iter()
                .filter(|(path, _, _)| Some(path) != own.as_ref())
                .map(|(path, line, _)| format!("{path}:{line}"))
                .collect();
            if !in_sources.is_empty() {
                reasons.push(format!("プロジェクトのソースに `{}` への参照があります（削除するとコンパイル エラーになります）: {}", row.code_name, in_sources.join("、")));
            }
            // The formulas of the workbook, as it was saved last.
            if let Ok(cells) = ras_ovba::sheet_references(&book, &row.code_name) {
                if !cells.is_empty() {
                    let shown: Vec<&str> = cells.iter().take(10).map(String::as_str).collect();
                    reasons.push(format!("ブックの数式・名前の定義に、シート「{}」への参照があります（削除すると #REF! になります）: {}{}", row.name, shown.join("、"), if cells.len() > shown.len() { format!(" ほか {} 件", cells.len() - shown.len()) } else { String::new() }));
                }
            }
            if !reasons.is_empty() && !force {
                return Err(Failure::Refused(reasons));
            }
            // The file steps aside first: when it cannot, nothing has changed yet.
            if let Some(file) = file {
                let moved = file.with_extension("ras.removing");
                std::fs::rename(&file, &moved).map_err(|e| format!("{}: {e}", file.display()))?;
                aside = Some((file, moved));
            }
        }
        SheetEdit::Rename {
            sheet,
            name,
            code_name,
        } => {
            outcome.action = "rename";
            let row = find(&sheets, sheet)?.clone();
            if name.is_none() && code_name.is_none() {
                return Err(Failure::Error("新しい名前を指定してください（`--name <見出しの名前>`、`--code-name <コード名>`）".into()));
            }
            if let Some(n) = name {
                check_name(n, Some(&row))?;
            }
            if let Some(c) = code_name.as_ref().filter(|c| **c != row.code_name) {
                check_code_name(c, Some(&row))?;
                source_edits = code_name_references(root, &row.code_name);
            }
        }
    }

    // The workbook.
    let report = match edit_book(&mut reach, root, &config, &book, edit) {
        Ok(report) => report,
        Err(message) => {
            if let Some((file, moved)) = aside {
                let _ = std::fs::rename(moved, file);
            }
            return Err(Failure::Error(message));
        }
    };
    outcome.name = report.name.clone();
    outcome.code_name = report.code_name.clone();
    outcome.before = report.before.clone();
    outcome.formulas = report.rewritten;

    // The files of the project.
    let files = (|| -> Result<(), String> {
        let objects = root.join(ras_compiler::OBJECTS_DIR);
        let written = crate::written_dir(root);
        match edit {
            SheetEdit::Add { .. } | SheetEdit::Copy { .. } => {
                let path = ras_compiler::object_code_path(root, &report.code_name);
                std::fs::create_dir_all(&objects)
                    .map_err(|e| format!("{}: {e}", objects.display()))?;
                let source = report
                    .before
                    .as_ref()
                    .filter(|_| matches!(edit, SheetEdit::Copy { .. }))
                    .and_then(|(_, code_name)| object_file(root, code_name));
                let text = match &source {
                    Some(source) => std::fs::read_to_string(source)
                        .map_err(|e| format!("{}: {e}", source.display()))?,
                    None => ras_form::import::sheet_file_text(
                        Some(&ras_form::import::Sheet {
                            code_name: report.code_name.clone(),
                            name: report.name.clone(),
                        }),
                        "ThisWorkbook",
                    ),
                };
                std::fs::write(&path, text).map_err(|e| format!("{}: {e}", path.display()))?;
                outcome.created.push(relative(root, &path));
            }
            SheetEdit::Remove { .. } => {
                if let Some((file, moved)) = &aside {
                    std::fs::remove_file(moved).map_err(|e| format!("{}: {e}", moved.display()))?;
                    outcome.removed.push(relative(root, file));
                }
                let _ = std::fs::remove_file(written.join(format!("{}.vba", report.code_name)));
            }
            SheetEdit::Rename { .. } => {
                let Some((old_name, old_code_name)) = &report.before else {
                    return Ok(());
                };
                if *old_code_name != report.code_name {
                    // The references first (the positions are those of the files as they are).
                    let mut by_file: Vec<(String, Vec<std::ops::Range<usize>>)> = Vec::new();
                    for (path, _, range) in &source_edits {
                        match by_file.iter_mut().find(|(p, _)| p == path) {
                            Some((_, ranges)) => ranges.push(range.clone()),
                            None => by_file.push((path.clone(), vec![range.clone()])),
                        }
                    }
                    for (path, mut ranges) in by_file {
                        let file = root.join(&path);
                        let mut text = std::fs::read_to_string(&file)
                            .map_err(|e| format!("{}: {e}", file.display()))?;
                        let bom = if text.starts_with('\u{feff}') {
                            '\u{feff}'.len_utf8()
                        } else {
                            0
                        };
                        ranges.sort_by_key(|r| std::cmp::Reverse(r.start));
                        ranges.dedup();
                        for range in ranges {
                            let (start, end) = (range.start + bom, range.end + bom);
                            if text.get(start..end).is_some_and(|t| same(t, old_code_name)) {
                                text.replace_range(start..end, &report.code_name);
                                outcome.references += 1;
                            }
                        }
                        std::fs::write(&file, text)
                            .map_err(|e| format!("{}: {e}", file.display()))?;
                    }
                    if let Some(file) = object_file(root, old_code_name) {
                        let to = ras_compiler::object_code_path(root, &report.code_name);
                        std::fs::rename(&file, &to)
                            .map_err(|e| format!("{}: {e}", file.display()))?;
                        outcome
                            .renamed
                            .push((relative(root, &file), relative(root, &to)));
                    }
                    let _ = std::fs::rename(
                        written.join(format!("{old_code_name}.vba")),
                        written.join(format!("{}.vba", report.code_name)),
                    );
                }
                if *old_name != report.name {
                    let mentions = tab_name_mentions(root, old_name);
                    if !mentions.is_empty() {
                        outcome.warnings.push(format!("ソースに、シートを見出しの名前「{old_name}」で指している箇所があります（書き換えていません。コード名 `{}` で参照すれば、見出しの名前を変えても影響を受けません）: {}", report.code_name, mentions.join("、")));
                    }
                }
            }
        }
        Ok(())
    })();
    if let Err(message) = files {
        // The workbook goes back to what it was.
        let undone = match inverse(edit, &report) {
            Some(back) => edit_book(&mut reach, root, &config, &book, &back).is_ok(),
            None => false,
        };
        return Err(Failure::Error(if undone {
            format!("{message}（ブックは元に戻しました）")
        } else {
            format!("{message}（ブックのシートは変更済みです）")
        }));
    }
    if let Reach::Live(session) = &mut reach {
        // What the compiler takes for the sheets of the workbook until it is saved.
        if let Ok(now) = session.sheets() {
            crate::save_live_code_names(
                root,
                &now.into_iter().map(|s| s.code_name).collect::<Vec<_>>(),
            );
        }
        outcome.warnings.push("Excel で開いているブックを変更しました。ブックを保存するまで、ブックのファイルには反映されません".into());
    }
    Ok(outcome)
}
