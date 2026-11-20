//! What a sync does to a VBA project (spec §14.6, §14.7), independent of how the project is reached.
//!
//! [`apply`] compares the generated modules with the components of the open project and
//! creates, rewrites or removes what differs. The project is behind [`VbProject`], so the
//! rules are tested without Excel.

use ras_compiler::emit::{has_drifted, same_code, OutputFile};
use ras_compiler::vba::VModuleKind;

use crate::SyncError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Standard,
    Class,
    Form,
    /// `ThisWorkbook` and the sheet modules.
    Document,
    Other,
}

/// A component of the project with its code-pane text.
#[derive(Debug, Clone)]
pub struct Component {
    pub name: String,
    pub kind: Kind,
    pub code: String,
}

/// The VBA project of the open workbook.
pub trait VbProject {
    fn components(&mut self) -> Result<Vec<Component>, SyncError>;
    /// Replaces the whole code-pane text of an existing component.
    fn set_code(&mut self, name: &str, code: &str) -> Result<(), SyncError>;
    /// Adds an empty standard or class module with the given name.
    fn add(&mut self, name: &str, kind: Kind) -> Result<(), SyncError>;
    /// Imports an exported module file (`VERSION 1.0 CLASS` … with its `Attribute` lines).
    fn import(&mut self, name: &str, file_text: &str) -> Result<(), SyncError>;
    /// Imports a user form from its `.frm` text and `.frx` contents.
    fn import_form(&mut self, name: &str, files: &ras_ovba::FormFiles) -> Result<(), SyncError>;
    /// Gives a component another name.
    fn rename(&mut self, name: &str, new_name: &str) -> Result<(), SyncError>;
    /// Makes the forms `names` ready to be renamed or rewritten: closes their designer windows
    /// and unloads their loaded instances (a form hidden with `Hide`, a modeless form). The VBE
    /// refuses to rename a loaded form.
    fn unload_forms(&mut self, names: &[String]) -> Result<(), SyncError>;
    fn remove(&mut self, name: &str) -> Result<(), SyncError>;
}

/// The user forms generated from `forms/` (spec §14.11).
#[derive(Debug, Clone, Copy)]
pub struct Forms<'a> {
    /// Code (starting with the marker line) and designer storage of every form.
    pub forms: &'a [ras_ovba::UserForm],
    /// Code page the designer properties are stored in (`target.codepage`).
    pub codepage: u16,
}

impl Forms<'_> {
    /// No forms.
    pub const NONE: Forms<'static> = Forms { forms: &[], codepage: 932 };
}

#[derive(Debug, Clone, Copy, Default)]
pub struct SyncOptions {
    /// Replace generated modules that were edited in the VBE (drift) and hand-written modules of the same name.
    pub overwrite: bool,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Report {
    pub created: Vec<String>,
    pub updated: Vec<String>,
    pub removed: Vec<String>,
    pub unchanged: Vec<String>,
    /// Generated modules that were edited in the VBE and therefore left alone (spec §14.7).
    pub drifted: Vec<String>,
}

impl Report {
    /// Number of modules that were written or removed.
    pub fn changed(&self) -> usize {
        self.created.len() + self.updated.len() + self.removed.len()
    }

    pub fn summary(&self) -> String {
        format!("作成 {} / 更新 {} / 削除 {} / 変更なし {}", self.created.len(), self.updated.len(), self.removed.len(), self.unchanged.len())
    }
}

/// Whether a module holds anything besides `Attribute` lines, `Option` statements and blank lines.
pub fn has_code(source: &str) -> bool {
    source.lines().map(str::trim).any(|l| !l.is_empty() && !l.starts_with("Attribute ") && !l.starts_with("VERSION ") && !l.to_ascii_lowercase().starts_with("option "))
}

fn same_name(a: &str, b: &str) -> bool {
    a.to_lowercase() == b.to_lowercase()
}

fn kind_of(k: &VModuleKind) -> Kind {
    match k {
        VModuleKind::Standard => Kind::Standard,
        VModuleKind::Class => Kind::Class,
        VModuleKind::Document => Kind::Document,
        VModuleKind::Form => Kind::Form,
    }
}

/// A class whose attributes differ from what `VBComponents.Add` creates can only be created by `Import`.
fn needs_import(f: &OutputFile) -> bool {
    f.kind == VModuleKind::Class && f.full_text.lines().take_while(|l| !l.starts_with("' RAS-GENERATED")).any(|l| l.starts_with("Attribute VB_") && !l.starts_with("Attribute VB_Name") && l.trim_end().ends_with("= True"))
}

/// The marker line of a form written from `forms/`: `' RAS-FORM <file> design=<hash> code=<hash>`.
fn form_marker(code: &str) -> Option<&str> {
    code.lines().map(str::trim_end).find(|l| !l.trim().is_empty() && !l.starts_with("Attribute ") && !l.starts_with("VERSION ")).filter(|l| l.starts_with(ras_ovba::RAS_FORM_MARKER))
}

/// The `design=` field of a marker line.
fn design_of(marker: &str) -> Option<&str> {
    marker.split_whitespace().find_map(|field| field.strip_prefix("design="))
}

/// A component name that no component of the project has.
fn spare_name(existing: &[Component]) -> String {
    (1..).map(|n| format!("RasOld{n}")).find(|candidate| !existing.iter().any(|c| same_name(&c.name, candidate))).expect("a free name")
}

enum Op<'a> {
    Create(&'a OutputFile),
    CreateForm(&'a ras_ovba::UserForm),
    /// The designer changed: the form is imported anew. `aside` is the name the old one gets meanwhile.
    ReplaceForm { old: String, aside: String, form: &'a ras_ovba::UserForm },
    /// Rewrite the code of the existing component `name` (its spelling in the project).
    Write { name: String, code: String, before: String },
    /// A component of the same name but of the other kind (module / class) is replaced.
    Replace { old: Component, file: &'a OutputFile },
    Remove(Component),
}

/// How to take back an operation that was already carried out.
enum Undo {
    Remove(String),
    Restore { name: String, code: String },
    Recreate(Component),
    /// Remove the new form `name` and give the old one, set aside as `aside`, its name back.
    PutBack { name: String, aside: String },
}

fn create(project: &mut dyn VbProject, f: &OutputFile) -> Result<(), SyncError> {
    if needs_import(f) {
        return project.import(&f.name, &f.full_text);
    }
    project.add(&f.name, kind_of(&f.kind))?;
    if let Err(e) = project.set_code(&f.name, &f.code) {
        let _ = project.remove(&f.name);
        return Err(e);
    }
    Ok(())
}

fn undo(project: &mut dyn VbProject, journal: Vec<Undo>) -> Vec<String> {
    let mut failed = Vec::new();
    for step in journal.into_iter().rev() {
        let (name, result) = match step {
            Undo::Remove(name) => (name.clone(), project.remove(&name)),
            Undo::Restore { name, code } => (name.clone(), project.set_code(&name, &code)),
            Undo::Recreate(c) => (c.name.clone(), project.add(&c.name, c.kind).and_then(|()| project.set_code(&c.name, &c.code))),
            Undo::PutBack { name, aside } => (name.clone(), project.remove(&name).and_then(|()| project.rename(&aside, &name))),
        };
        if result.is_err() {
            failed.push(name);
        }
    }
    failed
}

/// Brings the project in line with the generated modules.
///
/// Nothing is written when a check fails. When a write fails half-way, the modules already
/// written are put back as they were (spec §14.6 "一括性").
///
/// A form is imported as a whole when it is new or its design changed (the marker line of its
/// code records a hash of the design); otherwise only its code is rewritten. A form made in the
/// VBE is replaced only with `overwrite`.
pub fn apply(project: &mut dyn VbProject, files: &[OutputFile], forms: Forms, opts: &SyncOptions) -> Result<Report, SyncError> {
    // The code of a form (`forms/<Name>.ras`) travels with its designer (`forms`), not as a module of its own.
    let files: Vec<OutputFile> = files.iter().filter(|f| f.kind != VModuleKind::Form).cloned().collect();
    let files = &files[..];
    let existing = project.components()?;
    let find = |name: &str| existing.iter().find(|c| same_name(&c.name, name));
    let mut report = Report::default();
    let mut ops: Vec<Op> = Vec::new();
    let mut hand_written = Vec::new();

    for f in files {
        if f.kind == VModuleKind::Document {
            let Some(doc) = find(&f.name).filter(|c| c.kind == Kind::Document) else {
                return Err(SyncError::Failed(format!("ブックに `{}` というコード名のシート（またはブック）がありません。Excel 側でシートを追加するか、RAS 側の指定を直してください", f.name)));
            };
            // The VBE reports a procedure that exists twice only as a compile error dialog.
            let hand = ras_ovba::procedure_names(&ras_ovba::merge_managed_region(&doc.code, ""));
            for generated in ras_ovba::procedure_names(&f.code) {
                if let Some(clash) = hand.iter().find(|n| same_name(n, &generated)) {
                    return Err(SyncError::Failed(format!(
                        "ドキュメントモジュール `{}` には手書きの `{clash}` があり、RAS が生成するプロシージャと名前が重なります。手書き側の処理を RAS の `@on` ハンドラへ移すか、手書きのプロシージャを削除してください",
                        doc.name
                    )));
                }
            }
            let merged = ras_ovba::merge_managed_region(&doc.code, &f.code);
            if same_code(&merged, &doc.code) {
                report.unchanged.push(doc.name.clone());
            } else {
                ops.push(Op::Write { name: doc.name.clone(), code: merged, before: doc.code.clone() });
            }
            continue;
        }
        let Some(c) = find(&f.name) else {
            ops.push(Op::Create(f));
            continue;
        };
        if !matches!(c.kind, Kind::Standard | Kind::Class) {
            return Err(SyncError::Failed(format!("`{}` という名前は、ブックのシートやフォームがすでに使っています。RAS 側のファイル名を変えてください", c.name)));
        }
        match has_drifted(&c.code) {
            None if has_code(&c.code) && !opts.overwrite => {
                hand_written.push(c.name.clone());
                continue;
            }
            Some(true) if !opts.overwrite => {
                report.drifted.push(c.name.clone());
                continue;
            }
            _ => {}
        }
        if c.kind != kind_of(&f.kind) {
            ops.push(Op::Replace { old: c.clone(), file: f });
        } else if same_code(&c.code, &f.code) {
            report.unchanged.push(c.name.clone());
        } else {
            ops.push(Op::Write { name: c.name.clone(), code: f.code.clone(), before: c.code.clone() });
        }
    }
    if !hand_written.is_empty() {
        return Err(SyncError::Failed(format!(
            "ブックには RAS が生成したものではない同名のモジュールがあります: {}\n  このまま同期すると手書きのコードが失われます。VBE でモジュール名を変えるか、RAS 側のファイル名を変えてください。置き換えてよい場合は上書きを指定してください",
            hand_written.join(", ")
        )));
    }

    let mut hand_made = Vec::new();
    let mut set_aside: Vec<Component> = existing.clone();
    for form in forms.forms {
        if files.iter().any(|f| same_name(&f.name, &form.name)) {
            return Err(SyncError::Failed(format!("ユーザーフォーム `{}` と同じ名前の RAS モジュールがあります。どちらかの名前を変えてください", form.name)));
        }
        let Some(c) = find(&form.name) else {
            ops.push(Op::CreateForm(form));
            continue;
        };
        if c.kind != Kind::Form {
            return Err(SyncError::Failed(format!("ブックには `{}` という名前のフォームではないモジュールがあります。フォーム名を変えてください", c.name)));
        }
        let (old, new) = (form_marker(&c.code), form_marker(&form.code));
        match old {
            None if !opts.overwrite => hand_made.push(c.name.clone()),
            Some(old) if Some(old) == new => report.unchanged.push(c.name.clone()),
            Some(old) if design_of(old).is_some() && design_of(old) == new.and_then(design_of) => {
                ops.push(Op::Write { name: c.name.clone(), code: form.code.clone(), before: c.code.clone() });
            }
            _ => {
                let aside = spare_name(&set_aside);
                set_aside.push(Component { name: aside.clone(), kind: Kind::Form, code: String::new() });
                ops.push(Op::ReplaceForm { old: c.name.clone(), aside, form });
            }
        }
    }
    if !hand_made.is_empty() {
        return Err(SyncError::Failed(format!(
            "ブックには forms/ から生成したものではない同名のユーザーフォームがあります: {}\n  このまま同期すると VBE で作ったフォームとそのコードが失われます。フォーム名を変えるか、`ras form pull` で forms/ に取り出してください。置き換えてよい場合は上書きを指定してください",
            hand_made.join(", ")
        )));
    }

    // What RAS generated earlier and no longer generates.
    for c in &existing {
        if files.iter().any(|f| same_name(&f.name, &c.name)) {
            continue;
        }
        match c.kind {
            Kind::Standard | Kind::Class => match has_drifted(&c.code) {
                Some(true) if !opts.overwrite => report.drifted.push(c.name.clone()),
                Some(_) => ops.push(Op::Remove(c.clone())),
                None => {}
            },
            Kind::Document if ras_ovba::extract_managed_region(&c.code).is_some() => {
                ops.push(Op::Write { name: c.name.clone(), code: ras_ovba::merge_managed_region(&c.code, ""), before: c.code.clone() });
            }
            _ => {}
        }
    }

    // Forms that are replaced or whose code is rewritten must not be loaded.
    let mut busy_forms: Vec<String> = Vec::new();
    for op in &ops {
        let name = match op {
            Op::ReplaceForm { old, .. } => old,
            Op::Write { name, .. } if find(name).is_some_and(|c| c.kind == Kind::Form) => name,
            _ => continue,
        };
        if !busy_forms.iter().any(|n| same_name(n, name)) {
            busy_forms.push(name.clone());
        }
    }
    if !busy_forms.is_empty() {
        project.unload_forms(&busy_forms)?;
    }

    let export = |form: &ras_ovba::UserForm| ras_ovba::export_user_form(form, forms.codepage).map_err(|e| SyncError::Failed(format!("ユーザーフォーム `{}` を書き出せません: {e}", form.name)));
    let mut journal: Vec<Undo> = Vec::new();
    // Old forms that were set aside; they go once everything else worked.
    let mut leftovers: Vec<String> = Vec::new();
    let mut failure = None;
    for op in ops {
        let result = match op {
            Op::CreateForm(form) => export(form).and_then(|files| project.import_form(&form.name, &files)).map(|()| {
                journal.push(Undo::Remove(form.name.clone()));
                report.created.push(form.name.clone());
            }),
            Op::ReplaceForm { old, aside, form } => export(form).and_then(|files| {
                project.rename(&old, &aside)?;
                if let Err(e) = project.import_form(&form.name, &files) {
                    let _ = project.rename(&aside, &old);
                    return Err(e);
                }
                journal.push(Undo::PutBack { name: form.name.clone(), aside: aside.clone() });
                leftovers.push(aside);
                report.updated.push(form.name.clone());
                Ok(())
            }),
            Op::Create(f) => create(project, f).map(|()| {
                journal.push(Undo::Remove(f.name.clone()));
                report.created.push(f.name.clone());
            }),
            Op::Write { name, code, before } => project.set_code(&name, &code).map(|()| {
                journal.push(Undo::Restore { name: name.clone(), code: before });
                report.updated.push(name);
            }),
            Op::Replace { old, file } => project.remove(&old.name).and_then(|()| {
                journal.push(Undo::Recreate(old.clone()));
                create(project, file).map(|()| {
                    journal.push(Undo::Remove(file.name.clone()));
                    report.updated.push(file.name.clone());
                })
            }),
            Op::Remove(c) => project.remove(&c.name).map(|()| {
                report.removed.push(c.name.clone());
                journal.push(Undo::Recreate(c));
            }),
        };
        if let Err(e) = result {
            failure = Some(e);
            break;
        }
    }
    let Some(error) = failure else {
        let stuck: Vec<String> = leftovers.into_iter().filter(|name| project.remove(name).is_err()).collect();
        if !stuck.is_empty() {
            return Err(SyncError::Failed(format!("フォームは更新しましたが、置き換え前のフォームを削除できませんでした。VBE で削除してください: {}", stuck.join(", "))));
        }
        return Ok(report);
    };
    let not_restored = undo(project, journal);
    Err(match error {
        // Nothing to add: the caller retries later.
        SyncError::Busy(_) if not_restored.is_empty() => error,
        e if not_restored.is_empty() => SyncError::Failed(format!("{e}（書き込み済みのモジュールは元に戻しました）")),
        e => SyncError::Failed(format!("{e}（元に戻せなかったモジュール: {}）", not_restored.join(", "))),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use ras_compiler::emit::recorded_hash;

    /// A project in memory that behaves like the VBE where the rules depend on it.
    #[derive(Default)]
    struct Fake {
        components: Vec<Component>,
        /// `set_code` fails for this module.
        fail_on: Option<String>,
        imported: Vec<String>,
        /// Forms that were imported, with the size of their `.frx`.
        imported_forms: Vec<(String, usize)>,
        /// `import_form` fails for this form.
        fail_import: Option<String>,
        /// Forms that were unloaded before they were written.
        unloaded: Vec<String>,
        /// `unload_forms` fails (a form refused to close).
        fail_unload: bool,
    }

    impl Fake {
        fn with(components: &[(&str, Kind, &str)]) -> Fake {
            Fake { components: components.iter().map(|(n, k, c)| Component { name: n.to_string(), kind: *k, code: c.to_string() }).collect(), ..Fake::default() }
        }

        fn code(&self, name: &str) -> Option<&str> {
            self.components.iter().find(|c| c.name == name).map(|c| c.code.as_str())
        }
    }

    impl VbProject for Fake {
        fn components(&mut self) -> Result<Vec<Component>, SyncError> {
            Ok(self.components.clone())
        }

        fn set_code(&mut self, name: &str, code: &str) -> Result<(), SyncError> {
            if self.fail_on.as_deref() == Some(name) {
                return Err(SyncError::Failed("書き込めません".into()));
            }
            let c = self.components.iter_mut().find(|c| same_name(&c.name, name)).ok_or_else(|| SyncError::Failed(format!("no {name}")))?;
            c.code = code.to_string();
            Ok(())
        }

        fn add(&mut self, name: &str, kind: Kind) -> Result<(), SyncError> {
            assert!(!self.components.iter().any(|c| same_name(&c.name, name)), "{name} exists");
            // With "Require Variable Declaration" a new module is not empty.
            self.components.push(Component { name: name.to_string(), kind, code: "Option Explicit\r\n".into() });
            Ok(())
        }

        fn import(&mut self, name: &str, file_text: &str) -> Result<(), SyncError> {
            let code = file_text.lines().skip_while(|l| !l.starts_with("' RAS-GENERATED")).map(|l| format!("{l}\r\n")).collect();
            self.components.push(Component { name: name.to_string(), kind: Kind::Class, code });
            self.imported.push(name.to_string());
            Ok(())
        }

        fn import_form(&mut self, name: &str, files: &ras_ovba::FormFiles) -> Result<(), SyncError> {
            if self.fail_import.as_deref() == Some(name) {
                return Err(SyncError::Failed("取り込めません".into()));
            }
            assert!(!self.components.iter().any(|c| same_name(&c.name, name)), "{name} exists");
            assert!(files.frm.contains(&format!("OleObjectBlob   =   \"{name}.frx\":0000")), "{}", files.frm);
            assert!(files.frx.starts_with(b"LB"));
            let code = files.frm.split("\r\n").skip_while(|l| !l.starts_with("Attribute VB_Exposed")).skip(1).collect::<Vec<_>>().join("\r\n");
            self.components.push(Component { name: name.to_string(), kind: Kind::Form, code });
            self.imported_forms.push((name.to_string(), files.frx.len()));
            Ok(())
        }

        fn unload_forms(&mut self, names: &[String]) -> Result<(), SyncError> {
            if self.fail_unload {
                return Err(SyncError::Failed("閉じられません".into()));
            }
            self.unloaded.extend(names.iter().cloned());
            Ok(())
        }

        fn rename(&mut self, name: &str, new_name: &str) -> Result<(), SyncError> {
            assert!(!self.components.iter().any(|c| same_name(&c.name, new_name)), "{new_name} exists");
            let c = self.components.iter_mut().find(|c| same_name(&c.name, name)).ok_or_else(|| SyncError::Failed(format!("no {name}")))?;
            c.name = new_name.to_string();
            Ok(())
        }

        fn remove(&mut self, name: &str) -> Result<(), SyncError> {
            let before = self.components.len();
            self.components.retain(|c| !same_name(&c.name, name));
            if self.components.len() == before {
                return Err(SyncError::Failed(format!("no {name}")));
            }
            Ok(())
        }
    }

    /// Compiles a one-file project and returns its generated modules.
    fn build(source: &str) -> Vec<OutputFile> {
        let dir = std::env::temp_dir().join(format!("ras-sync-plan-{}-{:?}", std::process::id(), std::thread::current().id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\n").unwrap();
        std::fs::write(dir.join("src").join("main.ras"), source).unwrap();
        let (mut config, sources) = ras_compiler::load_project(&dir).unwrap();
        config.sheet_code_names = vec!["Sheet1".into()];
        let prog = ras_compiler::analyze(sources, config);
        let out = ras_compiler::build(&prog, &ras_compiler::BuildOptions { release: false, erl: None, no_attributes: true });
        let _ = std::fs::remove_dir_all(&dir);
        assert!(!out.files.is_empty(), "{:?} {:?}", prog.diagnostics, out.diagnostics);
        out.files
    }

    const HELLO: &str = "export function hello() {\n    msgBox(\"hi\");\n}\n";

    fn documents() -> Fake {
        Fake::with(&[("ThisWorkbook", Kind::Document, ""), ("Sheet1", Kind::Document, "")])
    }

    #[test]
    fn creates_then_skips_then_updates() {
        let files = build(HELLO);
        let mut book = documents();
        let report = apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        assert!(report.created.contains(&"Main".to_string()), "{report:?}");
        let main = book.code("Main").unwrap().to_string();
        assert!(main.starts_with("' RAS-GENERATED hash="), "{main}");
        assert_eq!(main.matches("Option Explicit").count(), 1, "the line a new module starts with must not be kept: {main}");

        // The VBE stores the text with its own casing and without the last line terminator.
        let stored = main.trim_end().replace("MsgBox", "msgbox");
        book.components.iter_mut().find(|c| c.name == "Main").unwrap().code = stored;
        let again = apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(again.changed(), 0, "{again:?}");
        assert!(again.unchanged.contains(&"Main".to_string()));

        let changed = build(&HELLO.replace("hi", "bye"));
        let report = apply(&mut book, &changed, Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(report.updated, ["Main"]);
        assert!(book.code("Main").unwrap().contains("bye"));
    }

    #[test]
    fn drifted_modules_are_left_alone_until_overwrite() {
        let files = build(HELLO);
        let mut book = documents();
        apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        let edited = book.code("Main").unwrap().replace("\"hi\"", "\"edited in the VBE\"");
        book.components.iter_mut().find(|c| c.name == "Main").unwrap().code = edited.clone();

        let changed = build(&HELLO.replace("hi", "bye"));
        let report = apply(&mut book, &changed, Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(report.drifted, ["Main"]);
        assert_eq!(book.code("Main").unwrap(), edited);

        let report = apply(&mut book, &changed, Forms::NONE, &SyncOptions { overwrite: true }).unwrap();
        assert_eq!(report.updated, ["Main"]);
        assert!(report.drifted.is_empty());
        assert!(book.code("Main").unwrap().contains("bye"));
    }

    #[test]
    fn hand_written_module_of_the_same_name_stops_the_sync() {
        let files = build(HELLO);
        let hand = "Option Explicit\r\nSub Mine()\r\nEnd Sub\r\n";
        let mut book = documents();
        book.components.push(Component { name: "main".into(), kind: Kind::Standard, code: hand.into() });
        let err = apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("main"), "{err}");
        assert_eq!(book.code("main").unwrap(), hand);
        assert_eq!(book.components.len(), 3, "nothing may be written when a check fails");

        apply(&mut book, &files, Forms::NONE, &SyncOptions { overwrite: true }).unwrap();
        assert!(book.code("main").unwrap().contains("RAS-GENERATED"));
    }

    #[test]
    fn modules_that_are_no_longer_generated_are_removed_but_hand_written_ones_stay() {
        let files = build(HELLO);
        let mut book = documents();
        book.components.push(Component { name: "Legacy".into(), kind: Kind::Standard, code: "Sub Old()\r\nEnd Sub\r\n".into() });
        apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        // The module generated above goes away when the generated set no longer has it.
        let mut renamed = files.clone();
        for f in &mut renamed {
            if f.name == "Main" {
                f.name = "Other".into();
            }
        }
        let report = apply(&mut book, &renamed, Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(report.removed, ["Main"]);
        assert!(report.created.contains(&"Other".to_string()));
        assert!(book.code("Main").is_none() && book.code("Legacy").is_some());
    }

    #[test]
    fn document_modules_keep_the_hand_written_part() {
        let files = build("@on(\"Workbook.Open\")\nfunction opened() {\n    msgBox(\"open\");\n}\n");
        let doc = files.iter().find(|f| f.kind == VModuleKind::Document).expect("a document module");
        let hand = "Option Explicit\r\nPrivate Sub Hand()\r\nEnd Sub\r\n";
        let mut book = Fake::with(&[(doc.name.as_str(), Kind::Document, hand), ("Sheet1", Kind::Document, "")]);
        let report = apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        assert!(report.updated.contains(&doc.name), "{report:?}");
        let code = book.code(&doc.name).unwrap().to_string();
        assert!(code.starts_with(hand) && code.contains(ras_ovba::RAS_BEGIN_MARKER), "{code}");
        assert_eq!(apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap().changed(), 0);

        // Without the handler the region is taken out again.
        let report = apply(&mut book, &build(HELLO), Forms::NONE, &SyncOptions::default()).unwrap();
        assert!(report.updated.contains(&doc.name), "{report:?}");
        assert_eq!(book.code(&doc.name).unwrap(), hand);

        // A hand-written procedure of the same name would not compile.
        let clash = format!("Private Sub {}()\r\nEnd Sub\r\n", ras_ovba::procedure_names(&doc.code)[0]);
        let mut book = Fake::with(&[(doc.name.as_str(), Kind::Document, clash.as_str())]);
        assert!(apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).is_err());
    }

    #[test]
    fn a_failed_write_puts_back_what_was_written() {
        let two = "export function hello() {\n    msgBox(\"hi\");\n}\n";
        let files = build(two);
        let mut book = documents();
        apply(&mut book, &files, Forms::NONE, &SyncOptions::default()).unwrap();
        let before: Vec<(String, String)> = book.components.iter().map(|c| (c.name.clone(), c.code.clone())).collect();

        // Every generated module changes; the last one cannot be written.
        let mut changed = build(&two.replace("hi", "bye"));
        let extra = OutputFile { name: "Extra".into(), file_name: "Extra.bas".into(), ..changed[0].clone() };
        changed.push(extra);
        book.fail_on = Some("Extra".into());
        let err = apply(&mut book, &changed, Forms::NONE, &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("元に戻しました"), "{err}");
        let after: Vec<(String, String)> = book.components.iter().map(|c| (c.name.clone(), c.code.clone())).collect();
        assert_eq!(after, before);
    }

    #[test]
    fn classes_with_attributes_are_imported() {
        let files = build(HELLO);
        let mut class = files[0].clone();
        class.name = "Registry".into();
        class.kind = VModuleKind::Class;
        class.full_text = format!("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"Registry\"\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\n{}", class.code);
        let mut plain = class.clone();
        plain.name = "Plain".into();
        plain.full_text = plain.full_text.replace("VB_PredeclaredId = True", "VB_PredeclaredId = False");
        let mut book = documents();
        apply(&mut book, &[class.clone(), plain], Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(book.imported, ["Registry"]);
        assert_eq!(recorded_hash(book.code("Registry").unwrap()), recorded_hash(&class.code));
        assert!(book.components.iter().any(|c| c.name == "Plain" && c.kind == Kind::Class));
    }

    /// A form as `ras_form::workspace::build` hands it over: marker line, code, designer streams.
    fn form(name: &str, design: &str, body: &str) -> ras_ovba::UserForm {
        let frame = format!("VERSION 5.00\r\nBegin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {name}\r\n   Caption = \"{name}\"\r\n   ClientHeight = 3600\r\n   ClientWidth = 4800\r\n   StartUpPosition = 1\r\nEnd\r\n");
        let mut streams = std::collections::BTreeMap::new();
        streams.insert("f".to_string(), design.as_bytes().to_vec());
        streams.insert("o".to_string(), Vec::new());
        streams.insert("\u{1}CompObj".to_string(), vec![0; 28]);
        streams.insert("\u{3}VBFrame".to_string(), frame.into_bytes());
        let code = format!("' RAS-FORM forms/{name}.form.json design={design} code={:x}\n{body}", body.len());
        ras_ovba::UserForm { name: name.into(), code, streams, storages: [(String::new(), [0; 16])].into() }
    }

    fn with_forms(forms: &[ras_ovba::UserForm]) -> Forms<'_> {
        Forms { forms, codepage: 932 }
    }

    #[test]
    fn forms_are_imported_and_updated() {
        let mut book = documents();
        let first = [form("frmInput", "d1", "Sub A()\nEnd Sub\n")];
        let report = apply(&mut book, &[], with_forms(&first), &SyncOptions::default()).unwrap();
        assert_eq!(report.created, ["frmInput"]);
        assert_eq!(book.imported_forms.len(), 1);
        let code = book.code("frmInput").unwrap().to_string();
        assert!(code.starts_with("' RAS-FORM forms/frmInput.form.json design=d1 "), "{code}");

        // The VBE stores the code its own way; the marker line decides.
        let stored = code.replace("Sub A()", "Sub a()").trim_end().to_string();
        book.components.iter_mut().find(|c| c.name == "frmInput").unwrap().code = stored;
        let report = apply(&mut book, &[], with_forms(&first), &SyncOptions::default()).unwrap();
        assert_eq!((report.changed(), report.unchanged.as_slice()), (0, ["frmInput".to_string()].as_slice()));

        // Only the code changed: it is rewritten in place.
        let code_only = [form("frmInput", "d1", "Sub A()\n    Beep\nEnd Sub\n")];
        let report = apply(&mut book, &[], with_forms(&code_only), &SyncOptions::default()).unwrap();
        assert_eq!(report.updated, ["frmInput"]);
        assert_eq!(book.imported_forms.len(), 1, "no import for a code change");
        assert_eq!(book.unloaded, ["frmInput"], "a loaded form is closed before its code is written");
        assert!(book.code("frmInput").unwrap().contains("Beep"));

        // The design changed: the form is imported again and the old one is gone.
        let redesigned = [form("frmInput", "d2", "Sub A()\n    Beep\nEnd Sub\n")];
        let report = apply(&mut book, &[], with_forms(&redesigned), &SyncOptions::default()).unwrap();
        assert_eq!(report.updated, ["frmInput"]);
        assert_eq!(book.imported_forms.len(), 2);
        assert!(book.code("frmInput").unwrap().contains("design=d2 "));
        assert_eq!(book.components.iter().filter(|c| c.kind == Kind::Form).count(), 1, "{:?}", book.components);
        assert_eq!(book.unloaded, ["frmInput", "frmInput"], "a loaded form is closed before it is replaced");

        // A form that cannot be closed is not touched.
        let again = [form("frmInput", "d3", "")];
        book.fail_unload = true;
        let before = book.code("frmInput").unwrap().to_string();
        assert!(apply(&mut book, &[], with_forms(&again), &SyncOptions::default()).is_err());
        assert_eq!(book.code("frmInput").unwrap(), before);
        book.fail_unload = false;

        // A form that is no longer in forms/ stays in the workbook.
        let report = apply(&mut book, &[], Forms::NONE, &SyncOptions::default()).unwrap();
        assert_eq!(report.changed(), 0);
        assert!(book.code("frmInput").is_some());
    }

    #[test]
    fn a_form_made_in_the_vbe_is_not_replaced_silently() {
        let hand = "Private Sub UserForm_Click()\r\nEnd Sub\r\n";
        let mut book = documents();
        book.components.push(Component { name: "FrmInput".into(), kind: Kind::Form, code: hand.into() });
        let forms = [form("frmInput", "d1", "")];
        let err = apply(&mut book, &[], with_forms(&forms), &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("FrmInput") && err.to_string().contains("ras form pull"), "{err}");
        assert_eq!(book.code("FrmInput").unwrap(), hand);

        let report = apply(&mut book, &[], with_forms(&forms), &SyncOptions { overwrite: true }).unwrap();
        assert_eq!(report.updated, ["frmInput"]);
        assert!(book.code("frmInput").unwrap().starts_with("' RAS-FORM "));
        assert!(book.code("FrmInput").is_none() && book.components.len() == 3);

        // Names used by something else.
        let mut book = documents();
        book.components.push(Component { name: "frmInput".into(), kind: Kind::Standard, code: String::new() });
        assert!(apply(&mut book, &[], with_forms(&forms), &SyncOptions { overwrite: true }).is_err());
        let files = build(HELLO);
        let clash = [form("Main", "d1", "")];
        let err = apply(&mut documents(), &files, with_forms(&clash), &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("同じ名前の RAS モジュール"), "{err}");
    }

    #[test]
    fn a_failed_form_import_keeps_the_old_form() {
        let mut book = documents();
        let first = [form("frmInput", "d1", "Sub A()\nEnd Sub\n"), form("frmOther", "d1", "")];
        apply(&mut book, &[], with_forms(&first), &SyncOptions::default()).unwrap();
        let before: Vec<(String, String)> = book.components.iter().map(|c| (c.name.clone(), c.code.clone())).collect();

        // The first form is replaced, the import of the second one fails: both are as before.
        let changed = [form("frmInput", "d2", "Sub A()\nEnd Sub\n"), form("frmOther", "d2", "")];
        book.fail_import = Some("frmOther".into());
        let err = apply(&mut book, &[], with_forms(&changed), &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("元に戻しました"), "{err}");
        let mut after: Vec<(String, String)> = book.components.iter().map(|c| (c.name.clone(), c.code.clone())).collect();
        let mut expected = before.clone();
        after.sort();
        expected.sort();
        assert_eq!(after, expected);
    }
}
