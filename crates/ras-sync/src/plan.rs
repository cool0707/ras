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
    fn remove(&mut self, name: &str) -> Result<(), SyncError>;
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
    }
}

/// A class whose attributes differ from what `VBComponents.Add` creates can only be created by `Import`.
fn needs_import(f: &OutputFile) -> bool {
    f.kind == VModuleKind::Class && f.full_text.lines().take_while(|l| !l.starts_with("' RAS-GENERATED")).any(|l| l.starts_with("Attribute VB_") && !l.starts_with("Attribute VB_Name") && l.trim_end().ends_with("= True"))
}

enum Op<'a> {
    Create(&'a OutputFile),
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
pub fn apply(project: &mut dyn VbProject, files: &[OutputFile], opts: &SyncOptions) -> Result<Report, SyncError> {
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

    let mut journal: Vec<Undo> = Vec::new();
    let mut failure = None;
    for op in ops {
        let result = match op {
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
    let Some(error) = failure else { return Ok(report) };
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
        let report = apply(&mut book, &files, &SyncOptions::default()).unwrap();
        assert!(report.created.contains(&"Main".to_string()), "{report:?}");
        let main = book.code("Main").unwrap().to_string();
        assert!(main.starts_with("' RAS-GENERATED hash="), "{main}");
        assert_eq!(main.matches("Option Explicit").count(), 1, "the line a new module starts with must not be kept: {main}");

        // The VBE stores the text with its own casing and without the last line terminator.
        let stored = main.trim_end().replace("MsgBox", "msgbox");
        book.components.iter_mut().find(|c| c.name == "Main").unwrap().code = stored;
        let again = apply(&mut book, &files, &SyncOptions::default()).unwrap();
        assert_eq!(again.changed(), 0, "{again:?}");
        assert!(again.unchanged.contains(&"Main".to_string()));

        let changed = build(&HELLO.replace("hi", "bye"));
        let report = apply(&mut book, &changed, &SyncOptions::default()).unwrap();
        assert_eq!(report.updated, ["Main"]);
        assert!(book.code("Main").unwrap().contains("bye"));
    }

    #[test]
    fn drifted_modules_are_left_alone_until_overwrite() {
        let files = build(HELLO);
        let mut book = documents();
        apply(&mut book, &files, &SyncOptions::default()).unwrap();
        let edited = book.code("Main").unwrap().replace("\"hi\"", "\"edited in the VBE\"");
        book.components.iter_mut().find(|c| c.name == "Main").unwrap().code = edited.clone();

        let changed = build(&HELLO.replace("hi", "bye"));
        let report = apply(&mut book, &changed, &SyncOptions::default()).unwrap();
        assert_eq!(report.drifted, ["Main"]);
        assert_eq!(book.code("Main").unwrap(), edited);

        let report = apply(&mut book, &changed, &SyncOptions { overwrite: true }).unwrap();
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
        let err = apply(&mut book, &files, &SyncOptions::default()).unwrap_err();
        assert!(err.to_string().contains("main"), "{err}");
        assert_eq!(book.code("main").unwrap(), hand);
        assert_eq!(book.components.len(), 3, "nothing may be written when a check fails");

        apply(&mut book, &files, &SyncOptions { overwrite: true }).unwrap();
        assert!(book.code("main").unwrap().contains("RAS-GENERATED"));
    }

    #[test]
    fn modules_that_are_no_longer_generated_are_removed_but_hand_written_ones_stay() {
        let files = build(HELLO);
        let mut book = documents();
        book.components.push(Component { name: "Legacy".into(), kind: Kind::Standard, code: "Sub Old()\r\nEnd Sub\r\n".into() });
        apply(&mut book, &files, &SyncOptions::default()).unwrap();
        // The module generated above goes away when the generated set no longer has it.
        let mut renamed = files.clone();
        for f in &mut renamed {
            if f.name == "Main" {
                f.name = "Other".into();
            }
        }
        let report = apply(&mut book, &renamed, &SyncOptions::default()).unwrap();
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
        let report = apply(&mut book, &files, &SyncOptions::default()).unwrap();
        assert!(report.updated.contains(&doc.name), "{report:?}");
        let code = book.code(&doc.name).unwrap().to_string();
        assert!(code.starts_with(hand) && code.contains(ras_ovba::RAS_BEGIN_MARKER), "{code}");
        assert_eq!(apply(&mut book, &files, &SyncOptions::default()).unwrap().changed(), 0);

        // Without the handler the region is taken out again.
        let report = apply(&mut book, &build(HELLO), &SyncOptions::default()).unwrap();
        assert!(report.updated.contains(&doc.name), "{report:?}");
        assert_eq!(book.code(&doc.name).unwrap(), hand);

        // A hand-written procedure of the same name would not compile.
        let clash = format!("Private Sub {}()\r\nEnd Sub\r\n", ras_ovba::procedure_names(&doc.code)[0]);
        let mut book = Fake::with(&[(doc.name.as_str(), Kind::Document, clash.as_str())]);
        assert!(apply(&mut book, &files, &SyncOptions::default()).is_err());
    }

    #[test]
    fn a_failed_write_puts_back_what_was_written() {
        let two = "export function hello() {\n    msgBox(\"hi\");\n}\n";
        let files = build(two);
        let mut book = documents();
        apply(&mut book, &files, &SyncOptions::default()).unwrap();
        let before: Vec<(String, String)> = book.components.iter().map(|c| (c.name.clone(), c.code.clone())).collect();

        // Every generated module changes; the last one cannot be written.
        let mut changed = build(&two.replace("hi", "bye"));
        let extra = OutputFile { name: "Extra".into(), file_name: "Extra.bas".into(), ..changed[0].clone() };
        changed.push(extra);
        book.fail_on = Some("Extra".into());
        let err = apply(&mut book, &changed, &SyncOptions::default()).unwrap_err();
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
        apply(&mut book, &[class.clone(), plain], &SyncOptions::default()).unwrap();
        assert_eq!(book.imported, ["Registry"]);
        assert_eq!(recorded_hash(book.code("Registry").unwrap()), recorded_hash(&class.code));
        assert!(book.components.iter().any(|c| c.name == "Plain" && c.kind == Kind::Class));
    }
}
