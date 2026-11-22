//! The sync thread of the language server (spec §14.2): keeps the workbook that is open in
//! Excel in line with the saved sources and reports the state to the editor (`ras/syncStatus`).
//!
//! All COM objects live on this thread. A macro started from the editor runs here too, so
//! nothing is written to the project while it runs (spec §14.6).

use std::collections::VecDeque;
use std::path::{Path, PathBuf};
use std::sync::mpsc::{Receiver, RecvTimeoutError, Sender};
use std::time::Duration;

use lsp_server::{Message, Notification, RequestId, Response};
use ras_compiler::emit::OutputFile;
use ras_compiler::{BuildOptions, Config};
use ras_sync::{Forms, Prepared, Report, Session, SyncError};
use serde_json::{json, Value};

/// How often the thread looks whether the workbook was opened or closed, and retries a sync Excel was too busy for.
const POLL: Duration = Duration::from_secs(2);

pub enum Job {
    /// The project the server works on (sent when it is known and when it changes).
    Root(PathBuf),
    /// Sources were saved.
    Saved,
    /// "今すぐ同期".
    Sync { id: RequestId, overwrite: bool },
    /// "Excel で開く".
    Open { id: RequestId },
    /// "マクロを実行".
    Run { id: RequestId, name: String },
    /// "Excelで実行" of a user form.
    ShowForm { id: RequestId, name: String },
    /// "シートを追加" / "複製" / "削除" / "名前を変更" (spec §14.16); `None`: the list of the sheets.
    Sheet {
        id: RequestId,
        edit: Option<ras_ovba::SheetEdit>,
        force: bool,
    },
}

/// The sheet command of `workspace/executeCommand` arguments: `{ action, sheet, name, codeName,
/// before, after, force }`.
pub fn sheet_edit(args: &Value) -> Option<(Option<ras_ovba::SheetEdit>, bool)> {
    use ras_ovba::{Place, SheetEdit};
    let text = |key: &str| {
        args[key]
            .as_str()
            .filter(|s| !s.is_empty())
            .map(str::to_string)
    };
    let place = match (text("before"), text("after")) {
        (Some(s), _) => Some(Place::Before(s)),
        (_, Some(s)) => Some(Place::After(s)),
        _ => None,
    };
    let edit = match args["action"].as_str()? {
        "list" => None,
        "add" => Some(SheetEdit::Add {
            name: text("name")?,
            code_name: text("codeName"),
            place: place.unwrap_or(Place::End),
        }),
        "copy" => Some(SheetEdit::Copy {
            sheet: text("sheet")?,
            name: text("name"),
            code_name: text("codeName"),
            place,
        }),
        "remove" => Some(SheetEdit::Remove {
            sheet: text("sheet")?,
        }),
        "rename" => Some(SheetEdit::Rename {
            sheet: text("sheet")?,
            name: text("name"),
            code_name: text("codeName"),
        }),
        _ => return None,
    };
    Some((edit, args["force"].as_bool().unwrap_or(false)))
}

/// Carries out a sheet command for the project at `root` and answers as JSON: `{ ok: true, … }`,
/// `{ ok: false, refused: [reasons] }` (a removal that needs `force`) or `{ ok: false, message }`.
pub fn run_sheet_command(root: &Path, edit: Option<&ras_ovba::SheetEdit>, force: bool) -> Value {
    let Some(edit) = edit else {
        return match ras_sync::sheets::list(root) {
            Ok((rows, live)) => {
                json!({ "ok": true, "live": live, "sheets": rows.iter().map(|r| json!({ "name": r.name, "codeName": r.code_name, "worksheet": r.worksheet, "visible": r.visible, "file": r.file })).collect::<Vec<_>>() })
            }
            Err(message) => json!({ "ok": false, "message": message }),
        };
    };
    match ras_sync::sheets::apply(root, edit, force) {
        Ok(o) => json!({
            "ok": true,
            "action": o.action,
            "name": o.name,
            "codeName": o.code_name,
            "before": o.before.as_ref().map(|(name, code_name)| json!({ "name": name, "codeName": code_name })),
            "live": o.live,
            "created": o.created,
            "removed": o.removed,
            "renamed": o.renamed.iter().map(|(from, to)| json!({ "from": from, "to": to })).collect::<Vec<_>>(),
            "references": o.references,
            "formulas": o.formulas,
            "warnings": o.warnings,
        }),
        Err(ras_sync::sheets::Failure::Refused(reasons)) => {
            json!({ "ok": false, "refused": reasons })
        }
        Err(ras_sync::sheets::Failure::Error(message)) => {
            json!({ "ok": false, "message": message })
        }
    }
}

pub type Send = Box<dyn Fn(Message) + std::marker::Send>;

/// Starts the sync thread. It ends when the returned sender is dropped.
pub fn spawn(send: Send) -> Sender<Job> {
    let (tx, rx) = std::sync::mpsc::channel();
    std::thread::Builder::new()
        .name("ras-sync".into())
        .spawn(move || {
            Worker {
                send,
                root: None,
                session: None,
                connected: false,
                pending: 0,
                drift: 0,
                retry: false,
                message: None,
                last_status: None,
                last_warning: None,
                last_take_over: None,
            }
            .run(rx)
        })
        .expect("the sync thread starts");
    tx
}

/// Why a sync did not happen.
enum Failure {
    /// The sources do not compile; the diagnostics are already in the editor.
    Build(String),
    Sync(SyncError),
}

impl Failure {
    fn message(&self) -> String {
        match self {
            Failure::Build(m) => m.clone(),
            Failure::Sync(e) => e.to_string(),
        }
    }
}

struct Worker {
    send: Send,
    root: Option<PathBuf>,
    session: Option<Session>,
    connected: bool,
    /// Saves that have not reached the workbook yet.
    pending: usize,
    drift: usize,
    /// The last sync was put off because Excel was busy: try again on the next tick.
    retry: bool,
    message: Option<String>,
    last_status: Option<Value>,
    last_warning: Option<String>,
    /// The last notification about a take-over: what could not be taken over is said once.
    last_take_over: Option<String>,
}

fn read_config(root: &Path) -> Result<Config, String> {
    let path = root.join("ras.toml");
    let text = std::fs::read_to_string(&path).map_err(|e| format!("{}: {e}", path.display()))?;
    Config::from_toml(&text).map_err(|e| format!("{}: {e}", path.display()))
}

fn book_path(root: &Path, config: &Config) -> Result<PathBuf, String> {
    match &config.project.workbook {
        Some(book) => Ok(root.join(book)),
        None => Err("ras.toml の project.workbook に同期先のブックが指定されていません".into()),
    }
}

/// What a sync writes: the generated modules, the user forms of `forms/` and their code page.
struct Built {
    files: Vec<OutputFile>,
    forms: Vec<ras_ovba::UserForm>,
    codepage: u16,
}

/// Compiles what is saved on disk, as `ras build` does, for `CodeModule.AddFromString`, and
/// generates the user forms as `ras pack` does.
fn compile(root: &Path) -> Result<Built, String> {
    let (config, sources) = crate::load_project(root)?;
    let prog = ras_compiler::analyze(sources, config);
    let out = ras_compiler::build(
        &prog,
        &BuildOptions {
            release: false,
            erl: None,
            no_attributes: true,
        },
    );
    let errors = prog
        .diagnostics
        .iter()
        .chain(&out.diagnostics)
        .filter(|d| d.is_error())
        .count();
    if errors > 0 {
        return Err(format!(
            "ビルド エラーが {errors} 件あるため、同期していません"
        ));
    }
    let codepage = prog.config.target.codepage;
    let compiled = out
        .files
        .iter()
        .filter(|f| f.kind == ras_compiler::vba::VModuleKind::Form)
        .map(|f| (f.name.clone(), f.code.clone()))
        .collect();
    let forms =
        ras_form::workspace::build(&ras_form::workspace::load(root)?, root, codepage, &compiled);
    if let Some(first) = forms
        .problems
        .iter()
        .find(|p| p.severity == ras_form::Severity::Error)
    {
        return Err(format!(
            "ユーザーフォームにエラーが {} 件あるため、同期していません（{}: {}）",
            forms.errors(),
            first.rel_path,
            first.message
        ));
    }
    Ok(Built {
        files: out.files,
        forms: forms.forms,
        codepage,
    })
}

fn describe(report: &Report) -> String {
    let mut text = if report.changed() > 0 {
        format!(
            "同期しました（{}）。VBA のグローバル変数はリセットされます",
            report.summary()
        )
    } else {
        "ブックは最新です".to_string()
    };
    if !report.drifted.is_empty() {
        text.push_str(&format!(
            "。VBE 上で編集されたため上書きしていないモジュール: {}",
            report.drifted.join(", ")
        ));
    }
    if !report.pending.is_empty() {
        text.push_str(&format!(
            "。取り込んだファイルに `// TODO(VBA):` の行が残っているため置き換えていないもの: {}",
            report.pending.join(", ")
        ));
    }
    if !report.foreign.is_empty() {
        text.push_str(&format!("。RAS が生成したコードがあるのに対応するファイルが無いため、そのまま残しているシート（シートの複製）: {}", report.foreign.join(", ")));
    }
    text
}

/// A short account of what a take-over wrote into the project, for a notification.
fn describe_take_over(prepared: &Prepared) -> Option<(i32, String)> {
    if prepared.imported.is_empty()
        && prepared.skipped.is_empty()
        && prepared.created.is_empty()
        && prepared.appended.is_empty()
    {
        return None;
    }
    let mut parts = Vec::new();
    if !prepared.created.is_empty() {
        let items: Vec<String> = prepared
            .created
            .iter()
            .map(|(path, what)| format!("{what} → {path}"))
            .collect();
        parts.push(format!(
            "シート・ブックのモジュールのファイルを作成しました（{}）",
            items.join("、")
        ));
    }
    if !prepared.appended.is_empty() {
        let items: Vec<String> = prepared
            .appended
            .iter()
            .map(|a| format!("{} の {} → {}", a.name, a.names.join("・"), a.path))
            .collect();
        parts.push(format!(
            "VBE で足されたコードを RAS に変換して取り込みました（{}）",
            items.join("、")
        ));
    }
    if !prepared.imported.is_empty() {
        let items: Vec<String> = prepared
            .imported
            .iter()
            .map(|i| format!("{} → {}", i.name, i.path))
            .collect();
        parts.push(format!(
            "ブックの VBA を RAS に変換して取り込みました（{}）。元の VBA は vba/ に保存しています",
            items.join("、")
        ));
    }
    let todos: Vec<String> = prepared
        .imported
        .iter()
        .filter(|i| !i.todos.is_empty())
        .map(|i| format!("{}（{} 行）", i.path, i.todos.len()))
        .collect();
    if !todos.is_empty() {
        parts.push(format!("変換できなかった行が `// TODO(VBA):` として残っています: {}。直すまでブックの元のコードは置き換えません", todos.join("、")));
    }
    parts.extend(prepared.skipped.iter().cloned());
    let level = if todos.is_empty() && prepared.skipped.is_empty() {
        3
    } else {
        2
    };
    Some((level, format!("RAS: {}", parts.join("。"))))
}

impl Worker {
    fn run(mut self, rx: Receiver<Job>) {
        let mut backlog: VecDeque<Job> = VecDeque::new();
        loop {
            let job = match backlog.pop_front() {
                Some(job) => job,
                None => match rx.recv_timeout(POLL) {
                    Ok(job) => job,
                    Err(RecvTimeoutError::Timeout) => {
                        self.tick();
                        continue;
                    }
                    Err(RecvTimeoutError::Disconnected) => return,
                },
            };
            match job {
                Job::Root(root) => {
                    if self.root.as_ref() != Some(&root) {
                        self.root = Some(root);
                        self.session = None;
                        self.connected = false;
                        self.tick();
                    }
                }
                Job::Saved => {
                    // An editor reports one save several times (the document, then the file watcher).
                    while let Ok(next) = rx.try_recv() {
                        if !matches!(next, Job::Saved) {
                            backlog.push_back(next);
                            break;
                        }
                    }
                    self.auto_sync(true);
                }
                Job::Sync { id, overwrite } => {
                    let result = match self.sync(overwrite) {
                        Ok(report) => {
                            json!({ "ok": true, "message": describe(&report), "drifted": report.drifted })
                        }
                        // No Excel has the workbook: the editor writes to the file instead (`ras pack`).
                        Err(e @ Failure::Sync(SyncError::NotOpen(_) | SyncError::Unsupported)) => {
                            json!({ "ok": false, "offline": true, "message": e.message() })
                        }
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
                Job::Open { id } => {
                    let opened = self.session().and_then(|s| s.open().map_err(Failure::Sync));
                    let result = match opened {
                        Ok(()) => {
                            self.connected = true;
                            match self.sync(false) {
                                Ok(report) => {
                                    json!({ "ok": true, "message": format!("Excel で開きました。{}", describe(&report)) })
                                }
                                Err(e) => {
                                    json!({ "ok": true, "message": format!("Excel で開きました。{}", e.message()) })
                                }
                            }
                        }
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
                Job::Run { id, name } => {
                    // The macro must be the one in the sources: sync first.
                    let result = match self
                        .sync(false)
                        .and_then(|_| self.session()?.run_macro(&name).map_err(Failure::Sync))
                    {
                        Ok(()) => json!({ "ok": true }),
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
                Job::Sheet { id, edit, force } => {
                    // On this thread: it is the one that talks to Excel.
                    let result = match self.root.clone() {
                        Some(root) => run_sheet_command(&root, edit.as_ref(), force),
                        None => {
                            json!({ "ok": false, "message": "ras.toml のあるフォルダーを開いてください" })
                        }
                    };
                    self.respond(id, result);
                }
                Job::ShowForm { id, name } => {
                    // The form must be the one in the sources: sync first.
                    let result = match self
                        .sync(false)
                        .and_then(|_| self.session()?.show_form(&name).map_err(Failure::Sync))
                    {
                        Ok(()) => json!({ "ok": true }),
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
            }
        }
    }

    fn respond(&mut self, id: RequestId, result: Value) {
        (self.send)(Message::Response(Response {
            id,
            result: Some(result),
            error: None,
        }));
        self.publish();
    }

    /// The session for the project's workbook.
    fn session(&mut self) -> Result<&mut Session, Failure> {
        let root = self
            .root
            .clone()
            .ok_or_else(|| Failure::Build("ras.toml のあるフォルダーを開いてください".into()))?;
        let book = read_config(&root)
            .and_then(|c| book_path(&root, &c))
            .map_err(Failure::Build)?;
        if self
            .session
            .as_ref()
            .is_none_or(|s| s.book() != std::path::absolute(&book).unwrap_or(book.clone()))
        {
            self.session = Some(Session::new(&book));
            self.connected = false;
        }
        Ok(self.session.as_mut().expect("just set"))
    }

    /// Connects, compiles and writes. Records the outcome for the status bar.
    fn sync(&mut self, overwrite: bool) -> Result<Report, Failure> {
        let result = (|| {
            let root = self.root.clone().ok_or_else(|| {
                Failure::Build("ras.toml のあるフォルダーを開いてください".into())
            })?;
            let connected = self.session()?.connect();
            self.connected = connected.is_ok();
            connected.map_err(Failure::Sync)?;
            // The hand-written VBA of the workbook is taken over into the project first (spec §14.7.1),
            // and the sheets without a file get one (spec §14.7.2).
            let config = read_config(&root).map_err(Failure::Build)?;
            let prepared = if config.sync.pull_unmanaged || config.sync.sheet_files {
                self.session()?
                    .take_over(&root, config.target.codepage, &config.sync)
                    .map_err(Failure::Sync)?
            } else {
                Prepared::default()
            };
            self.announce(&prepared);
            let built = compile(&root).map_err(Failure::Build)?;
            let forms = Forms {
                forms: &built.forms,
                codepage: built.codepage,
            };
            let report = self
                .session()?
                .apply(
                    &built.files,
                    forms,
                    &ras_sync::options(&root, &prepared, overwrite),
                )
                .map_err(Failure::Sync)?;
            ras_form::import::finish(&prepared).map_err(Failure::Build)?;
            // What was written, to tell later what was added in the VBE (spec §14.7.3).
            let mut not_written = report.drifted.clone();
            not_written.extend(report.pending.iter().cloned());
            not_written.extend(prepared.pending.iter().map(|p| p.generated.clone()));
            ras_sync::save_written(&root, &built.files, &not_written).map_err(Failure::Build)?;
            Ok(report)
        })();
        self.retry = matches!(result, Err(Failure::Sync(SyncError::Busy(_))));
        match &result {
            Ok(report) => {
                self.pending = 0;
                self.drift = report.drifted.len();
                self.message = Some(describe(report));
            }
            Err(e) => {
                self.pending = self.pending.max(1);
                self.message = Some(e.message());
            }
        }
        result
    }

    /// Tells the user what a take-over wrote into the project (the files appear in the editor).
    fn announce(&mut self, prepared: &Prepared) {
        // Only when something was written or could not be: what is pending is in the result of every sync.
        let Some((level, message)) = describe_take_over(prepared) else {
            return;
        };
        if prepared.imported.is_empty()
            && prepared.created.is_empty()
            && prepared.appended.is_empty()
            && self.last_take_over.as_ref() == Some(&message)
        {
            return;
        }
        self.last_take_over = Some(message.clone());
        (self.send)(Message::Notification(Notification::new(
            "window/showMessage".into(),
            json!({ "type": level, "message": message }),
        )));
        (self.send)(Message::Notification(Notification::new(
            "window/logMessage".into(),
            json!({ "type": 3, "message": prepared.messages().join("\n") }),
        )));
    }

    /// A sync nobody asked for: after a save, when the workbook was opened, or as a retry.
    fn auto_sync(&mut self, saved: bool) {
        let Some(root) = self.root.clone() else {
            return;
        };
        if !read_config(&root).is_ok_and(|c| c.sync.enabled) {
            self.publish();
            return;
        }
        if saved {
            self.pending += 1;
        }
        match self.sync(false) {
            // Expected states; the status bar shows them.
            Ok(_)
            | Err(Failure::Build(_))
            | Err(Failure::Sync(
                SyncError::Busy(_) | SyncError::NotOpen(_) | SyncError::Unsupported,
            )) => self.last_warning = None,
            Err(Failure::Sync(SyncError::Failed(message))) => {
                // Something the user has to act on (trust setting, a hand-written module of the same name …): say it once.
                if self.last_warning.as_ref() != Some(&message) {
                    (self.send)(Message::Notification(Notification::new(
                        "window/showMessage".into(),
                        json!({ "type": 2, "message": format!("RAS: 同期できませんでした。{message}") }),
                    )));
                    self.last_warning = Some(message);
                }
            }
        }
        self.publish();
    }

    /// Looks whether the workbook was opened or closed since the last look.
    fn tick(&mut self) {
        if self.root.is_none() {
            return;
        }
        let was_connected = self.connected;
        match self.session().map(|s| s.connect()) {
            Ok(Ok(())) => {
                self.connected = true;
                if !was_connected || self.retry {
                    return self.auto_sync(false);
                }
            }
            Ok(Err(e)) => {
                self.connected = false;
                self.retry = false;
                self.drift = 0;
                self.message = Some(e.to_string());
            }
            Err(e) => {
                self.connected = false;
                self.message = Some(e.message());
            }
        }
        self.publish();
    }

    /// Sends `ras/syncStatus` when something changed.
    fn publish(&mut self) {
        let workbook = self
            .session
            .as_ref()
            .map(|s| s.book().to_string_lossy().into_owned());
        let status = json!({
            "connected": self.connected,
            "workbook": workbook,
            "pending": if self.connected { self.pending } else { 0 },
            "drift": self.drift,
            "message": self.message,
        });
        if self.last_status.as_ref() != Some(&status) {
            (self.send)(Message::Notification(Notification::new(
                "ras/syncStatus".into(),
                status.clone(),
            )));
            self.last_status = Some(status);
        }
    }
}

#[cfg(test)]
mod sheet_tests {
    use super::*;

    #[test]
    fn sheet_commands_from_their_arguments() {
        use ras_ovba::{Place, SheetEdit};
        let (edit, force) = sheet_edit(
            &json!({ "action": "add", "name": "集計", "codeName": "", "after": "Sheet1" }),
        )
        .unwrap();
        assert_eq!(
            edit,
            Some(SheetEdit::Add {
                name: "集計".into(),
                code_name: None,
                place: Place::After("Sheet1".into())
            })
        );
        assert!(!force);
        let (edit, force) =
            sheet_edit(&json!({ "action": "remove", "sheet": "Sheet2", "force": true })).unwrap();
        assert_eq!(
            (edit, force),
            (
                Some(SheetEdit::Remove {
                    sheet: "Sheet2".into()
                }),
                true
            )
        );
        assert_eq!(
            sheet_edit(&json!({ "action": "list" })),
            Some((None, false))
        );
        // What is missing or unknown is no command.
        assert_eq!(sheet_edit(&json!({ "action": "add" })), None);
        assert_eq!(
            sheet_edit(&json!({ "action": "paint", "sheet": "Sheet1" })),
            None
        );
    }

    #[test]
    fn a_sheet_command_changes_the_workbook_file_and_the_project() {
        let root = std::env::temp_dir().join(format!("ras-lsp-sheet-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&root);
        std::fs::create_dir_all(root.join("src")).unwrap();
        std::fs::write(
            root.join("ras.toml"),
            "[project]\nname = \"T\"\nworkbook = \"Book.xlsm\"\n",
        )
        .unwrap();
        std::fs::write(
            root.join("src/main.ras"),
            "export function hello() {\n    msgBox(\"hi\");\n}\n",
        )
        .unwrap();
        std::fs::write(
            root.join("Book.xlsm"),
            ras_ovba::create_workbook(ras_ovba::BookKind::Xlsm, &["Sheet1"]),
        )
        .unwrap();
        let (edit, force) = sheet_edit(&json!({ "action": "add", "name": "集計" })).unwrap();
        let result = run_sheet_command(&root, edit.as_ref(), force);
        assert_eq!(
            (
                result["ok"].as_bool(),
                result["codeName"].as_str(),
                result["live"].as_bool()
            ),
            (Some(true), Some("Sheet2"), Some(false)),
            "{result}"
        );
        assert_eq!(result["created"][0], "objects/Sheet2.ras");
        assert!(root.join("objects/Sheet2.ras").is_file());
        let list = run_sheet_command(&root, None, false);
        assert_eq!(list["sheets"].as_array().map(Vec::len), Some(2), "{list}");
        assert_eq!(list["sheets"][1]["file"], "objects/Sheet2.ras");
        // A refusal says why.
        let (edit, force) = sheet_edit(&json!({ "action": "add", "name": "集計" })).unwrap();
        let result = run_sheet_command(&root, edit.as_ref(), force);
        assert_eq!(result["ok"], false);
        assert!(
            result["message"].as_str().unwrap().contains("既にあります"),
            "{result}"
        );
        let _ = std::fs::remove_dir_all(&root);
    }
}
