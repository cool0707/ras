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
use ras_sync::{Report, Session, SyncError, SyncOptions};
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
}

pub type Send = Box<dyn Fn(Message) + std::marker::Send>;

/// Starts the sync thread. It ends when the returned sender is dropped.
pub fn spawn(send: Send) -> Sender<Job> {
    let (tx, rx) = std::sync::mpsc::channel();
    std::thread::Builder::new()
        .name("ras-sync".into())
        .spawn(move || Worker { send, root: None, session: None, connected: false, pending: 0, drift: 0, retry: false, message: None, last_status: None, last_warning: None }.run(rx))
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

/// Compiles what is saved on disk, as `ras build` does, for `CodeModule.AddFromString`.
fn compile(root: &Path) -> Result<Vec<OutputFile>, String> {
    let (config, sources) = crate::load_project(root)?;
    let prog = ras_compiler::analyze(sources, config);
    let out = ras_compiler::build(&prog, &BuildOptions { release: false, erl: None, no_attributes: true });
    let errors = prog.diagnostics.iter().chain(&out.diagnostics).filter(|d| d.is_error()).count();
    if errors > 0 {
        return Err(format!("ビルド エラーが {errors} 件あるため、同期していません"));
    }
    Ok(out.files)
}

fn describe(report: &Report) -> String {
    let mut text = if report.changed() > 0 { format!("同期しました（{}）。VBA のグローバル変数はリセットされます", report.summary()) } else { "ブックは最新です".to_string() };
    if !report.drifted.is_empty() {
        text.push_str(&format!("。VBE 上で編集されたため上書きしていないモジュール: {}", report.drifted.join(", ")));
    }
    text
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
                        Ok(report) => json!({ "ok": true, "message": describe(&report), "drifted": report.drifted }),
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
                                Ok(report) => json!({ "ok": true, "message": format!("Excel で開きました。{}", describe(&report)) }),
                                Err(e) => json!({ "ok": true, "message": format!("Excel で開きました。{}", e.message()) }),
                            }
                        }
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
                Job::Run { id, name } => {
                    // The macro must be the one in the sources: sync first.
                    let result = match self.sync(false).and_then(|_| self.session()?.run_macro(&name).map_err(Failure::Sync)) {
                        Ok(()) => json!({ "ok": true }),
                        Err(e) => json!({ "ok": false, "message": e.message() }),
                    };
                    self.respond(id, result);
                }
            }
        }
    }

    fn respond(&mut self, id: RequestId, result: Value) {
        (self.send)(Message::Response(Response { id, result: Some(result), error: None }));
        self.publish();
    }

    /// The session for the project's workbook.
    fn session(&mut self) -> Result<&mut Session, Failure> {
        let root = self.root.clone().ok_or_else(|| Failure::Build("ras.toml のあるフォルダーを開いてください".into()))?;
        let book = read_config(&root).and_then(|c| book_path(&root, &c)).map_err(Failure::Build)?;
        if self.session.as_ref().is_none_or(|s| s.book() != std::path::absolute(&book).unwrap_or(book.clone())) {
            self.session = Some(Session::new(&book));
            self.connected = false;
        }
        Ok(self.session.as_mut().expect("just set"))
    }

    /// Connects, compiles and writes. Records the outcome for the status bar.
    fn sync(&mut self, overwrite: bool) -> Result<Report, Failure> {
        let result = (|| {
            let root = self.root.clone().ok_or_else(|| Failure::Build("ras.toml のあるフォルダーを開いてください".into()))?;
            let connected = self.session()?.connect();
            self.connected = connected.is_ok();
            connected.map_err(Failure::Sync)?;
            let files = compile(&root).map_err(Failure::Build)?;
            self.session()?.apply(&files, &SyncOptions { overwrite }).map_err(Failure::Sync)
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

    /// A sync nobody asked for: after a save, when the workbook was opened, or as a retry.
    fn auto_sync(&mut self, saved: bool) {
        let Some(root) = self.root.clone() else { return };
        if !read_config(&root).is_ok_and(|c| c.sync.enabled) {
            self.publish();
            return;
        }
        if saved {
            self.pending += 1;
        }
        match self.sync(false) {
            // Expected states; the status bar shows them.
            Ok(_) | Err(Failure::Build(_)) | Err(Failure::Sync(SyncError::Busy(_) | SyncError::NotOpen(_) | SyncError::Unsupported)) => self.last_warning = None,
            Err(Failure::Sync(SyncError::Failed(message))) => {
                // Something the user has to act on (trust setting, a hand-written module of the same name …): say it once.
                if self.last_warning.as_ref() != Some(&message) {
                    (self.send)(Message::Notification(Notification::new("window/showMessage".into(), json!({ "type": 2, "message": format!("RAS: 同期できませんでした。{message}") }))));
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
        let workbook = self.session.as_ref().map(|s| s.book().to_string_lossy().into_owned());
        let status = json!({
            "connected": self.connected,
            "workbook": workbook,
            "pending": if self.connected { self.pending } else { 0 },
            "drift": self.drift,
            "message": self.message,
        });
        if self.last_status.as_ref() != Some(&status) {
            (self.send)(Message::Notification(Notification::new("ras/syncStatus".into(), status.clone())));
            self.last_status = Some(status);
        }
    }
}
