//! `ras run`: the debug run (spec §14.10) from the command line, on the stand-in for Excel.
//! Also what `ras test` shares with it: finding a JavaScript run time and the `ras_book` addon.

use std::path::{Path, PathBuf};
use std::process::{Command, ExitCode, Stdio};

use serde_json::{json, Value as Json};

use ras_workbook::diff;

/// A program that runs JavaScript as Node does.
pub struct Node {
    pub program: PathBuf,
    /// VS Code's executable: it runs as Node with `ELECTRON_RUN_AS_NODE`.
    pub electron: bool,
}

impl Node {
    pub fn command(&self) -> Command {
        let mut c = Command::new(&self.program);
        if self.electron {
            c.env("ELECTRON_RUN_AS_NODE", "1");
        }
        c
    }
}

fn runs(program: &Path, electron: bool) -> bool {
    let mut c = Command::new(program);
    if electron {
        c.env("ELECTRON_RUN_AS_NODE", "1");
    }
    c.arg("--version").stdin(Stdio::null()).stdout(Stdio::null()).stderr(Stdio::null()).status().is_ok_and(|s| s.success())
}

/// The run time: `RAS_NODE`, then `node` on PATH, then the one VS Code comes with.
pub fn find_node() -> Result<Node, String> {
    if let Some(p) = std::env::var_os("RAS_NODE").filter(|p| !p.is_empty()) {
        let program = PathBuf::from(p);
        if runs(&program, false) {
            return Ok(Node { program, electron: false });
        }
        if runs(&program, true) {
            return Ok(Node { program, electron: true });
        }
        return Err(format!("RAS_NODE の {} を実行できません", program.display()));
    }
    if runs(Path::new("node"), false) {
        return Ok(Node { program: PathBuf::from("node"), electron: false });
    }
    let mut candidates: Vec<PathBuf> = Vec::new();
    if cfg!(windows) {
        for var in ["LOCALAPPDATA", "ProgramFiles"] {
            if let Some(base) = std::env::var_os(var) {
                let base = PathBuf::from(base);
                candidates.push(base.join("Programs").join("Microsoft VS Code").join("Code.exe"));
                candidates.push(base.join("Microsoft VS Code").join("Code.exe"));
            }
        }
    } else if cfg!(target_os = "macos") {
        candidates.push(PathBuf::from("/Applications/Visual Studio Code.app/Contents/MacOS/Electron"));
        if let Some(home) = std::env::var_os("HOME") {
            candidates.push(PathBuf::from(home).join("Applications/Visual Studio Code.app/Contents/MacOS/Electron"));
        }
    } else {
        candidates.extend(["/usr/share/code/code", "/usr/lib/code/code", "/opt/visual-studio-code/code", "/snap/code/current/usr/share/code/code"].map(PathBuf::from));
    }
    for program in candidates {
        if program.is_file() && runs(&program, true) {
            return Ok(Node { program, electron: true });
        }
    }
    Err("JavaScript のランタイムが見つかりません。Node.js をインストールして PATH を通すか、環境変数 RAS_NODE に node（または VS Code の実行ファイル）のパスを指定してください".into())
}

/// The `ras_book` addon: `RAS_BOOK_ADDON`, or `ras_book.node` next to this program. In a
/// development build the library Cargo built is next to it instead, and is copied to that name.
pub fn find_book_addon() -> Result<PathBuf, String> {
    if let Some(p) = std::env::var_os("RAS_BOOK_ADDON").filter(|p| !p.is_empty()) {
        let p = PathBuf::from(p);
        return if p.is_file() { Ok(p) } else { Err(format!("RAS_BOOK_ADDON の {} が見つかりません", p.display())) };
    }
    let exe = std::env::current_exe().map_err(|e| e.to_string())?;
    let dir = exe.parent().unwrap_or(Path::new("."));
    let addon = dir.join("ras_book.node");
    let built = dir.join(if cfg!(windows) {
        "ras_book.dll"
    } else if cfg!(target_os = "macos") {
        "libras_book.dylib"
    } else {
        "libras_book.so"
    });
    if built.is_file() {
        let newer = |a: &Path, b: &Path| match (a.metadata().and_then(|m| m.modified()), b.metadata().and_then(|m| m.modified())) {
            (Ok(x), Ok(y)) => x > y,
            _ => true,
        };
        if !addon.is_file() || newer(&built, &addon) {
            // Another `ras` may be copying it at the same time: copy aside, then rename.
            let tmp = dir.join(format!("ras_book.node.{}", std::process::id()));
            if std::fs::copy(&built, &tmp).is_ok() && std::fs::rename(&tmp, &addon).is_err() {
                let _ = std::fs::remove_file(&tmp);
            }
        }
    }
    if addon.is_file() {
        return Ok(addon);
    }
    Err(format!(
        "ブックの代役のアドオン ras_book.node が {} にありません。配布物の ras_book.node を ras と同じフォルダーに置くか、環境変数 RAS_BOOK_ADDON でパスを指定してください",
        dir.display()
    ))
}

/// Writes generated JavaScript into `dir`, emptied first (stale modules would still be found).
pub fn write_js(dir: &Path, files: &[ras_compiler::JsFile]) -> Result<(), String> {
    if dir.exists() {
        std::fs::remove_dir_all(dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    }
    for f in files {
        let path = dir.join(&f.rel_path);
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent).map_err(|e| format!("{}: {e}", parent.display()))?;
        }
        std::fs::write(&path, &f.text).map_err(|e| format!("{}: {e}", path.display()))?;
    }
    Ok(())
}

/// `vbYes`, `yes` or a number: the value `msgBox` returns.
pub fn msgbox_answer(text: &str) -> Result<i64, String> {
    const NAMES: &[(&str, i64)] = &[("ok", 1), ("cancel", 2), ("abort", 3), ("retry", 4), ("ignore", 5), ("yes", 6), ("no", 7)];
    let lower = text.trim().to_lowercase();
    let key = lower.strip_prefix("vb").unwrap_or(&lower);
    if let Some((_, v)) = NAMES.iter().find(|(n, _)| *n == key) {
        return Ok(*v);
    }
    lower.parse::<i64>().map_err(|_| format!("--msgbox `{text}` は指定できません（vbOK / vbCancel / vbAbort / vbRetry / vbIgnore / vbYes / vbNo か数値）"))
}

fn msgbox_name(v: i64) -> String {
    match v {
        1 => "vbOK".into(),
        2 => "vbCancel".into(),
        3 => "vbAbort".into(),
        4 => "vbRetry".into(),
        5 => "vbIgnore".into(),
        6 => "vbYes".into(),
        7 => "vbNo".into(),
        n => n.to_string(),
    }
}

#[derive(clap::Args, Clone, Default)]
pub struct RunArgs {
    /// 実行する関数（`Module.proc`。モジュール名を省いた `proc` も、1 つに決まれば使える）
    pub entry: Option<String>,
    /// 実行できる関数の一覧を出力する
    #[arg(long)]
    pub list: bool,
    /// 結果を JSON で出力する
    #[arg(long)]
    pub json: bool,
    /// msgBox の応答（vbYes など。複数指定すると表示された順に使う。指定がなければ vbOK）
    #[arg(long, value_name = "RESULT")]
    pub msgbox: Vec<String>,
    /// inputBox の応答（複数指定すると表示された順に使う。指定がなければ既定値、既定値もなければキャンセル）
    #[arg(long, value_name = "TEXT")]
    pub inputbox: Vec<String>,
    /// getOpenFilename の応答（ファイルのパス。false でキャンセル）
    #[arg(long, value_name = "PATH")]
    pub open_file: Vec<String>,
    /// getSaveAsFilename の応答（ファイルのパス。false でキャンセル）
    #[arg(long, value_name = "PATH")]
    pub save_file: Vec<String>,
    /// 手書き VBA（application.run で呼ぶマクロ）を RAS の関数で置き換える（例: --replace-run Legacy=Mocks.legacy）
    #[arg(long, value_name = "MACRO=MODULE.FUNC")]
    pub replace_run: Vec<String>,
    /// COM オブジェクト（createObject / new）を RAS のクラスで置き換える（例: --replace-object ADODB.Connection=Mocks.FakeConnection）
    #[arg(long, value_name = "PROGID=MODULE.CLASS")]
    pub replace_object: Vec<String>,
    /// 実行後のブックをグリッドで表示する HTML（build/view/）を作る。変更されたセルを強調する
    #[arg(long)]
    pub view: bool,
    /// 変更されたセルの差分を出力しない
    #[arg(long)]
    pub no_diff: bool,
}

/// The changed cells a run prints at most.
const MAX_CELLS: usize = 50;

pub fn run(root: &Path, out: &ras_compiler::JsOutput, args: &RunArgs) -> Result<ExitCode, String> {
    if args.list {
        if args.json {
            println!("{}", serde_json::to_string(&out.entries).unwrap_or_default());
        } else {
            for e in &out.entries {
                println!("{e}");
            }
        }
        return Ok(ExitCode::SUCCESS);
    }
    let entry = match &args.entry {
        Some(e) => e.clone(),
        None if out.entries.len() == 1 => out.entries[0].clone(),
        None if out.entries.is_empty() => return Err("実行できる関数がありません（引数なしの関数か、トップレベルの文が必要です）".into()),
        None => return Err(format!("実行する関数を指定してください（候補: {}）", out.entries.join(", "))),
    };
    let mut msgbox = Vec::new();
    for m in &args.msgbox {
        msgbox.push(msgbox_answer(m)?);
    }
    let node = find_node()?;
    let addon = find_book_addon()?;
    let dir = root.join("build").join("debug");
    let report_path = dir.join("ras-run.json");
    let _ = std::fs::remove_file(&report_path);
    let pairs = |list: &[String], option: &str| -> Result<serde_json::Map<String, Json>, String> {
        let mut out = serde_json::Map::new();
        for item in list {
            let (from, to) = item.split_once('=').filter(|(a, b)| !a.is_empty() && b.contains('.')).ok_or_else(|| format!("{option} `{item}` は `名前=モジュール名.関数名` の形で指定してください"))?;
            out.insert(from.to_string(), json!(to));
        }
        Ok(out)
    };
    let answers = json!({
        "msgbox": msgbox, "inputbox": args.inputbox, "openfile": args.open_file, "savefile": args.save_file,
        "run": pairs(&args.replace_run, "--replace-run")?, "createObject": pairs(&args.replace_object, "--replace-object")?,
    });
    let mut cmd = node.command();
    cmd.arg("--enable-source-maps").arg(dir.join("ras-main.js")).arg(&entry).current_dir(root);
    cmd.env("RAS_BOOK_ADDON", &addon).env("RAS_RUN_REPORT", &report_path).env("RAS_ANSWERS", answers.to_string());
    cmd.env_remove("RAS_COM_ADDON").env_remove("RAS_MOCK_HOST").env_remove("RAS_ENTRY");
    cmd.stdin(Stdio::null());
    if args.json {
        cmd.env("RAS_RUN_QUIET", "1").stdout(Stdio::piped()).stderr(Stdio::piped());
    }
    // Without --json the program's output is shown as it runs.
    let launch = |e: std::io::Error| format!("{} を起動できません: {e}", node.program.display());
    let (status, stderr) = if args.json {
        let o = cmd.output().map_err(launch)?;
        (o.status, o.stderr)
    } else {
        (cmd.status().map_err(launch)?, Vec::new())
    };
    let report: Option<Json> = std::fs::read_to_string(&report_path).ok().and_then(|t| serde_json::from_str(&t).ok());
    let Some(mut report) = report else {
        // The process ended without reporting (it crashed, or the launcher failed).
        eprint!("{}", String::from_utf8_lossy(&stderr));
        return Err(format!("実行が結果を残さずに終了しました（終了コード {}）", status.code().map(|c| c.to_string()).unwrap_or_else(|| "なし".into())));
    };
    let _ = std::fs::remove_file(&report_path);
    // What the program changed: each saved workbook against the file it was read from.
    let mut diffs: Vec<(String, String, Result<diff::Diff, String>)> = Vec::new();
    if !args.no_diff {
        for w in report["workbooks"].as_array().cloned().unwrap_or_default() {
            let (Some(file), Some(source)) = (w["file"].as_str(), w["source"].as_str()) else { continue };
            diffs.push((source.to_string(), file.to_string(), crate::book::compare(&root.join(source), &root.join(file))));
        }
    }
    if args.json {
        if let Some(list) = report["workbooks"].as_array_mut() {
            for w in list {
                let file = w["file"].as_str().unwrap_or_default().to_string();
                match diffs.iter().find(|d| d.1 == file) {
                    Some((_, _, Ok(d))) => w["diff"] = diff::to_json(d),
                    Some((_, _, Err(e))) => w["diffError"] = json!(e),
                    None => {}
                }
            }
        }
        if let Some(o) = report.as_object_mut() {
            if let Some(Json::Object(e)) = o.get_mut("error") {
                e.remove("stack");
            }
        }
        println!("{}", serde_json::to_string_pretty(&report).unwrap_or_default());
    } else {
        if let Some(e) = report["error"].as_object() {
            let at = match (e.get("file").and_then(Json::as_str), e.get("line").and_then(Json::as_i64)) {
                (Some(f), Some(l)) if l > 0 => format!("（{f}:{l}）"),
                _ => String::new(),
            };
            eprintln!("実行時エラー {}: {}{at}", e.get("code").and_then(Json::as_i64).unwrap_or(0), e.get("description").and_then(Json::as_str).unwrap_or_default());
            // The RAS frames of the stack, when there are more than the one above.
            let frames: Vec<&str> = e.get("stack").and_then(Json::as_str).unwrap_or_default().lines().filter(|l| l.trim_start().starts_with("at ")).collect();
            if frames.len() > 1 {
                for f in frames {
                    eprintln!("{f}");
                }
            }
        }
        let unanswered = report["dialogs"].as_array().map(|d| d.iter().filter(|x| x["answered"] == json!(false) && x["kind"] != json!("getOpenFilename") && x["kind"] != json!("getSaveAsFilename")).count()).unwrap_or(0);
        if unanswered > 0 {
            println!("ダイアログ {unanswered} 個に既定の応答をしました（--msgbox {} / --inputbox <文字列> で応答を指定できます）", msgbox_name(6));
        }
        for (source, file, d) in &diffs {
            match d {
                Ok(d) if d.is_empty() => println!("{source} → {file}: セルの変更はありません"),
                Ok(d) => {
                    println!("変更されたセル（{source} → {file}）:");
                    print!("{}", diff::text(d, MAX_CELLS));
                }
                Err(e) => eprintln!("warning: {file} の差分を取れませんでした: {e}"),
            }
        }
    }
    if args.view && !args.json {
        for w in report["workbooks"].as_array().cloned().unwrap_or_default() {
            let Some(file) = w["file"].as_str() else { continue };
            let before = w["source"].as_str().map(|s| root.join(s));
            match crate::book::write_view(&root.join(file), before.as_deref(), false, None, &root.join("build").join("view")) {
                Ok(page) => println!("グリッド表示: {}", page.strip_prefix(root).unwrap_or(&page).display()),
                Err(e) => eprintln!("warning: {file} のグリッド表示を作れませんでした: {e}"),
            }
        }
    }
    Ok(if report["ok"] == json!(true) { ExitCode::SUCCESS } else { ExitCode::from(1) })
}
