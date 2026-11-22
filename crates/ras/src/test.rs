//! `ras test` (spec §14.12): runs the test files under Node on the stand-in for Excel, one
//! process a file, and reports the results.

use std::collections::HashSet;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::process::{ExitCode, Stdio};
use std::time::{Duration, Instant};

use ras_compiler::testing::{self, Mode, TestFile, TestNode};
use ras_compiler::Program;
use serde_json::{json, Value as Json};

use crate::run::{find_book_addon, find_node, write_js, Node};

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
pub enum Reporter {
    Pretty,
    Json,
    Junit,
}

#[derive(clap::Args, Clone)]
pub struct TestArgs {
    /// 実行するテストファイル（またはフォルダー）。省略時はすべてのテストファイル
    pub files: Vec<PathBuf>,
    /// フルネーム（describe > test）に指定した文字列を含むテストだけを実行する
    #[arg(short = 't', long = "name", value_name = "TEXT")]
    pub name: Option<String>,
    /// 実行せずにテストの一覧を出力する
    #[arg(long)]
    pub list: bool,
    /// 出力の形式
    #[arg(long, value_enum, default_value_t = Reporter::Pretty)]
    pub reporter: Reporter,
    /// 変更を監視して、影響するテストファイルを実行し直す
    #[arg(long)]
    pub watch: bool,
    /// ブックの代役ではなく実際の Excel に対して実行する（未実装）
    #[arg(long)]
    pub excel: bool,
    /// test.only と、置き換えていないダイアログをエラーにする
    #[arg(long)]
    pub ci: bool,
    /// テストファイルを 1 つずつ実行する（現状は常に 1 つずつ実行する）
    #[arg(long)]
    pub sequence: bool,
    /// 失敗が N 件になったら打ち切る
    #[arg(long, value_name = "N", default_value_t = 0)]
    pub bail: u32,
}

/// The result of one test.
#[derive(Clone, Debug)]
struct TestResult {
    file: String,
    name: Vec<String>,
    line: u64,
    /// `passed` / `failed` / `skipped` / `todo`.
    status: String,
    duration_ms: u64,
    failure: Option<Json>,
    printed: Vec<String>,
    warnings: Vec<String>,
    workbook: Option<Json>,
}

impl TestResult {
    fn json(&self) -> Json {
        let mut o = json!({ "type": "test", "file": self.file, "name": self.name, "line": self.line, "status": self.status, "durationMs": self.duration_ms, "printed": self.printed });
        if let Some(f) = &self.failure {
            o["failure"] = f.clone();
        }
        if !self.warnings.is_empty() {
            o["warnings"] = json!(self.warnings);
        }
        if let Some(w) = &self.workbook {
            o["workbook"] = w.clone();
        }
        o
    }

    fn full_name(&self) -> String {
        self.name.join(" > ")
    }
}

fn strings(v: &Json) -> Vec<String> {
    v.as_array()
        .map(|a| {
            a.iter()
                .filter_map(|x| x.as_str().map(str::to_string))
                .collect()
        })
        .unwrap_or_default()
}

fn result_of(file: &str, e: &Json) -> TestResult {
    TestResult {
        file: file.to_string(),
        name: strings(&e["name"]),
        line: e["line"].as_u64().unwrap_or(0),
        status: e["status"].as_str().unwrap_or("failed").to_string(),
        duration_ms: e["durationMs"].as_u64().unwrap_or(0),
        failure: e.get("failure").cloned(),
        printed: strings(&e["printed"]),
        warnings: strings(&e["warnings"]),
        workbook: e.get("workbook").cloned(),
    }
}

/// The lines of a failure, as people read them.
fn failure_lines(f: &Json) -> Vec<String> {
    let mut out = Vec::new();
    let kind = f["kind"].as_str().unwrap_or("error");
    let head = match kind {
        "expect" => format!(
            "{}: {}",
            f["matcher"].as_str().unwrap_or("expect"),
            f["message"].as_str().unwrap_or_default()
        ),
        "timeout" => f["description"]
            .as_str()
            .unwrap_or("タイムアウトしました")
            .to_string(),
        "hook" => {
            let what = match (
                f["code"].as_i64(),
                f["description"].as_str(),
                f["message"].as_str(),
            ) {
                (Some(c), Some(d), _) => format!("実行時エラー {c}: {d}"),
                (_, _, Some(m)) => m.to_string(),
                _ => String::new(),
            };
            format!(
                "{} が失敗しました: {what}",
                f["hook"].as_str().unwrap_or("フック")
            )
        }
        _ => match f["code"].as_i64() {
            Some(c) => format!(
                "実行時エラー {c}: {}",
                f["description"].as_str().unwrap_or_default()
            ),
            None => f["description"].as_str().unwrap_or_default().to_string(),
        },
    };
    out.extend(head.lines().map(str::to_string));
    if let Some(l) = f["location"].as_object() {
        let col = l
            .get("column")
            .and_then(Json::as_u64)
            .filter(|c| *c > 0)
            .map(|c| format!(":{c}"))
            .unwrap_or_default();
        out.push(format!(
            "at {}:{}{col}",
            l.get("file").and_then(Json::as_str).unwrap_or_default(),
            l.get("line").and_then(Json::as_u64).unwrap_or(0)
        ));
    }
    out
}

fn xml(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            '&' => "&amp;".to_string(),
            '<' => "&lt;".to_string(),
            '>' => "&gt;".to_string(),
            '"' => "&quot;".to_string(),
            c if (c as u32) < 0x20 && c != '\n' && c != '\t' => " ".to_string(),
            c => c.to_string(),
        })
        .collect()
}

fn junit(results: &[TestResult], seconds: f64) -> String {
    let mut files: Vec<&str> = results.iter().map(|r| r.file.as_str()).collect();
    files.dedup();
    let count = |rs: &[&TestResult], s: &str| rs.iter().filter(|r| r.status == s).count();
    let all: Vec<&TestResult> = results.iter().collect();
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
    out.push_str(&format!(
        "<testsuites tests=\"{}\" failures=\"{}\" skipped=\"{}\" time=\"{seconds:.3}\">\n",
        all.len(),
        count(&all, "failed"),
        count(&all, "skipped") + count(&all, "todo")
    ));
    for file in files {
        let rs: Vec<&TestResult> = results.iter().filter(|r| r.file == file).collect();
        let time: u64 = rs.iter().map(|r| r.duration_ms).sum();
        out.push_str(&format!(
            "  <testsuite name=\"{}\" tests=\"{}\" failures=\"{}\" skipped=\"{}\" time=\"{:.3}\">\n",
            xml(file),
            rs.len(),
            count(&rs, "failed"),
            count(&rs, "skipped") + count(&rs, "todo"),
            time as f64 / 1000.0
        ));
        for r in rs {
            out.push_str(&format!(
                "    <testcase classname=\"{}\" name=\"{}\" time=\"{:.3}\"",
                xml(file),
                xml(&r.full_name()),
                r.duration_ms as f64 / 1000.0
            ));
            let body = match (r.status.as_str(), &r.failure) {
                ("failed", f) => {
                    let lines = f.as_ref().map(failure_lines).unwrap_or_default();
                    let mut text = lines.join("\n");
                    if !r.printed.is_empty() {
                        text.push_str(&format!("\ndebug.print:\n{}", r.printed.join("\n")));
                    }
                    format!(
                        "      <failure message=\"{}\">{}</failure>\n",
                        xml(lines.first().map(String::as_str).unwrap_or("failed")),
                        xml(&text)
                    )
                }
                ("skipped", _) => "      <skipped/>\n".to_string(),
                ("todo", _) => "      <skipped message=\"todo\"/>\n".to_string(),
                _ => String::new(),
            };
            if body.is_empty() {
                out.push_str("/>\n");
            } else {
                out.push_str(&format!(">\n{body}    </testcase>\n"));
            }
        }
        out.push_str("  </testsuite>\n");
    }
    out.push_str("</testsuites>\n");
    out
}

// ---------------------------------------------------------------------------- the list

fn node_json(n: &TestNode) -> Json {
    let mut o = json!({ "type": if n.suite { "describe" } else { "test" }, "name": n.name, "line": n.line, "column": n.column });
    if n.mode != Mode::Run {
        o["mode"] = json!(n.mode.as_str());
    }
    if n.template {
        o["each"] = json!(true);
    }
    if n.suite {
        o["children"] = n.children.iter().map(node_json).collect();
    }
    o
}

fn print_tree(nodes: &[TestNode], depth: usize) {
    for n in nodes {
        let note = match n.mode {
            Mode::Run => "",
            Mode::Skip => "（skip）",
            Mode::Only => "（only）",
            Mode::Todo => "（todo）",
        };
        let each = if n.template {
            "（each: 実行時にケースが決まります）"
        } else {
            ""
        };
        println!("{}{}{note}{each}", "  ".repeat(depth + 1), n.name);
        print_tree(&n.children, depth + 1);
    }
}

fn list(files: &[&TestFile], reporter: Reporter) {
    if reporter == Reporter::Json {
        let out: Vec<Json> = files.iter().map(|f| json!({ "file": f.file, "tests": f.nodes.iter().map(node_json).collect::<Vec<_>>() })).collect();
        println!(
            "{}",
            serde_json::to_string(&json!({ "files": out })).unwrap_or_default()
        );
        return;
    }
    for f in files {
        println!("{}", f.file);
        print_tree(&f.nodes, 0);
    }
}

// ---------------------------------------------------------------------------- running a file

struct Session<'a> {
    root: &'a Path,
    node: &'a Node,
    addon: &'a Path,
    args: &'a TestArgs,
    timeout_ms: u64,
    save_workbooks: &'a str,
}

/// What running a test file gave: its results, and what kept it from running at all.
struct FileRun {
    results: Vec<TestResult>,
    fatal: Option<String>,
}

fn run_file(
    s: &Session,
    file: &str,
    index: usize,
    bail_left: u32,
    on_result: &mut dyn FnMut(&TestResult),
) -> Result<FileRun, String> {
    let dir = s.root.join("build").join("test");
    let events = dir.join(format!(".events-{index}.jsonl"));
    let log = dir.join(format!(".output-{index}.log"));
    let _ = std::fs::remove_file(&events);
    let stem = Path::new(file)
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();
    let stem = stem.strip_suffix(".ras").unwrap_or(&stem).to_string();
    let options = json!({
        "filter": s.args.name,
        "ci": s.args.ci,
        "bail": bail_left,
        "saveWorkbooks": s.save_workbooks,
        "workbooksDir": format!("build/test/workbooks/{stem}"),
        "timeoutMs": s.timeout_ms,
    });
    let out = std::fs::File::create(&log).map_err(|e| format!("{}: {e}", log.display()))?;
    let err = out.try_clone().map_err(|e| e.to_string())?;
    let mut cmd = s.node.command();
    cmd.arg("--enable-source-maps")
        .arg(dir.join("ras-test-main.js"))
        .arg(file)
        .current_dir(s.root);
    cmd.env("RAS_BOOK_ADDON", s.addon)
        .env("RAS_TEST_EVENTS", &events)
        .env("RAS_TEST_OPTIONS", options.to_string());
    cmd.env_remove("RAS_COM_ADDON")
        .env_remove("RAS_MOCK_HOST")
        .env_remove("RAS_RUN_REPORT");
    cmd.stdin(Stdio::null()).stdout(out).stderr(err);
    let mut child = cmd
        .spawn()
        .map_err(|e| format!("{} を起動できません: {e}", s.node.program.display()))?;

    let mut run = FileRun {
        results: Vec::new(),
        fatal: None,
    };
    let mut collected: Vec<(Vec<String>, u64)> = Vec::new();
    let mut reported: HashSet<Vec<String>> = HashSet::new();
    let mut current: Option<(Vec<String>, u64, Instant, u64)> = None;
    let mut offset = 0u64;
    let mut pending = String::new();
    let mut done = false;
    let mut timed_out = false;
    let mut read_events = |run: &mut FileRun,
                           current: &mut Option<(Vec<String>, u64, Instant, u64)>,
                           collected: &mut Vec<(Vec<String>, u64)>,
                           reported: &mut HashSet<Vec<String>>,
                           done: &mut bool| {
        let Ok(mut f) = std::fs::File::open(&events) else {
            return;
        };
        if f.seek(SeekFrom::Start(offset)).is_err() {
            return;
        }
        let mut buf = Vec::new();
        if f.read_to_end(&mut buf).is_err() {
            return;
        }
        offset += buf.len() as u64;
        pending.push_str(&String::from_utf8_lossy(&buf));
        while let Some(at) = pending.find('\n') {
            let line: String = pending.drain(..=at).collect();
            let Ok(e) = serde_json::from_str::<Json>(line.trim()) else {
                continue;
            };
            match e["type"].as_str().unwrap_or_default() {
                "collected" => {
                    *collected = e["tests"]
                        .as_array()
                        .map(|a| {
                            a.iter()
                                .map(|t| (strings(&t["name"]), t["line"].as_u64().unwrap_or(0)))
                                .collect()
                        })
                        .unwrap_or_default();
                }
                "start" => {
                    *current = Some((
                        strings(&e["name"]),
                        e["line"].as_u64().unwrap_or(0),
                        Instant::now(),
                        e["timeoutMs"].as_u64().unwrap_or(s.timeout_ms),
                    ))
                }
                "test" => {
                    *current = None;
                    let r = result_of(file, &e);
                    reported.insert(r.name.clone());
                    on_result(&r);
                    run.results.push(r);
                }
                "hookError" => {
                    let mut name = strings(&e["name"]);
                    name.push("afterAll".into());
                    let r = TestResult {
                        file: file.to_string(),
                        name,
                        line: 0,
                        status: "failed".into(),
                        duration_ms: 0,
                        failure: e.get("failure").cloned(),
                        printed: Vec::new(),
                        warnings: Vec::new(),
                        workbook: None,
                    };
                    on_result(&r);
                    run.results.push(r);
                }
                "fatal" => {
                    let at = match (e["file"].as_str(), e["line"].as_u64()) {
                        (Some(f), Some(l)) if l > 0 => format!("（{f}:{l}）"),
                        _ => String::new(),
                    };
                    run.fatal = Some(format!("{}{at}", e["message"].as_str().unwrap_or_default()));
                }
                "done" => *done = true,
                _ => {}
            }
        }
    };
    let status = loop {
        read_events(
            &mut run,
            &mut current,
            &mut collected,
            &mut reported,
            &mut done,
        );
        if let Some(status) = child.try_wait().map_err(|e| e.to_string())? {
            read_events(
                &mut run,
                &mut current,
                &mut collected,
                &mut reported,
                &mut done,
            );
            break status;
        }
        if let Some((_, _, started, limit)) = &current {
            // JavaScript cannot be stopped from inside: the process is ended (spec §14.12.4).
            if started.elapsed() > Duration::from_millis(*limit) {
                timed_out = true;
                let _ = child.kill();
                let status = child.wait().map_err(|e| e.to_string())?;
                read_events(
                    &mut run,
                    &mut current,
                    &mut collected,
                    &mut reported,
                    &mut done,
                );
                break status;
            }
        }
        std::thread::sleep(Duration::from_millis(5));
    };
    let _ = std::fs::remove_file(&events);
    let output = std::fs::read_to_string(&log).unwrap_or_default();
    let _ = std::fs::remove_file(&log);
    if !done && run.fatal.is_none() {
        // The process ended before its tests did: a time-out, or it crashed.
        let mut extra: Vec<TestResult> = Vec::new();
        let mut push = |name: Vec<String>, line: u64, failure: Json, duration_ms: u64| {
            extra.push(TestResult {
                file: file.to_string(),
                name,
                line,
                status: "failed".into(),
                duration_ms,
                failure: Some(failure),
                printed: Vec::new(),
                warnings: Vec::new(),
                workbook: None,
            });
        };
        let code = status
            .code()
            .map(|c| c.to_string())
            .unwrap_or_else(|| "なし".into());
        let why = if timed_out {
            "タイムアウトでテストファイルの実行を打ち切ったため、実行されませんでした".to_string()
        } else {
            format!("テストの実行が途中で終了したため、実行されませんでした（終了コード {code}）")
        };
        if let Some((name, line, started, limit)) = current.take() {
            reported.insert(name.clone());
            let failure = if timed_out {
                json!({ "kind": "timeout", "description": format!("{limit} ミリ秒以内に終わりませんでした（タイムアウト）") })
            } else {
                json!({ "kind": "error", "description": format!("テストの実行中にプロセスが終了しました（終了コード {code}）") })
            };
            push(name, line, failure, started.elapsed().as_millis() as u64);
        }
        if collected.is_empty() && !timed_out {
            let tail: Vec<&str> = output
                .lines()
                .rev()
                .take(20)
                .collect::<Vec<_>>()
                .into_iter()
                .rev()
                .collect();
            run.fatal = Some(format!(
                "テストを実行できませんでした（終了コード {code}）\n{}",
                tail.join("\n")
            ));
        }
        for (name, line) in collected {
            if !reported.contains(&name) {
                push(
                    name,
                    line,
                    json!({ "kind": if timed_out { "timeout" } else { "error" }, "description": why }),
                    0,
                );
            }
        }
        for r in extra {
            on_result(&r);
            run.results.push(r);
        }
    }
    Ok(run)
}

// ---------------------------------------------------------------------------- the command

/// The test files `args` selects: those given (files, or folders holding them), or all.
fn select<'a>(
    root: &Path,
    files: &'a [TestFile],
    given: &[PathBuf],
) -> Result<Vec<&'a TestFile>, String> {
    if given.is_empty() {
        return Ok(files.iter().collect());
    }
    let cwd = std::env::current_dir().map_err(|e| e.to_string())?;
    let norm = |p: &Path| p.canonicalize().unwrap_or_else(|_| p.to_path_buf());
    let root_abs = norm(root);
    let mut out: Vec<&TestFile> = Vec::new();
    for g in given {
        // As typed from the current directory, or from the project root.
        let candidates = [norm(&cwd.join(g)), norm(&root.join(g))];
        let before = out.len();
        for f in files {
            let path = root_abs.join(&f.file);
            if candidates.iter().any(|c| path == *c || path.starts_with(c))
                && !out.iter().any(|x| x.file == f.file)
            {
                out.push(f);
            }
        }
        if out.len() == before
            && !files.iter().any(|f| {
                candidates
                    .iter()
                    .any(|c| root_abs.join(&f.file) == *c || root_abs.join(&f.file).starts_with(c))
            })
        {
            return Err(format!("{} にテストファイルがありません（テストファイルは ras.toml の [test] include に一致する .ras ファイルです）", g.display()));
        }
    }
    Ok(out)
}

fn status_mark(status: &str) -> &'static str {
    match status {
        "passed" => "✓",
        "failed" => "✗",
        "todo" => "□",
        _ => "-",
    }
}

fn print_result(r: &TestResult) {
    let time = if r.status == "passed" || r.status == "failed" {
        format!(" ({} ms)", r.duration_ms)
    } else {
        String::new()
    };
    let note = match r.status.as_str() {
        "skipped" => "（スキップ）",
        "todo" => "（未実装）",
        _ => "",
    };
    println!("  {} {}{note}{time}", status_mark(&r.status), r.full_name());
    if r.status == "failed" {
        if let Some(f) = &r.failure {
            for l in failure_lines(f) {
                println!("      {l}");
            }
        }
        if !r.printed.is_empty() {
            println!("      debug.print:");
            for l in &r.printed {
                println!("        {l}");
            }
        }
        match &r.workbook {
            Some(Json::String(w)) => println!("      ブック: {w}"),
            Some(Json::Array(ws)) => println!(
                "      ブック: {}",
                ws.iter()
                    .filter_map(Json::as_str)
                    .collect::<Vec<_>>()
                    .join(", ")
            ),
            _ => {}
        }
    }
    for w in &r.warnings {
        println!("      警告: {w}");
    }
}

/// Runs the selected test files once. Returns the exit code.
fn run_once(
    root: &Path,
    prog: &Program,
    files: &[&TestFile],
    args: &TestArgs,
    node: &Node,
    addon: &Path,
) -> Result<u8, String> {
    if args.ci {
        if let Some(f) = files.iter().find(|f| f.has_only()) {
            return Err(format!("{}: `only` の付いたテストがあります（--ci では `test.only` / `describe.only` を使えません）", f.file));
        }
    }
    let out = ras_compiler::build_js(prog, "../../");
    if out.diagnostics.iter().any(|d| d.is_error()) {
        let errors = crate::report(prog, &out.diagnostics, false);
        return Err(format!("{errors} 件のエラーがあります"));
    }
    write_js(&root.join("build").join("test"), &out.files)?;
    let session = Session {
        root,
        node,
        addon,
        args,
        timeout_ms: prog.config.test.timeout_ms,
        save_workbooks: &prog.config.test.save_workbooks,
    };
    let started = Instant::now();
    let mut results: Vec<TestResult> = Vec::new();
    let mut fatal = false;
    for (index, f) in files.iter().enumerate() {
        let failed = results.iter().filter(|r| r.status == "failed").count() as u32;
        if args.bail > 0 && failed >= args.bail {
            break;
        }
        match args.reporter {
            Reporter::Pretty => println!("{}", f.file),
            Reporter::Json => println!("{}", json!({ "type": "file", "file": f.file })),
            Reporter::Junit => {}
        }
        let reporter = args.reporter;
        let mut on_result = |r: &TestResult| match reporter {
            Reporter::Pretty => print_result(r),
            Reporter::Json => println!("{}", r.json()),
            Reporter::Junit => {}
        };
        let run = run_file(
            &session,
            &f.file,
            index,
            if args.bail > 0 { args.bail - failed } else { 0 },
            &mut on_result,
        )?;
        if let Some(message) = run.fatal {
            fatal = true;
            match args.reporter {
                Reporter::Json => println!(
                    "{}",
                    json!({ "type": "error", "file": f.file, "message": message })
                ),
                _ => eprintln!("error: {}: {message}", f.file),
            }
        }
        if args.reporter == Reporter::Pretty && run.results.is_empty() {
            println!("  （実行するテストはありません）");
        }
        results.extend(run.results);
    }
    let count = |s: &str| results.iter().filter(|r| r.status == s).count();
    let (passed, failed, skipped, todo) = (
        count("passed"),
        count("failed"),
        count("skipped"),
        count("todo"),
    );
    let seconds = started.elapsed().as_secs_f64();
    match args.reporter {
        Reporter::Pretty => {
            let mut parts = vec![format!("成功 {passed}")];
            for (n, label) in [(failed, "失敗"), (skipped, "スキップ"), (todo, "未実装")] {
                if n > 0 {
                    parts.push(format!("{label} {n}"));
                }
            }
            println!(
                "\nテスト {} 件: {}（{seconds:.2} 秒）",
                results.len(),
                parts.join("、")
            );
        }
        Reporter::Json => println!(
            "{}",
            json!({ "type": "summary", "files": files.len(), "tests": results.len(), "passed": passed, "failed": failed, "skipped": skipped, "todo": todo, "durationMs": (seconds * 1000.0) as u64 })
        ),
        Reporter::Junit => print!("{}", junit(&results, seconds)),
    }
    Ok(if fatal {
        2
    } else if failed > 0 {
        1
    } else {
        0
    })
}

/// The test files that (transitively) import one of the changed files, or are one of them.
fn affected<'a>(
    prog: &Program,
    files: &[&'a TestFile],
    changed: &HashSet<PathBuf>,
) -> Vec<&'a TestFile> {
    let norm = |p: &Path| p.canonicalize().unwrap_or_else(|_| p.to_path_buf());
    let changed: HashSet<PathBuf> = changed.iter().map(|p| norm(p)).collect();
    let hit: HashSet<usize> = (0..prog.modules.len())
        .filter(|&m| changed.contains(&norm(&prog.modules[m].path)))
        .collect();
    // A change of something that is no module (ras.toml, a workbook, a form's design): everything.
    if hit.len() < changed.len() {
        return files.to_vec();
    }
    files
        .iter()
        .copied()
        .filter(|f| {
            let mut seen = HashSet::new();
            let mut stack = vec![f.module];
            while let Some(m) = stack.pop() {
                if hit.contains(&m) {
                    return true;
                }
                if seen.insert(m) {
                    stack.extend(prog.modules[m].imports.iter().copied());
                }
            }
            false
        })
        .collect()
}

pub fn cmd(root: &Path, args: &TestArgs) -> Result<ExitCode, String> {
    // Loads and checks the project with its tests; `None` (after printing) when it has errors.
    let load = || -> Result<Option<Program>, String> {
        let prog = crate::analyze_with_tests(root)?;
        let errors = crate::report(&prog, &prog.diagnostics, false);
        if errors > 0 {
            eprintln!("{errors} 件のエラーがあります");
            return Ok(None);
        }
        Ok(Some(prog))
    };
    let Some(prog) = load()? else {
        return Ok(ExitCode::from(2));
    };
    let all = testing::list(&prog);
    let files = select(root, &all, &args.files)?;
    if args.list {
        list(&files, args.reporter);
        return Ok(ExitCode::SUCCESS);
    }
    if args.excel || prog.config.test.excel == "com" {
        return Err("実際の Excel に対するテストの実行（--excel、[test] excel = \"com\"）にはまだ対応していません。ブックの代役で実行するには、--excel を外してください".into());
    }
    if files.is_empty() {
        println!("テストファイルがありません（test/**/*.test.ras に置きます。場所は ras.toml の [test] include で変えられます）");
        return Ok(ExitCode::SUCCESS);
    }
    let node = find_node()?;
    let addon = find_book_addon()?;
    let code = run_once(root, &prog, &files, args, &node, &addon)?;
    if !args.watch {
        return Ok(ExitCode::from(code));
    }

    use notify::{RecursiveMode, Watcher};
    let (tx, rx) = std::sync::mpsc::channel();
    let mut watcher = notify::recommended_watcher(move |res: notify::Result<notify::Event>| {
        if let Ok(ev) = res {
            let _ = tx.send(ev);
        }
    })
    .map_err(|e| e.to_string())?;
    for sub in ["src", "types", "forms", "test", "ras.toml"] {
        let p = root.join(sub);
        if p.exists() {
            watcher
                .watch(&p, RecursiveMode::Recursive)
                .map_err(|e| e.to_string())?;
        }
    }
    println!("変更を監視しています（Ctrl+C で終了）");
    while let Ok(first) = rx.recv() {
        // Debounce: editors write files in several steps.
        let mut changed: HashSet<PathBuf> = HashSet::new();
        let mut take = |ev: notify::Event| {
            if !matches!(ev.kind, notify::EventKind::Access(_)) {
                changed.extend(ev.paths.into_iter().filter(|p| {
                    p.extension().is_some_and(|e| {
                        ["ras", "toml", "json", "xlsx", "xlsm", "csv", "txt"]
                            .iter()
                            .any(|k| e.eq_ignore_ascii_case(k))
                    })
                }));
            }
        };
        take(first);
        let deadline = Instant::now() + Duration::from_millis(150);
        while let Ok(ev) = rx.recv_timeout(deadline.saturating_duration_since(Instant::now())) {
            take(ev);
        }
        if changed.is_empty() {
            continue;
        }
        let prog = match load() {
            Ok(Some(p)) => p,
            Ok(None) => continue,
            Err(e) => {
                eprintln!("error: {e}");
                continue;
            }
        };
        let all = testing::list(&prog);
        let files = match select(root, &all, &args.files) {
            Ok(f) => affected(&prog, &f, &changed),
            Err(e) => {
                eprintln!("error: {e}");
                continue;
            }
        };
        if files.is_empty() {
            continue;
        }
        println!();
        if let Err(e) = run_once(root, &prog, &files, args, &node, &addon) {
            eprintln!("error: {e}");
        }
    }
    Ok(ExitCode::SUCCESS)
}
