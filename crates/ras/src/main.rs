//! `ras` — the RAS toolchain in a single binary (spec §7.3).

use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use clap::{Parser, Subcommand};
use ras_compiler::vba::VModuleKind;
use ras_compiler::{BuildOptions, BuildOutput, Config, Diagnostic, ModuleKind, Program, Severity};

#[derive(Parser)]
#[command(name = "ras", version, about = "RAS: TypeScript ライクな構文から VBA へトランスパイルする Better VBA")]
struct Cli {
    /// プロジェクトのルート（ras.toml のあるディレクトリ）。省略時はカレントディレクトリから上へ探す
    #[arg(long, global = true, value_name = "DIR")]
    project: Option<PathBuf>,
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// ras.toml と雛形を生成する
    Init {
        /// 作成先ディレクトリ
        #[arg(default_value = ".")]
        dir: PathBuf,
        /// 空のマクロ有効ブックも作成する（例: --workbook Book.xlsm）
        #[arg(long, value_name = "FILE")]
        workbook: Option<String>,
    },
    /// 型チェック・リントのみ行う
    Check,
    /// VBA ファイルを build/ に出力する
    Build(BuildArgs),
    /// 変更を監視して再ビルドする（ブックが閉じていればオフライン書き込みも行う）
    Watch(BuildArgs),
    /// Excel を起動せず、ブックファイル内の VBA プロジェクトを直接書き換える
    Pack {
        /// 対象ブック（省略時は ras.toml の project.workbook）
        book: Option<PathBuf>,
        #[command(flatten)]
        build: BuildArgs,
        /// .xlsx を .xlsm に変換して出力する（元ファイルは書き換えない）
        #[arg(long)]
        convert: bool,
        /// 出力先（省略時は対象ブックを置き換える）
        #[arg(long, value_name = "FILE")]
        output: Option<PathBuf>,
        /// VBE 上で編集された RAS 管理モジュール（ドリフト）を上書きする
        #[arg(long)]
        overwrite: bool,
        /// 不可視の Excel 経由で書き込む（未実装）
        #[arg(long)]
        via_excel: bool,
    },
    /// Excel を起動せず、ブックファイルから VBA ソースを vba/ に取り出す
    Unpack {
        /// 対象ブック（省略時は ras.toml の project.workbook）
        book: Option<PathBuf>,
    },
    /// ソースを整形する
    Fmt {
        /// 対象ファイル（省略時は src/ 以下のすべての .ras）
        files: Vec<PathBuf>,
        /// 書き換えず、整形が必要なファイルがあれば失敗する
        #[arg(long)]
        check: bool,
    },
    /// LSP サーバとして起動する（VS Code 拡張から呼ばれる）
    Lsp {
        #[arg(long)]
        stdio: bool,
    },
    /// 1回だけ Excel と同期する（ライブ同期。未実装）
    Sync {
        #[arg(long)]
        pull: bool,
    },
    /// VBA プロジェクトへのアクセス許可と RASBridge のインストール（未実装）
    Setup,
    /// COM タイプライブラリから型定義（.d.ras）を生成する（未実装）
    Typegen { typelib: String },
}

#[derive(clap::Args, Clone, Default)]
struct BuildArgs {
    /// 最適化レベル既定 2、行番号（Erl）なしでビルドする
    #[arg(long)]
    release: bool,
    /// 最適化レベル（0: 展開なし / 1: ループ内自動展開 / 2: 積極的）
    #[arg(short = 'O', value_name = "LEVEL")]
    opt: Option<u8>,
    /// 生成 VBA に行番号を付ける / 付けない（既定は ras.toml の build.erl）
    #[arg(long, value_name = "BOOL")]
    erl: Option<bool>,
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    match run(cli) {
        Ok(code) => code,
        Err(msg) => {
            eprintln!("error: {msg}");
            ExitCode::from(2)
        }
    }
}

fn find_root(explicit: Option<&Path>) -> Result<PathBuf, String> {
    if let Some(p) = explicit {
        return Ok(p.to_path_buf());
    }
    let cwd = std::env::current_dir().map_err(|e| e.to_string())?;
    let mut dir: &Path = &cwd;
    loop {
        if dir.join("ras.toml").is_file() {
            return Ok(dir.to_path_buf());
        }
        match dir.parent() {
            Some(p) => dir = p,
            None => return Err("ras.toml が見つかりません。`ras init` でプロジェクトを作成するか、--project で指定してください".into()),
        }
    }
}

fn workbook_path(root: &Path, cfg: &Config, explicit: Option<&Path>) -> Option<PathBuf> {
    let p = explicit.map(Path::to_path_buf).or_else(|| cfg.project.workbook.as_ref().map(PathBuf::from))?;
    Some(if p.is_absolute() { p } else { root.join(p) })
}

/// Loads and analyses the project. Sheet code names come from the workbook when it can be read.
fn analyze(root: &Path, args: &BuildArgs, book: Option<&Path>) -> Result<Program, String> {
    let (mut cfg, sources) = ras_compiler::load_project(root)?;
    cfg.opt_override = args.opt.or(if args.release { Some(cfg.build.opt_level.max(2)) } else { None });
    if let Some(book) = workbook_path(root, &cfg, book) {
        if book.is_file() {
            match ras_ovba::workbook_code_names(&book) {
                Ok(names) => cfg.sheet_code_names = names.sheets.into_iter().map(|s| s.code_name).collect(),
                Err(e) => eprintln!("warning: {} のシート名を読み取れませんでした: {e}", book.display()),
            }
        }
    }
    Ok(ras_compiler::analyze(sources, cfg))
}

/// Prints diagnostics (with the offending source line) and returns the number of errors.
fn report(prog: &Program, diags: &[Diagnostic], show_hints: bool) -> usize {
    let mut items: Vec<&Diagnostic> = diags
        .iter()
        .filter(|d| show_hints || d.severity <= Severity::Warning)
        .filter(|d| d.module.is_none_or(|m| prog.modules[m].kind == ModuleKind::User) || d.is_error())
        .collect();
    items.sort_by_key(|d| (d.module, u32::from(d.span.start())));
    for d in &items {
        let sev = match d.severity {
            Severity::Error => "error",
            Severity::Warning => "warning",
            Severity::Info => "info",
            Severity::Hint => "hint",
        };
        eprintln!("{sev}[{}]: {}", d.code, d.message);
        if let Some(m) = d.module {
            let md = &prog.modules[m];
            let start = u32::from(d.span.start());
            let (line, col) = md.line_col_utf16(start);
            eprintln!("  --> {}:{}:{}", md.rel_path, line + 1, col + 1);
            let ls = md.line_starts[line as usize] as usize;
            let le = md.text[ls..].find('\n').map_or(md.text.len(), |i| ls + i);
            let src = md.text[ls..le].trim_end_matches('\r');
            let prefix: String = md.text[ls..(start as usize).min(le)].chars().map(|c| if c == '\t' { '\t' } else { ' ' }).collect();
            let width = md.text[(start as usize).min(le)..(u32::from(d.span.end()) as usize).clamp(start as usize, le)].chars().count().max(1);
            eprintln!("   |\n{:>3}| {src}\n   | {prefix}{}", line + 1, "^".repeat(width));
        }
        eprintln!();
    }
    items.iter().filter(|d| d.is_error()).count()
}

fn encode(text: &str, codepage: u16) -> Vec<u8> {
    let enc = match codepage {
        932 => encoding_rs::SHIFT_JIS,
        936 => encoding_rs::GBK,
        949 => encoding_rs::EUC_KR,
        950 => encoding_rs::BIG5,
        874 => encoding_rs::WINDOWS_874,
        1250 => encoding_rs::WINDOWS_1250,
        1251 => encoding_rs::WINDOWS_1251,
        1253 => encoding_rs::WINDOWS_1253,
        1254 => encoding_rs::WINDOWS_1254,
        1255 => encoding_rs::WINDOWS_1255,
        1256 => encoding_rs::WINDOWS_1256,
        1257 => encoding_rs::WINDOWS_1257,
        1258 => encoding_rs::WINDOWS_1258,
        _ => encoding_rs::WINDOWS_1252,
    };
    enc.encode(text).0.into_owned()
}

/// Compiles the project; `None` when there were errors.
fn compile(root: &Path, args: &BuildArgs, book: Option<&Path>) -> Result<Option<(Program, BuildOutput)>, String> {
    let prog = analyze(root, args, book)?;
    let errors = report(&prog, &prog.diagnostics, false);
    if errors > 0 {
        eprintln!("{errors} 件のエラーがあります");
        return Ok(None);
    }
    let out = ras_compiler::build(&prog, &BuildOptions { release: args.release, erl: args.erl, no_attributes: false });
    let errors = report(&prog, &out.diagnostics, false);
    if errors > 0 {
        eprintln!("{errors} 件のエラーがあります");
        return Ok(None);
    }
    Ok(Some((prog, out)))
}

/// Writes the build products to `build/`, removing generated files of modules that no longer exist.
fn write_build(root: &Path, prog: &Program, out: &BuildOutput) -> Result<PathBuf, String> {
    let dir = root.join("build");
    std::fs::create_dir_all(&dir).map_err(|e| format!("{}: {e}", dir.display()))?;
    let mut keep: HashSet<String> = HashSet::new();
    for f in &out.files {
        // Module files are stored in the target code page, as the VBE expects (spec §13.7).
        let path = dir.join(&f.file_name);
        std::fs::write(&path, encode(&f.full_text, prog.config.target.codepage)).map_err(|e| format!("{}: {e}", path.display()))?;
        keep.insert(f.file_name.clone());
        let map = format!("{}.rasmap", f.name);
        std::fs::write(dir.join(&map), f.rasmap()).map_err(|e| e.to_string())?;
        keep.insert(map);
    }
    if let Ok(entries) = std::fs::read_dir(&dir) {
        for e in entries.flatten() {
            let name = e.file_name().to_string_lossy().to_string();
            let generated = [".bas", ".cls", ".doccls", ".rasmap"].iter().any(|ext| name.ends_with(ext));
            if generated && !keep.contains(&name) {
                let _ = std::fs::remove_file(e.path());
            }
        }
    }
    Ok(dir)
}

fn cmd_build(root: &Path, args: &BuildArgs) -> Result<ExitCode, String> {
    let Some((prog, out)) = compile(root, args, None)? else { return Ok(ExitCode::from(1)) };
    let dir = write_build(root, &prog, &out)?;
    println!("{} 個のモジュールを {} に出力しました", out.files.len(), dir.display());
    Ok(ExitCode::SUCCESS)
}

fn to_ovba_kind(k: &VModuleKind) -> ras_ovba::ModuleKind {
    match k {
        VModuleKind::Standard => ras_ovba::ModuleKind::Standard,
        VModuleKind::Class => ras_ovba::ModuleKind::Class,
        VModuleKind::Document => ras_ovba::ModuleKind::Document,
    }
}

struct PackArgs<'a> {
    book: Option<&'a Path>,
    convert: bool,
    output: Option<&'a Path>,
    overwrite: bool,
}

fn cmd_pack(root: &Path, args: &BuildArgs, p: &PackArgs) -> Result<ExitCode, String> {
    let (cfg, _) = ras_compiler::load_project(root)?;
    let Some(book) = workbook_path(root, &cfg, p.book) else {
        return Err("対象ブックが指定されていません。`ras pack <book>` か ras.toml の project.workbook で指定してください".into());
    };
    if !book.is_file() {
        return Err(format!("{} が見つかりません（`ras init --workbook` で空のブックを作成できます）", book.display()));
    }
    let Some((prog, out)) = compile(root, args, Some(&book))? else { return Ok(ExitCode::from(1)) };
    write_build(root, &prog, &out)?;

    // Modules that RAS generated earlier but that no longer exist are removed; hand-edited ones are reported (spec §14.7).
    let existing = ras_ovba::unpack(&book).map_err(|e| format!("{}: {e}", book.display()))?;
    let generated: HashSet<&str> = out.files.iter().map(|f| f.name.as_str()).collect();
    let mut remove = Vec::new();
    let mut drifted = Vec::new();
    if let Some(project) = &existing.project {
        for m in &project.modules {
            let Some(is_drifted) = ras_compiler::emit::has_drifted(&m.source) else { continue };
            if is_drifted {
                drifted.push(m.name.clone());
            }
            if !generated.contains(m.name.as_str()) && matches!(m.kind, ras_ovba::ModuleKind::Standard | ras_ovba::ModuleKind::Class) {
                remove.push(m.name.clone());
            }
        }
    }
    if !drifted.is_empty() && !p.overwrite {
        return Err(format!(
            "次の RAS 管理モジュールが VBE 上で編集されています（ドリフト）: {}\n  RAS の出力で上書きするには --overwrite を付けてください。VBE 側の内容は `ras unpack` で build/drift/ に保存できます",
            drifted.join(", ")
        ));
    }
    let mut modules = Vec::new();
    let mut document_code = Vec::new();
    for f in &out.files {
        match f.kind {
            VModuleKind::Document => document_code.push(ras_ovba::DocumentCode { code_name: f.name.clone(), managed_code: f.code.clone() }),
            _ => modules.push(ras_ovba::VbaModule { name: f.name.clone(), kind: to_ovba_kind(&f.kind), source: f.stream_text() }),
        }
    }
    let req = ras_ovba::PackRequest {
        book: book.clone(),
        output: p.output.map(Path::to_path_buf),
        convert: p.convert,
        codepage: prog.config.target.codepage,
        project_name: None,
        modules,
        document_code,
        remove_modules: remove,
        backup_dir: Some(root.join("build").join("backup")),
        backup_generations: prog.config.pack.backup_generations,
    };
    match ras_ovba::pack(&req) {
        Ok(rep) => {
            for w in &rep.warnings {
                eprintln!("warning: {w}");
            }
            println!(
                "{} に書き込みました（作成 {} / 更新 {} / 削除 {} / 変更なし {}）",
                rep.output.display(),
                rep.created.len(),
                rep.updated.len(),
                rep.removed.len(),
                rep.unchanged.len()
            );
            Ok(ExitCode::SUCCESS)
        }
        Err(ras_ovba::OvbaError::Locked) => Err(format!(
            "{} は Excel で開かれているため、オフライン書き込みはできません。ブックを閉じてから再実行してください（開いたままの反映にはライブ同期が必要ですが、この版では未実装です）",
            book.display()
        )),
        Err(e) => Err(format!("{}: {e}", book.display())),
    }
}

fn cmd_unpack(root: &Path, book: Option<&Path>) -> Result<ExitCode, String> {
    let (cfg, _) = ras_compiler::load_project(root)?;
    let Some(book) = workbook_path(root, &cfg, book) else {
        return Err("対象ブックが指定されていません".into());
    };
    let unpacked = ras_ovba::unpack(&book).map_err(|e| format!("{}: {e}", book.display()))?;
    let Some(project) = unpacked.project else {
        println!("{} に VBA プロジェクトはありません", book.display());
        return Ok(ExitCode::SUCCESS);
    };
    for name in &project.stomped_modules {
        eprintln!("warning: モジュール `{name}` はソースと P-code が一致しません（VBA stomping の疑い）。ソース側を正として取り出します");
    }
    let vba_dir = root.join("vba");
    let drift_dir = root.join("build").join("drift");
    let (mut written, mut managed, mut drifted) = (0, 0, 0);
    for m in &project.modules {
        // The Attribute header is not part of the code pane.
        let code: String = m.source.replace("\r\n", "\n").lines().skip_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION ")).map(|l| format!("{l}\n")).collect();
        match ras_compiler::emit::has_drifted(&m.source) {
            Some(false) => managed += 1,
            Some(true) => {
                // RAS-managed module edited in the VBE: keep the VBE side for the user to compare.
                std::fs::create_dir_all(&drift_dir).map_err(|e| e.to_string())?;
                std::fs::write(drift_dir.join(format!("{}.vba", m.name)), &code).map_err(|e| e.to_string())?;
                eprintln!("warning: RAS 管理モジュール `{}` は VBE 上で編集されています。VBE 側の内容を build/drift/{}.vba に保存しました", m.name, m.name);
                drifted += 1;
            }
            None => {
                let ext = match m.kind {
                    ras_ovba::ModuleKind::Standard => "bas",
                    ras_ovba::ModuleKind::Form => "frm",
                    _ => "cls",
                };
                // Document modules: only the hand-written part outside the RAS-managed region.
                let text = if m.kind == ras_ovba::ModuleKind::Document { ras_ovba::merge_managed_region(&m.source, "").replace("\r\n", "\n") } else { m.source.replace("\r\n", "\n") };
                if m.kind == ras_ovba::ModuleKind::Document && text.lines().all(|l| l.starts_with("Attribute ") || l.trim().is_empty()) {
                    continue;
                }
                std::fs::create_dir_all(&vba_dir).map_err(|e| e.to_string())?;
                std::fs::write(vba_dir.join(format!("{}.{ext}", m.name)), text).map_err(|e| e.to_string())?;
                written += 1;
            }
        }
    }
    println!("非管理モジュール {written} 個を {} に取り出しました（RAS 管理 {managed} 個、ドリフト {drifted} 個）", vba_dir.display());
    Ok(ExitCode::SUCCESS)
}

fn cmd_init(dir: &Path, workbook: Option<&str>) -> Result<ExitCode, String> {
    let toml_path = dir.join("ras.toml");
    if toml_path.exists() {
        return Err(format!("{} は既に存在します", toml_path.display()));
    }
    std::fs::create_dir_all(dir.join("src")).map_err(|e| e.to_string())?;
    let name: String = dir
        .canonicalize()
        .ok()
        .and_then(|p| p.file_name().map(|n| n.to_string_lossy().to_string()))
        .unwrap_or_else(|| "RASProject".into())
        .chars()
        .filter(|c| c.is_alphanumeric() || *c == '_')
        .collect();
    let name = if name.is_empty() { "RASProject".to_string() } else { name };
    let wb_line = match workbook {
        Some(w) => format!("workbook = \"{w}\""),
        None => "# workbook = \"Book.xlsm\"    # `ras pack` / 同期の対象ブック".to_string(),
    };
    let toml = format!(
        "[project]\nname = \"{name}\"\n{wb_line}\n\n[target]\nhost = \"excel\"\nvba6 = true             # VBA6（32bit Office 2007 以前）用の Declare も出力\ncodepage = 932\n\n[build]\nopt-level = 1           # 0: 展開なし / 1: ループ内自動展開 / 2: 積極的\nerl = true              # 行番号付与（エラー位置を RAS のソース行で表示）\n\n[pack]\nbackup-generations = 5\n\n[doevents]\nenabled = true\ninterval-ms = 100\nloop-threshold = 10000\n"
    );
    std::fs::write(&toml_path, toml).map_err(|e| e.to_string())?;
    let main = "/**\n * サンプルのマクロ。Excel の「マクロ」ダイアログから `hello` を実行できる。\n */\nexport function hello() {\n    const names = [\"RAS\", \"VBA\"];\n    const greeting = names.map(n => `Hello, ${n}!`).join(\"\\n\");\n    MsgBox(greeting);\n}\n";
    let main_path = dir.join("src").join("main.ras");
    if !main_path.exists() {
        std::fs::write(&main_path, main).map_err(|e| e.to_string())?;
    }
    let gitignore = dir.join(".gitignore");
    if !gitignore.exists() {
        std::fs::write(&gitignore, "/build/\n~$*\n").map_err(|e| e.to_string())?;
    }
    if let Some(w) = workbook {
        let path = dir.join(w);
        if path.exists() {
            eprintln!("warning: {} は既に存在するため作成しません", path.display());
        } else {
            let kind = if w.to_lowercase().ends_with(".xlam") { ras_ovba::BookKind::Xlam } else { ras_ovba::BookKind::Xlsm };
            std::fs::write(&path, ras_ovba::create_workbook(kind, &["Sheet1"])).map_err(|e| e.to_string())?;
        }
    }
    println!("{} を作成しました。`ras build` で VBA を生成できます", toml_path.display());
    Ok(ExitCode::SUCCESS)
}

fn cmd_fmt(root: Option<&Path>, files: &[PathBuf], check: bool) -> Result<ExitCode, String> {
    let mut targets: Vec<PathBuf> = files.to_vec();
    if targets.is_empty() {
        let root = find_root(root)?;
        let (_, sources) = ras_compiler::load_project(&root)?;
        targets = sources.into_iter().filter(|s| s.kind == ModuleKind::User).map(|s| s.path).collect();
    }
    let (mut changed, mut failed) = (0, 0);
    for path in &targets {
        let text = std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
        match ras_syntax::format(&text) {
            None => {
                eprintln!("{}: 構文エラーがあるため整形しません", path.display());
                failed += 1;
            }
            Some(f) if f != text => {
                changed += 1;
                if check {
                    println!("{}: 整形が必要です", path.display());
                } else {
                    std::fs::write(path, f).map_err(|e| format!("{}: {e}", path.display()))?;
                    println!("{}: 整形しました", path.display());
                }
            }
            Some(_) => {}
        }
    }
    Ok(if failed > 0 || (check && changed > 0) { ExitCode::from(1) } else { ExitCode::SUCCESS })
}

fn cmd_watch(root: &Path, args: &BuildArgs) -> Result<ExitCode, String> {
    use notify::{RecursiveMode, Watcher};
    let (tx, rx) = std::sync::mpsc::channel();
    let mut watcher = notify::recommended_watcher(move |res: notify::Result<notify::Event>| {
        if let Ok(ev) = res {
            let _ = tx.send(ev);
        }
    })
    .map_err(|e| e.to_string())?;
    for sub in ["src", "types", "ras.toml"] {
        let p = root.join(sub);
        if p.exists() {
            watcher.watch(&p, RecursiveMode::Recursive).map_err(|e| e.to_string())?;
        }
    }
    let rebuild = || {
        let (cfg, _) = match ras_compiler::load_project(root) {
            Ok(x) => x,
            Err(e) => return eprintln!("error: {e}"),
        };
        let has_book = workbook_path(root, &cfg, None).is_some_and(|b| b.is_file());
        let result = if has_book && cfg.sync.enabled {
            // The workbook is updated directly while it is closed (spec §14.9); live sync is not implemented.
            cmd_pack(root, args, &PackArgs { book: None, convert: false, output: None, overwrite: false })
        } else {
            cmd_build(root, args)
        };
        if let Err(e) = result {
            eprintln!("error: {e}");
        }
    };
    rebuild();
    println!("変更を監視しています（Ctrl+C で終了）");
    while let Ok(first) = rx.recv() {
        // Debounce: editors write files in several steps.
        let mut relevant = is_source_event(&first);
        let deadline = std::time::Instant::now() + std::time::Duration::from_millis(150);
        while let Ok(ev) = rx.recv_timeout(deadline.saturating_duration_since(std::time::Instant::now())) {
            relevant |= is_source_event(&ev);
        }
        if relevant {
            rebuild();
        }
    }
    Ok(ExitCode::SUCCESS)
}

fn is_source_event(ev: &notify::Event) -> bool {
    !matches!(ev.kind, notify::EventKind::Access(_)) && ev.paths.iter().any(|p| p.extension().is_some_and(|e| e == "ras" || e == "toml"))
}

fn unimplemented_cmd(what: &str, detail: &str) -> Result<ExitCode, String> {
    Err(format!("`{what}` はこの版では未実装です。{detail}"))
}

fn run(cli: Cli) -> Result<ExitCode, String> {
    let project = cli.project.as_deref();
    match &cli.command {
        Command::Init { dir, workbook } => cmd_init(dir, workbook.as_deref()),
        Command::Check => {
            let root = find_root(project)?;
            let prog = analyze(&root, &BuildArgs::default(), None)?;
            let errors = report(&prog, &prog.diagnostics, true);
            if errors > 0 {
                eprintln!("{errors} 件のエラーがあります");
                return Ok(ExitCode::from(1));
            }
            // Code generation can report further problems (size limits, unsupported constructs).
            let out = ras_compiler::build(&prog, &BuildOptions::default());
            let errors = report(&prog, &out.diagnostics, true);
            if errors > 0 {
                eprintln!("{errors} 件のエラーがあります");
                return Ok(ExitCode::from(1));
            }
            println!("問題は見つかりませんでした");
            Ok(ExitCode::SUCCESS)
        }
        Command::Build(args) => cmd_build(&find_root(project)?, args),
        Command::Watch(args) => cmd_watch(&find_root(project)?, args),
        Command::Pack { book, build, convert, output, overwrite, via_excel } => {
            if *via_excel {
                return unimplemented_cmd("ras pack --via-excel", "Excel を使わないオフライン書き込み（--via-excel なし）を使ってください。");
            }
            cmd_pack(&find_root(project)?, build, &PackArgs { book: book.as_deref(), convert: *convert, output: output.as_deref(), overwrite: *overwrite })
        }
        Command::Unpack { book } => cmd_unpack(&find_root(project)?, book.as_deref()),
        Command::Fmt { files, check } => cmd_fmt(project, files, *check),
        Command::Lsp { .. } => {
            ras_lsp::run_stdio().map_err(|e| e.to_string())?;
            Ok(ExitCode::SUCCESS)
        }
        Command::Sync { .. } => unimplemented_cmd("ras sync", "ブックを閉じた状態で `ras pack` を使うと、Excel を起動せずに VBA プロジェクトへ反映できます。"),
        Command::Setup => unimplemented_cmd("ras setup", "ライブ同期（RASBridge）が未実装のため、設定するものがありません。"),
        Command::Typegen { .. } => unimplemented_cmd("ras typegen", "同梱の Excel / VBA / Office の型定義は使えます。その他の COM ライブラリは types/*.d.ras に手書きしてください。"),
    }
}
