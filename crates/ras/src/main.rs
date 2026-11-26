//! `ras` — the RAS toolchain in a single binary (spec §7.3).

use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use clap::{Parser, Subcommand};

mod book;
mod explain;
mod forms;
mod run;
mod sheet;
mod test;
use ras_compiler::vba::VModuleKind;
use ras_compiler::{BuildOptions, BuildOutput, Config, Diagnostic, ModuleKind, Program, Severity};

#[derive(Parser)]
#[command(
    name = "ras",
    version,
    about = "RAS: TypeScript ライクな構文から VBA へトランスパイルする Better VBA"
)]
struct Cli {
    /// プロジェクトのルート（ras.toml のあるディレクトリ）。省略時はカレントディレクトリから上へ探す
    #[arg(long, global = true, value_name = "DIR")]
    project: Option<PathBuf>,
    /// 診断（エラー・警告）の出力形式。json は 1 行 1 件の JSON（コード・重大度・メッセージ・範囲・修正内容）を標準出力に出す
    #[arg(long, global = true, value_enum, default_value_t = MessageFormat::Human, value_name = "FORMAT")]
    message_format: MessageFormat,
    #[command(subcommand)]
    command: Command,
}

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
enum MessageFormat {
    Human,
    Json,
}

/// `--message-format json`: diagnostics are written as JSON Lines.
static JSON_MESSAGES: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);

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
    Check {
        /// 機械的に直せる警告（Excel / VBA のメンバー名の大文字小文字など）をソースに書き戻す
        #[arg(long)]
        fix: bool,
        /// stand-in: Excel を使わない実行（ras run・ras test・Excel のないデバッグ実行）のブックの代役で使えないものも警告する
        #[arg(long, value_enum, value_name = "TARGET")]
        target: Option<CheckTarget>,
    },
    /// VBA ファイルを build/ に出力する（--target js でデバッグ用の JavaScript を build/debug/ に出力する）
    Build {
        #[command(flatten)]
        build: BuildArgs,
        /// 出力形式（vba: VBA モジュール / js: デバッグ実行用の JavaScript とソースマップ）
        #[arg(long, value_enum, default_value_t = Target::Vba)]
        target: Target,
    },
    /// 変更を監視して再ビルドする（ブックが Excel で開いていればライブ同期、閉じていればオフライン書き込みを行う）
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
    /// ユーザーフォーム（forms/<Name>.form.json と forms/<Name>.ras）を扱う
    Form {
        #[command(subcommand)]
        action: FormCommand,
    },
    /// 関数を Excel なしで実行する（デバッグ実行と同じ JavaScript ランタイムとブックの代役。出力・エラー・変更されたセルを表示する）
    Run(run::RunArgs),
    /// 構文と Excel の操作の早見表（TypeScript との違い、型、関数、エラー処理、テスト、コマンド）を表示する
    Cheatsheet,
    /// エラー・警告のコード（E0301 など）の説明、典型的な原因、直し方を表示する
    Explain {
        /// コード（省略するとコードの一覧）
        code: Option<String>,
    },
    /// テスト（test/**/*.test.ras）を Excel なしで実行する
    Test(test::TestArgs),
    /// ブックの中身をテキストで見る・比べる（Excel は不要）
    Book {
        #[command(subcommand)]
        action: book::BookCommand,
    },
    /// シートの追加・複製・削除・名前の変更を、ブックと objects/ のファイルに一緒に行う
    Sheet {
        #[command(subcommand)]
        action: sheet::SheetCommand,
    },
    /// LSP サーバとして起動する（VS Code 拡張から呼ばれる）
    Lsp {
        #[arg(long)]
        stdio: bool,
    },
    /// Excel で開いているブックの VBA プロジェクトへ、モジュールとユーザーフォームを 1 回だけ反映する（ライブ同期。Windows のみ）
    Sync {
        #[command(flatten)]
        build: BuildArgs,
        /// VBE 上で編集された RAS 管理モジュール（ドリフト）と、同名の手書きモジュール・VBE で作ったフォームを上書きする
        #[arg(long)]
        overwrite: bool,
        /// Excel → ローカル方向に同期する（未実装）
        #[arg(long)]
        pull: bool,
    },
    /// VBA プロジェクトへのアクセス許可と RASBridge のインストール（未実装）
    Setup,
    /// COM タイプライブラリから型定義（.d.ras）を生成する（未実装）
    Typegen { typelib: String },
}

#[derive(Subcommand)]
enum FormCommand {
    /// 新しいユーザーフォームの雛形（forms/<Name>.form.json とコードビハインド forms/<Name>.ras）を作る
    New {
        /// フォーム名（VBA の識別子）
        name: String,
    },
    /// forms/ のユーザーフォームを検査する（ブックへの書き込みと同じ生成まで行う）
    Check,
    /// ブックにあるユーザーフォームを forms/ に取り出す（デザイン・コード・Image の画像）
    Pull {
        /// 取り出すフォーム名（省略時はブックのすべてのフォーム）
        names: Vec<String>,
        /// 対象ブック（省略時は ras.toml の project.workbook）
        #[arg(long, value_name = "FILE")]
        book: Option<PathBuf>,
        /// forms/ に既にあるファイルを置き換える
        #[arg(long)]
        force: bool,
    },
}

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
enum CheckTarget {
    /// Excel で動かす VBA（既定）
    Excel,
    /// ブックの代役
    StandIn,
}

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
enum Target {
    Vba,
    Js,
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
    JSON_MESSAGES.store(
        cli.message_format == MessageFormat::Json,
        std::sync::atomic::Ordering::Relaxed,
    );
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
    let p = explicit
        .map(Path::to_path_buf)
        .or_else(|| cfg.project.workbook.as_ref().map(PathBuf::from))?;
    Some(if p.is_absolute() { p } else { root.join(p) })
}

/// Loads and analyses the project. Sheet code names come from the workbook when it can be read.
/// `book` is a path already resolved with [`workbook_path`]; without it the project's workbook is used.
fn analyze(root: &Path, args: &BuildArgs, book: Option<&Path>) -> Result<Program, String> {
    analyze_project(root, args, book, false)
}

/// The project with its tests (spec §14.12.2): what `ras check` and `ras test` look at.
fn analyze_with_tests(root: &Path) -> Result<Program, String> {
    analyze_project(root, &BuildArgs::default(), None, true)
}

fn analyze_project(
    root: &Path,
    args: &BuildArgs,
    book: Option<&Path>,
    tests: bool,
) -> Result<Program, String> {
    let (mut cfg, sources) = if tests {
        ras_compiler::load_project_with_tests(root)?
    } else {
        ras_compiler::load_project(root)?
    };
    cfg.opt_override = args.opt.or(if args.release {
        Some(cfg.build.opt_level.max(2))
    } else {
        None
    });
    let book = match book {
        Some(b) => Some(b.to_path_buf()),
        None => workbook_path(root, &cfg, None),
    };
    if let Some(book) = book {
        if book.is_file() {
            if let Err(e) = ras_sync::read_workbook(&mut cfg, root, &book) {
                eprintln!(
                    "warning: {} のシート名を読み取れませんでした: {e}",
                    book.display()
                );
            }
        }
    }
    Ok(ras_compiler::analyze(sources, cfg))
}

/// Writes the mechanical corrections of `prog` back to the user's source files. Returns how many were applied.
fn apply_fixes(prog: &Program) -> Result<usize, String> {
    let mut count = 0;
    for md in prog.modules.iter().filter(|m| m.kind == ModuleKind::User) {
        let fixes: Vec<&ras_compiler::Fix> =
            prog.fixes.iter().filter(|f| f.module == md.id).collect();
        if fixes.is_empty() {
            continue;
        }
        let fixed = ras_compiler::apply_fixes(&md.text, &fixes);
        if fixed == *md.text {
            continue;
        }
        // The same place can be reported more than once (a generic function checked per instantiation).
        let mut places: Vec<u32> = fixes.iter().map(|f| u32::from(f.span.start())).collect();
        places.sort_unstable();
        places.dedup();
        count += places.len();
        // `load_project` drops a byte-order mark: keep the one the file has.
        let had_bom = std::fs::read(&md.path).is_ok_and(|b| b.starts_with(&[0xEF, 0xBB, 0xBF]));
        let out = if had_bom {
            format!("\u{feff}{fixed}")
        } else {
            fixed
        };
        std::fs::write(&md.path, out).map_err(|e| format!("{}: {e}", md.path.display()))?;
    }
    Ok(count)
}

/// Prints diagnostics (with the offending source line) and returns the number of errors.
fn report(prog: &Program, diags: &[Diagnostic], show_hints: bool) -> usize {
    let mut items: Vec<&Diagnostic> = diags
        .iter()
        .filter(|d| show_hints || d.severity <= Severity::Warning)
        .filter(|d| {
            d.module
                .is_none_or(|m| prog.modules[m].kind == ModuleKind::User)
                || d.is_error()
        })
        .collect();
    items.sort_by_key(|d| (d.module, u32::from(d.span.start())));
    if JSON_MESSAGES.load(std::sync::atomic::Ordering::Relaxed) {
        for d in &items {
            println!("{}", diagnostic_json(prog, d));
        }
        return items.iter().filter(|d| d.is_error()).count();
    }
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
            let prefix: String = md.text[ls..(start as usize).min(le)]
                .chars()
                .map(|c| if c == '\t' { '\t' } else { ' ' })
                .collect();
            let width = md.text[(start as usize).min(le)
                ..(u32::from(d.span.end()) as usize).clamp(start as usize, le)]
                .chars()
                .count()
                .max(1);
            eprintln!(
                "   |\n{:>3}| {src}\n   | {prefix}{}",
                line + 1,
                "^".repeat(width)
            );
        }
        eprintln!();
    }
    let errors = items.iter().filter(|d| d.is_error()).count();
    static HINTED: std::sync::Once = std::sync::Once::new();
    if errors > 0 {
        HINTED.call_once(|| {
            eprintln!(
                "（エラーの説明と直し方: ras explain <コード>。例: ras explain {}）",
                items
                    .iter()
                    .find(|d| d.is_error())
                    .map(|d| d.code)
                    .unwrap_or("E0301")
            )
        });
    }
    errors
}

/// A diagnostic as `--message-format json` writes it: lines and columns are 1-based, columns
/// in UTF-16 code units (as editors count them); `fixes` are the edits `ras check --fix` makes.
fn diagnostic_json(prog: &Program, d: &Diagnostic) -> serde_json::Value {
    use serde_json::json;
    let severity = match d.severity {
        Severity::Error => "error",
        Severity::Warning => "warning",
        Severity::Info => "info",
        Severity::Hint => "hint",
    };
    let mut out =
        json!({ "type": "diagnostic", "code": d.code, "severity": severity, "message": d.message });
    let range = |m: usize, span: ras_syntax::ast::Span| {
        let md = &prog.modules[m];
        let (l1, c1) = md.line_col_utf16(u32::from(span.start()));
        let (l2, c2) = md.line_col_utf16(u32::from(span.end()));
        json!({ "start": { "line": l1 + 1, "column": c1 + 1 }, "end": { "line": l2 + 1, "column": c2 + 1 } })
    };
    if let Some(m) = d.module {
        out["file"] = json!(prog.modules[m].rel_path);
        out["range"] = range(m, d.span);
        // The corrections of this diagnostic: those of its code inside its range.
        let fixes: Vec<serde_json::Value> = prog
            .fixes
            .iter()
            .filter(|f| f.module == m && f.code == d.code && f.span.start() >= d.span.start() && f.span.end() <= d.span.end())
            .map(|f| json!({ "file": prog.modules[m].rel_path, "range": range(m, f.span), "newText": f.replacement }))
            .collect();
        if !fixes.is_empty() {
            out["fixes"] = json!(fixes);
        }
    }
    if explain::find(d.code).is_some() {
        out["explain"] = json!(format!("ras explain {}", d.code));
    }
    out
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
fn compile(
    root: &Path,
    args: &BuildArgs,
    book: Option<&Path>,
) -> Result<Option<(Program, BuildOutput)>, String> {
    compile_with(root, args, book, false)
}

/// `no_attributes`: for code that goes through `CodeModule.AddFromString`, which rejects `Attribute` lines (spec §16.1.3).
fn compile_with(
    root: &Path,
    args: &BuildArgs,
    book: Option<&Path>,
    no_attributes: bool,
) -> Result<Option<(Program, BuildOutput)>, String> {
    let prog = analyze(root, args, book)?;
    let errors = report(&prog, &prog.diagnostics, false);
    if errors > 0 {
        eprintln!("{errors} 件のエラーがあります");
        return Ok(None);
    }
    let out = ras_compiler::build(
        &prog,
        &BuildOptions {
            release: args.release,
            erl: args.erl,
            no_attributes,
        },
    );
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
    for f in out.files.iter().filter(|f| f.kind != VModuleKind::Form) {
        // Module files are stored in the target code page, as the VBE expects (spec §13.7).
        let path = dir.join(&f.file_name);
        std::fs::write(&path, encode(&f.full_text, prog.config.target.codepage))
            .map_err(|e| format!("{}: {e}", path.display()))?;
        keep.insert(f.file_name.clone());
        let map = format!("{}.rasmap", f.name);
        std::fs::write(dir.join(&map), f.rasmap()).map_err(|e| e.to_string())?;
        keep.insert(map);
    }
    if let Ok(entries) = std::fs::read_dir(&dir) {
        for e in entries.flatten() {
            let name = e.file_name().to_string_lossy().to_string();
            let generated = [".bas", ".cls", ".doccls", ".rasmap"]
                .iter()
                .any(|ext| name.ends_with(ext));
            if generated && !keep.contains(&name) {
                let _ = std::fs::remove_file(e.path());
            }
        }
    }
    Ok(dir)
}

fn cmd_build(root: &Path, args: &BuildArgs) -> Result<ExitCode, String> {
    let Some((prog, out)) = compile(root, args, None)? else {
        return Ok(ExitCode::from(1));
    };
    if load_forms(root, prog.config.target.codepage, &out)?.is_none() {
        return Ok(ExitCode::from(1));
    }
    let dir = write_build(root, &prog, &out)?;
    let modules = out
        .files
        .iter()
        .filter(|f| f.kind != VModuleKind::Form)
        .count();
    println!(
        "{modules} 個のモジュールを {} に出力しました",
        dir.display()
    );
    Ok(ExitCode::SUCCESS)
}

/// `ras build --target js`: the debug build (spec §14.10). Runs the program as written,
/// so the optimisation options of the VBA build do not apply.
fn cmd_build_js(root: &Path, args: &BuildArgs) -> Result<ExitCode, String> {
    let Some(out) = build_debug(root, args)? else {
        return Ok(ExitCode::from(1));
    };
    let dir = root.join("build").join("debug");
    println!(
        "デバッグ用の JavaScript を {} に出力しました（実行できる関数: {}）",
        dir.display(),
        if out.entries.is_empty() {
            "なし".to_string()
        } else {
            out.entries.join(", ")
        }
    );
    Ok(ExitCode::SUCCESS)
}

/// Writes the debug build into build/debug/; `None` (after printing the problems) when the
/// program has errors.
fn build_debug(root: &Path, args: &BuildArgs) -> Result<Option<ras_compiler::JsOutput>, String> {
    let prog = analyze(root, args, None)?;
    let errors = report(&prog, &prog.diagnostics, false);
    if errors > 0 {
        eprintln!("{errors} 件のエラーがあります");
        return Ok(None);
    }
    let out = ras_compiler::build_js(&prog, "../../");
    let errors = report(&prog, &out.diagnostics, false);
    if errors > 0 {
        eprintln!("{errors} 件のエラーがあります");
        return Ok(None);
    }
    run::write_js(&root.join("build").join("debug"), &out.files)?;
    Ok(Some(out))
}

/// `ras run`: builds the debug build and runs an entry point of it under Node, on the
/// stand-in for Excel.
fn cmd_run(root: &Path, args: &run::RunArgs) -> Result<ExitCode, String> {
    let Some(out) = build_debug(root, &BuildArgs::default())? else {
        return Ok(ExitCode::from(2));
    };
    run::run(root, &out, args)
}

/// Loads, validates and generates the forms of `forms/` with the code compiled from their
/// `.ras` code-behinds; `None` (after printing the problems) when there are errors.
fn load_forms(
    root: &Path,
    codepage: u16,
    out: &BuildOutput,
) -> Result<Option<Vec<ras_ovba::UserForm>>, String> {
    let (_, built) = forms::build(root, codepage, &forms::compiled_code(&out.files))?;
    let errors = forms::report(&built);
    if errors > 0 {
        eprintln!("ユーザーフォームに {errors} 件のエラーがあります");
        return Ok(None);
    }
    Ok(Some(built.forms))
}

fn to_ovba_kind(k: &VModuleKind) -> ras_ovba::ModuleKind {
    match k {
        VModuleKind::Standard => ras_ovba::ModuleKind::Standard,
        VModuleKind::Class => ras_ovba::ModuleKind::Class,
        VModuleKind::Document => ras_ovba::ModuleKind::Document,
        VModuleKind::Form => ras_ovba::ModuleKind::Form,
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
        return Err(format!(
            "{} が見つかりません（`ras init --workbook` で空のブックを作成できます）",
            book.display()
        ));
    }
    // The hand-written VBA of the workbook is taken over into the project first (spec §14.7.1).
    let prepared = if cfg.sync.pull_unmanaged && p.output.is_none() && !is_open_in_excel(&book) {
        take_over_offline(root, &book, cfg.target.codepage)?
    } else {
        ras_sync::Prepared::default()
    };
    report_take_over(&prepared);
    let Some((prog, out)) = compile(root, args, Some(&book))? else {
        return Ok(ExitCode::from(1));
    };
    write_build(root, &prog, &out)?;
    let Some(user_forms) = load_forms(root, prog.config.target.codepage, &out)? else {
        return Ok(ExitCode::from(1));
    };
    let opts = ras_sync::options(root, &prepared, p.overwrite);
    let pending = |name: &str| {
        opts.pending
            .iter()
            .any(|(c, g)| c.eq_ignore_ascii_case(name) || g.eq_ignore_ascii_case(name))
    };
    let adopted = |name: &str| opts.adopt.iter().any(|n| n.eq_ignore_ascii_case(name));
    // What a pending take-over generates is not written yet.
    let user_forms: Vec<ras_ovba::UserForm> = user_forms
        .into_iter()
        .filter(|f| !pending(&f.name))
        .collect();
    let files: Vec<&ras_compiler::OutputFile> =
        out.files.iter().filter(|f| !pending(&f.name)).collect();

    // Modules that RAS generated earlier but that no longer exist are removed; hand-edited ones are reported (spec §14.7).
    let existing = ras_ovba::unpack(&book).map_err(|e| format!("{}: {e}", book.display()))?;
    let generated: HashSet<&str> = files
        .iter()
        .filter(|f| f.kind != VModuleKind::Form)
        .map(|f| f.name.as_str())
        .collect();
    let mut remove = Vec::new();
    let mut drifted = Vec::new();
    if let Some(project) = &existing.project {
        for m in &project.modules {
            // Document modules are decided below, each as a whole.
            if m.kind == ras_ovba::ModuleKind::Document {
                continue;
            }
            let Some(is_drifted) = ras_compiler::emit::has_drifted(&m.source) else {
                // Taken over into a RAS module of another name (`mod_Main` → `ModMain`).
                if m.kind == ras_ovba::ModuleKind::Standard
                    && adopted(&m.name)
                    && !generated.iter().any(|g| g.eq_ignore_ascii_case(&m.name))
                {
                    remove.push(m.name.clone());
                }
                continue;
            };
            // (What was added to it in the VBE was taken over: written again from its file.)
            if is_drifted && !adopted(&m.name) {
                drifted.push(m.name.clone());
            }
            if !generated.contains(m.name.as_str())
                && matches!(
                    m.kind,
                    ras_ovba::ModuleKind::Standard | ras_ovba::ModuleKind::Class
                )
            {
                remove.push(m.name.clone());
            }
        }
    }
    // A module of the same name that RAS did not generate is hand-written code: replacing it silently would lose it.
    let mut hand_written: Vec<String> = existing
        .project
        .iter()
        .flat_map(|project| &project.modules)
        .filter(|m| {
            matches!(
                m.kind,
                ras_ovba::ModuleKind::Standard
                    | ras_ovba::ModuleKind::Class
                    | ras_ovba::ModuleKind::Form
            )
        })
        .filter(|m| {
            ras_compiler::emit::has_drifted(&m.source).is_none()
                && has_code(&m.source)
                && !adopted(&m.name)
        })
        .filter(|m| {
            generated
                .iter()
                .any(|g| g.to_lowercase() == m.name.to_lowercase())
        })
        .map(|m| m.name.clone())
        .collect();
    let locked = || {
        format!(
            "{} は Excel で開かれているため、オフライン書き込みはできません。開いたまま反映するには `ras sync` を使ってください（Windows のみ）。または、ブックを閉じてから再実行してください{}",
            book.display(),
            stale_owner_hint(&book)
        )
    };
    // A workbook open in Excel is not written to, whatever else is wrong (spec §14.9.3).
    if is_open_in_excel(&book) {
        return Err(locked());
    }
    // The modules of the workbook and of the sheets are written whole from their files
    // (spec §12.2, §14.6).
    let mut document_code = Vec::new();
    let mut document_names = vec![existing.code_names.workbook.clone()];
    document_names.extend(
        existing
            .code_names
            .sheets
            .iter()
            .map(|s| s.code_name.clone()),
    );
    for f in files.iter().filter(|f| f.kind == VModuleKind::Document) {
        let name = if f.name.eq_ignore_ascii_case("ThisWorkbook") {
            existing.code_names.workbook.as_str()
        } else {
            f.name.as_str()
        };
        if !document_names
            .iter()
            .any(|n| n.to_lowercase() == name.to_lowercase())
        {
            return Err(format!("{}: ブックに `{}` というコード名のシートがありません（objects/{}.ras）。シートのコード名を変えたときは、ファイル名も変えてください", book.display(), f.name, f.name));
        }
    }
    for name in document_names.iter().filter(|n| !n.is_empty()) {
        let generated_name = if name.to_lowercase() == existing.code_names.workbook.to_lowercase() {
            "ThisWorkbook"
        } else {
            name.as_str()
        };
        if pending(name) || pending(generated_name) {
            continue;
        }
        let current = existing
            .project
            .iter()
            .flat_map(|project| &project.modules)
            .find(|m| {
                m.kind == ras_ovba::ModuleKind::Document
                    && m.name.to_lowercase() == name.to_lowercase()
            })
            .map(|m| m.source.as_str())
            .unwrap_or("");
        let code = files
            .iter()
            .find(|f| {
                f.kind == VModuleKind::Document
                    && f.name.to_lowercase() == generated_name.to_lowercase()
            })
            .map(|f| f.code.as_str());
        let own = opts
            .written
            .iter()
            .any(|n| n.to_lowercase() == generated_name.to_lowercase());
        match ras_sync::plan::document_action(current, code, adopted(name), p.overwrite, own) {
            ras_sync::plan::DocumentAction::Write(code) => document_code.push(ras_ovba::DocumentCode { code_name: name.clone(), code }),
            ras_sync::plan::DocumentAction::Drifted => drifted.push(name.clone()),
            ras_sync::plan::DocumentAction::HandWritten => hand_written.push(name.clone()),
            ras_sync::plan::DocumentAction::Foreign => eprintln!("warning: シート `{name}` のモジュールには RAS が生成したコードがありますが、対応するファイル（objects/{name}.ras）がありません（シートの複製でコードごと複製されたものです）。そのまま残しています。`ras sync`（Excel で開いた状態）でコピー元のファイルを複製できます"),
            ras_sync::plan::DocumentAction::Unchanged | ras_sync::plan::DocumentAction::Keep => {}
        }
    }
    if !hand_written.is_empty() && !p.overwrite {
        return Err(format!(
            "ブックには RAS が生成したものではない同名のモジュールがあります: {}\n  このまま書き込むと手書きのコードが失われます。`ras unpack` で vba/ に取り出してモジュール名を変えるか、RAS 側のファイル名を変えてください。置き換えてよい場合は --overwrite を付けてください",
            hand_written.join(", ")
        ));
    }
    if !drifted.is_empty() && !p.overwrite {
        return Err(format!(
            "次の RAS 管理モジュールが VBE 上で編集されています（ドリフト）: {}\n  RAS の出力で上書きするには --overwrite を付けてください。VBE 側の内容は `ras unpack` で build/drift/ に保存できます",
            drifted.join(", ")
        ));
    }
    // Forms: a name used by a RAS module, a module of another kind, or a form made in the VBE is not replaced silently.
    let mut hand_made_forms = Vec::new();
    for form in &user_forms {
        if generated.iter().any(|g| g.eq_ignore_ascii_case(&form.name)) {
            return Err(format!("ユーザーフォーム `{}` と同じ名前の RAS モジュールがあります。どちらかの名前を変えてください", form.name));
        }
        let Some(m) = existing
            .project
            .iter()
            .flat_map(|project| &project.modules)
            .find(|m| m.name.eq_ignore_ascii_case(&form.name))
        else {
            continue;
        };
        if m.kind != ras_ovba::ModuleKind::Form {
            return Err(format!("ブックには `{}` という名前のフォームではないモジュールがあります。フォーム名を変えてください", m.name));
        }
        if !ras_ovba::is_generated_form(&m.source) && !adopted(&m.name) {
            hand_made_forms.push(m.name.clone());
        }
    }
    if !hand_made_forms.is_empty() && !p.overwrite {
        return Err(format!(
            "ブックには forms/ から生成したものではない同名のユーザーフォームがあります: {}\n  このまま書き込むと VBE で作ったフォームとそのコードが失われます。フォーム名を変えるか、`ras form pull` で forms/ に取り出してください。置き換えてよい場合は --overwrite を付けてください",
            hand_made_forms.join(", ")
        ));
    }
    let mut modules = Vec::new();
    for f in files {
        match f.kind {
            // Written with the form's designer (`user_forms`); the document modules: above.
            VModuleKind::Form | VModuleKind::Document => {}
            _ => modules.push(ras_ovba::VbaModule {
                name: f.name.clone(),
                kind: to_ovba_kind(&f.kind),
                source: f.stream_text(),
            }),
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
        rename_modules: Vec::new(),
        prune_documents: false,
        forms: user_forms,
        backup_dir: Some(root.join("build").join("backup")),
        backup_generations: prog.config.pack.backup_generations,
    };
    match ras_ovba::pack(&req) {
        Ok(rep) => {
            ras_form::import::finish(&prepared)?;
            // What was written, to tell later what was added in the VBE (spec §14.7.3).
            if p.output.is_none() {
                let not_written: Vec<String> = out.files.iter().filter(|f| pending(&f.name)).map(|f| f.name.clone()).collect();
                ras_sync::save_written(root, &out.files, &not_written)?;
            }
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
        Err(ras_ovba::OvbaError::Locked) => Err(locked()),
        Err(ras_ovba::OvbaError::DuplicateProcedure { module, procedure }) => Err(format!(
            "{}: ドキュメントモジュール `{module}` には手書きの `{procedure}` があり、RAS が生成するプロシージャと名前が重なります（このまま書き込むと、ブックを開いたときに VBE が「名前が適切ではありません」のコンパイル エラーを出します）\n  手書き側の処理を RAS の `@on` ハンドラへ移すか、手書きのプロシージャを削除してください",
            book.display()
        )),
        Err(e) => Err(format!("{}: {e}", book.display())),
    }
}

/// `ras sync`: writes the generated VBA into the workbook that is open in Excel (spec §14.1–14.8).
///
/// `Ok(None)` when no Excel has the workbook open, or on a platform without Excel: the caller decides what to do then.
fn live_sync(root: &Path, args: &BuildArgs, overwrite: bool) -> Result<Option<ExitCode>, String> {
    let (cfg, _) = ras_compiler::load_project(root)?;
    let Some(book) = workbook_path(root, &cfg, None) else {
        return Err(
            "同期先のブックが指定されていません。ras.toml の project.workbook で指定してください"
                .into(),
        );
    };
    let mut session = ras_sync::Session::new(&book);
    match session.connect() {
        Ok(()) => {}
        Err(ras_sync::SyncError::NotOpen(_) | ras_sync::SyncError::Unsupported) => return Ok(None),
        Err(e) => return Err(e.to_string()),
    }
    // The hand-written VBA of the workbook is taken over into the project first (spec §14.7.1),
    // and the sheets without a file get one (spec §14.7.2).
    let prepared = if cfg.sync.pull_unmanaged || cfg.sync.sheet_files {
        session
            .take_over(root, cfg.target.codepage, &cfg.sync)
            .map_err(|e| e.to_string())?
    } else {
        ras_sync::Prepared::default()
    };
    report_take_over(&prepared);
    let Some((prog, out)) = compile_with(root, args, Some(&book), true)? else {
        return Ok(Some(ExitCode::from(1)));
    };
    write_build(root, &prog, &out)?;
    let codepage = prog.config.target.codepage;
    let Some(user_forms) = load_forms(root, codepage, &out)? else {
        return Ok(Some(ExitCode::from(1)));
    };
    let rep = session
        .apply(
            &out.files,
            ras_sync::Forms {
                forms: &user_forms,
                codepage,
            },
            &ras_sync::options(root, &prepared, overwrite),
        )
        .map_err(|e| e.to_string())?;
    ras_form::import::finish(&prepared)?;
    // What was written, to tell later what was added in the VBE (spec §14.7.3).
    let mut not_written = rep.drifted.clone();
    not_written.extend(rep.pending.iter().cloned());
    not_written.extend(prepared.pending.iter().map(|p| p.generated.clone()));
    ras_sync::save_written(root, &out.files, &not_written)?;
    if !rep.foreign.is_empty() {
        eprintln!("warning: 次のシートのモジュールには RAS が生成したコードがありますが、対応するファイルがありません（シートの複製でコードごと複製されたものです）。そのまま残しています: {}", rep.foreign.join(", "));
    }
    println!("{} に同期しました（{}）", book.display(), rep.summary());
    if rep.drifted.is_empty() {
        return Ok(Some(ExitCode::SUCCESS));
    }
    eprintln!(
        "warning: 次の RAS 管理モジュールは VBE 上で編集されているため、上書きしていません（ドリフト）: {}\n  RAS の出力で上書きするには `ras sync --overwrite` を実行してください",
        rep.drifted.join(", ")
    );
    Ok(Some(ExitCode::from(1)))
}

fn cmd_sync(root: &Path, args: &BuildArgs, overwrite: bool) -> Result<ExitCode, String> {
    match live_sync(root, args, overwrite)? {
        Some(code) => Ok(code),
        None if cfg!(windows) => Err("ブックが Excel で開かれていません。Excel で開いてから再実行するか、`ras pack` でファイルへ直接書き込んでください".into()),
        None => Err(format!("{}。`ras pack` でファイルへ直接書き込めます", ras_sync::SyncError::Unsupported)),
    }
}

/// Takes over the hand-written VBA of the workbook file `book` into the project (spec §14.7.1),
/// as the live sync does with the open workbook.
fn take_over_offline(
    root: &Path,
    book: &Path,
    codepage: u16,
) -> Result<ras_sync::Prepared, String> {
    let at = |e: ras_ovba::OvbaError| format!("{}: {e}", book.display());
    let unpacked = ras_ovba::unpack(book).map_err(at)?;
    let Some(project) = unpacked.project else {
        return Ok(ras_sync::Prepared::default());
    };
    let modules: Vec<ras_form::import::BookModule> = project
        .modules
        .iter()
        .map(|m| ras_form::import::BookModule {
            name: m.name.clone(),
            kind: m.kind,
            code: m.source.clone(),
        })
        .collect();
    let mut forms: Option<Vec<ras_ovba::UserForm>> = None;
    let mut form = |name: &str| -> Result<ras_ovba::UserForm, String> {
        if forms.is_none() {
            forms = Some(ras_ovba::unpack_user_forms(book).map_err(at)?);
        }
        forms
            .iter()
            .flatten()
            .find(|f| f.name.eq_ignore_ascii_case(name))
            .cloned()
            .ok_or_else(|| {
                format!(
                    "{} にフォーム `{name}` のデザイナがありません",
                    book.display()
                )
            })
    };
    let sheet_controls = ras_sync::sheet_controls(book).unwrap_or_default();
    let mut written = ras_sync::Written::new(root);
    let mut written = |name: &str| written.code(name);
    let mut taken = ras_form::import::Book {
        modules: &modules,
        workbook_code_name: &unpacked.code_names.workbook,
        sheet_controls: &sheet_controls,
        form: &mut form,
        written: &mut written,
    };
    let mut prepared = ras_form::import::prepare(root, &mut taken, codepage)?;
    // Code added in the VBE that does not compile as RAS goes back out of the file (spec §14.7.3).
    ras_sync::settle_appended(root, &mut prepared)?;
    Ok(prepared)
}

/// Prints what was taken over from the workbook and what waits for the user.
fn report_take_over(prepared: &ras_sync::Prepared) {
    for line in prepared.messages() {
        eprintln!("{line}");
    }
}

/// Whether Excel has the workbook open (its owner file `~$Book.xlsm` exists).
fn is_open_in_excel(book: &Path) -> bool {
    book.file_name()
        .and_then(|n| n.to_str())
        .is_some_and(|n| book.with_file_name(format!("~${n}")).exists())
}

/// Whether a module holds anything besides its `Attribute` header, `Option` statements and blank lines.
fn has_code(source: &str) -> bool {
    source.lines().map(str::trim).any(|l| {
        !l.is_empty()
            && !l.starts_with("Attribute ")
            && !l.starts_with("VERSION ")
            && !l.to_ascii_lowercase().starts_with("option ")
    })
}

/// Excel leaves its owner file (`~$Book.xlsm`) behind when it crashes; the workbook then looks open for ever.
fn stale_owner_hint(book: &Path) -> String {
    let owner = book
        .file_name()
        .and_then(|n| n.to_str())
        .map(|n| book.with_file_name(format!("~${n}")));
    match owner {
        Some(owner) if owner.exists() => format!("\n  Excel で開いていないのにこのエラーになる場合は、Excel の異常終了で残った {} を削除してください", owner.display()),
        _ => String::new(),
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
        let code: String = m
            .source
            .replace("\r\n", "\n")
            .lines()
            .skip_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION "))
            .map(|l| format!("{l}\n"))
            .collect();
        match ras_compiler::emit::has_drifted(&m.source) {
            Some(false) => managed += 1,
            Some(true) => {
                // RAS-managed module edited in the VBE: keep the VBE side for the user to compare.
                std::fs::create_dir_all(&drift_dir).map_err(|e| e.to_string())?;
                std::fs::write(drift_dir.join(format!("{}.vba", m.name)), &code)
                    .map_err(|e| e.to_string())?;
                eprintln!("warning: RAS 管理モジュール `{}` は VBE 上で編集されています。VBE 側の内容を build/drift/{}.vba に保存しました", m.name, m.name);
                drifted += 1;
            }
            None if m.kind == ras_ovba::ModuleKind::Form
                && ras_ovba::is_generated_form(&m.source) =>
            {
                managed += 1
            }
            None => {
                let ext = match m.kind {
                    ras_ovba::ModuleKind::Standard => "bas",
                    ras_ovba::ModuleKind::Form => "frm",
                    _ => "cls",
                };
                // Document modules: only the hand-written part outside the RAS-managed region.
                let text = if m.kind == ras_ovba::ModuleKind::Document {
                    ras_ovba::merge_managed_region(&m.source, "").replace("\r\n", "\n")
                } else {
                    m.source.replace("\r\n", "\n")
                };
                if m.kind == ras_ovba::ModuleKind::Document
                    && text
                        .lines()
                        .all(|l| l.starts_with("Attribute ") || l.trim().is_empty())
                {
                    continue;
                }
                std::fs::create_dir_all(&vba_dir).map_err(|e| e.to_string())?;
                std::fs::write(vba_dir.join(format!("{}.{ext}", m.name)), text)
                    .map_err(|e| e.to_string())?;
                written += 1;
            }
        }
    }
    println!("非管理モジュール {written} 個を {} に取り出しました（RAS 管理 {managed} 個、ドリフト {drifted} 個）", vba_dir.display());
    Ok(ExitCode::SUCCESS)
}

const AGENTS_TEMPLATE: &str = include_str!("../templates/AGENTS.md");
const CHEATSHEET: &str = include_str!("../../../docs/cheatsheet.md");

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
    let name = if name.is_empty() {
        "RASProject".to_string()
    } else {
        name
    };
    let wb_line = match workbook {
        Some(w) => format!("workbook = \"{w}\""),
        None => "# workbook = \"Book.xlsm\"    # `ras pack` / 同期の対象ブック".to_string(),
    };
    let toml = format!(
        "[project]\nname = \"{name}\"\n{wb_line}\n\n[target]\nhost = \"excel\"\nvba6 = true             # VBA6（32bit Office 2007 以前）用の Declare も出力\ncodepage = 932\n\n[build]\nopt-level = 1           # 0: 展開なし / 1: ループ内自動展開 / 2: 積極的\nerl = true              # 行番号付与（エラー位置を RAS のソース行で表示）\n\n[pack]\nbackup-generations = 5\n\n[doEvents]\nenabled = true\ninterval-ms = 100\nloop-threshold = 10000\n"
    );
    std::fs::write(&toml_path, toml).map_err(|e| e.to_string())?;
    let main = "/**\n * サンプルのマクロ。Excel の「マクロ」ダイアログから `hello` を実行できる。\n */\nexport function hello() {\n    const names = [\"RAS\", \"VBA\"];\n    const greeting = names.map(n => `Hello, ${n}!`).join(\"\\n\");\n    msgBox(greeting);\n}\n";
    let main_path = dir.join("src").join("main.ras");
    if !main_path.exists() {
        std::fs::write(&main_path, main).map_err(|e| e.to_string())?;
    }
    // What an AI agent working in the project reads first (the work flow, the rules, what cannot be used).
    let agents = dir.join("AGENTS.md");
    if !agents.exists() {
        std::fs::write(&agents, AGENTS_TEMPLATE).map_err(|e| e.to_string())?;
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
            let kind = if w.to_lowercase().ends_with(".xlam") {
                ras_ovba::BookKind::Xlam
            } else {
                ras_ovba::BookKind::Xlsm
            };
            std::fs::write(&path, ras_ovba::create_workbook(kind, &["Sheet1"]))
                .map_err(|e| e.to_string())?;
        }
    }
    println!(
        "{} を作成しました。`ras build` で VBA を生成できます",
        toml_path.display()
    );
    Ok(ExitCode::SUCCESS)
}

fn cmd_fmt(root: Option<&Path>, files: &[PathBuf], check: bool) -> Result<ExitCode, String> {
    let mut targets: Vec<PathBuf> = files.to_vec();
    if targets.is_empty() {
        let root = find_root(root)?;
        let (_, sources) = ras_compiler::load_project(&root)?;
        targets = sources
            .into_iter()
            .filter(|s| s.kind == ModuleKind::User)
            .map(|s| s.path)
            .collect();
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
    Ok(if failed > 0 || (check && changed > 0) {
        ExitCode::from(1)
    } else {
        ExitCode::SUCCESS
    })
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
    for sub in ["src", "types", "forms", "ras.toml"] {
        let p = root.join(sub);
        if p.exists() {
            watcher
                .watch(&p, RecursiveMode::Recursive)
                .map_err(|e| e.to_string())?;
        }
    }
    let rebuild = || {
        let (cfg, _) = match ras_compiler::load_project(root) {
            Ok(x) => x,
            Err(e) => return eprintln!("error: {e}"),
        };
        let has_book = workbook_path(root, &cfg, None).is_some_and(|b| b.is_file());
        let result = if has_book && cfg.sync.enabled {
            // Open in Excel: live sync. Closed: the file is updated directly (spec §14.9).
            match live_sync(root, args, false) {
                Ok(Some(code)) => Ok(code),
                Ok(None) => cmd_pack(
                    root,
                    args,
                    &PackArgs {
                        book: None,
                        convert: false,
                        output: None,
                        overwrite: false,
                    },
                ),
                Err(e) => Err(e),
            }
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
        while let Ok(ev) =
            rx.recv_timeout(deadline.saturating_duration_since(std::time::Instant::now()))
        {
            relevant |= is_source_event(&ev);
        }
        if relevant {
            rebuild();
        }
    }
    Ok(ExitCode::SUCCESS)
}

fn is_source_event(ev: &notify::Event) -> bool {
    !matches!(ev.kind, notify::EventKind::Access(_))
        && ev.paths.iter().any(|p| {
            p.extension().is_some_and(|e| {
                [
                    "ras", "toml", "json", "bmp", "jpg", "jpeg", "png", "gif", "webp", "tif",
                    "tiff", "ico",
                ]
                .iter()
                .any(|known| e.eq_ignore_ascii_case(known))
            })
        })
}

fn unimplemented_cmd(what: &str, detail: &str) -> Result<ExitCode, String> {
    Err(format!("`{what}` はこの版では未実装です。{detail}"))
}

fn run(cli: Cli) -> Result<ExitCode, String> {
    let project = cli.project.as_deref();
    match &cli.command {
        Command::Init { dir, workbook } => cmd_init(dir, workbook.as_deref()),
        Command::Check { fix, target } => {
            let root = find_root(project)?;
            // The tests are checked too (spec §14.12.1); they are no part of what is built.
            let mut prog = analyze_with_tests(&root)?;
            if *fix {
                let fixed = apply_fixes(&prog)?;
                if fixed > 0 {
                    println!("{fixed} 箇所を修正しました");
                    prog = analyze_with_tests(&root)?;
                }
            }
            let mut errors = report(&prog, &prog.diagnostics, true);
            if *target == Some(CheckTarget::StandIn) || (target.is_none() && prog.config.lint.stand_in) {
                report(&prog, &ras_compiler::standin::check(&prog), true);
            }
            // Code generation can report further problems (size limits, unsupported constructs).
            let mut compiled = HashMap::new();
            if errors == 0 {
                if prog.modules.iter().any(|m| m.is_test()) {
                    prog = analyze(&root, &BuildArgs::default(), None)?;
                }
                let out = ras_compiler::build(&prog, &BuildOptions::default());
                errors += report(&prog, &out.diagnostics, true);
                compiled = forms::compiled_code(&out.files);
            }
            errors += forms::report(&forms::build(&root, prog.config.target.codepage, &compiled)?.1);
            if errors > 0 {
                eprintln!("{errors} 件のエラーがあります");
                return Ok(ExitCode::from(1));
            }
            println!("問題は見つかりませんでした");
            Ok(ExitCode::SUCCESS)
        }
        Command::Build { build, target: Target::Vba } => cmd_build(&find_root(project)?, build),
        Command::Build { build, target: Target::Js } => cmd_build_js(&find_root(project)?, build),
        Command::Watch(args) => cmd_watch(&find_root(project)?, args),
        Command::Pack { book, build, convert, output, overwrite, via_excel } => {
            if *via_excel {
                return unimplemented_cmd("ras pack --via-excel", "Excel を使わないオフライン書き込み（--via-excel なし）を使ってください。");
            }
            cmd_pack(&find_root(project)?, build, &PackArgs { book: book.as_deref(), convert: *convert, output: output.as_deref(), overwrite: *overwrite })
        }
        Command::Unpack { book } => cmd_unpack(&find_root(project)?, book.as_deref()),
        Command::Fmt { files, check } => cmd_fmt(project, files, *check),
        Command::Form { action: FormCommand::New { name } } => {
            let root = find_root(project)?;
            let (spec, code) = forms::new_form(&root, name)?;
            println!("{} と {} を作成しました", spec.display(), code.display());
            Ok(ExitCode::SUCCESS)
        }
        Command::Form { action: FormCommand::Pull { names, book, force } } => {
            let root = find_root(project)?;
            let (cfg, _) = ras_compiler::load_project(&root)?;
            let Some(book) = workbook_path(&root, &cfg, book.as_deref()) else {
                return Err("対象ブックが指定されていません。`ras form pull --book <book>` か ras.toml の project.workbook で指定してください".into());
            };
            forms::pull(&root, &book, names, *force)
        }
        Command::Form { action: FormCommand::Check } => {
            let root = find_root(project)?;
            // The code-behinds are RAS: the project is compiled to get the code of the forms.
            let Some((prog, out)) = compile(&root, &BuildArgs::default(), None)? else { return Ok(ExitCode::from(1)) };
            let (count, built) = forms::build(&root, prog.config.target.codepage, &forms::compiled_code(&out.files))?;
            let errors = forms::report(&built);
            if errors > 0 {
                eprintln!("{errors} 件のエラーがあります");
                return Ok(ExitCode::from(1));
            }
            println!("ユーザーフォーム {count} 個に問題は見つかりませんでした");
            Ok(ExitCode::SUCCESS)
        }
        Command::Cheatsheet => {
            print!("{CHEATSHEET}");
            Ok(ExitCode::SUCCESS)
        }
        Command::Explain { code: None } => {
            print!("{}", explain::list());
            Ok(ExitCode::SUCCESS)
        }
        Command::Explain { code: Some(code) } => match explain::find(code) {
            Some(entry) => {
                print!("{}", explain::text(entry));
                Ok(ExitCode::SUCCESS)
            }
            None => Err(format!("`{code}` という診断のコードはありません（`ras explain` でコードの一覧を表示します）")),
        },
        Command::Run(args) => cmd_run(&find_root(project)?, args),
        Command::Test(args) => test::cmd(&find_root(project)?, args),
        Command::Book { action: book::BookCommand::Dump { file, sheet, range, format, formats, max_rows, max_cols, all } } => {
            let file = match file {
                Some(f) => f.clone(),
                None => {
                    let root = find_root(project)?;
                    let (cfg, _) = ras_compiler::load_project(&root)?;
                    workbook_path(&root, &cfg, None).ok_or("対象ブックを指定してください（ras.toml の project.workbook が未設定です）")?
                }
            };
            book::dump(&file, sheet, range.as_deref(), *format, *formats, if *all { None } else { Some((*max_rows, *max_cols)) })
        }
        Command::Book { action: book::BookCommand::View { file, diff, output, stdout, open, all, title, drawings, sheet, editable, run, controls } } => {
            // The project, when there is one: its workbook, and build/view/ for the page.
            let root = find_root(project).ok();
            let file = match (file, &root) {
                (Some(f), _) => f.clone(),
                (None, Some(root)) => {
                    let (cfg, _) = ras_compiler::load_project(root)?;
                    workbook_path(root, &cfg, None).ok_or("対象ブックを指定してください（ras.toml の project.workbook が未設定です）")?
                }
                (None, None) => return Err("対象ブックを指定してください".into()),
            };
            if *stdout {
                print!("{}", book::view_page(&file, diff.as_deref(), *all, title.as_deref(), drawings.as_deref(), sheet.as_deref(), *editable, *run, controls.as_deref())?);
                return Ok(ExitCode::SUCCESS);
            }
            let dir = root.unwrap_or_else(|| PathBuf::from(".")).join("build").join("view");
            let page = book::write_view(&file, diff.as_deref(), *all, output.as_deref(), &dir, drawings.as_deref(), sheet.as_deref())?;
            println!("{} を作成しました（ブラウザーで開くと、シートをグリッドで見られます）", page.display());
            if *open {
                book::open_in_browser(&page)?;
            }
            Ok(ExitCode::SUCCESS)
        }
        Command::Book { action: book::BookCommand::Edit { file, sheet, set, clear, insert, delete, shift, op, ops_file, output, view, editable, all } } => {
            // The project, when there is one: its workbook, and build/backup/ for the copy of the file as it was.
            let root = find_root(project).ok();
            let config = match &root {
                Some(root) => Some(ras_compiler::load_project(root)?.0),
                None => None,
            };
            let file = match (file, &root, &config) {
                (Some(f), _, _) => f.clone(),
                (None, Some(root), Some(cfg)) => workbook_path(root, cfg, None).ok_or("対象ブックを指定してください（ras.toml の project.workbook が未設定です）")?,
                _ => return Err("対象ブックを指定してください".into()),
            };
            let backup_dir = root.as_ref().map(|r| r.join("build").join("backup"));
            let backup = match (&backup_dir, &config) {
                (Some(dir), Some(cfg)) => Some((dir.as_path(), cfg.pack.backup_generations)),
                _ => None,
            };
            let args = book::EditArgs { sheet: sheet.as_deref(), set, clear, insert, delete, shift: shift.as_deref(), op, ops_file: ops_file.as_deref() };
            let written = match book::edit(&file, &args, output.as_deref(), backup) {
                Ok(path) => path,
                Err(Ok(message)) => {
                    eprintln!("error: {message}");
                    return Ok(ExitCode::from(1));
                }
                Err(Err(message)) => return Err(message),
            };
            if *view {
                print!("{}", book::view_page(&written, None, *all, None, None, sheet.as_deref(), *editable, false, None)?);
            } else {
                println!("{} を書き換えました", written.display());
            }
            Ok(ExitCode::SUCCESS)
        }
        Command::Book { action: book::BookCommand::Diff { before, after, format, max_cells } } => book::diff_cmd(before, after, *format, *max_cells),
        Command::Sheet { action } => sheet::run(&find_root(project)?, action),
        Command::Lsp { .. } => {
            ras_lsp::run_stdio().map_err(|e| e.to_string())?;
            Ok(ExitCode::SUCCESS)
        }
        Command::Sync { pull: true, .. } => unimplemented_cmd("ras sync --pull", "ブックを閉じた状態で `ras unpack` を使うと、VBA ソースを vba/ に取り出せます。"),
        Command::Sync { build, overwrite, pull: false } => cmd_sync(&find_root(project)?, build, *overwrite),
        Command::Setup => unimplemented_cmd(
            "ras setup",
            "ライブ同期に RASBridge は不要です。Excel の［トラスト センター］→［マクロの設定］で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。",
        ),
        Command::Typegen { .. } => unimplemented_cmd("ras typegen", "同梱の Excel / VBA / Office の型定義は使えます。その他の COM ライブラリは types/*.d.ras に手書きしてください。"),
    }
}
