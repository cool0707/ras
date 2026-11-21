//! `ras book dump` / `ras book diff`: a workbook's contents as text, without Excel.

use std::path::{Path, PathBuf};
use std::process::ExitCode;

use ras_workbook::book::Book;
use ras_workbook::{diff, dump, RangeRef};

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
pub enum DumpFormat {
    Text,
    Json,
    Csv,
    Markdown,
}

#[derive(clap::ValueEnum, Clone, Copy, PartialEq, Eq)]
pub enum DiffFormat {
    Text,
    Json,
}

#[derive(clap::Subcommand)]
pub enum BookCommand {
    /// ブックの中身をテキストで出力する（シートの指定がなければ、シートの一覧・使用範囲・数式・名前・テーブル）
    Dump {
        /// 対象ブック（省略時は ras.toml の project.workbook）
        file: Option<PathBuf>,
        /// セルを出力するシート（シート名・コード名・1 から始まる番号。複数指定できる）
        #[arg(long, value_name = "SHEET")]
        sheet: Vec<String>,
        /// 出力する範囲（例: A1:F20、Sheet1!A1:F20）。省略時はシートの使用範囲
        #[arg(long, value_name = "A1:B2")]
        range: Option<String>,
        /// 出力形式
        #[arg(long, value_enum, default_value_t = DumpFormat::Text)]
        format: DumpFormat,
        /// 表示形式とセルの表示（Text）も出力する
        #[arg(long)]
        formats: bool,
        /// 出力する行数の上限
        #[arg(long, value_name = "N", default_value_t = 200)]
        max_rows: u32,
        /// 出力する列数の上限
        #[arg(long, value_name = "N", default_value_t = 50)]
        max_cols: u32,
        /// 行数・列数の上限を付けずに出力する
        #[arg(long)]
        all: bool,
    },
    /// 2 つのブックのセル（値・数式・表示形式）と構造（シート、行・列の挿入と削除の推定）の差分を出力する
    Diff {
        before: PathBuf,
        after: PathBuf,
        /// 出力形式
        #[arg(long, value_enum, default_value_t = DiffFormat::Text)]
        format: DiffFormat,
        /// text 形式で出力するセルの数の上限
        #[arg(long, value_name = "N", default_value_t = 200)]
        max_cells: usize,
    },
}

pub fn open(path: &Path) -> Result<Book, String> {
    if !path.is_file() {
        return Err(format!("{} が見つかりません", path.display()));
    }
    let ext = path.extension().map(|e| e.to_string_lossy().to_lowercase()).unwrap_or_default();
    if matches!(ext.as_str(), "xls" | "xlsb") {
        return Err(format!("{}: .{ext} は読み取れません（.xlsx / .xlsm / .xltx / .xltm / .xlam に保存し直してください）", path.display()));
    }
    Book::open(path)
}

/// Splits `Sheet1!A1:B2` / `'My sheet'!A1` into the sheet and the range.
fn split_range(text: &str) -> (Option<String>, &str) {
    match text.rsplit_once('!') {
        Some((sheet, range)) => {
            let sheet = sheet.trim();
            let sheet = sheet.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')).map(|s| s.replace("''", "'")).unwrap_or_else(|| sheet.to_string());
            (Some(sheet), range)
        }
        None => (None, text),
    }
}

pub fn dump(file: &Path, sheets: &[String], range: Option<&str>, format: DumpFormat, formats: bool, limits: Option<(u32, u32)>) -> Result<ExitCode, String> {
    let mut book = open(file)?;
    let name = file.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
    let opts = dump::Options {
        format: match format {
            DumpFormat::Text => dump::Format::Text,
            DumpFormat::Json => dump::Format::Json,
            DumpFormat::Csv => dump::Format::Csv,
            DumpFormat::Markdown => dump::Format::Markdown,
        },
        formats,
        max_rows: limits.map(|l| l.0),
        max_cols: limits.map(|l| l.1),
    };
    let mut sheets: Vec<String> = sheets.to_vec();
    let mut area = None;
    if let Some(text) = range {
        let (sheet, r) = split_range(text);
        if let Some(s) = sheet {
            sheets = vec![s];
        }
        area = Some(RangeRef::parse(&r.replace('$', "")).ok_or_else(|| format!("範囲 `{r}` を読み取れません（A1:F20 のように指定してください）"))?);
        if sheets.is_empty() {
            return Err("--range には --sheet でシートを指定するか、Sheet1!A1:F20 のようにシート名を付けてください".into());
        }
    }
    if sheets.is_empty() {
        if matches!(format, DumpFormat::Csv) {
            return Err("csv 形式は --sheet でシートを指定したときに使えます".into());
        }
        print!("{}", dump::outline(&mut book, &name, &opts)?);
        return Ok(ExitCode::SUCCESS);
    }
    let mut parts = Vec::new();
    for key in &sheets {
        let si = book.find_sheet(key).ok_or_else(|| {
            let all: Vec<String> = (0..book.sheet_count()).filter_map(|i| book.sheet_name(i).map(str::to_string)).collect();
            format!("シート `{key}` が見つかりません（シート: {}）", all.join(", "))
        })?;
        parts.push(dump::sheet(&mut book, &name, si, area, &opts)?);
    }
    match format {
        // Several sheets: one JSON document.
        DumpFormat::Json if parts.len() > 1 => {
            let list: Vec<serde_json::Value> = parts.iter().filter_map(|p| serde_json::from_str(p).ok()).collect();
            println!("{}", serde_json::to_string_pretty(&list).unwrap_or_default());
        }
        _ => print!("{}", parts.join("\n")),
    }
    Ok(ExitCode::SUCCESS)
}

/// The differences between two workbook files.
pub fn compare(before: &Path, after: &Path) -> Result<diff::Diff, String> {
    let mut a = open(before)?;
    let mut b = open(after)?;
    diff::diff(&mut a, &mut b)
}

/// Exit code 0 when the workbooks are the same, 1 when they differ.
pub fn diff_cmd(before: &Path, after: &Path, format: DiffFormat, max_cells: usize) -> Result<ExitCode, String> {
    let d = compare(before, after)?;
    match format {
        DiffFormat::Text => print!("{}", diff::text(&d, max_cells)),
        DiffFormat::Json => println!("{}", serde_json::to_string_pretty(&diff::to_json(&d)).unwrap_or_default()),
    }
    Ok(if d.is_empty() { ExitCode::SUCCESS } else { ExitCode::from(1) })
}
