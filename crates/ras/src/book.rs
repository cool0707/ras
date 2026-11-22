//! `ras book`: a workbook's contents as text (`dump`, `diff`), as a grid (`view`), and edits of
//! its cells written into the file (`edit`), without Excel.

use std::path::{Path, PathBuf};
use std::process::ExitCode;

use ras_workbook::book::Book;
use ras_workbook::{diff, dump, patch, view, RangeRef};

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
    /// ブックを Excel のようなグリッドで表示する HTML を作る（値・数式・主な書式・結合セル・フィルター・図形・画像・グラフ・コントロール。--diff で変更されたセルを強調）
    View {
        /// 対象ブック（省略時は ras.toml の project.workbook）
        file: Option<PathBuf>,
        /// このブックからの変更を強調して表示する（例: 実行前のブック）
        #[arg(long, value_name = "BEFORE")]
        diff: Option<PathBuf>,
        /// 出力する HTML ファイル（省略時は build/view/<ブック名>.html）
        #[arg(short, long, value_name = "FILE")]
        output: Option<PathBuf>,
        /// HTML をファイルに書かず、標準出力に出す
        #[arg(long)]
        stdout: bool,
        /// 作った HTML を既定のブラウザーで開く
        #[arg(long)]
        open: bool,
        /// 行数・列数の上限（500 行 × 100 列）を付けずに表示する
        #[arg(long)]
        all: bool,
        /// 表示するブックの名前（省略時はファイル名）
        #[arg(long, value_name = "NAME")]
        title: Option<String>,
        /// 図形・画像・コントロールを読むブック（省略時は対象ブック）。実行で保存されたブックを、実行前のブックの図形と重ねて表示するときに指定する
        #[arg(long, value_name = "BOOK")]
        drawings: Option<PathBuf>,
        /// 最初に表示するシート（コード名かシート名。両方に当たるときはコード名。例: Sheet1）
        #[arg(long, value_name = "SHEET")]
        sheet: Option<String>,
        /// エディタのパネルで、セルの編集を受け付けるページにする（編集は `ras book edit` が行う）
        #[arg(long, hide = true)]
        editable: bool,
    },
    /// ブックのセルを Excel と同じように編集して、ファイルに書き込む（入力・クリア・コピーと移動・セルと行と列の挿入と削除）
    Edit {
        /// 対象ブック（省略時は ras.toml の project.workbook）
        file: Option<PathBuf>,
        /// 編集するシート（コード名・シート名・1 から始まる番号）。編集がシートを指定していないときに使う
        #[arg(long, value_name = "SHEET")]
        sheet: Option<String>,
        /// セルに入力する（例: --set B2=100、--set "C2==B2*2"、--set "A1:A5=x"）。複数指定できる
        #[arg(long, value_name = "RANGE=TEXT")]
        set: Vec<String>,
        /// セルの数式と値をクリアする（例: --clear A1:C3）。複数指定できる
        #[arg(long, value_name = "RANGE")]
        clear: Vec<String>,
        /// 行・列・セルを挿入する（例: --insert 3:4、--insert B:C、--insert B2:C3）。セルは --shift で動かす向きを指定する
        #[arg(long, value_name = "RANGE")]
        insert: Vec<String>,
        /// 行・列・セルを削除する（例: --delete 3:4、--delete B:C、--delete B2:C3）
        #[arg(long, value_name = "RANGE")]
        delete: Vec<String>,
        /// セルの挿入・削除で、ほかのセルを動かす向き（挿入: down / right、削除: up / left。既定は上下）
        #[arg(long, value_name = "DIR")]
        shift: Option<String>,
        /// 編集を JSON で指定する（1 つの編集か、その配列。例: '{"op":"copy","from":"A1:B2","to":"D1"}'）。複数指定できる
        #[arg(long, value_name = "JSON")]
        op: Vec<String>,
        /// 編集の JSON を読むファイル（`-` で標準入力）
        #[arg(long, value_name = "FILE")]
        ops_file: Option<PathBuf>,
        /// 書き換えたブックを別のファイルに書く（元のブックは変えない）
        #[arg(short, long, value_name = "FILE")]
        output: Option<PathBuf>,
        /// 書き換えた後のブックのグリッド（`ras book view --stdout` と同じ HTML）を標準出力に出す
        #[arg(long)]
        view: bool,
        /// --view の HTML を、セルの編集を受け付けるページにする
        #[arg(long, hide = true)]
        editable: bool,
        /// --view の HTML で、行数・列数の上限を付けずに表示する
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
    let ext = path
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();
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
            let sheet = sheet
                .strip_prefix('\'')
                .and_then(|s| s.strip_suffix('\''))
                .map(|s| s.replace("''", "'"))
                .unwrap_or_else(|| sheet.to_string());
            (Some(sheet), range)
        }
        None => (None, text),
    }
}

pub fn dump(
    file: &Path,
    sheets: &[String],
    range: Option<&str>,
    format: DumpFormat,
    formats: bool,
    limits: Option<(u32, u32)>,
) -> Result<ExitCode, String> {
    let mut book = open(file)?;
    let name = file
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();
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
        area = Some(RangeRef::parse(&r.replace('$', "")).ok_or_else(|| {
            format!("範囲 `{r}` を読み取れません（A1:F20 のように指定してください）")
        })?);
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
            let all: Vec<String> = (0..book.sheet_count())
                .filter_map(|i| book.sheet_name(i).map(str::to_string))
                .collect();
            format!(
                "シート `{key}` が見つかりません（シート: {}）",
                all.join(", ")
            )
        })?;
        parts.push(dump::sheet(&mut book, &name, si, area, &opts)?);
    }
    match format {
        // Several sheets: one JSON document.
        DumpFormat::Json if parts.len() > 1 => {
            let list: Vec<serde_json::Value> = parts
                .iter()
                .filter_map(|p| serde_json::from_str(p).ok())
                .collect();
            println!(
                "{}",
                serde_json::to_string_pretty(&list).unwrap_or_default()
            );
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
pub fn diff_cmd(
    before: &Path,
    after: &Path,
    format: DiffFormat,
    max_cells: usize,
) -> Result<ExitCode, String> {
    let d = compare(before, after)?;
    match format {
        DiffFormat::Text => print!("{}", diff::text(&d, max_cells)),
        DiffFormat::Json => println!(
            "{}",
            serde_json::to_string_pretty(&diff::to_json(&d)).unwrap_or_default()
        ),
    }
    Ok(if d.is_empty() {
        ExitCode::SUCCESS
    } else {
        ExitCode::from(1)
    })
}

/// The page that draws `file` as a grid, with what differs from `before` marked. The objects of
/// the sheets' drawings (shapes, pictures, controls) are those of `drawings` when it is given: a
/// workbook a run saved can draw some of them differently (umya-spreadsheet, which the stand-in
/// saves with, does not keep all of what drawings hold), and the run does not change them. `sheet` (a sheet's name or
/// code name) is the sheet shown first.
pub fn view_page(
    file: &Path,
    before: Option<&Path>,
    all: bool,
    title: Option<&str>,
    drawings: Option<&Path>,
    sheet: Option<&str>,
    editable: bool,
) -> Result<String, String> {
    let mut book = open(file)?;
    let focus = match sheet {
        Some(wanted) => Some(
            sheet_named(&book, wanted)
                .ok_or_else(|| format!("{}: シート `{wanted}` がありません", file.display()))?,
        ),
        None => None,
    };
    let name = title.map(str::to_string).unwrap_or_else(|| {
        file.file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default()
    });
    let changes = match before {
        // (Compared with a snapshot of itself, it is "the last time".)
        Some(b) => Some((
            if title.is_some() {
                "前回".to_string()
            } else {
                b.file_name()
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_default()
            },
            diff::diff(&mut open(b)?, &mut book)?,
        )),
        None => None,
    };
    let limits = if all {
        None
    } else {
        Some((ras_workbook::book::VIEW_ROWS, ras_workbook::book::VIEW_COLS))
    };
    // The drawings (shapes, pictures, controls) are read from the file itself.
    let source = drawings
        .and_then(|d| std::fs::read(d).ok())
        .or_else(|| std::fs::read(file).ok());
    let mut data = view::data(
        &mut book,
        &name,
        source.as_deref(),
        limits,
        changes.as_ref().map(|(n, d)| (n.as_str(), d)),
    )?;
    if let Some(focus) = focus {
        data["active"] = focus.into();
        data["focus"] = true.into();
    }
    if editable {
        data["editable"] = true.into();
        // For the completion of formulas typed into the cells.
        data["functions"] = view::functions();
    }
    Ok(view::page(&data))
}

/// The name of the sheet whose code name or name is `wanted` (the code name first, as the
/// editor asks by code name: `Sheet1` is the sheet of `objects/Sheet1.ras` even when another sheet
/// is named "Sheet1"; case does not matter).
fn sheet_named(book: &Book, wanted: &str) -> Option<String> {
    let count = book.sheet_count();
    let by_code = (0..count).find(|&si| {
        book.code_name(si)
            .is_some_and(|c| c.eq_ignore_ascii_case(wanted))
    });
    let si = by_code.or_else(|| {
        (0..count).find(|&si| {
            book.sheet_name(si)
                .is_some_and(|n| n.eq_ignore_ascii_case(wanted))
        })
    })?;
    book.sheet_name(si).map(str::to_string)
}

/// Writes the page for `file` into `dir` (`<name>.html`); returns its path.
pub fn write_view(
    file: &Path,
    before: Option<&Path>,
    all: bool,
    output: Option<&Path>,
    dir: &Path,
    drawings: Option<&Path>,
    sheet: Option<&str>,
) -> Result<PathBuf, String> {
    let page = view_page(file, before, all, None, drawings, sheet, false)?;
    let target = match output {
        Some(o) => o.to_path_buf(),
        None => dir.join(format!(
            "{}.html",
            file.file_stem()
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_else(|| "book".into())
        )),
    };
    if let Some(parent) = target.parent().filter(|p| !p.as_os_str().is_empty()) {
        std::fs::create_dir_all(parent).map_err(|e| format!("{}: {e}", parent.display()))?;
    }
    std::fs::write(&target, page).map_err(|e| format!("{}: {e}", target.display()))?;
    Ok(target)
}

/// What `ras book edit` is asked to do.
pub struct EditArgs<'a> {
    pub sheet: Option<&'a str>,
    pub set: &'a [String],
    pub clear: &'a [String],
    pub insert: &'a [String],
    pub delete: &'a [String],
    pub shift: Option<&'a str>,
    pub op: &'a [String],
    pub ops_file: Option<&'a Path>,
}

/// The edits of the command line, in the order: `--set`, `--clear`, `--insert`, `--delete`, then
/// those written as JSON.
fn edit_ops(args: &EditArgs) -> Result<Vec<patch::Op>, String> {
    use serde_json::json;
    let mut list: Vec<serde_json::Value> = Vec::new();
    for entry in args.set {
        let (range, text) = entry.split_once('=').ok_or_else(|| {
            format!(
                "--set `{entry}`: `B2=100` のように、範囲と入力する内容を `=` で区切ってください"
            )
        })?;
        list.push(json!({ "op": "set", "range": range.trim(), "text": text }));
    }
    for range in args.clear {
        list.push(json!({ "op": "clear", "range": range }));
    }
    for (kind, ranges) in [("insert", args.insert), ("delete", args.delete)] {
        for range in ranges {
            let mut op = json!({ "op": kind, "range": range });
            if let Some(shift) = args.shift {
                op["shift"] = shift.into();
            }
            list.push(op);
        }
    }
    let mut texts: Vec<String> = args.op.to_vec();
    if let Some(file) = args.ops_file {
        texts.push(if file.as_os_str() == "-" {
            let mut text = String::new();
            std::io::Read::read_to_string(&mut std::io::stdin(), &mut text)
                .map_err(|e| format!("標準入力を読めません: {e}"))?;
            text
        } else {
            std::fs::read_to_string(file).map_err(|e| format!("{}: {e}", file.display()))?
        });
    }
    for text in texts {
        let value: serde_json::Value = serde_json::from_str(&text)
            .map_err(|e| format!("編集の JSON を読み取れません: {e}"))?;
        match value {
            serde_json::Value::Array(items) => list.extend(items),
            other => list.push(other),
        }
    }
    if list.is_empty() {
        return Err(
            "編集が指定されていません（--set / --clear / --insert / --delete / --op / --ops-file）"
                .into(),
        );
    }
    patch::parse_ops(&serde_json::Value::Array(list), args.sheet)
}

/// `ras book edit`: makes the edits in the workbook file and writes it (to `output`, or over the
/// file itself with a copy of the previous one in `backup`). Returns the path written.
///
/// Errors: `Err(Err(..))` when the workbook cannot be read at all, `Err(Ok(..))` when the edit
/// cannot be made (what is asked is refused, the workbook is open in Excel); nothing is written.
pub fn edit(
    file: &Path,
    args: &EditArgs,
    output: Option<&Path>,
    backup: Option<(&Path, usize)>,
) -> Result<PathBuf, Result<String, String>> {
    let first = open(file).map_err(Err)?.sheet_name(0).map(str::to_string);
    edit_file(file, args, first, output, backup).map_err(Ok)
}

fn edit_file(
    file: &Path,
    args: &EditArgs,
    first: Option<String>,
    output: Option<&Path>,
    backup: Option<(&Path, usize)>,
) -> Result<PathBuf, String> {
    // Without a sheet named, the edits are of the first one.
    let mut edit_args = EditArgs { ..*args };
    let sheet = args.sheet.map(str::to_string).or(first);
    edit_args.sheet = sheet.as_deref();
    let ops = edit_ops(&edit_args)?;
    let locked = || {
        format!(
            "{} は Excel で開かれているため、ファイルを書き換えられません（Excel でブックを閉じてから行ってください）",
            file.display()
        )
    };
    let bytes = std::fs::read(file).map_err(|e| format!("{}: {e}", file.display()))?;
    let new = patch::apply(&bytes, &ops)?;
    if let Some(output) = output {
        if let Some(parent) = output.parent().filter(|p| !p.as_os_str().is_empty()) {
            std::fs::create_dir_all(parent).map_err(|e| format!("{}: {e}", parent.display()))?;
        }
        std::fs::write(output, new).map_err(|e| format!("{}: {e}", output.display()))?;
        return Ok(output.to_path_buf());
    }
    if new == bytes {
        return Ok(file.to_path_buf());
    }
    ras_ovba::replace_workbook(file, &new, backup.map(|b| b.0), backup.map_or(0, |b| b.1))
        .map_err(|e| match e {
            ras_ovba::OvbaError::Locked => locked(),
            other => format!("{}: {other}", file.display()),
        })?;
    Ok(file.to_path_buf())
}

/// Opens a file with the program the system has for it (a page: the browser).
pub fn open_in_browser(path: &Path) -> Result<(), String> {
    let mut cmd = if cfg!(windows) {
        let mut c = std::process::Command::new("cmd");
        c.args(["/c", "start", ""]);
        c
    } else if cfg!(target_os = "macos") {
        std::process::Command::new("open")
    } else {
        std::process::Command::new("xdg-open")
    };
    cmd.arg(path)
        .stdin(std::process::Stdio::null())
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null());
    cmd.spawn()
        .map(|_| ())
        .map_err(|e| format!("{} を開けませんでした: {e}", path.display()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_sheet_is_found_by_its_code_name_first() {
        // "集計" has the code name Sheet1; the sheet named "Sheet1" has Sheet2.
        let mut book = Book::new_book();
        book.rename_sheet(0, "集計").unwrap();
        book.add_sheet(1, Some("Sheet1")).unwrap();
        assert_eq!(book.code_name(1).as_deref(), Some("Sheet2"));
        assert_eq!(sheet_named(&book, "sheet1").as_deref(), Some("集計"));
        assert_eq!(sheet_named(&book, "Sheet2").as_deref(), Some("Sheet1"));
        assert_eq!(sheet_named(&book, "集計").as_deref(), Some("集計"));
        assert_eq!(sheet_named(&book, "Sheet9"), None);
    }
}
