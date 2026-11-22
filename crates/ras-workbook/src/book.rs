//! A workbook as the debug run time's stand-in for Excel sees it (spec §14.10.8).
//!
//! Two copies of the workbook are kept:
//!
//! * the formula engine's model (`gridcraft_model::Workbook`), which holds the values and
//!   formulas of the cells and is what the program reads and writes. Every write is noted and
//!   what depends on it is calculated again before the next read, as Excel does in automatic
//!   calculation.
//! * the file as umya-spreadsheet read it, which holds everything else (styles, number formats,
//!   page setup, ...). The formats the program sets go there (`book/format.rs`), and the cells
//!   are copied into it when the workbook is saved.
//!
//! What moves cells is done to both: cells inserted and deleted (`book/edit.rs`), cut, copied
//! and pasted (`book/paste.rs`), sorted and filtered (`book/sort.rs`).
//!
//! How the file is read into the engine follows xljs (`src/api/xlsx_calc.rs`), from which the
//! engine (`vendor/gridcraft`) is taken: shared, array and dynamic-array formulas, implicit
//! intersection of formulas saved without dynamic-array metadata, and formulas the engine would
//! not work out as Excel does (unknown functions, external references, data tables) keep the
//! value they were saved with.
//!
//! Positions are 0-based here (`CellRef`, `RangeRef`); the addon converts.

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read as _};
use std::path::Path;
use std::sync::Arc;

use gridcraft_calc::Calc;
use gridcraft_core::{Array, CellError, CellRef, DateSystem, RangeRef, Value};
use gridcraft_formula::{Expr, SheetSel};
use gridcraft_model as gm;
use gridcraft_numfmt::NumberFormat;
use gridcraft_xlsx_formulas::{FunctionInfo, Names, SaveAs, apply_legacy, fmla, save_as};
use umya_spreadsheet::{
    Cell, CellErrorType, CellFormula, CellFormulaValues, CellRawValue, EnumTrait, Worksheet,
};

pub type Result<T> = std::result::Result<T, String>;

/// What the engine asks about worksheet functions when it reads formulas saved without
/// dynamic-array metadata.
const FUNCTIONS: FunctionInfo = FunctionInfo {
    is_builtin: |name| gridcraft_functions::lookup(name).is_some(),
    is_scalar_arg: |name, i| gridcraft_functions::lookup(name).is_none_or(|s| s.is_scalar_arg(i)),
};

/// Functions the engine has only a stand-in for.
const STAND_INS: &[&str] = &[
    "CUBEKPIMEMBER",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "CUBEVALUE",
    "FILTERXML",
    "WEBSERVICE",
];

pub const MAX_ROWS: u32 = gridcraft_core::addr::MAX_ROWS;
pub const MAX_COLS: u32 = gridcraft_core::addr::MAX_COLS;

/// The number format Excel gives a cell when a date is written into it: the short date of the
/// system (the built-in formats 14 and 22), which `NumberFormat` spells as the United States
/// do. What the cell shows, and `NumberFormatLocal`, are the system's: Japanese here, as
/// everywhere in the stand-in (`local_format`).
const DATE_FORMAT: &str = "m/d/yyyy";
const DATE_TIME_FORMAT: &str = "m/d/yyyy h:mm";
const TIME_FORMAT: &str = "h:mm:ss";

/// The number format a Japanese Excel shows a cell by, and says its `NumberFormatLocal` is.
pub fn local_format(code: &str) -> &str {
    match code {
        DATE_FORMAT => "yyyy/m/d",
        DATE_TIME_FORMAT => "yyyy/m/d h:mm",
        "General" => "G/標準",
        other => other,
    }
}

/// The built-in number formats (a `numFmtId` without a `numFmt` of the file) whose code a
/// Japanese Excel 2021 shows differently from the United States' (read from Excel itself).
/// Formats with Japanese eras (27-29, 36, 50, 51, 54, 57, 58) are left out: they are not drawn.
fn local_builtin(id: u32) -> Option<&'static str> {
    Some(match id {
        5 | 63 => "¥#,##0;¥-#,##0",
        6 | 64 => "¥#,##0;[Red]¥-#,##0",
        7 | 65 => "¥#,##0.00;¥-#,##0.00",
        8 | 66 => "¥#,##0.00;[Red]¥-#,##0.00",
        14 | 71 | 72 => "yyyy/m/d",
        22 | 78 => "yyyy/m/d h:mm",
        23 => "$#,##0_);($#,##0)",
        24 => "$#,##0_);[Red]($#,##0)",
        25 => "$#,##0.00_);($#,##0.00)",
        26 => "$#,##0.00_);[Red]($#,##0.00)",
        30 => "m/d/yy",
        31 => "yyyy\"年\"m\"月\"d\"日\"",
        32 => "h\"時\"mm\"分\"",
        33 => "h\"時\"mm\"分\"ss\"秒\"",
        34 | 52 | 55 => "yyyy\"年\"m\"月\"",
        35 | 53 | 56 => "m\"月\"d\"日\"",
        37 => "#,##0;-#,##0",
        38 => "#,##0;[Red]-#,##0",
        39 => "#,##0.00;-#,##0.00",
        40 => "#,##0.00;[Red]-#,##0.00",
        41 => "_ * #,##0_ ;_ * -#,##0_ ;_ * \"-\"_ ;_ @_ ",
        42 => "_ ¥* #,##0_ ;_ ¥* -#,##0_ ;_ ¥* \"-\"_ ;_ @_ ",
        43 => "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * \"-\"??_ ;_ @_ ",
        44 => "_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* \"-\"??_ ;_ @_ ",
        59 => "0",
        60 => "0.00",
        61 => "#,##0",
        62 => "#,##0.00",
        67 => "0%",
        68 => "0.00%",
        69 => "# ?/?",
        70 => "# ??/??",
        73 => "d-mmm-yy",
        74 => "d-mmm",
        75 => "mmm-yy",
        76 => "h:mm",
        77 => "h:mm:ss",
        79 => "mm:ss",
        80 => "[h]:mm:ss",
        81 => "mm:ss.0",
        _ => return None,
    })
}

/// A number under 1E-4 in the General format, as Excel shows it when all its digits fit in 11
/// characters in fixed notation (0.000012345, not 1.2345E-05); None for the others.
fn general_small(n: f64) -> Option<String> {
    if n == 0.0 || !n.is_finite() || n.abs() >= 1e-4 {
        return None;
    }
    // The 15 significant digits Excel keeps.
    let sci = format!("{:.14e}", n.abs());
    let (mantissa, exp) = sci.split_once('e')?;
    let exp: i32 = exp.parse().ok()?;
    let digits = mantissa.replace('.', "");
    let digits = digits.trim_end_matches('0');
    let zeros = usize::try_from(-exp - 1).ok()?;
    // (Excel's 11 characters do not count the minus sign: -0.000012345.)
    let s = format!("0.{}{digits}", "0".repeat(zeros));
    let sign = if n < 0.0 { "-" } else { "" };
    (s.len() <= 11).then(|| format!("{sign}{s}"))
}

// ---------------------------------------------------------------------------- values

/// A value as the program sees it (VBA's `Variant` of a cell).
#[derive(Clone, Debug, PartialEq)]
pub enum Val {
    Empty,
    Num(f64),
    Text(String),
    Bool(bool),
    /// A serial date (the cell's number format is a date format).
    Date(f64),
    /// VBA's `CVErr` number (`xlErrNA` = 2042).
    Err(i32),
}

/// VBA's error number of a cell error (`XlCVError`).
pub fn error_number(e: CellError) -> i32 {
    match e {
        CellError::Null => 2000,
        CellError::Div0 => 2007,
        CellError::Value => 2015,
        CellError::Ref => 2023,
        CellError::Name => 2029,
        CellError::Num => 2036,
        CellError::NA => 2042,
        CellError::GettingData => 2043,
        CellError::Spill => 2045,
        CellError::Blocked => 2047,
        CellError::Calc => 2050,
        _ => 2015,
    }
}

fn error_of_number(n: i32) -> CellError {
    match n {
        2000 => CellError::Null,
        2007 => CellError::Div0,
        2023 => CellError::Ref,
        2029 => CellError::Name,
        2036 => CellError::Num,
        2042 => CellError::NA,
        2043 => CellError::GettingData,
        2045 => CellError::Spill,
        2047 => CellError::Blocked,
        2050 => CellError::Calc,
        _ => CellError::Value,
    }
}

/// The value of an engine value for the program. `date`: the cell shows a date.
pub fn val_of(v: &Value, date: bool) -> Val {
    match v {
        Value::Empty => Val::Empty,
        Value::Number(n) if date => Val::Date(*n),
        Value::Number(n) => Val::Num(*n),
        Value::Text(s) => Val::Text(s.to_string()),
        Value::Bool(b) => Val::Bool(*b),
        // Excel shows 0 for a circular reference (and warns).
        Value::Error(CellError::Circ) => Val::Num(0.0),
        Value::Error(e) => Val::Err(error_number(*e)),
        Value::Array(a) => val_of(a.data.first().unwrap_or(&Value::Empty), date),
    }
}

pub fn value_of_val(v: &Val) -> Value {
    match v {
        Val::Empty => Value::Empty,
        Val::Num(n) | Val::Date(n) => Value::number(*n),
        Val::Text(s) => Value::text(s.as_str()),
        Val::Bool(b) => Value::Bool(*b),
        Val::Err(n) => Value::Error(error_of_number(*n)),
    }
}

fn error_to(e: CellError) -> CellErrorType {
    match e {
        CellError::Null => CellErrorType::Null,
        CellError::Div0 => CellErrorType::Div0,
        CellError::Ref => CellErrorType::Ref,
        CellError::Name => CellErrorType::Name,
        CellError::Num => CellErrorType::Num,
        CellError::NA | CellError::GettingData => CellErrorType::NA,
        _ => CellErrorType::Value,
    }
}

fn error_of(e: &CellErrorType) -> CellError {
    CellError::parse(&e.to_string()).unwrap_or(CellError::Value)
}

/// The value of a cell of the file for the engine. `formula`: the cell has one, so an empty
/// text is a result ("") and not a blank cell.
fn file_value(c: &Cell, formula: bool) -> Value {
    match c.raw_value() {
        CellRawValue::Empty => Value::Empty,
        CellRawValue::Numeric(n) => Value::number(*n),
        CellRawValue::Bool(v) => Value::Bool(*v),
        CellRawValue::Error(e) => Value::Error(error_of(e)),
        other => {
            let s = other.to_string();
            if s.is_empty() && !formula {
                Value::Empty
            } else {
                Value::text(s)
            }
        }
    }
}

fn set_result(c: &mut Cell, v: &Value) {
    match v {
        Value::Empty => {
            c.set_formula_result_blank();
        }
        Value::Number(n) => {
            c.set_formula_result_number(*n);
        }
        Value::Bool(x) => {
            c.set_formula_result_bool(*x);
        }
        Value::Error(e) => {
            c.set_formula_result_error(error_to(*e));
        }
        Value::Text(s) => {
            c.set_formula_result_string(s.to_string());
        }
        Value::Array(_) => set_result(c, &v.scalar()),
    }
}

fn set_plain(c: &mut Cell, v: &Value) {
    let cv = c.cell_value_mut();
    match v {
        Value::Empty => {
            cv.set_blank();
        }
        Value::Number(n) => {
            cv.set_value_number(*n);
        }
        Value::Bool(x) => {
            cv.set_value_bool(*x);
        }
        Value::Error(e) => {
            cv.remove_formula();
            cv.set_error(error_to(*e).to_string());
        }
        Value::Text(s) => {
            cv.set_value_string(s.to_string());
        }
        Value::Array(_) => set_plain(c, &v.scalar()),
    }
}

fn at_of(c: &Cell) -> CellRef {
    let co = c.coordinate();
    CellRef::new(
        co.row_num().saturating_sub(1),
        co.col_num().saturating_sub(1),
    )
}

/// umya-spreadsheet's (column, row), 1-based.
fn coord(at: CellRef) -> (u32, u32) {
    (at.col + 1, at.row + 1)
}

// ---------------------------------------------------------------------------- reading the file

/// How a formula cell was saved.
#[derive(Clone, Copy, PartialEq)]
enum Kind {
    Legacy,
    Dynamic(RangeRef),
    Array(RangeRef),
}

struct Made {
    cell: gm::Cell,
    kind: Kind,
}

/// A cell of the file as the engine takes it. A formula the engine cannot read comes back
/// without one (the saved value stays).
fn make(c: &Cell) -> Made {
    let plain = |formula: bool| Made {
        cell: gm::Cell::value(file_value(c, formula)),
        kind: Kind::Legacy,
    };
    let Some(f) = c.formula_obj() else {
        return plain(false);
    };
    let text = f.text().trim();
    if *f.formula_type() == CellFormulaValues::DataTable || text.is_empty() {
        return plain(true);
    }
    let body = text.strip_prefix('=').unwrap_or(text);
    if gridcraft_formula::parse(body).is_err() {
        return plain(true);
    }
    let mut formula = gm::Formula::new(&fmla::from_file(body));
    let at = at_of(c);
    let dynamic = f.dynamic_array();
    let kind = if *f.formula_type() == CellFormulaValues::Array {
        let area = RangeRef::parse(f.reference())
            .filter(|r| r.start == at)
            .unwrap_or(RangeRef::cell(at));
        if dynamic {
            Kind::Dynamic(area)
        } else {
            formula.array = Some(area);
            Kind::Array(area)
        }
    } else if dynamic {
        Kind::Dynamic(RangeRef::cell(at))
    } else {
        Kind::Legacy
    };
    formula.dynamic = dynamic;
    Made {
        cell: gm::Cell {
            value: file_value(c, true),
            formula: Some(Arc::new(formula)),
            ..Default::default()
        },
        kind,
    }
}

/// What a sheet has besides its cells that formulas depend on: tables, merged cells, hidden rows.
fn layout(ws: &Worksheet, id: &mut u32) -> gm::Sheet {
    let mut sh = gm::Sheet::new(ws.name());
    let range = |s: &str| RangeRef::parse(&s.replace('$', ""));
    for m in ws.merge_cells() {
        if let Some(r) = range(&m.range()) {
            sh.merges.push(r);
        }
    }
    let mut hidden: Vec<u32> = Vec::new();
    for r in ws.row_dimensions() {
        if r.hidden() && r.row_num() >= 1 {
            hidden.push(r.row_num() - 1);
            sh.rows.insert(
                r.row_num() - 1,
                gm::LineInfo {
                    hidden: true,
                    ..Default::default()
                },
            );
        }
    }
    for c in ws.column_dimensions() {
        if c.hidden() && c.col_num() >= 1 {
            sh.cols.insert(
                c.col_num() - 1,
                gm::LineInfo {
                    hidden: true,
                    ..Default::default()
                },
            );
        }
    }
    let filtered = |r: &RangeRef| hidden.iter().any(|h| *h > r.start.row && *h <= r.end.row);
    let criterion = || {
        vec![(
            0u32,
            gm::FilterCriterion::Values {
                values: Vec::new(),
                blanks: false,
            },
        )]
    };
    if let Some((f, r)) = ws
        .auto_filter()
        .and_then(|f| Some((f, range(&f.range().range())?)))
    {
        let mut criteria: Vec<(u32, gm::FilterCriterion)> = f
            .filter_columns()
            .iter()
            .filter_map(|c| Some((c.column_id(), file_criterion(c.criteria()?)?)))
            .collect();
        if criteria.is_empty() && filtered(&r) {
            // Filtered by something the file's reader does not keep (colours, icons): the rows
            // it hides are filtered rows all the same.
            criteria = criterion();
        }
        sh.autofilter = Some(gm::AutoFilter { range: r, criteria });
    }
    for t in ws.tables() {
        if !t.is_ok() {
            continue;
        }
        let (a, z) = t.area();
        let r = RangeRef::new(
            CellRef::new(a.row_num() - 1, a.col_num() - 1),
            CellRef::new(z.row_num() - 1, z.col_num() - 1),
        );
        *id += 1;
        sh.tables.push(gm::Table {
            id: *id,
            name: (if t.display_name().is_empty() {
                t.name()
            } else {
                t.display_name()
            })
            .to_string(),
            range: r,
            header_row: true,
            totals_row: t.totals_row_shown() || t.totals_row_count() > 0,
            columns: t
                .columns()
                .iter()
                .map(|c| gm::TableColumn {
                    name: c.name().to_string(),
                    totals: Default::default(),
                    totals_label: None,
                    formula: None,
                    unique_name: None,
                    xml: None,
                })
                .collect(),
            style: String::new(),
            banded_rows: false,
            banded_cols: false,
            first_col: false,
            last_col: false,
            filter_button: true,
            filter: if filtered(&r) {
                criterion()
            } else {
                Vec::new()
            },
            table_type: None,
            connection_id: None,
        });
    }
    sh
}

/// What a column of a filter of the file shows, for the engine.
fn file_criterion(c: &umya_spreadsheet::FilterCriteria) -> Option<gm::FilterCriterion> {
    use umya_spreadsheet::FilterCriteria as F;
    Some(match c {
        F::Values { values, blank } => gm::FilterCriterion::Values {
            values: values.clone(),
            blanks: *blank,
        },
        F::Custom { tests, and } => gm::FilterCriterion::Custom {
            a: tests.first()?.clone(),
            b: tests.get(1).cloned(),
            and: *and,
        },
        F::Top10 { top, percent, val } => gm::FilterCriterion::Top10 {
            bottom: !*top,
            percent: *percent,
            count: *val as u32,
        },
    })
}

/// The workbook for the engine without its cells: sheets, names, calculation options.
fn outline(file: &umya_spreadsheet::Workbook, date1904: bool) -> gm::Workbook {
    let mut wb = gm::Workbook::new();
    wb.date_system = if date1904 {
        DateSystem::D1904
    } else {
        DateSystem::D1900
    };
    let options = file.calculation_properties();
    wb.calc.iterative = options.iterate();
    wb.calc.set_max_iterations(options.iterate_count());
    wb.calc.set_max_change(options.iterate_delta());
    let sheets = file.sheet_collection();
    let mut id = 0;
    wb.sheets = sheets
        .iter()
        .map(|ws| Arc::new(layout(ws, &mut id)))
        .collect();
    let mut seen: HashSet<(String, Option<usize>)> = HashSet::new();
    let global = file.defined_names().iter().map(|d| (None, d));
    let local = sheets.iter().enumerate().flat_map(|(i, ws)| {
        ws.defined_names()
            .iter()
            .map(move |d| (d.has_local_sheet_id().then_some(i), d))
    });
    for (scope, d) in global.chain(local) {
        let name = d.name();
        if name.is_empty()
            || name.to_ascii_lowercase().starts_with("_xlnm.")
            || !seen.insert((name.to_lowercase(), scope))
        {
            continue;
        }
        wb.names.push(gm::DefinedName {
            name: name.to_string(),
            scope,
            formula: fmla::from_file(&d.address()),
            comment: String::new(),
            hidden: d.hidden(),
        });
    }
    wb
}

/// Whether the engine would not work out `e` as Excel does.
fn unsupported(wb: &gm::Workbook, sheet: usize, e: &Expr) -> bool {
    let mut locals: Vec<&str> = Vec::new();
    e.walk(&mut |x| {
        if let Expr::Call(name, args) = x
            && (name == "LET" || name == "LAMBDA")
        {
            let last = args.len().saturating_sub(1);
            for (i, a) in args.iter().enumerate() {
                if let Expr::Name(n) = a
                    && i < last
                    && (name == "LAMBDA" || i % 2 == 0)
                {
                    locals.push(n);
                }
            }
        }
    });
    let local = |n: &str| locals.iter().any(|l| gridcraft_formula::same_name(l, n));
    let external = |s: &str| s.contains('[');
    let check = |x: &Expr| -> bool {
        match x {
            Expr::Call(name, _) => {
                let known = gridcraft_calc::is_known_function(name)
                    && !STAND_INS.contains(&name.to_ascii_uppercase().as_str());
                !(known || local(name) || wb.name(name, sheet).is_some())
            }
            Expr::Ref(r) => match &r.sheet {
                SheetSel::Named(s) => external(s),
                SheetSel::Span(a, z) => external(a) || external(z),
                SheetSel::Current => false,
            },
            Expr::Name(n) => {
                let (scope, bare) = match n.rsplit_once('!') {
                    Some((s, _)) if external(s) => return true,
                    Some((s, bare)) => (
                        wb.sheet_index(&gridcraft_core::unquote_sheet(s))
                            .unwrap_or(sheet),
                        bare,
                    ),
                    None => (sheet, n.as_str()),
                };
                !(local(bare) || wb.name(bare, scope).is_some() || wb.table(bare).is_some())
            }
            Expr::Struct(st) => !st.table.is_empty() && wb.table(&st.table).is_none(),
            _ => false,
        }
    };
    let mut found = false;
    e.walk(&mut |x| {
        if !found {
            found = check(x);
        }
    });
    found
}

fn read_date1904(path: &Path) -> bool {
    let Ok(file) = std::fs::File::open(path) else {
        return false;
    };
    date1904_of(file)
}

fn date1904_of<R: std::io::Read + std::io::Seek>(file: R) -> bool {
    let Ok(mut zip) = zip::ZipArchive::new(file) else {
        return false;
    };
    let Ok(mut part) = zip.by_name("xl/workbook.xml") else {
        return false;
    };
    let mut s = String::new();
    if part.read_to_string(&mut s).is_err() {
        return false;
    }
    s.find("<workbookPr")
        .and_then(|i| s[i..].find('>').map(|j| &s[i..i + j]))
        .is_some_and(|tag| tag.contains("date1904=\"1\"") || tag.contains("date1904=\"true\""))
}

fn seed() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_nanos() as u64)
        .unwrap_or(1)
}

// ---------------------------------------------------------------------------- the workbook

/// What `get` reads of the cells of a range.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Read {
    /// `Range.Value`: dates as dates.
    Value,
    /// `Range.Value2`: dates as numbers.
    Value2,
    /// `Range.Formula`: formulas as formulas saved without dynamic-array metadata are written.
    Formula,
    /// `Range.Formula2`: formulas as dynamic-array Excel shows them.
    Formula2,
    /// `Range.FormulaR1C1`.
    FormulaR1C1,
    /// `Range.Text`: the values as the cells show them.
    Text,
    /// `Range.NumberFormat`.
    NumberFormat,
}

/// The direction of `Range.End`.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Dir {
    Up,
    Down,
    Left,
    Right,
}

pub struct Book {
    file: umya_spreadsheet::Workbook,
    model: gm::Workbook,
    calc: Calc,
    /// Cells written since the engine last calculated.
    pending: Vec<gridcraft_calc::Key>,
    /// Something other than cells changed (sheets, names): the graph is built again.
    rebuild: bool,
    names: Names,
    /// Number formats parsed so far.
    formats: HashMap<String, NumberFormat>,
    /// `Application.Calculation` is manual: formulas are calculated when the program asks.
    manual: bool,
    /// The sheets (by their id in the model) whose rows `Range.AdvancedFilter` hid in place,
    /// with the rows of the list.
    advanced: HashMap<u64, RangeRef>,
}

impl Book {
    /// A new workbook with one sheet, as `Workbooks.Add` makes.
    pub fn new_book() -> Book {
        let mut file = umya_spreadsheet::new_file();
        if let Ok(ws) = file.sheet_mut(0) {
            ws.set_code_name("Sheet1");
        }
        Book::from_file(file, false)
    }

    pub fn open(path: &Path) -> Result<Book> {
        let file = umya_spreadsheet::reader::xlsx::read(path)
            .map_err(|e| format!("{} を開けません（{e}）", path.display()))?;
        Ok(Book::from_file(file, read_date1904(path)))
    }

    pub fn open_bytes(data: &[u8]) -> Result<Book> {
        let file = umya_spreadsheet::reader::xlsx::read_reader(Cursor::new(data), true)
            .map_err(|e| format!("ブックを開けません（{e}）"))?;
        Ok(Book::from_file(file, date1904_of(Cursor::new(data))))
    }

    fn from_file(mut file: umya_spreadsheet::Workbook, date1904: bool) -> Book {
        file.read_sheet_collection();
        // A workbook without a VBA project has no code names in its file (umya-spreadsheet writes
        // them with macros only): they are what Excel gives the sheets, Sheet1, Sheet2, ...
        if file
            .sheet_collection()
            .iter()
            .all(|ws| ws.code_name().is_none_or(str::is_empty))
        {
            for (i, ws) in file.sheet_collection_mut().iter_mut().enumerate() {
                ws.set_code_name(format!("Sheet{}", i + 1));
            }
        }
        // Column widths are kept as `ColumnWidth` has them (see `widths_for_file`).
        for ws in file.sheet_collection_mut() {
            let default = ws.sheet_format_properties().default_column_width();
            if default > 0.0 {
                ws.sheet_format_properties_mut()
                    .set_default_column_width(format::width_from_file(default));
            }
            for c in ws.column_dimensions_mut() {
                let w = format::width_from_file(c.width());
                c.set_width(w);
            }
        }
        let mut model = outline(&file, date1904);
        let mut calc = Calc::new();
        calc.seed(seed());
        let mut formulas: Vec<(usize, CellRef, Kind)> = Vec::new();
        let mut filled: Vec<(usize, CellRef, CellRef, Value)> = Vec::new();
        for (si, ws) in file.sheet_collection().iter().enumerate() {
            let cells = ws.cells();
            // cell -> first cell of the array formula that fills it
            let mut cover: HashMap<CellRef, CellRef> = HashMap::new();
            for c in &cells {
                let Some(f) = c
                    .formula_obj()
                    .filter(|f| *f.formula_type() == CellFormulaValues::Array)
                else {
                    continue;
                };
                let at = at_of(c);
                let Some(r) =
                    RangeRef::parse(f.reference()).filter(|r| r.start == at && !r.is_single())
                else {
                    continue;
                };
                cover.extend(r.iter().filter(|p| *p != at).map(|p| (p, at)));
            }
            let Some(sh) = model.sheet_mut(si) else {
                continue;
            };
            for c in &cells {
                let at = at_of(c);
                if !c.is_formula()
                    && let Some(first) = cover.get(&at)
                {
                    let v = file_value(c, false);
                    if !v.is_empty() {
                        filled.push((si, at, *first, v));
                    }
                    continue;
                }
                let m = make(c);
                if m.cell.formula.is_some() {
                    formulas.push((si, at, m.kind));
                }
                if m.cell.has_content() {
                    sh.cells.set(at, m.cell);
                }
            }
        }
        let mut book = Book {
            file,
            model,
            calc,
            pending: Vec::new(),
            rebuild: false,
            names: Names::default(),
            formats: HashMap::new(),
            manual: false,
            advanced: HashMap::new(),
        };
        let mut legacy: Vec<Vec<CellRef>> = vec![Vec::new(); book.model.sheets.len()];
        let mut fresh = Vec::new();
        for (si, at, kind) in formulas {
            if book.settle(si, at) {
                continue;
            }
            match kind {
                Kind::Legacy => legacy[si].push(at),
                Kind::Array(r) | Kind::Dynamic(r) if !r.is_single() => {
                    if let Some(sh) = book.model.sheet_mut(si) {
                        sh.spill_ranges.insert(at, r);
                    }
                }
                _ => {}
            }
            if book
                .model
                .sheet(si)
                .and_then(|s| s.cell(at))
                .is_some_and(|c| c.value.is_empty())
            {
                fresh.push((si, at));
            }
        }
        for (si, at, first, v) in filled {
            let Some(sh) = book.model.sheet_mut(si) else {
                continue;
            };
            if sh.spill_ranges.contains_key(&first) {
                sh.spill.insert(at, v);
            } else {
                sh.cells.set(at, gm::Cell::value(v));
            }
        }
        apply_legacy(
            &mut book.model,
            legacy
                .into_iter()
                .enumerate()
                .filter(|(_, v)| !v.is_empty())
                .collect(),
            FUNCTIONS,
        );
        // LENB and the like count a full-width character as two bytes, as Japanese Excel does.
        gridcraft_functions::set_dbcs(true);
        book.calc.rebuild(&book.model);
        // The formulas without a saved value, and those that change on their own (NOW, RAND).
        book.calc.cells_changed(&mut book.model, &fresh);
        book
    }

    /// Checks a formula against the workbook's functions, names and tables. One the engine would
    /// not work out as Excel does loses its formula and keeps its saved value: true then.
    fn settle(&mut self, si: usize, at: CellRef) -> bool {
        let Some(expr) = self
            .model
            .sheet(si)
            .and_then(|s| s.cell(at))
            .and_then(|c| c.formula.as_ref())
            .and_then(|f| f.expr_arc())
        else {
            return true;
        };
        if !unsupported(&self.model, si, &expr) {
            return false;
        }
        if let Some(sh) = self.model.sheet_mut(si) {
            match sh.cells.get_mut(at) {
                Some(c) if !c.value.is_empty() => c.formula = None,
                _ => {
                    sh.remove_cell(at);
                }
            }
        }
        true
    }

    /// Brings the values of the formulas up to date with what was written, unless calculation
    /// is manual.
    fn flush(&mut self) {
        if !self.manual {
            self.recalculate();
        }
    }

    /// `Application.Calculation`: manual stops formulas from being calculated as cells change
    /// (a formula is still worked out when it is written into a cell), until `calculate`.
    pub fn set_manual(&mut self, manual: bool) {
        if manual && !self.manual {
            // What was written while calculation was automatic is calculated.
            self.recalculate();
        }
        self.manual = manual;
    }

    pub fn manual(&self) -> bool {
        self.manual
    }

    /// Whether the workbook was saved with manual calculation.
    pub fn saved_manual(&self) -> bool {
        self.file
            .calculation_properties()
            .calc_mode()
            .value_string()
            == "manual"
    }

    /// The formulas that wait to be calculated again (`Application.Calculate`, and every read
    /// while calculation is automatic).
    pub fn recalculate(&mut self) {
        if self.rebuild {
            self.rebuild = false;
            self.pending.clear();
            self.calc.recalc_all(&mut self.model);
        } else if !self.pending.is_empty() {
            let keys = std::mem::take(&mut self.pending);
            self.calc.cells_changed(&mut self.model, &keys);
        }
    }

    /// Calculates what waits to be calculated, and tells which sheets had formulas calculated
    /// (for the `Calculate` events): those with a formula that was written, or that depends on
    /// a cell that was.
    pub fn calculated_sheets(&mut self) -> Vec<usize> {
        if self.manual {
            return Vec::new();
        }
        let mut sheets: HashSet<usize> = HashSet::new();
        if self.rebuild {
            for (i, sh) in self.model.sheets.iter().enumerate() {
                if sh.cells.iter().any(|(_, c)| c.formula.is_some()) {
                    sheets.insert(i);
                }
            }
        } else if !self.pending.is_empty() {
            let mut seen: HashSet<gridcraft_calc::Key> = HashSet::new();
            let mut queue: Vec<gridcraft_calc::Key> = self.pending.clone();
            for k in &self.pending {
                if self
                    .model
                    .sheet(k.0)
                    .and_then(|s| s.cell(k.1))
                    .is_some_and(|c| c.formula.is_some())
                {
                    sheets.insert(k.0);
                }
            }
            let mut out = Vec::new();
            while let Some(k) = queue.pop() {
                out.clear();
                self.calc.graph.dependents(k.0, k.1, &mut out);
                for d in out.drain(..) {
                    if seen.insert(d) {
                        sheets.insert(d.0);
                        queue.push(d);
                    }
                }
            }
        }
        self.recalculate();
        let mut sheets: Vec<usize> = sheets.into_iter().collect();
        sheets.sort_unstable();
        sheets
    }

    /// `Application.CalculateFull`: every formula again.
    pub fn calculate(&mut self) {
        self.pending.clear();
        self.rebuild = false;
        self.calc.recalc_all(&mut self.model);
    }

    pub fn date_system(&self) -> DateSystem {
        self.model.date_system
    }

    // ---- sheets

    pub fn sheet_count(&self) -> usize {
        self.model.sheets.len()
    }

    fn check_sheet(&self, si: usize) -> Result<()> {
        if si < self.model.sheets.len() {
            Ok(())
        } else {
            Err("インデックスが有効範囲にありません".into())
        }
    }

    pub fn sheet_name(&self, si: usize) -> Option<&str> {
        self.model.sheet(si).map(|s| s.name.as_str())
    }

    pub fn sheet_index(&self, name: &str) -> Option<usize> {
        self.model.sheet_index(name)
    }

    pub fn code_name(&self, si: usize) -> Option<String> {
        self.file
            .sheet(si)
            .ok()
            .and_then(|ws| ws.code_name().map(str::to_string))
            .filter(|s| !s.is_empty())
    }

    /// `Worksheet.Visible`: -1 visible, 0 hidden, 2 very hidden.
    pub fn visible(&self, si: usize) -> i32 {
        match self
            .file
            .sheet(si)
            .map(|ws| ws.sheet_state().to_string())
            .unwrap_or_default()
            .as_str()
        {
            "hidden" => 0,
            "veryHidden" => 2,
            _ => -1,
        }
    }

    pub fn set_visible(&mut self, si: usize, v: i32) -> Result<()> {
        self.check_sheet(si)?;
        let state = match v {
            0 => "hidden",
            2 => "veryHidden",
            _ => "visible",
        };
        if state != "visible"
            && (0..self.sheet_count())
                .filter(|i| *i != si && self.visible(*i) == -1)
                .count()
                == 0
        {
            return Err("ブックのすべてのシートを非表示にすることはできません".into());
        }
        if let Ok(ws) = self.file.sheet_mut(si) {
            ws.set_sheet_state(state.to_string());
        }
        Ok(())
    }

    pub fn rename_sheet(&mut self, si: usize, name: &str) -> Result<()> {
        self.check_sheet(si)?;
        let old = self
            .model
            .sheet(si)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        if old == name {
            return Ok(());
        }
        self.model
            .check_sheet_name(name, Some(si))
            .map_err(|e| format!("シート名 \"{name}\" は使えません（{e}）"))?;
        self.file
            .set_sheet_name(si, name)
            .map_err(|e| format!("シート名 \"{name}\" は使えません（{e}）"))?;
        if let Some(sh) = self.model.sheet_mut(si) {
            sh.name = name.to_string();
        }
        self.rewrite_formulas(|e| gridcraft_formula::adjust::rename_sheet(e, &old, name));
        Ok(())
    }

    /// Adds a sheet before position `at` (`sheet_count()` to add it last). Returns its index.
    pub fn add_sheet(&mut self, at: usize, name: Option<&str>) -> Result<usize> {
        let at = at.min(self.sheet_count());
        let name = match name {
            Some(n) => {
                self.model
                    .check_sheet_name(n, None)
                    .map_err(|e| format!("シート名 \"{n}\" は使えません（{e}）"))?;
                n.to_string()
            }
            None => {
                // The first name that is free, as Excel names a sheet added to a workbook it opened.
                let mut i = 1;
                loop {
                    let n = format!("Sheet{i}");
                    if self.model.sheet_index(&n).is_none() {
                        break n;
                    }
                    i += 1;
                }
            }
        };
        let code = {
            let used: HashSet<String> = (0..self.sheet_count())
                .filter_map(|i| self.code_name(i))
                .map(|s| s.to_lowercase())
                .collect();
            let mut i = 1;
            loop {
                let n = format!("Sheet{i}");
                if !used.contains(&n.to_lowercase()) {
                    break n;
                }
                i += 1;
            }
        };
        let ws = self
            .file
            .new_sheet(name.clone())
            .map_err(|e| format!("シートを追加できません（{e}）"))?;
        ws.set_code_name(code);
        let last = self.file.sheet_collection().len() - 1;
        self.file.sheet_collection_mut()[at..=last].rotate_right(1);
        self.model.sheets.insert(at, Arc::new(gm::Sheet::new(name)));
        self.rebuild = true;
        Ok(at)
    }

    pub fn delete_sheet(&mut self, si: usize) -> Result<()> {
        self.check_sheet(si)?;
        if self.sheet_count() == 1 {
            return Err("ブックには少なくとも 1 枚のシートが必要です".into());
        }
        let order: Vec<String> = self.model.sheets.iter().map(|s| s.name.clone()).collect();
        let name = order[si].clone();
        self.file
            .remove_sheet(si)
            .map_err(|e| format!("シートを削除できません（{e}）"))?;
        self.model.sheets.remove(si);
        for n in &mut self.model.names {
            n.scope = match n.scope {
                Some(s) if s == si => None,
                Some(s) if s > si => Some(s - 1),
                s => s,
            };
        }
        self.model
            .names
            .retain(|n| n.scope.is_some() || !n.formula.is_empty());
        self.rewrite_formulas(|e| gridcraft_formula::adjust::delete_sheet(e, &name, &order));
        Ok(())
    }

    /// `Worksheet.Copy`: a copy of sheet `si` before position `at`, named as Excel names it
    /// (`Sheet1 (2)`). Returns its index.
    pub fn copy_sheet(&mut self, si: usize, at: usize) -> Result<usize> {
        self.check_sheet(si)?;
        let at = at.min(self.sheet_count());
        let old = self
            .model
            .sheet(si)
            .map(|s| s.name.clone())
            .unwrap_or_default();
        let name = (2..)
            .map(|i| {
                let suffix = format!(" ({i})");
                let base: String = old.chars().take(31 - suffix.chars().count()).collect();
                format!("{base}{suffix}")
            })
            .find(|n| self.model.sheet_index(n).is_none())
            .unwrap_or_default();
        self.flush();
        self.store();
        let code = {
            let used: HashSet<String> = (0..self.sheet_count())
                .filter_map(|i| self.code_name(i))
                .map(|s| s.to_lowercase())
                .collect();
            (1..)
                .map(|i| format!("Sheet{i}"))
                .find(|n| !used.contains(&n.to_lowercase()))
                .unwrap_or_default()
        };
        let ws = self
            .file
            .copy_sheet(si, &name, at)
            .map_err(|e| format!("シートをコピーできません（{e}）"))?;
        ws.set_code_name(code);
        let mut copy: gm::Sheet = self.model.sheet(si).cloned().expect("checked");
        copy.name = name;
        copy.uid = gm::SheetUid::fresh();
        self.model.sheets.insert(at, Arc::new(copy));
        // The names of the sheet copied belong to the copy too; the others keep their sheets.
        let source = if at <= si { si + 1 } else { si };
        let mut added = Vec::new();
        for n in &mut self.model.names {
            if let Some(s) = n.scope.as_mut()
                && *s >= at
            {
                *s += 1;
            }
            if n.scope == Some(source) {
                added.push(gm::DefinedName {
                    scope: Some(at),
                    ..n.clone()
                });
            }
        }
        self.model.names.extend(added);
        self.relayout(at);
        self.names = Names::default();
        self.rebuild = true;
        Ok(at)
    }

    /// Moves a sheet so that it ends up at position `to`.
    pub fn move_sheet(&mut self, from: usize, to: usize) -> Result<()> {
        self.check_sheet(from)?;
        let to = to.min(self.sheet_count() - 1);
        if from == to {
            return Ok(());
        }
        let sheets = self.file.sheet_collection_mut();
        if from < to {
            sheets[from..=to].rotate_left(1);
            self.model.sheets[from..=to].rotate_left(1);
        } else {
            sheets[to..=from].rotate_right(1);
            self.model.sheets[to..=from].rotate_right(1);
        }
        self.rebuild = true;
        Ok(())
    }

    /// Rewrites every formula and name (a sheet was renamed or deleted).
    fn rewrite_formulas(&mut self, f: impl Fn(Expr) -> Expr) {
        for si in 0..self.model.sheets.len() {
            let Some(sh) = self.model.sheet_mut(si) else {
                continue;
            };
            let changes: Vec<(CellRef, gm::Formula)> = sh
                .cells
                .iter()
                .filter_map(|(at, c)| {
                    let old = c.formula.as_ref()?;
                    let e = old.expr()?;
                    let mut new = gm::Formula::from_expr(f(e));
                    new.array = old.array;
                    new.dynamic = old.dynamic;
                    (new.text != old.text).then_some((at, new))
                })
                .collect();
            for (at, formula) in changes {
                if let Some(c) = sh.cells.get_mut(at) {
                    c.formula = Some(Arc::new(formula));
                }
            }
        }
        for n in &mut self.model.names {
            if let Ok(e) = gridcraft_formula::parse(&n.formula) {
                n.formula = gridcraft_formula::print(&f(e));
            }
        }
        self.names = Names::default();
        self.rebuild = true;
    }

    // ---- names

    /// The workbook's names: (name, sheet of a sheet's own name, formula without `=`).
    pub fn names(&self) -> Vec<(String, Option<usize>, String)> {
        self.model
            .names
            .iter()
            .map(|n| (n.name.clone(), n.scope, n.formula.clone()))
            .collect()
    }

    pub fn add_name(&mut self, name: &str, scope: Option<usize>, formula: &str) -> Result<()> {
        let body = formula.trim();
        let body = body.strip_prefix('=').unwrap_or(body).to_string();
        if name.is_empty() || name.contains([' ', '!', '"']) || CellRef::parse(name).is_some() {
            return Err(format!("名前 \"{name}\" は使えません"));
        }
        self.model
            .names
            .retain(|n| !(n.scope == scope && gridcraft_formula::same_name(&n.name, name)));
        self.model.names.push(gm::DefinedName {
            name: name.to_string(),
            scope,
            formula: body,
            comment: String::new(),
            hidden: false,
        });
        self.names = Names::default();
        self.rebuild = true;
        Ok(())
    }

    pub fn delete_name(&mut self, name: &str, scope: Option<usize>) {
        self.model
            .names
            .retain(|n| !(n.scope == scope && gridcraft_formula::same_name(&n.name, name)));
        self.names = Names::default();
        self.rebuild = true;
    }

    /// The range a reference or a name refers to, as `Range("...")` reads it on sheet `si`:
    /// `A1`, `$A$1:B2`, `A:A`, `1:3`, `Sheet2!A1`, `'Sheet 2'!A1`, a defined name.
    pub fn resolve(&mut self, si: usize, text: &str) -> Option<(usize, RangeRef)> {
        let text = text.trim();
        let e = gridcraft_formula::parse(text.strip_prefix('=').unwrap_or(text)).ok()?;
        self.resolve_expr(si, &e, 0)
    }

    fn resolve_expr(&mut self, si: usize, e: &Expr, depth: usize) -> Option<(usize, RangeRef)> {
        if depth > 8 {
            return None;
        }
        match e {
            Expr::Ref(r) => {
                let sheet = match &r.sheet {
                    SheetSel::Current => si,
                    SheetSel::Named(n) => self.model.sheet_index(n)?,
                    SheetSel::Span(..) => return None,
                };
                Some((sheet, r.range()))
            }
            Expr::Paren(inner) => self.resolve_expr(si, inner, depth + 1),
            Expr::Name(n) => {
                let (scope, bare) = match n.rsplit_once('!') {
                    Some((s, bare)) => (
                        self.model.sheet_index(&gridcraft_core::unquote_sheet(s))?,
                        bare,
                    ),
                    None => (si, n.as_str()),
                };
                let formula = self.model.name(bare, scope)?.formula.clone();
                let inner = gridcraft_formula::parse(&formula).ok()?;
                self.resolve_expr(scope, &inner, depth + 1)
            }
            _ => None,
        }
    }

    // ---- cells

    fn number_format_code(&self, si: usize, at: CellRef) -> String {
        let Ok(ws) = self.file.sheet(si) else {
            return "General".into();
        };
        let code = |style: &umya_spreadsheet::Style| {
            style
                .number_format()
                .map(|f| f.format_code().to_string())
                .filter(|s| !s.is_empty())
        };
        match ws.cell(coord(at)) {
            Some(c) => code(c.style()),
            None => code(&self.style_at(si, at)),
        }
        .unwrap_or_else(|| "General".into())
    }

    /// The number format a Japanese Excel shows the cell by: a built-in format as Japanese Excel
    /// has it (`local_builtin`), the others as `local_format` spells them.
    fn shown_format_code(&self, si: usize, at: CellRef) -> String {
        let id = self.file.sheet(si).ok().and_then(|ws| {
            let style = match ws.cell(coord(at)) {
                Some(c) => c.style().clone(),
                None => self.style_at(si, at),
            };
            style.number_format().map(|f| f.number_format_id())
        });
        if let Some(code) = id.and_then(local_builtin) {
            return code.to_string();
        }
        let code = self.number_format_code(si, at);
        if code == "General" {
            return code;
        }
        local_format(&code).to_string()
    }

    fn format(&mut self, code: &str) -> &NumberFormat {
        if !self.formats.contains_key(code) {
            self.formats
                .insert(code.to_string(), NumberFormat::parse(code));
        }
        &self.formats[code]
    }

    fn is_date_format(&mut self, si: usize, at: CellRef) -> bool {
        let code = self.number_format_code(si, at);
        code != "General" && self.format(&code).is_date()
    }

    /// The value of a cell (what it shows: its own value, or what a formula spilled into it).
    pub fn value(&mut self, si: usize, at: CellRef) -> Value {
        self.flush();
        self.model
            .sheet(si)
            .map(|s| s.value(at))
            .unwrap_or_default()
    }

    /// Reads the cells of `r` (row by row).
    pub fn get(&mut self, si: usize, r: RangeRef, what: Read) -> Result<Vec<Vec<Val>>> {
        self.check_sheet(si)?;
        if r.count() > 50_000_000 {
            return Err("範囲が大きすぎます".into());
        }
        self.flush();
        let mut out = Vec::with_capacity(r.height() as usize);
        for row in r.start.row..=r.end.row {
            let mut line = Vec::with_capacity(r.width() as usize);
            for col in r.start.col..=r.end.col {
                line.push(self.read(si, CellRef::new(row, col), what));
            }
            out.push(line);
        }
        Ok(out)
    }

    fn read(&mut self, si: usize, at: CellRef, what: Read) -> Val {
        let value = self
            .model
            .sheet(si)
            .map(|s| s.value(at))
            .unwrap_or_default();
        match what {
            Read::Value2 => val_of(&value, false),
            Read::Value => {
                let date = matches!(value, Value::Number(_)) && self.is_date_format(si, at);
                val_of(&value, date)
            }
            Read::Formula | Read::Formula2 => {
                Val::Text(self.formula_text(si, at, what == Read::Formula, &value))
            }
            Read::FormulaR1C1 => {
                let text = self.formula_text(si, at, true, &value);
                match text.strip_prefix('=').map(gridcraft_formula::parse) {
                    Some(Ok(e)) => Val::Text(format!("={}", gridcraft_formula::print_r1c1(&e, at))),
                    _ => Val::Text(text),
                }
            }
            Read::Text => Val::Text(self.text(si, at, &value)),
            Read::NumberFormat => Val::Text(self.number_format_code(si, at)),
        }
    }

    fn formula_text(&mut self, si: usize, at: CellRef, legacy: bool, value: &Value) -> String {
        let formula = self
            .model
            .sheet(si)
            .and_then(|s| s.cell(at))
            .and_then(|c| c.formula.clone());
        let Some(f) = formula else {
            // A cell without a formula: its value as it would be typed.
            let own = self
                .model
                .sheet(si)
                .and_then(|s| s.cell(at))
                .map(|c| c.value.clone())
                .unwrap_or_default();
            return match own {
                Value::Number(n) if self.is_date_format(si, at) => self.display(
                    &Value::Number(n),
                    if n.fract() == 0.0 {
                        DATE_FORMAT
                    } else {
                        DATE_TIME_FORMAT
                    },
                ),
                Value::Empty => String::new(),
                _ => {
                    let _ = value;
                    own.display()
                }
            };
        };
        if legacy && f.array.is_none() {
            let shown = match save_as(&self.model, si, at, &f, FUNCTIONS, &mut self.names) {
                SaveAs::Legacy(e) => gridcraft_formula::print(&e),
                _ => f.text.clone(),
            };
            return format!("={shown}");
        }
        format!("={}", f.text)
    }

    fn display(&mut self, v: &Value, code: &str) -> String {
        let sys = self.model.date_system;
        let code = if code == "General" {
            code
        } else {
            local_format(code)
        };
        gridcraft_numfmt::format_value(v, self.format(code), sys).text
    }

    fn text(&mut self, si: usize, at: CellRef, value: &Value) -> String {
        self.formatted(si, at, value).text
    }

    /// What the cell shows and the color its number format gives it (`[Red]`).
    fn formatted(&mut self, si: usize, at: CellRef, value: &Value) -> gridcraft_numfmt::Formatted {
        let code = self.shown_format_code(si, at);
        match value {
            Value::Empty => gridcraft_numfmt::Formatted::default(),
            Value::Error(CellError::Circ) => gridcraft_numfmt::Formatted {
                text: "0".into(),
                ..Default::default()
            },
            Value::Number(n) if code == "General" && general_small(*n).is_some() => {
                gridcraft_numfmt::Formatted {
                    text: general_small(*n).unwrap_or_default(),
                    numeric: true,
                    ..Default::default()
                }
            }
            _ => {
                let sys = self.model.date_system;
                gridcraft_numfmt::format_value(value, self.format(&code), sys)
            }
        }
    }

    /// The color the number format of a cell gives what it shows (`[Red]`), as RRGGBB.
    pub fn format_color(&mut self, si: usize, at: CellRef) -> Option<String> {
        let value = self.value(si, at);
        let [r, g, b] = self.formatted(si, at, &value).color?.rgb()?;
        Some(format!("{r:02X}{g:02X}{b:02X}"))
    }

    /// Notes that a cell changed.
    fn touch(&mut self, si: usize, at: CellRef) {
        self.pending.push((si, at));
    }

    /// Writes values into the cells of `r`, as `Range.Value = ...` does: `rows` is laid over the
    /// range as Excel lays an array over it (a single value fills every cell, a single row every
    /// row; cells beyond the array get `#N/A`). Text is read as Excel reads what is typed into a
    /// cell (`"123"` is a number, `"=A1*2"` a formula) unless the cell is formatted as text.
    pub fn set_values(&mut self, si: usize, r: RangeRef, rows: &[Vec<Val>]) -> Result<()> {
        self.check_sheet(si)?;
        let array = to_array(rows);
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let v =
                    array.get_broadcast((row - r.start.row) as usize, (col - r.start.col) as usize);
                let at = CellRef::new(row, col);
                match v {
                    Value::Text(s) => self.enter(si, at, &s, false)?,
                    _ => {
                        let date = rows_val(
                            rows,
                            (row - r.start.row) as usize,
                            (col - r.start.col) as usize,
                        )
                        .is_some_and(|v| matches!(v, Val::Date(_)));
                        self.put_value(si, at, v, date);
                    }
                }
            }
        }
        Ok(())
    }

    /// Writes formulas (or values typed as text) into the cells of `r`, as `Range.Formula = ...`
    /// (`dynamic`: `Range.Formula2`) does. A formula written into a range of several cells is
    /// written into each, with its relative references moved as Excel moves them.
    pub fn set_formulas(
        &mut self,
        si: usize,
        r: RangeRef,
        rows: &[Vec<Val>],
        dynamic: bool,
    ) -> Result<()> {
        self.set_formulas_as(si, r, rows, dynamic, false)
    }

    /// `Range.FormulaR1C1 = rows`: the references are relative to each cell.
    pub fn set_formulas_r1c1(&mut self, si: usize, r: RangeRef, rows: &[Vec<Val>]) -> Result<()> {
        self.set_formulas_as(si, r, rows, false, true)
    }

    fn set_formulas_as(
        &mut self,
        si: usize,
        r: RangeRef,
        rows: &[Vec<Val>],
        dynamic: bool,
        r1c1: bool,
    ) -> Result<()> {
        self.check_sheet(si)?;
        let single = !r1c1 && rows.len() == 1 && rows[0].len() == 1;
        let array = to_array(rows);
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let (dr, dc) = (row - r.start.row, col - r.start.col);
                let v = array.get_broadcast(dr as usize, dc as usize);
                let at = CellRef::new(row, col);
                match v {
                    Value::Text(s) if single && s.starts_with('=') && (dr, dc) != (0, 0) => {
                        let moved = match gridcraft_formula::parse(&s[1..]) {
                            Ok(e) => format!(
                                "={}",
                                gridcraft_formula::print(
                                    &gridcraft_formula::adjust::shift_relative(
                                        e, dr as i64, dc as i64
                                    )
                                )
                            ),
                            Err(_) => s.to_string(),
                        };
                        self.enter(si, at, &moved, dynamic)?
                    }
                    Value::Text(s) if r1c1 && s.starts_with('=') => {
                        self.enter(si, at, &r1c1_to_a1(&s, at)?, dynamic)?
                    }
                    Value::Text(s) => self.enter(si, at, &s, dynamic)?,
                    other => {
                        let date = rows_val(rows, dr as usize, dc as usize)
                            .is_some_and(|v| matches!(v, Val::Date(_)));
                        self.put_value(si, at, other, date)
                    }
                }
            }
        }
        Ok(())
    }

    /// Whether the engine has a worksheet function (`WorksheetFunction.Name`).
    pub fn is_function(name: &str) -> bool {
        gridcraft_calc::is_known_function(&name.to_ascii_uppercase())
            && !STAND_INS.contains(&name.to_ascii_uppercase().as_str())
    }

    /// Writes a value (not text to be read) into a cell.
    fn put_value(&mut self, si: usize, at: CellRef, v: Value, date: bool) {
        if date && let Value::Number(n) = v {
            let code = self.number_format_code(si, at);
            if code == "General" {
                let f = if n.fract() == 0.0 {
                    DATE_FORMAT
                } else if n < 1.0 {
                    TIME_FORMAT
                } else {
                    DATE_TIME_FORMAT
                };
                self.set_number_format_at(si, at, f);
            }
        }
        let Some(sh) = self.model.sheet_mut(si) else {
            return;
        };
        if v.is_empty() {
            sh.remove_cell(at);
        } else {
            sh.set_value(at, v);
        }
        self.touch(si, at);
    }

    /// Writes text into a cell as Excel reads what is typed into it.
    fn enter(&mut self, si: usize, at: CellRef, text: &str, dynamic: bool) -> Result<()> {
        let code = self.number_format_code(si, at);
        let as_text = code == "@";
        if !as_text && let Some(body) = text.strip_prefix('=').filter(|b| !b.trim().is_empty()) {
            return self.put_formula(si, at, body, dynamic);
        }
        if as_text {
            self.put_value(
                si,
                at,
                if text.is_empty() {
                    Value::Empty
                } else {
                    Value::text(text)
                },
                false,
            );
            return Ok(());
        }
        let parsed = gridcraft_core::parse::parse_input(text, self.model.date_system);
        if let Some(f) = parsed.format
            && code == "General"
        {
            let f = match f {
                "yyyy-mm-dd" | "m/d/yyyy" => DATE_FORMAT,
                "m/d/yyyy h:mm" => DATE_TIME_FORMAT,
                // A month and a day without a year, as a Japanese Excel formats it.
                "d-mmm" => "m\"月\"d\"日\"",
                other => other,
            };
            self.set_number_format_at(si, at, f);
        }
        self.put_value(si, at, parsed.value, false);
        Ok(())
    }

    fn put_formula(&mut self, si: usize, at: CellRef, body: &str, dynamic: bool) -> Result<()> {
        let body = fmla::from_file(body.trim());
        if let Err(e) = gridcraft_formula::parse(&body) {
            return Err(format!("数式 \"={body}\" を読めません（{}）", e.msg));
        }
        let mut formula = gm::Formula::new(&body);
        formula.dynamic = dynamic;
        if self.number_format_code(si, at) == "General"
            && let Some(code) = formula
                .expr
                .as_deref()
                .and_then(|e| self.implied_format(si, e))
        {
            self.set_number_format_at(si, at, &code);
        }
        let Some(sh) = self.model.sheet_mut(si) else {
            return Ok(());
        };
        let style = sh.style_id(at);
        sh.set_cell(
            at,
            gm::Cell {
                value: Value::Empty,
                formula: Some(Arc::new(formula)),
                style,
                py: None,
            },
        );
        if !dynamic {
            // `Range.Formula`: the formula as Excel without dynamic arrays reads it.
            apply_legacy(&mut self.model, vec![(si, vec![at])], FUNCTIONS);
        }
        if self.manual {
            // Excel works a formula out when it is entered, also in manual calculation; what
            // depends on it waits.
            let expr = self
                .model
                .sheet(si)
                .and_then(|s| s.cell(at))
                .and_then(|c| c.formula.as_ref())
                .and_then(|f| f.expr_arc());
            if let Some(e) = expr {
                let v = gridcraft_calc::recalc::evaluate_expr(&self.model, si, at, &e).scalar();
                if let Some(c) = self.model.sheet_mut(si).and_then(|s| s.cells.get_mut(at)) {
                    c.value = if v == Value::Error(CellError::Circ) {
                        Value::number(0.0)
                    } else {
                        v
                    };
                }
            }
        }
        self.touch(si, at);
        Ok(())
    }

    /// The number format Excel gives a formula typed into a cell formatted as General: a date
    /// format for `TODAY()`, `DATE(...)` and the like, and the format of the first cell the formula
    /// reads when that one shows a date (`=A1+1`).
    fn implied_format(&mut self, si: usize, e: &Expr) -> Option<String> {
        let mut found: Option<String> = None;
        let mut first_ref: Option<(usize, CellRef)> = None;
        e.walk(&mut |x| {
            if found.is_some() {
                return;
            }
            match x {
                Expr::Call(name, _) => match name.as_str() {
                    "TODAY" | "DATE" | "DATEVALUE" | "EDATE" | "EOMONTH" | "WORKDAY"
                    | "WORKDAY.INTL" => found = Some(DATE_FORMAT.into()),
                    "NOW" => found = Some(DATE_TIME_FORMAT.into()),
                    "TIME" | "TIMEVALUE" => found = Some(TIME_FORMAT.into()),
                    _ => {}
                },
                Expr::Ref(r) if first_ref.is_none() => {
                    let sheet = match &r.sheet {
                        SheetSel::Current => Some(si),
                        SheetSel::Named(n) => self.model.sheet_index(n),
                        SheetSel::Span(..) => None,
                    };
                    if let Some(sheet) = sheet {
                        first_ref = Some((sheet, r.range().start));
                    }
                }
                _ => {}
            }
        });
        if found.is_some() {
            return found;
        }
        if let Some((sheet, at)) = first_ref
            && self.is_date_format(sheet, at)
        {
            return Some(self.number_format_code(sheet, at));
        }
        self.inherited_format(si, e)
    }

    /// The number format a formula takes from the cells it adds up: the format of the first of
    /// them that has one. As Excel 2021 does (tools/standin-check): `=SUM(C2:C4)`, `=C2+1` and
    /// `=D2+C2` take the `#,##0.00` of column C; `=C2*D2` and `=ROUND(C3*D3,0)` stay General.
    fn inherited_format(&mut self, si: usize, e: &Expr) -> Option<String> {
        use gridcraft_formula::ast::{BinOp, UnOp};
        match e {
            Expr::Paren(x) | Expr::Unary(UnOp::Neg | UnOp::Plus, x) => self.inherited_format(si, x),
            Expr::Binary(BinOp::Add | BinOp::Sub, a, b) => self
                .inherited_format(si, a)
                .or_else(|| self.inherited_format(si, b)),
            Expr::Call(name, args)
                if matches!(name.as_str(), "SUM" | "AVERAGE" | "MIN" | "MAX" | "MEDIAN") =>
            {
                args.iter().find_map(|a| self.inherited_format(si, a))
            }
            Expr::Ref(r) => {
                let sheet = match &r.sheet {
                    SheetSel::Current => si,
                    SheetSel::Named(n) => self.model.sheet_index(n)?,
                    SheetSel::Span(..) => return None,
                };
                let code = self.number_format_code(sheet, r.range().start);
                (code != "General" && code != "@").then_some(code)
            }
            _ => None,
        }
    }

    /// `Range.ClearContents` (`formats`: `Range.Clear`, which also clears the number formats).
    pub fn clear(&mut self, si: usize, r: RangeRef, formats: bool) -> Result<()> {
        self.check_sheet(si)?;
        let cells: Vec<CellRef> = self
            .model
            .sheet(si)
            .map(|s| s.cells.iter_range(r).map(|(at, _)| at).collect())
            .unwrap_or_default();
        if let Some(sh) = self.model.sheet_mut(si) {
            for at in &cells {
                sh.remove_cell(*at);
            }
        }
        for at in cells {
            self.touch(si, at);
        }
        if formats {
            self.clear_formats(si, r)?;
        }
        Ok(())
    }

    /// The range from the first to the last cell with content (`UsedRange`), None when empty.
    pub fn used_range(&mut self, si: usize) -> Option<RangeRef> {
        self.flush();
        let sh = self.model.sheet(si)?;
        let mut r = sh.used_range();
        for at in sh.spill.keys() {
            r = Some(r.map_or(RangeRef::cell(*at), |x| x.union(&RangeRef::cell(*at))));
        }
        // Formatted cells count too, as in Excel.
        if let Ok(ws) = self.file.sheet(si) {
            for c in ws.cells() {
                let at = at_of(c);
                r = Some(r.map_or(RangeRef::cell(at), |x| x.union(&RangeRef::cell(at))));
            }
        }
        r
    }

    fn filled(&self, si: usize, at: CellRef) -> bool {
        self.model.sheet(si).is_some_and(|s| {
            s.cell(at)
                .is_some_and(|c| c.formula.is_some() || !c.value.is_empty())
                || s.spill.get(&at).is_some_and(|v| !v.is_empty())
        })
    }

    /// The bounds of the cells with content, for `End` (None when the sheet is empty).
    fn content_bounds(&self, si: usize) -> Option<RangeRef> {
        let sh = self.model.sheet(si)?;
        let mut r: Option<RangeRef> = None;
        for (at, c) in sh.cells.iter() {
            if c.formula.is_some() || !c.value.is_empty() {
                r = Some(r.map_or(RangeRef::cell(at), |x| x.union(&RangeRef::cell(at))));
            }
        }
        for at in sh.spill.keys() {
            r = Some(r.map_or(RangeRef::cell(*at), |x| x.union(&RangeRef::cell(*at))));
        }
        r
    }

    /// `Range.End`: where Ctrl+arrow goes from `at`.
    pub fn end(&mut self, si: usize, at: CellRef, dir: Dir) -> CellRef {
        self.flush();
        let (dr, dc): (i64, i64) = match dir {
            Dir::Up => (-1, 0),
            Dir::Down => (1, 0),
            Dir::Left => (0, -1),
            Dir::Right => (0, 1),
        };
        let edge = match dir {
            Dir::Up => CellRef::new(0, at.col),
            Dir::Down => CellRef::new(MAX_ROWS - 1, at.col),
            Dir::Left => CellRef::new(at.row, 0),
            Dir::Right => CellRef::new(at.row, MAX_COLS - 1),
        };
        let Some(next) = at.offset(dr, dc) else {
            return at;
        };
        let Some(bounds) = self.content_bounds(si) else {
            return edge;
        };
        if self.filled(si, at) && self.filled(si, next) {
            let mut cur = next;
            while let Some(n) = cur.offset(dr, dc) {
                if !self.filled(si, n) {
                    break;
                }
                cur = n;
            }
            return cur;
        }
        // The next cell with content: none beyond the bounds.
        let mut cur = next;
        loop {
            // Jump over the empty part outside the bounds.
            match dir {
                Dir::Up if cur.row > bounds.end.row => cur.row = bounds.end.row,
                Dir::Left if cur.col > bounds.end.col => cur.col = bounds.end.col,
                Dir::Down if cur.row < bounds.start.row => cur.row = bounds.start.row,
                Dir::Right if cur.col < bounds.start.col => cur.col = bounds.start.col,
                _ => {}
            }
            let past = match dir {
                Dir::Up => cur.row < bounds.start.row,
                Dir::Down => cur.row > bounds.end.row,
                Dir::Left => cur.col < bounds.start.col,
                Dir::Right => cur.col > bounds.end.col,
            };
            if past {
                return edge;
            }
            if self.filled(si, cur) {
                return cur;
            }
            match cur.offset(dr, dc) {
                Some(n) => cur = n,
                None => return cur,
            }
        }
    }

    /// `Range.CurrentRegion`: the block of cells with content around `at`, bounded by empty rows
    /// and columns.
    pub fn current_region(&mut self, si: usize, at: CellRef) -> RangeRef {
        self.flush();
        let mut r = RangeRef::cell(at);
        loop {
            let grown = RangeRef::new(r.start.offset_clamped(-1, -1), r.end.offset_clamped(1, 1));
            let mut next = r;
            let mut any = false;
            for p in grown.iter() {
                if !r.contains(p) && self.filled(si, p) {
                    next = next.union(&RangeRef::cell(p));
                    any = true;
                }
            }
            if !any || next == r {
                // A single empty cell is its own region.
                return r;
            }
            r = next;
        }
    }

    // ---- number formats

    fn set_number_format_at(&mut self, si: usize, at: CellRef, code: &str) {
        self.inherit_style(si, at);
        if let Ok(ws) = self.file.sheet_mut(si) {
            ws.style_mut(coord(at))
                .number_format_mut()
                .set_format_code(code);
        }
    }

    /// `Range.NumberFormat = code`. Whole rows and columns are formatted where cells are used.
    pub fn set_number_format(&mut self, si: usize, r: RangeRef, code: &str) -> Result<()> {
        self.check_sheet(si)?;
        let r = if r.count() > 1_000_000 {
            match self.used_range(si).and_then(|u| u.intersection(&r)) {
                Some(x) => x,
                None => return Ok(()),
            }
        } else {
            r
        };
        for at in r.iter() {
            self.set_number_format_at(si, at, code);
        }
        Ok(())
    }

    // ---- formulas worked out on the side

    /// `Application.Evaluate` / `Worksheet.Evaluate`: a formula worked out on sheet `si`.
    pub fn evaluate(&mut self, si: usize, formula: &str) -> Value {
        self.flush();
        let body = formula.trim();
        let body = fmla::from_file(body.strip_prefix('=').unwrap_or(body));
        match gridcraft_formula::parse(&body) {
            Ok(e) => gridcraft_calc::recalc::evaluate_expr(&self.model, si, CellRef::new(0, 0), &e),
            Err(_) => Value::Error(CellError::Name),
        }
    }

    // ---- saving

    /// Copies the cells into the file and writes it to `path` (`.xlsx` / `.xlsm`). A workbook
    /// without macros is written as `.xlsx` (umya-spreadsheet marks the file by its macros, and
    /// Excel refuses an `.xlsm` file marked as an `.xlsx` one). Returns the path written.
    pub fn save(&mut self, path: &Path) -> Result<std::path::PathBuf> {
        self.flush();
        self.store();
        // Excel calculates everything when it opens the file, unless the program left the
        // calculation manual: the values are then the ones it left.
        self.file
            .calculation_properties_mut()
            .set_full_calc_on_load(!self.manual);
        let mode = if self.manual {
            umya_spreadsheet::CalculateModeValues::Manual
        } else {
            umya_spreadsheet::CalculateModeValues::Auto
        };
        self.file.calculation_properties_mut().set_calc_mode(mode);
        let macro_ext = path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("xlsm"));
        let path = if macro_ext && !self.file.has_macros() {
            path.with_extension("xlsx")
        } else {
            path.to_path_buf()
        };
        let kept = self.widths_for_file();
        let written = umya_spreadsheet::writer::xlsx::write(&self.file, &path);
        self.widths_back(kept);
        written.map_err(|e| format!("{} に保存できません（{e}）", path.display()))?;
        Ok(path)
    }

    /// Makes the column widths what the file has, and returns what they were (`widths_back`).
    /// A width of the file is the `ColumnWidth` of Excel and the padding of a cell, and
    /// umya-spreadsheet gives every column it has touched a width of its own (8.38): written as
    /// it is, Excel would show the columns the program wrote to narrower than the others
    /// (7.75), and a `ColumnWidth` of 20 as 19.38.
    fn widths_for_file(&mut self) -> Vec<(f64, Vec<umya_spreadsheet::structs::Column>)> {
        let mut kept = Vec::new();
        for ws in self.file.sheet_collection_mut() {
            let default = ws.sheet_format_properties().default_column_width();
            let columns = ws.column_dimensions().to_vec();
            kept.push((default, columns));
            if default > 0.0 {
                ws.sheet_format_properties_mut()
                    .set_default_column_width(format::width_to_file(default));
            }
            let untouched = |c: &umya_spreadsheet::structs::Column| {
                default <= 0.0
                    && c.width() == format::DEFAULT_COLUMN_WIDTH
                    && !c.hidden()
                    && !c.best_fit()
                    && *c.style() == umya_spreadsheet::Style::default()
            };
            let columns = ws.column_dimensions_mut();
            columns.retain(|c| !untouched(c));
            for c in columns.iter_mut() {
                let w = format::width_to_file(c.width());
                c.set_width(w);
            }
        }
        kept
    }

    fn widths_back(&mut self, kept: Vec<(f64, Vec<umya_spreadsheet::structs::Column>)>) {
        for (ws, (default, columns)) in self.file.sheet_collection_mut().iter_mut().zip(kept) {
            if default > 0.0 {
                ws.sheet_format_properties_mut()
                    .set_default_column_width(default);
            }
            *ws.column_dimensions_mut() = columns;
        }
    }

    /// The bytes of the file as `save` would write them.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        self.flush();
        self.store();
        let mut buf = Vec::new();
        let kept = self.widths_for_file();
        let written = umya_spreadsheet::writer::xlsx::write_writer(&self.file, &mut buf);
        self.widths_back(kept);
        written.map_err(|e| format!("保存できません（{e}）"))?;
        Ok(buf)
    }

    /// Copies the cells of the engine's model into the file.
    fn store(&mut self) {
        // Names.
        let sheets = self.model.sheets.len();
        let mut global = Vec::new();
        for n in &self.model.names {
            let mut d = umya_spreadsheet::DefinedName::default();
            d.set_name(n.name.clone());
            d.set_address(fmla::text_to_file(&n.formula));
            d.set_hidden(n.hidden);
            match n.scope {
                Some(s) if s < sheets => d.set_local_sheet_id(s as u32),
                _ => {}
            }
            global.push(d);
        }
        // umya keeps the names of the file in two places; they all go into the workbook's list.
        let builtin: Vec<umya_spreadsheet::DefinedName> = self
            .file
            .defined_names()
            .iter()
            .filter(|d| d.name().to_ascii_lowercase().starts_with("_xlnm."))
            .cloned()
            .collect();
        for ws in self.file.sheet_collection_mut() {
            ws.defined_names_mut()
                .retain(|d| d.name().to_ascii_lowercase().starts_with("_xlnm."));
        }
        self.file
            .set_defined_names(builtin.into_iter().chain(global).collect::<Vec<_>>());

        for si in 0..sheets {
            let fresh: Vec<CellRef> = match (self.model.sheet(si), self.file.sheet(si)) {
                (Some(sh), Ok(ws)) => sh
                    .cells
                    .iter()
                    .map(|(at, _)| at)
                    .chain(sh.spill.keys().copied())
                    .filter(|at| ws.cell(coord(*at)).is_none())
                    .collect(),
                _ => Vec::new(),
            };
            for at in fresh {
                self.inherit_style(si, at);
            }
            self.store_filter(si);
            let Some(sh) = self.model.sheet(si) else {
                continue;
            };
            let mut writes: Vec<(CellRef, Option<FileFormula>, Value)> = Vec::new();
            for (at, c) in sh.cells.iter() {
                let formula = c.formula.as_ref().map(|f| {
                    if let Some(area) = f.array {
                        (fmla::to_file(f), Some(area), false)
                    } else {
                        match save_as(&self.model, si, at, f, FUNCTIONS, &mut self.names) {
                            SaveAs::Unchanged => (fmla::to_file(f), None, false),
                            SaveAs::Legacy(e) => {
                                (fmla::to_file(&gm::Formula::from_expr(e)), None, false)
                            }
                            SaveAs::Dynamic => (
                                fmla::to_file(f),
                                Some(
                                    sh.spill_ranges
                                        .get(&at)
                                        .copied()
                                        .unwrap_or(RangeRef::cell(at)),
                                ),
                                true,
                            ),
                        }
                    }
                });
                writes.push((at, formula, c.value.clone()));
            }
            let spilled: Vec<(CellRef, Value)> =
                sh.spill.iter().map(|(at, v)| (*at, v.clone())).collect();
            let present: HashSet<CellRef> = writes
                .iter()
                .map(|w| w.0)
                .chain(spilled.iter().map(|s| s.0))
                .collect();
            let Ok(ws) = self.file.sheet_mut(si) else {
                continue;
            };
            // Cells the program emptied.
            let gone: Vec<CellRef> = ws
                .cells()
                .iter()
                .filter(|c| !c.raw_value().is_empty() || c.is_formula())
                .map(|c| at_of(c))
                .filter(|at| !present.contains(at))
                .collect();
            for at in gone {
                ws.cell_mut(coord(at)).cell_value_mut().set_blank();
            }
            for (at, formula, value) in writes {
                let c = ws.cell_mut(coord(at));
                match formula {
                    Some((text, area, dynamic)) => {
                        c.cell_value_mut().set_blank();
                        if let Some(area) = area {
                            let mut f = CellFormula::default();
                            f.set_text(text);
                            f.set_formula_type(CellFormulaValues::Array);
                            f.set_reference(area.a1());
                            if dynamic {
                                f.set_dynamic_array(true);
                            }
                            c.cell_value_mut().set_formula_obj(f);
                        } else {
                            c.set_formula(text);
                        }
                        if value == Value::Error(CellError::Circ) {
                            c.set_formula_result_number(0.0);
                        } else {
                            set_result(c, &value);
                        }
                    }
                    None => set_plain(c, &value),
                }
            }
            for (at, v) in spilled {
                let c = ws.cell_mut(coord(at));
                set_plain(c, &v);
            }
        }
    }
}

impl Book {
    /// Writes the filter of a sheet (its range and what it filters by) into the file.
    fn store_filter(&mut self, si: usize) {
        let Some(filter) = self.model.sheet(si).map(|s| s.autofilter.clone()) else {
            return;
        };
        let Ok(ws) = self.file.sheet_mut(si) else {
            return;
        };
        let Some(f) = filter else {
            ws.remove_auto_filter();
            return;
        };
        ws.set_auto_filter(f.range.a1());
        let Some(out) = ws.auto_filter_mut() else {
            return;
        };
        for (offset, criterion) in &f.criteria {
            let mut column = umya_spreadsheet::FilterColumn::default();
            column.set_column_id(*offset);
            match criterion {
                gm::FilterCriterion::Values { values, blanks } => {
                    column.set_values(values.clone(), *blanks);
                }
                gm::FilterCriterion::Custom { a, b, and } => {
                    let mut tests = vec![(a.0.clone(), a.1.clone())];
                    tests.extend(b.iter().cloned());
                    column.set_custom(tests, *and);
                }
                gm::FilterCriterion::Top10 {
                    bottom,
                    percent,
                    count,
                } => {
                    column.set_top10(!*bottom, *percent, f64::from(*count));
                }
                // Kept in the file as it was read, when it was.
                _ => continue,
            }
            out.add_filter_column(column);
        }
    }
}

/// Rewrites the R1C1 references of a formula as A1 references for the cell `at`
/// (`=R[-1]C*2` in B3 is `=B2*2`). Text in quotes stays as it is.
fn r1c1_to_a1(formula: &str, at: CellRef) -> Result<String> {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let word = |c: char| c.is_alphanumeric() || c == '_' || c == '.' || c == '\\';
    // One part of a reference: `R`, `R2`, `R[-1]` (or the same with `C`). Returns its end.
    let part = |i: usize, letter: char| -> Option<usize> {
        if chars.get(i).map(|c| c.to_ascii_uppercase()) != Some(letter) {
            return None;
        }
        let mut j = i + 1;
        if chars.get(j) == Some(&'[') {
            j += 1;
            if chars.get(j) == Some(&'-') || chars.get(j) == Some(&'+') {
                j += 1;
            }
            let start = j;
            while chars.get(j).is_some_and(char::is_ascii_digit) {
                j += 1;
            }
            if j == start || chars.get(j) != Some(&']') {
                return None;
            }
            j += 1;
        } else {
            while chars.get(j).is_some_and(char::is_ascii_digit) {
                j += 1;
            }
        }
        Some(j)
    };
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            let start = i;
            i += 1;
            while i < chars.len() {
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        i += 2;
                        continue;
                    }
                    break;
                }
                i += 1;
            }
            i = (i + 1).min(chars.len());
            out.extend(&chars[start..i]);
            continue;
        }
        let boundary = i == 0 || !word(chars[i - 1]);
        if boundary && let Some(r_end) = part(i, 'R') {
            let c_end = part(r_end, 'C');
            let end = c_end.unwrap_or(r_end);
            let next = chars.get(end).copied();
            if !next.is_some_and(|n| word(n) || n == '(' || n == '!') {
                let text: String = chars[i..end].iter().collect();
                // A whole row (`R2`) or column (`C3`) needs both halves for the conversion.
                let full = if c_end.is_some() {
                    text.clone()
                } else {
                    format!("{text}C1")
                };
                if let Ok(Some(a1)) = gridcraft_calc::eval::r1c1_to_a1(&full, at) {
                    let a1 = if c_end.is_some() {
                        a1
                    } else {
                        a1.trim_start_matches(|ch: char| ch.is_ascii_alphabetic() || ch == '$')
                            .to_string()
                    };
                    out.push_str(&a1);
                    i = end;
                    continue;
                }
            }
        }
        if boundary && let Some(end) = part(i, 'C') {
            let next = chars.get(end).copied();
            if !next.is_some_and(|n| word(n) || n == '(' || n == '!') {
                let text: String = chars[i..end].iter().collect();
                if let Ok(Some(a1)) = gridcraft_calc::eval::r1c1_to_a1(&format!("R1{text}"), at) {
                    out.push_str(
                        a1.trim_end_matches(|ch: char| ch.is_ascii_digit())
                            .trim_end_matches('$'),
                    );
                    i = end;
                    continue;
                }
            }
        }
        out.push(c);
        i += 1;
    }
    Ok(out)
}

/// A formula as the file holds it: its text, the area of an array formula, whether it is a
/// dynamic-array formula.
type FileFormula = (String, Option<RangeRef>, bool);

fn rows_val(rows: &[Vec<Val>], r: usize, c: usize) -> Option<&Val> {
    let r = if rows.len() == 1 { 0 } else { r };
    let row = rows.get(r)?;
    row.get(if row.len() == 1 { 0 } else { c })
}

fn to_array(rows: &[Vec<Val>]) -> Array {
    let height = rows.len().max(1);
    let width = rows.iter().map(Vec::len).max().unwrap_or(1).max(1);
    let mut data = Vec::with_capacity(height * width);
    for r in 0..height {
        for c in 0..width {
            data.push(
                rows.get(r)
                    .and_then(|row| row.get(c))
                    .map(value_of_val)
                    .unwrap_or(Value::Error(CellError::NA)),
            );
        }
    }
    Array::new(height, width, data).unwrap_or_else(|| Array::scalar(Value::Empty))
}

mod cond;
mod edit;
mod format;
pub(crate) use format::width_of_px;
mod inspect;
mod paste;
pub mod snapshot;
mod sort;
pub(crate) mod tables;
pub mod view;

pub(crate) use edit::{adjusted, moved};
pub use gridcraft_model::FilterCriterion;
pub use inspect::{CellInfo, FormulaRegion, RuleInfo, SheetInfo, TableInfo};
pub use paste::{Paste, PasteOp, PasteWhat};
pub use sort::{SortKey, Special, split_criterion};
pub use view::{VIEW_COLS, VIEW_ROWS};

#[cfg(test)]
mod tests;
