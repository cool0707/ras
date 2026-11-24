//! Conversion of VBA to RAS, used by `ras form pull` to turn the code of a
//! form made in the VBE into a RAS code-behind (`forms/<Name>.ras`, spec §14.11.6),
//! and by the sync to take over the hand-written standard modules and document
//! modules of a workbook (`src/<Name>.ras`, spec §14.7.1).
//!
//! The conversion is syntactic: VBA is parsed into a small tree and printed as RAS.
//! Host names are written as RAS spells them (`MsgBox` → `msgBox`, `.Text` → `.text`),
//! controls of the form become `this.<control>`, event procedures (`btnOK_Click`,
//! `Worksheet_Change`) become `@on("btnOK.Click")` / `@on("Sheet1.Change")` handlers.
//! What has no RAS counterpart (`GoTo`, `ReDim`, `Resume Next` in a handler,
//! user-defined types …) is kept as a `// TODO(VBA):`
//! comment with the original line and reported, so that nothing is lost silently.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use crate::config::{Config, FormDecl};
use crate::model::{Program, Sym};
use crate::names;

/// Result of a conversion.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Converted {
    /// The RAS source (LF line ends).
    pub text: String,
    /// What could not be converted: (1-based line in `text`, message).
    pub todos: Vec<(usize, String)>,
}

/// Converts the code of a form (its code pane, without `Attribute` lines) to a RAS code-behind.
pub fn convert_form(vba: &str, form: &FormDecl) -> Converted {
    convert(vba, &Target::Form(form), &[])
}

/// What a module of VBA becomes in RAS.
#[derive(Clone, Copy, Debug)]
pub enum Target<'a> {
    /// The code-behind of a form (`forms/<Name>.ras`).
    Form(&'a FormDecl),
    /// A module of `src/`, from a standard module.
    Standard,
    /// The file of a sheet or of `ThisWorkbook` (`objects/<CodeName>.ras`, spec §12.2), from
    /// the hand-written code of that document module: the sheet with this code name, or
    /// `ThisWorkbook` (`workbook`). Its event procedures become `@on` handlers
    /// (`Worksheet.Change`; `CommandButton1.Click` for a control on the sheet); `Me` and its own
    /// members used without a qualifier (`Range`, `Me.Cells`, a control) become `this.…`.
    Document {
        code_name: &'a str,
        workbook: bool,
        controls: &'a [crate::config::FormControl],
    },
}

/// Another module converted with this one, whose public names the converted code may use:
/// they are imported from the RAS module it becomes.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Peer {
    /// Name of the VBA module (for a document module, its code name).
    pub vba_name: String,
    /// Key of the RAS module it becomes: its path in `src/` without `.ras` (`Module1`).
    pub key: String,
    /// A sheet or `ThisWorkbook`: its public procedures stay members of the sheet
    /// (`Sheet1.Foo`), which every module can refer to without an import (spec §12.2).
    pub document: bool,
    /// Its public names: VBA name and RAS name, and whether it is a function.
    pub names: Vec<PublicName>,
}

/// A public name of a [`Peer`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PublicName {
    pub vba: String,
    pub ras: String,
    pub is_function: bool,
}

/// Converts a module of VBA (its code pane, without `Attribute` lines). Names that the
/// `peers` make public are imported from them.
pub fn convert(vba: &str, target: &Target, peers: &[Peer]) -> Converted {
    if vba.trim().is_empty() {
        return Converted::default();
    }
    let host = HostIndex::new();
    let lines = logical_lines(vba);
    let mut parser = Parser::new(&lines);
    let items = parser.module();
    let mut em = Emitter::new(&host, *target, peers, &lines);
    em.with_events = parser.with_events;
    em.module(&items);
    em.finish()
}

/// The public names of a module of VBA, as [`convert`] names them in RAS (for a [`Peer`]).
/// Event procedures of a document module are not public names.
pub fn public_names(vba: &str, target: &Target) -> Vec<PublicName> {
    let host = HostIndex::new();
    let lines = logical_lines(vba);
    let mut parser = Parser::new(&lines);
    let items = parser.module();
    let mut em = Emitter::new(&host, *target, &[], &lines);
    em.with_events = parser.with_events;
    let mut out = Vec::new();
    for item in &items {
        match item {
            Item::Proc(p) if em.event_of(&p.name).is_none() && !is_private(p.vis.as_deref()) => {
                let mut ras = Emitter::user_ident(&p.name);
                if matches!(p.kind.as_str(), "Get" | "Let" | "Set") {
                    ras = format!("{}{ras}", p.kind.to_ascii_lowercase());
                }
                if !out.iter().any(|n: &PublicName| n.ras == ras) {
                    out.push(PublicName {
                        vba: p.name.clone(),
                        ras,
                        is_function: p.kind == "Function" || p.kind == "Get",
                    });
                }
            }
            Item::Stmt(
                s @ Stmt {
                    kind: StmtKind::Dim { vars, .. } | StmtKind::Const(vars),
                    ..
                },
            ) if is_public_decl(&lines[s.line]) => {
                for v in vars {
                    out.push(PublicName {
                        vba: v.name.clone(),
                        ras: Emitter::user_ident(&v.name),
                        is_function: false,
                    });
                }
            }
            Item::Enum {
                name, public: true, ..
            } => out.push(PublicName {
                vba: name.clone(),
                ras: name.clone(),
                is_function: false,
            }),
            _ => {}
        }
    }
    out
}

fn is_private(vis: Option<&str>) -> bool {
    vis.is_some_and(|v| v.eq_ignore_ascii_case("private") || v.eq_ignore_ascii_case("friend"))
}

/// A module-level `Public` / `Global` declaration.
fn is_public_decl(line: &Line) -> bool {
    line.toks
        .first()
        .is_some_and(|t| t.is_kw("Public") || t.is_kw("Global"))
}

// ---------------------------------------------------------------------------------------------
// Host names

/// The names of the bundled type definitions, from an analysed empty program.
struct HostIndex {
    /// Lower-cased name → (RAS name, is a function, number of required parameters).
    globals: HashMap<String, (String, bool, usize)>,
    /// Lower-cased class name → qualified RAS type (`Excel.Range`, `MSForms.TextBox`).
    classes: HashMap<String, String>,
    /// Members of `Excel.Worksheet` and `Excel.Workbook`: lower-cased VBA name → (RAS name, is a method).
    worksheet: HashMap<String, (String, bool)>,
    workbook: HashMap<String, (String, bool)>,
}

impl HostIndex {
    fn new() -> HostIndex {
        let prog: Program = crate::analyze(Vec::new(), Config::default());
        let mut globals = HashMap::new();
        for (name, sym) in &prog.globals {
            let entry = match sym {
                Sym::Fn(f) => (name.clone(), true, prog.fns[*f].required_params()),
                _ => (name.clone(), false, 0),
            };
            globals.insert(name.to_lowercase(), entry);
        }
        let mut classes = HashMap::new();
        // Earlier namespaces win: `MSForms.Label`, not a later `Label`.
        for ns in ["MSForms", "Excel", "VBA", "Office", "Scripting", "VBScript"] {
            let Some(members) = prog.namespaces.get(ns) else {
                continue;
            };
            for (name, sym) in members {
                if matches!(sym, Sym::Class(_) | Sym::Enum(_)) {
                    classes
                        .entry(name.to_lowercase())
                        .or_insert_with(|| format!("{ns}.{name}"));
                }
            }
        }
        let members = |class: &str| {
            let mut out = HashMap::new();
            let Some(Sym::Class(id)) = prog.namespaces.get("Excel").and_then(|ns| ns.get(class))
            else {
                return out;
            };
            let def = &prog.classes[*id];
            for f in &def.fields {
                out.insert(f.vba_name().to_lowercase(), (f.name.clone(), false));
            }
            for &m in &def.methods {
                let f = &prog.fns[m];
                out.entry(f.vba_name().to_lowercase())
                    .or_insert_with(|| (f.name.clone(), true));
            }
            out
        };
        let (worksheet, workbook) = (members("Worksheet"), members("Workbook"));
        HostIndex {
            globals,
            classes,
            worksheet,
            workbook,
        }
    }
}

// ---------------------------------------------------------------------------------------------
// Lexer

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    /// An identifier or keyword (type-declaration character removed), and that character.
    Word(String, Option<char>),
    /// `[name]`: an identifier that may be a keyword.
    Bracketed(String),
    /// A number already written as RAS (`123`, `1.5`, `0xFF`).
    Num(String, bool),
    Str(String),
    Date(String),
    Op(&'static str),
    /// Statement separator `:`.
    Colon,
}

impl Tok {
    fn is_kw(&self, kw: &str) -> bool {
        matches!(self, Tok::Word(w, None) if w.eq_ignore_ascii_case(kw))
    }

    fn is_op(&self, op: &str) -> bool {
        matches!(self, Tok::Op(o) if *o == op)
    }
}

/// One logical line: continuation lines joined, its comment split off.
#[derive(Clone, Debug)]
struct Line {
    toks: Vec<Tok>,
    comment: Option<String>,
    /// Original text (all physical lines).
    src: String,
    /// A label at the start (`Handler:` or a line number).
    label: Option<String>,
    blank_before: bool,
    /// The line could not be read.
    bad: bool,
}

const OPS: &[&str] = &[
    ":=", "<>", "<=", ">=", "=", "<", ">", "+", "-", "*", "/", "\\", "^", "&", "(", ")", ",", ".",
    "!", ";",
];

/// Words after which a `#` is a file number (`Print #1`, `As #f`), not a date literal.
const FILE_NUMBER_AFTER: &[&str] = &[
    "print", "write", "input", "get", "put", "close", "seek", "lock", "unlock", "width", "as",
    "eof", "lof", "loc", "fileattr",
];

fn lex(code: &str) -> Result<Vec<Tok>, ()> {
    let chars: Vec<char> = code.chars().collect();
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() {
            i += 1;
            continue;
        }
        if c == '"' {
            let mut s = String::new();
            i += 1;
            loop {
                match chars.get(i) {
                    None => return Err(()),
                    Some('"') if chars.get(i + 1) == Some(&'"') => {
                        s.push('"');
                        i += 2;
                    }
                    Some('"') => {
                        i += 1;
                        break;
                    }
                    Some(ch) => {
                        s.push(*ch);
                        i += 1;
                    }
                }
            }
            out.push(Tok::Str(s));
            continue;
        }
        if c == '#' {
            // A file number: `Print #1, …`, `Open … As #f`, `Close #1, #2`, `Input(5, #1)`.
            let file_number = match out.last() {
                Some(Tok::Word(w, None)) => {
                    FILE_NUMBER_AFTER.contains(&w.to_ascii_lowercase().as_str())
                }
                Some(Tok::Op(",") | Tok::Op("(")) => true,
                _ => false,
            };
            // Date literal `#1/2/2020#` (a `#` elsewhere is a file number or a suffix).
            if let Some(end) = chars[i + 1..].iter().position(|&ch| ch == '#') {
                let inner: String = chars[i + 1..i + 1 + end].iter().collect();
                if !inner.is_empty()
                    && inner
                        .chars()
                        .all(|ch| ch.is_ascii_alphanumeric() || " /-:.,".contains(ch))
                    && (!file_number || inner.contains(['/', '-', ':']) || {
                        // `#Jan 1, 2020#`: a month written as a name.
                        let lower = inner.to_ascii_lowercase();
                        [
                            "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct",
                            "nov", "dec",
                        ]
                        .iter()
                        .any(|m| {
                            lower
                                .split(|c: char| !c.is_ascii_alphabetic())
                                .any(|w| w.starts_with(m))
                        })
                    })
                {
                    out.push(Tok::Date(inner));
                    i += end + 2;
                    continue;
                }
            }
            if file_number {
                out.push(Tok::Op("#"));
                i += 1;
                continue;
            }
            return Err(());
        }
        if c == '&'
            && matches!(chars.get(i + 1), Some('H' | 'h' | 'O' | 'o'))
            && chars.get(i + 2).is_some_and(|ch| ch.is_ascii_hexdigit())
        {
            let octal = matches!(chars[i + 1], 'O' | 'o');
            let mut j = i + 2;
            while j < chars.len() && chars[j].is_ascii_hexdigit() {
                j += 1;
            }
            let digits: String = chars[i + 2..j].iter().collect();
            if matches!(chars.get(j), Some('&' | '%' | '^')) {
                j += 1;
            }
            let text = if octal {
                i64::from_str_radix(&digits, 8)
                    .map(|v| v.to_string())
                    .map_err(|_| ())?
            } else {
                format!("0x{}", digits.to_uppercase())
            };
            out.push(Tok::Num(text, false));
            i = j;
            continue;
        }
        if c.is_ascii_digit()
            || (c == '.' && chars.get(i + 1).is_some_and(|ch| ch.is_ascii_digit()))
        {
            let mut j = i;
            while j < chars.len() && (chars[j].is_ascii_digit() || chars[j] == '.') {
                j += 1;
            }
            if j < chars.len()
                && matches!(chars[j], 'E' | 'e' | 'D' | 'd')
                && chars
                    .get(j + 1)
                    .is_some_and(|ch| ch.is_ascii_digit() || *ch == '-' || *ch == '+')
            {
                j += 2;
                while j < chars.len() && chars[j].is_ascii_digit() {
                    j += 1;
                }
            }
            let mut text: String = chars[i..j]
                .iter()
                .collect::<String>()
                .replace(['D', 'd'], "e");
            let mut float = text.contains(['.', 'e', 'E']);
            match chars.get(j) {
                Some('#') | Some('!') | Some('@') => {
                    float = true;
                    j += 1;
                }
                Some('&') | Some('%') | Some('^') => j += 1,
                _ => {}
            }
            if text.starts_with('.') {
                text.insert(0, '0');
            }
            if float && !text.contains(['.', 'e', 'E']) {
                text.push_str(".0");
            }
            out.push(Tok::Num(text, float));
            i = j;
            continue;
        }
        if c.is_alphabetic() || c == '_' {
            let mut j = i;
            while j < chars.len() && (chars[j].is_alphanumeric() || chars[j] == '_') {
                j += 1;
            }
            let word: String = chars[i..j].iter().collect();
            let suffix = match chars.get(j) {
                // `x&` is a suffix only when not followed by an operand (`a &b` is a concatenation).
                Some(&ch @ ('$' | '%' | '&' | '!' | '#' | '@'))
                    if !chars
                        .get(j + 1)
                        .is_some_and(|n| n.is_alphanumeric() || *n == '"' || *n == '_') =>
                {
                    j += 1;
                    Some(ch)
                }
                _ => None,
            };
            out.push(Tok::Word(word, suffix));
            i = j;
            continue;
        }
        if c == '[' {
            let end = chars[i..].iter().position(|&ch| ch == ']').ok_or(())?;
            out.push(Tok::Bracketed(chars[i + 1..i + end].iter().collect()));
            i += end + 1;
            continue;
        }
        if c == ':' && chars.get(i + 1) != Some(&'=') {
            out.push(Tok::Colon);
            i += 1;
            continue;
        }
        let rest: String = chars[i..chars.len().min(i + 2)].iter().collect();
        let Some(op) = OPS.iter().find(|op| rest.starts_with(**op)) else {
            return Err(());
        };
        out.push(Tok::Op(op));
        i += op.chars().count();
    }
    Ok(out)
}

/// Splits a physical line into code and comment (`'` or `Rem` outside strings).
fn split_comment(line: &str) -> (String, Option<String>) {
    let mut in_str = false;
    for (i, ch) in line.char_indices() {
        match ch {
            '"' => in_str = !in_str,
            '\'' if !in_str => return (line[..i].to_string(), Some(line[i + 1..].to_string())),
            _ => {}
        }
    }
    let trimmed = line.trim_start();
    if trimmed.len() >= 3
        && trimmed[..3].eq_ignore_ascii_case("rem")
        && trimmed[3..].chars().next().is_none_or(char::is_whitespace)
    {
        return (String::new(), Some(trimmed[3..].trim_start().to_string()));
    }
    (line.to_string(), None)
}

fn logical_lines(code: &str) -> Vec<Line> {
    let text = code.replace("\r\n", "\n").replace('\r', "\n");
    let physical: Vec<&str> = text.lines().collect();
    let mut out = Vec::new();
    let mut i = 0;
    let mut blank = false;
    while i < physical.len() {
        let mut src = physical[i].to_string();
        let (mut code, mut comment) = split_comment(physical[i]);
        // ` _` at the end continues the statement on the next line.
        while comment.is_none()
            && (code.trim_end().ends_with(" _") || code.trim_end() == "_")
            && i + 1 < physical.len()
        {
            i += 1;
            src.push('\n');
            src.push_str(physical[i]);
            let trimmed = code.trim_end();
            code = format!("{} ", &trimmed[..trimmed.len() - 1]);
            let (next, c) = split_comment(physical[i]);
            code.push_str(next.trim_start());
            comment = c;
        }
        i += 1;
        if code.trim().is_empty() && comment.is_none() {
            blank = true;
            continue;
        }
        // A label: `Name:` or a line number at the start.
        let mut label = None;
        let trimmed = code.trim_start();
        let word: String = trimmed
            .chars()
            .take_while(|c| c.is_alphanumeric() || *c == '_')
            .collect();
        if !word.is_empty() {
            let after = &trimmed[word.len()..];
            if word.chars().all(|c| c.is_ascii_digit()) {
                label = Some(word.clone());
                code = after.to_string();
            } else if after.starts_with(':') && !after.starts_with(":=") && !is_keyword(&word) {
                label = Some(word.clone());
                code = after[1..].to_string();
            }
        }
        let (toks, bad) = match lex(&code) {
            Ok(t) => (t, false),
            Err(()) => (Vec::new(), true),
        };
        out.push(Line {
            toks,
            comment,
            src,
            label,
            blank_before: blank,
            bad,
        });
        blank = false;
    }
    out
}

const KEYWORDS: &[&str] = &[
    "and",
    "as",
    "byref",
    "byval",
    "call",
    "case",
    "const",
    "declare",
    "dim",
    "do",
    "each",
    "else",
    "elseif",
    "end",
    "enum",
    "eqv",
    "erase",
    "error",
    "event",
    "exit",
    "false",
    "for",
    "friend",
    "function",
    "get",
    "global",
    "gosub",
    "goto",
    "if",
    "imp",
    "implements",
    "in",
    "is",
    "let",
    "like",
    "loop",
    "me",
    "mod",
    "new",
    "next",
    "not",
    "nothing",
    "on",
    "option",
    "optional",
    "or",
    "paramarray",
    "preserve",
    "private",
    "property",
    "public",
    "raiseevent",
    "redim",
    "resume",
    "return",
    "select",
    "set",
    "static",
    "step",
    "stop",
    "sub",
    "then",
    "to",
    "true",
    "type",
    "typeof",
    "until",
    "wend",
    "while",
    "with",
    "withevents",
    "xor",
];

fn is_keyword(word: &str) -> bool {
    KEYWORDS.contains(&word.to_ascii_lowercase().as_str())
}

// ---------------------------------------------------------------------------------------------
// Syntax tree

#[derive(Clone, Debug)]
enum Expr {
    Num(String, bool),
    Str(String),
    Date(String),
    Bool(bool),
    Nothing,
    Empty,
    Null,
    Me,
    Ident(String, Option<char>),
    /// `obj.name`; `None` is the object of the enclosing `With`.
    Member(Option<Box<Expr>>, String),
    Bang,
    Call(Box<Expr>, Vec<Arg>),
    Unary(&'static str, Box<Expr>),
    Binary(String, Box<Expr>, Box<Expr>),
    Paren(Box<Expr>),
    New(String),
    TypeOf,
}

#[derive(Clone, Debug)]
struct Arg {
    name: Option<String>,
    value: Option<Expr>,
}

#[derive(Clone, Debug)]
struct TypeName {
    name: String,
    /// `As New T`
    new: bool,
}

#[derive(Clone, Debug)]
struct VarDecl {
    name: String,
    suffix: Option<char>,
    ty: Option<TypeName>,
    /// Declared with parentheses: an array; the bounds (empty for a dynamic array).
    dims: Option<Vec<Expr>>,
    /// `Const` value.
    value: Option<Expr>,
}

#[derive(Clone, Debug)]
enum CaseItem {
    Value(Expr),
    Range(Expr, Expr),
    Is(String, Expr),
}

#[derive(Clone, Debug)]
enum StmtKind {
    Dim {
        vars: Vec<VarDecl>,
        is_static: bool,
    },
    Const(Vec<VarDecl>),
    Assign {
        target: Expr,
        value: Expr,
    },
    Call(Expr),
    If {
        arms: Vec<(Expr, Vec<Stmt>)>,
        els: Option<Vec<Stmt>>,
    },
    Select {
        subject: Expr,
        cases: Vec<(Vec<CaseItem>, Vec<Stmt>)>,
        els: Option<Vec<Stmt>>,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    ForEach {
        var: String,
        coll: Expr,
        body: Vec<Stmt>,
    },
    /// `Do [While|Until c] … Loop [While|Until c]`; the bool is "Until".
    Do {
        pre: Option<(bool, Expr)>,
        post: Option<(bool, Expr)>,
        body: Vec<Stmt>,
    },
    With {
        obj: Expr,
        body: Vec<Stmt>,
    },
    Exit(String),
    OnErrorResumeNext,
    OnErrorGoTo0,
    OnErrorGoTo(String),
    Label(String),
    /// A statement of the files opened by number (`Open`, `Print #` …): the RAS function and
    /// its arguments; with `result`, `result = func(args)` (`Line Input #`).
    File {
        func: &'static str,
        args: Vec<FileArg>,
        result: Option<Expr>,
    },
    /// A statement without RAS counterpart: the reason.
    Todo(String),
    /// Only a comment.
    Nothing,
}

#[derive(Clone, Debug)]
enum FileArg {
    Value(Expr),
    /// A variable the statement assigns (`Input #1, a`).
    Target(Expr),
    /// A word written as is (`forOutput`).
    Word(&'static str),
}

#[derive(Clone, Debug)]
struct Stmt {
    kind: StmtKind,
    /// Index of the line the statement starts on.
    line: usize,
    /// Comment at the end of the line (the last statement of a line carries it).
    comment: Option<String>,
    blank_before: bool,
}

#[derive(Clone, Debug)]
struct Param {
    name: String,
    suffix: Option<char>,
    byval: Option<bool>,
    optional: bool,
    paramarray: bool,
    ty: Option<TypeName>,
    array: bool,
    default: Option<Expr>,
}

#[derive(Clone, Debug)]
struct Proc {
    /// `Sub`, `Function`, `Get`, `Let`, `Set`.
    kind: String,
    /// `Public`, `Private`, `Friend` or none.
    vis: Option<String>,
    is_static: bool,
    name: String,
    suffix: Option<char>,
    params: Vec<Param>,
    ret: Option<TypeName>,
    ret_array: bool,
    body: Vec<Stmt>,
    line: usize,
}

#[derive(Clone, Debug)]
enum Item {
    Stmt(Stmt),
    Proc(Proc),
    Enum {
        name: String,
        public: bool,
        members: Vec<(String, Option<Expr>, Option<String>)>,
        line: usize,
    },
    /// Lines kept as a TODO (from, to inclusive), with the reason.
    Todo {
        from: usize,
        to: usize,
        reason: String,
    },
}

// ---------------------------------------------------------------------------------------------
// Parser

struct Parser<'a> {
    lines: &'a [Line],
    /// Current line and token.
    li: usize,
    ti: usize,
    /// `Next i, j` closes more than one loop.
    pending_next: usize,
    /// Variables declared `WithEvents`.
    with_events: Vec<String>,
}

type PResult<T> = Result<T, String>;

impl<'a> Parser<'a> {
    fn new(lines: &'a [Line]) -> Parser<'a> {
        Parser {
            lines,
            li: 0,
            ti: 0,
            pending_next: 0,
            with_events: Vec::new(),
        }
    }

    fn at_end(&self) -> bool {
        self.li >= self.lines.len()
    }

    fn toks(&self) -> &'a [Tok] {
        self.lines.get(self.li).map(|l| &l.toks[..]).unwrap_or(&[])
    }

    fn peek(&self) -> Option<&'a Tok> {
        self.toks().get(self.ti)
    }

    fn peek_at(&self, n: usize) -> Option<&'a Tok> {
        self.toks().get(self.ti + n)
    }

    fn bump(&mut self) -> Option<&'a Tok> {
        let t = self.toks().get(self.ti);
        if t.is_some() {
            self.ti += 1;
        }
        t
    }

    fn eat_kw(&mut self, kw: &str) -> bool {
        if self.peek().is_some_and(|t| t.is_kw(kw)) {
            self.ti += 1;
            true
        } else {
            false
        }
    }

    fn eat_op(&mut self, op: &str) -> bool {
        if self.peek().is_some_and(|t| t.is_op(op)) {
            self.ti += 1;
            true
        } else {
            false
        }
    }

    fn expect_op(&mut self, op: &str) -> PResult<()> {
        if self.eat_op(op) {
            Ok(())
        } else {
            Err(format!("`{op}` がありません"))
        }
    }

    /// End of the current statement: end of line, `:` or (in a single-line `If`) `Else`.
    fn at_stmt_end(&self) -> bool {
        matches!(self.peek(), None | Some(Tok::Colon))
    }

    fn next_line(&mut self) {
        self.li += 1;
        self.ti = 0;
    }

    /// Moves past the end of the current statement.
    fn finish_stmt(&mut self) {
        if self.peek() == Some(&Tok::Colon) {
            self.ti += 1;
            if self.peek().is_none() {
                self.next_line();
            }
        } else {
            self.next_line();
        }
    }

    /// Whether the statement at the cursor starts with these keywords.
    fn starts_with(&self, kws: &[&str]) -> bool {
        kws.iter()
            .enumerate()
            .all(|(i, kw)| self.peek_at(i).is_some_and(|t| t.is_kw(kw)))
    }

    fn ident(&mut self) -> PResult<(String, Option<char>)> {
        match self.bump() {
            Some(Tok::Word(w, s)) => Ok((w.clone(), *s)),
            Some(Tok::Bracketed(w)) => Ok((w.clone(), None)),
            _ => Err("名前がありません".into()),
        }
    }

    fn module(&mut self) -> Vec<Item> {
        let mut items = Vec::new();
        while !self.at_end() {
            let start = self.li;
            if self.ti == 0 && self.lines[self.li].bad {
                items.push(Item::Todo {
                    from: start,
                    to: start,
                    reason: "読み取れない行です".into(),
                });
                self.next_line();
                continue;
            }
            if self.toks().is_empty() {
                items.push(Item::Stmt(self.comment_stmt()));
                continue;
            }
            let mut j = 0;
            let mut vis = None;
            while let Some(Tok::Word(w, None)) = self.peek_at(j) {
                let lw = w.to_ascii_lowercase();
                if ["public", "private", "friend", "global", "static"].contains(&lw.as_str()) {
                    if lw != "static" {
                        vis = Some(w.clone());
                    }
                    j += 1;
                } else {
                    break;
                }
            }
            let head = self.peek_at(j).cloned();
            let kw = |k: &str| head.as_ref().is_some_and(|t| t.is_kw(k));
            if kw("Sub") || kw("Function") || kw("Property") {
                match self.procedure() {
                    Ok(p) => items.push(Item::Proc(p)),
                    Err(reason) => {
                        let to = self.skip_to_end_of_proc(start);
                        items.push(Item::Todo {
                            from: start,
                            to,
                            reason,
                        });
                    }
                }
                continue;
            }
            if kw("Enum") {
                match self.enumeration(vis.as_deref()) {
                    Ok(e) => items.push(e),
                    Err(reason) => {
                        let to = self.skip_to("End", "Enum", start);
                        items.push(Item::Todo {
                            from: start,
                            to,
                            reason,
                        });
                    }
                }
                continue;
            }
            if kw("Type") {
                let to = self.skip_to("End", "Type", start);
                items.push(Item::Todo { from: start, to, reason: "ユーザー定義型（Type）には対応していません。src/ のモジュールに struct として宣言してください".into() });
                continue;
            }
            if kw("Option") {
                let base = self.starts_with(&["Option", "Base"]);
                if base {
                    items.push(Item::Todo {
                        from: start,
                        to: start,
                        reason: "Option Base には対応していません（RAS の配列は 0 始まりです）"
                            .into(),
                    });
                }
                self.next_line();
                continue;
            }
            if kw("Declare") {
                items.push(Item::Todo { from: start, to: start, reason: "Declare には対応していません。src/ のモジュールに `declare module` で宣言してください".into() });
                self.next_line();
                continue;
            }
            if kw("Implements") || kw("Event") {
                items.push(Item::Todo {
                    from: start,
                    to: start,
                    reason: "Implements / Event には対応していません".into(),
                });
                self.next_line();
                continue;
            }
            // Module-level declarations and anything else.
            let stmt = self.statement();
            items.push(Item::Stmt(stmt));
        }
        items
    }

    fn comment_stmt(&mut self) -> Stmt {
        let line = &self.lines[self.li];
        let s = Stmt {
            kind: StmtKind::Nothing,
            line: self.li,
            comment: line.comment.clone(),
            blank_before: line.blank_before,
        };
        self.next_line();
        s
    }

    /// Skips to the line that starts with `End <what>`; returns its index.
    fn skip_to(&mut self, end: &str, what: &str, from: usize) -> usize {
        self.li = from;
        self.ti = 0;
        while !self.at_end() {
            let hit = self.starts_with(&[end, what]);
            let at = self.li;
            self.next_line();
            if hit {
                return at;
            }
        }
        self.lines.len().saturating_sub(1)
    }

    fn skip_to_end_of_proc(&mut self, from: usize) -> usize {
        self.li = from;
        self.ti = 0;
        while !self.at_end() {
            let hit = self.starts_with(&["End", "Sub"])
                || self.starts_with(&["End", "Function"])
                || self.starts_with(&["End", "Property"]);
            let at = self.li;
            self.next_line();
            if hit {
                return at;
            }
        }
        self.lines.len().saturating_sub(1)
    }

    fn type_name(&mut self) -> PResult<TypeName> {
        let new = self.eat_kw("New");
        let (mut name, _) = self.ident()?;
        while self.eat_op(".") {
            let (part, _) = self.ident()?;
            name = format!("{name}.{part}");
        }
        if self.eat_op("*") {
            // `String * 10`
            self.bump();
        }
        Ok(TypeName { name, new })
    }

    fn procedure(&mut self) -> PResult<Proc> {
        let line = self.li;
        let mut vis = None;
        let mut is_static = false;
        loop {
            if let Some(Tok::Word(w, None)) = self.peek() {
                let lw = w.to_ascii_lowercase();
                if ["public", "private", "friend"].contains(&lw.as_str()) {
                    vis = Some(w.clone());
                    self.ti += 1;
                    continue;
                }
                if lw == "static" {
                    is_static = true;
                    self.ti += 1;
                    continue;
                }
            }
            break;
        }
        let kind = if self.eat_kw("Sub") {
            "Sub".to_string()
        } else if self.eat_kw("Function") {
            "Function".to_string()
        } else if self.eat_kw("Property") {
            match self.ident()?.0.to_ascii_lowercase().as_str() {
                "get" => "Get".into(),
                "let" => "Let".into(),
                "set" => "Set".into(),
                _ => return Err("Property の種類が不明です".into()),
            }
        } else {
            return Err("プロシージャを読めません".into());
        };
        let (name, suffix) = self.ident()?;
        let mut params = Vec::new();
        if self.eat_op("(") && !self.eat_op(")") {
            loop {
                params.push(self.param()?);
                if self.eat_op(")") {
                    break;
                }
                self.expect_op(",")?;
            }
        }
        let mut ret = None;
        let mut ret_array = false;
        if self.eat_kw("As") {
            ret = Some(self.type_name()?);
            if self.eat_op("(") {
                self.expect_op(")")?;
                ret_array = true;
            }
        }
        if !self.at_stmt_end() {
            return Err("プロシージャの宣言を読めません".into());
        }
        self.finish_stmt();
        let end_kw = match kind.as_str() {
            "Sub" => "Sub",
            "Function" => "Function",
            _ => "Property",
        };
        let body = self.block(&|p| p.starts_with(&["End", end_kw]))?;
        if self.at_end() {
            return Err(format!("End {end_kw} がありません"));
        }
        self.finish_stmt();
        Ok(Proc {
            kind,
            vis,
            is_static,
            name,
            suffix,
            params,
            ret,
            ret_array,
            body,
            line,
        })
    }

    fn param(&mut self) -> PResult<Param> {
        let mut p = Param {
            name: String::new(),
            suffix: None,
            byval: None,
            optional: false,
            paramarray: false,
            ty: None,
            array: false,
            default: None,
        };
        loop {
            if self.eat_kw("Optional") {
                p.optional = true;
            } else if self.eat_kw("ByVal") {
                p.byval = Some(true);
            } else if self.eat_kw("ByRef") {
                p.byval = Some(false);
            } else if self.eat_kw("ParamArray") {
                p.paramarray = true;
            } else {
                break;
            }
        }
        let (name, suffix) = self.ident()?;
        p.name = name;
        p.suffix = suffix;
        if self.eat_op("(") {
            self.expect_op(")")?;
            p.array = true;
        }
        if self.eat_kw("As") {
            p.ty = Some(self.type_name()?);
        }
        if self.eat_op("=") {
            p.default = Some(self.expr()?);
        }
        Ok(p)
    }

    fn enumeration(&mut self, vis: Option<&str>) -> PResult<Item> {
        let line = self.li;
        while !self.eat_kw("Enum") {
            self.bump().ok_or("Enum を読めません")?;
        }
        let (name, _) = self.ident()?;
        self.next_line();
        let mut members = Vec::new();
        while !self.at_end() && !self.starts_with(&["End", "Enum"]) {
            let comment = self.lines[self.li].comment.clone();
            if self.toks().is_empty() {
                self.next_line();
                continue;
            }
            let (m, _) = self.ident()?;
            let value = if self.eat_op("=") {
                Some(self.expr()?)
            } else {
                None
            };
            members.push((m, value, comment));
            self.next_line();
        }
        if self.at_end() {
            return Err("End Enum がありません".into());
        }
        self.next_line();
        Ok(Item::Enum {
            name,
            public: vis.is_some_and(|v| v.eq_ignore_ascii_case("public")),
            members,
            line,
        })
    }

    /// Statements until `stop` holds at the start of a statement.
    fn block(&mut self, stop: &dyn Fn(&Parser) -> bool) -> PResult<Vec<Stmt>> {
        let mut out = Vec::new();
        loop {
            if self.at_end() || self.pending_next > 0 {
                return Ok(out);
            }
            if self.ti == 0 {
                if self.lines[self.li].bad {
                    let line = &self.lines[self.li];
                    out.push(Stmt {
                        kind: StmtKind::Todo("読み取れない行です".into()),
                        line: self.li,
                        comment: None,
                        blank_before: line.blank_before,
                    });
                    self.next_line();
                    continue;
                }
                if let Some(label) = self.lines[self.li].label.clone() {
                    let line = &self.lines[self.li];
                    out.push(Stmt {
                        kind: StmtKind::Label(label),
                        line: self.li,
                        comment: None,
                        blank_before: line.blank_before,
                    });
                }
                if self.toks().is_empty() {
                    let line = &self.lines[self.li];
                    if line.comment.is_some() || line.label.is_none() {
                        out.push(self.comment_stmt());
                    } else {
                        self.next_line();
                    }
                    continue;
                }
            }
            if stop(self) {
                return Ok(out);
            }
            out.push(self.statement());
        }
    }

    /// One statement; problems become a `Todo` that covers the rest of the line.
    fn statement(&mut self) -> Stmt {
        let li = self.li;
        let first = self.ti == 0;
        let line = &self.lines[li];
        let blank_before = first && line.blank_before;
        let save = (self.li, self.ti);
        let kind = match self.statement_kind() {
            Ok(k) => k,
            Err(reason) => {
                // Skip the rest of the line.
                if self.li == save.0 {
                    self.next_line();
                }
                StmtKind::Todo(reason)
            }
        };
        // The comment belongs to the last statement of its line.
        let comment = if self.li != li {
            self.lines[li].comment.clone()
        } else {
            None
        };
        Stmt {
            kind,
            line: li,
            comment,
            blank_before,
        }
    }

    /// Parses a statement and moves past its end.
    fn statement_kind(&mut self) -> PResult<StmtKind> {
        let Some(head) = self.peek().cloned() else {
            self.finish_stmt();
            return Ok(StmtKind::Nothing);
        };
        let kw = |k: &str| head.is_kw(k);
        if kw("Dim") || kw("Private") || kw("Public") || kw("Global") || kw("Static") || kw("Const")
        {
            let mut is_static = false;
            let mut is_const = false;
            let mut with_events = false;
            while let Some(Tok::Word(w, None)) = self.peek() {
                match w.to_ascii_lowercase().as_str() {
                    "dim" | "private" | "public" | "global" => self.ti += 1,
                    "static" => {
                        is_static = true;
                        self.ti += 1;
                    }
                    "const" => {
                        is_const = true;
                        self.ti += 1;
                    }
                    // The events of the object the variable holds: `@on("<variable>.<Event>")`.
                    "withevents" => {
                        with_events = true;
                        self.ti += 1;
                    }
                    _ => break,
                }
            }
            let mut vars = Vec::new();
            loop {
                let (name, suffix) = self.ident()?;
                if with_events {
                    self.with_events.push(name.clone());
                }
                let mut v = VarDecl {
                    name,
                    suffix,
                    ty: None,
                    dims: None,
                    value: None,
                };
                if self.eat_op("(") {
                    let mut dims = Vec::new();
                    if !self.eat_op(")") {
                        loop {
                            let lo = self.expr()?;
                            if self.eat_kw("To") {
                                dims.push(self.expr()?);
                                let _ = lo;
                            } else {
                                dims.push(lo);
                            }
                            if self.eat_op(")") {
                                break;
                            }
                            self.expect_op(",")?;
                        }
                    }
                    v.dims = Some(dims);
                }
                if self.eat_kw("As") {
                    v.ty = Some(self.type_name()?);
                }
                if is_const {
                    self.expect_op("=")?;
                    v.value = Some(self.expr()?);
                }
                vars.push(v);
                if !self.eat_op(",") {
                    break;
                }
            }
            self.end_of_stmt()?;
            return Ok(if is_const {
                StmtKind::Const(vars)
            } else {
                StmtKind::Dim { vars, is_static }
            });
        }
        if kw("If") {
            return self.if_stmt();
        }
        if kw("Select") {
            return self.select_stmt();
        }
        if kw("For") {
            return self.for_stmt();
        }
        if kw("Do") {
            self.ti += 1;
            let pre = self.loop_cond()?;
            self.end_of_stmt()?;
            let body = self.block(&|p| p.starts_with(&["Loop"]))?;
            if self.at_end() {
                return Err("Loop がありません".into());
            }
            self.ti += 1;
            let post = self.loop_cond()?;
            self.end_of_stmt()?;
            return Ok(StmtKind::Do { pre, post, body });
        }
        if kw("While") {
            self.ti += 1;
            let cond = self.expr()?;
            self.end_of_stmt()?;
            let body = self.block(&|p| p.starts_with(&["Wend"]))?;
            if self.at_end() {
                return Err("Wend がありません".into());
            }
            self.ti += 1;
            self.end_of_stmt()?;
            return Ok(StmtKind::Do {
                pre: Some((false, cond)),
                post: None,
                body,
            });
        }
        if kw("With") {
            self.ti += 1;
            let obj = self.expr()?;
            self.end_of_stmt()?;
            let body = self.block(&|p| p.starts_with(&["End", "With"]))?;
            if self.at_end() {
                return Err("End With がありません".into());
            }
            self.ti += 2;
            self.end_of_stmt()?;
            return Ok(StmtKind::With { obj, body });
        }
        if kw("Exit") {
            self.ti += 1;
            let (what, _) = self.ident()?;
            self.end_of_stmt()?;
            return Ok(StmtKind::Exit(what.to_ascii_lowercase()));
        }
        if kw("On") {
            self.ti += 1;
            if self.eat_kw("Error") {
                if self.eat_kw("Resume") && self.eat_kw("Next") {
                    self.end_of_stmt()?;
                    return Ok(StmtKind::OnErrorResumeNext);
                }
                if self.eat_kw("GoTo") {
                    let target = match self.bump() {
                        Some(Tok::Num(n, _)) => n.clone(),
                        Some(Tok::Word(w, _)) => w.clone(),
                        _ => return Err("On Error GoTo の飛び先がありません".into()),
                    };
                    self.end_of_stmt()?;
                    return Ok(if target == "0" {
                        StmtKind::OnErrorGoTo0
                    } else {
                        StmtKind::OnErrorGoTo(target)
                    });
                }
            }
            return Err("On … GoTo / GoSub には対応していません".into());
        }
        if kw("GoTo") || kw("GoSub") || kw("Return") || kw("Resume") {
            return Err(format!("{} には対応していません", head_word(&head)));
        }
        if kw("ReDim") || kw("Erase") {
            return Err(format!(
                "{} には対応していません（RAS の配列は push などで伸ばします）",
                head_word(&head)
            ));
        }
        if kw("Stop") || kw("End") && self.peek_at(1).is_none_or(|t| *t == Tok::Colon) {
            return Err(format!("{} には対応していません", head_word(&head)));
        }
        if kw("RaiseEvent") || kw("LSet") || kw("RSet") {
            return Err(format!("{} には対応していません", head_word(&head)));
        }
        if let Some(k) = self.file_statement()? {
            return Ok(k);
        }
        if kw("Set") || kw("Let") {
            self.ti += 1;
            let target = self.postfix()?;
            self.expect_op("=")?;
            let value = self.expr()?;
            self.end_of_stmt()?;
            return Ok(StmtKind::Assign { target, value });
        }
        if kw("Call") {
            self.ti += 1;
            let e = self.postfix()?;
            self.end_of_stmt()?;
            return Ok(StmtKind::Call(e));
        }
        // An assignment or a call.
        let target = self.postfix()?;
        if self.eat_op("=") {
            let value = self.expr()?;
            self.end_of_stmt()?;
            return Ok(StmtKind::Assign { target, value });
        }
        if self.at_stmt_end() || self.peek().is_some_and(|t| t.is_kw("Else")) {
            self.end_of_stmt()?;
            return Ok(StmtKind::Call(target));
        }
        // `Foo a, b`: arguments without parentheses.
        let (callee, mut args) = match target {
            // `Foo (a), b`: the parentheses belonged to the first argument.
            Expr::Call(callee, inner)
                if self.peek().is_some_and(|t| t.is_op(","))
                    && inner.len() == 1
                    && inner[0].name.is_none() =>
            {
                (
                    *callee,
                    vec![Arg {
                        name: None,
                        value: inner[0].value.clone().map(|v| Expr::Paren(Box::new(v))),
                    }],
                )
            }
            other => (other, Vec::new()),
        };
        if !args.is_empty() {
            self.expect_op(",")?;
        }
        loop {
            if self.peek().is_some_and(|t| t.is_op(",")) {
                args.push(Arg {
                    name: None,
                    value: None,
                });
                self.ti += 1;
                continue;
            }
            args.push(self.arg()?);
            if !self.eat_op(",") {
                break;
            }
            if self.at_stmt_end() {
                args.push(Arg {
                    name: None,
                    value: None,
                });
                break;
            }
        }
        self.end_of_stmt()?;
        Ok(StmtKind::Call(Expr::Call(Box::new(callee), args)))
    }

    /// The statements of the files opened by number; `None` when the statement is not one.
    fn file_statement(&mut self) -> PResult<Option<StmtKind>> {
        let Some(Tok::Word(head, None)) = self.peek().cloned() else {
            return Ok(None);
        };
        let head = head.to_ascii_lowercase();
        let next = self.peek_at(1).cloned();
        let hash_next = next.as_ref().is_some_and(|t| t.is_op("#"));
        // `Name = …`, `Width = …`: variables, not the statements.
        let assigns = next
            .as_ref()
            .is_some_and(|t| t.is_op("=") || t.is_op(".") || t.is_op("(") || t.is_op("!"));
        let has_word = |p: &Self, w: &str| {
            p.toks()[p.ti..]
                .iter()
                .take_while(|t| **t != Tok::Colon)
                .any(|t| t.is_kw(w))
        };
        let file = |args: Vec<FileArg>, func: &'static str| StmtKind::File {
            func,
            args,
            result: None,
        };
        let kind = match head.as_str() {
            "open" if !assigns && has_word(self, "For") => {
                self.ti += 1;
                let path = self.expr()?;
                if !self.eat_kw("For") {
                    return Err("Open … For を読めません".into());
                }
                let (mode, _) = self.ident()?;
                let mode = match mode.to_ascii_lowercase().as_str() {
                    "input" => "forInput",
                    "output" => "forOutput",
                    "append" => "forAppend",
                    "binary" => "forBinary",
                    "random" => "forRandom",
                    _ => return Err(format!("Open のモード `{mode}` を読めません")),
                };
                let mut access = None;
                if self.eat_kw("Access") {
                    access = Some(if self.eat_kw("Read") {
                        if self.eat_kw("Write") {
                            "accessReadWrite"
                        } else {
                            "accessRead"
                        }
                    } else if self.eat_kw("Write") {
                        "accessWrite"
                    } else {
                        return Err("Open … Access を読めません".into());
                    });
                }
                let mut share = None;
                if self.eat_kw("Shared") {
                    share = Some("lockShared");
                } else if self.eat_kw("Lock") {
                    share = Some(if self.eat_kw("Read") {
                        if self.eat_kw("Write") {
                            "lockReadWrite"
                        } else {
                            "lockRead"
                        }
                    } else if self.eat_kw("Write") {
                        "lockWrite"
                    } else {
                        return Err("Open … Lock を読めません".into());
                    });
                }
                if !self.eat_kw("As") {
                    return Err("Open … As を読めません".into());
                }
                let n = self.file_number()?;
                let len = if self.eat_kw("Len") {
                    self.expect_op("=")?;
                    Some(self.expr()?)
                } else {
                    None
                };
                let mut args = vec![FileArg::Value(n), FileArg::Value(path), FileArg::Word(mode)];
                if access.is_some() || share.is_some() || len.is_some() {
                    args.push(FileArg::Word(access.unwrap_or("accessDefault")));
                }
                if share.is_some() || len.is_some() {
                    args.push(FileArg::Word(share.unwrap_or("lockDefault")));
                }
                if let Some(len) = len {
                    args.push(FileArg::Value(len));
                }
                file(args, "fileOpen")
            }
            "close" if !assigns => {
                self.ti += 1;
                let mut args = Vec::new();
                if !self.at_stmt_end() {
                    loop {
                        args.push(FileArg::Value(self.file_number()?));
                        if !self.eat_op(",") {
                            break;
                        }
                    }
                }
                file(args, "fileClose")
            }
            "reset" if next.is_none() || next == Some(Tok::Colon) => {
                self.ti += 1;
                file(vec![], "reset")
            }
            "name" if !assigns && has_word(self, "As") => {
                self.ti += 1;
                let from = self.expr()?;
                if !self.eat_kw("As") {
                    return Err("Name … As を読めません".into());
                }
                let to = self.expr()?;
                file(vec![FileArg::Value(from), FileArg::Value(to)], "rename")
            }
            "print" if hash_next => {
                self.ti += 1;
                let n = self.file_number()?;
                let mut args = vec![FileArg::Value(n)];
                let mut newline = true;
                if self.eat_op(",") {
                    while !self.at_stmt_end() && !self.peek().is_some_and(|t| t.is_kw("Else")) {
                        if self.eat_op(";") {
                            newline = false;
                            continue;
                        }
                        if self
                            .peek()
                            .is_some_and(|t| t.is_op(",") || t.is_kw("Spc") || t.is_kw("Tab"))
                        {
                            return Err("Print # の `,`（次のタブ位置）・Spc・Tab には対応していません（`;` で区切るか、文字列を組み立ててください）".into());
                        }
                        args.push(FileArg::Value(self.expr()?));
                        newline = true;
                    }
                }
                file(args, if newline { "printLine" } else { "print" })
            }
            "write" if hash_next => {
                self.ti += 1;
                let n = self.file_number()?;
                let mut args = vec![FileArg::Value(n)];
                let mut newline = true;
                if self.eat_op(",") {
                    while !self.at_stmt_end() && !self.peek().is_some_and(|t| t.is_kw("Else")) {
                        args.push(FileArg::Value(self.expr()?));
                        newline = true;
                        if self.eat_op(",") || self.eat_op(";") {
                            newline = false;
                        } else {
                            break;
                        }
                    }
                }
                file(args, if newline { "writeLine" } else { "write" })
            }
            "input" if hash_next => {
                self.ti += 1;
                let n = self.file_number()?;
                self.expect_op(",")?;
                let mut args = vec![FileArg::Value(n)];
                loop {
                    args.push(FileArg::Target(self.postfix()?));
                    if !self.eat_op(",") {
                        break;
                    }
                }
                file(args, "input")
            }
            "line" if next.as_ref().is_some_and(|t| t.is_kw("Input")) => {
                self.ti += 2;
                let n = self.file_number()?;
                self.expect_op(",")?;
                let target = self.postfix()?;
                StmtKind::File {
                    func: "lineInput",
                    args: vec![FileArg::Value(n)],
                    result: Some(target),
                }
            }
            "get" | "put" if !assigns && next.is_some() => {
                self.ti += 1;
                let n = self.file_number()?;
                self.expect_op(",")?;
                let rec = if self.peek().is_some_and(|t| t.is_op(",")) {
                    None
                } else {
                    Some(self.expr()?)
                };
                self.expect_op(",")?;
                let var = self.postfix()?;
                let get = head == "get";
                let mut args = vec![
                    FileArg::Value(n),
                    if get {
                        FileArg::Target(var)
                    } else {
                        FileArg::Value(var)
                    },
                ];
                if let Some(r) = rec {
                    args.push(FileArg::Value(r));
                }
                file(args, if get { "fileGet" } else { "filePut" })
            }
            "seek" if !assigns && next.is_some() => {
                self.ti += 1;
                let n = self.file_number()?;
                self.expect_op(",")?;
                let pos = self.expr()?;
                file(vec![FileArg::Value(n), FileArg::Value(pos)], "seek")
            }
            "lock" | "unlock" if !assigns && next.is_some() => {
                self.ti += 1;
                let n = self.file_number()?;
                let mut args = vec![FileArg::Value(n)];
                if self.eat_op(",") {
                    // `Lock #1, To 5` starts at the first record.
                    let from = if self.peek().is_some_and(|t| t.is_kw("To")) {
                        Expr::Num("1".into(), false)
                    } else {
                        self.expr()?
                    };
                    args.push(FileArg::Value(from));
                    if self.eat_kw("To") {
                        args.push(FileArg::Value(self.expr()?));
                    }
                }
                file(args, if head == "lock" { "lock" } else { "unlock" })
            }
            "width" if hash_next => {
                self.ti += 1;
                let n = self.file_number()?;
                self.expect_op(",")?;
                let w = self.expr()?;
                file(vec![FileArg::Value(n), FileArg::Value(w)], "fileWidth")
            }
            _ => return Ok(None),
        };
        Ok(Some(self.finish_file(kind)?))
    }

    fn finish_file(&mut self, kind: StmtKind) -> PResult<StmtKind> {
        self.end_of_stmt()?;
        Ok(kind)
    }

    /// `#n` or `n`.
    fn file_number(&mut self) -> PResult<Expr> {
        self.eat_op("#");
        self.expr()
    }

    fn end_of_stmt(&mut self) -> PResult<()> {
        if self.peek().is_some_and(|t| t.is_kw("Else")) {
            // Single-line If: the caller continues with `Else`.
            return Ok(());
        }
        if !self.at_stmt_end() {
            return Err("文の終わりを読めません".into());
        }
        self.finish_stmt();
        Ok(())
    }

    fn loop_cond(&mut self) -> PResult<Option<(bool, Expr)>> {
        if self.eat_kw("While") {
            return Ok(Some((false, self.expr()?)));
        }
        if self.eat_kw("Until") {
            return Ok(Some((true, self.expr()?)));
        }
        Ok(None)
    }

    fn if_stmt(&mut self) -> PResult<StmtKind> {
        self.ti += 1;
        let cond = self.expr()?;
        if !self.eat_kw("Then") {
            return Err("Then がありません".into());
        }
        if !self.at_stmt_end() || self.peek() == Some(&Tok::Colon) && self.peek_at(1).is_some() {
            // Single-line If: `If c Then a: b Else d`
            if self.peek() == Some(&Tok::Colon) {
                self.ti += 1;
            }
            let line = self.li;
            let then = self.inline_stmts(line)?;
            let els = if self.li == line && self.eat_kw("Else") {
                Some(self.inline_stmts(line)?)
            } else {
                None
            };
            if self.li == line {
                self.next_line();
            }
            return Ok(StmtKind::If {
                arms: vec![(cond, then)],
                els,
            });
        }
        self.finish_stmt();
        let mut arms = Vec::new();
        let mut els = None;
        let stop = |p: &Parser| {
            p.starts_with(&["ElseIf"])
                || p.starts_with(&["Else"])
                || p.starts_with(&["End", "If"])
                || p.starts_with(&["EndIf"])
        };
        let body = self.block(&stop)?;
        arms.push((cond, body));
        loop {
            if self.at_end() {
                return Err("End If がありません".into());
            }
            if self.eat_kw("ElseIf") {
                let c = self.expr()?;
                if !self.eat_kw("Then") {
                    return Err("Then がありません".into());
                }
                self.end_of_stmt()?;
                let body = self.block(&stop)?;
                arms.push((c, body));
            } else if self.eat_kw("Else") {
                if self.peek() == Some(&Tok::Colon) {
                    self.ti += 1;
                }
                if self.peek().is_none() {
                    self.next_line();
                }
                els = Some(self.block(&stop)?);
            } else {
                // End If
                if !self.eat_kw("EndIf") {
                    self.ti += 2;
                }
                self.end_of_stmt()?;
                break;
            }
        }
        Ok(StmtKind::If { arms, els })
    }

    /// Statements of a single-line If, up to `Else` or the end of the line.
    fn inline_stmts(&mut self, line: usize) -> PResult<Vec<Stmt>> {
        let mut out = Vec::new();
        while self.li == line
            && self.peek().is_some()
            && !self.peek().is_some_and(|t| t.is_kw("Else"))
        {
            let kind = self.statement_kind()?;
            out.push(Stmt {
                kind,
                line,
                comment: None,
                blank_before: false,
            });
        }
        Ok(out)
    }

    fn select_stmt(&mut self) -> PResult<StmtKind> {
        self.ti += 1;
        if !self.eat_kw("Case") {
            return Err("Select Case を読めません".into());
        }
        let subject = self.expr()?;
        self.end_of_stmt()?;
        let stop = |p: &Parser| p.starts_with(&["Case"]) || p.starts_with(&["End", "Select"]);
        // Comments before the first Case are dropped with their lines.
        let _ = self.block(&stop)?;
        let mut cases = Vec::new();
        let mut els = None;
        loop {
            if self.at_end() {
                return Err("End Select がありません".into());
            }
            if self.starts_with(&["End", "Select"]) {
                self.ti += 2;
                self.end_of_stmt()?;
                break;
            }
            self.ti += 1; // Case
            if self.eat_kw("Else") {
                if self.peek() == Some(&Tok::Colon) {
                    self.ti += 1;
                }
                if self.peek().is_none() {
                    self.next_line();
                }
                els = Some(self.block(&stop)?);
                continue;
            }
            let mut items = Vec::new();
            loop {
                if self.eat_kw("Is") {
                    let op = match self.bump() {
                        Some(Tok::Op(op)) if ["=", "<>", "<", ">", "<=", ">="].contains(op) => {
                            op.to_string()
                        }
                        _ => return Err("Case Is の比較演算子を読めません".into()),
                    };
                    items.push(CaseItem::Is(op, self.expr()?));
                } else {
                    let a = self.expr()?;
                    if self.eat_kw("To") {
                        items.push(CaseItem::Range(a, self.expr()?));
                    } else {
                        items.push(CaseItem::Value(a));
                    }
                }
                if !self.eat_op(",") {
                    break;
                }
            }
            if self.peek() == Some(&Tok::Colon) {
                self.ti += 1;
            }
            if self.peek().is_none() {
                self.next_line();
            }
            let body = self.block(&stop)?;
            cases.push((items, body));
        }
        Ok(StmtKind::Select {
            subject,
            cases,
            els,
        })
    }

    fn for_stmt(&mut self) -> PResult<StmtKind> {
        self.ti += 1;
        if self.eat_kw("Each") {
            let (var, _) = self.ident()?;
            if !self.eat_kw("In") {
                return Err("For Each … In を読めません".into());
            }
            let coll = self.expr()?;
            self.end_of_stmt()?;
            let body = self.loop_body()?;
            return Ok(StmtKind::ForEach { var, coll, body });
        }
        let var = self.postfix()?;
        self.expect_op("=")?;
        let from = self.expr()?;
        if !self.eat_kw("To") {
            return Err("For … To を読めません".into());
        }
        let to = self.expr()?;
        let step = if self.eat_kw("Step") {
            Some(self.expr()?)
        } else {
            None
        };
        self.end_of_stmt()?;
        let body = self.loop_body()?;
        Ok(StmtKind::For {
            var,
            from,
            to,
            step,
            body,
        })
    }

    fn loop_body(&mut self) -> PResult<Vec<Stmt>> {
        let body = self.block(&|p| p.starts_with(&["Next"]))?;
        if self.pending_next > 0 {
            self.pending_next -= 1;
            return Ok(body);
        }
        if self.at_end() {
            return Err("Next がありません".into());
        }
        self.ti += 1;
        let mut names: usize = 0;
        while !self.at_stmt_end() {
            self.bump();
            names += 1;
            if !self.eat_op(",") {
                break;
            }
        }
        // `Next j, i` also closes the enclosing loop.
        self.pending_next = names.saturating_sub(1);
        self.end_of_stmt()?;
        Ok(body)
    }

    // ----------------------------------------------------------------------------- expressions

    fn expr(&mut self) -> PResult<Expr> {
        self.binary(0)
    }

    fn binary(&mut self, level: usize) -> PResult<Expr> {
        // Lowest to highest precedence.
        const LEVELS: &[&[&str]] = &[&["imp"], &["eqv"], &["xor"], &["or"], &["and"]];
        if level < LEVELS.len() {
            let mut lhs = self.binary(level + 1)?;
            while let Some(Tok::Word(w, None)) = self.peek() {
                let lw = w.to_ascii_lowercase();
                if !LEVELS[level].contains(&lw.as_str()) {
                    break;
                }
                self.ti += 1;
                let rhs = self.binary(level + 1)?;
                lhs = Expr::Binary(lw, Box::new(lhs), Box::new(rhs));
            }
            return Ok(lhs);
        }
        self.not_expr()
    }

    fn not_expr(&mut self) -> PResult<Expr> {
        if self.eat_kw("Not") {
            return Ok(Expr::Unary("not", Box::new(self.not_expr()?)));
        }
        self.comparison()
    }

    fn comparison(&mut self) -> PResult<Expr> {
        let mut lhs = self.concat()?;
        loop {
            let op = match self.peek() {
                Some(Tok::Op(op)) if ["=", "<>", "<", ">", "<=", ">="].contains(op) => {
                    op.to_string()
                }
                Some(t) if t.is_kw("Is") => "is".into(),
                Some(t) if t.is_kw("Like") => "like".into(),
                _ => break,
            };
            self.ti += 1;
            let rhs = self.concat()?;
            lhs = Expr::Binary(op, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn concat(&mut self) -> PResult<Expr> {
        let mut lhs = self.additive()?;
        while self.eat_op("&") {
            let rhs = self.additive()?;
            lhs = Expr::Binary("&".into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn additive(&mut self) -> PResult<Expr> {
        let mut lhs = self.modulo()?;
        loop {
            let op = if self.eat_op("+") {
                "+"
            } else if self.eat_op("-") {
                "-"
            } else {
                break;
            };
            let rhs = self.modulo()?;
            lhs = Expr::Binary(op.into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn modulo(&mut self) -> PResult<Expr> {
        let mut lhs = self.intdiv()?;
        while self.eat_kw("Mod") {
            let rhs = self.intdiv()?;
            lhs = Expr::Binary("mod".into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn intdiv(&mut self) -> PResult<Expr> {
        let mut lhs = self.multiplicative()?;
        while self.eat_op("\\") {
            let rhs = self.multiplicative()?;
            lhs = Expr::Binary("\\".into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn multiplicative(&mut self) -> PResult<Expr> {
        let mut lhs = self.unary()?;
        loop {
            let op = if self.eat_op("*") {
                "*"
            } else if self.eat_op("/") {
                "/"
            } else {
                break;
            };
            let rhs = self.unary()?;
            lhs = Expr::Binary(op.into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn unary(&mut self) -> PResult<Expr> {
        if self.eat_op("-") {
            return Ok(Expr::Unary("-", Box::new(self.unary()?)));
        }
        if self.eat_op("+") {
            return self.unary();
        }
        self.power()
    }

    fn power(&mut self) -> PResult<Expr> {
        let mut lhs = self.postfix()?;
        while self.eat_op("^") {
            // `2 ^ -1`
            let rhs = if self.eat_op("-") {
                Expr::Unary("-", Box::new(self.postfix()?))
            } else {
                self.postfix()?
            };
            lhs = Expr::Binary("^".into(), Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn postfix(&mut self) -> PResult<Expr> {
        let mut e = if self.peek().is_some_and(|t| t.is_op(".")) {
            // `.member` inside `With`
            self.ti += 1;
            let (name, _) = self.ident()?;
            Expr::Member(None, name)
        } else {
            self.primary()?
        };
        loop {
            if self.eat_op(".") {
                let (name, _) = self.ident()?;
                e = Expr::Member(Some(Box::new(e)), name);
            } else if self.eat_op("!") {
                let (name, _) = self.ident()?;
                let _ = (e, name);
                e = Expr::Bang;
            } else if self.peek().is_some_and(|t| t.is_op("(")) {
                self.ti += 1;
                let mut args = Vec::new();
                if !self.eat_op(")") {
                    loop {
                        if self.peek().is_some_and(|t| t.is_op(",")) {
                            args.push(Arg {
                                name: None,
                                value: None,
                            });
                        } else if self.peek().is_some_and(|t| t.is_op(")")) {
                            args.push(Arg {
                                name: None,
                                value: None,
                            });
                            self.ti += 1;
                            break;
                        } else {
                            args.push(self.arg()?);
                        }
                        if self.eat_op(")") {
                            break;
                        }
                        self.expect_op(",")?;
                    }
                }
                e = Expr::Call(Box::new(e), args);
            } else {
                break;
            }
        }
        Ok(e)
    }

    fn arg(&mut self) -> PResult<Arg> {
        if let (Some(Tok::Word(w, _)), Some(Tok::Op(":="))) = (self.peek(), self.peek_at(1)) {
            let name = w.clone();
            self.ti += 2;
            return Ok(Arg {
                name: Some(name),
                value: Some(self.expr()?),
            });
        }
        if self.eat_kw("ByVal") {
            return Err("引数の ByVal には対応していません".into());
        }
        Ok(Arg {
            name: None,
            value: Some(self.expr()?),
        })
    }

    fn primary(&mut self) -> PResult<Expr> {
        let Some(t) = self.bump().cloned() else {
            return Err("式がありません".into());
        };
        Ok(match t {
            Tok::Num(n, f) => Expr::Num(n, f),
            Tok::Str(s) => Expr::Str(s),
            Tok::Date(d) => Expr::Date(d),
            Tok::Op("(") => {
                let e = self.expr()?;
                self.expect_op(")")?;
                Expr::Paren(Box::new(e))
            }
            Tok::Bracketed(w) => Expr::Ident(w, None),
            Tok::Word(w, s) => match w.to_ascii_lowercase().as_str() {
                "true" => Expr::Bool(true),
                "false" => Expr::Bool(false),
                "nothing" => Expr::Nothing,
                "empty" => Expr::Empty,
                "null" => Expr::Null,
                "me" => Expr::Me,
                "new" => Expr::New(self.type_name()?.name),
                "typeof" => {
                    let e = self.postfix()?;
                    if !self.eat_kw("Is") {
                        return Err("TypeOf … Is を読めません".into());
                    }
                    let _ = (e, self.type_name()?);
                    Expr::TypeOf
                }
                "addressof" => return Err("AddressOf には対応していません".into()),
                _ if is_keyword(&w)
                    && !["mid", "string", "date", "error", "get", "input", "len"]
                        .contains(&w.to_ascii_lowercase().as_str()) =>
                {
                    return Err(format!("`{w}` を式として読めません"));
                }
                _ => Expr::Ident(w, s),
            },
            // A file number among the arguments (`Input(5, #1)`).
            Tok::Op("#") => return self.primary(),
            Tok::Op(op) => return Err(format!("`{op}` を式として読めません")),
            Tok::Colon => return Err("式がありません".into()),
        })
    }
}

fn head_word(t: &Tok) -> String {
    match t {
        Tok::Word(w, _) => w.clone(),
        _ => String::new(),
    }
}

// ---------------------------------------------------------------------------------------------
// Emitter

#[derive(Clone, Debug)]
struct VarInfo {
    /// RAS spelling.
    name: String,
    /// RAS type (`int`, `MSForms.ReturnBoolean`, `string[]`).
    ty: String,
}

impl VarInfo {
    /// Indexed with `[]` rather than called (`a(i)` → `a[i]`).
    fn indexable(&self) -> bool {
        !Emitter::is_object_type(&self.ty)
    }

    /// `MSForms.ReturnBoolean` and the like: the value is the default member `Value`.
    fn has_value_member(&self) -> bool {
        self.ty.starts_with("MSForms.Return")
    }

    fn is_numeric(&self) -> bool {
        matches!(
            self.ty.as_str(),
            "int" | "short" | "byte" | "long" | "float" | "double" | "currency" | "decimal"
        )
    }
}

#[derive(Clone, Debug)]
struct ProcInfo {
    /// RAS name.
    name: String,
    is_function: bool,
    exported: bool,
}

/// Precedence of an emitted RAS expression (higher binds tighter).
mod prec {
    pub const OR: u8 = 1;
    pub const AND: u8 = 2;
    pub const BITOR: u8 = 3;
    pub const BITXOR: u8 = 4;
    pub const BITAND: u8 = 5;
    pub const EQ: u8 = 6;
    pub const REL: u8 = 7;
    pub const ADD: u8 = 9;
    pub const MUL: u8 = 10;
    pub const POW: u8 = 11;
    pub const UNARY: u8 = 12;
    pub const CAST: u8 = 8;
    pub const PRIMARY: u8 = 20;
}

type EResult = Result<(String, u8), String>;

struct Emitter<'a> {
    host: &'a HostIndex,
    target: Target<'a>,
    peers: &'a [Peer],
    lines: &'a [Line],
    out: Vec<String>,
    todos: Vec<(usize, String)>,
    indent: usize,
    /// Variables declared `WithEvents`: their event procedures are handlers.
    with_events: Vec<String>,
    controls: HashMap<String, String>,
    module_vars: HashMap<String, VarInfo>,
    procs: HashMap<String, ProcInfo>,
    /// Members of user enums: lower-cased member → `Enum.member`.
    enum_members: HashMap<String, String>,
    enums: HashMap<String, String>,
    /// Locals of the procedure being emitted.
    locals: HashMap<String, VarInfo>,
    /// Lower-cased name of the function being emitted and its result variable.
    result: Option<(String, String)>,
    result_used: bool,
    with_stack: Vec<String>,
    with_n: usize,
    /// Inside a `catch (e)` converted from an `On Error GoTo` handler.
    in_catch: bool,
    /// Names already used at module level (RAS names, lower-cased).
    taken: HashSet<String>,
    /// `Integer` is `short` (parameters of event procedures).
    exact_integer: bool,
    /// Code compiled by RAS: handler (lower-cased) → its events (`btnOK.Click`, re-entrant).
    handlers: HashMap<String, Vec<(String, bool)>>,
    /// Event procedures RAS generated (skipped) and `RAS_ModInit`.
    glue: HashSet<String>,
    /// Initial values of module variables, from a generated `RAS_ModInit` (lower-cased name → RAS).
    inits: HashMap<String, String>,
    /// Names imported from peers: module key → RAS names.
    imports: BTreeMap<String, BTreeSet<String>>,
    /// Peers used as a namespace (`Module1.Foo`): module key → alias.
    namespace_imports: BTreeMap<String, String>,
}

impl<'a> Emitter<'a> {
    fn new(
        host: &'a HostIndex,
        target: Target<'a>,
        peers: &'a [Peer],
        lines: &'a [Line],
    ) -> Emitter<'a> {
        let controls = match target {
            Target::Form(form) => form
                .controls
                .iter()
                .map(|c| (c.name.to_lowercase(), c.name.clone()))
                .collect(),
            Target::Document { controls, .. } => controls
                .iter()
                .map(|c| (c.name.to_lowercase(), c.name.clone()))
                .collect(),
            _ => HashMap::new(),
        };
        Emitter {
            host,
            target,
            peers,
            lines,
            out: Vec::new(),
            todos: Vec::new(),
            indent: 0,
            with_events: Vec::new(),
            controls,
            module_vars: HashMap::new(),
            procs: HashMap::new(),
            enum_members: HashMap::new(),
            enums: HashMap::new(),
            locals: HashMap::new(),
            result: None,
            result_used: false,
            with_stack: Vec::new(),
            with_n: 0,
            in_catch: false,
            taken: HashSet::new(),
            exact_integer: false,
            handlers: HashMap::new(),
            glue: HashSet::new(),
            inits: HashMap::new(),
            imports: BTreeMap::new(),
            namespace_imports: BTreeMap::new(),
        }
    }

    fn finish(mut self) -> Converted {
        while self.out.last().is_some_and(|l| l.is_empty()) {
            self.out.pop();
        }
        // Imports of the names used from the peers.
        let from = |key: &str| match self.target {
            // A code-behind lies in `forms/`, the file of a sheet in `objects/`; module keys
            // are relative to `src/`.
            Target::Form(_) | Target::Document { .. } => format!("../{key}"),
            _ => format!("./{key}"),
        };
        let mut head: Vec<String> = Vec::new();
        for (key, names) in &self.imports {
            head.push(format!(
                "import {{ {} }} from \"{}\";",
                names.iter().cloned().collect::<Vec<_>>().join(", "),
                from(key)
            ));
        }
        for (key, alias) in &self.namespace_imports {
            head.push(format!("import * as {alias} from \"{}\";", from(key)));
        }
        if !head.is_empty() && !self.out.is_empty() {
            head.push(String::new());
        }
        let shift = head.len();
        head.append(&mut self.out);
        let mut text = head.join("\n");
        if !text.is_empty() {
            text.push('\n');
        }
        let todos = self
            .todos
            .into_iter()
            .map(|(line, message)| (line + shift, message))
            .collect();
        Converted { text, todos }
    }

    /// What `Me` is: the form, the sheet, or the workbook.
    fn self_ref(&self) -> Option<String> {
        match self.target {
            Target::Form(_) | Target::Document { .. } => Some("this".into()),
            Target::Standard => None,
        }
    }

    /// Members of the document a document module belongs to (`Range` of a sheet).
    fn document_members(&self) -> Option<&'a HashMap<String, (String, bool)>> {
        match self.target {
            Target::Document { workbook: true, .. } => Some(&self.host.workbook),
            Target::Document { .. } => Some(&self.host.worksheet),
            _ => None,
        }
    }

    /// Where the converted code ends up, for messages.
    fn place(&self) -> &'static str {
        match self.target {
            Target::Form(_) => "フォームのコードビハインド",
            Target::Document { workbook: true, .. } => "`ThisWorkbook` のファイル",
            Target::Document { .. } => "シートのファイル",
            _ => "RAS のモジュール",
        }
    }

    /// A public name of a peer used without a qualifier: imported by name.
    fn peer_name(&mut self, lower: &str) -> Option<PublicName> {
        let peers = self.peers;
        let (peer, name) = peers.iter().filter(|p| !p.document).find_map(|p| {
            p.names
                .iter()
                .find(|n| n.vba.to_lowercase() == lower)
                .map(|n| (p, n))
        })?;
        self.imports
            .entry(peer.key.clone())
            .or_default()
            .insert(name.ras.clone());
        Some(name.clone())
    }

    /// `Module1.Foo`: a public name of a peer used with its module's name.
    fn peer_member(&mut self, module: &str, lower: &str) -> Option<(String, PublicName)> {
        let peers = self.peers;
        let peer = peers
            .iter()
            .find(|p| p.vba_name.eq_ignore_ascii_case(module))?;
        let name = peer.names.iter().find(|n| n.vba.to_lowercase() == lower)?;
        if peer.document {
            // A member of the sheet (of the workbook): `Sheet1.foo`, `thisWorkbook.foo`.
            let object = if peer.vba_name.eq_ignore_ascii_case("ThisWorkbook") {
                "thisWorkbook".to_string()
            } else {
                peer.vba_name.clone()
            };
            return Some((object, name.clone()));
        }
        let alias = crate::collect::vba_module_name(&peer.key);
        self.namespace_imports
            .insert(peer.key.clone(), alias.clone());
        Some((alias, name.clone()))
    }

    fn line(&mut self, s: impl AsRef<str>) {
        let s = s.as_ref();
        if s.is_empty() {
            self.out.push(String::new());
        } else {
            self.out.push(format!("{}{s}", "    ".repeat(self.indent)));
        }
    }

    fn blank(&mut self) {
        if self
            .out
            .last()
            .is_some_and(|l| !l.is_empty() && !l.trim_end().ends_with('{'))
        {
            self.out.push(String::new());
        }
    }

    fn comment(&mut self, c: &str) {
        let c = c.trim_end();
        if c.is_empty() {
            self.line("//");
        } else if c.starts_with(' ') {
            self.line(format!("//{c}"));
        } else {
            self.line(format!("// {c}"));
        }
    }

    /// Keeps original VBA lines as comments and records the problem.
    fn todo(&mut self, from: usize, to: usize, reason: &str) {
        self.line(format!("// TODO(VBA): {reason}"));
        self.todos.push((self.out.len(), reason.to_string()));
        for l in from..=to.min(self.lines.len().saturating_sub(1)) {
            for src in self.lines[l].src.lines() {
                self.line(format!("//   {}", src.trim()));
            }
        }
    }

    /// A RAS identifier for a user name (RAS keywords get a trailing `_`).
    fn user_ident(name: &str) -> String {
        if names::is_ras_keyword(name)
            || [
                "of",
                "as",
                "from",
                "struct",
                "declare",
                "module",
                "namespace",
                "type",
                "Math",
                "console",
                "platform",
            ]
            .contains(&name)
        {
            format!("{name}_")
        } else {
            name.to_string()
        }
    }

    fn ras_type(&mut self, t: Option<&TypeName>, suffix: Option<char>, array: bool) -> String {
        let base = match t {
            Some(t) => self.type_of(&t.name),
            None => match suffix {
                Some('$') => "string".into(),
                Some('%') if self.exact_integer => "short".into(),
                Some('%') => "int".into(),
                Some('&') => "int".into(),
                Some('!') => "float".into(),
                Some('#') => "double".into(),
                Some('@') => "currency".into(),
                _ => "any".into(),
            },
        };
        if array {
            format!("{base}[]")
        } else {
            base
        }
    }

    fn type_of(&mut self, name: &str) -> String {
        let lower = name.to_ascii_lowercase();
        let lower = lower.strip_prefix("vba.").unwrap_or(&lower).to_string();
        match lower.as_str() {
            "long" => "int".into(),
            // VBA's 16-bit Integer overflows at 32767; `int` keeps the code working (event
            // parameters keep `short`, their declaration must match).
            "integer" if self.exact_integer => "short".into(),
            "integer" => "int".into(),
            "byte" => "byte".into(),
            "single" => "float".into(),
            "double" => "double".into(),
            "currency" => "currency".into(),
            "string" => "string".into(),
            "boolean" => "boolean".into(),
            "date" => "date".into(),
            "variant" => "any".into(),
            "object" => "object".into(),
            "longlong" => "long".into(),
            "longptr" => "ptr".into(),
            "decimal" => "decimal".into(),
            _ => {
                let short = lower.rsplit('.').next().unwrap_or(&lower).to_string();
                if let Some(e) = self.enums.get(&short) {
                    return e.clone();
                }
                if let Target::Form(form) = self.target {
                    if short == form.name.to_ascii_lowercase() || short == "userform" {
                        return "MSForms.UserForm".into();
                    }
                }
                match self.host.classes.get(&short) {
                    Some(q) => q.clone(),
                    None => "object".into(),
                }
            }
        }
    }

    fn is_object_type(t: &str) -> bool {
        !matches!(
            t,
            "int"
                | "short"
                | "byte"
                | "float"
                | "double"
                | "currency"
                | "string"
                | "boolean"
                | "date"
                | "any"
                | "long"
                | "ptr"
                | "decimal"
        ) && !t.ends_with("[]")
    }

    // ----------------------------------------------------------------------------- module

    fn module(&mut self, items: &[Item]) {
        // Names first: procedures, module variables, enums.
        for item in items {
            match item {
                Item::Proc(p) => {
                    if self.event_of(&p.name).is_none() {
                        let name = Self::user_ident(&p.name);
                        self.taken.insert(name.to_lowercase());
                        let exported = !is_private(p.vis.as_deref());
                        self.procs.insert(
                            p.name.to_lowercase(),
                            ProcInfo {
                                name,
                                is_function: p.kind == "Function" || p.kind == "Get",
                                exported,
                            },
                        );
                    }
                }
                Item::Stmt(Stmt {
                    kind: StmtKind::Dim { vars, .. } | StmtKind::Const(vars),
                    ..
                }) => {
                    for v in vars {
                        let ty = self.ras_type(v.ty.as_ref(), v.suffix, v.dims.is_some());
                        let name = Self::user_ident(&v.name);
                        self.taken.insert(name.to_lowercase());
                        self.module_vars
                            .insert(v.name.to_lowercase(), VarInfo { name, ty });
                    }
                }
                Item::Enum { name, members, .. } => {
                    self.enums.insert(name.to_lowercase(), name.clone());
                    self.taken.insert(name.to_lowercase());
                    for (m, _, _) in members {
                        self.enum_members
                            .insert(m.to_lowercase(), format!("{name}.{m}"));
                    }
                }
                _ => {}
            }
        }
        self.find_glue(items);
        let mut first = true;
        for item in items {
            if let Item::Proc(p) = item {
                if self.glue.contains(&p.name.to_lowercase()) {
                    continue;
                }
            }
            match item {
                Item::Stmt(s) => {
                    if s.blank_before && !first {
                        self.blank();
                    }
                    self.module_stmt(s);
                }
                Item::Proc(p) => {
                    if !first {
                        self.blank();
                    }
                    self.procedure(p);
                }
                Item::Enum {
                    name,
                    public,
                    members,
                    line,
                } => {
                    if !first {
                        self.blank();
                    }
                    let export = match self.target {
                        Target::Form(_) => {
                            if *public {
                                self.todo(*line, *line, "Public な列挙型はフォームのコードビハインドで公開できません（フォームの中だけで使います）");
                            }
                            ""
                        }
                        Target::Document { .. } => {
                            if *public {
                                self.todo(*line, *line, "Public な列挙型はシート・ThisWorkbook のファイルで公開できません（このファイルの中だけで使います。ほかのモジュールでも使うなら src/ のモジュールに移してください）");
                            }
                            ""
                        }
                        _ if *public => "export ",
                        _ => "",
                    };
                    self.line(format!("{export}enum {name} {{"));
                    self.indent += 1;
                    let n = members.len();
                    for (i, (m, value, comment)) in members.iter().enumerate() {
                        let sep = if i + 1 < n { "," } else { "" };
                        let value = match value.as_ref().map(|v| self.expr(v)) {
                            Some(Ok((v, _))) => format!(" = {v}"),
                            _ => String::new(),
                        };
                        let comment = comment
                            .as_ref()
                            .map(|c| format!(" // {}", c.trim()))
                            .unwrap_or_default();
                        self.line(format!("{m}{value}{sep}{comment}"));
                    }
                    self.indent -= 1;
                    self.line("}");
                }
                Item::Todo { from, to, reason } => {
                    if !first {
                        self.blank();
                    }
                    self.todo(*from, *to, reason);
                }
            }
            first = false;
        }
    }

    fn module_stmt(&mut self, s: &Stmt) {
        match &s.kind {
            StmtKind::Dim { vars, is_static: _ } => {
                let public = is_public_decl(&self.lines[s.line]);
                let form = !matches!(self.target, Target::Standard);
                if public && matches!(self.target, Target::Document { .. }) {
                    self.todo(s.line, s.line, "Public 変数はシート・ThisWorkbook のファイルで公開できません（このファイルの中だけの変数にしました。外から使うときは export した関数で受け渡してください）");
                } else if public && form {
                    self.todo(s.line, s.line, "Public 変数はフォームのコードビハインドで公開できません（フォームの中だけの変数にしました。外から使うときは export した関数で受け渡してください）");
                }
                let start = self.out.len();
                self.var_decls(vars, s);
                if public && !form {
                    self.export_lines(start);
                }
            }
            StmtKind::Const(vars) => {
                let public = is_public_decl(&self.lines[s.line]);
                let form = !matches!(self.target, Target::Standard);
                if public && matches!(self.target, Target::Document { .. }) {
                    self.todo(s.line, s.line, "Public 定数はシート・ThisWorkbook のファイルで公開できません（このファイルの中だけの定数にしました。ほかのモジュールでも使うなら src/ のモジュールに移してください）");
                } else if public && form {
                    self.todo(s.line, s.line, "Public 定数はフォームのコードビハインドで公開できません（フォームの中だけの定数にしました）");
                }
                let start = self.out.len();
                self.const_decls(vars, s);
                if public && !form {
                    self.export_lines(start);
                }
            }
            StmtKind::Nothing => {
                if let Some(c) = &s.comment {
                    self.comment(c);
                }
            }
            StmtKind::Todo(reason) => self.todo(s.line, s.line, reason),
            _ => self.todo(s.line, s.line, "モジュールレベルの文には対応していません"),
        }
    }

    /// Adds `export` to the `let` / `const` declarations written from line `start` on.
    fn export_lines(&mut self, start: usize) {
        for line in &mut self.out[start..] {
            let indent = line.len() - line.trim_start().len();
            if line[indent..].starts_with("let ") || line[indent..].starts_with("const ") {
                line.insert_str(indent, "export ");
            }
        }
    }

    fn trailing(&self, s: &Stmt) -> String {
        s.comment
            .as_ref()
            .map(|c| format!(" // {}", c.trim()))
            .unwrap_or_default()
    }

    fn var_decls(&mut self, vars: &[VarDecl], s: &Stmt) {
        for (i, v) in vars.iter().enumerate() {
            let ty = self.ras_type(v.ty.as_ref(), v.suffix, v.dims.is_some());
            let name = Self::user_ident(&v.name);
            let comment = if i + 1 == vars.len() {
                self.trailing(s)
            } else {
                String::new()
            };
            if self.result.is_some() || self.indent > 0 {
                self.locals.insert(
                    v.name.to_lowercase(),
                    VarInfo {
                        name: name.clone(),
                        ty: ty.clone(),
                    },
                );
            }
            if v.ty.as_ref().is_some_and(|t| t.new) {
                self.line(format!("let {name} = new {ty}();{comment}"));
            } else if v.dims.as_ref().is_some_and(|d| !d.is_empty()) {
                self.todo(s.line, s.line, "大きさを指定した配列には対応していません（空の配列にしました。push で要素を追加してください）");
                self.line(format!("let {name}: {ty} = [];{comment}"));
            } else if v.dims.is_some() {
                self.line(format!("let {name}: {ty} = [];{comment}"));
            } else if let Some(init) = self
                .inits
                .get(&v.name.to_lowercase())
                .filter(|_| self.indent == 0)
                .cloned()
            {
                self.line(format!("let {name}: {ty} = {init};{comment}"));
            } else {
                self.line(format!("let {name}: {ty};{comment}"));
            }
        }
    }

    fn const_decls(&mut self, vars: &[VarDecl], s: &Stmt) {
        for (i, v) in vars.iter().enumerate() {
            let name = Self::user_ident(&v.name);
            let comment = if i + 1 == vars.len() {
                self.trailing(s)
            } else {
                String::new()
            };
            let ty = if v.ty.is_some() || v.suffix.is_some() {
                format!(": {}", self.ras_type(v.ty.as_ref(), v.suffix, false))
            } else {
                String::new()
            };
            if self.indent > 0 {
                let ty = self.ras_type(v.ty.as_ref(), v.suffix, false);
                self.locals.insert(
                    v.name.to_lowercase(),
                    VarInfo {
                        name: name.clone(),
                        ty,
                    },
                );
            }
            match v.value.as_ref().map(|e| self.expr(e)) {
                Some(Ok((value, _))) => self.line(format!("const {name}{ty} = {value};{comment}")),
                Some(Err(reason)) => self.todo(s.line, s.line, &reason),
                None => {}
            }
        }
    }

    /// Recognises what RAS generated for a form compiled from a `.ras` code-behind (spec §14.11.2):
    /// event procedures that only call handlers (`@on`), and `RAS_ModInit` (initial values of
    /// the form's variables). They are turned back into decorators and initialisers.
    fn find_glue(&mut self, items: &[Item]) {
        let procs: HashSet<String> = items
            .iter()
            .filter_map(|i| match i {
                Item::Proc(p) => Some(p.name.to_lowercase()),
                _ => None,
            })
            .collect();
        let called = |e: &Expr| -> Option<String> {
            let name = match e {
                Expr::Ident(n, _) => n,
                Expr::Call(c, _) => match &**c {
                    Expr::Ident(n, _) => n,
                    _ => return None,
                },
                _ => return None,
            };
            procs
                .contains(&name.to_lowercase())
                .then(|| name.to_lowercase())
        };
        let is_flag =
            |e: &Expr| matches!(e, Expr::Ident(n, _) if n.to_lowercase().starts_with("ras_busy"));
        for item in items {
            let Item::Proc(p) = item else { continue };
            if p.name.eq_ignore_ascii_case("RAS_ModInit") {
                let mut inits = Vec::new();
                let all = p.body.iter().all(|st| match &st.kind {
                    StmtKind::Assign {
                        target: Expr::Ident(n, _),
                        value,
                    } if self.module_vars.contains_key(&n.to_lowercase()) => {
                        inits.push((n.to_lowercase(), value.clone()));
                        true
                    }
                    StmtKind::Nothing | StmtKind::Label(_) => true,
                    _ => false,
                });
                if all {
                    for (n, v) in inits {
                        if let Ok((t, _)) = self.expr(&v) {
                            self.inits.insert(n, t);
                        }
                    }
                    self.glue.insert(p.name.to_lowercase());
                }
                continue;
            }
            let Some((object, event)) = self.event_of(&p.name) else {
                continue;
            };
            let mut found: Vec<(String, bool)> = Vec::new();
            let mut generated = false;
            let ok = p.body.iter().all(|st| match &st.kind {
                StmtKind::Dim { is_static: true, vars } => vars.iter().all(|v| v.name.to_lowercase().starts_with("ras_busy")),
                StmtKind::OnErrorGoTo(l) => l.eq_ignore_ascii_case("RAS_OnError"),
                StmtKind::Label(l) => l.eq_ignore_ascii_case("RAS_OnError"),
                StmtKind::Exit(w) => w == "sub",
                StmtKind::Assign { target, .. } => is_flag(target),
                StmtKind::Nothing => true,
                StmtKind::Call(e) => {
                    if let Expr::Call(c, _) | Expr::Member(Some(c), _) = e {
                        if matches!(&**c, Expr::Ident(n, _) if n.eq_ignore_ascii_case("RAS_EventError")) {
                            generated = true;
                            return true;
                        }
                    }
                    match called(e) {
                        Some(h) if h == "ras_modinit" => true,
                        Some(h) => {
                            found.push((h, true));
                            true
                        }
                        None => false,
                    }
                }
                StmtKind::If { arms, els: None } if arms.len() == 1 && matches!(&arms[0].0, Expr::Unary("not", f) if is_flag(f)) => arms[0].1.iter().all(|inner| match &inner.kind {
                    StmtKind::Assign { target, .. } => is_flag(target),
                    StmtKind::Call(e) => match called(e) {
                        Some(h) => {
                            found.push((h, false));
                            true
                        }
                        None => false,
                    },
                    _ => false,
                }),
                _ => false,
            });
            if ok && generated {
                self.glue.insert(p.name.to_lowercase());
                for (h, reentrant) in found {
                    self.handlers
                        .entry(h)
                        .or_default()
                        .push((format!("{object}.{event}"), reentrant));
                }
            }
        }
    }

    /// `(object, event)` when `name` is an event procedure of the form or one of its controls,
    /// or of the sheet or workbook of a document module (`Sheet1`, `Change`; `Workbook`, `Open`).
    fn event_of(&self, name: &str) -> Option<(String, String)> {
        // `app_SheetChange` of `Private WithEvents app As Application` (spec §12.2.2): where
        // VBA allows such a variable, in the module of a form, of a sheet or of the workbook.
        if !matches!(self.target, Target::Standard) {
            for (i, _) in name.match_indices('_') {
                let (object, event) = (&name[..i], &name[i + 1..]);
                if let Some(variable) = self
                    .with_events
                    .iter()
                    .find(|v| !event.is_empty() && v.eq_ignore_ascii_case(object))
                {
                    return Some((Self::user_ident(variable), event.to_string()));
                }
            }
        }
        match self.target {
            Target::Form(form) => {
                for (i, _) in name.match_indices('_') {
                    let (object, event) = (&name[..i], &name[i + 1..]);
                    let host_type = if object.eq_ignore_ascii_case("UserForm") {
                        Some(("UserForm".to_string(), "UserForm".to_string()))
                    } else {
                        form.controls
                            .iter()
                            .find(|c| c.name.eq_ignore_ascii_case(object))
                            .map(|c| (c.control_type.clone(), c.name.clone()))
                    };
                    if let Some((ty, object)) = host_type {
                        if let Some(def) = crate::userform::find_event(&ty, event) {
                            return Some((object, def.name.to_string()));
                        }
                    }
                }
                None
            }
            Target::Standard => None,
            Target::Document {
                workbook, controls, ..
            } => {
                let prefix = if workbook { "Workbook" } else { "Worksheet" };
                if let Some((object, event)) = name.split_once('_') {
                    if object.eq_ignore_ascii_case(prefix) {
                        let event = crate::lower::document_event(workbook, event)?;
                        return Some((prefix.to_string(), event.to_string()));
                    }
                }
                // An ActiveX control on the sheet (spec §12.2.1).
                for (i, _) in name.match_indices('_') {
                    let (object, event) = (&name[..i], &name[i + 1..]);
                    if let Some(c) = controls
                        .iter()
                        .find(|c| c.name.eq_ignore_ascii_case(object))
                    {
                        if let Some((event, _)) =
                            crate::lower::sheet_control_event(&c.control_type, event)
                        {
                            return Some((c.name.clone(), event));
                        }
                        // A control that is not an MSForms one: its events are not known, the
                        // procedure is taken as it is declared.
                        if crate::userform::events_of(&c.control_type).is_empty()
                            && !event.is_empty()
                        {
                            return Some((c.name.clone(), event.to_string()));
                        }
                    }
                }
                None
            }
        }
    }

    /// `Worksheet_Xxx` / `Workbook_Xxx` of a document module for an event RAS cannot handle.
    fn unsupported_event(&self, name: &str) -> bool {
        let Target::Document { workbook, .. } = self.target else {
            return false;
        };
        let prefix = if workbook { "workbook_" } else { "worksheet_" };
        name.to_lowercase().starts_with(prefix) && self.event_of(name).is_none()
    }

    /// A name not used at module level yet.
    fn fresh(&mut self, base: &str) -> String {
        let mut name = base.to_string();
        let mut n = 2;
        while self.taken.contains(&name.to_lowercase())
            || self.controls.contains_key(&name.to_lowercase())
        {
            name = format!("{base}{n}");
            n += 1;
        }
        self.taken.insert(name.to_lowercase());
        name
    }

    fn procedure(&mut self, p: &Proc) {
        self.locals.clear();
        self.with_stack.clear();
        self.with_n = 0;
        self.result = None;
        self.result_used = false;
        let event = self.event_of(&p.name);
        if self.unsupported_event(&p.name) {
            self.todo(p.line, p.line, "RAS が対応していないイベントのプロシージャです（ふつうの関数にしました。イベントでは呼ばれません）");
        }
        if p.is_static {
            self.todo(
                p.line,
                p.line,
                "Static プロシージャには対応していません（変数は呼び出しごとに初期化されます）",
            );
        }
        let assigned = assigned_names(&p.body);
        let mut params = Vec::new();
        for prm in &p.params {
            let name = Self::user_ident(&prm.name);
            self.exact_integer = event.is_some() && !matches!(self.target, Target::Standard);
            let ty = self.ras_type(prm.ty.as_ref(), prm.suffix, prm.array || prm.paramarray);
            self.exact_integer = false;
            self.locals.insert(
                prm.name.to_lowercase(),
                VarInfo {
                    name: name.clone(),
                    ty: ty.clone(),
                },
            );
            if prm.paramarray {
                self.todo(
                    p.line,
                    p.line,
                    "ParamArray には対応していません（配列の引数にしました）",
                );
            }
            // VBA passes by reference unless ByVal: kept only when the procedure assigns to it.
            let byref = prm.byval != Some(true)
                && !prm.array
                && assigned.contains(&prm.name.to_lowercase());
            let mut s = format!("{}{name}", if byref { "byref " } else { "" });
            match (&prm.default, prm.optional) {
                (Some(d), _) => match self.expr(d) {
                    Ok((d, _)) => s.push_str(&format!(": {ty} = {d}")),
                    Err(_) => s.push_str(&format!("?: {ty}")),
                },
                (None, true) => s.push_str(&format!("?: {ty}")),
                (None, false) => s.push_str(&format!(": {ty}")),
            }
            params.push(s);
        }
        let header = match &event {
            Some((object, event)) => {
                let base = format!(
                    "{}{event}",
                    if object == "Workbook" || object == "Worksheet" {
                        "on".to_string()
                    } else {
                        names::js_name(object)
                    }
                );
                let name = self.fresh(&base);
                self.line(format!("@on(\"{object}.{event}\")"));
                format!("function {name}({})", params.join(", "))
            }
            None => {
                for (event, reentrant) in self
                    .handlers
                    .get(&p.name.to_lowercase())
                    .cloned()
                    .unwrap_or_default()
                {
                    self.line(if reentrant {
                        format!("@on(\"{event}\", {{ reentrant: true }})")
                    } else {
                        format!("@on(\"{event}\")")
                    });
                }
                let info = self
                    .procs
                    .get(&p.name.to_lowercase())
                    .cloned()
                    .unwrap_or(ProcInfo {
                        name: p.name.clone(),
                        is_function: false,
                        exported: false,
                    });
                let export = if info.exported { "export " } else { "" };
                let mut name = info.name.clone();
                if matches!(p.kind.as_str(), "Get" | "Let" | "Set") {
                    let place = self.place();
                    self.todo(
                        p.line,
                        p.line,
                        &format!("Property プロシージャは{place}では関数にしました"),
                    );
                    name = format!("{}{}", p.kind.to_ascii_lowercase(), name);
                }
                let ret = if p.kind == "Function" || p.kind == "Get" {
                    let t = self.ras_type(p.ret.as_ref(), p.suffix, p.ret_array);
                    let result = if self.locals.contains_key("result") {
                        "ras_result".to_string()
                    } else {
                        "result".to_string()
                    };
                    self.result = Some((p.name.to_lowercase(), result));
                    format!(": {t}")
                } else {
                    String::new()
                };
                format!("{export}function {name}({}){ret}", params.join(", "))
            }
        };
        let ret_ty = header.rsplit_once("): ").map(|(_, t)| t.to_string());
        self.line(format!("{header} {{"));
        self.indent += 1;
        let start = self.out.len();
        let body = restructure(&p.body);
        self.block(&body);
        if let (Some((_, result)), Some(t)) = (self.result.clone(), ret_ty) {
            if self.result_used {
                let init = match t.as_str() {
                    "int" | "short" | "byte" | "long" | "currency" | "ptr" => " = 0",
                    "float" | "double" => " = 0.0",
                    "string" => " = \"\"",
                    "boolean" => " = false",
                    t if t.ends_with("[]") => " = []",
                    _ => "",
                };
                self.out.insert(
                    start,
                    format!("{}let {result}: {t}{init};", "    ".repeat(self.indent)),
                );
                if !self
                    .out
                    .last()
                    .is_some_and(|l| l.trim_start().starts_with("return "))
                {
                    self.line(format!("return {result};"));
                }
            }
        }
        self.indent -= 1;
        self.line("}");
        self.result = None;
    }

    // ----------------------------------------------------------------------------- statements

    fn block(&mut self, stmts: &[Stmt]) {
        let mut i = 0;
        while i < stmts.len() {
            let s = &stmts[i];
            if s.blank_before && i > 0 {
                self.blank();
            }
            // `Fn = x` followed by `Exit Function` (or the end): `return x;`
            if let (
                Some((fname, _)),
                StmtKind::Assign {
                    target: Expr::Ident(n, _),
                    value,
                },
            ) = (self.result.clone(), &s.kind)
            {
                if n.to_lowercase() == fname && !self.locals.contains_key(&fname) {
                    let next_exit = stmts.get(i + 1).is_some_and(|n| matches!(&n.kind, StmtKind::Exit(w) if w == "function" || w == "property"));
                    let last = i + 1 == stmts.len() && self.indent == 1;
                    if next_exit || last {
                        match self.expr(value) {
                            Ok((v, _)) => self.line(format!("return {v};{}", self.trailing(s))),
                            Err(reason) => self.todo(s.line, s.line, &reason),
                        }
                        i += if next_exit { 2 } else { 1 };
                        continue;
                    }
                }
            }
            self.stmt(s);
            i += 1;
        }
    }

    fn stmt(&mut self, s: &Stmt) {
        if self.try_stmt(s) {
            return;
        }
        let tail = self.trailing(s);
        match &s.kind {
            StmtKind::Nothing => {
                if let Some(c) = &s.comment {
                    self.comment(c);
                }
            }
            StmtKind::Todo(reason) => self.todo(s.line, s.line, reason),
            StmtKind::File { func, args, result } => {
                match self.file_stmt(func, args, result.as_ref()) {
                    Ok(text) => self.line(format!("{text};{tail}")),
                    Err(reason) => self.todo(s.line, s.line, &reason),
                }
            }
            StmtKind::Dim { vars, is_static } => {
                if *is_static {
                    self.todo(s.line, s.line, "Static 変数には対応していません（呼び出しごとに初期化されます。値を残すにはモジュールの変数にしてください）");
                }
                self.var_decls(vars, s);
            }
            StmtKind::Const(vars) => self.const_decls(vars, s),
            StmtKind::Assign { target, value } => match self.assign(target, value) {
                Ok(text) => self.line(format!("{text};{tail}")),
                Err(reason) => self.todo(s.line, s.line, &reason),
            },
            StmtKind::Call(e) => match self.call_stmt(e) {
                Ok(text) => self.line(format!("{text};{tail}")),
                Err(reason) => self.todo(s.line, s.line, &reason),
            },
            StmtKind::If { arms, els } => {
                for (i, (cond, body)) in arms.iter().enumerate() {
                    let c = match self.cond(cond) {
                        Ok(c) => c,
                        Err(reason) => {
                            self.todo(s.line, s.line, &reason);
                            "false".into()
                        }
                    };
                    if i == 0 {
                        self.line(format!("if ({c}) {{{tail}"));
                    } else {
                        self.out_close_open(&format!("}} else if ({c}) {{"));
                    }
                    self.indent += 1;
                    self.block(body);
                    self.indent -= 1;
                }
                if let Some(els) = els {
                    self.out_close_open("} else {");
                    self.indent += 1;
                    self.block(els);
                    self.indent -= 1;
                }
                self.line("}");
            }
            StmtKind::Select {
                subject,
                cases,
                els,
            } => self.select(s, subject, cases, els.as_deref()),
            StmtKind::For {
                var,
                from,
                to,
                step,
                body,
            } => {
                let r = (|| -> Result<String, String> {
                    let v = self.lvalue(var)?;
                    let (f, _) = self.expr(from)?;
                    let (t, _) = self.expr(to)?;
                    let declared = matches!(var, Expr::Ident(n, _) if self.locals.contains_key(&n.to_lowercase()) || self.module_vars.contains_key(&n.to_lowercase()));
                    let init = if declared {
                        format!("{v} = {f}")
                    } else {
                        format!("let {v} = {f}")
                    };
                    Ok(match step {
                        None => format!("for ({init}; {v} <= {t}; {v}++)"),
                        Some(st) => {
                            let negative = matches!(st, Expr::Unary("-", _));
                            let (sv, _) = self.expr(st)?;
                            match (negative, sv.as_str()) {
                                (true, "-1") => format!("for ({init}; {v} >= {t}; {v}--)"),
                                (true, _) => format!("for ({init}; {v} >= {t}; {v} += {sv})"),
                                (false, "1") => format!("for ({init}; {v} <= {t}; {v}++)"),
                                (false, _) if matches!(st, Expr::Num(..)) => {
                                    format!("for ({init}; {v} <= {t}; {v} += {sv})")
                                }
                                _ => {
                                    return Err("Step が定数でない For には対応していません".into())
                                }
                            }
                        }
                    })
                })();
                match r {
                    Ok(head) => {
                        self.line(format!("{head} {{{tail}"));
                        self.indent += 1;
                        self.block(body);
                        self.indent -= 1;
                        self.line("}");
                    }
                    Err(reason) => self.todo_block(s, &reason, body),
                }
            }
            StmtKind::ForEach { var, coll, body } => match self.expr(coll) {
                Ok((c, _)) => {
                    let name = Self::user_ident(var);
                    let saved = self.locals.insert(
                        var.to_lowercase(),
                        VarInfo {
                            name: name.clone(),
                            ty: "any".into(),
                        },
                    );
                    self.line(format!("for (const {name} of {c}) {{{tail}"));
                    self.indent += 1;
                    self.block(body);
                    self.indent -= 1;
                    self.line("}");
                    if let Some(v) = saved {
                        self.locals.insert(var.to_lowercase(), v);
                    }
                }
                Err(reason) => self.todo_block(s, &reason, body),
            },
            StmtKind::Do { pre, post, body } => {
                let cond = |em: &mut Self, c: &Option<(bool, Expr)>| -> Result<String, String> {
                    match c {
                        None => Ok("true".into()),
                        Some((false, e)) => em.cond(e),
                        Some((true, e)) => {
                            let (t, p) = em.cond_expr(e)?;
                            Ok(if p >= prec::UNARY {
                                format!("!{t}")
                            } else {
                                format!("!({t})")
                            })
                        }
                    }
                };
                match (pre, post) {
                    (_, Some(_)) => match cond(self, post) {
                        Ok(c) => {
                            if pre.is_some() {
                                self.todo(s.line, s.line, "Do と Loop の両方に条件がある繰り返しは、Loop 側の条件だけにしました");
                            }
                            self.line(format!("do {{{tail}"));
                            self.indent += 1;
                            self.block(body);
                            self.indent -= 1;
                            self.line(format!("}} while ({c});"));
                        }
                        Err(reason) => self.todo_block(s, &reason, body),
                    },
                    (_, None) => match cond(self, pre) {
                        Ok(c) => {
                            self.line(format!("while ({c}) {{{tail}"));
                            self.indent += 1;
                            self.block(body);
                            self.indent -= 1;
                            self.line("}");
                        }
                        Err(reason) => self.todo_block(s, &reason, body),
                    },
                }
            }
            StmtKind::With { obj, body } => match self.with_target(obj) {
                Ok((text, temp)) => {
                    if let Some(init) = temp {
                        self.line(format!("{init}{tail}"));
                    }
                    self.with_stack.push(text);
                    self.block(body);
                    self.with_stack.pop();
                }
                Err(reason) => self.todo_block(s, &reason, body),
            },
            StmtKind::Exit(what) => match what.as_str() {
                "sub" | "property" if self.result.is_none() => self.line(format!("return;{tail}")),
                "function" | "property" => {
                    let r = self.result.clone().map(|r| r.1).unwrap_or_default();
                    self.result_used = true;
                    self.line(format!("return {r};{tail}"));
                }
                "sub" => self.line(format!("return;{tail}")),
                "for" | "do" => self.line(format!("break;{tail}")),
                _ => self.todo(s.line, s.line, "この Exit には対応していません"),
            },
            // Line numbers (`Erl`, also what RAS writes) are dropped.
            StmtKind::Label(l) if l.chars().all(|c| c.is_ascii_digit()) => {}
            StmtKind::Label(_) => self.todo(s.line, s.line, "行ラベルには対応していません"),
            StmtKind::OnErrorResumeNext | StmtKind::OnErrorGoTo0 | StmtKind::OnErrorGoTo(_) => self
                .todo(
                    s.line,
                    s.line,
                    "この位置の On Error には対応していません（try / catch に書き換えてください）",
                ),
        }
    }

    /// Replaces the `{` line just closed: `} else {`.
    fn out_close_open(&mut self, text: &str) {
        self.line(text);
    }

    fn todo_block(&mut self, s: &Stmt, reason: &str, body: &[Stmt]) {
        self.todo(s.line, s.line, reason);
        self.line("{");
        self.indent += 1;
        self.block(body);
        self.indent -= 1;
        self.line("}");
    }

    fn select(
        &mut self,
        s: &Stmt,
        subject: &Expr,
        cases: &[(Vec<CaseItem>, Vec<Stmt>)],
        els: Option<&[Stmt]>,
    ) {
        let tail = self.trailing(s);
        let subj = match self.value_expr(subject) {
            Ok(x) => x,
            Err(reason) => {
                self.todo(s.line, s.line, &reason);
                return;
            }
        };
        // A selector with side effects or a cost is evaluated once.
        let simple = matches!(subject, Expr::Ident(..) | Expr::Me)
            || matches!(subject, Expr::Member(Some(o), _) if matches!(**o, Expr::Ident(..) | Expr::Me));
        let subj_text = if simple {
            subj.0
        } else {
            self.with_n += 1;
            let name = format!("sel{}", self.with_n);
            self.line(format!("const {name} = {};", subj.0));
            name
        };
        let wrap = |t: &str, p: u8| {
            if p > prec::REL {
                t.to_string()
            } else {
                format!("({t})")
            }
        };
        let mut first = true;
        for (items, body) in cases {
            let mut conds = Vec::new();
            for item in items {
                let c = match item {
                    CaseItem::Value(v) => self
                        .expr(v)
                        .map(|(t, p)| format!("{subj_text} == {}", wrap(&t, p))),
                    CaseItem::Range(a, b) => self.expr(a).and_then(|(a, ap)| {
                        self.expr(b).map(|(b, bp)| {
                            format!(
                                "{subj_text} >= {} && {subj_text} <= {}",
                                wrap(&a, ap),
                                wrap(&b, bp)
                            )
                        })
                    }),
                    CaseItem::Is(op, v) => {
                        let op = match op.as_str() {
                            "=" => "==",
                            "<>" => "!=",
                            other => other,
                        };
                        let op = op.to_string();
                        self.expr(v)
                            .map(|(t, p)| format!("{subj_text} {op} {}", wrap(&t, p)))
                    }
                };
                match c {
                    Ok(c) => {
                        conds.push(if items.len() > 1 && matches!(item, CaseItem::Range(..)) {
                            format!("({c})")
                        } else {
                            c
                        })
                    }
                    Err(reason) => {
                        self.todo(s.line, s.line, &reason);
                        conds.push("false".into());
                    }
                }
            }
            let cond = conds.join(" || ");
            if first {
                self.line(format!("if ({cond}) {{{tail}"));
                first = false;
            } else {
                self.line(format!("}} else if ({cond}) {{"));
            }
            self.indent += 1;
            self.block(body);
            self.indent -= 1;
        }
        let els = els.filter(|e| {
            e.iter()
                .any(|s| !matches!(s.kind, StmtKind::Nothing) || s.comment.is_some())
        });
        if let Some(els) = els {
            if first {
                self.line("{");
            } else {
                self.line("} else {");
            }
            self.indent += 1;
            self.block(els);
            self.indent -= 1;
            self.line("}");
        } else if !first {
            self.line("}");
        }
    }

    fn with_target(&mut self, obj: &Expr) -> Result<(String, Option<String>), String> {
        let (text, _) = self.expr(obj)?;
        // A path without calls is repeated as written; anything else is evaluated once.
        let pure = text
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.');
        if pure {
            return Ok((text, None));
        }
        self.with_n += 1;
        let name = format!("w{}", self.with_n);
        Ok((name.clone(), Some(format!("const {name} = {text};"))))
    }

    fn assign(&mut self, target: &Expr, value: &Expr) -> Result<String, String> {
        if let (Some((fname, result)), Expr::Ident(n, _)) = (self.result.clone(), target) {
            if n.to_lowercase() == fname && !self.locals.contains_key(&fname) {
                self.result_used = true;
                let (v, _) = self.expr(value)?;
                return Ok(format!("{result} = {v}"));
            }
        }
        // `lst.Selected(i) = True`: RAS cannot assign to a property that takes arguments.
        if matches!(target, Expr::Call(callee, _) if matches!(**callee, Expr::Member(..))) {
            return Err("引数を取るプロパティへの代入には対応していません（メソッドがあればそれを使ってください）".into());
        }
        let t = self.lvalue(target)?;
        // VBA's True is -1 when it is stored in a number (`Cancel = True` in QueryClose).
        if let (Some(var), Expr::Bool(b)) = (self.var_of(target), value) {
            if var.is_numeric() {
                return Ok(format!("{t} = {}", if *b { "-1" } else { "0" }));
            }
        }
        let (v, _) = self.value_expr(value)?;
        Ok(format!("{t} = {v}"))
    }

    /// The variable an expression names.
    fn var_of(&self, e: &Expr) -> Option<VarInfo> {
        let Expr::Ident(n, _) = e else { return None };
        let lower = n.to_lowercase();
        self.locals
            .get(&lower)
            .or_else(|| self.module_vars.get(&lower))
            .cloned()
    }

    /// An assignment target; a control assigned directly gets its default member `value`.
    fn lvalue(&mut self, e: &Expr) -> Result<String, String> {
        let (t, _) = self.expr(e)?;
        if self.is_control_ref(e) {
            return Ok(format!("{t}.value"));
        }
        Ok(t)
    }

    /// A statement of the files opened by number as a call of its RAS function (spec §15.3.2).
    fn file_stmt(
        &mut self,
        func: &str,
        args: &[FileArg],
        result: Option<&Expr>,
    ) -> Result<String, String> {
        let mut parts = Vec::new();
        for a in args {
            parts.push(match a {
                FileArg::Value(e) => self.value_expr(e)?.0,
                FileArg::Target(e) => self.lvalue(e)?,
                FileArg::Word(w) => (*w).to_string(),
            });
        }
        let call = format!("{func}({})", parts.join(", "));
        match result {
            Some(target) => Ok(format!("{} = {call}", self.lvalue(target)?)),
            None => Ok(call),
        }
    }

    fn call_stmt(&mut self, e: &Expr) -> Result<String, String> {
        match e {
            Expr::Call(..) => Ok(self.expr(e)?.0),
            Expr::Ident(name, _) => {
                let lower = name.to_lowercase();
                // `Unload Me` style statements were parsed as calls with arguments.
                if let Some(p) = self.procs.get(&lower) {
                    return Ok(format!("{}()", p.name));
                }
                let (t, _) = self.expr(e)?;
                Ok(if t.ends_with(')') {
                    t
                } else {
                    format!("{t}()")
                })
            }
            Expr::Member(..) => {
                let (t, _) = self.expr(e)?;
                Ok(format!("{t}()"))
            }
            _ => Err("この文には対応していません".into()),
        }
    }

    // ----------------------------------------------------------------------------- expressions

    fn cond(&mut self, e: &Expr) -> Result<String, String> {
        Ok(self.cond_expr(e)?.0)
    }

    /// A condition: a control stands for its value (a CheckBox).
    fn cond_expr(&mut self, e: &Expr) -> EResult {
        self.value_expr(e)
    }

    /// An expression used as a value: a bare control is its `value`.
    fn value_expr(&mut self, e: &Expr) -> EResult {
        let (t, p) = self.expr(e)?;
        if self.is_control_ref(e) {
            return Ok((format!("{t}.value"), prec::PRIMARY));
        }
        Ok((t, p))
    }

    /// A control, or a `MSForms.Return…` argument: used as a value, it stands for its `Value`.
    fn is_control_ref(&self, e: &Expr) -> bool {
        if self.var_of(e).is_some_and(|v| v.has_value_member()) {
            return true;
        }
        match e {
            Expr::Ident(n, _) => {
                !self.locals.contains_key(&n.to_lowercase())
                    && self.controls.contains_key(&n.to_lowercase())
            }
            Expr::Member(Some(o), n) => {
                matches!(**o, Expr::Me) && self.controls.contains_key(&n.to_lowercase())
            }
            Expr::Member(None, n) => {
                self.with_stack.last().is_some_and(|w| w == "this")
                    && self.controls.contains_key(&n.to_lowercase())
            }
            _ => false,
        }
    }

    fn operand(&mut self, e: &Expr, min: u8) -> Result<String, String> {
        let (t, p) = self.value_expr(e)?;
        Ok(if p < min { format!("({t})") } else { t })
    }

    fn is_boolean_like(e: &Expr) -> bool {
        match e {
            Expr::Bool(_) => true,
            Expr::Binary(op, ..) => matches!(
                op.as_str(),
                "=" | "<>" | "<" | ">" | "<=" | ">=" | "is" | "like" | "and" | "or" | "not"
            ),
            Expr::Unary("not", _) => true,
            Expr::Paren(inner) => Self::is_boolean_like(inner),
            _ => false,
        }
    }

    /// Known to be a floating-point value (then `/` divides as in VBA without a cast).
    fn is_float(&self, e: &Expr) -> bool {
        match e {
            Expr::Num(_, f) => *f,
            Expr::Paren(inner) | Expr::Unary("-", inner) => self.is_float(inner),
            Expr::Binary(op, a, b) if matches!(op.as_str(), "+" | "-" | "*" | "/" | "^") => {
                op == "/" || self.is_float(a) || self.is_float(b)
            }
            _ => self
                .var_of(e)
                .is_some_and(|v| matches!(v.ty.as_str(), "double" | "float" | "currency")),
        }
    }

    fn is_number(e: &Expr) -> bool {
        match e {
            Expr::Num(..) => true,
            Expr::Unary("-", inner) | Expr::Paren(inner) => Self::is_number(inner),
            _ => false,
        }
    }

    fn expr(&mut self, e: &Expr) -> EResult {
        use prec::*;
        Ok(match e {
            Expr::Num(n, _) => (n.clone(), PRIMARY),
            Expr::Str(s) => (ras_string(s), PRIMARY),
            Expr::Date(d) => {
                return Err(format!(
                    "日付リテラル #{d}# には対応していません（dateSerial などで書いてください）"
                ))
            }
            Expr::Bool(b) => ((if *b { "true" } else { "false" }).into(), PRIMARY),
            Expr::Nothing => ("null".into(), PRIMARY),
            Expr::Empty => return Err("Empty には対応していません".into()),
            Expr::Null => return Err("Null には対応していません".into()),
            Expr::Me => match self.self_ref() {
                Some(me) => (me, PRIMARY),
                None => return Err("標準モジュールの Me には対応していません".into()),
            },
            Expr::Ident(name, _) => (self.name_ref(name, false), PRIMARY),
            Expr::Member(obj, name) => (self.member(obj.as_deref(), name, false)?, PRIMARY),
            Expr::Bang => return Err("`!` によるメンバー参照には対応していません".into()),
            Expr::Call(callee, args) => (self.call(callee, args)?, PRIMARY),
            Expr::Paren(inner) => {
                let (t, _) = self.expr(inner)?;
                (format!("({t})"), PRIMARY)
            }
            Expr::New(t) => {
                let ty = self.type_of(t);
                if ty == "object" {
                    return Err(format!("New {t} には対応していません"));
                }
                (format!("new {ty}()"), PRIMARY)
            }
            Expr::TypeOf => return Err("TypeOf … Is には対応していません".into()),
            Expr::Unary(op, inner) => match *op {
                "-" => (format!("-{}", self.operand(inner, UNARY)?), UNARY),
                _ => {
                    // `Not`: logical unless applied to a number.
                    if Self::is_number(inner) {
                        (format!("~{}", self.operand(inner, UNARY)?), UNARY)
                    } else {
                        (format!("!{}", self.operand(inner, UNARY)?), UNARY)
                    }
                }
            },
            Expr::Binary(op, a, b) => match op.as_str() {
                "&" => (self.concat(e)?, PRIMARY),
                "+" | "-" => (
                    format!(
                        "{} {op} {}",
                        self.operand(a, ADD)?,
                        self.operand(b, ADD + 1)?
                    ),
                    ADD,
                ),
                "*" => (
                    format!("{} * {}", self.operand(a, MUL)?, self.operand(b, MUL + 1)?),
                    MUL,
                ),
                "/" => {
                    // VBA's `/` always divides as Double; RAS divides integers as integers.
                    let float = self.is_float(a) || self.is_float(b);
                    if float {
                        (
                            format!("{} / {}", self.operand(a, MUL)?, self.operand(b, MUL + 1)?),
                            MUL,
                        )
                    } else {
                        (
                            format!(
                                "({} as double) / {}",
                                self.operand(a, CAST)?,
                                self.operand(b, MUL + 1)?
                            ),
                            MUL,
                        )
                    }
                }
                "\\" => (
                    format!(
                        "({} as int) / ({} as int)",
                        self.operand(a, CAST)?,
                        self.operand(b, CAST)?
                    ),
                    MUL,
                ),
                "mod" => (
                    format!("{} % {}", self.operand(a, MUL)?, self.operand(b, MUL + 1)?),
                    MUL,
                ),
                "^" => (
                    format!("{} ** {}", self.operand(a, POW + 1)?, self.operand(b, POW)?),
                    POW,
                ),
                "=" | "is" => (
                    format!("{} == {}", self.operand(a, REL)?, self.operand(b, REL)?),
                    EQ,
                ),
                "<>" => (
                    format!("{} != {}", self.operand(a, REL)?, self.operand(b, REL)?),
                    EQ,
                ),
                "<" | ">" | "<=" | ">=" => (
                    format!("{} {op} {}", self.operand(a, ADD)?, self.operand(b, ADD)?),
                    REL,
                ),
                "and" | "or" | "xor" => {
                    let bitwise = (Self::is_number(a) || Self::is_number(b))
                        && !Self::is_boolean_like(a)
                        && !Self::is_boolean_like(b);
                    match (op.as_str(), bitwise) {
                        ("and", false) => (
                            format!("{} && {}", self.operand(a, AND)?, self.operand(b, AND + 1)?),
                            AND,
                        ),
                        ("or", false) => (
                            format!("{} || {}", self.operand(a, OR)?, self.operand(b, OR + 1)?),
                            OR,
                        ),
                        ("xor", false) => (
                            format!("{} != {}", self.operand(a, REL)?, self.operand(b, REL)?),
                            EQ,
                        ),
                        ("and", true) => (
                            format!(
                                "{} & {}",
                                self.operand(a, BITAND)?,
                                self.operand(b, BITAND + 1)?
                            ),
                            BITAND,
                        ),
                        ("or", true) => (
                            format!(
                                "{} | {}",
                                self.operand(a, BITOR)?,
                                self.operand(b, BITOR + 1)?
                            ),
                            BITOR,
                        ),
                        _ => (
                            format!(
                                "{} ^ {}",
                                self.operand(a, BITXOR)?,
                                self.operand(b, BITXOR + 1)?
                            ),
                            BITXOR,
                        ),
                    }
                }
                "like" => return Err("Like 演算子には対応していません".into()),
                _ => return Err(format!("{op} 演算子には対応していません")),
            },
        })
    }

    /// `a & b & c` as a template literal.
    fn concat(&mut self, e: &Expr) -> Result<String, String> {
        fn parts<'e>(e: &'e Expr, out: &mut Vec<&'e Expr>) {
            match e {
                Expr::Binary(op, a, b) if op == "&" => {
                    parts(a, out);
                    parts(b, out);
                }
                other => out.push(other),
            }
        }
        let mut list = Vec::new();
        parts(e, &mut list);
        let mut s = String::from("`");
        for p in list {
            match p {
                Expr::Str(text) => {
                    for ch in text.chars() {
                        match ch {
                            '`' => s.push_str("\\`"),
                            '\\' => s.push_str("\\\\"),
                            '$' => s.push_str("\\$"),
                            '\n' => s.push_str("\\n"),
                            '\r' => s.push_str("\\r"),
                            '\t' => s.push_str("\\t"),
                            other => s.push(other),
                        }
                    }
                }
                other => {
                    let (t, _) = self.value_expr(other)?;
                    s.push_str(&format!("${{{t}}}"));
                }
            }
        }
        s.push('`');
        Ok(s)
    }

    /// A name used on its own. `called`: arguments follow.
    fn name_ref(&mut self, name: &str, called: bool) -> String {
        let lower = name.to_lowercase();
        if let Some(v) = self.locals.get(&lower) {
            return v.name.clone();
        }
        if let Some((fname, result)) = &self.result {
            if *fname == lower && !called {
                self.result_used = true;
                return result.clone();
            }
        }
        if let Some(c) = self.controls.get(&lower) {
            return format!("this.{c}");
        }
        if let Some(v) = self.module_vars.get(&lower) {
            return v.name.clone();
        }
        if let Some(p) = self.procs.get(&lower) {
            return if called || !p.is_function {
                p.name.clone()
            } else {
                format!("{}()", p.name)
            };
        }
        if let Some(m) = self.enum_members.get(&lower) {
            return m.clone();
        }
        if let Target::Form(form) = self.target {
            if lower == form.name.to_lowercase() {
                return "this".into();
            }
        }
        // In a document module, the members of its sheet (workbook) come before the globals.
        if let (Some(members), Some(me)) = (self.document_members(), self.self_ref()) {
            if let Some((ras, is_fn)) = members.get(&lower) {
                return if *is_fn && !called {
                    format!("{me}.{ras}()")
                } else {
                    format!("{me}.{ras}")
                };
            }
        }
        if lower == "vbnullstring" {
            return "\"\"".into();
        }
        if let Some(name) = self.peer_name(&lower) {
            return if name.is_function && !called {
                format!("{}()", name.ras)
            } else {
                name.ras
            };
        }
        if let Some((ras, is_fn, _)) = self.host.globals.get(&lower) {
            return if *is_fn && !called {
                format!("{ras}()")
            } else {
                ras.clone()
            };
        }
        Self::user_ident(name)
    }

    fn member(&mut self, obj: Option<&Expr>, name: &str, called: bool) -> Result<String, String> {
        let lower = name.to_lowercase();
        // `Module1.Foo`: a public name of another module that is converted too.
        if let Some(Expr::Ident(module, _)) = obj {
            let shadowed = {
                let m = module.to_lowercase();
                self.locals.contains_key(&m)
                    || self.module_vars.contains_key(&m)
                    || self.procs.contains_key(&m)
            };
            if !shadowed {
                if let Some((alias, n)) = self.peer_member(module, &lower) {
                    return Ok(if n.is_function && !called {
                        format!("{alias}.{}()", n.ras)
                    } else {
                        format!("{alias}.{}", n.ras)
                    });
                }
            }
        }
        let receiver = match obj {
            Some(o) => Some(self.expr(o)?.0),
            None => match self.with_stack.last() {
                Some(w) => Some(w.clone()),
                None => return Err(format!("With の外の .{name} には対応していません")),
            },
        };
        let receiver = receiver.unwrap_or_default();
        if Some(&receiver) == self.self_ref().as_ref() {
            if let Some(c) = self.controls.get(&lower) {
                return Ok(format!("this.{c}"));
            }
            if let Some(p) = self.procs.get(&lower).cloned() {
                // Private procedures are called directly; exported ones are members of the form.
                let form = !matches!(self.target, Target::Standard);
                let base = if p.exported && form {
                    format!("this.{}", p.name)
                } else {
                    p.name.clone()
                };
                return Ok(if called || !p.is_function {
                    base
                } else {
                    format!("{base}()")
                });
            }
            if let Some(v) = self.module_vars.get(&lower) {
                return Ok(v.name.clone());
            }
        }
        if self.in_catch && receiver == "err" {
            match lower.as_str() {
                "number" => return Ok("e.code".into()),
                "description" => return Ok("e.description".into()),
                "source" => return Ok("e.source".into()),
                _ => {}
            }
        }
        let wrapped = if receiver.chars().all(|c| {
            c.is_alphanumeric()
                || c == '_'
                || c == '.'
                || c == '('
                || c == ')'
                || c == '['
                || c == ']'
                || c == '"'
                || c == ','
                || c == ' '
        }) && !receiver.starts_with('(')
        {
            receiver
        } else {
            format!("({receiver})")
        };
        Ok(format!("{wrapped}.{}", names::js_name(name)))
    }

    fn call(&mut self, callee: &Expr, args: &[Arg]) -> Result<String, String> {
        // Arrays: `a(i)` → `a[i]`, `a(i, j)` → `a[i][j]`.
        if let Expr::Ident(name, _) = callee {
            let lower = name.to_lowercase();
            let var = self
                .locals
                .get(&lower)
                .or_else(|| self.module_vars.get(&lower))
                .cloned();
            if let Some(v) = var {
                if v.indexable() && args.iter().all(|a| a.name.is_none() && a.value.is_some()) {
                    let mut s = v.name.clone();
                    for a in args {
                        let (t, _) = self.value_expr(a.value.as_ref().unwrap())?;
                        s.push_str(&format!("[{t}]"));
                    }
                    return Ok(s);
                }
            }
        }
        // The functions of the files opened by number that RAS writes differently.
        if let Expr::Ident(name, _) = callee {
            let lower = name.to_lowercase();
            if !self.procs.contains_key(&lower) {
                let value = |cv: &mut Self, i: usize| -> Result<String, String> {
                    match args.get(i).and_then(|a| a.value.as_ref()) {
                        Some(v) => Ok(cv.value_expr(v)?.0),
                        None => Err(format!("{name} の引数がありません")),
                    }
                };
                match lower.as_str() {
                    "input" if args.len() == 2 => {
                        let (count, n) = (value(self, 0)?, value(self, 1)?);
                        return Ok(format!("inputString({n}, {count})"));
                    }
                    "seek" if args.len() == 1 => return Ok(format!("seek({})", value(self, 0)?)),
                    "inputb" => return Err("InputB には対応していません".into()),
                    _ => {}
                }
            }
        }
        let f = match callee {
            Expr::Ident(name, _) => self.name_ref(name, true),
            Expr::Member(obj, name) => self.member(obj.as_deref(), name, true)?,
            other => self.expr(other)?.0,
        };
        let host_fn = matches!(callee, Expr::Ident(name, _) if !self.procs.contains_key(&name.to_lowercase()) && self.host.globals.contains_key(&name.to_lowercase()));
        let user_fn =
            matches!(callee, Expr::Ident(name, _) if self.procs.contains_key(&name.to_lowercase()));
        // Trailing omitted arguments are dropped.
        let mut n = args.len();
        while n > 0 && args[n - 1].value.is_none() {
            n -= 1;
        }
        let mut parts = Vec::new();
        for a in &args[..n] {
            let v = match &a.value {
                Some(v) => self.value_expr(v)?.0,
                None => "undefined".into(),
            };
            match &a.name {
                Some(name) if user_fn => parts.push(format!("{}: {v}", Self::user_ident(name))),
                Some(name) => parts.push(format!("{}: {v}", names::js_name(name))),
                None => parts.push(v),
            }
        }
        let _ = host_fn;
        Ok(format!("{f}({})", parts.join(", ")))
    }
}

/// A RAS string literal.
fn ras_string(s: &str) -> String {
    let mut out = String::from("\"");
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            other => out.push(other),
        }
    }
    out.push('"');
    out
}

/// Lower-cased names assigned anywhere in `stmts` (`x = …`, `Set x = …`, loop variables).
fn assigned_names(stmts: &[Stmt]) -> HashSet<String> {
    let mut out = HashSet::new();
    fn root(e: &Expr) -> Option<String> {
        match e {
            Expr::Ident(n, _) => Some(n.to_lowercase()),
            _ => None,
        }
    }
    fn walk(stmts: &[Stmt], out: &mut HashSet<String>) {
        for s in stmts {
            match &s.kind {
                StmtKind::Assign { target, .. } => out.extend(root(target)),
                StmtKind::File { args, result, .. } => {
                    out.extend(result.as_ref().and_then(root));
                    for a in args {
                        if let FileArg::Target(t) = a {
                            out.extend(root(t));
                        }
                    }
                }
                StmtKind::For { var, body, .. } => {
                    out.extend(root(var));
                    walk(body, out);
                }
                StmtKind::ForEach { var, body, .. } => {
                    out.insert(var.to_lowercase());
                    walk(body, out);
                }
                StmtKind::If { arms, els } => {
                    for (_, b) in arms {
                        walk(b, out);
                    }
                    if let Some(e) = els {
                        walk(e, out);
                    }
                }
                StmtKind::Select { cases, els, .. } => {
                    for (_, b) in cases {
                        walk(b, out);
                    }
                    if let Some(e) = els {
                        walk(e, out);
                    }
                }
                StmtKind::Do { body, .. } | StmtKind::With { body, .. } => walk(body, out),
                _ => {}
            }
        }
    }
    walk(stmts, &mut out);
    out
}

/// Turns `On Error` patterns of a procedure body into try / catch statements:
///
/// * `On Error Resume Next` … `On Error GoTo 0` (or the end of the block) → `try { … } catch { resume; }`
/// * `On Error GoTo L` … `Exit Sub` / `L:` … (end) → `try { … } catch (e) { … }`
fn restructure(stmts: &[Stmt]) -> Vec<Stmt> {
    let mut out: Vec<Stmt> = Vec::new();
    let mut i = 0;
    while i < stmts.len() {
        let s = &stmts[i];
        match &s.kind {
            StmtKind::OnErrorResumeNext => {
                let end = stmts[i + 1..]
                    .iter()
                    .position(|x| {
                        matches!(
                            x.kind,
                            StmtKind::OnErrorGoTo0
                                | StmtKind::OnErrorResumeNext
                                | StmtKind::OnErrorGoTo(_)
                        )
                    })
                    .map(|p| i + 1 + p)
                    .unwrap_or(stmts.len());
                let body = restructure(&stmts[i + 1..end]);
                out.push(Stmt {
                    kind: StmtKind::With {
                        obj: Expr::Ident(TRY_RESUME.into(), None),
                        body,
                    },
                    ..s.clone()
                });
                i = if end < stmts.len() && matches!(stmts[end].kind, StmtKind::OnErrorGoTo0) {
                    end + 1
                } else {
                    end
                };
            }
            StmtKind::OnErrorGoTo(label) => {
                let Some(at) = stmts[i + 1..]
                    .iter()
                    .position(
                        |x| matches!(&x.kind, StmtKind::Label(l) if l.eq_ignore_ascii_case(label)),
                    )
                    .map(|p| i + 1 + p)
                else {
                    out.push(Stmt {
                        kind: StmtKind::Todo(format!("On Error GoTo {label} の飛び先がありません")),
                        ..s.clone()
                    });
                    i += 1;
                    continue;
                };
                let mut body: Vec<Stmt> = stmts[i + 1..at].to_vec();
                // The `Exit Sub` before the handler.
                if let Some(last) = body
                    .iter()
                    .rposition(|x| !matches!(x.kind, StmtKind::Nothing))
                {
                    if matches!(&body[last].kind, StmtKind::Exit(w) if w == "sub" || w == "function" || w == "property")
                    {
                        body.remove(last);
                    }
                }
                let mut handler: Vec<Stmt> = stmts[at + 1..].to_vec();
                for h in handler.iter_mut() {
                    if let StmtKind::Todo(reason) = &h.kind {
                        if reason.starts_with("Resume") {
                            h.kind = StmtKind::Todo("エラー処理の Resume には対応していません（catch の後に処理が続きます）".into());
                        }
                    }
                }
                handler.retain(|h| !matches!(h.kind, StmtKind::Label(_)));
                out.push(Stmt {
                    kind: StmtKind::With {
                        obj: Expr::Ident(TRY_CATCH.into(), None),
                        body: restructure(&body),
                    },
                    ..s.clone()
                });
                out.push(Stmt {
                    kind: StmtKind::With {
                        obj: Expr::Ident(CATCH.into(), None),
                        body: restructure(&handler),
                    },
                    line: stmts[at].line,
                    comment: None,
                    blank_before: false,
                });
                i = stmts.len();
            }
            _ => {
                let mut s = s.clone();
                s.kind = match s.kind {
                    StmtKind::If { arms, els } => StmtKind::If {
                        arms: arms
                            .into_iter()
                            .map(|(c, b)| (c, restructure(&b)))
                            .collect(),
                        els: els.map(|e| restructure(&e)),
                    },
                    StmtKind::For {
                        var,
                        from,
                        to,
                        step,
                        body,
                    } => StmtKind::For {
                        var,
                        from,
                        to,
                        step,
                        body: restructure(&body),
                    },
                    StmtKind::ForEach { var, coll, body } => StmtKind::ForEach {
                        var,
                        coll,
                        body: restructure(&body),
                    },
                    StmtKind::Do { pre, post, body } => StmtKind::Do {
                        pre,
                        post,
                        body: restructure(&body),
                    },
                    StmtKind::With { obj, body } => StmtKind::With {
                        obj,
                        body: restructure(&body),
                    },
                    StmtKind::Select {
                        subject,
                        cases,
                        els,
                    } => StmtKind::Select {
                        subject,
                        cases: cases
                            .into_iter()
                            .map(|(c, b)| (c, restructure(&b)))
                            .collect(),
                        els: els.map(|e| restructure(&e)),
                    },
                    other => other,
                };
                out.push(s);
                i += 1;
            }
        }
    }
    out
}

// Markers for the try statements made by `restructure` (reusing `With` nodes).
const TRY_RESUME: &str = "\u{0}try_resume";
const TRY_CATCH: &str = "\u{0}try_catch";
const CATCH: &str = "\u{0}catch";

impl Emitter<'_> {
    /// `With` nodes that stand for try statements.
    fn try_stmt(&mut self, s: &Stmt) -> bool {
        let StmtKind::With {
            obj: Expr::Ident(marker, None),
            body,
        } = &s.kind
        else {
            return false;
        };
        match marker.as_str() {
            TRY_RESUME => {
                self.line("try {");
                self.indent += 1;
                self.block(body);
                self.indent -= 1;
                self.line("} catch {");
                self.line("    resume;");
                self.line("}");
            }
            TRY_CATCH => {
                self.line("try {");
                self.indent += 1;
                self.block(body);
                self.indent -= 1;
            }
            CATCH => {
                self.line("} catch (e) {");
                self.indent += 1;
                let saved = std::mem::replace(&mut self.in_catch, true);
                self.block(body);
                self.in_catch = saved;
                self.indent -= 1;
                self.line("}");
            }
            _ => return false,
        }
        true
    }
}
