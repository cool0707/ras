//! A small VBA abstract syntax tree and its pretty-printer.
//!
//! The lowering pass builds these trees; the printer takes care of operator
//! precedence, indentation, optional `Erl` line numbers and the source map.

use std::fmt::Write;

#[derive(Clone, Debug, PartialEq)]
pub enum VExpr {
    /// A literal or constant, printed verbatim (`1&`, `"abc"`, `True`, `Nothing`).
    Lit(String),
    Name(String),
    Member(Box<VExpr>, String),
    /// Function call, parameterised property access or array indexing.
    Call(Box<VExpr>, Vec<VExpr>),
    Not(Box<VExpr>),
    Neg(Box<VExpr>),
    Bin(&'static str, Box<VExpr>, Box<VExpr>),
    /// An omitted optional argument (`f(a, , c)`).
    Missing,
}

pub fn lit(s: impl Into<String>) -> VExpr {
    VExpr::Lit(s.into())
}

pub fn name(s: impl Into<String>) -> VExpr {
    VExpr::Name(s.into())
}

pub fn call(f: impl Into<String>, args: Vec<VExpr>) -> VExpr {
    VExpr::Call(Box::new(VExpr::Name(f.into())), args)
}

pub fn bin(op: &'static str, l: VExpr, r: VExpr) -> VExpr {
    VExpr::Bin(op, Box::new(l), Box::new(r))
}

pub fn long(n: i64) -> VExpr {
    VExpr::Lit(format!("{n}&"))
}

impl VExpr {
    pub fn member(self, name: impl Into<String>) -> VExpr {
        VExpr::Member(Box::new(self), name.into())
    }

    /// Array element / default-member access; literal indices are printed bare.
    pub fn index(self, args: Vec<VExpr>) -> VExpr {
        VExpr::Call(
            Box::new(self),
            args.into_iter().map(VExpr::bare_index).collect(),
        )
    }

    pub fn not(self) -> VExpr {
        match self {
            VExpr::Not(inner) => *inner,
            VExpr::Bin(op, l, r) if negate_cmp(op).is_some() => {
                VExpr::Bin(negate_cmp(op).unwrap(), l, r)
            }
            VExpr::Lit(s) if s == "True" => lit("False"),
            VExpr::Lit(s) if s == "False" => lit("True"),
            other => VExpr::Not(Box::new(other)),
        }
    }

    pub fn is_lit(&self, text: &str) -> bool {
        matches!(self, VExpr::Lit(s) if s == text)
    }

    /// If this is an integer literal such as `5&` (or a bare `5`), returns its value.
    pub fn as_long(&self) -> Option<i64> {
        match self {
            VExpr::Lit(s) => s.strip_suffix('&').unwrap_or(s).parse().ok(),
            _ => None,
        }
    }

    /// Array bounds and indices are printed without the `&` suffix (`arr(0)`, `ReDim a(0 To 2)`).
    pub fn bare_index(self) -> VExpr {
        match self.as_long() {
            Some(v) => VExpr::Lit(v.to_string()),
            None => self,
        }
    }

    /// `self + n`, folding integer literals (`i + 1& - 1&` → `i`).
    pub fn plus(self, n: i64) -> VExpr {
        if n == 0 {
            return self;
        }
        if let Some(v) = self.as_long() {
            return long(v + n);
        }
        if let VExpr::Bin(op @ ("+" | "-"), l, r) = &self {
            if let Some(k) = r.as_long() {
                let total = if *op == "+" { k + n } else { n - k };
                return (**l).clone().plus(total);
            }
        }
        if n > 0 {
            bin("+", self, long(n))
        } else {
            bin("-", self, long(-n))
        }
    }

    fn prec(&self) -> u8 {
        match self {
            VExpr::Lit(s) if s.starts_with('-') => 12,
            VExpr::Lit(_)
            | VExpr::Name(_)
            | VExpr::Member(..)
            | VExpr::Call(..)
            | VExpr::Missing => 14,
            VExpr::Neg(_) => 12,
            VExpr::Not(_) => 5,
            VExpr::Bin(op, ..) => bin_prec(op),
        }
    }

    pub fn print(&self) -> String {
        let mut s = String::new();
        self.write(&mut s);
        s
    }

    fn write_operand(&self, out: &mut String, parens: bool) {
        if parens {
            out.push('(');
            self.write(out);
            out.push(')');
        } else {
            self.write(out);
        }
    }

    fn write(&self, out: &mut String) {
        match self {
            VExpr::Lit(s) | VExpr::Name(s) => out.push_str(s),
            VExpr::Missing => {}
            VExpr::Member(obj, name) => {
                obj.write_operand(out, obj.prec() < 14);
                out.push('.');
                out.push_str(name);
            }
            VExpr::Call(f, args) => {
                f.write_operand(out, f.prec() < 14);
                out.push('(');
                write_args(out, args);
                out.push(')');
            }
            VExpr::Not(x) => {
                out.push_str("Not ");
                x.write_operand(out, x.prec() < 14);
            }
            VExpr::Neg(x) => {
                out.push('-');
                x.write_operand(out, x.prec() <= 12);
            }
            VExpr::Bin(op, l, r) => {
                let p = bin_prec(op);
                l.write_operand(out, l.prec() < p);
                out.push(' ');
                out.push_str(op);
                out.push(' ');
                r.write_operand(out, r.prec() <= p);
            }
        }
    }
}

fn write_args(out: &mut String, args: &[VExpr]) {
    // Trailing omitted arguments are dropped.
    let mut n = args.len();
    while n > 0 && args[n - 1] == VExpr::Missing {
        n -= 1;
    }
    for (i, a) in args[..n].iter().enumerate() {
        if i > 0 {
            out.push_str(if *a == VExpr::Missing { "," } else { ", " });
        }
        a.write(out);
    }
}

fn negate_cmp(op: &str) -> Option<&'static str> {
    Some(match op {
        "=" => "<>",
        "<>" => "=",
        "<" => ">=",
        ">=" => "<",
        ">" => "<=",
        "<=" => ">",
        _ => return None,
    })
}

fn bin_prec(op: &str) -> u8 {
    match op {
        "^" => 13,
        "*" | "/" => 11,
        "\\" => 10,
        "Mod" => 9,
        "+" | "-" => 8,
        "&" => 7,
        "=" | "<>" | "<" | ">" | "<=" | ">=" | "Is" | "Like" => 6,
        "And" => 4,
        "Or" => 3,
        "Xor" => 2,
        _ => 1,
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum VStmt {
    /// Marks the RAS source line the following statements originate from.
    Pos(u32),
    Assign {
        target: VExpr,
        value: VExpr,
        set: bool,
    },
    /// Statement call: `target arg1, arg2`.
    Call {
        target: VExpr,
        args: Vec<VExpr>,
    },
    If {
        cond: VExpr,
        then: Vec<VStmt>,
        elifs: Vec<(VExpr, Vec<VStmt>)>,
        els: Vec<VStmt>,
    },
    For {
        var: String,
        from: VExpr,
        to: VExpr,
        step: Option<VExpr>,
        body: Vec<VStmt>,
    },
    ForEach {
        var: String,
        coll: VExpr,
        body: Vec<VStmt>,
    },
    /// `Do [While pre] ... Loop [While post]`
    Do {
        pre: Option<VExpr>,
        post: Option<VExpr>,
        body: Vec<VStmt>,
    },
    Select {
        expr: VExpr,
        cases: Vec<(Vec<VExpr>, Vec<VStmt>)>,
        els: Option<Vec<VStmt>>,
    },
    Label(String),
    GoTo(String),
    /// `Exit Sub` / `Exit Function` / `Exit For` / `Exit Do` / `Exit Property`
    Exit(&'static str),
    ReDim {
        preserve: bool,
        target: VExpr,
        dims: Vec<(VExpr, VExpr)>,
    },
    Erase(VExpr),
    Comment(String),
    /// Any other single-line statement, printed verbatim (`On Error GoTo 0`).
    Raw(String),
}

pub fn assign(target: VExpr, value: VExpr) -> VStmt {
    VStmt::Assign {
        target,
        value,
        set: false,
    }
}

pub fn set(target: VExpr, value: VExpr) -> VStmt {
    VStmt::Assign {
        target,
        value,
        set: true,
    }
}

pub fn call_stmt(f: impl Into<String>, args: Vec<VExpr>) -> VStmt {
    VStmt::Call {
        target: VExpr::Name(f.into()),
        args,
    }
}

pub fn if_then(cond: VExpr, then: Vec<VStmt>) -> VStmt {
    VStmt::If {
        cond,
        then,
        elifs: Vec::new(),
        els: Vec::new(),
    }
}

pub fn if_else(cond: VExpr, then: Vec<VStmt>, els: Vec<VStmt>) -> VStmt {
    VStmt::If {
        cond,
        then,
        elifs: Vec::new(),
        els,
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Clone, Debug)]
pub struct VParam {
    pub name: String,
    pub byref: bool,
    pub ty: String,
    pub is_array: bool,
    pub optional_default: Option<String>,
}

#[derive(Clone, Debug)]
pub struct VProc {
    pub kind: ProcKind,
    pub public: bool,
    /// `Friend` instead of `Public` (class members whose signature mentions a `Type`).
    pub friend: bool,
    pub name: String,
    pub params: Vec<VParam>,
    /// Return type declaration (`As Long`, `As Long()`), without the `As`.
    pub ret: Option<String>,
    /// `Dim` declarations hoisted to the top: (name, type declaration, is array).
    pub locals: Vec<(String, String, bool)>,
    pub body: Vec<VStmt>,
    /// Lines printed before the procedure (doc comment).
    pub leading_comments: Vec<String>,
    /// `Attribute name.VB_Description = "..."` etc., printed right after the header.
    pub attributes: Vec<String>,
    /// RAS source file (index into the module's source list) for the source map.
    pub source: Option<usize>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum VModuleKind {
    Standard,
    Class,
    /// Code for a document module (`ThisWorkbook`, a sheet); merged into the
    /// RAS-managed region by sync / pack, never imported as a file.
    Document,
    /// Code module of a user form (spec §14.11): the whole code pane, written
    /// together with the form's designer storage by `ras pack` / `ras sync`.
    Form,
}

#[derive(Clone, Debug)]
pub struct VModule {
    pub name: String,
    pub kind: VModuleKind,
    /// `Option Private Module`
    pub private_module: bool,
    /// Class attributes (`VB_PredeclaredId` …) — (name, value).
    pub class_attrs: Vec<(String, String)>,
    /// `Implements X` lines (class modules).
    pub implements: Vec<String>,
    /// Module-level declarations (Declare, Const, variables, Type, Enum), verbatim blocks.
    pub decls: Vec<String>,
    pub procs: Vec<VProc>,
    /// RAS source files (relative paths) that contributed to this module.
    pub sources: Vec<String>,
    pub leading_comments: Vec<String>,
}

impl VModule {
    pub fn new(name: impl Into<String>, kind: VModuleKind) -> Self {
        VModule {
            name: name.into(),
            kind,
            private_module: false,
            class_attrs: Vec::new(),
            implements: Vec::new(),
            decls: Vec::new(),
            procs: Vec::new(),
            sources: Vec::new(),
            leading_comments: Vec::new(),
        }
    }
}

/// One `(vba line, source index, ras line)` triple; lines are 1-based.
pub type Mapping = (u32, u32, u32);

pub struct PrintOptions {
    /// Emit VBA line numbers equal to the RAS source line (`Erl` support, spec §14.8).
    pub erl: bool,
    /// Emit `Attribute` lines (file / offline output). Live sync omits them (spec §16.1.3).
    pub attributes: bool,
    /// Also target VBA6: declarations using `LongPtr` are emitted for both (spec §11.3).
    pub vba6: bool,
}

pub struct Printed {
    pub text: String,
    pub mappings: Vec<Mapping>,
}

struct Printer<'a> {
    out: String,
    line: u32,
    opts: &'a PrintOptions,
    mappings: Vec<Mapping>,
    cur_pos: Option<u32>,
    cur_source: u32,
    used_numbers: std::collections::HashSet<u32>,
}

/// Maximum line number VBA accepts as a numeric line label.
const MAX_LINE_NUMBER: u32 = 65529;

impl Printer<'_> {
    fn raw_line(&mut self, s: &str) {
        self.out.push_str(s);
        self.out.push_str("\r\n");
        self.line += 1;
    }

    /// Prints one statement line; `numberable` lines may carry an `Erl` number.
    fn stmt_line(&mut self, indent: usize, s: &str, numberable: bool) {
        let width = indent * 4;
        if let Some(pos) = self.cur_pos {
            self.mappings.push((self.line + 1, self.cur_source, pos));
            if numberable
                && self.opts.erl
                && pos <= MAX_LINE_NUMBER
                && self.used_numbers.insert(pos)
            {
                let num = pos.to_string();
                let pad = width.saturating_sub(num.len()).max(1);
                let _ = write!(self.out, "{num}{}", " ".repeat(pad));
                self.raw_line(s);
                return;
            }
        }
        let mut line = " ".repeat(width);
        line.push_str(s);
        self.raw_line(&line);
    }

    fn block(&mut self, indent: usize, stmts: &[VStmt]) {
        for s in stmts {
            self.stmt(indent, s);
        }
    }

    fn stmt(&mut self, indent: usize, s: &VStmt) {
        match s {
            VStmt::Pos(line) => self.cur_pos = Some(*line),
            VStmt::Assign { target, value, set } => {
                let text = format!(
                    "{}{} = {}",
                    if *set { "Set " } else { "" },
                    target.print(),
                    value.print()
                );
                self.stmt_line(indent, &text, true);
            }
            VStmt::Call { target, args } => {
                let mut text = target.print();
                let mut a = String::new();
                write_args(&mut a, args);
                // Every argument omitted: `wb.Close`, not `wb.Close `.
                if !a.is_empty() {
                    if a.starts_with('(') {
                        // `f (x), y` would be parsed as a parenthesised first argument.
                        text = format!("Call {text}({a})");
                    } else {
                        text.push(' ');
                        text.push_str(&a);
                    }
                }
                self.stmt_line(indent, &text, true);
            }
            VStmt::If {
                cond,
                then,
                elifs,
                els,
            } => {
                self.stmt_line(indent, &format!("If {} Then", cond.print()), true);
                self.block(indent + 1, then);
                for (c, body) in elifs {
                    self.stmt_line(indent, &format!("ElseIf {} Then", c.print()), false);
                    self.block(indent + 1, body);
                }
                if !els.is_empty() {
                    self.stmt_line(indent, "Else", false);
                    self.block(indent + 1, els);
                }
                self.stmt_line(indent, "End If", false);
            }
            VStmt::For {
                var,
                from,
                to,
                step,
                body,
            } => {
                let mut head = format!(
                    "For {var} = {} To {}",
                    from.clone().bare_index().print(),
                    to.clone().bare_index().print()
                );
                if let Some(st) = step {
                    let _ = write!(head, " Step {}", st.clone().bare_index().print());
                }
                self.stmt_line(indent, &head, true);
                self.block(indent + 1, body);
                self.stmt_line(indent, &format!("Next {var}"), false);
            }
            VStmt::ForEach { var, coll, body } => {
                self.stmt_line(indent, &format!("For Each {var} In {}", coll.print()), true);
                self.block(indent + 1, body);
                self.stmt_line(indent, &format!("Next {var}"), false);
            }
            VStmt::Do { pre, post, body } => {
                match pre {
                    Some(c) => self.stmt_line(indent, &format!("Do While {}", c.print()), true),
                    None => self.stmt_line(indent, "Do", true),
                }
                self.block(indent + 1, body);
                match post {
                    Some(c) => self.stmt_line(indent, &format!("Loop While {}", c.print()), false),
                    None => self.stmt_line(indent, "Loop", false),
                }
            }
            VStmt::Select { expr, cases, els } => {
                self.stmt_line(indent, &format!("Select Case {}", expr.print()), true);
                for (tests, body) in cases {
                    let list = tests
                        .iter()
                        .map(|t| t.print())
                        .collect::<Vec<_>>()
                        .join(", ");
                    self.stmt_line(indent, &format!("Case {list}"), false);
                    self.block(indent + 1, body);
                }
                if let Some(body) = els {
                    self.stmt_line(indent, "Case Else", false);
                    self.block(indent + 1, body);
                }
                self.stmt_line(indent, "End Select", false);
            }
            VStmt::Label(l) => {
                // Labels start in column 1.
                self.raw_line(&format!("{l}:"));
            }
            VStmt::GoTo(l) => self.stmt_line(indent, &format!("GoTo {l}"), true),
            VStmt::Exit(what) => self.stmt_line(indent, &format!("Exit {what}"), true),
            VStmt::ReDim {
                preserve,
                target,
                dims,
            } => {
                let d = dims
                    .iter()
                    .map(|(lo, hi)| {
                        format!(
                            "{} To {}",
                            lo.clone().bare_index().print(),
                            hi.clone().bare_index().print()
                        )
                    })
                    .collect::<Vec<_>>()
                    .join(", ");
                let p = if *preserve { "Preserve " } else { "" };
                self.stmt_line(indent, &format!("ReDim {p}{}({d})", target.print()), true);
            }
            VStmt::Erase(target) => {
                self.stmt_line(indent, &format!("Erase {}", target.print()), true)
            }
            VStmt::Comment(text) => {
                for l in text.lines() {
                    let saved = self.cur_pos.take();
                    self.stmt_line(indent, &format!("' {l}"), false);
                    self.cur_pos = saved;
                }
            }
            VStmt::Raw(text) => self.stmt_line(indent, text, !text.starts_with('#')),
        }
    }

    fn proc(&mut self, p: &VProc) {
        for c in &p.leading_comments {
            self.raw_line(&format!("' {c}"));
        }
        let kw = match p.kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            ProcKind::PropertyGet => "Property Get",
            ProcKind::PropertyLet => "Property Let",
            ProcKind::PropertySet => "Property Set",
        };
        let end_kw = match p.kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            _ => "Property",
        };
        let params = p
            .params
            .iter()
            .map(|a| {
                let mut s = String::new();
                if a.optional_default.is_some() {
                    s.push_str("Optional ");
                }
                s.push_str(if a.byref { "ByRef " } else { "ByVal " });
                s.push_str(&a.name);
                if a.is_array {
                    s.push_str("()");
                }
                let _ = write!(s, " As {}", a.ty);
                if let Some(d) = &a.optional_default {
                    let _ = write!(s, " = {d}");
                }
                s
            })
            .collect::<Vec<_>>()
            .join(", ");
        let vis = if p.friend {
            "Friend"
        } else if p.public {
            "Public"
        } else {
            "Private"
        };
        let mut head = format!("{vis} {kw} {}({params})", p.name);
        if let Some(r) = &p.ret {
            let _ = write!(head, " As {r}");
        }
        // `LongLong` only exists in 64-bit Office: the whole procedure is conditional.
        let win64_only = head.contains("LongLong");
        if win64_only {
            self.raw_line("#If Win64 Then");
        }
        if self.opts.vba6 && head.contains("LongPtr") {
            self.raw_line("#If VBA7 Then");
            self.raw_line(&head);
            self.raw_line("#Else");
            self.raw_line(&head.replace("LongPtr", "Long"));
            self.raw_line("#End If");
        } else {
            self.raw_line(&head);
        }
        if self.opts.attributes {
            for a in &p.attributes {
                self.raw_line(a);
            }
        }
        for (name, ty, is_array) in &p.locals {
            let arr = if *is_array { "()" } else { "" };
            if self.opts.vba6 && ty == "LongPtr" {
                self.raw_line("    #If VBA7 Then");
                self.raw_line(&format!("        Dim {name}{arr} As LongPtr"));
                self.raw_line("    #Else");
                self.raw_line(&format!("        Dim {name}{arr} As Long"));
                self.raw_line("    #End If");
            } else if ty == "LongLong" && !win64_only {
                self.raw_line("    #If Win64 Then");
                self.raw_line(&format!("        Dim {name}{arr} As LongLong"));
                self.raw_line("    #End If");
            } else {
                self.raw_line(&format!("    Dim {name}{arr} As {ty}"));
            }
        }
        if !p.locals.is_empty() && !p.body.is_empty() {
            self.raw_line("");
        }
        self.cur_pos = None;
        self.cur_source = p.source.unwrap_or(0) as u32;
        self.used_numbers.clear();
        self.block(1, &p.body);
        self.cur_pos = None;
        self.raw_line(&format!("End {end_kw}"));
        if win64_only {
            self.raw_line("#End If");
        }
    }
}

/// Header lines that exist only in exported files (`.bas` / `.cls`), not in a
/// VBE code pane.
pub fn file_header(m: &VModule) -> String {
    let mut s = String::new();
    match m.kind {
        VModuleKind::Standard | VModuleKind::Document | VModuleKind::Form => {}
        VModuleKind::Class => {
            s.push_str("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n");
        }
    }
    if !matches!(m.kind, VModuleKind::Document | VModuleKind::Form) {
        let _ = write!(s, "Attribute VB_Name = \"{}\"\r\n", m.name);
    }
    if m.kind == VModuleKind::Class {
        let get = |key: &str, default: &str| {
            m.class_attrs
                .iter()
                .find(|(k, _)| k == key)
                .map(|(_, v)| v.clone())
                .unwrap_or_else(|| default.to_string())
        };
        let _ = write!(
            s,
            "Attribute VB_GlobalNameSpace = {}\r\n",
            get("VB_GlobalNameSpace", "False")
        );
        let _ = write!(
            s,
            "Attribute VB_Creatable = {}\r\n",
            get("VB_Creatable", "False")
        );
        let _ = write!(
            s,
            "Attribute VB_PredeclaredId = {}\r\n",
            get("VB_PredeclaredId", "False")
        );
        let _ = write!(
            s,
            "Attribute VB_Exposed = {}\r\n",
            get("VB_Exposed", "False")
        );
    }
    s
}

/// Prints the code-pane part of a module (everything after the `Attribute` header).
pub fn print_module(m: &VModule, opts: &PrintOptions) -> Printed {
    let mut p = Printer {
        out: String::new(),
        line: 0,
        opts,
        mappings: Vec::new(),
        cur_pos: None,
        cur_source: 0,
        used_numbers: Default::default(),
    };
    for c in &m.leading_comments {
        p.raw_line(&format!("' {c}"));
    }
    // A document module is generated whole, like the others (spec §12.2); `Option Private Module`
    // cannot be written in an object module.
    p.raw_line("Option Explicit");
    p.raw_line("Option Compare Binary");
    if m.private_module && m.kind == VModuleKind::Standard {
        p.raw_line("Option Private Module");
    }
    for i in &m.implements {
        p.raw_line(&format!("Implements {i}"));
    }
    if !m.decls.is_empty() {
        if p.line > 0 {
            p.raw_line("");
        }
        for d in &m.decls {
            for l in d.lines() {
                p.raw_line(l);
            }
        }
    }
    for proc in &m.procs {
        if p.line > 0 {
            p.raw_line("");
        }
        p.proc(proc);
    }
    Printed {
        text: p.out,
        mappings: p.mappings,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn precedence_and_parens() {
        let e = bin("*", bin("+", name("a"), name("b")), name("c"));
        assert_eq!(e.print(), "(a + b) * c");
        let e = bin("-", name("a"), bin("-", name("b"), name("c")));
        assert_eq!(e.print(), "a - (b - c)");
        let e = bin(
            "And",
            bin("=", name("a"), long(1)),
            bin("<", name("b"), long(2)).not(),
        );
        assert_eq!(e.print(), "a = 1& And b >= 2&");
        let e = VExpr::Not(Box::new(call("f", vec![name("x")])));
        assert_eq!(e.print(), "Not f(x)");
        let e = VExpr::Not(Box::new(bin("And", name("a"), name("b"))));
        assert_eq!(e.print(), "Not (a And b)");
        let e = bin("^", VExpr::Neg(Box::new(name("a"))), long(2));
        assert_eq!(e.print(), "(-a) ^ 2&");
    }

    #[test]
    fn plus_folds_constants() {
        assert_eq!(name("i").plus(1).plus(-1).print(), "i");
        assert_eq!(name("n").plus(-1).print(), "n - 1&");
        assert_eq!(long(3).plus(2).print(), "5&");
        assert_eq!(bin("-", name("r"), long(2)).plus(2).print(), "r");
    }

    #[test]
    fn omitted_arguments() {
        let e = call(
            "Find",
            vec![
                lit("\"x\""),
                VExpr::Missing,
                name("xlValues"),
                VExpr::Missing,
            ],
        );
        assert_eq!(e.print(), "Find(\"x\",, xlValues)");
    }

    #[test]
    fn call_statement_avoids_paren_trap() {
        let opts = PrintOptions {
            erl: false,
            attributes: true,
            vba6: true,
        };
        let mut m = VModule::new("M", VModuleKind::Standard);
        m.procs.push(VProc {
            kind: ProcKind::Sub,
            public: true,
            friend: false,
            name: "T".into(),
            params: vec![],
            ret: None,
            locals: vec![],
            body: vec![
                call_stmt(
                    "foo",
                    vec![bin("*", bin("+", name("a"), name("b")), long(2)), name("c")],
                ),
                call_stmt("bar", vec![name("a")]),
            ],
            leading_comments: vec![],
            attributes: vec![],
            source: None,
        });
        let text = print_module(&m, &opts).text;
        assert!(text.contains("Call foo((a + b) * 2&, c)"), "{text}");
        assert!(text.contains("    bar a\r\n"), "{text}");
    }
}
