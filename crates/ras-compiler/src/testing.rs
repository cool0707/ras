//! Test files (spec §14.12): what they may contain, and their tests listed without running them.
//!
//! `describe` / `test` calls are only written at the top level of a test file and in the body
//! of a `describe`, with a string literal for the name, so the tree of tests is read from the
//! syntax alone (`ras test --list`).

use ras_syntax::ast::{self, ArrowBody, Expr, ExprKind, Item, Span, Stmt, StmtKind};

use crate::diag::{Diagnostic, Severity};
use crate::model::*;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Mode {
    Run,
    Skip,
    Only,
    Todo,
}

impl Mode {
    pub fn as_str(self) -> &'static str {
        match self {
            Mode::Run => "run",
            Mode::Skip => "skip",
            Mode::Only => "only",
            Mode::Todo => "todo",
        }
    }
}

/// A `describe` or a `test` of a test file.
#[derive(Clone, Debug)]
pub struct TestNode {
    pub suite: bool,
    pub name: String,
    /// 1-based line and column (UTF-16) of the call.
    pub line: u32,
    pub column: u32,
    pub mode: Mode,
    /// Made by `.each` from a table that is not an array literal: `name` is the template.
    pub template: bool,
    pub children: Vec<TestNode>,
}

#[derive(Clone, Debug)]
pub struct TestFile {
    pub module: ModuleId,
    /// Path relative to the project root.
    pub file: String,
    pub nodes: Vec<TestNode>,
}

impl TestFile {
    pub fn has_only(&self) -> bool {
        fn any(nodes: &[TestNode]) -> bool {
            nodes
                .iter()
                .any(|n| n.mode == Mode::Only || any(&n.children))
        }
        any(&self.nodes)
    }
}

/// What a call at the level of declarations is.
enum Decl<'a> {
    /// `describe(...)`, `test(...)` and their `.skip` / `.only` / `.todo` / `.each(table)(...)`.
    Node {
        suite: bool,
        mode: Mode,
        table: Option<&'a Expr>,
        args: &'a [Expr],
        span: Span,
    },
    /// `beforeAll(...)` and the other hooks.
    Hook,
}

fn base(e: &Expr) -> Option<bool> {
    match &e.peel().kind {
        ExprKind::Name(id) if id.text == "describe" => Some(true),
        ExprKind::Name(id) if id.text == "test" || id.text == "it" => Some(false),
        _ => None,
    }
}

fn classify(e: &Expr) -> Option<Decl<'_>> {
    let e = e.peel();
    let ExprKind::Call { callee, args, .. } = &e.kind else {
        return None;
    };
    let c = callee.peel();
    if let ExprKind::Name(id) = &c.kind {
        if matches!(
            id.text.as_str(),
            "beforeAll" | "afterAll" | "beforeEach" | "afterEach"
        ) {
            return Some(Decl::Hook);
        }
    }
    if let Some(suite) = base(c) {
        return Some(Decl::Node {
            suite,
            mode: Mode::Run,
            table: None,
            args,
            span: e.span,
        });
    }
    match &c.kind {
        ExprKind::Member { object, name, .. } => {
            let suite = base(object)?;
            let mode = match name.text.as_str() {
                "skip" => Mode::Skip,
                "only" => Mode::Only,
                "todo" if !suite => Mode::Todo,
                _ => return None,
            };
            Some(Decl::Node {
                suite,
                mode,
                table: None,
                args,
                span: e.span,
            })
        }
        // test.each(table)(name, body)
        ExprKind::Call {
            callee: inner,
            args: table,
            ..
        } => {
            let ExprKind::Member { object, name, .. } = &inner.peel().kind else {
                return None;
            };
            if name.text != "each" {
                return None;
            }
            let suite = base(object)?;
            Some(Decl::Node {
                suite,
                mode: Mode::Run,
                table: table.first(),
                args,
                span: e.span,
            })
        }
        _ => None,
    }
}

/// A literal as `%s` writes it.
fn literal_text(e: &Expr) -> Option<String> {
    match &e.peel().kind {
        ExprKind::Int { value, .. } => Some(value.to_string()),
        ExprKind::Float { text, .. } => Some(text.clone()),
        ExprKind::Str(s) => Some(s.clone()),
        ExprKind::Bool(b) => Some(if *b { "True" } else { "False" }.to_string()),
        ExprKind::Unary {
            op: ast::UnaryOp::Neg,
            operand,
        } => literal_text(operand).map(|t| format!("-{t}")),
        _ => None,
    }
}

/// The name of a case of `.each`: `%s` `%i` `%d` `%f` take the values of the row in order,
/// `%%` is a percent sign, `$field` the field of a row that is a struct. The run time
/// (`ras-test.js`) does the same with the values.
fn case_name(template: &str, row: &Expr) -> Option<String> {
    let (values, fields): (Vec<String>, Vec<(String, String)>) = match &row.peel().kind {
        ExprKind::Array(items) => (
            items.iter().map(literal_text).collect::<Option<_>>()?,
            Vec::new(),
        ),
        ExprKind::Object(props) => (
            Vec::new(),
            props
                .iter()
                .map(|p| literal_text(&p.value).map(|v| (p.key.text.clone(), v)))
                .collect::<Option<_>>()?,
        ),
        _ => (vec![literal_text(row)?], Vec::new()),
    };
    let mut out = String::new();
    let mut next = 0;
    let chars: Vec<char> = template.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '%' && i + 1 < chars.len() {
            match chars[i + 1] {
                '%' => out.push('%'),
                's' | 'i' | 'd' | 'f' => {
                    let v = values.get(next).cloned().unwrap_or_default();
                    next += 1;
                    out.push_str(&match chars[i + 1] {
                        'i' | 'd' => v.split('.').next().unwrap_or_default().to_string(),
                        _ => v,
                    });
                }
                other => {
                    out.push('%');
                    out.push(other);
                }
            }
            i += 2;
            continue;
        }
        if c == '$' {
            let name: String = chars[i + 1..]
                .iter()
                .take_while(|c| c.is_alphanumeric() || **c == '_')
                .collect();
            if let Some((_, v)) = fields.iter().find(|(k, _)| *k == name) {
                out.push_str(v);
                i += 1 + name.chars().count();
                continue;
            }
        }
        out.push(c);
        i += 1;
    }
    Some(out)
}

struct Walker<'p> {
    md: &'p ModuleData,
    module: ModuleId,
    diags: Vec<Diagnostic>,
}

impl Walker<'_> {
    fn err(&mut self, span: Span, code: &'static str, message: impl Into<String>) {
        self.diags.push(Diagnostic {
            severity: Severity::Error,
            code,
            message: message.into(),
            module: Some(self.module),
            span,
        });
    }

    /// A statement at the level of declarations: `out` gets the tests it declares.
    fn stmt(&mut self, s: &Stmt, top: bool, out: &mut Vec<TestNode>) {
        match &s.kind {
            StmtKind::Var(_) | StmtKind::Empty | StmtKind::Error => {}
            StmtKind::Expr(e) => self.expr(e, s.span, top, out),
            _ => self.not_allowed(s.span, top),
        }
    }

    fn not_allowed(&mut self, span: Span, top: bool) {
        let place = if top {
            "テストファイルのトップレベル"
        } else {
            "`describe` の本体"
        };
        self.err(span, "E0702", format!("{place}に書けるのは、`const` / `let` の宣言と `describe` / `test` / フックの呼び出しだけです。準備の処理は `beforeAll` / `beforeEach` かテストの本体に書いてください"));
    }

    fn expr(&mut self, e: &Expr, span: Span, top: bool, out: &mut Vec<TestNode>) {
        let Some(decl) = classify(e) else {
            return self.not_allowed(span, top);
        };
        let Decl::Node {
            suite,
            mode,
            table,
            args,
            span,
        } = decl
        else {
            return;
        };
        let what = if suite { "describe" } else { "test" };
        let name = match args.first().map(|a| &a.peel().kind) {
            Some(ExprKind::Str(s)) => s.clone(),
            Some(_) => {
                self.err(args[0].span, "E0703", format!("`{what}` の名前は文字列リテラルで書いてください（実行せずにテストを列挙するため、テンプレート文字列や変数は使えません）"));
                return;
            }
            None => return,
        };
        let (line, column) = self.md.line_col_utf16(u32::from(span.start()));
        let mut children = Vec::new();
        if suite {
            match args.get(1).map(|a| &a.peel().kind) {
                Some(ExprKind::Arrow(f)) => match &f.body {
                    ArrowBody::Block(b) => {
                        for s in &b.stmts {
                            self.stmt(s, false, &mut children);
                        }
                    }
                    ArrowBody::Expr(x) => self.expr(x, x.span, false, &mut children),
                },
                Some(_) => self.err(
                    args[1].span,
                    "E0702",
                    "`describe` の本体はラムダ（`() => { ... }`）で書いてください",
                ),
                None => {}
            }
        }
        let node = |name: String, template: bool, children: Vec<TestNode>| TestNode {
            suite,
            name,
            line: line + 1,
            column: column + 1,
            mode,
            template,
            children,
        };
        match table {
            None => out.push(node(name, false, children)),
            Some(t) => {
                let rows = match &t.peel().kind {
                    ExprKind::Array(rows) => rows
                        .iter()
                        .map(|r| case_name(&name, r))
                        .collect::<Option<Vec<String>>>(),
                    _ => None,
                };
                match rows {
                    Some(names) => {
                        out.extend(names.into_iter().map(|n| node(n, false, children.clone())))
                    }
                    // The cases are only known when the table is evaluated.
                    None => out.push(node(name, true, children)),
                }
            }
        }
    }
}

fn walk(prog: &Program, m: ModuleId) -> (TestFile, Vec<Diagnostic>) {
    let md = &prog.modules[m];
    let mut w = Walker {
        md,
        module: m,
        diags: Vec::new(),
    };
    let mut nodes = Vec::new();
    for item in &md.ast.items {
        match item {
            Item::Stmt(s) => w.stmt(s, true, &mut nodes),
            Item::Fn(f) => {
                for d in &f.decorators {
                    if matches!(
                        d.name.text.as_str(),
                        "macro" | "init" | "udf" | "on" | "ribbon"
                    ) {
                        w.err(d.span, "E0704", format!("テストファイルには `@{}` を書けません（テストはブックに入りません）", d.name.text));
                    }
                }
            }
            _ => {}
        }
    }
    (
        TestFile {
            module: m,
            file: md.rel_path.clone(),
            nodes,
        },
        w.diags,
    )
}

/// The tests of every test file, in the order of their paths.
pub fn list(prog: &Program) -> Vec<TestFile> {
    let mut files: Vec<TestFile> = (0..prog.modules.len())
        .filter(|&m| prog.is_test_file(m))
        .map(|m| walk(prog, m).0)
        .collect();
    files.sort_by(|a, b| a.file.cmp(&b.file));
    files
}

/// Reports what a test file may not contain (`E0702`–`E0704`).
pub fn check(prog: &mut Program) {
    for m in 0..prog.modules.len() {
        if prog.is_test_file(m) {
            let (_, diags) = walk(prog, m);
            prog.diagnostics.extend(diags);
        }
    }
}
