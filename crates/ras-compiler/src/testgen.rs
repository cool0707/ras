//! The skeleton of a test for a function (spec §14.12.10): what the editors' "テストを生成"
//! (the code lens of VS Code, the button of the playground) writes into `test/`.
//!
//! The test file of `src/<key>.ras` is `test/<key>.test.ras`. A new file gets an `import` of
//! the function and a `describe` with one `test`; in a file that already exists, the name is
//! added to its `import` (or a new `import` is written after the others), and the `test` goes
//! at the end of the `describe` named after the function (or a new `describe` at the end).
//!
//! The arguments are declared as constants with their types and a placeholder value, and the
//! result is compared with a placeholder of its type, so the test compiles as written and fails
//! until its values are filled in.

use std::collections::BTreeMap;

use ras_syntax::ast::{ArrowBody, ExprKind, Item, StmtKind};

use crate::model::*;

/// A change of the text of the test file, as offsets (bytes) into its current text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Edit {
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Clone, Debug)]
pub struct GeneratedTest {
    /// The test file, relative to the project root (`test/report.test.ras`).
    pub path: String,
    /// Whether the file is new (`edits` is then one insertion into an empty text).
    pub created: bool,
    /// The changes of the file's current text (an empty text when `created`).
    pub edits: Vec<Edit>,
    /// The text of the file after the edits.
    pub text: String,
    /// The line (0-based) of the new `test` in `text`.
    pub line: u32,
}

/// Whether the editors offer a test for `f`: an exported function of a module of `src/` with a
/// body, which a test file can `import` and call (spec §14.12.2).
pub fn can_generate(prog: &Program, f: FnId) -> bool {
    let d = &prog.fns[f];
    let md = &prog.modules[d.module];
    md.kind == ModuleKind::User
        && md.form.is_none()
        && !md.is_test()
        && md.rel_path.starts_with("src/")
        && d.kind == FnKind::Free
        && d.accessor.is_none()
        && d.exported
        && d.decl.as_ref().is_some_and(|x| x.body.is_some())
        && d.attrs.on_events.is_empty()
        && !d.attrs.init
}

/// The functions of a module the editors offer a test for, in declaration order.
pub fn targets(prog: &Program, m: ModuleId) -> Vec<FnId> {
    prog.modules[m]
        .fns
        .iter()
        .copied()
        .filter(|&f| can_generate(prog, f))
        .collect()
}

/// The test file of a module of `src/`.
pub fn test_path(md: &ModuleData) -> String {
    format!("test/{}.test.ras", md.key)
}

/// The skeleton of a test of `f` (see [`can_generate`]) for its test file, whose current text
/// is `existing` (`None` when the file does not exist yet).
pub fn generate(prog: &Program, f: FnId, existing: Option<&str>) -> Option<GeneratedTest> {
    if !can_generate(prog, f) {
        return None;
    }
    let d = &prog.fns[f];
    let md = &prog.modules[d.module];
    let path = test_path(md);
    let from = crate::collect::root_import_path(prog, d.module);

    let mut gen = Gen {
        prog,
        imports: BTreeMap::new(),
    };
    gen.import(d.module, &d.name);
    let body = gen.test_body(d);
    // The names of the module of the function first, then the other modules in order.
    let mut imports: Vec<(String, Vec<String>)> = Vec::new();
    for (m, names) in &gen.imports {
        let other = &prog.modules[*m];
        let src = if *m == d.module {
            from.clone()
        } else {
            crate::collect::root_import_path(prog, other.id)
        };
        let entry = (src, names.clone());
        if *m == d.module {
            imports.insert(0, entry);
        } else {
            imports.push(entry);
        }
    }

    let indent = "    ";
    let test_text = |level: usize| -> String {
        let pad = indent.repeat(level);
        let mut s = String::new();
        for line in &body {
            if line.is_empty() {
                s.push('\n');
            } else {
                s.push_str(&pad);
                s.push_str(line);
                s.push('\n');
            }
        }
        s
    };

    let Some(existing) = existing else {
        let mut text = String::new();
        for (src, names) in &imports {
            text.push_str(&format!(
                "import {{ {} }} from \"{src}\";\n",
                names.join(", ")
            ));
        }
        text.push('\n');
        text.push_str(&format!("describe({}, () => {{\n", quote(&d.name)));
        let line = text.matches('\n').count() as u32;
        text.push_str(&test_text(1));
        text.push_str("});\n");
        return Some(GeneratedTest {
            path,
            created: true,
            edits: vec![Edit {
                start: 0,
                end: 0,
                text: text.clone(),
            }],
            text,
            line,
        });
    };

    let ast = ras_syntax::parse(existing).ast;
    let mut edits = Vec::new();
    // The imports: added to an `import` of the same module, or written after the last one.
    let last_import = ast.items.iter().filter_map(|i| match i {
        Item::Import(im) => Some(im),
        _ => None,
    });
    let mut after_imports = None;
    for (src, names) in &imports {
        let same = ast.items.iter().find_map(|i| match i {
            Item::Import(im) if im.namespace.is_none() && root_path(&path, &im.source) == *src => {
                Some(im)
            }
            _ => None,
        });
        match same {
            Some(im) => {
                let missing: Vec<&String> = names
                    .iter()
                    .filter(|n| {
                        !im.names
                            .iter()
                            .any(|x| x.alias.is_none() && x.name.text == **n)
                    })
                    .collect();
                if missing.is_empty() {
                    continue;
                }
                let added = missing
                    .iter()
                    .map(|s| s.as_str())
                    .collect::<Vec<_>>()
                    .join(", ");
                match im.names.last() {
                    Some(n) => {
                        let at = u32::from(n.alias.as_ref().unwrap_or(&n.name).span.end());
                        edits.push(Edit {
                            start: at,
                            end: at,
                            text: format!(", {added}"),
                        });
                    }
                    None => {
                        // `import {} from "..."`: the braces are rewritten.
                        let start = u32::from(im.span.start());
                        let end = u32::from(im.source_span.start());
                        edits.push(Edit {
                            start,
                            end,
                            text: format!("import {{ {added} }} from "),
                        });
                    }
                }
            }
            None => {
                let at =
                    *after_imports.get_or_insert_with(|| match last_import.clone().next_back() {
                        Some(im) => line_end(existing, u32::from(im.span.end())),
                        None => 0,
                    });
                let line = format!("import {{ {} }} from \"{src}\";\n", names.join(", "));
                // After the last `import`, or at the top of a file without one.
                let text = if at > 0 && !existing[..at as usize].ends_with('\n') {
                    format!("\n{line}")
                } else if at == 0 && !existing.is_empty() && !existing.starts_with("import") {
                    format!("{line}\n")
                } else {
                    line
                };
                edits.push(Edit {
                    start: at,
                    end: at,
                    text,
                });
            }
        }
    }

    // The test: at the end of the `describe` named after the function, or in a new one.
    let describe = ast.items.iter().find_map(|i| {
        let Item::Stmt(s) = i else { return None };
        let StmtKind::Expr(e) = &s.kind else {
            return None;
        };
        let ExprKind::Call { callee, args, .. } = &e.peel().kind else {
            return None;
        };
        if !matches!(&callee.peel().kind, ExprKind::Name(id) if id.text == "describe") {
            return None;
        }
        if !matches!(args.first().map(|a| &a.peel().kind), Some(ExprKind::Str(s)) if *s == d.name) {
            return None;
        }
        let ExprKind::Arrow(arrow) = &args.get(1)?.peel().kind else {
            return None;
        };
        match &arrow.body {
            ArrowBody::Block(b) => Some(b.span),
            ArrowBody::Expr(_) => None,
        }
    });
    let test_edit = match describe {
        Some(block) => {
            // Before the line of the closing brace of the `describe`'s body.
            let close = u32::from(block.end()).saturating_sub(1);
            let at = line_start(existing, close);
            let inner = &existing[u32::from(block.start()) as usize + 1..close as usize];
            let text = if inner.trim().is_empty() {
                test_text(1)
            } else {
                format!("\n{}", test_text(1))
            };
            if at <= u32::from(block.start()) {
                // `describe("f", () => { ... })` on one line: the body is opened.
                Edit {
                    start: close,
                    end: close,
                    text: format!("\n{text}"),
                }
            } else {
                Edit {
                    start: at,
                    end: at,
                    text,
                }
            }
        }
        None => {
            let end = existing.len() as u32;
            let sep = if existing.is_empty() || existing.ends_with("\n\n") {
                ""
            } else if existing.ends_with('\n') {
                "\n"
            } else {
                "\n\n"
            };
            Edit {
                start: end,
                end,
                text: format!(
                    "{sep}describe({}, () => {{\n{}}});\n",
                    quote(&d.name),
                    test_text(1)
                ),
            }
        }
    };
    edits.push(test_edit.clone());
    edits.sort_by_key(|e| (e.start, e.end));
    let text = apply(existing, &edits);
    // The line of the new test: the edits before it shift it.
    let shift: usize = edits
        .iter()
        .filter(|e| e.start < test_edit.start)
        .map(|e| e.text.len() as isize - (e.end - e.start) as isize)
        .sum::<isize>() as usize;
    let at = test_edit.start as usize + shift;
    let lead = test_edit
        .text
        .find("test(")
        .map(|i| test_edit.text[..i].matches('\n').count())
        .unwrap_or(0);
    let line = (text[..at].matches('\n').count() + lead) as u32;
    Some(GeneratedTest {
        path,
        created: false,
        edits,
        text,
        line,
    })
}

/// The text with the edits (sorted, not overlapping) applied.
pub fn apply(text: &str, edits: &[Edit]) -> String {
    let mut out = String::with_capacity(text.len());
    let mut pos = 0usize;
    for e in edits {
        out.push_str(&text[pos..e.start as usize]);
        out.push_str(&e.text);
        pos = e.end as usize;
    }
    out.push_str(&text[pos..]);
    out
}

fn line_end(text: &str, at: u32) -> u32 {
    match text[at as usize..].find('\n') {
        Some(i) => at + i as u32 + 1,
        None => text.len() as u32,
    }
}

fn line_start(text: &str, at: u32) -> u32 {
    let before = &text[..at as usize];
    if !before[before.rfind('\n').map_or(0, |i| i + 1)..]
        .trim()
        .is_empty()
    {
        // Something precedes the brace on its line (`  x; }`): the test goes before the brace.
        return at;
    }
    before.rfind('\n').map_or(0, |i| i as u32 + 1)
}

/// The first sheet of the workbook: its code name, or `thisWorkbook.worksheets(1)` when the
/// project has no workbook.
fn sheet_expr(prog: &Program) -> String {
    prog.config
        .sheet_code_names
        .first()
        .cloned()
        .unwrap_or_else(|| "thisWorkbook.worksheets(1)".into())
}

fn quote(s: &str) -> String {
    let mut out = String::from('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            _ => out.push(c),
        }
    }
    out.push('"');
    out
}

struct Gen<'a> {
    prog: &'a Program,
    /// The names to import, by module (in the order of the modules).
    imports: BTreeMap<ModuleId, Vec<String>>,
}

impl Gen<'_> {
    fn import(&mut self, m: ModuleId, name: &str) {
        let names = self.imports.entry(m).or_default();
        if !names.iter().any(|n| n == name) {
            names.push(name.to_string());
        }
    }

    /// The name of a declaration of a user module as the test writes it, imported when needed.
    /// `None` when the test cannot name it (not exported).
    fn named(&mut self, m: ModuleId, name: &str) -> Option<String> {
        let md = &self.prog.modules[m];
        if md.kind != ModuleKind::User {
            return Some(name.to_string());
        }
        if !md.exports.contains_key(name) || !md.rel_path.starts_with("src/") {
            return None;
        }
        self.import(m, name);
        Some(name.to_string())
    }

    /// The type as the test writes it in an annotation (`None`: left to inference).
    fn annot(&mut self, t: &Type) -> Option<String> {
        let prog = self.prog;
        match t {
            Type::Error | Type::Void | Type::Null | Type::Ns(_) | Type::Param(_) => None,
            Type::Enum(id) => {
                let e = &prog.enums[*id];
                if e.host {
                    return Some(e.name.clone());
                }
                self.named(e.module, &e.name)
            }
            Type::Struct(id) => {
                let s = &prog.structs[*id];
                match s.kind {
                    StructKind::Named => self.named(s.module, &s.name),
                    _ => {
                        for (_, ft) in &s.fields {
                            self.annot(ft)?;
                        }
                        Some(prog.show(t))
                    }
                }
            }
            Type::Class(id, args) => {
                let c = &prog.classes[*id];
                if c.namespace.is_none() && c.kind != ClassKind::Host {
                    self.named(c.module, &c.name)?;
                }
                for a in args {
                    self.annot(a)?;
                }
                Some(prog.show(t))
            }
            Type::Array(e) => self.annot(e).map(|_| prog.show(t)),
            Type::Map(k, v) => {
                self.annot(k)?;
                self.annot(v)?;
                Some(prog.show(t))
            }
            Type::Union(ms) => {
                for m in ms {
                    if !matches!(m, Type::Null) {
                        self.annot(m)?;
                    }
                }
                Some(prog.show(t))
            }
            Type::Func(f) => {
                for p in &f.params {
                    self.annot(p)?;
                }
                if f.ret != Type::Void {
                    self.annot(&f.ret)?;
                }
                Some(prog.show(t))
            }
            _ => Some(prog.show(t)),
        }
    }

    /// A value of the type to start the test with.
    fn sample(&mut self, t: &Type) -> String {
        let prog = self.prog;
        match t {
            t if t.is_integer() => "0".into(),
            Type::Float | Type::Double | Type::Currency | Type::Decimal => "0.0".into(),
            Type::Str | Type::WStr | Type::AStr => "\"\"".into(),
            Type::Bool => "false".into(),
            Type::Date => "dateSerial(2000, 1, 1)".into(),
            Type::Array(_) => "[]".into(),
            Type::Map(..) => match self.annot(t) {
                Some(s) => format!("new {s}()"),
                None => "null".into(),
            },
            Type::Enum(id) => {
                let e = &prog.enums[*id];
                match (self.annot(t), e.members.first()) {
                    (Some(name), Some(m)) if !e.host => format!("{name}.{}", m.name),
                    (_, Some(m)) if e.host => m.name.clone(),
                    _ => "0".into(),
                }
            }
            Type::Struct(id) => {
                let s = &prog.structs[*id];
                match s.kind {
                    StructKind::Tuple => format!(
                        "[{}]",
                        s.fields
                            .iter()
                            .map(|(_, ft)| self.sample(ft))
                            .collect::<Vec<_>>()
                            .join(", ")
                    ),
                    _ => {
                        let fields: Vec<String> = s
                            .fields
                            .iter()
                            .zip(&s.optional)
                            .filter(|(_, o)| !**o)
                            .map(|((n, ft), _)| format!("{n}: {}", self.sample(ft)))
                            .collect();
                        if fields.is_empty() {
                            "{}".into()
                        } else {
                            format!("{{ {} }}", fields.join(", "))
                        }
                    }
                }
            }
            Type::Class(id, _) => {
                let c = &prog.classes[*id];
                match (c.namespace.as_deref(), c.name.as_str()) {
                    (Some("Excel"), "Worksheet") => sheet_expr(prog),
                    (Some("Excel"), "Range") => format!("{}.range(\"A1\")", sheet_expr(prog)),
                    (Some("Excel"), "Workbook") => "thisWorkbook".into(),
                    (Some("Excel"), "Application") => "application".into(),
                    _ => "null".into(),
                }
            }
            Type::Func(f) => {
                let params: Vec<String> = f
                    .params
                    .iter()
                    .enumerate()
                    .map(|(i, p)| match self.annot(p) {
                        Some(a) => format!("p{}: {a}", i + 1),
                        None => format!("p{}", i + 1),
                    })
                    .collect();
                let ret = if f.ret == Type::Void {
                    "{}".into()
                } else {
                    let r = f.ret.clone();
                    self.sample(&r)
                };
                format!("({}) => {ret}", params.join(", "))
            }
            Type::Union(ms) => match ms.iter().find(|m| !matches!(m, Type::Null)) {
                Some(m) => {
                    let m = m.clone();
                    self.sample(&m)
                }
                None => "null".into(),
            },
            _ => "null".into(),
        }
    }

    /// The statements of the `test`, one per line, without indentation.
    fn test_body(&mut self, d: &FnDef) -> Vec<String> {
        let mut lines = Vec::new();
        let mut decls = Vec::new();
        let mut args = Vec::new();
        let mut receiver = None;
        for (i, p) in d.params.iter().enumerate() {
            let is_this = i == 0 && d.is_extension;
            if !is_this && (p.optional || p.default.is_some() || p.rest) {
                continue;
            }
            let name = if is_this {
                "target".to_string()
            } else {
                p.name.clone()
            };
            let value = self.sample(&p.ty);
            let decl = match self.annot(&p.ty) {
                Some(a) => format!("const {name}: {a} = {value};"),
                None => format!("const {name} = {value};"),
            };
            decls.push(decl);
            if is_this {
                receiver = Some(name);
            } else {
                args.push(name);
            }
        }
        let call = match &receiver {
            Some(r) => format!("{r}.{}({})", d.name, args.join(", ")),
            None => format!("{}({})", d.name, args.join(", ")),
        };
        let ret = d.ret_ty();
        let has_args = !decls.is_empty();
        let name = if ret == Type::Void {
            format!("{} を実行する", d.name)
        } else {
            format!("{} の戻り値", d.name)
        };
        lines.push(format!("test({}, () => {{", quote(&name)));
        if has_args {
            lines.push("    // 準備: 引数の値を書き換える".into());
            for decl in decls {
                lines.push(format!("    {decl}"));
            }
            lines.push(String::new());
        }
        match &ret {
            Type::Void => {
                lines.push(format!("    {call};"));
                lines.push(String::new());
                lines.push("    // 確認: 実行した後の状態（セルの値など）を確かめる。例:".into());
                let sheet = sheet_expr(self.prog);
                lines.push(format!(
                    "    // expect({sheet}.range(\"A1\").value).toBe(\"期待する値\");"
                ));
            }
            t => {
                lines.push(format!("    const actual = {call};"));
                lines.push(String::new());
                lines.push("    // 確認: 期待する値に書き換える".into());
                let check = match t {
                    Type::Array(_) | Type::Struct(_) | Type::Map(..) => {
                        let t = t.clone();
                        format!("expect(actual).toEqual({});", self.sample(&t))
                    }
                    t if t.is_object() && !matches!(t, Type::Union(_)) => {
                        "expect(actual).not.toBeNothing();".to_string()
                    }
                    Type::Float | Type::Double | Type::Currency | Type::Decimal => {
                        "expect(actual).toBeCloseTo(0.0);".to_string()
                    }
                    Type::Any | Type::CellValue | Type::Param(_) | Type::Error => {
                        "expect(actual).toBe(\"\");".to_string()
                    }
                    t => {
                        let t = t.clone();
                        format!("expect(actual).toBe({});", self.sample(&t))
                    }
                };
                lines.push(format!("    {check}"));
            }
        }
        lines.push("});".into());
        if has_args {
            lines.push(format!(
                "test.todo({});",
                quote("境界の値（0・空の文字列・最大の値など）")
            ));
        }
        lines
    }
}

/// The path from the project root an `import` in the file at `file` names (a relative path,
/// deprecated, is resolved from the file: `../src/report` in `test/report.test.ras`).
fn root_path(file: &str, source: &str) -> String {
    if !source.starts_with('.') {
        return source.strip_suffix(".ras").unwrap_or(source).to_string();
    }
    let mut parts: Vec<&str> = file
        .rsplit_once('/')
        .map(|(d, _)| d.split('/').collect())
        .unwrap_or_default();
    for seg in source.split('/') {
        match seg {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    let path = parts.join("/");
    path.strip_suffix(".ras").unwrap_or(&path).to_string()
}
