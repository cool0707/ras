//! Position-based queries on an analysed program (hover, completion, go to
//! definition, document symbols) — the semantic half of the language server.

use ras_syntax::ast::{self, Expr, ExprKind, Item, PassMode, Span};

use crate::astutil::{walk_expr, walk_stmt, walk_stmts};
use crate::intrinsics;
use crate::model::*;

/// Calls `f` for every expression of a module (function bodies, initializers, decorators).
pub fn walk_module<'a>(module: &'a ast::Module, f: &mut dyn FnMut(&'a Expr)) {
    fn func<'a>(d: &'a ast::FnDecl, f: &mut dyn FnMut(&'a Expr)) {
        let mut fe = |e: &'a Expr| {
            f(e);
            true
        };
        for p in &d.params {
            if let Some(x) = &p.default {
                walk_expr(x, &mut |_| true, &mut fe);
            }
        }
        if let Some(b) = &d.body {
            walk_stmts(&b.stmts, &mut |_| true, &mut fe);
        }
    }
    for item in &module.items {
        match item {
            Item::Fn(d) => func(d, f),
            Item::Class(c) => {
                for m in &c.members {
                    match m {
                        ast::ClassMember::Field(fl) => {
                            if let Some(x) = &fl.init {
                                walk_expr(x, &mut |_| true, &mut |e| {
                                    f(e);
                                    true
                                });
                            }
                        }
                        ast::ClassMember::Method(md) => func(&md.func, f),
                        ast::ClassMember::Ctor(d) => func(d, f),
                        ast::ClassMember::Accessor(a) => func(&a.func, f),
                    }
                }
            }
            Item::Var(v) => {
                if let Some(x) = &v.init {
                    walk_expr(x, &mut |_| true, &mut |e| {
                        f(e);
                        true
                    });
                }
            }
            Item::Stmt(s) => walk_stmt(s, &mut |_| true, &mut |e| {
                f(e);
                true
            }),
            _ => {}
        }
    }
}

/// Innermost expression whose span contains `offset`.
pub fn expr_at(module: &ast::Module, offset: u32) -> Option<&Expr> {
    let mut best: Option<&Expr> = None;
    walk_module(module, &mut |e| {
        let (s, t) = (u32::from(e.span.start()), u32::from(e.span.end()));
        if s <= offset && offset <= t && best.is_none_or(|b| e.span.len() <= b.span.len()) {
            best = Some(e);
        }
    });
    best
}

/// Function whose declaration contains `offset` (falls back to the module's `Main` / initializer).
pub fn fn_at(prog: &Program, m: ModuleId, offset: u32) -> Option<FnId> {
    let md = &prog.modules[m];
    let mut best: Option<(FnId, u32)> = None;
    for &f in &md.fns {
        let d = &prog.fns[f];
        if matches!(d.kind, FnKind::Main | FnKind::ModInit) {
            continue;
        }
        let Some(decl) = &d.decl else { continue };
        let (s, t) = (u32::from(decl.span.start()), u32::from(decl.span.end()));
        if s <= offset && offset <= t && best.is_none_or(|b| t - s < b.1) {
            best = Some((f, t - s));
        }
    }
    best.map(|b| b.0).or(md.main_fn).or(md.init_fn)
}

fn vba_type(t: &Type) -> Option<&'static str> {
    Some(match t {
        Type::Byte => "Byte",
        Type::Short => "Integer",
        Type::Int => "Long",
        Type::Long => "LongLong",
        Type::Float => "Single",
        Type::Double => "Double",
        Type::Currency => "Currency",
        Type::Decimal | Type::Any => "Variant",
        Type::Ptr => "LongPtr",
        Type::Date => "Date",
        Type::Str => "String",
        Type::Bool => "Boolean",
        Type::Object | Type::Map(..) | Type::Func(_) => "Object",
        _ => return None,
    })
}

/// `function name<T>(a: int, byref b: double): boolean`
pub fn signature(prog: &Program, f: FnId) -> String {
    let d = &prog.fns[f];
    let mut s = String::new();
    if d.inline == ast::InlineMode::Inline {
        s.push_str("inline ");
    }
    s.push_str(match d.kind {
        FnKind::Getter { .. } => "get ",
        FnKind::Setter { .. } => "set ",
        FnKind::Method { .. } | FnKind::Ctor { .. } | FnKind::Abstract { .. } | FnKind::Host { class: Some(_) } => "",
        _ => "function ",
    });
    s.push_str(if d.name.is_empty() { "(既定メンバー)" } else { &d.name });
    if !d.type_params.is_empty() {
        s.push('<');
        s.push_str(&d.type_params.iter().map(|p| p.name.clone()).collect::<Vec<_>>().join(", "));
        s.push('>');
    }
    s.push('(');
    let ps: Vec<String> = d
        .params
        .iter()
        .map(|p| {
            let mut t = String::new();
            if p.mode == PassMode::ByRef {
                t.push_str("byref ");
            }
            t.push_str(if d.locals.get(p.local).is_some_and(|l| l.kind == LocalKind::This) { "this" } else { &p.name });
            if p.optional || p.default.is_some() {
                t.push('?');
            }
            t.push_str(": ");
            t.push_str(&prog.show(&p.ty));
            t
        })
        .collect();
    s.push_str(&ps.join(", "));
    s.push_str("): ");
    s.push_str(&prog.show(&d.ret_ty()));
    s
}

pub struct Hover {
    pub span: Span,
    pub markdown: String,
}

fn code_block(code: &str, doc: Option<&str>) -> String {
    let mut s = format!("```ras\n{code}\n```");
    if let Some(d) = doc {
        s.push_str("\n\n");
        s.push_str(d);
    }
    s
}

fn with_vba(prog: &Program, head: &str, ty: &Type) -> String {
    match vba_type(ty) {
        Some(v) => format!("{head}: {}\n// VBA: {v}", prog.show(ty)),
        None => format!("{head}: {}", prog.show(ty)),
    }
}

fn describe_res(prog: &Program, m: ModuleId, offset: u32, res: &Res, ty: &Type) -> Option<String> {
    Some(match res {
        Res::Local(l) => {
            let f = fn_at(prog, m, offset)?;
            let d = prog.fns[f].locals.get(*l)?;
            let kw = match d.kind {
                LocalKind::Const | LocalKind::Iter => "const",
                LocalKind::Param | LocalKind::LambdaParam | LocalKind::This => "(引数)",
                _ => "let",
            };
            code_block(&with_vba(prog, &format!("{kw} {}", d.name), ty), None)
        }
        Res::Global(v) => {
            let d = &prog.vars[*v];
            let head = format!("{} {}", if d.is_const { "const" } else { "let" }, d.name);
            match &d.host_name {
                // A host global: the VBA line shows its spelling there instead of the type.
                Some(host) => code_block(&format!("{head}: {}\n// VBA: {host}", prog.show(&d.ty)), d.doc.as_deref()),
                None => code_block(&with_vba(prog, &head, &d.ty), d.doc.as_deref()),
            }
        }
        Res::Fn(f) | Res::Method(f) => {
            let d = &prog.fns[*f];
            match (&d.host_name, d.kind.clone()) {
                (Some(host), FnKind::Host { class: Some(c) }) => code_block(&format!("{}\n// VBA: {}.{host}", signature(prog, *f), prog.classes[c].name), d.doc.as_deref()),
                (Some(host), _) => code_block(&format!("{}\n// VBA: {host}", signature(prog, *f)), d.doc.as_deref()),
                _ => code_block(&signature(prog, *f), d.doc.as_deref()),
            }
        }
        Res::ExtMethod(f) => {
            let d = &prog.fns[*f];
            let md = &prog.modules[d.module];
            code_block(&format!("{}\n// 拡張メソッド（{}）→ VBA: {}.{}(…)", signature(prog, *f), md.key, md.vba_name, d.name), d.doc.as_deref())
        }
        Res::Property { get, set } => {
            let f = get.or(*set)?;
            code_block(&format!("(プロパティ) {}: {}", prog.fns[f].name, prog.show(ty)), prog.fns[f].doc.as_deref())
        }
        Res::Field(c, i) => {
            let fl = &prog.classes[*c].fields[*i];
            let head = format!("(プロパティ) {}.{}", prog.class_display(*c), fl.name);
            match &fl.host_name {
                Some(host) => code_block(&format!("{head}: {}\n// VBA: {}.{host}", prog.show(ty), prog.classes[*c].name), fl.doc.as_deref()),
                None => code_block(&with_vba(prog, &head, ty), fl.doc.as_deref()),
            }
        }
        Res::Class(c) => {
            let d = &prog.classes[*c];
            let kw = if d.kind == ClassKind::Class { "class" } else { "interface" };
            code_block(&format!("{kw} {}", prog.class_display(*c)), d.doc.as_deref())
        }
        Res::Enum(e) => code_block(&format!("enum {}", prog.enums[*e].name), prog.enums[*e].doc.as_deref()),
        Res::EnumMember(e, i) => {
            let d = &prog.enums[*e];
            code_block(&format!("{}.{} = {}", d.name, d.members[*i].name, d.members[*i].value), d.members[*i].doc.as_deref())
        }
        Res::Module(t) => code_block(&format!("module \"{}\"  // VBA: {}", prog.modules[*t].key, prog.modules[*t].vba_name), None),
        Res::Namespace(n) => code_block(&format!("namespace {n}"), None),
        Res::Builtin(n) => code_block(&format!("(組み込み) {n}"), None),
        Res::Intrinsic(_) => code_block(&format!("(組み込み): {}", prog.show(ty)), None),
        Res::Dynamic => code_block("(遅延バインディング): any", Some("静的な型が無いため、実行時に解決されます。")),
    })
}

pub fn hover(prog: &Program, m: ModuleId, offset: u32) -> Option<Hover> {
    let md = &prog.modules[m];
    // Declarations first: function names and local declarations are not expressions.
    for &f in &md.fns {
        let d = &prog.fns[f];
        if d.name_span.contains_inclusive(offset.into()) && !d.name_span.is_empty() && !matches!(d.kind, FnKind::Main | FnKind::ModInit) {
            return Some(Hover { span: d.name_span, markdown: code_block(&signature(prog, f), d.doc.as_deref()) });
        }
    }
    if let Some(f) = fn_at(prog, m, offset) {
        for l in &prog.fns[f].locals {
            if l.span.contains_inclusive(offset.into()) && !l.span.is_empty() && !l.name.starts_with("ras$") {
                return Some(Hover { span: l.span, markdown: code_block(&with_vba(prog, &l.name, &l.ty), None) });
            }
        }
    }
    let e = expr_at(&md.ast, offset)?;
    let ty = md.info.ty(e.id);
    let (span, target) = match &e.kind {
        ExprKind::Member { name, .. } if name.span.contains_inclusive(offset.into()) => (name.span, e),
        ExprKind::Name(_) | ExprKind::This => (e.span, e),
        _ => (e.span, e),
    };
    let markdown = match md.info.res.get(&target.id) {
        Some(res) => describe_res(prog, m, offset, res, &ty)?,
        None if ty.is_error() => return None,
        None => code_block(&with_vba(prog, "(式)", &ty), None),
    };
    Some(Hover { span, markdown })
}

/// Location of the declaration the expression at `offset` refers to.
pub fn definition(prog: &Program, m: ModuleId, offset: u32) -> Option<(ModuleId, Span)> {
    let md = &prog.modules[m];
    let e = expr_at(&md.ast, offset)?;
    match md.info.res.get(&e.id)? {
        Res::Local(l) => {
            let f = fn_at(prog, m, offset)?;
            Some((m, prog.fns[f].locals.get(*l)?.span))
        }
        Res::Global(v) => Some((prog.vars[*v].module, prog.vars[*v].name_span)),
        Res::Fn(f) | Res::Method(f) | Res::ExtMethod(f) => Some((prog.fns[*f].module, prog.fns[*f].name_span)),
        Res::Property { get, set } => {
            let f = get.or(*set)?;
            Some((prog.fns[f].module, prog.fns[f].name_span))
        }
        Res::Field(c, i) => Some((prog.classes[*c].module, prog.classes[*c].fields[*i].span)),
        Res::Class(c) => Some((prog.classes[*c].module, prog.classes[*c].name_span)),
        Res::Enum(e) => Some((prog.enums[*e].module, prog.enums[*e].name_span)),
        Res::EnumMember(e, i) => Some((prog.enums[*e].module, prog.enums[*e].members[*i].span)),
        Res::Module(t) => Some((*t, Span::default())),
        _ => None,
    }
}

/// LSP `CompletionItemKind` values.
pub mod kind {
    pub const METHOD: u32 = 2;
    pub const FUNCTION: u32 = 3;
    pub const FIELD: u32 = 5;
    pub const VARIABLE: u32 = 6;
    pub const CLASS: u32 = 7;
    pub const MODULE: u32 = 9;
    pub const PROPERTY: u32 = 10;
    pub const ENUM: u32 = 13;
    pub const KEYWORD: u32 = 14;
    pub const ENUM_MEMBER: u32 = 20;
    pub const CONSTANT: u32 = 21;
}

pub struct Completion {
    pub label: String,
    pub kind: u32,
    pub detail: String,
    pub doc: Option<String>,
    /// Sorts type members before extension methods (spec §16.2.1).
    pub sort: u8,
}

fn item(label: &str, kind: u32, detail: impl Into<String>, doc: Option<&String>, sort: u8) -> Completion {
    Completion { label: label.to_string(), kind, detail: detail.into(), doc: doc.cloned(), sort }
}

fn sym_item(prog: &Program, name: &str, sym: &Sym, sort: u8) -> Option<Completion> {
    Some(match sym {
        Sym::Fn(f) => item(name, kind::FUNCTION, signature(prog, *f), prog.fns[*f].doc.as_ref(), sort),
        Sym::Class(c) => item(name, kind::CLASS, prog.class_display(*c), prog.classes[*c].doc.as_ref(), sort),
        Sym::Enum(e) => item(name, kind::ENUM, format!("enum {}", prog.enums[*e].name), prog.enums[*e].doc.as_ref(), sort),
        Sym::EnumMember(e, i) => {
            let d = &prog.enums[*e];
            item(name, kind::ENUM_MEMBER, format!("{} = {}", d.name, d.members[*i].value), d.members[*i].doc.as_ref(), sort)
        }
        Sym::Var(v) => item(name, if prog.vars[*v].is_const { kind::CONSTANT } else { kind::VARIABLE }, prog.show(&prog.vars[*v].ty), prog.vars[*v].doc.as_ref(), sort),
        Sym::Module(t) => item(name, kind::MODULE, format!("module \"{}\"", prog.modules[*t].key), None, sort),
        Sym::Namespace(n) => item(name, kind::MODULE, format!("namespace {n}"), None, sort),
        Sym::Struct(_) | Sym::Alias(_) => item(name, kind::CLASS, "型", None, sort),
    })
}

fn class_members(prog: &Program, c: ClassId, statics: bool, out: &mut Vec<Completion>) {
    let d = &prog.classes[c];
    for f in &d.fields {
        if f.is_static == statics && !out.iter().any(|o| o.label == f.name) {
            out.push(item(&f.name, kind::FIELD, prog.show(&f.ty), f.doc.as_ref(), 0));
        }
    }
    for &m in &d.methods {
        let fd = &prog.fns[m];
        let is_static = matches!(fd.kind, FnKind::Method { is_static: true, .. } | FnKind::Getter { is_static: true, .. } | FnKind::Setter { is_static: true, .. });
        if is_static != statics || out.iter().any(|o| o.label == fd.name) {
            continue;
        }
        let k = if matches!(fd.kind, FnKind::Getter { .. } | FnKind::Setter { .. }) { kind::PROPERTY } else { kind::METHOD };
        out.push(item(&fd.name, k, signature(prog, m), fd.doc.as_ref(), 0));
    }
    for base in &d.extends {
        if let Type::Class(b, _) = base {
            class_members(prog, *b, statics, out);
        }
    }
}

fn members_of(prog: &Program, m: ModuleId, ty: &Type) -> Vec<Completion> {
    let mut out = Vec::new();
    let builtin = |names: &[&str], out: &mut Vec<Completion>| {
        for n in names {
            out.push(item(n, kind::METHOD, "(組み込み)", None, 0));
        }
    };
    match ty {
        Type::Class(c, _) => {
            class_members(prog, *c, false, &mut out);
            let d = &prog.classes[*c];
            if d.kind == ClassKind::Host && d.namespace.as_deref() == Some("Excel") {
                match d.name.as_str() {
                    "Range" => builtin(&["toArray", "fromArray", "toColumn", "toRow"], &mut out),
                    "Worksheet" => builtin(&["usedTable"], &mut out),
                    _ => {}
                }
            }
            // Array-style methods of a collection; a member of the type with the same name wins (`range.find`).
            if prog.collection_elem(ty).is_some() {
                for n in intrinsics::COLLECTION_MEMBERS {
                    if !out.iter().any(|o| o.label == *n) {
                        out.push(item(n, kind::METHOD, "(組み込み・コレクション)", None, 0));
                    }
                }
            }
        }
        Type::Struct(s) => {
            let d = &prog.structs[*s];
            if d.kind != StructKind::Tuple {
                for (name, fty) in &d.fields {
                    out.push(item(name, kind::FIELD, prog.show(fty), None, 0));
                }
            }
        }
        Type::Array(_) if ty.dims() == 2 => builtin(&["length", "rows", "cols"], &mut out),
        Type::Array(_) => builtin(intrinsics::ARRAY_MEMBERS, &mut out),
        Type::Str => builtin(intrinsics::STRING_MEMBERS, &mut out),
        Type::Map(..) => builtin(intrinsics::MAP_MEMBERS, &mut out),
        Type::Ns(NsRef::Module(t)) => {
            for (n, s) in &prog.modules[*t].exports {
                out.extend(sym_item(prog, n, s, 0));
            }
        }
        Type::Ns(NsRef::Host(ns)) => {
            for (n, s) in prog.namespaces.get(ns).into_iter().flatten() {
                out.extend(sym_item(prog, n, s, 0));
            }
        }
        Type::Ns(NsRef::Enum(e)) => {
            for mem in &prog.enums[*e].members {
                out.push(item(&mem.name, kind::ENUM_MEMBER, mem.value.to_string(), mem.doc.as_ref(), 0));
            }
        }
        Type::Ns(NsRef::Class(c)) => class_members(prog, *c, true, &mut out),
        Type::Ns(NsRef::Builtin(ns)) => match *ns {
            "Math" => builtin(intrinsics::MATH_MEMBERS, &mut out),
            "console" => builtin(&["log"], &mut out),
            "Platform" => builtin(&["is64bit", "isVBA7"], &mut out),
            _ => {}
        },
        t if t.is_numeric() => builtin(&["toString", "toFixed"], &mut out),
        Type::Bool | Type::Date => builtin(&["toString"], &mut out),
        // `any` / `object`: no member list (late binding, spec §16.2.1).
        _ => {}
    }
    // Extension methods whose receiver type matches come after the members of the type.
    if ty.is_value() && !matches!(ty, Type::Any | Type::Object | Type::Error) {
        for &f in &prog.modules[m].ext_methods {
            let d = &prog.fns[f];
            let Some(p0) = d.params.first() else { continue };
            let fits = p0.ty == *ty || (p0.ty.contains_param() && p0.ty.is_array() == ty.is_array()) || matches!((&p0.ty, ty), (Type::Class(a, _), Type::Class(b, _)) if prog.class_conforms(*b, *a));
            if fits && !out.iter().any(|o| o.label == d.name && o.sort == 1) {
                out.push(item(&d.name, kind::METHOD, format!("(拡張) {}", signature(prog, f)), d.doc.as_ref(), 1));
            }
        }
    }
    out
}

const KEYWORDS: &[&str] = &[
    "let", "const", "function", "return", "if", "else", "for", "of", "while", "do", "break", "continue", "switch", "case", "default",
    "try", "catch", "finally", "throw", "new", "class", "interface", "enum", "implements", "import", "export", "from", "as", "this",
    "null", "true", "false", "byref", "byval", "inline", "noinline", "resume", "using", "declare", "instanceof", "typeof",
];

pub fn completions(prog: &Program, m: ModuleId, offset: u32) -> Vec<Completion> {
    let md = &prog.modules[m];
    let text: &str = &md.text;
    let offset = (offset as usize).min(text.len());
    // Start of the identifier being typed.
    let mut p = offset;
    while p > 0 {
        let Some(c) = text[..p].chars().next_back() else { break };
        if c.is_alphanumeric() || c == '_' || c == '$' {
            p -= c.len_utf8();
        } else {
            break;
        }
    }
    if text[..p].ends_with('.') {
        let dot = (p - 1) as u32;
        // The receiver is the largest expression that ends right at the dot.
        let mut recv: Option<&Expr> = None;
        walk_module(&md.ast, &mut |e| {
            let end = match &e.kind {
                // error recovery: `obj.` parses as a member access with an empty name
                ExprKind::Member { object, name, .. } if name.text.is_empty() || u32::from(name.span.start()) > dot => u32::from(object.span.end()),
                _ => u32::from(e.span.end()),
            };
            if end == dot && !matches!(&e.kind, ExprKind::Member { name, .. } if name.text.is_empty()) && recv.is_none_or(|r| e.span.len() >= r.span.len()) {
                recv = Some(e);
            }
        });
        return match recv {
            Some(e) => members_of(prog, m, &md.info.ty(e.id)),
            None => Vec::new(),
        };
    }
    let mut out = Vec::new();
    if let Some(f) = fn_at(prog, m, offset as u32) {
        for l in &prog.fns[f].locals {
            if !l.name.is_empty() && !l.name.starts_with("ras$") && u32::from(l.span.start()) < offset as u32 && !out.iter().any(|o: &Completion| o.label == l.name) && l.kind != LocalKind::This {
                out.push(item(&l.name, kind::VARIABLE, prog.show(&l.ty), None, 0));
            }
        }
    }
    for (n, s) in &md.scope {
        if n.starts_with("ras$") {
            continue;
        }
        out.extend(sym_item(prog, n, s, 1));
    }
    for (n, s) in &prog.globals {
        if !md.scope.contains_key(n) {
            out.extend(sym_item(prog, n, s, 3));
        }
    }
    for ns in prog.namespaces.keys() {
        out.push(item(ns, kind::MODULE, format!("namespace {ns}"), None, 2));
    }
    for ns in intrinsics::BUILTIN_NAMESPACES {
        out.push(item(ns, kind::MODULE, "(組み込み)", None, 2));
    }
    for f in intrinsics::GLOBAL_FNS {
        out.push(item(f, kind::FUNCTION, "(組み込み)", None, 2));
    }
    for k in KEYWORDS.iter().chain(crate::collect::PRIMITIVE_NAMES.iter()) {
        out.push(item(k, kind::KEYWORD, "", None, 4));
    }
    out
}

/// LSP `SymbolKind` values.
pub mod symbol_kind {
    pub const CLASS: u32 = 5;
    pub const METHOD: u32 = 6;
    pub const FIELD: u32 = 8;
    pub const ENUM: u32 = 10;
    pub const INTERFACE: u32 = 11;
    pub const FUNCTION: u32 = 12;
    pub const VARIABLE: u32 = 13;
    pub const CONSTANT: u32 = 14;
}

pub struct Symbol {
    pub name: String,
    pub kind: u32,
    pub detail: String,
    pub span: Span,
    pub name_span: Span,
    pub children: Vec<Symbol>,
}

pub fn symbols(prog: &Program, m: ModuleId) -> Vec<Symbol> {
    let md = &prog.modules[m];
    let mut out = Vec::new();
    for &f in &md.fns {
        let d = &prog.fns[f];
        if !matches!(d.kind, FnKind::Free | FnKind::Declare { .. }) || d.name.is_empty() {
            continue;
        }
        let span = d.decl.as_ref().map(|x| x.span).unwrap_or(d.name_span);
        out.push(Symbol { name: d.name.clone(), kind: symbol_kind::FUNCTION, detail: signature(prog, f), span, name_span: d.name_span, children: Vec::new() });
    }
    for &c in &md.classes {
        let d = &prog.classes[c];
        let mut children = Vec::new();
        for fl in &d.fields {
            children.push(Symbol { name: fl.name.clone(), kind: symbol_kind::FIELD, detail: prog.show(&fl.ty), span: fl.span, name_span: fl.span, children: Vec::new() });
        }
        for &me in d.methods.iter().chain(d.ctor.iter()) {
            let fd = &prog.fns[me];
            let span = fd.decl.as_ref().map(|x| x.span).unwrap_or(fd.name_span);
            children.push(Symbol { name: fd.name.clone(), kind: symbol_kind::METHOD, detail: signature(prog, me), span, name_span: fd.name_span, children: Vec::new() });
        }
        let span = children.iter().fold(d.name_span, |acc, ch| acc.cover(ch.span));
        let kind = if d.kind == ClassKind::Class { symbol_kind::CLASS } else { symbol_kind::INTERFACE };
        out.push(Symbol { name: d.name.clone(), kind, detail: String::new(), span, name_span: d.name_span, children });
    }
    for &e in &md.enums {
        let d = &prog.enums[e];
        out.push(Symbol { name: d.name.clone(), kind: symbol_kind::ENUM, detail: String::new(), span: d.name_span, name_span: d.name_span, children: Vec::new() });
    }
    for &v in &md.vars {
        let d = &prog.vars[v];
        let kind = if d.is_const { symbol_kind::CONSTANT } else { symbol_kind::VARIABLE };
        out.push(Symbol { name: d.name.clone(), kind, detail: prog.show(&d.ty), span: d.name_span, name_span: d.name_span, children: Vec::new() });
    }
    out.sort_by_key(|s| u32::from(s.span.start()));
    out
}
