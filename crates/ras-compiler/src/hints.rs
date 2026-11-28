//! Inlay hints, semantic tokens and folding ranges (spec §16.2): what the editor shows on
//! top of the text.

use std::collections::HashMap;

use ras_syntax::ast::{Decorator, Expr, ExprKind, Ident, Item, Span, Stmt, StmtKind, TypeRef};
use ras_syntax::SyntaxKind;

use crate::astutil::{visit_module, Visit};
use crate::model::*;
use crate::query::{fn_at, vba_type};
use crate::refs::{res_target, type_target, Target};

// ---------------------------------------------------------------------------- inlay hints

/// LSP `InlayHintKind`.
pub mod hint_kind {
    pub const TYPE: u32 = 1;
    pub const PARAMETER: u32 = 2;
}

pub struct InlayHint {
    pub offset: u32,
    pub label: String,
    pub kind: u32,
    pub pad_left: bool,
    pub pad_right: bool,
    pub tooltip: Option<String>,
}

struct Hints<'p> {
    prog: &'p Program,
    m: ModuleId,
    out: Vec<InlayHint>,
}

impl Hints<'_> {
    /// `: T` after a declaration without a type annotation.
    fn type_hint(&mut self, name: &Ident, ty: &Type) {
        if name.text.is_empty()
            || name.text.starts_with("ras$")
            || matches!(ty, Type::Error | Type::Void | Type::Null | Type::Ns(_))
        {
            return;
        }
        let tooltip = vba_type(ty).map(|v| format!("推論された型（VBA: {v}）"));
        self.out.push(InlayHint {
            offset: name.span.end().into(),
            label: format!(": {}", self.prog.show(ty)),
            kind: hint_kind::TYPE,
            pad_left: false,
            pad_right: false,
            tooltip,
        });
    }

    /// The type of the local or top-level variable declared at `name`.
    fn declared(&self, name: &Ident) -> Option<Type> {
        let md = &self.prog.modules[self.m];
        let off: u32 = name.span.start().into();
        if let Some(&l) = md.info.decl_locals.get(&off) {
            let f = fn_at(self.prog, self.m, off)?;
            return self.prog.fns[f]
                .locals
                .get(l)
                .filter(|d| d.span == name.span)
                .map(|d| d.ty.clone());
        }
        if let Some(&v) = md.info.decl_globals.get(&off) {
            return Some(self.prog.vars[v].ty.clone());
        }
        md.vars
            .iter()
            .map(|&v| &self.prog.vars[v])
            .find(|v| v.name_span == name.span)
            .map(|v| v.ty.clone())
    }

    fn var_hint(&mut self, name: &Ident, init: Option<&Expr>) {
        // The type is written right there.
        if init.is_some_and(|i| matches!(i.peel().kind, ExprKind::New { .. } | ExprKind::As { .. }))
        {
            return;
        }
        if let Some(ty) = self.declared(name) {
            self.type_hint(name, &ty);
        }
    }
}

impl<'a> Visit<'a> for Hints<'_> {
    fn stmt(&mut self, s: &'a Stmt) {
        match &s.kind {
            StmtKind::Var(v) if v.ty.is_none() => self.var_hint(&v.name, v.init.as_ref()),
            StmtKind::ForOf { name, ty: None, .. } => self.var_hint(name, None),
            _ => {}
        }
    }

    fn expr(&mut self, e: &'a Expr) {
        let prog = self.prog;
        let md = &prog.modules[self.m];
        match &e.kind {
            ExprKind::Arrow(a) => {
                for p in a.params.iter().filter(|p| p.ty.is_none()) {
                    self.var_hint(&p.name, None);
                }
            }
            ExprKind::Call {
                callee,
                type_args,
                args,
            } => {
                let Some(call) = md.info.calls.get(&e.id) else {
                    return;
                };
                let (f, skip) = match call.target {
                    CallTarget::Fn(f)
                    | CallTarget::Method(f)
                    | CallTarget::Host(f)
                    | CallTarget::DefaultMember(f) => (f, 0),
                    CallTarget::Ext(f) => (f, 1),
                    _ => return,
                };
                // Type arguments inferred for a generic function.
                if type_args.is_empty()
                    && prog.fns[f].is_generic()
                    && !call.type_args.is_empty()
                    && call
                        .type_args
                        .iter()
                        .all(|t| !t.is_error() && !t.contains_param())
                {
                    let end = match &callee.kind {
                        ExprKind::Member { name, .. } => name.span.end(),
                        _ => callee.span.end(),
                    };
                    let label = format!(
                        "<{}>",
                        call.type_args
                            .iter()
                            .map(|t| prog.show(t))
                            .collect::<Vec<_>>()
                            .join(", ")
                    );
                    self.out.push(InlayHint {
                        offset: end.into(),
                        label,
                        kind: hint_kind::TYPE,
                        pad_left: false,
                        pad_right: false,
                        tooltip: Some("推論された型引数".into()),
                    });
                }
                self.byref_hints(e, f, skip, args);
            }
            ExprKind::New { ty, args } => {
                let Some(call) = md.info.calls.get(&e.id) else {
                    return;
                };
                let CallTarget::Ctor(c) = call.target else {
                    return;
                };
                if matches!(ty, TypeRef::Named { args, .. } if args.is_empty())
                    && !prog.classes[c].type_params.is_empty()
                    && !call.type_args.is_empty()
                    && call.type_args.iter().all(|t| !t.is_error())
                {
                    let label = format!(
                        "<{}>",
                        call.type_args
                            .iter()
                            .map(|t| prog.show(t))
                            .collect::<Vec<_>>()
                            .join(", ")
                    );
                    self.out.push(InlayHint {
                        offset: ty.span().end().into(),
                        label,
                        kind: hint_kind::TYPE,
                        pad_left: false,
                        pad_right: false,
                        tooltip: Some("推論された型引数".into()),
                    });
                }
                if let Some(f) = prog.classes[c].ctor {
                    self.byref_hints(e, f, 0, args);
                }
            }
            _ => {}
        }
    }
}

impl Hints<'_> {
    /// `byref` before the arguments passed by reference (spec §16.2).
    fn byref_hints(&mut self, call: &Expr, f: FnId, skip: usize, args: &[Expr]) {
        let md = &self.prog.modules[self.m];
        let d = &self.prog.fns[f];
        for (i, a) in md.info.call_args(call.id, args).iter().enumerate() {
            let Some(p) = d.params.get(i + skip) else {
                break;
            };
            if p.mode == ras_syntax::ast::PassMode::ByRef
                && !a.span.is_empty()
                && !matches!(a.kind, ExprKind::Undefined)
            {
                let tooltip = format!(
                    "`{}` は参照渡し（byref）です。呼び出し先で書き換えられることがあります",
                    p.name
                );
                self.out.push(InlayHint {
                    offset: a.span.start().into(),
                    label: "byref".into(),
                    kind: hint_kind::PARAMETER,
                    pad_left: false,
                    pad_right: true,
                    tooltip: Some(tooltip),
                });
            }
        }
    }
}

/// The inlay hints of the byte range `from..=to` of module `m`: inferred types of variables
/// and lambda parameters, inferred type arguments, and `byref` at the arguments passed by
/// reference (spec §16.2).
pub fn inlay_hints(prog: &Program, m: ModuleId, from: u32, to: u32) -> Vec<InlayHint> {
    let md = &prog.modules[m];
    let mut h = Hints {
        prog,
        m,
        out: Vec::new(),
    };
    for item in md.ast.items.iter() {
        if let Item::Var(v) = item {
            if v.ty.is_none() && !v.declared {
                h.var_hint(&v.name, v.init.as_ref());
            }
        }
    }
    visit_module(&md.ast, &mut h);
    let mut out: Vec<InlayHint> = h
        .out
        .into_iter()
        .filter(|x| from <= x.offset && x.offset <= to)
        .collect();
    out.sort_by_key(|x| x.offset);
    out
}

// ---------------------------------------------------------------------------- semantic tokens

/// The legend of the semantic tokens: LSP token types…
pub const TOKEN_TYPES: &[&str] = &[
    "namespace",
    "type",
    "class",
    "enum",
    "interface",
    "struct",
    "typeParameter",
    "parameter",
    "variable",
    "property",
    "enumMember",
    "function",
    "method",
];
/// …and modifiers. `byref` and `extension` are RAS's own (declared by the VS Code extension).
pub const TOKEN_MODIFIERS: &[&str] = &[
    "declaration",
    "readonly",
    "static",
    "defaultLibrary",
    "deprecated",
    "byref",
    "extension",
];

mod tt {
    pub const NAMESPACE: u32 = 0;
    pub const TYPE: u32 = 1;
    pub const CLASS: u32 = 2;
    pub const ENUM: u32 = 3;
    pub const INTERFACE: u32 = 4;
    pub const STRUCT: u32 = 5;
    pub const TYPE_PARAMETER: u32 = 6;
    pub const PARAMETER: u32 = 7;
    pub const VARIABLE: u32 = 8;
    pub const PROPERTY: u32 = 9;
    pub const ENUM_MEMBER: u32 = 10;
    pub const FUNCTION: u32 = 11;
    pub const METHOD: u32 = 12;
}

mod tm {
    pub const DECLARATION: u32 = 1 << 0;
    pub const READONLY: u32 = 1 << 1;
    pub const STATIC: u32 = 1 << 2;
    pub const DEFAULT_LIBRARY: u32 = 1 << 3;
    pub const DEPRECATED: u32 = 1 << 4;
    pub const BYREF: u32 = 1 << 5;
    pub const EXTENSION: u32 = 1 << 6;
}

fn is_lib(prog: &Program, m: ModuleId) -> bool {
    prog.modules[m].path.starts_with("<stdlib>")
}

fn deprecated(doc: Option<&String>) -> u32 {
    if doc.is_some_and(|d| d.contains("@deprecated")) {
        tm::DEPRECATED
    } else {
        0
    }
}

/// Token type and modifiers of a declaration.
fn classify(prog: &Program, t: &Target) -> (u32, u32) {
    match t {
        Target::Local { func, local } => {
            let l = &prog.fns[*func].locals[*local];
            let ty = if matches!(
                l.kind,
                LocalKind::Param | LocalKind::LambdaParam | LocalKind::This
            ) {
                tt::PARAMETER
            } else {
                tt::VARIABLE
            };
            let mut mods = 0;
            if matches!(l.kind, LocalKind::Const | LocalKind::Iter) {
                mods |= tm::READONLY;
            }
            if l.byref
                || prog.fns[*func]
                    .params
                    .iter()
                    .any(|p| p.local == *local && p.mode == ras_syntax::ast::PassMode::ByRef)
            {
                mods |= tm::BYREF;
            }
            (ty, mods)
        }
        Target::Global(v) => {
            let d = &prog.vars[*v];
            let mut mods = deprecated(d.doc.as_ref());
            if d.is_const {
                mods |= tm::READONLY;
            }
            if is_lib(prog, d.module) {
                mods |= tm::DEFAULT_LIBRARY;
            }
            (tt::VARIABLE, mods)
        }
        Target::Fn(f) => {
            let d = &prog.fns[*f];
            let mut mods = deprecated(d.doc.as_ref());
            if is_lib(prog, d.module) {
                mods |= tm::DEFAULT_LIBRARY;
            }
            let ty = match d.kind {
                FnKind::Method { is_static, .. } => {
                    if is_static {
                        mods |= tm::STATIC;
                    }
                    tt::METHOD
                }
                FnKind::Getter { is_static, .. } | FnKind::Setter { is_static, .. } => {
                    if is_static {
                        mods |= tm::STATIC;
                    }
                    tt::PROPERTY
                }
                FnKind::Abstract { .. } | FnKind::Host { class: Some(_) } | FnKind::Ctor { .. } => {
                    tt::METHOD
                }
                _ if d.is_extension => {
                    mods |= tm::EXTENSION;
                    tt::FUNCTION
                }
                _ => tt::FUNCTION,
            };
            (ty, mods)
        }
        Target::Class(c) => {
            let d = &prog.classes[*c];
            let mut mods = deprecated(d.doc.as_ref());
            if is_lib(prog, d.module) {
                mods |= tm::DEFAULT_LIBRARY;
            }
            (
                if d.kind == ClassKind::Interface {
                    tt::INTERFACE
                } else {
                    tt::CLASS
                },
                mods,
            )
        }
        Target::Enum(e) => (
            tt::ENUM,
            if is_lib(prog, prog.enums[*e].module) {
                tm::DEFAULT_LIBRARY
            } else {
                0
            } | deprecated(prog.enums[*e].doc.as_ref()),
        ),
        Target::EnumMember(e, i) => {
            let lib = if is_lib(prog, prog.enums[*e].module) {
                tm::DEFAULT_LIBRARY
            } else {
                0
            };
            (
                tt::ENUM_MEMBER,
                tm::READONLY | lib | deprecated(prog.enums[*e].members[*i].doc.as_ref()),
            )
        }
        Target::Field(c, i) => {
            let f = &prog.classes[*c].fields[*i];
            let mut mods = deprecated(f.doc.as_ref());
            if f.readonly {
                mods |= tm::READONLY;
            }
            if f.is_static {
                mods |= tm::STATIC;
            }
            if is_lib(prog, prog.classes[*c].module) {
                mods |= tm::DEFAULT_LIBRARY;
            }
            (tt::PROPERTY, mods)
        }
        Target::Struct(_) => (tt::STRUCT, 0),
        Target::StructField(..) => (tt::PROPERTY, 0),
        Target::Alias { .. } => (tt::TYPE, 0),
    }
}

struct Tokens<'p> {
    prog: &'p Program,
    m: ModuleId,
    /// Byte offset → (length in bytes, type, modifiers).
    out: HashMap<u32, (u32, u32, u32)>,
}

impl Tokens<'_> {
    fn push(&mut self, span: Span, ty: u32, mods: u32) {
        if !span.is_empty() {
            self.out
                .entry(span.start().into())
                .or_insert((span.len().into(), ty, mods));
        }
    }

    fn target(&mut self, span: Span, t: &Target, decl: bool) {
        let (ty, mut mods) = classify(self.prog, t);
        if decl {
            mods |= tm::DECLARATION;
        }
        self.push(span, ty, mods);
    }
}

impl<'a> Visit<'a> for Tokens<'_> {
    fn expr(&mut self, e: &'a Expr) {
        let prog = self.prog;
        let md = &prog.modules[self.m];
        let (span, is_member) = match &e.kind {
            ExprKind::Name(id) => (id.span, false),
            ExprKind::Member { name, .. } => (name.span, true),
            ExprKind::Object(props) => {
                if matches!(md.info.ty(e.id), Type::Struct(_)) {
                    for p in props.iter().filter(|p| p.value.span != p.key.span) {
                        self.push(p.key.span, tt::PROPERTY, 0);
                    }
                }
                return;
            }
            ExprKind::NamedArg { name, .. } => {
                self.push(name.span, tt::PARAMETER, 0);
                return;
            }
            _ => return,
        };
        if let Some(t) = res_target(prog, self.m, e) {
            // An extension method called with `.` is a method of its receiver.
            let ext = matches!(md.info.res.get(&e.id), Some(Res::ExtMethod(_)));
            let (ty, mut mods) = classify(prog, &t);
            if ext {
                mods |= tm::EXTENSION;
            }
            self.push(
                span,
                if ext || (is_member && ty == tt::FUNCTION) {
                    tt::METHOD
                } else {
                    ty
                },
                mods,
            );
            return;
        }
        match md.info.res.get(&e.id) {
            Some(Res::Module(_)) => self.push(span, tt::NAMESPACE, 0),
            Some(Res::Namespace(_) | Res::Builtin(_)) => {
                self.push(span, tt::NAMESPACE, tm::DEFAULT_LIBRARY)
            }
            Some(Res::Intrinsic(i)) => self.push(
                span,
                if i.is_property() {
                    tt::PROPERTY
                } else if is_member {
                    tt::METHOD
                } else {
                    tt::FUNCTION
                },
                tm::DEFAULT_LIBRARY,
            ),
            Some(Res::Dynamic) if is_member => self.push(span, tt::PROPERTY, 0),
            _ => {}
        }
    }

    fn decorator(&mut self, _d: &'a Decorator) {}

    fn named_type(&mut self, path: &'a [Ident], _span: Span, tps: &[&'a str]) {
        let Some(last) = path.last() else { return };
        if path.len() == 1 && tps.contains(&last.text.as_str()) {
            self.push(last.span, tt::TYPE_PARAMETER, 0);
            return;
        }
        if path.len() == 2 {
            let lib = if self.prog.namespaces.contains_key(&path[0].text) {
                tm::DEFAULT_LIBRARY
            } else {
                0
            };
            self.push(path[0].span, tt::NAMESPACE, lib);
        }
        if let Some(t) = type_target(self.prog, self.m, path, tps) {
            self.target(last.span, &t, false);
        }
    }
}

/// The semantic tokens of module `m`, encoded as LSP wants them (relative positions in UTF-16
/// code units, five numbers per token).
pub fn semantic_tokens(prog: &Program, m: ModuleId) -> Vec<u32> {
    let md = &prog.modules[m];
    let mut toks = Tokens {
        prog,
        m,
        out: HashMap::new(),
    };
    // Declarations.
    for &f in &md.fns {
        let d = &prog.fns[f];
        if !matches!(d.kind, FnKind::Main | FnKind::ModInit | FnKind::Ctor { .. }) {
            toks.target(d.name_span, &Target::Fn(f), true);
        }
        for (l, local) in d.locals.iter().enumerate() {
            if !local.name.starts_with("ras$")
                && local.kind != LocalKind::This
                && md.info.decl_locals.get(&u32::from(local.span.start())) == Some(&l)
            {
                toks.target(local.span, &Target::Local { func: f, local: l }, true);
            }
        }
    }
    for &c in &md.classes {
        toks.target(prog.classes[c].name_span, &Target::Class(c), true);
        for i in 0..prog.classes[c].fields.len() {
            toks.target(prog.classes[c].fields[i].span, &Target::Field(c, i), true);
        }
    }
    for &e in &md.enums {
        toks.target(prog.enums[e].name_span, &Target::Enum(e), true);
        for i in 0..prog.enums[e].members.len() {
            toks.target(
                prog.enums[e].members[i].span,
                &Target::EnumMember(e, i),
                true,
            );
        }
    }
    for &v in &md.vars {
        toks.target(prog.vars[v].name_span, &Target::Global(v), true);
    }
    for &s in &md.structs {
        if prog.structs[s].kind == StructKind::Named {
            toks.target(prog.structs[s].name_span, &Target::Struct(s), true);
        }
    }
    for item in md.ast.items.iter() {
        match item {
            Item::Struct(s) => {
                for f in &s.fields {
                    toks.push(f.name.span, tt::PROPERTY, tm::DECLARATION);
                }
            }
            Item::TypeAlias(a) => toks.push(a.name.span, tt::TYPE, tm::DECLARATION),
            Item::Import(imp) => {
                let Some(t) = crate::collect::import_target(prog, m, &imp.source) else {
                    continue;
                };
                for n in &imp.names {
                    let Some(sym) = prog.modules[t].exports.get(&n.name.text) else {
                        continue;
                    };
                    let target = match sym {
                        Sym::Alias(_) => {
                            toks.push(n.name.span, tt::TYPE, 0);
                            continue;
                        }
                        Sym::Fn(f) => Target::Fn(*f),
                        Sym::Class(c) => Target::Class(*c),
                        Sym::Enum(e) => Target::Enum(*e),
                        Sym::EnumMember(e, i) => Target::EnumMember(*e, *i),
                        Sym::Var(v) => Target::Global(*v),
                        Sym::Struct(s) => Target::Struct(*s),
                        Sym::Module(_) | Sym::Namespace(_) => continue,
                    };
                    toks.target(n.name.span, &target, false);
                    if let Some(a) = &n.alias {
                        toks.target(a.span, &target, true);
                    }
                }
                if let Some(ns) = &imp.namespace {
                    toks.push(ns.span, tt::NAMESPACE, tm::DECLARATION);
                }
            }
            _ => {}
        }
    }
    visit_module(&md.ast, &mut toks);
    let mut list: Vec<(u32, (u32, u32, u32))> = toks.out.into_iter().collect();
    list.sort_by_key(|x| x.0);
    let mut data = Vec::with_capacity(list.len() * 5);
    let (mut last_line, mut last_col) = (0u32, 0u32);
    let mut last_end = 0u32;
    for (start, (len, ty, mods)) in list {
        if start < last_end {
            continue; // overlapping
        }
        let text = md
            .text
            .get(start as usize..(start + len) as usize)
            .unwrap_or("");
        if text.contains('\n') {
            continue;
        }
        let (line, col) = md.line_col_utf16(start);
        let delta_col = if line == last_line {
            col - last_col
        } else {
            col
        };
        data.extend([
            line - last_line,
            delta_col,
            text.encode_utf16().count() as u32,
            ty,
            mods,
        ]);
        (last_line, last_col, last_end) = (line, col, start + len);
    }
    data
}

// ---------------------------------------------------------------------------- folding

/// A folding range: 0-based first and last line, and its LSP kind (`comment`, `imports`, `region`).
pub struct Fold {
    pub start: u32,
    pub end: u32,
    pub kind: Option<&'static str>,
}

/// The folding ranges of module `m`: brackets spanning lines, comments, the imports and
/// `// #region` … `// #endregion`.
pub fn folding_ranges(prog: &Program, m: ModuleId) -> Vec<Fold> {
    let md = &prog.modules[m];
    let line = |off: u32| md.line_of(off) - 1;
    let mut out = Vec::new();
    let mut stack: Vec<u32> = Vec::new();
    let mut regions: Vec<u32> = Vec::new();
    // Consecutive line comments: (first line, last line).
    let mut comments: Option<(u32, u32)> = None;
    let flush = |c: &mut Option<(u32, u32)>, out: &mut Vec<Fold>| {
        if let Some((a, b)) = c.take() {
            if b > a {
                out.push(Fold {
                    start: a,
                    end: b,
                    kind: Some("comment"),
                });
            }
        }
    };
    for t in ras_syntax::lex(&md.text) {
        let text = &md.text[t.start as usize..t.end as usize];
        match t.kind {
            SyntaxKind::L_BRACE | SyntaxKind::L_BRACK | SyntaxKind::L_PAREN => stack.push(t.start),
            SyntaxKind::R_BRACE | SyntaxKind::R_BRACK | SyntaxKind::R_PAREN => {
                if let Some(open) = stack.pop() {
                    let (a, b) = (line(open), line(t.start));
                    // A closing bracket at the start of its line stays visible.
                    let leads = md.text[..t.start as usize]
                        .rsplit('\n')
                        .next()
                        .is_some_and(|before| before.trim().is_empty());
                    let end = if leads { b.saturating_sub(1) } else { b };
                    if end > a {
                        out.push(Fold {
                            start: a,
                            end,
                            kind: None,
                        });
                    }
                }
            }
            SyntaxKind::BLOCK_COMMENT => {
                flush(&mut comments, &mut out);
                let (a, b) = (line(t.start), line(t.end));
                if b > a {
                    out.push(Fold {
                        start: a,
                        end: b,
                        kind: Some("comment"),
                    });
                }
            }
            SyntaxKind::LINE_COMMENT => {
                let l = line(t.start);
                let body = text.trim_start_matches('/').trim_start();
                if body.starts_with("#region") {
                    flush(&mut comments, &mut out);
                    regions.push(l);
                    continue;
                }
                if body.starts_with("#endregion") {
                    flush(&mut comments, &mut out);
                    if let Some(a) = regions.pop() {
                        out.push(Fold {
                            start: a,
                            end: l,
                            kind: Some("region"),
                        });
                    }
                    continue;
                }
                // Only comments on lines of their own fold together.
                let own_line = md.text[..t.start as usize]
                    .rsplit('\n')
                    .next()
                    .is_some_and(|b| b.trim().is_empty());
                match comments {
                    Some((a, b)) if own_line && l == b + 1 => comments = Some((a, l)),
                    _ => {
                        flush(&mut comments, &mut out);
                        if own_line {
                            comments = Some((l, l));
                        }
                    }
                }
            }
            SyntaxKind::WHITESPACE => {}
            _ => flush(&mut comments, &mut out),
        }
    }
    flush(&mut comments, &mut out);
    // The imports at the top.
    let imports: Vec<Span> = md
        .ast
        .items
        .iter()
        .filter_map(|i| {
            if let Item::Import(imp) = i {
                Some(imp.span)
            } else {
                None
            }
        })
        .collect();
    if let (Some(first), Some(last)) = (imports.first(), imports.last()) {
        let (a, b) = (
            line(first.start().into()),
            line(u32::from(last.end()).saturating_sub(1)),
        );
        if b > a {
            out.push(Fold {
                start: a,
                end: b,
                kind: Some("imports"),
            });
        }
    }
    out.sort_by_key(|f| (f.start, std::cmp::Reverse(f.end)));
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{analyze, Config, SourceInput};
    use std::path::PathBuf;

    fn program(text: &str) -> Program {
        let src = SourceInput {
            path: PathBuf::from("/p/src/main.ras"),
            rel_path: "src/main.ras".into(),
            key: "main".into(),
            kind: ModuleKind::User,
            text: text.into(),
        };
        analyze(vec![src], Config::default())
    }

    #[test]
    fn hints_for_inferred_types_type_arguments_and_byref() {
        let src = "function first<T>(xs: T[]): T {\n    return xs[0];\n}\n\nfunction bump(byref n: int, step: int) {\n    n += step;\n}\n\nconst limit = 10;\n\nexport function run() {\n    let total = 0;\n    const p = new Collection();\n    const names = [\"a\", \"b\"].map(s => s + \"!\");\n    for (const n of names) {\n        bump(total, n.length);\n    }\n    bump(step: 1, n: total);\n    const head = first(names);\n}\n";
        let prog = program(src);
        let m = prog.module_by_key("main").unwrap();
        let hints = inlay_hints(&prog, m, 0, src.len() as u32);
        let shown: Vec<(String, String)> = hints
            .iter()
            .map(|h| {
                (
                    src[..h.offset as usize]
                        .lines()
                        .last()
                        .unwrap_or("")
                        .trim()
                        .to_string(),
                    h.label.clone(),
                )
            })
            .collect();
        let has = |before: &str, label: &str| {
            shown.iter().any(|(b, l)| b.ends_with(before) && l == label)
        };
        assert!(has("const limit", ": int"), "{shown:?}");
        assert!(has("let total", ": int"), "{shown:?}");
        assert!(
            !shown.iter().any(|(b, _)| b.ends_with("const p")),
            "`new` shows its type: {shown:?}"
        );
        assert!(has("const names", ": string[]"), "{shown:?}");
        assert!(has("map(s", ": string"), "lambda parameter: {shown:?}");
        assert!(has("for (const n", ": string"), "{shown:?}");
        assert!(has("bump(", "byref"), "{shown:?}");
        assert!(
            has("bump(step: 1, n:", "byref"),
            "named arguments: {shown:?}"
        );
        assert!(has("first", "<string>"), "{shown:?}");
        assert_eq!(hints.iter().filter(|h| h.label == "byref").count(), 2);
    }

    #[test]
    fn semantic_tokens_classify_names() {
        let src = "function add(byref acc: int, x: int) {\n    acc += x;\n}\n\nexport function isBlank(this: string): boolean {\n    return this.length == 0;\n}\n\nexport function run(ws: Excel.Worksheet) {\n    let n = 0;\n    add(n, ws.cells(1, 1).row);\n    const b = \"\".isBlank();\n}\n";
        let prog = program(src);
        let m = prog.module_by_key("main").unwrap();
        let data = semantic_tokens(&prog, m);
        assert_eq!(data.len() % 5, 0);
        // Decode back to (text, type, modifiers).
        let md = &prog.modules[m];
        let mut out = Vec::new();
        let (mut line, mut col) = (0u32, 0u32);
        for t in data.chunks(5) {
            if t[0] > 0 {
                col = 0;
            }
            line += t[0];
            col += t[1];
            let start = md.offset_of(line, col) as usize;
            let end = md.offset_of(line, col + t[2]) as usize;
            out.push((
                md.text[start..end].to_string(),
                TOKEN_TYPES[t[3] as usize],
                t[4],
            ));
        }
        let find = |text: &str, n: usize| {
            out.iter()
                .filter(|o| o.0 == text)
                .nth(n)
                .cloned()
                .unwrap_or_else(|| panic!("{text}: {out:?}"))
        };
        assert_eq!(find("acc", 0).1, "parameter");
        assert!(find("acc", 0).2 & tm::BYREF != 0 && find("acc", 0).2 & tm::DECLARATION != 0);
        assert!(find("acc", 1).2 & tm::BYREF != 0 && find("acc", 1).2 & tm::DECLARATION == 0);
        assert_eq!(find("add", 1).1, "function");
        assert_eq!(find("isBlank", 1).1, "method");
        assert!(find("isBlank", 1).2 & tm::EXTENSION != 0);
        assert_eq!(find("Excel", 0).1, "namespace");
        assert_eq!(find("Worksheet", 0).1, "class");
        assert!(find("cells", 0).2 & tm::DEFAULT_LIBRARY != 0);
        assert_eq!(find("row", 0).1, "property");
        assert_eq!(find("n", 0).1, "variable");
    }

    #[test]
    fn folds_blocks_comments_and_imports() {
        let src = "// a\n// b\n/**\n * doc\n */\nexport function run() {\n    const xs = [\n        1,\n    ];\n    // #region setup\n    msgBox(\"x\");\n    // #endregion\n}\n";
        let prog = program(src);
        let m = prog.module_by_key("main").unwrap();
        let folds: Vec<(u32, u32, Option<&str>)> = folding_ranges(&prog, m)
            .into_iter()
            .map(|f| (f.start, f.end, f.kind))
            .collect();
        assert!(folds.contains(&(0, 1, Some("comment"))), "{folds:?}");
        assert!(folds.contains(&(2, 4, Some("comment"))), "{folds:?}");
        assert!(folds.contains(&(5, 11, None)), "{folds:?}");
        assert!(folds.contains(&(6, 7, None)), "{folds:?}");
        assert!(folds.contains(&(9, 11, Some("region"))), "{folds:?}");
    }
}
