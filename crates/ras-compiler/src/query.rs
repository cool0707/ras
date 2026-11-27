//! Position-based queries on an analysed program (hover, completion, go to
//! definition, document symbols) — the semantic half of the language server.

use ras_syntax::ast::{self, Expr, ExprKind, Item, PassMode, Span};

use crate::astutil::{walk_expr, walk_stmt, walk_stmts};
use crate::formula;
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

pub(crate) fn vba_type(t: &Type) -> Option<&'static str> {
    Some(match t {
        Type::Byte => "Byte",
        Type::Short => "Integer",
        Type::Int => "Long",
        Type::Long => "LongLong",
        Type::Float => "Single",
        Type::Double => "Double",
        Type::Currency => "Currency",
        Type::Decimal | Type::Any | Type::CellValue => "Variant",
        Type::Ptr => "LongPtr",
        Type::Date => "Date",
        Type::Str => "String",
        Type::Bool => "Boolean",
        Type::Object | Type::Map(..) | Type::Func(_) => "Object",
        Type::Union(_) if t.is_object() => "Object",
        Type::Union(_) => "Variant",
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
        FnKind::Method { .. }
        | FnKind::Ctor { .. }
        | FnKind::Abstract { .. }
        | FnKind::Host { class: Some(_) } => "",
        _ => "function ",
    });
    s.push_str(if d.name.is_empty() {
        "(既定メンバー)"
    } else {
        &d.name
    });
    if !d.type_params.is_empty() {
        s.push('<');
        s.push_str(
            &d.type_params
                .iter()
                .map(|p| p.name.clone())
                .collect::<Vec<_>>()
                .join(", "),
        );
        s.push('>');
    }
    s.push('(');
    let ps: Vec<String> = d.params.iter().map(|p| param_text(prog, d, p)).collect();
    s.push_str(&ps.join(", "));
    s.push_str("): ");
    s.push_str(&prog.show(&d.ret_ty()));
    s
}

/// `byref name?: type` of one parameter.
fn param_text(prog: &Program, d: &FnDef, p: &ParamDef) -> String {
    let mut t = String::new();
    if p.mode == PassMode::ByRef {
        t.push_str("byref ");
    }
    if p.rest {
        t.push_str("...");
    }
    t.push_str(
        if d.locals
            .get(p.local)
            .is_some_and(|l| l.kind == LocalKind::This)
        {
            "this"
        } else {
            &p.name
        },
    );
    if p.optional || p.default.is_some() {
        t.push('?');
    }
    t.push_str(": ");
    t.push_str(&prog.show(&p.ty));
    t
}

/// The signature of a member the compiler adds (spec §9.4), for hover and parameter information.
pub struct BuiltinSig {
    pub name: &'static str,
    /// Per parameter: its text in the label, its name, its description.
    pub params: &'static [(&'static str, &'static str, &'static str)],
    pub ret: &'static str,
    pub doc: &'static str,
    /// Whether the parameters can be given by name (spec §3.3).
    pub named: bool,
}

impl BuiltinSig {
    pub fn label(&self) -> String {
        let ps: Vec<&str> = self.params.iter().map(|p| p.0).collect();
        format!("{}({}): {}", self.name, ps.join(", "), self.ret)
    }
}

/// `range.extend(up, down, left, right)` (spec §9.4.1).
pub const RANGE_EXTEND_SIG: BuiltinSig = BuiltinSig {
    name: "extend",
    params: &[
        ("up?: int", "up", "上に足す行数（負の数で縮める。省略は 0）"),
        ("down?: int", "down", "下に足す行数（負の数で縮める。省略は 0）"),
        ("left?: int", "left", "左に足す列数（負の数で縮める。省略は 0）"),
        ("right?: int", "right", "右に足す列数（負の数で縮める。省略は 0）"),
    ],
    ret: "Excel.Range",
    doc: "範囲の上・下・左・右に行・列を足した範囲を返す（元の範囲は変えない）。名前付き引数でも書ける。シートの端を越える・大きさが 0 以下になると実行時エラー 1004。\n\n例: `range(\"B3:C3\").extend(down: 2)` → B3:C5、`extend(1, 1, 1, 1)` → A2:D4",
    named: true,
};

/// `ws.range(top, left, bottom, right)` (spec §9.4.2).
pub const SHEET_RANGE_SIG: BuiltinSig = BuiltinSig {
    name: "range",
    params: &[
        ("top: int", "top", "上の行の番号"),
        ("left: int", "left", "左の列の番号"),
        ("bottom: int", "bottom", "下の行の番号"),
        ("right: int", "right", "右の列の番号"),
    ],
    ret: "Excel.Range",
    doc: "(top, left) のセルから (bottom, right) のセルまでの範囲を返す（RAS の拡張。VBA では `Range(Cells(top, left), Cells(bottom, right))`）。\n\n例: `Sheet1.range(2, 4, 8, 9)` → D2:I8",
    named: false,
};

/// The signature of an intrinsic member, when it has one to show.
fn intrinsic_sig(i: intrinsics::Intr) -> Option<&'static BuiltinSig> {
    match i {
        intrinsics::Intr::RangeExtend => Some(&RANGE_EXTEND_SIG),
        _ => None,
    }
}

/// The signatures RAS adds to host function `f` (overloads; spec §9.4.2).
fn extra_sigs(prog: &Program, f: FnId) -> &'static [BuiltinSig] {
    if prog.is_sheet_range(f) {
        std::slice::from_ref(&SHEET_RANGE_SIG)
    } else {
        &[]
    }
}

/// A hover of `sig`: its label, then its description and parameters.
fn builtin_hover(sig: &BuiltinSig, vba: Option<&str>) -> String {
    let mut code = format!("(組み込み) {}", sig.label());
    if let Some(v) = vba {
        code.push_str("\n// VBA: ");
        code.push_str(v);
    }
    let mut doc = sig.doc.to_string();
    for (_, n, d) in sig.params {
        doc.push_str(&format!("\n\n*@param* `{n}` — {d}"));
    }
    code_block(&code, Some(&doc))
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
            // A narrowed local (spec §2.6.4) is still stored in its declared type.
            let head = format!("{kw} {}", d.name);
            match vba_type(&d.ty) {
                Some(v) if *ty != d.ty => code_block(
                    &format!(
                        "{head}: {}\n// VBA: {v}（宣言は `{}`。ここでは型が絞り込まれています）",
                        prog.show(ty),
                        prog.show(&d.ty)
                    ),
                    None,
                ),
                _ => code_block(&with_vba(prog, &head, ty), None),
            }
        }
        Res::Global(v) => {
            let d = &prog.vars[*v];
            let head = format!("{} {}", if d.is_const { "const" } else { "let" }, d.name);
            match &d.host_name {
                // A host global: the VBA line shows its spelling there instead of the type.
                Some(host) => code_block(
                    &format!("{head}: {}\n// VBA: {host}", prog.show(&d.ty)),
                    d.doc.as_deref(),
                ),
                None => code_block(&with_vba(prog, &head, &d.ty), d.doc.as_deref()),
            }
        }
        Res::Fn(f) | Res::Method(f) => {
            let d = &prog.fns[*f];
            match (&d.host_name, d.kind.clone()) {
                (Some(host), FnKind::Host { class: Some(c) }) => {
                    let mut code = signature(prog, *f);
                    let mut doc = d.doc.clone().unwrap_or_default();
                    // The forms RAS adds (spec §9.4.2), as overloads.
                    for x in extra_sigs(prog, *f) {
                        code.push('\n');
                        code.push_str(&x.label());
                        doc.push_str(&format!("\n\n`{}`: {}", x.label(), x.doc));
                    }
                    code_block(
                        &format!("{code}\n// VBA: {}.{host}", prog.classes[c].name),
                        Some(doc.trim()).filter(|d| !d.is_empty()),
                    )
                }
                (Some(host), _) => code_block(
                    &format!("{}\n// VBA: {host}", signature(prog, *f)),
                    d.doc.as_deref(),
                ),
                _ => code_block(&signature(prog, *f), d.doc.as_deref()),
            }
        }
        Res::ExtMethod(f) => {
            let d = &prog.fns[*f];
            let md = &prog.modules[d.module];
            code_block(
                &format!(
                    "{}\n// 拡張メソッド（{}）→ VBA: {}.{}(…)",
                    signature(prog, *f),
                    md.key,
                    md.vba_name,
                    d.name
                ),
                d.doc.as_deref(),
            )
        }
        Res::Property { get, set } => {
            let f = get.or(*set)?;
            code_block(
                &format!("(プロパティ) {}: {}", prog.fns[f].name, prog.show(ty)),
                prog.fns[f].doc.as_deref(),
            )
        }
        Res::Field(c, i) => {
            let fl = &prog.classes[*c].fields[*i];
            let head = format!("(プロパティ) {}.{}", prog.class_display(*c), fl.name);
            match &fl.host_name {
                Some(_) => code_block(
                    &format!(
                        "{head}: {}\n// VBA: {}.{}",
                        prog.show(ty),
                        prog.classes[*c].name,
                        fl.emitted_name()
                    ),
                    fl.doc.as_deref(),
                ),
                None => code_block(&with_vba(prog, &head, ty), fl.doc.as_deref()),
            }
        }
        Res::Class(c) => {
            let d = &prog.classes[*c];
            let kw = if d.kind == ClassKind::Class {
                "class"
            } else {
                "interface"
            };
            code_block(
                &format!("{kw} {}", prog.class_display(*c)),
                d.doc.as_deref(),
            )
        }
        Res::Enum(e) => code_block(
            &format!("enum {}", prog.enums[*e].name),
            prog.enums[*e].doc.as_deref(),
        ),
        Res::EnumMember(e, i) => {
            let d = &prog.enums[*e];
            code_block(
                &format!(
                    "{}.{} = {}",
                    d.name, d.members[*i].name, d.members[*i].value
                ),
                d.members[*i].doc.as_deref(),
            )
        }
        Res::Module(t) => code_block(
            &format!(
                "module \"{}\"  // VBA: {}",
                prog.modules[*t].key, prog.modules[*t].vba_name
            ),
            None,
        ),
        Res::Namespace(n) => code_block(&format!("namespace {n}"), None),
        Res::Builtin(n) => code_block(&format!("(組み込み) {n}"), None),
        Res::Intrinsic(i) => match intrinsic_sig(*i) {
            Some(sig) => builtin_hover(sig, Some("RAS_RangeExtend(…)")),
            None => code_block(&format!("(組み込み): {}", prog.show(ty)), None),
        },
        Res::Dynamic => code_block(
            "(遅延バインディング): any",
            Some("静的な型が無いため、実行時に解決されます。"),
        ),
    })
}

pub fn hover(prog: &Program, m: ModuleId, offset: u32) -> Option<Hover> {
    let md = &prog.modules[m];
    // A function of the formula of `` fx`…` `` (spec §2.7).
    if let Some(h) = formula::hover_at(&md.text, (offset as usize).min(md.text.len())) {
        let (f, (a, b)) = h?;
        let mut markdown = format!("```\n{}\n```", f.signature);
        if !f.description.is_empty() {
            markdown.push_str("\n\n");
            markdown.push_str(&f.description);
        }
        return Some(Hover {
            span: Span::new((a as u32).into(), (b as u32).into()),
            markdown,
        });
    }
    // Declarations first: function names and local declarations are not expressions.
    for &f in &md.fns {
        let d = &prog.fns[f];
        if d.name_span.contains_inclusive(offset.into())
            && !d.name_span.is_empty()
            && !matches!(d.kind, FnKind::Main | FnKind::ModInit)
        {
            return Some(Hover {
                span: d.name_span,
                markdown: code_block(&signature(prog, f), d.doc.as_deref()),
            });
        }
    }
    if let Some(f) = fn_at(prog, m, offset) {
        for l in &prog.fns[f].locals {
            if l.span.contains_inclusive(offset.into())
                && !l.span.is_empty()
                && !l.name.starts_with("ras$")
            {
                return Some(Hover {
                    span: l.span,
                    markdown: code_block(&with_vba(prog, &l.name, &l.ty), None),
                });
            }
        }
    }
    let e = expr_at(&md.ast, offset)?;
    let ty = md.info.ty(e.id);
    let (span, target) = match &e.kind {
        ExprKind::Member { name, .. } if name.span.contains_inclusive(offset.into()) => {
            (name.span, e)
        }
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
        Res::Fn(f) | Res::Method(f) | Res::ExtMethod(f) => {
            Some((prog.fns[*f].module, prog.fns[*f].name_span))
        }
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
    /// Text inserted instead of the label (`name: ` for a named argument).
    pub insert: Option<String>,
    /// Text the editor filters on instead of the label.
    pub filter: Option<String>,
    /// The `import` that makes the name visible, inserted when the item is accepted (an
    /// export of a module that is not imported yet, spec §16.2.1).
    pub import: Option<(Span, String)>,
    /// The text the item replaces (when it is not the identifier at the cursor: a function
    /// of a formula with `.` in its name, spec §2.7).
    pub range: Option<Span>,
    /// Shows the signature once the item is accepted (a function of a formula).
    pub parameter_hints: bool,
}

fn item(
    label: &str,
    kind: u32,
    detail: impl Into<String>,
    doc: Option<&String>,
    sort: u8,
) -> Completion {
    Completion {
        label: label.to_string(),
        kind,
        detail: detail.into(),
        doc: doc.cloned(),
        sort,
        insert: None,
        filter: None,
        import: None,
        range: None,
        parameter_hints: false,
    }
}

fn sym_item(prog: &Program, name: &str, sym: &Sym, sort: u8) -> Option<Completion> {
    Some(match sym {
        Sym::Fn(f) => item(
            name,
            kind::FUNCTION,
            signature(prog, *f),
            prog.fns[*f].doc.as_ref(),
            sort,
        ),
        Sym::Class(c) => item(
            name,
            kind::CLASS,
            prog.class_display(*c),
            prog.classes[*c].doc.as_ref(),
            sort,
        ),
        Sym::Enum(e) => item(
            name,
            kind::ENUM,
            format!("enum {}", prog.enums[*e].name),
            prog.enums[*e].doc.as_ref(),
            sort,
        ),
        Sym::EnumMember(e, i) => {
            let d = &prog.enums[*e];
            item(
                name,
                kind::ENUM_MEMBER,
                format!("{} = {}", d.name, d.members[*i].value),
                d.members[*i].doc.as_ref(),
                sort,
            )
        }
        Sym::Var(v) => item(
            name,
            if prog.vars[*v].is_const {
                kind::CONSTANT
            } else {
                kind::VARIABLE
            },
            prog.show(&prog.vars[*v].ty),
            prog.vars[*v].doc.as_ref(),
            sort,
        ),
        Sym::Module(t) => item(
            name,
            kind::MODULE,
            format!("module \"{}\"", prog.modules[*t].key),
            None,
            sort,
        ),
        Sym::Namespace(n) => item(name, kind::MODULE, format!("namespace {n}"), None, sort),
        Sym::Struct(_) | Sym::Alias(_) => item(name, kind::CLASS, "型", None, sort),
    })
}

fn class_members(prog: &Program, c: ClassId, statics: bool, out: &mut Vec<Completion>) {
    let d = &prog.classes[c];
    for f in &d.fields {
        if f.is_static == statics && !out.iter().any(|o| o.label == f.name) {
            out.push(item(
                &f.name,
                kind::FIELD,
                prog.show(&f.ty),
                f.doc.as_ref(),
                0,
            ));
        }
    }
    for &m in &d.methods {
        let fd = &prog.fns[m];
        let is_static = matches!(
            fd.kind,
            FnKind::Method {
                is_static: true,
                ..
            } | FnKind::Getter {
                is_static: true,
                ..
            } | FnKind::Setter {
                is_static: true,
                ..
            }
        );
        if is_static != statics || out.iter().any(|o| o.label == fd.name) {
            continue;
        }
        let k = if matches!(fd.kind, FnKind::Getter { .. } | FnKind::Setter { .. }) {
            kind::PROPERTY
        } else {
            kind::METHOD
        };
        let mut detail = signature(prog, m);
        let extra = extra_sigs(prog, m);
        if !extra.is_empty() {
            detail.push_str(&format!("（+{} 個のオーバーロード）", extra.len()));
        }
        out.push(item(&fd.name, k, detail, fd.doc.as_ref(), 0));
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
    let builtin_sig = |sig: &BuiltinSig, out: &mut Vec<Completion>| {
        let doc = sig.doc.to_string();
        out.push(item(sig.name, kind::METHOD, sig.label(), Some(&doc), 0));
    };
    match ty {
        // The members every member type of the union has.
        Type::Union(ms) => {
            let mut lists = ms.iter().map(|t| members_of(prog, m, t));
            let first = lists.next().unwrap_or_default();
            let rest: Vec<Vec<Completion>> = lists.collect();
            out = first
                .into_iter()
                .filter(|c| rest.iter().all(|l| l.iter().any(|o| o.label == c.label)))
                .collect();
        }
        Type::Class(c, _) => {
            class_members(prog, *c, false, &mut out);
            let d = &prog.classes[*c];
            if d.kind == ClassKind::Host && d.namespace.as_deref() == Some("Excel") {
                match d.name.as_str() {
                    "Range" => {
                        builtin(&["toArray", "fromArray", "toColumn", "toRow"], &mut out);
                        builtin_sig(&RANGE_EXTEND_SIG, &mut out);
                    }
                    "Worksheet" => builtin(&["usedTable"], &mut out),
                    _ => {}
                }
            }
            // A user form (spec §14.11.9): `showModal()` anywhere, `close` in its code-behind.
            if let Some(fm) = prog.form_module_of(*c) {
                if !out.iter().any(|o| o.label == "showModal") {
                    out.push(item("showModal", kind::METHOD, "(組み込み) showModal(): モーダルで表示し、this.close(値) の値を返す（閉じたフォームはアンロードする）", None, 0));
                }
                if fm == m && !out.iter().any(|o| o.label == "close") {
                    out.push(item(
                        "close",
                        kind::METHOD,
                        "(組み込み) close(値?): フォームを閉じ、showModal() に値を返す",
                        None,
                        0,
                    ));
                }
            }
            // Handlers attached at run time (spec §12.2.3).
            if crate::lower::listen_target(prog, *c).is_some() {
                for n in intrinsics::EVENT_MEMBERS {
                    if !out.iter().any(|o| o.label == *n) {
                        out.push(item(n, kind::METHOD, "(組み込み・イベントハンドラを付ける) on(\"<イベント>\", ハンドラ): RAS_Listener", None, 0));
                    }
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
        Type::Array(_) if ty.dims() == 2 => {
            builtin(&["length", "rows", "cols", "transpose"], &mut out)
        }
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
                out.push(item(
                    &mem.name,
                    kind::ENUM_MEMBER,
                    mem.value.to_string(),
                    mem.doc.as_ref(),
                    0,
                ));
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
    if ty.is_value() && !matches!(ty, Type::Any | Type::CellValue | Type::Object | Type::Error) {
        for &f in &prog.modules[m].ext_methods {
            let d = &prog.fns[f];
            let Some(p0) = d.params.first() else { continue };
            let fits = p0.ty == *ty
                || (p0.ty.contains_param() && p0.ty.is_array() == ty.is_array())
                || matches!((&p0.ty, ty), (Type::Class(a, _), Type::Class(b, _)) if prog.class_conforms(*b, *a));
            if fits && !out.iter().any(|o| o.label == d.name && o.sort == 1) {
                out.push(item(
                    &d.name,
                    kind::METHOD,
                    format!("(拡張) {}", signature(prog, f)),
                    d.doc.as_ref(),
                    1,
                ));
            }
        }
        // Extension methods of modules not imported yet: accepting one adds the import.
        for (t, name, sym) in importable(prog, m) {
            let Sym::Fn(f) = sym else { continue };
            let d = &prog.fns[*f];
            let Some(p0) = d.params.first().filter(|_| d.is_extension) else {
                continue;
            };
            if prog.modules[m].ext_methods.contains(f) || out.iter().any(|o| o.label == *name) {
                continue;
            }
            let fits = p0.ty == *ty
                || (p0.ty.contains_param() && p0.ty.is_array() == ty.is_array())
                || matches!((&p0.ty, ty), (Type::Class(a, _), Type::Class(b, _)) if prog.class_conforms(*b, *a));
            if fits {
                let source = import_source(prog, m, t);
                let mut c = item(
                    name,
                    kind::METHOD,
                    format!(
                        "(拡張) {}  — \"{source}\" からインポート",
                        signature(prog, *f)
                    ),
                    d.doc.as_ref(),
                    2,
                );
                c.import = import_edit(prog, m, t, name);
                out.push(c);
            }
        }
    }
    out
}

/// The exports of the other project modules that module `m` can import: (module, name, symbol).
fn importable(prog: &Program, m: ModuleId) -> Vec<(ModuleId, &String, &Sym)> {
    let mut out = Vec::new();
    for (t, md) in prog.modules.iter().enumerate() {
        if t == m || md.kind != ModuleKind::User || md.form.is_some() {
            continue;
        }
        for (name, sym) in &md.exports {
            if !matches!(sym, Sym::Module(_) | Sym::Namespace(_)) {
                out.push((t, name, sym));
            }
        }
    }
    out.sort_by(|a, b| (a.1, a.0).cmp(&(b.1, b.0)));
    out
}

/// The path module `m` imports module `t` by: its path from the project root (`src/util`,
/// `src/lib/str`; spec §15.4).
pub fn import_source(prog: &Program, _m: ModuleId, t: ModuleId) -> String {
    crate::collect::root_import_path(prog, t)
}

/// The edit that imports `name` from module `t` into module `m`: added to an existing
/// `import { … } from` of that module, or a new `import` line after the others.
pub fn import_edit(prog: &Program, m: ModuleId, t: ModuleId, name: &str) -> Option<(Span, String)> {
    let md = &prog.modules[m];
    let imports: Vec<&ast::Import> = md
        .ast
        .items
        .iter()
        .filter_map(|i| {
            if let Item::Import(imp) = i {
                Some(imp)
            } else {
                None
            }
        })
        .collect();
    for imp in &imports {
        if crate::collect::import_target(prog, m, &imp.source) != Some(t) {
            continue;
        }
        if let Some(last) = imp.names.last() {
            let end = last.alias.as_ref().unwrap_or(&last.name).span.end();
            return Some((Span::empty(end), format!(", {name}")));
        }
    }
    let line = format!(
        "import {{ {name} }} from \"{}\";\n",
        import_source(prog, m, t)
    );
    let at = match imports.last() {
        // The start of the line after the last import.
        Some(last) => {
            let end = u32::from(last.span.end()) as usize;
            md.text[end..]
                .find('\n')
                .map(|i| end + i + 1)
                .unwrap_or(md.text.len())
        }
        None => 0,
    };
    let text = if at == md.text.len() && !md.text.is_empty() && !md.text.ends_with('\n') {
        format!("\n{line}")
    } else {
        line
    };
    Some((Span::empty((at as u32).into()), text))
}

/// A quick fix that imports a name the module uses without importing it.
pub struct ImportFix {
    pub title: String,
    pub span: Span,
    pub text: String,
}

/// Imports for the names in `from..=to` that are not found (`E0201` names, `E0170` types,
/// `E0202` members that an extension method of another module would provide).
pub fn import_fixes(prog: &Program, m: ModuleId, from: u32, to: u32) -> Vec<ImportFix> {
    let md = &prog.modules[m];
    let mut out: Vec<ImportFix> = Vec::new();
    for d in prog
        .diagnostics
        .iter()
        .filter(|d| d.module == Some(m) && matches!(d.code, "E0201" | "E0170" | "E0202"))
    {
        let (s, e) = (u32::from(d.span.start()), u32::from(d.span.end()));
        if e < from || to < s {
            continue;
        }
        let written = md.text.get(s as usize..e as usize).unwrap_or("");
        // The name: the identifier the span starts with (a type) or ends with (a member).
        let ident = |t: &str| {
            t.chars()
                .all(|c| c.is_alphanumeric() || c == '_' || c == '$')
                && !t.is_empty()
        };
        let name = match d.code {
            "E0170" if d.message.starts_with("型 `") => written
                .split(|c: char| !(c.is_alphanumeric() || c == '_' || c == '$'))
                .next()
                .unwrap_or(""),
            "E0201" => written,
            "E0202" => written,
            _ => continue,
        };
        if !ident(name) {
            continue;
        }
        for (t, n, sym) in importable(prog, m) {
            if n != name {
                continue;
            }
            let is_ext = matches!(sym, Sym::Fn(f) if prog.fns[*f].is_extension);
            let wanted = match d.code {
                "E0201" => {
                    matches!(sym, Sym::Fn(_) | Sym::Var(_) | Sym::Class(_) | Sym::Enum(_))
                        && !is_ext
                }
                "E0170" => matches!(
                    sym,
                    Sym::Class(_) | Sym::Enum(_) | Sym::Struct(_) | Sym::Alias(_)
                ),
                _ => is_ext,
            };
            if !wanted {
                continue;
            }
            let Some((span, text)) = import_edit(prog, m, t, name) else {
                continue;
            };
            let title = format!(
                "\"{}\" から `{name}` をインポート",
                import_source(prog, m, t)
            );
            if !out.iter().any(|o| o.title == title) {
                out.push(ImportFix { title, span, text });
            }
        }
    }
    out
}

pub(crate) const KEYWORDS: &[&str] = &[
    "let",
    "const",
    "function",
    "return",
    "if",
    "else",
    "for",
    "of",
    "while",
    "do",
    "break",
    "continue",
    "switch",
    "case",
    "default",
    "try",
    "catch",
    "finally",
    "throw",
    "new",
    "class",
    "interface",
    "enum",
    "implements",
    "import",
    "export",
    "from",
    "as",
    "this",
    "null",
    "true",
    "false",
    "byref",
    "byval",
    "inline",
    "noinline",
    "resume",
    "debugger",
    "using",
    "declare",
    "instanceof",
    "typeof",
];

/// Start of the identifier that ends at `offset`.
fn ident_start(text: &str, offset: usize) -> usize {
    let mut p = offset;
    while p > 0 {
        let Some(c) = text[..p].chars().next_back() else {
            break;
        };
        if c.is_alphanumeric() || c == '_' || c == '$' {
            p -= c.len_utf8();
        } else {
            break;
        }
    }
    p
}

/// The argument list `p` is at the start of an argument of: the offset of its `(` and the
/// arguments before the one being typed (spec §3.3). Works on the text, since the buffer
/// is usually incomplete while typing.
fn arg_list_at(text: &str, p: usize) -> Option<(usize, Vec<&str>)> {
    let before = text[..p].trim_end();
    if !before.ends_with('(') && !before.ends_with(',') {
        return None;
    }
    let (paren, mut args) = enclosing_args(text, p)?;
    args.pop(); // the argument being typed
    Some((paren, args))
}

/// The unclosed `(` before `p` and the arguments written after it up to `p`, the one being
/// typed last.
fn enclosing_args(text: &str, p: usize) -> Option<(usize, Vec<&str>)> {
    // Walk back to the unclosed `(`, splitting the arguments at the top-level commas.
    let bytes = text.as_bytes();
    let mut depth = 0usize;
    let mut end = p;
    let mut args = Vec::new();
    let mut i = p;
    while i > 0 {
        i -= 1;
        match bytes[i] {
            b')' | b']' | b'}' => depth += 1,
            b'(' | b'[' | b'{' if depth > 0 => depth -= 1,
            b'(' => {
                args.push(&text[i + 1..end]);
                args.reverse();
                return Some((i, args));
            }
            b'[' | b'{' => return None,
            b';' if depth == 0 => return None,
            b',' if depth == 0 => {
                args.push(&text[i + 1..end]);
                end = i;
            }
            q @ (b'"' | b'\'' | b'`') => {
                // Skip a string literal on the same line.
                let line_start = text[..i].rfind('\n').map_or(0, |n| n + 1);
                i = text[line_start..i]
                    .rfind(q as char)
                    .map(|j| line_start + j)?;
            }
            _ => {}
        }
    }
    None
}

/// `name` of an argument written as `name: value`.
fn arg_name(arg: &str) -> Option<&str> {
    let a = arg.trim_start();
    let n = a.len()
        - a.trim_start_matches(|c: char| c.is_alphanumeric() || c == '_' || c == '$')
            .len();
    (n > 0 && a[n..].trim_start().starts_with(':')).then(|| &a[..n])
}

/// `f<T>(`: the type arguments sit between the callee and the `(`.
fn generic_callee(text: &str, callee: Span, paren: u32) -> bool {
    let (end, paren) = (u32::from(callee.end()) as usize, paren as usize);
    end < paren
        && text[end..paren].trim_start().starts_with('<')
        && text[..paren].trim_end().ends_with('>')
}

/// Function (and the leading parameters not written in the argument list) whose argument
/// list opens at `paren`.
/// What a call calls, for the editor: a function (with the number of leading parameters not
/// written in the argument list), or a member the compiler adds.
enum Callee {
    Fn(FnId, usize),
    Builtin(&'static BuiltinSig),
}

fn callee_at(prog: &Program, m: ModuleId, paren: u32) -> Option<Callee> {
    let md = &prog.modules[m];
    let mut callee: Option<&Expr> = None;
    let mut fallback: Option<&Expr> = None;
    let mut ctor: Option<FnId> = None;
    walk_module(&md.ast, &mut |e| match &e.kind {
        ExprKind::Call {
            callee: c,
            type_args,
            ..
        } if u32::from(c.span.end()) == paren
            || (!type_args.is_empty() && generic_callee(&md.text, c.span, paren)) =>
        {
            callee = Some(c)
        }
        ExprKind::New { ty, .. } if u32::from(ty.span().end()) == paren => {
            if let Type::Class(c, _) = md.info.ty(e.id) {
                ctor = prog.classes[c].ctor;
            }
        }
        _ if u32::from(e.span.end()) == paren
            && fallback.is_none_or(|f| e.span.len() >= f.span.len()) =>
        {
            fallback = Some(e)
        }
        _ => {}
    });
    if let Some(f) = ctor {
        return Some(Callee::Fn(f, 0));
    }
    let e = callee.or(fallback)?;
    match md.info.res.get(&e.id) {
        Some(Res::Fn(f) | Res::Method(f)) => Some(Callee::Fn(*f, 0)),
        Some(Res::ExtMethod(f)) => Some(Callee::Fn(*f, 1)),
        Some(Res::Intrinsic(i)) => intrinsic_sig(*i).map(Callee::Builtin),
        // A call of a value: the default member of its class.
        _ => match md.info.ty(e.id) {
            Type::Class(c, _) => prog.classes[c].call_sig.map(|f| Callee::Fn(f, 0)),
            _ => None,
        },
    }
}

/// `name: ` items for the parameters of the call whose argument list `p` is in, not yet
/// given positionally or by name. The flag tells whether a named argument came before, so
/// that only names can follow.
fn named_arg_items(prog: &Program, m: ModuleId, p: usize) -> Option<(Vec<Completion>, bool)> {
    let text: &str = &prog.modules[m].text;
    let (paren, args) = arg_list_at(text, p)?;
    let names: Vec<&str> = args.iter().filter_map(|a| arg_name(a)).collect();
    let positional = args.iter().take_while(|a| arg_name(a).is_none()).count();
    let mut out = Vec::new();
    let (f, skip) = match callee_at(prog, m, paren as u32)? {
        Callee::Fn(f, skip) => (f, skip),
        Callee::Builtin(sig) if sig.named => {
            for (i, (text, n, _)) in sig.params.iter().enumerate() {
                if i < positional || names.contains(n) {
                    continue;
                }
                let ty = text.split_once(": ").map(|x| x.1).unwrap_or("");
                out.push(Completion {
                    label: format!("{n}:"),
                    kind: kind::VARIABLE,
                    detail: format!("{ty}（省略可）"),
                    doc: Some(code_block(&sig.label(), None)),
                    sort: 0,
                    insert: Some(format!("{n}: ")),
                    filter: Some(n.to_string()),
                    import: None,
                    range: None,
                    parameter_hints: false,
                });
            }
            return Some((out, !names.is_empty()));
        }
        Callee::Builtin(_) => return None,
    };
    let d = &prog.fns[f];
    for (i, pd) in d.params.iter().enumerate().skip(skip) {
        let is_this = d
            .locals
            .get(pd.local)
            .is_some_and(|l| l.kind == LocalKind::This);
        if pd.rest
            || is_this
            || pd.name.is_empty()
            || i - skip < positional
            || names.contains(&pd.name.as_str())
        {
            continue;
        }
        let mut detail = prog.show(&pd.ty);
        if pd.optional || pd.default.is_some() {
            detail.push_str("（省略可）");
        }
        out.push(Completion {
            label: format!("{}:", pd.name),
            kind: kind::VARIABLE,
            detail,
            doc: Some(code_block(&signature(prog, f), None)),
            sort: 0,
            insert: Some(format!("{}: ", pd.name)),
            filter: Some(pd.name.clone()),
            import: None,
            range: None,
            parameter_hints: false,
        });
    }
    Some((out, !names.is_empty()))
}

/// Only the named-argument items at `offset` (completion triggered by `(` or `,`).
pub fn named_arg_completions(prog: &Program, m: ModuleId, offset: u32) -> Vec<Completion> {
    let text: &str = &prog.modules[m].text;
    let offset = (offset as usize).min(text.len());
    if formula::formula_at(text, offset).is_some() {
        return Vec::new();
    }
    named_arg_items(prog, m, ident_start(text, offset))
        .map(|x| x.0)
        .unwrap_or_default()
}

/// A documentation comment without its `@param` lines, and the description of each parameter.
pub fn split_doc(doc: &str) -> (String, Vec<(String, String)>) {
    let mut body: Vec<&str> = Vec::new();
    let mut params: Vec<(String, String)> = Vec::new();
    let mut open = false;
    for line in doc.lines() {
        let t = line.trim();
        if let Some(rest) = t.strip_prefix("@param") {
            let mut rest = rest.trim_start();
            // JSDoc type annotations (`{string}`) are ignored.
            if rest.starts_with('{') {
                rest = rest.split_once('}').map(|x| x.1).unwrap_or("").trim_start();
            }
            let (name, desc) = rest.split_once(char::is_whitespace).unwrap_or((rest, ""));
            // `[name]` / `[name=default]`: an optional parameter.
            let name = name
                .trim_start_matches('[')
                .split(['=', ']'])
                .next()
                .unwrap_or("");
            let desc = desc.trim_start();
            let desc = desc.strip_prefix("- ").unwrap_or(desc);
            params.push((name.to_string(), desc.to_string()));
            open = true;
        } else if open && !t.is_empty() && !t.starts_with('@') {
            let last = params.last_mut().unwrap();
            if !last.1.is_empty() {
                last.1.push(' ');
            }
            last.1.push_str(t);
        } else {
            open = false;
            body.push(line);
        }
    }
    (body.join("\n").trim().to_string(), params)
}

/// Parameter information for a call (`textDocument/signatureHelp`): its signatures (more than
/// one for the overloads RAS adds, spec §9.4.2) and the one that fits the arguments.
pub struct SignatureHelp {
    pub signatures: Vec<Signature>,
    pub active_signature: u32,
}

/// One signature of [`SignatureHelp`].
pub struct Signature {
    /// `name(a: int, b?: string): boolean`
    pub label: String,
    pub doc: Option<String>,
    /// Per parameter: its range in the label in UTF-16 code units, and its `@param` description.
    pub params: Vec<((u32, u32), Option<String>)>,
    /// The parameter of the argument at the cursor (`None` when it cannot be told: a
    /// positional argument after a named one).
    pub active: Option<u32>,
}

/// The signature of the call whose argument list `offset` is in, with the parameter being
/// written (spec §16.2). Works on the text, like named-argument completion.
pub fn signature_help(prog: &Program, m: ModuleId, offset: u32) -> Option<SignatureHelp> {
    let text: &str = &prog.modules[m].text;
    if let Some(sig) = formula::signature_at(text, (offset as usize).min(text.len())) {
        let (f, params, active) = sig?;
        return Some(SignatureHelp {
            signatures: vec![Signature {
                label: f.signature.clone(),
                doc: Some(f.description.clone()).filter(|d| !d.is_empty()),
                params: params
                    .into_iter()
                    .map(|(a, b)| ((a as u32, b as u32), None))
                    .collect(),
                active: active.map(|a| a as u32),
            }],
            active_signature: 0,
        });
    }
    let (paren, args) = enclosing_args(text, (offset as usize).min(text.len()))?;
    let (current, before) = args.split_last()?;
    let (f, skip) = match callee_at(prog, m, paren as u32)? {
        Callee::Fn(f, skip) => (f, skip),
        Callee::Builtin(sig) => {
            return Some(SignatureHelp {
                signatures: vec![builtin_signature(sig, current, before)],
                active_signature: 0,
            })
        }
    };
    let mut signatures = vec![fn_signature(prog, f, skip, current, before)];
    let extra = extra_sigs(prog, f);
    signatures.extend(extra.iter().map(|x| builtin_signature(x, current, before)));
    // `ws.range(2, 4, …`: the form with numbers once there are more arguments than the host
    // member takes, or the first is a number.
    let host_max = prog.fns[f].params.len() - skip;
    let numeric = args
        .first()
        .is_some_and(|a| a.trim_start().starts_with(|c: char| c.is_ascii_digit()));
    let active_signature = if !extra.is_empty() && (args.len() > host_max || numeric) {
        1
    } else {
        0
    };
    Some(SignatureHelp {
        signatures,
        active_signature,
    })
}

/// The [`Signature`] of `sig` with the argument being written.
fn builtin_signature(sig: &BuiltinSig, current: &str, before: &[&str]) -> Signature {
    let utf16 = |s: &str| s.encode_utf16().count() as u32;
    let mut label = format!("{}(", sig.name);
    let mut params = Vec::new();
    for (i, (text, _, doc)) in sig.params.iter().enumerate() {
        if i > 0 {
            label.push_str(", ");
        }
        let start = utf16(&label);
        label.push_str(text);
        params.push(((start, utf16(&label)), Some(doc.to_string())));
    }
    label.push_str("): ");
    label.push_str(sig.ret);
    let active = match arg_name(current) {
        Some(n) if sig.named => sig.params.iter().position(|p| p.1 == n),
        Some(_) => None,
        None if before.iter().any(|a| arg_name(a).is_some()) => None,
        None if before.len() < sig.params.len() => Some(before.len()),
        None => None,
    };
    Signature {
        label,
        doc: Some(sig.doc.to_string()),
        params,
        active: active.map(|a| a as u32),
    }
}

/// The [`Signature`] of function `f` (`skip` leading parameters not written) with the argument
/// being written.
fn fn_signature(prog: &Program, f: FnId, skip: usize, current: &str, before: &[&str]) -> Signature {
    let d = &prog.fns[f];
    let utf16 = |s: &str| s.encode_utf16().count() as u32;
    let mut label = String::new();
    match d.kind {
        FnKind::Ctor { class } => {
            label.push_str("new ");
            label.push_str(&prog.class_display(class));
        }
        _ => label.push_str(if d.name.is_empty() {
            "(既定メンバー)"
        } else {
            &d.name
        }),
    }
    if !d.type_params.is_empty() {
        label.push('<');
        label.push_str(
            &d.type_params
                .iter()
                .map(|p| p.name.clone())
                .collect::<Vec<_>>()
                .join(", "),
        );
        label.push('>');
    }
    label.push('(');
    let (body, pdocs) = d.doc.as_deref().map(split_doc).unwrap_or_default();
    let mut params = Vec::new();
    for (i, pd) in d.params.iter().enumerate().skip(skip) {
        if i > skip {
            label.push_str(", ");
        }
        let start = utf16(&label);
        label.push_str(&param_text(prog, d, pd));
        let doc = pdocs
            .iter()
            .find(|(n, _)| *n == pd.name)
            .map(|x| x.1.clone())
            .filter(|s| !s.is_empty());
        params.push(((start, utf16(&label)), doc));
    }
    label.push(')');
    if !matches!(d.kind, FnKind::Ctor { .. }) {
        label.push_str(": ");
        label.push_str(&prog.show(&d.ret_ty()));
    }
    let shown = d.params.len() - skip;
    let active = match arg_name(current) {
        Some(n) => d.params.iter().skip(skip).position(|p| p.name == n),
        // Positional after a named argument: a mistake, no parameter to show.
        None if before.iter().any(|a| arg_name(a).is_some()) => None,
        None if before.len() < shown => Some(before.len()),
        None if d.params.last().is_some_and(|p| p.rest) => Some(shown - 1),
        None => None,
    };
    let mut doc = body;
    if skip == 1 {
        let recv = prog.show(&d.params[0].ty);
        let md = &prog.modules[d.module];
        doc = format!(
            "（拡張メソッド: `{recv}` に対する `{}` の関数）\n\n{doc}",
            md.key
        )
        .trim_end()
        .to_string();
    }
    Signature {
        label,
        doc: (!doc.is_empty()).then_some(doc),
        params,
        active: active.map(|a| a as u32),
    }
}

/// The declaration of the type of what is named at `offset` (`textDocument/typeDefinition`).
pub fn type_definition(prog: &Program, m: ModuleId, offset: u32) -> Option<(ModuleId, Span)> {
    use crate::refs::Target;
    let (target, _) = crate::refs::target_at(prog, m, offset)?;
    let ty = match &target {
        Target::Local { func, local } => prog.fns[*func].locals[*local].ty.clone(),
        Target::Global(v) => prog.vars[*v].ty.clone(),
        Target::Fn(f) => match prog.fns[*f].kind {
            FnKind::Setter { .. } => prog.fns[*f]
                .params
                .first()
                .map(|p| p.ty.clone())
                .unwrap_or(Type::Void),
            FnKind::Ctor { class } => Type::Class(class, Vec::new()),
            _ => prog.fns[*f].ret_ty(),
        },
        Target::Field(c, i) => prog.classes[*c].fields[*i].ty.clone(),
        Target::StructField(s, i) => prog.structs[*s].fields[*i].1.clone(),
        Target::Class(c) => Type::Class(*c, Vec::new()),
        Target::Enum(e) | Target::EnumMember(e, _) => Type::Enum(*e),
        Target::Struct(s) => Type::Struct(*s),
        Target::Alias { .. } => return None,
    };
    let mut t = &ty;
    loop {
        t = match t {
            Type::Array(e) => e,
            Type::Map(_, v) => v,
            _ => break,
        };
    }
    match t {
        Type::Class(c, _) | Type::Ns(NsRef::Class(c)) => {
            Some((prog.classes[*c].module, prog.classes[*c].name_span))
        }
        Type::Enum(e) | Type::Ns(NsRef::Enum(e)) => {
            Some((prog.enums[*e].module, prog.enums[*e].name_span))
        }
        Type::Struct(s) if prog.structs[*s].kind == StructKind::Named => {
            Some((prog.structs[*s].module, prog.structs[*s].name_span))
        }
        _ => None,
    }
}

/// `query` matches `name` when its characters appear in it in order, ignoring case.
fn fuzzy(name: &str, query: &str) -> bool {
    let mut chars = name.chars().flat_map(char::to_lowercase);
    query
        .chars()
        .flat_map(char::to_lowercase)
        .all(|q| chars.any(|c| c == q))
}

/// The declarations of the project's modules whose name matches `query`
/// (`workspace/symbol`): (module, symbol, name of the enclosing class).
pub fn workspace_symbols(prog: &Program, query: &str) -> Vec<(ModuleId, Symbol, Option<String>)> {
    let mut out = Vec::new();
    for md in prog.modules.iter().filter(|md| md.kind == ModuleKind::User) {
        for mut s in symbols(prog, md.id) {
            for c in std::mem::take(&mut s.children) {
                if fuzzy(&c.name, query) {
                    out.push((md.id, c, Some(s.name.clone())));
                }
            }
            if fuzzy(&s.name, query) {
                out.push((md.id, s, None));
            }
        }
    }
    out
}

pub fn completions(prog: &Program, m: ModuleId, offset: u32) -> Vec<Completion> {
    let md = &prog.modules[m];
    let text: &str = &md.text;
    let offset = (offset as usize).min(text.len());
    // In the formula of `` fx`…` ``: Excel's functions (spec §2.7).
    if let Some(list) = formula::completions(text, offset) {
        return list.into_iter().map(formula_item).collect();
    }
    let p = ident_start(text, offset);
    if text[..p].ends_with('.') {
        let dot = (p - 1) as u32;
        // The receiver is the largest expression that ends right at the dot.
        let mut recv: Option<&Expr> = None;
        walk_module(&md.ast, &mut |e| {
            let end = match &e.kind {
                // error recovery: `obj.` parses as a member access with an empty name
                ExprKind::Member { object, name, .. }
                    if name.text.is_empty() || u32::from(name.span.start()) > dot =>
                {
                    u32::from(object.span.end())
                }
                _ => u32::from(e.span.end()),
            };
            if end == dot
                && !matches!(&e.kind, ExprKind::Member { name, .. } if name.text.is_empty())
                && recv.is_none_or(|r| e.span.len() >= r.span.len())
            {
                recv = Some(e);
            }
        });
        return match recv {
            Some(e) => members_of(prog, m, &md.info.ty(e.id)),
            None => Vec::new(),
        };
    }
    // Named arguments (spec §3.3) come first; after one, only names can follow.
    let mut out = Vec::new();
    if let Some((named, only)) = named_arg_items(prog, m, p) {
        out = named;
        if only {
            return out;
        }
    }
    if let Some(f) = fn_at(prog, m, offset as u32) {
        for l in &prog.fns[f].locals {
            if !l.name.is_empty()
                && !l.name.starts_with("ras$")
                && u32::from(l.span.start()) < offset as u32
                && !out.iter().any(|o: &Completion| o.label == l.name)
                && l.kind != LocalKind::This
            {
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
    // Exports of modules not imported yet (spec §16.2.1): accepting one adds the import.
    let mut offered: Vec<&str> = Vec::new();
    for (t, name, sym) in importable(prog, m) {
        if md.scope.contains_key(name)
            || offered.contains(&name.as_str())
            || matches!(sym, Sym::Fn(f) if prog.fns[*f].is_extension)
        {
            continue;
        }
        if let Some(mut c) = sym_item(prog, name, sym, 3) {
            c.detail = format!(
                "{}  — \"{}\" からインポート",
                c.detail,
                import_source(prog, m, t)
            );
            c.import = import_edit(prog, m, t, name);
            offered.push(name);
            out.push(c);
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
    for k in KEYWORDS
        .iter()
        .chain(crate::collect::PRIMITIVE_NAMES.iter())
    {
        out.push(item(k, kind::KEYWORD, "", None, 4));
    }
    out
}

/// A worksheet function offered in a formula: `NAME(` replacing the word written so far.
fn formula_item(c: formula::FormulaCompletion) -> Completion {
    let f = c.function;
    let mut doc = format!("```\n{}\n```", f.signature);
    if !f.description.is_empty() {
        doc.push_str("\n\n");
        doc.push_str(&f.description);
    }
    let mut it = item(&f.name, kind::FUNCTION, "Excel の関数", Some(&doc), 0);
    it.insert = Some(format!("{}(", f.name));
    it.range = Some(Span::new(
        (c.replace.0 as u32).into(),
        (c.replace.1 as u32).into(),
    ));
    it.parameter_hints = true;
    it
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
        out.push(Symbol {
            name: d.name.clone(),
            kind: symbol_kind::FUNCTION,
            detail: signature(prog, f),
            span,
            name_span: d.name_span,
            children: Vec::new(),
        });
    }
    for &c in &md.classes {
        let d = &prog.classes[c];
        let mut children = Vec::new();
        for fl in &d.fields {
            children.push(Symbol {
                name: fl.name.clone(),
                kind: symbol_kind::FIELD,
                detail: prog.show(&fl.ty),
                span: fl.span,
                name_span: fl.span,
                children: Vec::new(),
            });
        }
        for &me in d.methods.iter().chain(d.ctor.iter()) {
            let fd = &prog.fns[me];
            let span = fd.decl.as_ref().map(|x| x.span).unwrap_or(fd.name_span);
            children.push(Symbol {
                name: fd.name.clone(),
                kind: symbol_kind::METHOD,
                detail: signature(prog, me),
                span,
                name_span: fd.name_span,
                children: Vec::new(),
            });
        }
        let span = children
            .iter()
            .fold(d.name_span, |acc, ch| acc.cover(ch.span));
        let kind = if d.kind == ClassKind::Class {
            symbol_kind::CLASS
        } else {
            symbol_kind::INTERFACE
        };
        out.push(Symbol {
            name: d.name.clone(),
            kind,
            detail: String::new(),
            span,
            name_span: d.name_span,
            children,
        });
    }
    for &e in &md.enums {
        let d = &prog.enums[e];
        out.push(Symbol {
            name: d.name.clone(),
            kind: symbol_kind::ENUM,
            detail: String::new(),
            span: d.name_span,
            name_span: d.name_span,
            children: Vec::new(),
        });
    }
    for &v in &md.vars {
        let d = &prog.vars[v];
        let kind = if d.is_const {
            symbol_kind::CONSTANT
        } else {
            symbol_kind::VARIABLE
        };
        out.push(Symbol {
            name: d.name.clone(),
            kind,
            detail: prog.show(&d.ty),
            span: d.name_span,
            name_span: d.name_span,
            children: Vec::new(),
        });
    }
    out.sort_by_key(|s| u32::from(s.span.start()));
    out
}
