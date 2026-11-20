//! Find all references, document highlights and rename (spec §16.2): the declaration under
//! the cursor and every place in the project that names it.

use std::collections::HashSet;

use ras_syntax::ast::{self, Decorator, Expr, ExprId, ExprKind, Ident, Item, Span};
use ras_syntax::SyntaxKind;

use crate::astutil::{visit_module, Visit};
use crate::collect::{import_target, PRIMITIVE_NAMES};
use crate::model::*;
use crate::names;
use crate::query::fn_at;

/// A declaration that references and renames are about.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum Target {
    /// Local variable or parameter (`local` is the slot in the function's `locals`).
    Local { func: FnId, local: LocalId },
    Global(VarId),
    /// Function, method, accessor, extension method or `declare` function.
    Fn(FnId),
    Class(ClassId),
    Enum(EnumId),
    EnumMember(EnumId, usize),
    Field(ClassId, usize),
    Struct(StructId),
    StructField(StructId, usize),
    /// `type Name = …` declared in `module`.
    Alias { module: ModuleId, name: String },
}

impl Target {
    pub fn name<'s>(&'s self, prog: &'s Program) -> &'s str {
        match self {
            Target::Local { func, local } => &prog.fns[*func].locals[*local].name,
            Target::Global(v) => &prog.vars[*v].name,
            Target::Fn(f) => &prog.fns[*f].name,
            Target::Class(c) => &prog.classes[*c].name,
            Target::Enum(e) => &prog.enums[*e].name,
            Target::EnumMember(e, i) => &prog.enums[*e].members[*i].name,
            Target::Field(c, i) => &prog.classes[*c].fields[*i].name,
            Target::Struct(s) => &prog.structs[*s].name,
            Target::StructField(s, i) => &prog.structs[*s].fields[*i].0,
            Target::Alias { name, .. } => name,
        }
    }

    /// The module of the declaration.
    pub fn module(&self, prog: &Program) -> ModuleId {
        match self {
            Target::Local { func, .. } | Target::Fn(func) => prog.fns[*func].module,
            Target::Global(v) => prog.vars[*v].module,
            Target::Class(c) | Target::Field(c, _) => prog.classes[*c].module,
            Target::Enum(e) | Target::EnumMember(e, _) => prog.enums[*e].module,
            Target::Struct(s) | Target::StructField(s, _) => prog.structs[*s].module,
            Target::Alias { module, .. } => *module,
        }
    }

    /// The name in the declaration (empty when the declaration has no source).
    pub fn decl_span(&self, prog: &Program) -> Span {
        match self {
            Target::Local { func, local } => prog.fns[*func].locals[*local].span,
            Target::Global(v) => prog.vars[*v].name_span,
            Target::Fn(f) => prog.fns[*f].name_span,
            Target::Class(c) => prog.classes[*c].name_span,
            Target::Enum(e) => prog.enums[*e].name_span,
            Target::EnumMember(e, i) => prog.enums[*e].members[*i].span,
            Target::Field(c, i) => prog.classes[*c].fields[*i].span,
            Target::Struct(s) => prog.structs[*s].name_span,
            Target::StructField(s, i) => struct_decl(prog, *s).and_then(|d| d.fields.get(*i)).map(|f| f.name.span).unwrap_or_default(),
            Target::Alias { module, name } => prog.modules[*module]
                .ast
                .items
                .iter()
                .find_map(|item| match item {
                    Item::TypeAlias(a) if a.name.text == *name => Some(a.name.span),
                    _ => None,
                })
                .unwrap_or_default(),
        }
    }
}

/// How a reference is written.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Shorthand {
    No,
    /// The key of `{ x }` (a field of the struct).
    Key,
    /// The value of `{ x }` (the variable).
    Value,
}

#[derive(Clone, Debug)]
pub struct Reference {
    pub module: ModuleId,
    pub span: Span,
    /// The declaration itself.
    pub decl: bool,
    /// Assigned to (`x = …`, `x += …`, `x++`).
    pub write: bool,
    pub shorthand: Shorthand,
}

/// A replacement of `span` in `module` by `text`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Edit {
    pub module: ModuleId,
    pub span: Span,
    pub text: String,
}

fn struct_decl(prog: &Program, s: StructId) -> Option<&ast::StructDecl> {
    let d = &prog.structs[s];
    if d.kind != StructKind::Named {
        return None;
    }
    prog.modules[d.module].ast.items.iter().find_map(|item| match item {
        Item::Struct(x) if x.name.span == d.name_span => Some(x),
        _ => None,
    })
}

/// The module and name of the `type` declaration that `name` stands for in module `m`
/// (following imports).
fn alias_origin(prog: &Program, m: ModuleId, name: &str, depth: usize) -> Option<(ModuleId, String)> {
    if depth > 16 {
        return None;
    }
    let items = &prog.modules[m].ast.items;
    if items.iter().any(|i| matches!(i, Item::TypeAlias(a) if a.name.text == name)) {
        return Some((m, name.to_string()));
    }
    for item in items.iter() {
        let Item::Import(imp) = item else { continue };
        for n in &imp.names {
            if n.alias.as_ref().unwrap_or(&n.name).text == name {
                let t = import_target(prog, m, &imp.source)?;
                return alias_origin(prog, t, &n.name.text, depth + 1);
            }
        }
    }
    None
}

fn sym_target(prog: &Program, m: ModuleId, name: &str, sym: &Sym) -> Option<Target> {
    Some(match sym {
        Sym::Fn(f) => Target::Fn(*f),
        Sym::Class(c) => Target::Class(*c),
        Sym::Enum(e) => Target::Enum(*e),
        Sym::EnumMember(e, i) => Target::EnumMember(*e, *i),
        Sym::Var(v) => Target::Global(*v),
        Sym::Struct(s) => Target::Struct(*s),
        Sym::Alias(_) => {
            let (module, name) = alias_origin(prog, m, name, 0)?;
            Target::Alias { module, name }
        }
        Sym::Module(_) | Sym::Namespace(_) => return None,
    })
}

/// Names a type annotation resolves without a declaration.
const BUILTIN_TYPE_NAMES: &[&str] = &["Array", "Map", "Comparable", "Numeric", "ObjectType", "null", "undefined", "wstr", "astr", "number"];

/// What the named type `path` (written in module `m`) declares, as `collect::resolve_type` finds it.
pub(crate) fn type_target(prog: &Program, m: ModuleId, path: &[Ident], tps: &[&str]) -> Option<Target> {
    let md = &prog.modules[m];
    match path {
        [one] => {
            let n = one.text.as_str();
            if tps.contains(&n) || PRIMITIVE_NAMES.contains(&n) || BUILTIN_TYPE_NAMES.contains(&n) {
                return None;
            }
            let sym = md.scope.get(n).or_else(|| prog.globals.get(n))?;
            sym_target(prog, m, n, sym)
        }
        [ns, name] => match md.scope.get(&ns.text) {
            Some(Sym::Module(t)) => sym_target(prog, *t, &name.text, prog.modules[*t].exports.get(&name.text)?),
            _ => sym_target(prog, m, &name.text, prog.namespaces.get(&ns.text)?.get(&name.text)?),
        },
        _ => None,
    }
}

/// The identifier of a name or member expression.
fn ident_span(e: &Expr) -> Option<Span> {
    match &e.kind {
        ExprKind::Name(id) => Some(id.span),
        ExprKind::Member { name, .. } => Some(name.span),
        _ => None,
    }
}

/// What a name or member expression refers to.
pub(crate) fn res_target(prog: &Program, m: ModuleId, e: &Expr) -> Option<Target> {
    let md = &prog.modules[m];
    Some(match md.info.res.get(&e.id)? {
        Res::Local(l) => Target::Local { func: fn_at(prog, m, e.span.start().into())?, local: *l },
        Res::Global(v) => Target::Global(*v),
        Res::Fn(f) | Res::Method(f) | Res::ExtMethod(f) => Target::Fn(*f),
        Res::Property { get, set } => Target::Fn(get.or(*set)?),
        Res::Field(c, i) => Target::Field(*c, *i),
        Res::Class(c) => Target::Class(*c),
        Res::Enum(x) => Target::Enum(*x),
        Res::EnumMember(x, i) => Target::EnumMember(*x, *i),
        // A field of a struct is accessed by name.
        Res::Dynamic => match &e.kind {
            ExprKind::Member { object, name, .. } => match md.info.ty(object.id) {
                Type::Struct(s) if prog.structs[s].kind == StructKind::Named => Target::StructField(s, prog.structs[s].fields.iter().position(|f| f.0 == name.text)?),
                _ => return None,
            },
            _ => return None,
        },
        _ => return None,
    })
}

/// The function a call (or `new`) with named arguments calls.
fn callee_fn(prog: &Program, m: ModuleId, call: ExprId) -> Option<FnId> {
    match &prog.modules[m].info.calls.get(&call)?.target {
        CallTarget::Fn(f) | CallTarget::Method(f) | CallTarget::Ext(f) | CallTarget::Host(f) | CallTarget::DefaultMember(f) => Some(*f),
        CallTarget::Ctor(c) => prog.classes[*c].ctor,
        _ => None,
    }
}

/// The parameter of `f` that the named argument `name` gives.
fn named_param(prog: &Program, f: FnId, name: &str) -> Option<Target> {
    let p = prog.fns[f].params.iter().find(|p| p.name == name)?;
    Some(Target::Local { func: f, local: p.local })
}

fn call_args(e: &Expr) -> Option<&[Expr]> {
    match &e.kind {
        ExprKind::Call { args, .. } | ExprKind::New { args, .. } if args.iter().any(|a| matches!(a.kind, ExprKind::NamedArg { .. })) => Some(args),
        _ => None,
    }
}

/// The function named by `@onError(handler)` / `@fallback(fn)`: (span of the name, the name).
fn decorator_fn_name<'a>(d: &'a Decorator, text: &str) -> Option<(Span, &'a str)> {
    if !matches!(d.name.text.as_str(), "onError" | "fallback") {
        return None;
    }
    match &d.args.first()?.peel().kind {
        ExprKind::Name(id) => Some((id.span, &id.text)),
        ExprKind::Str(s) => {
            let a = d.args.first()?.peel();
            let (start, end) = (u32::from(a.span.start()) as usize, u32::from(a.span.end()) as usize);
            // Only a plain literal: its contents are the name.
            let src = text.get(start..end)?;
            (src.len() == s.len() + 2 && &src[1..src.len() - 1] == s).then(|| (Span::new(((start + 1) as u32).into(), ((end - 1) as u32).into()), s.as_str()))
        }
        _ => None,
    }
}

// ---------------------------------------------------------------------------- target at cursor

struct FindAt<'p> {
    prog: &'p Program,
    m: ModuleId,
    offset: u32,
    found: Option<(Target, Span)>,
}

impl FindAt<'_> {
    fn at(&self, s: Span) -> bool {
        !s.is_empty() && s.contains_inclusive(self.offset.into())
    }

    fn set(&mut self, t: Option<Target>, span: Span) {
        if self.found.is_none() {
            if let Some(t) = t {
                self.found = Some((t, span));
            }
        }
    }
}

impl<'a> Visit<'a> for FindAt<'_> {
    fn expr(&mut self, e: &'a Expr) {
        if self.found.is_some() {
            return;
        }
        if let Some(span) = ident_span(e).filter(|s| self.at(*s)) {
            let t = res_target(self.prog, self.m, e);
            self.set(t, span);
            return;
        }
        if let ExprKind::Object(props) = &e.kind {
            if let Type::Struct(s) = self.prog.modules[self.m].info.ty(e.id) {
                for p in props {
                    // The key of a shorthand `{ x }` is also the variable: the variable wins there.
                    if self.at(p.key.span) && p.value.span != p.key.span && self.prog.structs[s].kind == StructKind::Named {
                        let t = self.prog.structs[s].fields.iter().position(|f| f.0 == p.key.text).map(|i| Target::StructField(s, i));
                        self.set(t, p.key.span);
                    }
                }
            }
        }
        if let Some(args) = call_args(e) {
            for a in args {
                if let ExprKind::NamedArg { name, .. } = &a.kind {
                    if self.at(name.span) {
                        let t = callee_fn(self.prog, self.m, e.id).and_then(|f| named_param(self.prog, f, &name.text));
                        self.set(t, name.span);
                    }
                }
            }
        }
    }

    fn decorator(&mut self, d: &'a Decorator) {
        let md = &self.prog.modules[self.m];
        if let Some((span, name)) = decorator_fn_name(d, &md.text).filter(|(s, _)| self.at(*s)) {
            let t = match md.scope.get(name) {
                Some(Sym::Fn(f)) => Some(Target::Fn(*f)),
                _ => None,
            };
            self.set(t, span);
        }
    }

    fn named_type(&mut self, path: &'a [Ident], _span: Span, tps: &[&'a str]) {
        if let Some(last) = path.last().filter(|l| self.at(l.span)) {
            let t = type_target(self.prog, self.m, path, tps);
            self.set(t, last.span);
        }
    }
}

/// The declaration named at `offset` (at a declaration or at a use) and the span of that name.
pub fn target_at(prog: &Program, m: ModuleId, offset: u32) -> Option<(Target, Span)> {
    let md = &prog.modules[m];
    let at = |s: Span| !s.is_empty() && s.contains_inclusive(offset.into());
    // Declarations.
    for &f in &md.fns {
        let d = &prog.fns[f];
        if at(d.name_span) && !matches!(d.kind, FnKind::Main | FnKind::ModInit | FnKind::Ctor { .. }) {
            return Some((Target::Fn(f), d.name_span));
        }
    }
    for &c in &md.classes {
        let d = &prog.classes[c];
        if at(d.name_span) {
            return Some((Target::Class(c), d.name_span));
        }
        if let Some(i) = d.fields.iter().position(|f| at(f.span)) {
            return Some((Target::Field(c, i), d.fields[i].span));
        }
    }
    for &e in &md.enums {
        let d = &prog.enums[e];
        if at(d.name_span) {
            return Some((Target::Enum(e), d.name_span));
        }
        if let Some(i) = d.members.iter().position(|x| at(x.span)) {
            return Some((Target::EnumMember(e, i), d.members[i].span));
        }
    }
    for &v in &md.vars {
        if at(prog.vars[v].name_span) {
            return Some((Target::Global(v), prog.vars[v].name_span));
        }
    }
    for &s in &md.structs {
        if prog.structs[s].kind != StructKind::Named {
            continue;
        }
        if at(prog.structs[s].name_span) {
            return Some((Target::Struct(s), prog.structs[s].name_span));
        }
        if let Some(d) = struct_decl(prog, s) {
            if let Some(i) = d.fields.iter().position(|f| at(f.name.span)) {
                return Some((Target::StructField(s, i), d.fields[i].name.span));
            }
        }
    }
    for item in md.ast.items.iter() {
        match item {
            Item::TypeAlias(a) if at(a.name.span) => return Some((Target::Alias { module: m, name: a.name.text.clone() }, a.name.span)),
            Item::Import(imp) => {
                for n in &imp.names {
                    let span = match &n.alias {
                        _ if at(n.name.span) => n.name.span,
                        Some(a) if at(a.span) => a.span,
                        _ => continue,
                    };
                    let t = import_target(prog, m, &imp.source)?;
                    let sym = prog.modules[t].exports.get(&n.name.text)?;
                    return Some((sym_target(prog, t, &n.name.text, sym)?, span));
                }
            }
            _ => {}
        }
    }
    if let Some(f) = fn_at(prog, m, offset) {
        for (l, d) in prog.fns[f].locals.iter().enumerate() {
            if at(d.span) && !d.name.starts_with("ras$") && !d.name.is_empty() {
                return Some((Target::Local { func: f, local: l }, d.span));
            }
        }
    }
    // Uses.
    let mut find = FindAt { prog, m, offset, found: None };
    visit_module(&md.ast, &mut find);
    find.found
}

// ---------------------------------------------------------------------------- references

/// The members a rename has to change together with `t`: the getter and setter of a property,
/// and the members of the same name in the interfaces a class implements and the classes that
/// implement an interface (spec §16.2).
pub fn related(prog: &Program, t: &Target) -> Vec<Target> {
    let (class, name, is_static) = match t {
        Target::Fn(f) => match prog.fns[*f].kind {
            FnKind::Method { class, is_static } | FnKind::Getter { class, is_static } | FnKind::Setter { class, is_static } => (class, prog.fns[*f].name.clone(), is_static),
            FnKind::Abstract { class } => (class, prog.fns[*f].name.clone(), false),
            _ => return vec![t.clone()],
        },
        Target::Field(c, i) => (*c, prog.classes[*c].fields[*i].name.clone(), prog.classes[*c].fields[*i].is_static),
        _ => return vec![t.clone()],
    };
    let members = |c: ClassId| -> Vec<Target> {
        let d = &prog.classes[c];
        let mut out: Vec<Target> = d.fields.iter().enumerate().filter(|(_, f)| f.name == name && f.is_static == is_static).map(|(i, _)| Target::Field(c, i)).collect();
        for &me in &d.methods {
            let fd = &prog.fns[me];
            let st = matches!(fd.kind, FnKind::Method { is_static: true, .. } | FnKind::Getter { is_static: true, .. } | FnKind::Setter { is_static: true, .. });
            if fd.name == name && st == is_static && !matches!(fd.kind, FnKind::Ctor { .. }) {
                out.push(Target::Fn(me));
            }
        }
        out
    };
    let mut classes = vec![class];
    if prog.classes[class].kind != ClassKind::Host {
        loop {
            let before = classes.len();
            for d in 0..prog.classes.len() {
                if classes.contains(&d) || prog.classes[d].kind == ClassKind::Host || members(d).is_empty() {
                    continue;
                }
                if classes.iter().any(|&x| prog.class_conforms(d, x) || prog.class_conforms(x, d)) {
                    classes.push(d);
                }
            }
            if classes.len() == before {
                break;
            }
        }
    }
    let mut out: Vec<Target> = classes.iter().flat_map(|&c| members(c)).collect();
    if !out.contains(t) {
        out.insert(0, t.clone());
    }
    out
}

struct Scan<'p> {
    prog: &'p Program,
    m: ModuleId,
    targets: &'p [Target],
    has_locals: bool,
    out: Vec<Reference>,
    writes: HashSet<ExprId>,
    shorthand_keys: HashSet<u32>,
}

impl Scan<'_> {
    fn push(&mut self, span: Span, write: bool, shorthand: Shorthand) {
        self.out.push(Reference { module: self.m, span, decl: false, write, shorthand });
    }
}

impl<'a> Visit<'a> for Scan<'_> {
    fn expr(&mut self, e: &'a Expr) {
        let md = &self.prog.modules[self.m];
        match &e.kind {
            ExprKind::Assign { target, .. } => {
                self.writes.insert(target.peel().id);
            }
            ExprKind::IncDec { operand, .. } => {
                self.writes.insert(operand.peel().id);
            }
            ExprKind::Object(props) => {
                for p in props.iter().filter(|p| p.value.span == p.key.span) {
                    self.shorthand_keys.insert(p.key.span.start().into());
                }
                if let Type::Struct(s) = md.info.ty(e.id) {
                    for p in props {
                        if let Some(i) = self.prog.structs[s].fields.iter().position(|f| f.0 == p.key.text) {
                            if self.targets.contains(&Target::StructField(s, i)) {
                                let sh = if p.value.span == p.key.span { Shorthand::Key } else { Shorthand::No };
                                self.push(p.key.span, false, sh);
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        if let Some(span) = ident_span(e) {
            let local = matches!(md.info.res.get(&e.id), Some(Res::Local(_)));
            if !local || self.has_locals {
                if let Some(t) = res_target(self.prog, self.m, e) {
                    if self.targets.contains(&t) {
                        let sh = if self.shorthand_keys.contains(&span.start().into()) { Shorthand::Value } else { Shorthand::No };
                        let write = self.writes.contains(&e.id);
                        self.push(span, write, sh);
                    }
                }
            }
        }
        if let Some(args) = call_args(e) {
            if let Some(f) = callee_fn(self.prog, self.m, e.id) {
                for a in args {
                    if let ExprKind::NamedArg { name, .. } = &a.kind {
                        if named_param(self.prog, f, &name.text).is_some_and(|t| self.targets.contains(&t)) {
                            self.push(name.span, false, Shorthand::No);
                        }
                    }
                }
            }
        }
    }

    fn decorator(&mut self, d: &'a Decorator) {
        let md = &self.prog.modules[self.m];
        if let Some((span, name)) = decorator_fn_name(d, &md.text) {
            if matches!(md.scope.get(name), Some(Sym::Fn(f)) if self.targets.contains(&Target::Fn(*f))) {
                self.push(span, false, Shorthand::No);
            }
        }
    }

    fn named_type(&mut self, path: &'a [Ident], _span: Span, tps: &[&'a str]) {
        if let (Some(t), Some(last)) = (type_target(self.prog, self.m, path, tps), path.last()) {
            if self.targets.contains(&t) {
                self.push(last.span, false, Shorthand::No);
            }
        }
    }
}

fn scanned(prog: &Program, m: ModuleId) -> bool {
    !prog.modules[m].path.starts_with("<stdlib>")
}

/// Every place that names one of `targets` (declarations included), in module and source order.
pub fn references_of(prog: &Program, targets: &[Target]) -> Vec<Reference> {
    let mut out = Vec::new();
    for t in targets {
        let span = t.decl_span(prog);
        if !span.is_empty() {
            out.push(Reference { module: t.module(prog), span, decl: true, write: false, shorthand: Shorthand::No });
        }
    }
    let has_locals = targets.iter().any(|t| matches!(t, Target::Local { .. }));
    let local_modules: Vec<ModuleId> = targets.iter().filter(|t| matches!(t, Target::Local { .. })).map(|t| t.module(prog)).collect();
    for m in 0..prog.modules.len() {
        if !scanned(prog, m) {
            continue;
        }
        let md = &prog.modules[m];
        let mut scan = Scan { prog, m, targets, has_locals: has_locals && local_modules.contains(&m), out: Vec::new(), writes: HashSet::new(), shorthand_keys: HashSet::new() };
        visit_module(&md.ast, &mut scan);
        out.append(&mut scan.out);
        for item in md.ast.items.iter() {
            let Item::Import(imp) = item else { continue };
            let Some(t) = import_target(prog, m, &imp.source) else { continue };
            for n in &imp.names {
                let found = prog.modules[t].exports.get(&n.name.text).and_then(|s| sym_target(prog, t, &n.name.text, s));
                if found.is_some_and(|x| targets.contains(&x)) {
                    out.push(Reference { module: m, span: n.name.span, decl: false, write: false, shorthand: Shorthand::No });
                }
            }
        }
    }
    out.sort_by_key(|r| (r.module, u32::from(r.span.start()), !r.decl));
    out.dedup_by(|b, a| a.module == b.module && a.span == b.span);
    out
}

/// The references of the declaration named at `offset`.
pub fn references(prog: &Program, m: ModuleId, offset: u32) -> Vec<Reference> {
    match target_at(prog, m, offset) {
        Some((t, _)) => references_of(prog, &related(prog, &t)),
        None => Vec::new(),
    }
}

// ---------------------------------------------------------------------------- rename

/// Whether the declaration can be renamed in the sources (it is in the project, not bound to
/// a name outside RAS).
fn renamable(prog: &Program, t: &Target) -> Result<(), String> {
    let name = t.name(prog);
    let md = &prog.modules[t.module(prog)];
    if md.kind != ModuleKind::User || md.path.starts_with("<stdlib>") {
        return Err(format!("`{name}` は型定義（.d.ras）または同梱のライブラリで宣言されているため、名前を変更できません"));
    }
    match t {
        Target::Fn(f) => {
            let d = &prog.fns[*f];
            match &d.kind {
                FnKind::Main | FnKind::ModInit | FnKind::Ctor { .. } => return Err("この名前は変更できません".into()),
                FnKind::Declare { alias: None, .. } => return Err(format!("`{name}` は DLL の関数名です。`@alias(\"{name}\")` を付けると RAS での名前を変更できます")),
                FnKind::Method { class, .. } | FnKind::Getter { class, .. } | FnKind::Setter { class, .. } | FnKind::Abstract { class } if prog.classes[*class].kind == ClassKind::Host => {
                    return Err(format!("`{name}` は型定義のメンバーのため、名前を変更できません"));
                }
                _ => {}
            }
            if d.host_name.is_some() {
                return Err(format!("`{name}` は型定義のメンバーのため、名前を変更できません"));
            }
        }
        Target::Local { func, local } => {
            if prog.fns[*func].locals[*local].kind == LocalKind::This {
                return Err("`this` の名前は変更できません".into());
            }
        }
        Target::Field(c, _) if prog.classes[*c].kind == ClassKind::Host => {
            let is_form = md.form.as_ref().is_some_and(|f| f.class == Some(*c));
            return Err(if is_form {
                format!("`{name}` はフォームのコントロールです。名前はデザイナー（forms/{}.form.json）で変更してください", md.vba_name)
            } else {
                format!("`{name}` は型定義のメンバーのため、名前を変更できません")
            });
        }
        Target::Class(c) if prog.classes[*c].kind == ClassKind::Host => return Err(format!("`{name}` はフォームまたは型定義の型のため、名前を変更できません")),
        Target::Global(v) if prog.vars[*v].declared || prog.vars[*v].host_name.is_some() => return Err(format!("`{name}` は型定義の名前のため、名前を変更できません")),
        Target::Enum(e) | Target::EnumMember(e, _) if prog.enums[*e].host => return Err(format!("`{name}` は型定義の列挙型のため、名前を変更できません")),
        Target::Struct(s) | Target::StructField(s, _) if prog.structs[*s].kind != StructKind::Named => return Err("無名の型のフィールドは名前を変更できません".into()),
        _ => {}
    }
    if t.decl_span(prog).is_empty() {
        return Err(format!("`{name}` は名前を変更できません"));
    }
    Ok(())
}

/// The span and current name of the declaration named at `offset`, or why it cannot be renamed.
pub fn prepare_rename(prog: &Program, m: ModuleId, offset: u32) -> Result<(Span, String), String> {
    let Some((t, span)) = target_at(prog, m, offset) else {
        return Err("名前を変更できるシンボルではありません".into());
    };
    let name = t.name(prog).to_string();
    let written = prog.modules[m].text.get(usize::from(span.start())..usize::from(span.end())).unwrap_or("");
    if !written.is_empty() && written != name && !is_quoted(written, &name) {
        return Err(format!("`{written}` はインポート時の別名です。元の名前 `{name}` の宣言から名前を変更してください"));
    }
    for r in related(prog, &t) {
        renamable(prog, &r)?;
    }
    Ok((span, name))
}

fn is_quoted(written: &str, name: &str) -> bool {
    written.len() == name.len() + 2 && (written.starts_with('"') || written.starts_with('\'')) && &written[1..written.len() - 1] == name
}

/// Identifiers the case rule (§13.6, lint rule 6) compares: `(module, name span, name)`.
fn user_identifiers(prog: &Program) -> Vec<(ModuleId, Span, String)> {
    let mut items = Vec::new();
    for m in &prog.modules {
        if m.kind != ModuleKind::User {
            continue;
        }
        for &f in &m.fns {
            let d = &prog.fns[f];
            if !matches!(d.kind, FnKind::Main | FnKind::ModInit) && !d.name.is_empty() && d.name != "constructor" {
                items.push((m.id, d.name_span, d.name.clone()));
            }
            for l in &d.locals {
                if !l.name.is_empty() && l.kind != LocalKind::This {
                    items.push((m.id, l.span, l.name.clone()));
                }
            }
        }
        for &c in &m.classes {
            let d = &prog.classes[c];
            if d.kind == ClassKind::Host {
                continue;
            }
            items.push((m.id, d.name_span, d.name.clone()));
            for fld in &d.fields {
                items.push((m.id, fld.span, fld.name.clone()));
            }
        }
        for &v in &m.vars {
            items.push((m.id, prog.vars[v].name_span, prog.vars[v].name.clone()));
        }
        for &e in &m.enums {
            items.push((m.id, prog.enums[e].name_span, prog.enums[e].name.clone()));
        }
    }
    items
}

/// Names used in the expressions and types of a module that are not member names.
struct NamesUsed<'a> {
    name: &'a str,
    /// Only within this byte range (`None`: the whole module).
    within: Option<(u32, u32)>,
    found: bool,
}

impl<'a, 'b> Visit<'a> for NamesUsed<'b> {
    fn expr(&mut self, e: &'a Expr) {
        if let ExprKind::Name(id) = &e.kind {
            let s: u32 = id.span.start().into();
            if id.text == self.name && self.within.is_none_or(|(a, b)| a <= s && s <= b) {
                self.found = true;
            }
        }
    }

    fn named_type(&mut self, path: &'a [Ident], _span: Span, _tps: &[&'a str]) {
        if let Some(first) = path.first() {
            let s: u32 = first.span.start().into();
            if first.text == self.name && self.within.is_none_or(|(a, b)| a <= s && s <= b) {
                self.found = true;
            }
        }
    }
}

fn name_used(prog: &Program, m: ModuleId, name: &str, within: Option<(u32, u32)>) -> bool {
    let mut v = NamesUsed { name, within, found: false };
    visit_module(&prog.modules[m].ast, &mut v);
    v.found
}

/// Words that can come between a documentation comment and the name of a function.
const MODIFIERS: &[&str] = &["export", "declare", "private", "public", "protected", "static", "readonly", "inline", "noinline", "function", "get", "set", "async"];

/// The `name` after each `@param` of the documentation comment of `decl`.
fn doc_param_spans(text: &str, decl: &ast::FnDecl, name: &str) -> Vec<Span> {
    let start = usize::from(decl.span.start()).min(text.len());
    let Some(close) = text[..start].rfind("*/") else { return Vec::new() };
    if !text[close + 2..start].split_whitespace().all(|w| MODIFIERS.contains(&w)) {
        return Vec::new();
    }
    let Some(open) = text[..close].rfind("/**") else { return Vec::new() };
    let doc = &text[open..close];
    let mut out = Vec::new();
    let mut from = 0;
    while let Some(i) = doc[from..].find("@param") {
        let mut p = from + i + "@param".len();
        from = p;
        let skip_ws = |p: usize| p + doc[p..].len() - doc[p..].trim_start_matches([' ', '\t']).len();
        p = skip_ws(p);
        // JSDoc type annotation: `@param {string} name`.
        if doc[p..].starts_with('{') {
            let Some(end) = doc[p..].find('}') else { continue };
            p = skip_ws(p + end + 1);
        }
        let rest = &doc[p..];
        let next = rest[name.len().min(rest.len())..].chars().next();
        if rest.starts_with(name) && !next.is_some_and(|c| c.is_alphanumeric() || c == '_' || c == '$') {
            let s = (open + p) as u32;
            out.push(Span::new(s.into(), (s + name.len() as u32).into()));
        }
    }
    out
}

/// Checks a new name for the declaration: an identifier that is not a keyword or a VBA
/// reserved word and does not collide with another name (spec §13.6, §16.2).
fn check_new_name(prog: &Program, targets: &[Target], refs: &[Reference], new: &str) -> Result<(), String> {
    let toks: Vec<_> = ras_syntax::lex(new).into_iter().filter(|t| !t.kind.is_trivia()).collect();
    let whole = toks.len() == 1 && toks[0].start == 0 && toks[0].end as usize == new.len();
    if !whole || !(toks[0].kind == SyntaxKind::IDENT || toks[0].kind.is_keyword()) {
        return Err(format!("`{new}` は識別子として使えません"));
    }
    if toks[0].kind.is_keyword() || names::is_ras_keyword(new) || crate::query::KEYWORDS.contains(&new) || PRIMITIVE_NAMES.contains(&new) || BUILTIN_TYPE_NAMES.contains(&new) {
        return Err(format!("`{new}` は RAS のキーワードまたは組み込みの型名のため使えません"));
    }
    if names::is_reserved(new) {
        return Err(format!("`{new}` は VBA の予約語のため使えません"));
    }
    let lower = new.to_lowercase();
    if lower.starts_with("ras$") || lower.starts_with("ras_") {
        return Err(format!("`{new}`: `RAS_` で始まる名前はコンパイラが生成するコードのために予約されています"));
    }
    // Identifiers that differ only by case collide in VBA (spec §13.6).
    let own: HashSet<(ModuleId, u32)> = targets.iter().map(|t| (t.module(prog), u32::from(t.decl_span(prog).start()))).collect();
    for (m, span, name) in user_identifiers(prog) {
        if !own.contains(&(m, u32::from(span.start()))) && name != new && name.to_lowercase() == lower {
            return Err(format!("`{new}` は `{name}` と大文字小文字だけが異なります。VBA は大文字小文字を区別しないため使えません"));
        }
    }
    let primary = &targets[0];
    match primary {
        Target::Local { func, .. } => {
            let d = &prog.fns[*func];
            if d.locals.iter().any(|l| l.name == new) {
                return Err(format!("`{new}` は `{}` の中で既に使われています", fn_label(prog, *func)));
            }
            let within = match (&d.kind, &d.decl) {
                (FnKind::Main | FnKind::ModInit, _) | (_, None) => None,
                (_, Some(decl)) => Some((u32::from(decl.span.start()), u32::from(decl.span.end()))),
            };
            if name_used(prog, d.module, new, within) {
                return Err(format!("`{new}` は `{}` の中で別のものを指す名前として使われているため使えません", fn_label(prog, *func)));
            }
        }
        Target::Fn(_) | Target::Field(..) if related_class(prog, primary).is_some() => {
            for t in targets {
                let Some(c) = related_class(prog, t) else { continue };
                let has = prog.find_field(c, new).is_some() || !prog.find_methods(c, new).is_empty();
                if has {
                    return Err(format!("`{}` には既に `{new}` があります", prog.class_display(c)));
                }
            }
        }
        Target::EnumMember(e, _) => {
            if prog.enums[*e].members.iter().any(|x| x.name == new) {
                return Err(format!("列挙型 `{}` には既に `{new}` があります", prog.enums[*e].name));
            }
        }
        Target::StructField(s, _) => {
            if prog.structs[*s].fields.iter().any(|f| f.0 == new) {
                return Err(format!("構造体 `{}` には既に `{new}` があります", prog.structs[*s].name));
            }
        }
        // Module-level declarations: the module, and the modules that import it by its name.
        _ => {
            let home = primary.module(prog);
            let old = primary.name(prog);
            let mut affected = vec![home];
            for (m, md) in prog.modules.iter().enumerate() {
                for item in md.ast.items.iter() {
                    let Item::Import(imp) = item else { continue };
                    if import_target(prog, m, &imp.source) == Some(home) && imp.names.iter().any(|n| n.name.text == old && n.alias.is_none()) && !affected.contains(&m) {
                        affected.push(m);
                    }
                }
            }
            // A code-behind is the form's code module: its controls and the UserForm members
            // share the names of its procedures and variables.
            if let Some(c) = prog.modules[home].form.as_ref().and_then(|f| f.class) {
                if prog.find_field(c, new).is_some() || !prog.find_methods(c, new).is_empty() {
                    return Err(format!("`{new}` はフォーム `{}` のコントロールまたはメンバーと同じ名前のため使えません", prog.class_display(c)));
                }
            }
            for &m in &affected {
                let md = &prog.modules[m];
                if md.scope.contains_key(new) {
                    return Err(format!("`{new}` はモジュール `{}` で既に宣言されています", md.rel_path));
                }
                if name_used(prog, m, new, None) {
                    return Err(format!("`{new}` はモジュール `{}` で別のもの（Excel / VBA の名前など）を指して使われているため使えません", md.rel_path));
                }
            }
            // A local of the same name would hide the declaration at a use.
            for r in refs.iter().filter(|r| !r.decl) {
                if let Some(f) = fn_at(prog, r.module, r.span.start().into()) {
                    if prog.fns[f].locals.iter().any(|l| l.name == new) {
                        return Err(format!("`{new}` は `{}` のローカル変数と重なるため使えません", fn_label(prog, f)));
                    }
                }
            }
        }
    }
    Ok(())
}

/// The class a member target belongs to.
fn related_class(prog: &Program, t: &Target) -> Option<ClassId> {
    match t {
        Target::Field(c, _) => Some(*c),
        Target::Fn(f) => match prog.fns[*f].kind {
            FnKind::Method { class, .. } | FnKind::Getter { class, .. } | FnKind::Setter { class, .. } | FnKind::Abstract { class } => Some(class),
            _ => None,
        },
        _ => None,
    }
}

fn fn_label(prog: &Program, f: FnId) -> String {
    let d = &prog.fns[f];
    match d.kind {
        FnKind::Main | FnKind::ModInit => format!("モジュール {}", prog.modules[d.module].rel_path),
        _ => d.name.clone(),
    }
}

/// The edits that rename the declaration named at `offset` to `new`, or why it cannot be renamed.
pub fn rename(prog: &Program, m: ModuleId, offset: u32, new: &str) -> Result<Vec<Edit>, String> {
    let (_, old) = prepare_rename(prog, m, offset)?;
    let (t, _) = target_at(prog, m, offset).ok_or("名前を変更できるシンボルではありません")?;
    let targets = related(prog, &t);
    if new == old {
        return Ok(Vec::new());
    }
    let refs = references_of(prog, &targets);
    check_new_name(prog, &targets, &refs, new)?;
    let mut edits = Vec::new();
    for r in &refs {
        let md = &prog.modules[r.module];
        let Some(written) = md.text.get(usize::from(r.span.start())..usize::from(r.span.end())) else { continue };
        let text = match r.shorthand {
            Shorthand::Key => format!("{new}: {written}"),
            Shorthand::Value => format!("{written}: {new}"),
            Shorthand::No if written == old => new.to_string(),
            Shorthand::No if is_quoted(written, &old) => format!("{q}{new}{q}", q = &written[..1]),
            // An import alias (`import { a as b }`): its uses keep the alias.
            Shorthand::No => continue,
        };
        edits.push(Edit { module: r.module, span: r.span, text });
    }
    // `@param old` in the documentation comment of the function.
    if let Target::Local { func, local } = &t {
        let d = &prog.fns[*func];
        if let (true, Some(decl)) = (d.params.iter().any(|p| p.local == *local), &d.decl) {
            for span in doc_param_spans(&prog.modules[d.module].text, decl, &old) {
                edits.push(Edit { module: d.module, span, text: new.to_string() });
            }
        }
    }
    Ok(edits)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{analyze, Config, SourceInput};
    use std::path::PathBuf;

    fn program(files: &[(&str, &str)]) -> Program {
        let sources = files
            .iter()
            .map(|(key, text)| SourceInput { path: PathBuf::from(format!("/p/src/{key}.ras")), rel_path: format!("src/{key}.ras"), key: key.to_string(), kind: ModuleKind::User, text: text.to_string() })
            .collect();
        analyze(sources, Config::default())
    }

    fn module(prog: &Program, key: &str) -> ModuleId {
        prog.module_by_key(key).unwrap()
    }

    /// Byte offset of the `n`-th (0-based) occurrence of `needle` in module `key`, plus `delta`.
    fn at(prog: &Program, key: &str, needle: &str, n: usize) -> u32 {
        let text = &prog.modules[module(prog, key)].text;
        text.match_indices(needle).nth(n).unwrap_or_else(|| panic!("{needle} #{n} in {key}")).0 as u32
    }

    /// The text after applying `edits` to module `key`.
    fn apply(prog: &Program, key: &str, edits: &[Edit]) -> String {
        let m = module(prog, key);
        let mut text = prog.modules[m].text.to_string();
        let mut mine: Vec<&Edit> = edits.iter().filter(|e| e.module == m).collect();
        mine.sort_by_key(|e| std::cmp::Reverse(u32::from(e.span.start())));
        for e in mine {
            text.replace_range(usize::from(e.span.start())..usize::from(e.span.end()), &e.text);
        }
        text
    }

    fn errors(prog: &Program) -> Vec<String> {
        prog.diagnostics.iter().filter(|d| d.severity == crate::Severity::Error).map(|d| d.message.clone()).collect()
    }

    #[test]
    fn renames_a_function_across_modules() {
        let util = "/** 2 倍。 */\nexport function twice(x: int): int {\n    return x * 2;\n}\n\n@onError(twice)\nfunction guarded() {\n}\n";
        let main = "import { twice } from \"./util\";\nimport * as u from \"./util\";\n\nexport function run() {\n    const a = twice(1) + u.twice(2);\n}\n";
        let prog = program(&[("util", util), ("main", main)]);
        assert!(errors(&prog).is_empty(), "{:?}", errors(&prog));
        let off = at(&prog, "main", "twice(1)", 0);
        let refs = references(&prog, module(&prog, "main"), off);
        assert_eq!(refs.len(), 5, "{refs:?}");
        assert_eq!(refs.iter().filter(|r| r.decl).count(), 1);
        let edits = rename(&prog, module(&prog, "main"), off, "doubled").unwrap();
        assert_eq!(apply(&prog, "util", &edits), util.replace("twice", "doubled"));
        assert_eq!(apply(&prog, "main", &edits), main.replace("twice", "doubled"));
        let renamed = program(&[("util", &apply(&prog, "util", &edits)), ("main", &apply(&prog, "main", &edits))]);
        assert!(errors(&renamed).is_empty(), "{:?}", errors(&renamed));
    }

    #[test]
    fn renames_a_parameter_with_named_arguments_and_docs() {
        let src = "/**\n * ラベル。\n * @param count 個数\n * @param {string} unit 単位\n */\nexport function label(count: int, unit: string = \"個\"): string {\n    return count + unit;\n}\n\nexport function run() {\n    const s = label(unit: \"本\", count: 3);\n    const t = label(2);\n}\n";
        let prog = program(&[("main", src)]);
        assert!(errors(&prog).is_empty(), "{:?}", errors(&prog));
        let m = module(&prog, "main");
        let edits = rename(&prog, m, at(&prog, "main", "count + unit", 0), "n").unwrap();
        let out = apply(&prog, "main", &edits);
        assert!(out.contains("@param n 個数") && out.contains("label(n: int") && out.contains("return n + unit") && !out.contains("count: 3") && out.contains("n: 3"), "{out}");
        let edits = rename(&prog, m, at(&prog, "main", "unit: \"本\"", 0), "suffix").unwrap();
        let out = apply(&prog, "main", &edits);
        assert!(out.contains("@param {string} suffix 単位") && out.contains("suffix: string") && out.contains("label(suffix: \"本\""), "{out}");
    }

    #[test]
    fn locals_are_scoped_to_their_function() {
        let src = "export function a() {\n    let total = 0;\n    total += 1;\n    const xs = [1].map(v => v + total);\n}\n\nexport function b() {\n    let total = 2;\n    msgBox(total);\n}\n";
        let prog = program(&[("main", src)]);
        let m = module(&prog, "main");
        let refs = references(&prog, m, at(&prog, "main", "total", 0));
        assert_eq!(refs.len(), 3, "{refs:?}");
        assert!(refs[0].decl && refs[1].write && !refs[2].write, "{refs:?}");
        let edits = rename(&prog, m, at(&prog, "main", "total", 1), "sum").unwrap();
        assert_eq!(apply(&prog, "main", &edits), src.replacen("total", "sum", 3));
        // the lambda parameter
        let edits = rename(&prog, m, at(&prog, "main", "v =>", 0), "item").unwrap();
        assert!(apply(&prog, "main", &edits).contains("map(item => item + total)"));
    }

    #[test]
    fn rejects_bad_names() {
        let src = "let counter = 0;\n\nexport function run() {\n    let other = 1;\n    counter = other + range(\"A1\").row;\n}\n\nfunction helper(): int {\n    let value2 = 3;\n    return counter;\n}\n";
        let prog = program(&[("main", src)]);
        let m = module(&prog, "main");
        let off = at(&prog, "main", "counter", 0);
        let err = |new: &str| rename(&prog, m, off, new).unwrap_err();
        assert!(err("1x").contains("識別子"));
        assert!(err("function").contains("キーワード"));
        assert!(err("Dim").contains("予約語"));
        assert!(err("Other").contains("大文字小文字"), "{}", err("Other"));
        assert!(err("helper").contains("既に宣言"), "{}", err("helper"));
        assert!(err("range").contains("別のもの"), "{}", err("range"));
        assert!(err("value2").contains("ローカル変数"), "{}", err("value2"));
        assert!(rename(&prog, m, off, "count2").is_ok());
        // a local cannot take a name its function already uses
        let off = at(&prog, "main", "other", 0);
        assert!(rename(&prog, m, off, "counter").unwrap_err().contains("別のもの"));
        // host names and `this` cannot be renamed
        assert!(prepare_rename(&prog, m, at(&prog, "main", "range", 0)).is_err());
        assert!(prepare_rename(&prog, m, at(&prog, "main", "row", 0)).is_err());
    }

    #[test]
    fn renames_interface_members_with_their_implementations() {
        let src = "interface Shape {\n    area(): double;\n    name: string;\n}\n\nclass Square implements Shape {\n    name: string = \"sq\";\n    side: double = 1;\n    area(): double {\n        return this.side * this.side;\n    }\n}\n\nclass Circle implements Shape {\n    name: string = \"c\";\n    area(): double {\n        return 3.14;\n    }\n}\n\nexport function run() {\n    const s: Shape = new Square();\n    const sq = new Square();\n    msgBox(s.area() + sq.area());\n    sq.name = \"x\";\n}\n";
        let prog = program(&[("main", src)]);
        assert!(errors(&prog).is_empty(), "{:?}", errors(&prog));
        let m = module(&prog, "main");
        let edits = rename(&prog, m, at(&prog, "main", "area", 0), "size").unwrap();
        let out = apply(&prog, "main", &edits);
        assert_eq!(out.matches("size()").count(), 5, "{out}");
        assert!(!out.contains("area"), "{out}");
        let edits = rename(&prog, m, at(&prog, "main", "sq.name", 0) + 3, "title").unwrap();
        let out = apply(&prog, "main", &edits);
        assert!(!out.contains("name"), "{out}");
        // a member that exists already
        assert!(rename(&prog, m, at(&prog, "main", "side", 0), "area").unwrap_err().contains("既に"));
        // the class name: declaration, `implements`, type annotations and `new`
        let refs = references(&prog, m, at(&prog, "main", "Shape", 0));
        assert_eq!(refs.len(), 4, "{refs:?}");
        let edits = rename(&prog, m, at(&prog, "main", "Square", 2), "Box").unwrap();
        assert_eq!(apply(&prog, "main", &edits), src.replace("Square", "Box"));
    }

    #[test]
    fn struct_fields_enums_and_shorthand() {
        let src = "struct Point {\n    x: double;\n    y: double;\n}\n\nenum Color {\n    Red,\n    Blue,\n}\n\nexport function run() {\n    const x = 1.0;\n    const p: Point = { x, y: 2 };\n    const c = Color.Red;\n    msgBox(p.x + p.y);\n}\n";
        let prog = program(&[("main", src)]);
        assert!(errors(&prog).is_empty(), "{:?}", errors(&prog));
        let m = module(&prog, "main");
        let edits = rename(&prog, m, at(&prog, "main", "p.x", 0) + 2, "left").unwrap();
        let out = apply(&prog, "main", &edits);
        assert!(out.contains("left: double") && out.contains("{ left: x, y: 2 }") && out.contains("p.left + p.y"), "{out}");
        let edits = rename(&prog, m, at(&prog, "main", "const x", 0) + 6, "x0").unwrap();
        assert!(apply(&prog, "main", &edits).contains("{ x: x0, y: 2 }"));
        let edits = rename(&prog, m, at(&prog, "main", "Red", 1), "Crimson").unwrap();
        assert_eq!(apply(&prog, "main", &edits), src.replace("Red", "Crimson"));
        let refs = references(&prog, m, at(&prog, "main", "Point", 0));
        assert_eq!(refs.len(), 2, "{refs:?}");
    }

    #[test]
    fn an_import_alias_is_not_renamed_from_its_uses() {
        let util = "export function twice(x: int): int {\n    return x * 2;\n}\n";
        let main = "import { twice as dbl } from \"./util\";\n\nexport function run() {\n    msgBox(dbl(1));\n}\n";
        let prog = program(&[("util", util), ("main", main)]);
        let m = module(&prog, "main");
        assert!(prepare_rename(&prog, m, at(&prog, "main", "dbl(1)", 0)).unwrap_err().contains("別名"));
        let edits = rename(&prog, module(&prog, "util"), at(&prog, "util", "twice", 0), "doubled").unwrap();
        assert_eq!(apply(&prog, "main", &edits), main.replace("twice", "doubled"));
        assert_eq!(references(&prog, m, at(&prog, "main", "dbl(1)", 0)).len(), 3);
    }

    #[test]
    fn imports_are_added_next_to_the_existing_ones() {
        let strs = "export function isBlank(this: string): boolean {\n    return this.length == 0;\n}\n\nexport function pad(s: string): string {\n    return s;\n}\n";
        let main = "import { pad } from \"../lib/str\";\n\nexport function run() {\n    msgBox(trim2(\"a\") + \"b\".isBlank());\n}\n";
        let prog = program(&[("lib/str", strs), ("lib/more", "export function trim2(s: string): string {\n    return s;\n}\n"), ("app/main", main)]);
        let m = module(&prog, "app/main");
        let (lib, more) = (module(&prog, "lib/str"), module(&prog, "lib/more"));
        assert_eq!(crate::query::import_source(&prog, m, lib), "../lib/str");
        assert_eq!(crate::query::import_source(&prog, lib, more), "./more");
        // added to the import of the same module, or as a new line after the imports
        let (span, text) = crate::query::import_edit(&prog, m, lib, "isBlank").unwrap();
        assert_eq!((u32::from(span.start()), text.as_str()), (12, ", isBlank"));
        let (span, text) = crate::query::import_edit(&prog, m, more, "trim2").unwrap();
        assert_eq!((u32::from(span.start()), text.as_str()), (34, "import { trim2 } from \"../lib/more\";\n"));
        let fixes = crate::query::import_fixes(&prog, m, 0, main.len() as u32);
        let titles: Vec<&str> = fixes.iter().map(|f| f.title.as_str()).collect();
        assert_eq!(titles, ["\"../lib/more\" から `trim2` をインポート", "\"../lib/str\" から `isBlank` をインポート"]);
    }

    #[test]
    fn form_code_behind_members() {
        let mut config = Config::default();
        config.forms.push(crate::FormDecl { name: "frmInput".into(), controls: vec![crate::FormControl { name: "btnOK".into(), control_type: "CommandButton".into() }] });
        let form = "export function entered(): string {\n    return this.btnOK.caption;\n}\n";
        let main = "export function run() {\n    msgBox(frmInput.entered());\n}\n";
        let sources = vec![
            SourceInput { path: PathBuf::from("/p/forms/frmInput.ras"), rel_path: "forms/frmInput.ras".into(), key: "<forms>/frmInput".into(), kind: ModuleKind::User, text: form.into() },
            SourceInput { path: PathBuf::from("/p/src/main.ras"), rel_path: "src/main.ras".into(), key: "main".into(), kind: ModuleKind::User, text: main.into() },
        ];
        let prog = analyze(sources, config);
        assert!(errors(&prog).is_empty(), "{:?}", errors(&prog));
        let f = module(&prog, "<forms>/frmInput");
        let off = at(&prog, "<forms>/frmInput", "entered", 0);
        let edits = rename(&prog, f, off, "typedName").unwrap();
        assert_eq!(apply(&prog, "main", &edits), main.replace("entered", "typedName"));
        assert!(rename(&prog, f, off, "btnOK").unwrap_err().contains("コントロール"));
        assert!(rename(&prog, f, off, "show").unwrap_err().contains("コントロール"));
        assert!(prepare_rename(&prog, f, at(&prog, "<forms>/frmInput", "btnOK", 0)).unwrap_err().contains("form.json"));
    }
}
