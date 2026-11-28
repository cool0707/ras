//! Generic AST walkers used by the analyses in the lowering pass.

use ras_syntax::ast::*;

/// Pre-order walk. `fs` / `fe` return `false` to skip the children of a node.
pub fn walk_stmt<'a>(
    s: &'a Stmt,
    fs: &mut dyn FnMut(&'a Stmt) -> bool,
    fe: &mut dyn FnMut(&'a Expr) -> bool,
) {
    if !fs(s) {
        return;
    }
    match &s.kind {
        StmtKind::Var(v) => {
            if let Some(i) = &v.init {
                walk_expr(i, fs, fe);
            }
        }
        StmtKind::Expr(e) | StmtKind::Throw(e) => walk_expr(e, fs, fe),
        StmtKind::Block(b) => walk_stmts(&b.stmts, fs, fe),
        StmtKind::If { cond, then, els } => {
            walk_expr(cond, fs, fe);
            walk_stmt(then, fs, fe);
            if let Some(e) = els {
                walk_stmt(e, fs, fe);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            walk_stmts(init, fs, fe);
            if let Some(c) = cond {
                walk_expr(c, fs, fe);
            }
            for u in update {
                walk_expr(u, fs, fe);
            }
            walk_stmt(body, fs, fe);
        }
        StmtKind::ForOf { iter, body, .. } => {
            walk_expr(iter, fs, fe);
            walk_stmt(body, fs, fe);
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            walk_expr(cond, fs, fe);
            walk_stmt(body, fs, fe);
        }
        StmtKind::Switch { disc, cases } => {
            walk_expr(disc, fs, fe);
            for c in cases {
                if let Some(t) = &c.test {
                    walk_expr(t, fs, fe);
                }
                walk_stmts(&c.body, fs, fe);
            }
        }
        StmtKind::Try(t) => {
            walk_stmts(&t.block.stmts, fs, fe);
            if let Some(c) = &t.catch {
                walk_stmts(&c.body.stmts, fs, fe);
            }
            if let Some(f) = &t.finally {
                walk_stmts(&f.stmts, fs, fe);
            }
        }
        StmtKind::Return(Some(e)) => walk_expr(e, fs, fe),
        StmtKind::Labeled { body, .. } => walk_stmt(body, fs, fe),
        StmtKind::Return(None)
        | StmtKind::Break(_)
        | StmtKind::Continue(_)
        | StmtKind::Resume
        | StmtKind::Debugger
        | StmtKind::Fn(_)
        | StmtKind::Empty
        | StmtKind::Error => {}
    }
}

pub fn walk_stmts<'a>(
    stmts: &'a [Stmt],
    fs: &mut dyn FnMut(&'a Stmt) -> bool,
    fe: &mut dyn FnMut(&'a Expr) -> bool,
) {
    for s in stmts {
        walk_stmt(s, fs, fe);
    }
}

pub fn walk_expr<'a>(
    e: &'a Expr,
    fs: &mut dyn FnMut(&'a Stmt) -> bool,
    fe: &mut dyn FnMut(&'a Expr) -> bool,
) {
    if !fe(e) {
        return;
    }
    match &e.kind {
        ExprKind::Template(parts) => {
            for p in parts {
                if let TemplatePart::Expr(x) = p {
                    walk_expr(x, fs, fe);
                }
            }
        }
        ExprKind::Paren(x) | ExprKind::NonNull(x) | ExprKind::NamedArg { value: x, .. } => {
            walk_expr(x, fs, fe)
        }
        ExprKind::Array(items) => {
            for i in items {
                walk_expr(i, fs, fe);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                walk_expr(&p.value, fs, fe);
            }
        }
        ExprKind::Arrow(a) => match &a.body {
            ArrowBody::Expr(x) => walk_expr(x, fs, fe),
            ArrowBody::Block(b) => walk_stmts(&b.stmts, fs, fe),
        },
        ExprKind::Unary { operand, .. } | ExprKind::IncDec { operand, .. } => {
            walk_expr(operand, fs, fe)
        }
        ExprKind::Binary { lhs, rhs, .. } => {
            walk_expr(lhs, fs, fe);
            walk_expr(rhs, fs, fe);
        }
        ExprKind::Assign { target, value, .. } => {
            walk_expr(target, fs, fe);
            walk_expr(value, fs, fe);
        }
        ExprKind::Ternary { cond, then, els } => {
            walk_expr(cond, fs, fe);
            walk_expr(then, fs, fe);
            walk_expr(els, fs, fe);
        }
        ExprKind::Call { callee, args, .. } => {
            walk_expr(callee, fs, fe);
            for a in args {
                walk_expr(a, fs, fe);
            }
        }
        ExprKind::New { args, .. } => {
            for a in args {
                walk_expr(a, fs, fe);
            }
        }
        ExprKind::Member { object, .. } => walk_expr(object, fs, fe),
        ExprKind::Index { object, index } => {
            walk_expr(object, fs, fe);
            walk_expr(index, fs, fe);
        }
        ExprKind::FromEnd(x) => walk_expr(x, fs, fe),
        ExprKind::Subscripts(xs) => {
            for x in xs {
                walk_expr(x, fs, fe);
            }
        }
        ExprKind::Range { start, end } => {
            for x in [start, end].into_iter().flatten() {
                walk_expr(x, fs, fe);
            }
        }
        ExprKind::As { expr, .. } | ExprKind::InstanceOf { expr, .. } => walk_expr(expr, fs, fe),
        ExprKind::Int { .. }
        | ExprKind::Float { .. }
        | ExprKind::Str(_)
        | ExprKind::Bool(_)
        | ExprKind::Null
        | ExprKind::Undefined
        | ExprKind::Name(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::Error => {}
    }
}

/// True when any expression in `s` satisfies `pred`.
pub fn any_expr<'a>(s: &'a Stmt, pred: &mut dyn FnMut(&'a Expr) -> bool) -> bool {
    let mut found = false;
    walk_stmt(s, &mut |_| true, &mut |e| {
        if !found && pred(e) {
            found = true;
        }
        !found
    });
    found
}

pub fn any_expr_in<'a>(e: &'a Expr, pred: &mut dyn FnMut(&'a Expr) -> bool) -> bool {
    let mut found = false;
    walk_expr(e, &mut |_| true, &mut |x| {
        if !found && pred(x) {
            found = true;
        }
        !found
    });
    found
}

/// True when any statement in `s` satisfies `pred` (arrow function bodies included).
pub fn any_stmt<'a>(s: &'a Stmt, pred: &mut dyn FnMut(&'a Stmt) -> bool) -> bool {
    let found = std::cell::Cell::new(false);
    walk_stmt(
        s,
        &mut |x| {
            if !found.get() && pred(x) {
                found.set(true);
            }
            !found.get()
        },
        &mut |_| !found.get(),
    );
    found.get()
}

/// Callbacks of [`visit_module`]: everything a module names, in source order.
pub trait Visit<'a> {
    /// Every expression, decorator arguments and parameter defaults included.
    fn expr(&mut self, _e: &'a Expr) {}
    fn stmt(&mut self, _s: &'a Stmt) {}
    fn decorator(&mut self, _d: &'a Decorator) {}
    /// A named type (`T`, `Excel.Range`, `Map<K, V>`) with the type parameters in scope.
    fn named_type(&mut self, _path: &'a [Ident], _span: Span, _type_params: &[&'a str]) {}
}

/// Walks a whole module: items, members, statements, expressions and type annotations.
pub fn visit_module<'a>(module: &'a Module, v: &mut dyn Visit<'a>) {
    let mut tps = Vec::new();
    for item in &module.items {
        visit_item(item, v, &mut tps);
    }
}

fn visit_item<'a>(item: &'a Item, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    match item {
        Item::Import(_) => {}
        Item::Fn(d) => visit_fn(d, v, tps),
        Item::Accessor(a) => visit_fn(&a.func, v, tps),
        Item::Class(c) => {
            for d in &c.decorators {
                visit_decorator(d, v, tps);
            }
            let n = tps.len();
            visit_type_params(&c.type_params, v, tps);
            for t in c.extends.iter().chain(&c.implements) {
                visit_type(t, v, tps);
            }
            for m in &c.members {
                match m {
                    ClassMember::Field(f) => {
                        if let Some(t) = &f.ty {
                            visit_type(t, v, tps);
                        }
                        if let Some(x) = &f.init {
                            visit_expr(x, v, tps);
                        }
                    }
                    ClassMember::Method(md) => visit_fn(&md.func, v, tps),
                    ClassMember::Ctor(d) => visit_fn(d, v, tps),
                    ClassMember::Accessor(a) => visit_fn(&a.func, v, tps),
                }
            }
            tps.truncate(n);
        }
        Item::Interface(i) => {
            let n = tps.len();
            visit_type_params(&i.type_params, v, tps);
            for t in &i.extends {
                visit_type(t, v, tps);
            }
            for m in &i.members {
                match m {
                    InterfaceMember::Prop(p) => {
                        if let Some(t) = &p.ty {
                            visit_type(t, v, tps);
                        }
                    }
                    InterfaceMember::Method(d) => visit_fn(d, v, tps),
                    InterfaceMember::Index { param, ty, .. } => {
                        visit_param(param, v, tps);
                        visit_type(ty, v, tps);
                    }
                }
            }
            tps.truncate(n);
        }
        Item::Enum(e) => {
            for m in &e.members {
                if let Some(x) = &m.value {
                    visit_expr(x, v, tps);
                }
            }
        }
        Item::TypeAlias(a) => {
            let n = tps.len();
            visit_type_params(&a.type_params, v, tps);
            visit_type(&a.ty, v, tps);
            tps.truncate(n);
        }
        Item::Var(d) => visit_var(d, v, tps),
        Item::DeclareModule(d) => {
            for f in &d.functions {
                visit_fn(f, v, tps);
            }
        }
        Item::Struct(s) => {
            for f in &s.fields {
                if let Some(t) = &f.ty {
                    visit_type(t, v, tps);
                }
            }
        }
        Item::Namespace(ns) => {
            for i in &ns.items {
                visit_item(i, v, tps);
            }
        }
        Item::Stmt(s) => visit_stmt(s, v, tps),
        Item::RawVba(_) => {}
    }
}

fn visit_type_params<'a>(params: &'a [TypeParam], v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    for p in params {
        tps.push(&p.name.text);
    }
    for p in params {
        if let Some(c) = &p.constraint {
            visit_type(c, v, tps);
        }
    }
}

fn visit_decorator<'a>(d: &'a Decorator, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    v.decorator(d);
    for a in &d.args {
        visit_expr(a, v, tps);
    }
}

fn visit_param<'a>(p: &'a Param, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    if let Some(t) = &p.ty {
        visit_type(t, v, tps);
    }
    if let Some(x) = &p.default {
        visit_expr(x, v, tps);
    }
}

fn visit_fn<'a>(d: &'a FnDecl, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    for x in &d.decorators {
        visit_decorator(x, v, tps);
    }
    let n = tps.len();
    visit_type_params(&d.type_params, v, tps);
    for p in &d.params {
        visit_param(p, v, tps);
    }
    if let Some(t) = &d.ret {
        visit_type(t, v, tps);
    }
    if let Some(b) = &d.body {
        for s in &b.stmts {
            visit_stmt(s, v, tps);
        }
    }
    tps.truncate(n);
}

fn visit_var<'a>(d: &'a VarDecl, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    if let Some(t) = &d.ty {
        visit_type(t, v, tps);
    }
    if let Some(x) = &d.init {
        visit_expr(x, v, tps);
    }
}

fn visit_stmts<'a>(stmts: &'a [Stmt], v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    for s in stmts {
        visit_stmt(s, v, tps);
    }
}

fn visit_stmt<'a>(s: &'a Stmt, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    v.stmt(s);
    match &s.kind {
        StmtKind::Var(d) => visit_var(d, v, tps),
        StmtKind::Expr(e) | StmtKind::Throw(e) | StmtKind::Return(Some(e)) => visit_expr(e, v, tps),
        StmtKind::Block(b) => visit_stmts(&b.stmts, v, tps),
        StmtKind::If { cond, then, els } => {
            visit_expr(cond, v, tps);
            visit_stmt(then, v, tps);
            if let Some(e) = els {
                visit_stmt(e, v, tps);
            }
        }
        StmtKind::For {
            decorators,
            init,
            cond,
            update,
            body,
        } => {
            for d in decorators {
                visit_decorator(d, v, tps);
            }
            visit_stmts(init, v, tps);
            if let Some(c) = cond {
                visit_expr(c, v, tps);
            }
            for u in update {
                visit_expr(u, v, tps);
            }
            visit_stmt(body, v, tps);
        }
        StmtKind::ForOf {
            decorators,
            ty,
            iter,
            body,
            ..
        } => {
            for d in decorators {
                visit_decorator(d, v, tps);
            }
            if let Some(t) = ty {
                visit_type(t, v, tps);
            }
            visit_expr(iter, v, tps);
            visit_stmt(body, v, tps);
        }
        StmtKind::While {
            decorators,
            cond,
            body,
        }
        | StmtKind::DoWhile {
            decorators,
            body,
            cond,
        } => {
            for d in decorators {
                visit_decorator(d, v, tps);
            }
            visit_expr(cond, v, tps);
            visit_stmt(body, v, tps);
        }
        StmtKind::Switch { disc, cases } => {
            visit_expr(disc, v, tps);
            for c in cases {
                if let Some(t) = &c.test {
                    visit_expr(t, v, tps);
                }
                visit_stmts(&c.body, v, tps);
            }
        }
        StmtKind::Try(t) => {
            visit_stmts(&t.block.stmts, v, tps);
            if let Some(c) = &t.catch {
                visit_stmts(&c.body.stmts, v, tps);
            }
            if let Some(f) = &t.finally {
                visit_stmts(&f.stmts, v, tps);
            }
        }
        StmtKind::Labeled { body, .. } => visit_stmt(body, v, tps),
        StmtKind::Fn(d) => visit_fn(d, v, tps),
        StmtKind::Return(None)
        | StmtKind::Break(_)
        | StmtKind::Continue(_)
        | StmtKind::Resume
        | StmtKind::Debugger
        | StmtKind::Empty
        | StmtKind::Error => {}
    }
}

fn visit_expr<'a>(e: &'a Expr, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    v.expr(e);
    match &e.kind {
        ExprKind::Template(parts) => {
            for p in parts {
                if let TemplatePart::Expr(x) = p {
                    visit_expr(x, v, tps);
                }
            }
        }
        ExprKind::Paren(x) | ExprKind::NonNull(x) | ExprKind::NamedArg { value: x, .. } => {
            visit_expr(x, v, tps)
        }
        ExprKind::Array(items) => {
            for i in items {
                visit_expr(i, v, tps);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                visit_expr(&p.value, v, tps);
            }
        }
        ExprKind::Arrow(a) => {
            for p in &a.params {
                visit_param(p, v, tps);
            }
            if let Some(t) = &a.ret {
                visit_type(t, v, tps);
            }
            match &a.body {
                ArrowBody::Expr(x) => visit_expr(x, v, tps),
                ArrowBody::Block(b) => visit_stmts(&b.stmts, v, tps),
            }
        }
        ExprKind::Unary { operand, .. } | ExprKind::IncDec { operand, .. } => {
            visit_expr(operand, v, tps)
        }
        ExprKind::Binary { lhs, rhs, .. } => {
            visit_expr(lhs, v, tps);
            visit_expr(rhs, v, tps);
        }
        ExprKind::Assign { target, value, .. } => {
            visit_expr(target, v, tps);
            visit_expr(value, v, tps);
        }
        ExprKind::Ternary { cond, then, els } => {
            visit_expr(cond, v, tps);
            visit_expr(then, v, tps);
            visit_expr(els, v, tps);
        }
        ExprKind::Call {
            callee,
            type_args,
            args,
        } => {
            visit_expr(callee, v, tps);
            for t in type_args {
                visit_type(t, v, tps);
            }
            for a in args {
                visit_expr(a, v, tps);
            }
        }
        ExprKind::New { ty, args } => {
            visit_type(ty, v, tps);
            for a in args {
                visit_expr(a, v, tps);
            }
        }
        ExprKind::Member { object, .. } => visit_expr(object, v, tps),
        ExprKind::Index { object, index } => {
            visit_expr(object, v, tps);
            visit_expr(index, v, tps);
        }
        ExprKind::FromEnd(x) => visit_expr(x, v, tps),
        ExprKind::Subscripts(xs) => {
            for x in xs {
                visit_expr(x, v, tps);
            }
        }
        ExprKind::Range { start, end } => {
            for x in [start, end].into_iter().flatten() {
                visit_expr(x, v, tps);
            }
        }
        ExprKind::As { expr, ty } | ExprKind::InstanceOf { expr, ty } => {
            visit_expr(expr, v, tps);
            visit_type(ty, v, tps);
        }
        ExprKind::Int { .. }
        | ExprKind::Float { .. }
        | ExprKind::Str(_)
        | ExprKind::Bool(_)
        | ExprKind::Null
        | ExprKind::Undefined
        | ExprKind::Name(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::Error => {}
    }
}

fn visit_type<'a>(t: &'a TypeRef, v: &mut dyn Visit<'a>, tps: &mut Vec<&'a str>) {
    match t {
        TypeRef::Named { path, args, span } => {
            v.named_type(path, *span, tps);
            for a in args {
                visit_type(a, v, tps);
            }
        }
        TypeRef::Array { elem, .. } => visit_type(elem, v, tps),
        TypeRef::Fn { params, ret, .. } => {
            for p in params {
                visit_param(p, v, tps);
            }
            visit_type(ret, v, tps);
        }
        TypeRef::Nullable { inner, .. } => visit_type(inner, v, tps),
        TypeRef::Tuple { elems, .. } | TypeRef::Union { elems, .. } => {
            for e in elems {
                visit_type(e, v, tps);
            }
        }
        TypeRef::Object { fields, .. } => {
            for f in fields {
                if let Some(t) = &f.ty {
                    visit_type(t, v, tps);
                }
            }
        }
        TypeRef::Error { .. } => {}
    }
}
