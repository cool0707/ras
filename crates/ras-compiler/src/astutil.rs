//! Generic AST walkers used by the analyses in the lowering pass.

use ras_syntax::ast::*;

/// Pre-order walk. `fs` / `fe` return `false` to skip the children of a node.
pub fn walk_stmt<'a>(s: &'a Stmt, fs: &mut dyn FnMut(&'a Stmt) -> bool, fe: &mut dyn FnMut(&'a Expr) -> bool) {
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
        StmtKind::For { init, cond, update, body, .. } => {
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
        | StmtKind::Fn(_)
        | StmtKind::Empty
        | StmtKind::Error => {}
    }
}

pub fn walk_stmts<'a>(stmts: &'a [Stmt], fs: &mut dyn FnMut(&'a Stmt) -> bool, fe: &mut dyn FnMut(&'a Expr) -> bool) {
    for s in stmts {
        walk_stmt(s, fs, fe);
    }
}

pub fn walk_expr<'a>(e: &'a Expr, fs: &mut dyn FnMut(&'a Stmt) -> bool, fe: &mut dyn FnMut(&'a Expr) -> bool) {
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
        ExprKind::Paren(x) | ExprKind::NonNull(x) | ExprKind::NamedArg { value: x, .. } => walk_expr(x, fs, fe),
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
        ExprKind::Unary { operand, .. } | ExprKind::IncDec { operand, .. } => walk_expr(operand, fs, fe),
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
