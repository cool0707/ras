//! Lint W0313 (spec §6 rule 16): a value that may be a cell error value (`#N/A` ...) is
//! compared, computed with, converted or passed to a function without being checked first.
//! VBA raises "type mismatch" (13) for all of these, only when the cell actually holds an error.
//!
//! Where a value may be an error: the members whose doc comment has `@errorValue`
//! (`Range.value`, `Application.vLookup`, `CVErr` ..., and the user's own functions), the cell
//! values read in bulk (`toArray()` ...), and the elements of all of these. Locals (and the
//! module variables of `Main`) that hold such a value are followed within the function body:
//! `isError(v)`, `typeof v == "..."`, `v == null` and `Excel.errorToNull(v)` clear them, as
//! does assigning another value. Values are not followed into the functions they are passed to.

use std::collections::HashSet;

use ras_syntax::ast::{
    self, ArrowBody, BinaryOp, Expr, ExprKind, Stmt, StmtKind, TemplatePart, UnaryOp,
};

use crate::check::{assigns_in_expr, assigns_in_stmt, named_value, stmt_exits};
use crate::diag::{Diagnostic, Severity};
use crate::intrinsics::Intr;
use crate::model::*;

/// The doc comment tag of members that may return an error value.
const TAG: &str = "@errorValue";

/// Host functions that take an error value without a type mismatch (they inspect it, print it,
/// or return an error value for it).
const ACCEPTING: &[&str] = &[
    "IsError",
    "IsEmpty",
    "IsNull",
    "IsObject",
    "IsArray",
    "IsNumeric",
    "IsDate",
    "IsMissing",
    "TypeName",
    "VarType",
    "IfError",
    "IfNa",
    "IsNA",
    "IsErr",
    "IsText",
    "IsNonText",
    "IsNumber",
    "IsLogical",
    "Print",
];

/// `Application` methods that return an error value instead of raising (`Application.Match`).
const APPLICATION_ACCEPTING: &[&str] =
    &["Match", "Index", "VLookup", "HLookup", "Lookup", "XLookup"];

/// Tests that are true only for error values / only for values that are not error values.
const IS_ERROR: &[&str] = &["IsError"];
const NOT_ERROR_WHEN_TRUE: &[&str] = &[
    "IsNumeric",
    "IsDate",
    "IsEmpty",
    "IsNull",
    "IsObject",
    "IsArray",
];

#[derive(Clone, Copy, PartialEq, Eq, Hash, Debug)]
enum Key {
    Local(LocalId),
    Global(VarId),
}

#[derive(Clone, Copy)]
enum Use {
    Compare,
    Compute,
    Condition,
    Argument,
    Convert,
    Index,
    Switch,
}

impl Use {
    fn what(self) -> &'static str {
        match self {
            Use::Compare => "比較する",
            Use::Compute => "演算に使う",
            Use::Condition => "条件に使う",
            Use::Argument => "関数の引数に渡す",
            Use::Convert => "型を変換する",
            Use::Index => "添字に使う",
            Use::Switch => "`switch` で比べる",
        }
    }
}

pub(crate) fn lint(prog: &mut Program) {
    let mut out = Vec::new();
    for f in 0..prog.fns.len() {
        let d = &prog.fns[f];
        let m = &prog.modules[d.module];
        if m.kind != ModuleKind::User || m.is_test() {
            continue;
        }
        let Some(decl) = d.decl.clone() else {
            continue;
        };
        let Some(body) = &decl.body else {
            continue;
        };
        let mut l = Lint {
            prog,
            module: d.module,
            f,
            tainted: HashSet::new(),
            clean: Vec::new(),
            rets: vec![d.ret.clone()],
            out: Vec::new(),
        };
        l.stmts(&body.stmts);
        out.extend(l.out);
    }
    prog.diagnostics.extend(out);
}

struct Lint<'a> {
    prog: &'a Program,
    module: ModuleId,
    f: FnId,
    /// Variables that may hold an error value here.
    tainted: HashSet<Key>,
    /// Variables known not to hold an error value here (`if (!isError(v)) { ... }`), innermost last.
    clean: Vec<Key>,
    /// Return types of the function and the arrow functions being walked (innermost last).
    rets: Vec<Option<Type>>,
    out: Vec<Diagnostic>,
}

/// A type whose values can be an error value, or contain one (`any`, `any[][]`).
fn holds_error(t: &Type) -> bool {
    match t {
        Type::Any => true,
        Type::Array(e) => holds_error(e),
        _ => false,
    }
}

fn tagged(doc: Option<&str>) -> bool {
    doc.is_some_and(|d| {
        d.match_indices(TAG).any(|(i, _)| {
            !d[i + TAG.len()..]
                .chars()
                .next()
                .is_some_and(|c| c.is_alphanumeric() || c == '_')
        })
    })
}

fn named(name: &str, list: &[&str]) -> bool {
    list.iter().any(|n| n.eq_ignore_ascii_case(name))
}

impl Lint<'_> {
    fn info(&self) -> &TypeInfo {
        &self.prog.modules[self.module].info
    }

    fn ty(&self, e: &Expr) -> Type {
        self.info()
            .expr_types
            .get(e.id as usize)
            .cloned()
            .unwrap_or(Type::Error)
    }

    fn key_of(&self, e: &Expr) -> Option<Key> {
        let e = e.peel();
        if !matches!(e.kind, ExprKind::Name(_)) {
            return None;
        }
        match self.info().res.get(&e.id)? {
            Res::Local(l) => Some(Key::Local(*l)),
            Res::Global(v) => Some(Key::Global(*v)),
            _ => None,
        }
    }

    /// The variable declared by the name at `name` (`let x`, `for (const x of ...)`).
    fn decl_key(&self, name: &ast::Ident) -> Option<Key> {
        let offset = u32::from(name.span.start());
        let info = self.info();
        info.decl_locals
            .get(&offset)
            .map(|l| Key::Local(*l))
            .or_else(|| info.decl_globals.get(&offset).map(|v| Key::Global(*v)))
    }

    fn key_ty(&self, k: Key) -> Type {
        match k {
            Key::Local(l) => self.prog.fns[self.f]
                .locals
                .get(l)
                .map(|d| d.ty.clone())
                .unwrap_or(Type::Error),
            Key::Global(v) => self.prog.vars[v].ty.clone(),
        }
    }

    fn key_name(&self, k: Key) -> &str {
        match k {
            Key::Local(l) => self.prog.fns[self.f]
                .locals
                .get(l)
                .map(|d| d.name.as_str())
                .unwrap_or(""),
            Key::Global(v) => &self.prog.vars[v].name,
        }
    }

    fn is_tainted(&self, k: Key) -> bool {
        self.tainted.contains(&k) && !self.clean.contains(&k)
    }

    /// The function a call expression calls, when it is a declared one.
    fn callee_fn(&self, call: &Expr) -> Option<FnId> {
        match &self.info().calls.get(&call.id)?.target {
            CallTarget::Fn(f)
            | CallTarget::Method(f)
            | CallTarget::Ext(f)
            | CallTarget::Host(f)
            | CallTarget::DefaultMember(f) => Some(*f),
            _ => None,
        }
    }

    fn callee_intr(&self, call: &Expr) -> Option<Intr> {
        match &self.info().calls.get(&call.id)?.target {
            CallTarget::Intrinsic(i) => Some(*i),
            _ => None,
        }
    }

    fn is_application_method(&self, f: FnId) -> bool {
        match self.prog.fns[f].kind {
            FnKind::Host { class: Some(class) } => self.prog.classes[class].name == "Application",
            _ => false,
        }
    }

    /// `e` may evaluate to an error value (or to an array holding one).
    fn may_err(&self, e: &Expr) -> bool {
        match &e.kind {
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => self.may_err(inner),
            ExprKind::Name(_) => self.key_of(e).is_some_and(|k| self.is_tainted(k)),
            ExprKind::Member { .. } => match self.info().res.get(&e.id) {
                Some(Res::Field(c, i)) => self.prog.classes[*c]
                    .fields
                    .get(*i)
                    .is_some_and(|fd| tagged(fd.doc.as_deref())),
                Some(Res::Property { get: Some(f), .. }) => {
                    tagged(self.prog.fns[*f].doc.as_deref())
                }
                _ => false,
            },
            ExprKind::Call { .. } => {
                if let Some(f) = self.callee_fn(e) {
                    tagged(self.prog.fns[f].doc.as_deref())
                } else {
                    matches!(
                        self.callee_intr(e),
                        Some(
                            Intr::RangeToArray
                                | Intr::SheetUsedTable
                                | Intr::RangeToColumn
                                | Intr::RangeToRow
                        )
                    )
                }
            }
            ExprKind::Index { object, .. } => self.may_err(object),
            ExprKind::Ternary { then, els, .. } => self.may_err(then) || self.may_err(els),
            ExprKind::Binary {
                op: BinaryOp::Coalesce,
                lhs,
                rhs,
            } => self.may_err(lhs) || self.may_err(rhs),
            ExprKind::Assign {
                op: None, value, ..
            } => self.may_err(value),
            ExprKind::As { expr, .. } => holds_error(&self.ty(e)) && self.may_err(expr),
            _ => false,
        }
    }

    /// How the value is shown in the message.
    fn label(&self, e: &Expr) -> String {
        match &e.kind {
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => self.label(inner),
            ExprKind::Name(id) => id.text.clone(),
            ExprKind::Member { object, name, .. } => {
                format!("{}.{}", self.label(object), name.text)
            }
            ExprKind::Call { callee, .. } => format!("{}(…)", self.label(callee)),
            ExprKind::Index { object, .. } => format!("{}[…]", self.label(object)),
            ExprKind::This => "this".into(),
            _ => "…".into(),
        }
    }

    fn use_(&mut self, e: &Expr, how: Use) {
        if !holds_error(&self.ty(e)) || !self.may_err(e) {
            return;
        }
        let message = format!(
            "`{}` はセルのエラー値（`#N/A` など）のことがあり、そのまま{}と実行時エラー（型が一致しません）になります。先に `isError()` で確かめるか、`Excel.errorToNull()` で `null` にしてください",
            self.label(e),
            how.what()
        );
        self.out.push(Diagnostic {
            severity: Severity::Warning,
            code: "W0313",
            message,
            module: Some(self.module),
            span: e.span,
        });
    }

    /// Records the value assigned to a variable that can hold an error value.
    fn set(&mut self, k: Key, value: &Expr) {
        if self.may_err(value) {
            self.tainted.insert(k);
        } else {
            self.tainted.remove(&k);
        }
    }

    // ------------------------------------------------------------------ guards

    /// Variables that hold no error value when `cond` is true / false.
    fn guard(&self, cond: &Expr) -> (Vec<Key>, Vec<Key>) {
        let none = || (Vec::new(), Vec::new());
        match &cond.peel().kind {
            ExprKind::Unary {
                op: UnaryOp::Not,
                operand,
            } => {
                let (t, f) = self.guard(operand);
                (f, t)
            }
            ExprKind::Binary {
                op: BinaryOp::And,
                lhs,
                rhs,
            } => {
                let (mut t, _) = self.guard(lhs);
                t.extend(self.guard(rhs).0);
                (t, Vec::new())
            }
            ExprKind::Binary {
                op: BinaryOp::Or,
                lhs,
                rhs,
            } => {
                let (_, mut f) = self.guard(lhs);
                f.extend(self.guard(rhs).1);
                (Vec::new(), f)
            }
            ExprKind::Call { args, .. } => {
                let Some(k) = args.first().and_then(|a| self.key_of(named_value(a))) else {
                    return none();
                };
                let name = match self.callee_fn(cond.peel()) {
                    Some(f) => self.prog.fns[f].vba_name().to_string(),
                    None if self.callee_intr(cond.peel()) == Some(Intr::IsNullish) => {
                        return (vec![k], Vec::new());
                    }
                    None => return none(),
                };
                if named(&name, IS_ERROR) {
                    (Vec::new(), vec![k])
                } else if named(&name, NOT_ERROR_WHEN_TRUE) {
                    (vec![k], Vec::new())
                } else {
                    none()
                }
            }
            ExprKind::Binary {
                op: op @ (BinaryOp::Eq | BinaryOp::Ne),
                lhs,
                rhs,
            } => {
                let test = |a: &Expr, b: &Expr| -> Option<(Vec<Key>, Vec<Key>)> {
                    match (&a.peel().kind, &b.peel().kind) {
                        // `typeof v == "Error"`
                        (
                            ExprKind::Unary {
                                op: UnaryOp::TypeOf,
                                operand,
                            },
                            ExprKind::Str(name),
                        ) => {
                            let k = self.key_of(operand)?;
                            Some(if name.eq_ignore_ascii_case("Error") {
                                (Vec::new(), vec![k])
                            } else {
                                (vec![k], Vec::new())
                            })
                        }
                        // `v == null`
                        (_, ExprKind::Null | ExprKind::Undefined) => {
                            Some((vec![self.key_of(a)?], Vec::new()))
                        }
                        _ => None,
                    }
                };
                let Some((t, f)) = test(lhs, rhs).or_else(|| test(rhs, lhs)) else {
                    return none();
                };
                if *op == BinaryOp::Eq {
                    (t, f)
                } else {
                    (f, t)
                }
            }
            _ => none(),
        }
    }

    /// Runs `f` with `keys` known to hold no error value (those that `assigned` says the region
    /// does not assign).
    fn with_clean(
        &mut self,
        keys: Vec<Key>,
        assigned: &dyn Fn(&str) -> bool,
        f: impl FnOnce(&mut Self),
    ) {
        let mark = self.clean.len();
        let keys: Vec<Key> = keys
            .into_iter()
            .filter(|k| !assigned(self.key_name(*k)))
            .collect();
        self.clean.extend(keys);
        f(self);
        self.clean.truncate(mark);
    }

    // ------------------------------------------------------------------ statements

    fn stmts(&mut self, stmts: &[Stmt]) {
        let mark = self.clean.len();
        for (i, s) in stmts.iter().enumerate() {
            self.stmt(s);
            // `if (isError(v)) return;` — the rest of the list sees `v` without an error value.
            if let StmtKind::If { cond, then, els } = &s.kind {
                let then_exits = stmt_exits(then);
                let else_exits = els.as_ref().is_some_and(|e| stmt_exits(e));
                if then_exits != else_exits {
                    let (t, f) = self.guard(cond);
                    let keys = if then_exits { f } else { t };
                    let rest = &stmts[i + 1..];
                    for k in keys {
                        let name = self.key_name(k);
                        if !rest.iter().any(|r| assigns_in_stmt(r, name)) {
                            self.clean.push(k);
                        }
                    }
                }
            }
        }
        self.clean.truncate(mark);
    }

    fn cond(&mut self, e: &Expr) {
        self.expr(e);
        self.use_(e, Use::Condition);
    }

    fn stmt(&mut self, s: &Stmt) {
        match &s.kind {
            StmtKind::Var(v) => self.var(v),
            StmtKind::Expr(e) | StmtKind::Throw(e) => self.expr(e),
            StmtKind::Block(b) => self.stmts(&b.stmts),
            StmtKind::If { cond, then, els } => {
                self.cond(cond);
                let (t, f) = self.guard(cond);
                let before = self.tainted.clone();
                self.with_clean(t, &|n| assigns_in_stmt(then, n), |l| l.stmt(then));
                let after_then = std::mem::replace(&mut self.tainted, before);
                if let Some(e) = els {
                    self.with_clean(f, &|n| assigns_in_stmt(e, n), |l| l.stmt(e));
                }
                self.tainted.extend(after_then);
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
                ..
            } => {
                for st in init {
                    self.stmt(st);
                }
                let before = self.tainted.clone();
                let keys = match cond {
                    Some(c) => {
                        self.cond(c);
                        self.guard(c).0
                    }
                    None => Vec::new(),
                };
                let assigned = |n: &str| {
                    assigns_in_stmt(body, n) || update.iter().any(|u| assigns_in_expr(u, n))
                };
                self.with_clean(keys, &assigned, |l| {
                    l.stmt(body);
                    for u in update {
                        l.expr(u);
                    }
                });
                self.tainted.extend(before);
            }
            StmtKind::ForOf {
                name, iter, body, ..
            } => {
                self.expr(iter);
                if let Some(k) = self.decl_key(name) {
                    if self.may_err(iter) && holds_error(&self.key_ty(k)) {
                        self.tainted.insert(k);
                    } else {
                        self.tainted.remove(&k);
                    }
                }
                let before = self.tainted.clone();
                self.stmt(body);
                self.tainted.extend(before);
            }
            StmtKind::While { cond, body, .. } => {
                self.cond(cond);
                let (t, _) = self.guard(cond);
                let before = self.tainted.clone();
                self.with_clean(t, &|n| assigns_in_stmt(body, n), |l| l.stmt(body));
                self.tainted.extend(before);
            }
            StmtKind::DoWhile { body, cond, .. } => {
                self.stmt(body);
                self.cond(cond);
            }
            StmtKind::Switch { disc, cases } => {
                self.expr(disc);
                self.use_(disc, Use::Switch);
                let before = self.tainted.clone();
                let mut after = before.clone();
                for c in cases {
                    if let Some(t) = &c.test {
                        self.expr(t);
                        self.use_(t, Use::Switch);
                    }
                    self.tainted = before.clone();
                    self.stmts(&c.body);
                    after.extend(self.tainted.drain());
                }
                self.tainted = after;
            }
            StmtKind::Try(t) => {
                let before = self.tainted.clone();
                self.stmts(&t.block.stmts);
                if let Some(c) = &t.catch {
                    // The `try` block may have stopped anywhere.
                    self.tainted.extend(before);
                    self.stmts(&c.body.stmts);
                }
                if let Some(f) = &t.finally {
                    self.stmts(&f.stmts);
                }
            }
            StmtKind::Return(Some(e)) => {
                self.expr(e);
                if let Some(Some(t)) = self.rets.last() {
                    if !holds_error(t) && !matches!(t, Type::Void | Type::Error) {
                        self.use_(e, Use::Convert);
                    }
                }
            }
            StmtKind::Labeled { body, .. } => self.stmt(body),
            _ => {}
        }
    }

    fn var(&mut self, v: &ast::VarDecl) {
        let key = self.decl_key(&v.name);
        let Some(init) = &v.init else {
            if let Some(k) = key {
                self.tainted.remove(&k);
            }
            return;
        };
        self.expr(init);
        match key {
            Some(k) if holds_error(&self.key_ty(k)) => self.set(k, init),
            Some(k) if !matches!(self.key_ty(k), Type::Error) => self.use_(init, Use::Convert),
            _ => {}
        }
    }

    // ------------------------------------------------------------------ expressions

    fn expr(&mut self, e: &Expr) {
        match &e.kind {
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => self.expr(inner),
            ExprKind::Template(parts) => {
                for p in parts {
                    if let TemplatePart::Expr(x) = p {
                        self.expr(x);
                    }
                }
            }
            ExprKind::Array(items) => {
                for i in items {
                    self.expr(i);
                }
            }
            ExprKind::Object(props) => {
                for p in props {
                    self.expr(&p.value);
                }
            }
            ExprKind::Arrow(a) => self.arrow(e, a),
            ExprKind::Unary { op, operand } => {
                self.expr(operand);
                match op {
                    UnaryOp::TypeOf => {}
                    UnaryOp::Not => self.use_(operand, Use::Condition),
                    _ => self.use_(operand, Use::Compute),
                }
            }
            ExprKind::IncDec { operand, .. } => {
                self.expr(operand);
                self.use_(operand, Use::Compute);
            }
            ExprKind::Binary { op, lhs, rhs } => self.binary(*op, lhs, rhs),
            ExprKind::Assign { op, target, value } => self.assign(*op, target, value),
            ExprKind::Ternary { cond, then, els } => {
                self.cond(cond);
                let (t, f) = self.guard(cond);
                self.with_clean(t, &|n| assigns_in_expr(then, n), |l| l.expr(then));
                self.with_clean(f, &|n| assigns_in_expr(els, n), |l| l.expr(els));
            }
            ExprKind::Call { callee, args, .. } => self.call(e, callee, args),
            ExprKind::New { args, .. } => {
                for a in args {
                    let v = named_value(a);
                    self.expr(v);
                    self.use_(v, Use::Argument);
                }
            }
            ExprKind::Member { object, .. } => self.expr(object),
            ExprKind::Index { object, index } => {
                self.expr(object);
                self.expr(index);
                self.use_(index, Use::Index);
            }
            ExprKind::As { expr, .. } => {
                self.expr(expr);
                if !holds_error(&self.ty(e)) {
                    self.use_(expr, Use::Convert);
                }
            }
            ExprKind::InstanceOf { expr, .. } => self.expr(expr),
            ExprKind::NamedArg { value, .. } => self.expr(value),
            _ => {}
        }
    }

    fn arrow(&mut self, e: &Expr, a: &ast::ArrowFn) {
        let ret = match self.ty(e) {
            Type::Func(ft) => Some(ft.ret.clone()),
            _ => None,
        };
        self.rets.push(ret);
        match &a.body {
            ArrowBody::Expr(x) => {
                self.expr(x);
                if let Some(Some(t)) = self.rets.last() {
                    if !holds_error(t) && !matches!(t, Type::Void | Type::Error) {
                        self.use_(x, Use::Convert);
                    }
                }
            }
            ArrowBody::Block(b) => self.stmts(&b.stmts),
        }
        self.rets.pop();
    }

    fn binary(&mut self, op: BinaryOp, lhs: &Expr, rhs: &Expr) {
        match op {
            BinaryOp::And | BinaryOp::Or => {
                self.cond(lhs);
                let (t, f) = self.guard(lhs);
                let keys = if op == BinaryOp::And { t } else { f };
                self.with_clean(keys, &|n| assigns_in_expr(rhs, n), |l| l.cond(rhs));
            }
            BinaryOp::Coalesce => {
                self.expr(lhs);
                self.expr(rhs);
            }
            BinaryOp::Eq | BinaryOp::Ne => {
                self.expr(lhs);
                self.expr(rhs);
                // `v == null` is `IsEmpty` / `IsNull`, which take an error value.
                let null = |x: &Expr| {
                    matches!(x.peel().kind, ExprKind::Null | ExprKind::Undefined)
                        || self.ty(x) == Type::Null
                };
                if !null(lhs) && !null(rhs) {
                    self.use_(lhs, Use::Compare);
                    self.use_(rhs, Use::Compare);
                }
            }
            _ => {
                self.expr(lhs);
                self.expr(rhs);
                let how = if op.is_comparison() {
                    Use::Compare
                } else {
                    Use::Compute
                };
                self.use_(lhs, how);
                self.use_(rhs, how);
            }
        }
    }

    fn assign(&mut self, op: Option<BinaryOp>, target: &Expr, value: &Expr) {
        self.expr(value);
        match &target.peel().kind {
            ExprKind::Index { object, index } => {
                self.expr(object);
                self.expr(index);
                self.use_(index, Use::Index);
            }
            ExprKind::Member { object, .. } => self.expr(object),
            _ => {}
        }
        match op {
            Some(BinaryOp::Coalesce | BinaryOp::And | BinaryOp::Or) => {}
            Some(_) => {
                self.use_(target, Use::Compute);
                self.use_(value, Use::Compute);
            }
            None => match self.key_of(target) {
                Some(k) if holds_error(&self.key_ty(k)) => self.set(k, value),
                Some(_) => self.use_(value, Use::Convert),
                None => {
                    let t = self.ty(target);
                    if !holds_error(&t) && !matches!(t, Type::Error) {
                        self.use_(value, Use::Convert);
                    }
                }
            },
        }
    }

    fn call(&mut self, e: &Expr, callee: &Expr, args: &[Expr]) {
        self.expr(callee);
        let accepts = match (self.callee_fn(e), self.callee_intr(e)) {
            (Some(f), _) => {
                let name = self.prog.fns[f].vba_name();
                named(name, ACCEPTING)
                    || (named(name, APPLICATION_ACCEPTING) && self.is_application_method(f))
            }
            (None, Some(i)) => matches!(
                i,
                Intr::ExcelErrorToNull | Intr::ExcelMatch | Intr::IsNullish | Intr::ConsoleLog
            ),
            // A closure or a late-bound call: nothing is known about the parameters.
            (None, None) => false,
        };
        for a in args {
            let v = named_value(a);
            self.expr(v);
            if !accepts {
                self.use_(v, Use::Argument);
            }
        }
    }
}
