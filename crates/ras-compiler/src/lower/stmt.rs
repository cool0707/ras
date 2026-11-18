//! Statement lowering: control flow, loops, try / catch / finally, returns.

use std::collections::HashSet;

use ras_syntax::ast::{self, BinaryOp, Expr, ExprKind, Stmt, StmtKind, VarKind};

use super::{Binding, Breakable, LoopKind, Lowerer, PendingExit, RetMode, TryCtx, TryKind};
use crate::astutil::{any_expr, any_expr_in, any_stmt};
use crate::intrinsics::Intr;
use crate::model::*;
use crate::vba::*;

/// Does the body need the procedure-level error handler (`RAS_Handler`)?
pub(crate) fn needs_handler(stmts: &[Stmt]) -> bool {
    stmts.iter().any(|s| {
        any_stmt(s, &mut |x| match &x.kind {
            StmtKind::Try(t) => t.finally.is_some() || t.catch.as_ref().is_some_and(|c| !c.is_resume_block()),
            StmtKind::Var(v) => v.kind == VarKind::Using,
            _ => false,
        })
    })
}

/// Constant trip count of `for (let i = A; i < B; i++)` with literal bounds.
fn const_trip_count(s: &Stmt) -> Option<i128> {
    let StmtKind::For { init, cond: Some(cond), update, .. } = &s.kind else { return None };
    let [ast::Stmt { kind: StmtKind::Var(v), .. }] = init.as_slice() else { return None };
    let from = crate::collect::const_int(v.init.as_ref()?)?;
    let ExprKind::Binary { op, rhs, .. } = &cond.peel().kind else { return None };
    let to = crate::collect::const_int(rhs)?;
    if update.len() != 1 {
        return None;
    }
    Some(match op {
        BinaryOp::Lt => (to - from).max(0),
        BinaryOp::Le => (to - from + 1).max(0),
        BinaryOp::Gt => (from - to).max(0),
        BinaryOp::Ge => (from - to + 1).max(0),
        _ => return None,
    })
}

fn loop_decorator(s: &Stmt, name: &str) -> bool {
    let decorators = match &s.kind {
        StmtKind::For { decorators, .. } | StmtKind::ForOf { decorators, .. } | StmtKind::While { decorators, .. } | StmtKind::DoWhile { decorators, .. } => decorators,
        StmtKind::Labeled { body, .. } => return loop_decorator(body, name),
        _ => return false,
    };
    decorators.iter().any(|d| d.name.text == name)
}

fn loop_body(s: &Stmt) -> Option<&Stmt> {
    match &s.kind {
        StmtKind::For { body, .. } | StmtKind::ForOf { body, .. } | StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } => Some(body),
        _ => None,
    }
}

/// Chooses the loops that get a `DoEvents` yield point (spec §10.4.1): the innermost
/// loop of each nest whose trip count is not a small constant.
pub(crate) fn plan_yields(lw: &Lowerer, f: FnId, stmts: &[Stmt], enabled: bool, out: &mut HashSet<u32>) {
    let module = lw.prog.fns[f].module;
    let info = &lw.prog.modules[module].info;
    let threshold = lw.cfg.doevents.loop_threshold as i128;
    // Returns true when a loop inside `s` was chosen.
    fn visit(lw: &Lowerer, info: &TypeInfo, s: &Stmt, enabled: bool, threshold: i128, out: &mut HashSet<u32>) -> bool {
        let Some(body) = loop_body(s) else {
            let mut chosen = false;
            match &s.kind {
                StmtKind::Block(b) => {
                    for x in &b.stmts {
                        chosen |= visit(lw, info, x, enabled, threshold, out);
                    }
                }
                StmtKind::If { then, els, .. } => {
                    chosen |= visit(lw, info, then, enabled, threshold, out);
                    if let Some(e) = els {
                        chosen |= visit(lw, info, e, enabled, threshold, out);
                    }
                }
                StmtKind::Switch { cases, .. } => {
                    for c in cases {
                        for x in &c.body {
                            chosen |= visit(lw, info, x, enabled, threshold, out);
                        }
                    }
                }
                StmtKind::Try(t) => {
                    for x in &t.block.stmts {
                        chosen |= visit(lw, info, x, enabled, threshold, out);
                    }
                    if let Some(c) = &t.catch {
                        for x in &c.body.stmts {
                            chosen |= visit(lw, info, x, enabled, threshold, out);
                        }
                    }
                    if let Some(fin) = &t.finally {
                        for x in &fin.stmts {
                            chosen |= visit(lw, info, x, enabled, threshold, out);
                        }
                    }
                }
                StmtKind::Labeled { body, .. } => chosen |= visit(lw, info, body, enabled, threshold, out),
                _ => {}
            }
            return chosen;
        };
        let enabled_here = (enabled || loop_decorator(s, "doevents")) && !loop_decorator(s, "nodoevents");
        let inner = visit(lw, info, body, enabled_here, threshold, out);
        if !enabled_here || inner {
            return inner;
        }
        let small = matches!(const_trip_count(s), Some(n) if n < threshold);
        // A small constant loop still qualifies when its body does slow work (spec §10.4.1, 4th bullet)
        // — but then the yield point belongs to the enclosing loop.
        if small && !loop_decorator(s, "doevents") {
            return false;
        }
        let _ = (lw, info);
        out.insert(u32::from(s.span.start()));
        true
    }
    for s in stmts {
        visit(lw, info, s, enabled, threshold, out);
    }
}

enum Fin<'a> {
    Block(&'a ast::Block),
    Vba(Vec<VStmt>),
}

/// `for` header that can become `For i = a To b [Step k]` (spec §8.5).
pub(crate) struct ForShape<'a> {
    pub local: LocalId,
    pub decl: Option<&'a ast::VarDecl>,
    pub from: &'a Expr,
    pub bound: &'a Expr,
    pub op: BinaryOp,
    pub step: i64,
}

impl<'p> Lowerer<'p> {
    pub(crate) fn stmts(&mut self, stmts: &[Stmt], tail: bool) {
        for (i, s) in stmts.iter().enumerate() {
            if let StmtKind::Var(v) = &s.kind {
                if v.kind == VarKind::Using {
                    self.mark_pos(s);
                    self.using_stmt(v, &stmts[i + 1..]);
                    return;
                }
            }
            self.stmt(s, tail && i + 1 == stmts.len());
        }
    }

    fn mark_pos(&mut self, s: &Stmt) {
        // Inlined bodies keep the position of the call site.
        if self.pc.frames.len() == 1 && !matches!(s.kind, StmtKind::Block(_) | StmtKind::Empty) {
            let line = self.line_of(s.span);
            self.emit(VStmt::Pos(line));
        }
    }

    pub(crate) fn stmt(&mut self, s: &Stmt, tail: bool) {
        self.mark_pos(s);
        match &s.kind {
            StmtKind::Empty | StmtKind::Error | StmtKind::Resume | StmtKind::Fn(_) => {}
            StmtKind::Var(v) => self.var_stmt(v),
            StmtKind::Expr(e) => self.expr_stmt(e),
            StmtKind::Block(b) => self.stmts(&b.stmts, tail),
            StmtKind::If { cond, then, els } => self.if_stmt(cond, then, els.as_deref(), tail),
            StmtKind::For { .. } => self.for_stmt(s, None),
            StmtKind::ForOf { .. } => self.for_of_stmt(s, None),
            StmtKind::While { cond, body, .. } => self.while_stmt(s, cond, body, None),
            StmtKind::DoWhile { body, cond, .. } => self.do_while_stmt(s, body, cond, None),
            StmtKind::Labeled { label, body } => match &body.kind {
                StmtKind::For { .. } => self.for_stmt(body, Some(label.text.clone())),
                StmtKind::ForOf { .. } => self.for_of_stmt(body, Some(label.text.clone())),
                StmtKind::While { cond, body: b, .. } => self.while_stmt(body, cond, b, Some(label.text.clone())),
                StmtKind::DoWhile { body: b, cond, .. } => self.do_while_stmt(body, b, cond, Some(label.text.clone())),
                _ => self.stmt(body, tail),
            },
            StmtKind::Switch { disc, cases } => self.switch_stmt(disc, cases),
            StmtKind::Try(t) => self.try_stmt(t),
            StmtKind::Throw(e) => self.throw_stmt(e),
            StmtKind::Return(v) => self.return_stmt(v.as_ref(), tail),
            StmtKind::Break(label) => self.break_stmt(label.as_ref().map(|l| l.text.as_str())),
            StmtKind::Continue(label) => self.continue_stmt(label.as_ref().map(|l| l.text.as_str())),
        }
    }

    fn var_stmt(&mut self, v: &ast::VarDecl) {
        let offset = u32::from(v.name.span.start());
        let info = self.info();
        if let Some(&var) = info.decl_globals.get(&offset) {
            let ty = self.prog.vars[var].ty.clone();
            let target = if self.is_demoted(var) {
                // No longer a module-level name: the local may keep it.
                self.global_names.remove(&v.name.text.to_lowercase());
                let n = self.local_name(&v.name.text, &ty);
                self.pc.demoted.insert(var, n.clone());
                VExpr::Name(n)
            } else {
                if self.const_literal(var).is_some() {
                    return; // emitted as a VBA `Const`
                }
                self.var_ref(var)
            };
            if let Some(init) = &v.init {
                self.expr_into(init, &target, &ty);
            }
            return;
        }
        let Some(&local) = info.decl_locals.get(&offset) else { return };
        let ty = self.local_ty(local);
        let n = self.local_name(&v.name.text, &ty);
        self.bind(local, Binding::Var(n.clone()));
        match &v.init {
            Some(init) => self.expr_into(init, &VExpr::Name(n), &ty),
            // `Dim` runs once per procedure: re-initialise explicitly where the declaration can run again (spec §13.1).
            None if self.pc.loop_depth > 0 || self.frame().inl.is_some() => self.reset_var(&VExpr::Name(n), &ty),
            None => {}
        }
    }

    fn scoped(&mut self, s: &Stmt, tail: bool) -> Vec<VStmt> {
        self.capture(|lw| lw.stmt(s, tail)).0
    }

    fn if_stmt(&mut self, cond: &Expr, then: &Stmt, els: Option<&Stmt>, tail: bool) {
        // `if (Platform.is64bit)` must be conditional compilation: `LongLong` does not compile on 32-bit Office.
        if matches!(self.info().res.get(&cond.peel().id), Some(Res::Intrinsic(Intr::PlatformIs64))) {
            self.emit(VStmt::Raw("#If Win64 Then".into()));
            let t = self.scoped(then, tail);
            for x in t {
                self.emit(x);
            }
            if let Some(e) = els {
                self.emit(VStmt::Raw("#Else".into()));
                let b = self.scoped(e, tail);
                for x in b {
                    self.emit(x);
                }
            }
            self.emit(VStmt::Raw("#End If".into()));
            return;
        }
        let c = self.cond(cond);
        let then_b = self.scoped(then, tail);
        let else_b = match els {
            Some(e) => self.scoped(e, tail),
            None => Vec::new(),
        };
        self.emit(flatten_if(c, then_b, else_b));
    }

    // ------------------------------------------------------------------ loops

    fn yield_point(&mut self, s: &Stmt) -> Vec<VStmt> {
        let plan = self.current_yield_plan();
        if self.pc.suppress_yield || !plan.contains(&u32::from(s.span.start())) {
            return Vec::new();
        }
        self.pc.yield_n += 1;
        let counter = self.declare_raw(&format!("ras_y{}", self.pc.yield_n), "Long", false);
        self.note(ast::Span::at(s.span.start(), 3.into()), true, "doevents", "DoEvents を自動挿入しました（間引き付き）");
        vec![
            assign(name(counter.clone()), name(counter.clone()).plus(1)),
            VStmt::Raw(format!("If ({counter} And ras_yieldMask) = 0& Then RAS_Yield")),
        ]
    }

    pub(crate) fn enter_loop(&mut self, kind: LoopKind, label: Option<String>) {
        let tries_base = self.pc.tries.len();
        self.pc.breakables.push(Breakable { kind, label, break_label: None, continue_label: None, tries_base });
        self.pc.loop_depth += 1;
    }

    /// Lowers a loop body; appends the `continue` label when it was used.
    fn loop_body(&mut self, body: &Stmt) -> Vec<VStmt> {
        let mut b = self.scoped(body, false);
        if let Some(l) = self.pc.breakables.last().and_then(|x| x.continue_label.clone()) {
            b.push(VStmt::Label(l));
        }
        b
    }

    fn leave_loop(&mut self) {
        self.pc.loop_depth -= 1;
        let b = self.pc.breakables.pop().unwrap();
        if let Some(l) = b.break_label {
            self.emit(VStmt::Label(l));
        }
    }

    fn while_stmt(&mut self, s: &Stmt, cond: &Expr, body: &Stmt, label: Option<String>) {
        let caps = self.push_plan(s);
        self.enter_loop(LoopKind::Do, label);
        let always = matches!(cond.peel().kind, ExprKind::Bool(true));
        let (pre, c) = self.capture(|lw| lw.cond(cond));
        let mut b = self.loop_body(body);
        b.extend(self.yield_point(s));
        if always {
            self.emit(VStmt::Do { pre: None, post: None, body: b });
        } else if pre.is_empty() {
            self.emit(VStmt::Do { pre: Some(c), post: None, body: b });
        } else {
            // The condition needs statements: evaluate it inside the loop (spec §8.3).
            let mut full = pre;
            full.push(if_then(c.not(), vec![VStmt::Exit("Do")]));
            full.extend(b);
            self.emit(VStmt::Do { pre: None, post: None, body: full });
        }
        self.leave_loop();
        self.finish_push_plan(caps);
    }

    fn do_while_stmt(&mut self, s: &Stmt, body: &Stmt, cond: &Expr, label: Option<String>) {
        let caps = self.push_plan(s);
        self.enter_loop(LoopKind::Do, label);
        let mut b = self.loop_body(body);
        b.extend(self.yield_point(s));
        let (pre, c) = self.capture(|lw| lw.cond(cond));
        b.extend(pre);
        self.emit(VStmt::Do { pre: None, post: Some(c), body: b });
        self.leave_loop();
        self.finish_push_plan(caps);
    }

    pub(crate) fn local_of(&self, e: &Expr) -> Option<LocalId> {
        match (&e.peel().kind, self.info().res.get(&e.peel().id)) {
            (ExprKind::Name(_), Some(Res::Local(l))) => Some(*l),
            _ => None,
        }
    }

    /// True when `local` may be modified inside `s` (assignment, `++`, or passed where it can be changed).
    pub(crate) fn modifies_local(&self, s: &Stmt, local: LocalId) -> bool {
        let info = self.info();
        let is_local = |e: &Expr| matches!((&e.peel().kind, info.res.get(&e.peel().id)), (ExprKind::Name(_), Some(Res::Local(l))) if *l == local);
        any_expr(s, &mut |e| match &e.kind {
            ExprKind::Assign { target, .. } => is_local(target),
            ExprKind::IncDec { operand, .. } => is_local(operand),
            ExprKind::Call { callee, args, .. } => {
                // byref arguments and mutating methods (`arr.push`)
                let target = info.calls.get(&e.id).map(|c| &c.target);
                let byref_arg = match target {
                    Some(CallTarget::Fn(f) | CallTarget::Method(f) | CallTarget::Ext(f) | CallTarget::Host(f)) => {
                        let def = &self.prog.fns[*f];
                        let skip = usize::from(matches!(target, Some(CallTarget::Ext(_))));
                        let recv_byref = matches!(target, Some(CallTarget::Ext(_)))
                            && def.params.first().is_some_and(|p| p.mode == ast::PassMode::ByRef || p.ty.is_array())
                            && matches!(&callee.peel().kind, ExprKind::Member { object, .. } if is_local(object));
                        recv_byref
                            || args.iter().enumerate().any(|(i, a)| {
                                is_local(a) && def.params.get(i + skip).is_none_or(|p| p.mode == ast::PassMode::ByRef || p.ty.is_array())
                            })
                    }
                    Some(CallTarget::Intrinsic(i)) => {
                        matches!(i, Intr::ArrPush | Intr::ArrPop | Intr::ArrSort | Intr::ArrReverse | Intr::ArrClear)
                            && matches!(&callee.peel().kind, ExprKind::Member { object, .. } if is_local(object))
                    }
                    Some(CallTarget::Dynamic | CallTarget::Closure | CallTarget::DefaultMember(_)) => args.iter().any(&is_local),
                    _ => false,
                };
                byref_arg
            }
            _ => false,
        })
    }


    /// Is `e` guaranteed to keep its value while `body` runs?
    pub(crate) fn is_invariant(&self, e: &Expr, body: &Stmt) -> bool {
        let info = self.info();
        match &e.peel().kind {
            ExprKind::Int { .. } | ExprKind::Float { .. } => true,
            ExprKind::Name(_) => match info.res.get(&e.peel().id) {
                Some(Res::Local(l)) => !self.modifies_local(body, *l) && !self.frame_local_is_byref(*l),
                Some(Res::Global(v)) => {
                    let d = &self.prog.vars[*v];
                    (d.is_const && !d.declared) || (!d.declared && !d.reassigned)
                }
                Some(Res::EnumMember(..)) => true,
                _ => false,
            },
            ExprKind::Member { object, .. } => {
                // `arr.length` of a local array that the body does not touch otherwise
                matches!(info.res.get(&e.peel().id), Some(Res::Intrinsic(Intr::ArrLength | Intr::StrLength)))
                    && match self.local_of(object) {
                        Some(l) => !self.modifies_local(body, l) && !self.frame_local_is_byref(l),
                        None => false,
                    }
            }
            ExprKind::Unary { op: ast::UnaryOp::Neg, operand } => self.is_invariant(operand, body),
            ExprKind::Binary { op: BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul, lhs, rhs } => self.is_invariant(lhs, body) && self.is_invariant(rhs, body),
            _ => false,
        }
    }

    fn frame_local_is_byref(&self, l: LocalId) -> bool {
        self.frame().fn_id.and_then(|f| self.prog.fns[f].locals.get(l)).is_some_and(|d| d.byref)
    }

    pub(crate) fn for_shape<'a>(&self, init: &'a [Stmt], cond: Option<&'a Expr>, update: &'a [Expr], body: &'a Stmt) -> Option<ForShape<'a>> {
        let info = self.info();
        let [first] = init else { return None };
        let (local, decl, from) = match &first.kind {
            StmtKind::Var(v) => {
                let l = *info.decl_locals.get(&u32::from(v.name.span.start()))?;
                (l, Some(v), v.init.as_ref()?)
            }
            StmtKind::Expr(e) => match &e.kind {
                ExprKind::Assign { op: None, target, value } => (self.local_of(target)?, None, &**value),
                _ => return None,
            },
            _ => return None,
        };
        if !self.local_ty(local).is_integer() {
            return None;
        }
        let ExprKind::Binary { op, lhs, rhs } = &cond?.peel().kind else { return None };
        if self.local_of(lhs)? != local || !matches!(op, BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge) {
            return None;
        }
        let [upd] = update else { return None };
        let step = match &upd.peel().kind {
            ExprKind::IncDec { inc, operand, .. } if self.local_of(operand)? == local => {
                if *inc {
                    1
                } else {
                    -1
                }
            }
            ExprKind::Assign { op: Some(o @ (BinaryOp::Add | BinaryOp::Sub)), target, value } if self.local_of(target)? == local => {
                let k = i64::try_from(crate::collect::const_int(value)?).ok()?;
                if k <= 0 {
                    return None;
                }
                if *o == BinaryOp::Add {
                    k
                } else {
                    -k
                }
            }
            _ => return None,
        };
        let ascending = matches!(op, BinaryOp::Lt | BinaryOp::Le);
        if ascending != (step > 0) {
            return None;
        }
        // VBA evaluates the bound once: only valid when neither the variable nor the bound changes in the body.
        if self.modifies_local(body, local) || !self.is_invariant(rhs, body) {
            return None;
        }
        // The bound must not depend on the loop variable itself.
        if any_expr_in(rhs, &mut |x| self.local_of(x) == Some(local)) {
            return None;
        }
        Some(ForShape { local, decl, from, bound: rhs, op: *op, step })
    }

    fn for_stmt(&mut self, s: &Stmt, label: Option<String>) {
        // Range ⇔ array bulk transfer (spec §9)
        if self.pc.bulk.is_empty() {
            if let Some(plan) = self.bulk_plan(s) {
                self.emit_bulk(s, plan, label);
                return;
            }
        }
        self.for_stmt_plain(s, label);
    }

    pub(crate) fn for_stmt_plain(&mut self, s: &Stmt, label: Option<String>) {
        let StmtKind::For { init, cond, update, body, .. } = &s.kind else { return };
        let caps = self.push_plan(s);
        if let Some(shape) = self.for_shape(init, cond.as_ref(), update, body) {
            let from = self.expr(shape.from);
            let bound = self.expr(shape.bound);
            let to = match shape.op {
                BinaryOp::Lt => bound.plus(-1),
                BinaryOp::Gt => bound.plus(1),
                _ => bound,
            };
            let var = match shape.decl {
                Some(v) => match self.frame().locals.get(&shape.local) {
                    // already bound by the bulk-transfer plan
                    Some(Binding::Var(n)) => n.clone(),
                    _ => {
                        let ty = self.local_ty(shape.local);
                        let n = self.local_name(&v.name.text, &ty);
                        self.bind(shape.local, Binding::Var(n.clone()));
                        n
                    }
                },
                None => self.local_expr(shape.local).print(),
            };
            self.enter_loop(LoopKind::For, label);
            let mut b = self.loop_body(body);
            b.extend(self.yield_point(s));
            let step = (shape.step != 1).then(|| long(shape.step));
            self.emit(VStmt::For { var, from, to, step, body: b });
            self.leave_loop();
            self.finish_push_plan(caps);
            return;
        }
        // General form: init; Do While cond ... update; Loop
        for i in init {
            self.stmt(i, false);
        }
        self.enter_loop(LoopKind::Do, label);
        let (pre, c) = match cond {
            Some(c) => {
                let (pre, v) = self.capture(|lw| lw.cond(c));
                (pre, Some(v))
            }
            None => (Vec::new(), None),
        };
        let mut b = self.loop_body(body);
        let (upd, _) = self.capture(|lw| {
            for u in update {
                lw.expr_stmt(u);
            }
        });
        b.extend(upd);
        b.extend(self.yield_point(s));
        match c {
            None => self.emit(VStmt::Do { pre: None, post: None, body: b }),
            Some(c) if pre.is_empty() => self.emit(VStmt::Do { pre: Some(c), post: None, body: b }),
            Some(c) => {
                let mut full = pre;
                full.push(if_then(c.not(), vec![VStmt::Exit("Do")]));
                full.extend(b);
                self.emit(VStmt::Do { pre: None, post: None, body: full });
            }
        }
        self.leave_loop();
        self.finish_push_plan(caps);
    }

    fn for_of_stmt(&mut self, s: &Stmt, label: Option<String>) {
        let StmtKind::ForOf { name: var_name, iter, body, .. } = &s.kind else { return };
        let info = self.info();
        let Some(&local) = info.decl_locals.get(&u32::from(var_name.span.start())) else { return };
        let var_ty = self.local_ty(local);
        let iter_ty = self.ty(iter);
        let caps = self.push_plan(s);
        // `for (const k of map.keys())` enumerates the dictionary directly.
        let map_iter = match (&iter.peel().kind, info.calls.get(&iter.peel().id).map(|c| &c.target)) {
            (ExprKind::Call { callee, .. }, Some(CallTarget::Intrinsic(i @ (Intr::MapKeys | Intr::MapValues)))) => match &callee.peel().kind {
                ExprKind::Member { object, .. } => Some((*i, &**object)),
                _ => None,
            },
            _ => None,
        };
        if iter_ty.is_array() && map_iter.is_none() {
            let src = self.stable_array(iter);
            let idx = self.declare_raw("ras_i", "Long", false);
            let n = self.local_name(&var_name.text, &var_ty);
            self.bind(local, Binding::Var(n.clone()));
            self.enter_loop(LoopKind::For, label);
            let (mut b, _) = self.capture(|lw| {
                let elem_ty = iter_ty.elem().cloned().unwrap_or(Type::Any);
                lw.store(&VExpr::Name(n.clone()), &var_ty, src.clone().index(vec![name(idx.clone())]), &elem_ty);
            });
            b.extend(self.loop_body(body));
            b.extend(self.yield_point(s));
            self.emit(VStmt::For { var: idx, from: long(0), to: call("RAS_Len", vec![src]).plus(-1), step: None, body: b });
            self.leave_loop();
            self.finish_push_plan(caps);
            return;
        }
        // Collections: `For Each`. The control variable must be a Variant or an object.
        let coll = match map_iter {
            Some((i, obj)) => {
                let o = self.expr(obj);
                o.member(if i == Intr::MapKeys { "Keys" } else { "Items" })
            }
            None => self.expr(iter),
        };
        let n = self.local_name(&var_name.text, &var_ty);
        self.bind(local, Binding::Var(n.clone()));
        let direct = var_ty.is_object() || matches!(var_ty, Type::Any) || self.vtype(&var_ty).0 == "Variant";
        self.enter_loop(LoopKind::For, label);
        let (control, mut b) = if direct && !var_ty.is_array() {
            (n, Vec::new())
        } else {
            let tmp = self.declare_raw("ras_v", "Variant", false);
            let (pre, _) = self.capture(|lw| lw.store(&VExpr::Name(n.clone()), &var_ty, name(tmp.clone()), &Type::Any));
            (tmp, pre)
        };
        b.extend(self.loop_body(body));
        b.extend(self.yield_point(s));
        self.emit(VStmt::ForEach { var: control, coll, body: b });
        self.leave_loop();
        self.finish_push_plan(caps);
    }

    // ------------------------------------------------------------------ push capacity doubling (spec §13.3)

    /// Finds local arrays that are only pushed / indexed / measured inside loop `s`
    /// and switches them to capacity-doubling mode for the duration of the loop.
    fn push_plan(&mut self, s: &Stmt) -> Vec<(String, String, String, Type)> {
        if self.opt == 0 || self.pc.loop_depth > 0 || !self.pc.tries.is_empty() {
            return Vec::new();
        }
        let info = self.info();
        let mut pushed: Vec<LocalId> = Vec::new();
        any_expr(s, &mut |e| {
            if let (ExprKind::Call { callee, .. }, Some(CallTarget::Intrinsic(Intr::ArrPush))) = (&e.kind, info.calls.get(&e.id).map(|c| &c.target)) {
                if let ExprKind::Member { object, .. } = &callee.peel().kind {
                    if let Some(l) = self.local_of(object) {
                        if !pushed.contains(&l) {
                            pushed.push(l);
                        }
                    }
                }
            }
            false
        });
        if pushed.is_empty() {
            return Vec::new();
        }
        // An early exit would leave the array over-allocated.
        let escapes = any_stmt(s, &mut |x| matches!(x.kind, StmtKind::Return(_) | StmtKind::Throw(_) | StmtKind::Try(_)));
        if escapes || self.has_escaping_calls(s) {
            return Vec::new();
        }
        let mut out = Vec::new();
        for l in pushed {
            let Some(Binding::Var(arr)) = self.frame().locals.get(&l).cloned() else { continue };
            let def_ok = self.frame().fn_id.and_then(|f| self.prog.fns[f].locals.get(l)).is_some_and(|d| matches!(d.kind, LocalKind::Let | LocalKind::Const) && !d.captured);
            if !def_ok || self.pc.push_caps.contains_key(&arr) {
                continue;
            }
            // Every use of the array inside the loop must be `a.push(x)`, `a.length` or `a[i]`.
            let mut ok = true;
            let mut allowed: HashSet<ast::ExprId> = HashSet::new();
            crate::astutil::walk_stmt(s, &mut |_| true, &mut |e| {
                match &e.kind {
                    ExprKind::Member { object, .. } if self.local_of(object) == Some(l) => {
                        if matches!(info.res.get(&e.id), Some(Res::Intrinsic(Intr::ArrPush | Intr::ArrLength))) {
                            allowed.insert(object.peel().id);
                        }
                    }
                    ExprKind::Index { object, .. } if self.local_of(object) == Some(l) => {
                        allowed.insert(object.peel().id);
                    }
                    ExprKind::Name(_) if self.local_of(e) == Some(l) && !allowed.contains(&e.id) => ok = false,
                    _ => {}
                }
                true
            });
            if !ok {
                continue;
            }
            let ty = self.local_ty(l);
            let len = self.declare_raw(&format!("ras_len_{arr}"), "Long", false);
            let cap = self.declare_raw(&format!("ras_cap_{arr}"), "Long", false);
            self.emit(assign(name(len.clone()), call("RAS_Len", vec![name(arr.clone())])));
            self.emit(assign(name(cap.clone()), name(len.clone())));
            self.pc.push_caps.insert(arr.clone(), (len.clone(), cap.clone()));
            out.push((arr, len, cap, ty));
        }
        out
    }

    /// Calls that could raise or re-enter and observe the over-allocated array.
    fn has_escaping_calls(&self, s: &Stmt) -> bool {
        let info = self.info();
        any_expr(s, &mut |e| matches!(&e.kind, ExprKind::Arrow(_)) || (matches!(e.kind, ExprKind::Call { .. }) && matches!(info.calls.get(&e.id).map(|c| &c.target), Some(CallTarget::Closure))))
    }

    fn finish_push_plan(&mut self, caps: Vec<(String, String, String, Type)>) {
        for (arr, len, cap, ty) in caps {
            self.pc.push_caps.remove(&arr);
            // Trim to the real size after the loop.
            let (_, is_typed_array) = self.vtype(&ty);
            let clear = if is_typed_array { VStmt::Erase(name(arr.clone())) } else { assign(name(arr.clone()), lit("Empty")) };
            self.emit(VStmt::If {
                cond: bin("=", name(len.clone()), long(0)),
                then: vec![clear],
                elifs: vec![(bin("<", name(len.clone()), name(cap)), vec![VStmt::ReDim { preserve: true, target: name(arr), dims: vec![(long(0), name(len).plus(-1))] }])],
                els: vec![],
            });
        }
    }

    // ------------------------------------------------------------------ switch

    fn switch_stmt(&mut self, disc: &Expr, cases: &[ast::SwitchCase]) {
        let d = self.expr(disc);
        let tries_base = self.pc.tries.len();
        self.pc.breakables.push(Breakable { kind: LoopKind::Switch, label: None, break_label: None, continue_label: None, tries_base });
        let mut out_cases: Vec<(Vec<VExpr>, Vec<VStmt>)> = Vec::new();
        let mut default: Option<Vec<VStmt>> = None;
        let mut pending_tests: Vec<VExpr> = Vec::new();
        let mut pending_default = false;
        let mut fall_label: Option<String> = None;
        for (i, c) in cases.iter().enumerate() {
            match &c.test {
                Some(t) => {
                    let v = self.expr(t);
                    pending_tests.push(v);
                }
                None => pending_default = true,
            }
            // `case 1: case 2: body` shares one body.
            if c.body.is_empty() && i + 1 < cases.len() {
                continue;
            }
            let mut body = c.body.as_slice();
            let ends_with_break = matches!(body.last().map(|s| &s.kind), Some(StmtKind::Break(None)));
            if ends_with_break {
                body = &body[..body.len() - 1];
            }
            let terminates = ends_with_break || matches!(body.last().map(|s| &s.kind), Some(StmtKind::Return(_) | StmtKind::Throw(_) | StmtKind::Continue(_) | StmtKind::Break(Some(_))));
            let incoming = fall_label.take();
            let (mut b, _) = self.capture(|lw| {
                if let Some(l) = incoming {
                    lw.emit(VStmt::Label(l));
                }
                lw.stmts(body, false);
            });
            // Fall-through into the next case is expressed with GoTo (spec §13.5).
            if !terminates && i + 1 < cases.len() && !body.is_empty() {
                let l = self.label("RAS_Case");
                b.push(VStmt::GoTo(l.clone()));
                fall_label = Some(l);
            }
            if pending_default {
                if i + 1 != cases.len() {
                    self.err(c.span, "E0620", "`default` は `switch` の最後に書いてください");
                }
                default = Some(b);
            } else {
                out_cases.push((std::mem::take(&mut pending_tests), b));
            }
            pending_tests.clear();
            pending_default = false;
        }
        self.emit(VStmt::Select { expr: d, cases: out_cases, els: default });
        let b = self.pc.breakables.pop().unwrap();
        if let Some(l) = b.break_label {
            self.emit(VStmt::Label(l));
        }
    }

    // ------------------------------------------------------------------ try / catch / finally (spec §5)

    fn restore_on_error(&self) -> VStmt {
        VStmt::Raw(if self.pc.uses_handler { "On Error GoTo RAS_Handler".into() } else { "On Error GoTo 0".into() })
    }

    fn set_try_id(&mut self, id: u32) {
        self.pc.uses_try_id = true;
        self.pc.cur_try_id = id;
        self.emit(assign(name("ras_tryId"), long(id as i64)));
    }

    fn new_try_id(&mut self) -> u32 {
        self.pc.try_n += 1;
        self.pc.try_n
    }

    fn try_stmt(&mut self, t: &ast::TryStmt) {
        match (&t.catch, &t.finally) {
            (Some(c), _) if c.is_resume_block() => {
                // spec §5.2: errors inside the block are ignored.
                self.emit(VStmt::Raw("On Error Resume Next".into()));
                let outer_id = self.pc.cur_try_id;
                self.pc.tries.push(TryCtx { kind: TryKind::ResumeNext, outer_id });
                self.stmts(&t.block.stmts, false);
                self.pc.tries.pop();
                let restore = self.restore_on_error();
                self.emit(restore);
            }
            (Some(c), None) => self.try_catch(&t.block, c),
            (c, Some(f)) => self.try_finally(&t.block.stmts, c.as_ref(), Fin::Block(f)),
            (None, None) => self.stmts(&t.block.stmts, false),
        }
    }

    fn bind_catch_param(&mut self, c: &ast::CatchClause) {
        let Some(p) = &c.param else { return };
        let Some(&local) = self.info().decl_locals.get(&u32::from(p.span.start())) else { return };
        let used = self.frame().fn_id.and_then(|f| self.prog.fns[f].locals.get(local)).is_none_or(|l| l.used);
        let ty = self.local_ty(local);
        let n = self.local_name(&p.text, &ty);
        self.bind(local, Binding::Var(n.clone()));
        if used {
            self.emit(set(name(n), call("RAS_CaughtError", vec![])));
        }
    }

    fn try_catch(&mut self, block: &ast::Block, c: &ast::CatchClause) {
        let outer = self.pc.cur_try_id;
        let id = self.new_try_id();
        self.set_try_id(id);
        self.pc.tries.push(TryCtx { kind: TryKind::Catch, outer_id: outer });
        self.stmts(&block.stmts, false);
        self.pc.tries.pop();
        self.set_try_id(outer);
        let end = self.label("RAS_EndTry");
        let catch = self.label("RAS_Catch");
        self.emit(VStmt::GoTo(end.clone()));
        self.emit(VStmt::Label(catch.clone()));
        // The handler leaves via `Resume`, so errors in the catch body reach the outer try (spec §5.1).
        self.set_try_id(outer);
        self.bind_catch_param(c);
        self.stmts(&c.body.stmts, false);
        self.emit(VStmt::Label(end));
        self.pc.handler_arms.push((id, vec![VStmt::Raw(format!("Resume {catch}"))]));
    }

    fn try_finally(&mut self, body: &[Stmt], catch: Option<&ast::CatchClause>, fin: Fin) {
        let outer = self.pc.cur_try_id;
        let id = self.new_try_id();
        let exit_var = self.declare_raw(&format!("ras_exit{id}"), "Long", false);
        let pend = self.declare_raw(&format!("ras_pend{id}"), "Object", false);
        let fin_label = self.label("RAS_Finally");
        self.emit(assign(name(exit_var.clone()), long(0)));
        self.set_try_id(id);
        // Exit code 1 is always "re-raise the pending error".
        self.pc.tries.push(TryCtx { kind: TryKind::Finally { label: fin_label.clone(), exit_var: exit_var.clone(), pending: vec![PendingExit::Rethrow] }, outer_id: outer });
        self.stmts(body, false);
        let to_finally = |id: u32| {
            (id, vec![assign(name(exit_var.clone()), long(1)), set(name(pend.clone()), call("RAS_CaughtError", vec![])), VStmt::Raw(format!("Resume {fin_label}"))])
        };
        if let Some(c) = catch {
            let catch_label = self.label("RAS_Catch");
            let catch_id = self.new_try_id();
            self.emit(VStmt::GoTo(fin_label.clone()));
            self.emit(VStmt::Label(catch_label.clone()));
            self.set_try_id(catch_id);
            self.bind_catch_param(c);
            self.stmts(&c.body.stmts, false);
            self.pc.handler_arms.push((id, vec![VStmt::Raw(format!("Resume {catch_label}"))]));
            self.pc.handler_arms.push(to_finally(catch_id));
        } else {
            self.pc.handler_arms.push(to_finally(id));
        }
        let ctx = self.pc.tries.pop().unwrap();
        self.emit(VStmt::Label(fin_label.clone()));
        self.set_try_id(outer);
        match fin {
            Fin::Block(b) => self.stmts(&b.stmts, false),
            Fin::Vba(stmts) => {
                for s in stmts {
                    self.emit(s);
                }
            }
        }
        // Continue with whatever left the try block.
        let TryKind::Finally { pending, .. } = ctx.kind else { return };
        let mut cases = Vec::new();
        for (i, p) in pending.into_iter().enumerate() {
            let (b, _) = self.capture(|lw| match p {
                PendingExit::Rethrow => lw.emit(call_stmt("RAS_Throw", vec![name(pend.clone())])),
                PendingExit::Return => lw.leave_function(false),
                PendingExit::Jump(label, base) => lw.exit_through(base, Some(label)),
            });
            cases.push((vec![long(i as i64 + 1)], b));
        }
        self.emit(VStmt::Select { expr: name(exit_var), cases, els: None });
    }

    fn using_stmt(&mut self, v: &ast::VarDecl, rest: &[Stmt]) {
        let Some(init) = &v.init else { return };
        let info = self.info();
        let is_fast = matches!(info.calls.get(&init.peel().id).map(|c| &c.target), Some(CallTarget::Intrinsic(Intr::FastMode)));
        let fin = if is_fast {
            let mask = super::expr::fast_mode_mask(init);
            if mask == 63 {
                self.emit(call_stmt("RAS_SuspendAll", vec![]));
                vec![call_stmt("RAS_ResumeAll", vec![])]
            } else {
                self.emit(call_stmt("RAS_Suspend", vec![long(mask)]));
                vec![call_stmt("RAS_Resume", vec![long(mask)])]
            }
        } else {
            let Some(&local) = info.decl_locals.get(&u32::from(v.name.span.start())) else { return };
            let ty = self.local_ty(local);
            let n = self.local_name(&v.name.text, &ty);
            self.bind(local, Binding::Var(n.clone()));
            self.expr_into(init, &VExpr::Name(n.clone()), &ty);
            vec![if_then(bin("Is", name(n.clone()), lit("Nothing")).not_is(), vec![VStmt::Call { target: name(n).member("dispose"), args: vec![] }])]
        };
        self.try_finally(rest, None, Fin::Vba(fin));
    }

    fn throw_stmt(&mut self, e: &Expr) {
        let src = lit(format!("\"{}\"", self.pc.display_name));
        let raise = VExpr::Name("Err".into()).member("Raise");
        let default_code = bin("+", name("vbObjectError"), long(513));
        match &e.peel().kind {
            ExprKind::New { ty: ast::TypeRef::Named { path, .. }, args } if path.len() == 1 && path[0].text == "Error" => {
                let msg = args.first().map(|a| self.expr(a)).unwrap_or_else(|| lit("\"\""));
                let code = args.get(1).map(|a| self.expr(a)).unwrap_or(default_code);
                self.emit(VStmt::Call { target: raise, args: vec![code, src, msg] });
            }
            _ if self.ty(e) == Type::Str => {
                let msg = self.expr(e);
                self.emit(VStmt::Call { target: raise, args: vec![default_code, src, msg] });
            }
            _ => {
                let v = self.expr(e);
                self.emit(call_stmt("RAS_Throw", vec![v]));
            }
        }
    }

    // ------------------------------------------------------------------ exits

    fn return_stmt(&mut self, value: Option<&Expr>, tail: bool) {
        if let Some(v) = value {
            match self.pc.rets.last() {
                Some(RetMode::Inline { var: Some(var), ty, .. }) => {
                    let (var, ret_ty) = (var.clone(), ty.clone());
                    self.expr_into(v, &VExpr::Name(var), &ret_ty);
                }
                Some(RetMode::Inline { var: None, .. }) => self.expr_stmt(v),
                _ => {
                    if self.pc.ret_ty == Type::Void {
                        self.expr_stmt(v);
                    } else {
                        let target = VExpr::Name(self.pc.proc_name.clone());
                        let ty = self.pc.ret_ty.clone();
                        self.expr_into(v, &target, &ty);
                    }
                }
            }
        }
        self.leave_function(tail);
    }

    /// Leaves the current (possibly inlined) function, running enclosing `finally` blocks first.
    pub(crate) fn leave_function(&mut self, tail: bool) {
        let base = match self.pc.rets.last() {
            Some(RetMode::Inline { tries_base, .. }) => *tries_base,
            _ => 0,
        };
        if self.route_through_finally(base, || PendingExit::Return) {
            return;
        }
        match self.pc.rets.last_mut() {
            Some(RetMode::Inline { label, used, .. }) => {
                if !tail {
                    *used = true;
                    let l = label.clone();
                    self.restore_try_state(base);
                    self.emit(VStmt::GoTo(l));
                }
            }
            _ => {
                if tail {
                    return;
                }
                match self.pc.exit_label.clone() {
                    Some(l) => {
                        self.pc.exit_label_used = true;
                        self.emit(VStmt::GoTo(l));
                    }
                    None => self.emit(VStmt::Exit(self.pc.exit_kw)),
                }
            }
        }
    }

    /// If a `finally` region lies between the current position and `base`, records the pending
    /// exit there and jumps to its `finally` block. Returns true when it did.
    fn route_through_finally(&mut self, base: usize, make: impl FnOnce() -> PendingExit) -> bool {
        let mut resume_next = false;
        for i in (base..self.pc.tries.len()).rev() {
            match &mut self.pc.tries[i].kind {
                TryKind::ResumeNext => resume_next = true,
                TryKind::Catch => {}
                TryKind::Finally { label, exit_var, pending } => {
                    pending.push(make());
                    let (code, label, exit_var) = (pending.len() as i64, label.clone(), exit_var.clone());
                    if resume_next {
                        let r = self.restore_on_error();
                        self.emit(r);
                    }
                    self.emit(assign(name(exit_var), long(code)));
                    self.emit(VStmt::GoTo(label));
                    return true;
                }
            }
        }
        false
    }

    /// Restores `On Error` / `ras_tryId` before jumping out of the try regions above `base`.
    fn restore_try_state(&mut self, base: usize) {
        if self.pc.tries.len() <= base {
            return;
        }
        if self.pc.tries[base..].iter().any(|t| matches!(t.kind, TryKind::ResumeNext)) {
            let r = self.restore_on_error();
            self.emit(r);
        }
        let outer = self.pc.tries[base].outer_id;
        if self.pc.tries[base..].iter().any(|t| !matches!(t.kind, TryKind::ResumeNext)) {
            self.emit(assign(name("ras_tryId"), long(outer as i64)));
        }
    }

    /// Jumps to `label` (or performs `Exit`-less fallthrough when `None`), unwinding try regions above `base`.
    pub(crate) fn exit_through(&mut self, base: usize, label: Option<String>) {
        let Some(label) = label else { return };
        let l2 = label.clone();
        if self.route_through_finally(base, move || PendingExit::Jump(l2, base)) {
            return;
        }
        self.restore_try_state(base);
        self.emit(VStmt::GoTo(label));
    }

    fn break_stmt(&mut self, label: Option<&str>) {
        let idx = match label {
            Some(l) => self.pc.breakables.iter().rposition(|b| b.label.as_deref() == Some(l)),
            None => self.pc.breakables.len().checked_sub(1),
        };
        let Some(idx) = idx else { return };
        let innermost_loop = self.pc.breakables.iter().rposition(|b| b.kind != LoopKind::Switch);
        let base = self.pc.breakables[idx].tries_base;
        let has_finally = self.pc.tries[base.min(self.pc.tries.len())..].iter().any(|t| matches!(t.kind, TryKind::Finally { .. }));
        let kind = self.pc.breakables[idx].kind.clone();
        if kind != LoopKind::Switch && Some(idx) == innermost_loop && !has_finally {
            self.restore_try_state(base);
            self.emit(VStmt::Exit(if kind == LoopKind::For { "For" } else { "Do" }));
            return;
        }
        // `Exit For` only leaves one level: use GoTo for outer loops (spec §13.5).
        let l = match &self.pc.breakables[idx].break_label {
            Some(l) => l.clone(),
            None => {
                let l = self.label(if kind == LoopKind::Switch { "RAS_EndSwitch" } else { "RAS_Break" });
                self.pc.breakables[idx].break_label = Some(l.clone());
                l
            }
        };
        self.exit_through(base, Some(l));
    }

    fn continue_stmt(&mut self, label: Option<&str>) {
        let idx = match label {
            Some(l) => self.pc.breakables.iter().rposition(|b| b.label.as_deref() == Some(l)),
            None => self.pc.breakables.iter().rposition(|b| b.kind != LoopKind::Switch),
        };
        let Some(idx) = idx else { return };
        let base = self.pc.breakables[idx].tries_base;
        let l = match &self.pc.breakables[idx].continue_label {
            Some(l) => l.clone(),
            None => {
                let l = self.label("RAS_Continue");
                self.pc.breakables[idx].continue_label = Some(l.clone());
                l
            }
        };
        self.exit_through(base, Some(l));
    }

}

/// Builds `If c Then ... ElseIf ... Else ... End If`, folding `else { if ... }` chains.
pub(crate) fn flatten_if(cond: VExpr, then: Vec<VStmt>, mut els: Vec<VStmt>) -> VStmt {
    let significant: Vec<usize> = els.iter().enumerate().filter(|(_, s)| !matches!(s, VStmt::Pos(_))).map(|(i, _)| i).collect();
    if significant.len() == 1 {
        if let VStmt::If { .. } = &els[significant[0]] {
            let VStmt::If { cond: c2, then: t2, elifs: e2, els: l2 } = els.remove(significant[0]) else { unreachable!() };
            let mut elifs = vec![(c2, t2)];
            elifs.extend(e2);
            return VStmt::If { cond, then, elifs, els: l2 };
        }
    }
    VStmt::If { cond, then, elifs: Vec::new(), els }
}

trait NotIs {
    fn not_is(self) -> VExpr;
}

impl NotIs for VExpr {
    /// `Not (x Is Nothing)`
    fn not_is(self) -> VExpr {
        VExpr::Not(Box::new(self))
    }
}
