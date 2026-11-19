//! Lambdas: loop fusion for array methods (spec §4.1 rule 1) and closure
//! classes for lambdas that escape (rule 2).

use std::collections::HashMap;

use ras_syntax::ast::{self, ArrowBody, Expr, ExprKind};

use super::expr::mentions;
use super::stmt::needs_handler;
use super::{Binding, Frame, Lowerer, RetMode};
use crate::astutil::any_expr_in;
use crate::diag::Severity;
use crate::intrinsics::Intr;
use crate::model::*;
use crate::names;
use crate::vba::*;

/// A callback: either an arrow function literal (expanded in place) or a closure object.
enum Callback<'a> {
    Arrow(&'a Expr, &'a ast::ArrowFn),
    /// Closure object and the number of parameters it declares.
    Obj(VExpr, usize),
}

struct Stage<'a> {
    is_filter: bool,
    arg: &'a Expr,
    /// Element type after this stage.
    out_ty: Type,
}

enum Terminal<'a> {
    Collect { out: VExpr, counter: Option<String>, elem_ty: Type },
    ForEach(&'a Callback<'a>),
    Reduce { acc: VExpr, acc_ty: Type, cb: &'a Callback<'a> },
    Any { flag: VExpr, cb: &'a Callback<'a>, every: bool },
    Find { out: VExpr, ty: Type, cb: &'a Callback<'a>, index: bool },
}

fn simple_elem(v: &VExpr) -> bool {
    match v {
        VExpr::Lit(_) | VExpr::Name(_) => true,
        VExpr::Call(f, args) => matches!(**f, VExpr::Name(_)) && args.iter().all(|a| matches!(a, VExpr::Lit(_) | VExpr::Name(_))),
        _ => false,
    }
}

impl<'p> Lowerer<'p> {
    /// Splits `src.filter(f).map(g)` into its stages and the source array expression.
    fn collect_chain<'a>(&self, mut cur: &'a Expr) -> (Vec<Stage<'a>>, &'a Expr) {
        let info = self.info();
        let mut stages = Vec::new();
        loop {
            let c = cur.peel();
            let ExprKind::Call { callee, args, .. } = &c.kind else { break };
            let Some(CallTarget::Intrinsic(i @ (Intr::ArrMap | Intr::ArrFilter))) = info.calls.get(&c.id).map(|x| &x.target) else { break };
            let ExprKind::Member { object, .. } = &callee.peel().kind else { break };
            let Some(arg) = args.first() else { break };
            stages.push(Stage { is_filter: *i == Intr::ArrFilter, arg, out_ty: self.ty(c).elem().cloned().unwrap_or(Type::Any) });
            cur = object;
        }
        stages.reverse();
        (stages, cur)
    }

    fn prepare_cb<'a>(&mut self, arg: &'a Expr) -> Callback<'a> {
        match &arg.peel().kind {
            ExprKind::Arrow(a) => Callback::Arrow(arg.peel(), a),
            _ => {
                // A function-typed value: evaluate the closure once, call `Apply` per element.
                let arity = match self.ty(arg) {
                    Type::Func(f) => f.params.len(),
                    _ => usize::MAX,
                };
                let v = self.expr(arg);
                Callback::Obj(self.spill(v, &Type::Object), arity)
            }
        }
    }

    /// Expands one callback application. Returns its value (`Empty` when `want` is false).
    fn apply_cb(&mut self, cb: &Callback, vals: Vec<(VExpr, Type)>, want: bool) -> VExpr {
        match cb {
            Callback::Obj(o, arity) => {
                let target = o.clone().member("Apply");
                // A late-bound call with extra arguments fails at run time.
                let args: Vec<VExpr> = vals.into_iter().take(*arity).map(|v| v.0).collect();
                if want {
                    target.index(args)
                } else {
                    self.emit(VStmt::Call { target, args });
                    lit("Empty")
                }
            }
            Callback::Arrow(e, arrow) => {
                let info = self.info();
                let params = info.lambda_params.get(&e.id).cloned().unwrap_or_default();
                let fn_id = self.frame().fn_id;
                for (i, &l) in params.iter().enumerate() {
                    let Some((v, vty)) = vals.get(i) else { continue };
                    let reassigned = fn_id.and_then(|f| self.prog.fns[f].locals.get(l)).is_some_and(|d| d.reassigned);
                    if !reassigned && simple_elem(v) {
                        self.bind(l, Binding::Expr(v.clone()));
                    } else {
                        let pty = self.local_ty(l);
                        let n = self.local_name(&arrow.params[i].name.text, &pty);
                        self.store(&VExpr::Name(n.clone()), &pty, v.clone(), vty);
                        self.bind(l, Binding::Var(n));
                    }
                }
                match &arrow.body {
                    ArrowBody::Expr(b) => {
                        if want {
                            self.expr(b)
                        } else {
                            self.expr_stmt(b);
                            lit("Empty")
                        }
                    }
                    ArrowBody::Block(b) => {
                        let ret_ty = match self.ty(e) {
                            Type::Func(f) => f.ret,
                            _ => Type::Void,
                        };
                        let var = (ret_ty != Type::Void && want).then(|| self.declare("ras_r", &ret_ty));
                        let label = self.label("RAS_LambdaEnd");
                        let tries_base = self.pc.tries.len();
                        self.pc.rets.push(RetMode::Inline { var: var.clone(), ty: ret_ty, label: label.clone(), used: false, tries_base });
                        self.stmts(&b.stmts, true);
                        if let Some(RetMode::Inline { used: true, .. }) = self.pc.rets.pop() {
                            self.emit(VStmt::Label(label));
                        }
                        var.map(VExpr::Name).unwrap_or_else(|| lit("Empty"))
                    }
                }
            }
        }
    }

    fn run_stages(&mut self, stages: &[Stage], cbs: &[Callback], i: usize, cur: VExpr, cur_ty: Type, idx: &str, term: &Terminal) {
        let ix = || (name(idx), Type::Int);
        if i < stages.len() {
            let st = &stages[i];
            if st.is_filter {
                let c = self.apply_cb(&cbs[i], vec![(cur.clone(), cur_ty.clone()), ix()], true);
                let (body, _) = self.capture(|lw| lw.run_stages(stages, cbs, i + 1, cur, cur_ty, idx, term));
                self.emit(if_then(c, body));
            } else {
                let v = self.apply_cb(&cbs[i], vec![(cur, cur_ty), ix()], true);
                self.run_stages(stages, cbs, i + 1, v, st.out_ty.clone(), idx, term);
            }
            return;
        }
        match term {
            Terminal::Collect { out, counter, elem_ty } => match counter {
                Some(k) => {
                    self.store(&out.clone().index(vec![name(k.clone())]), elem_ty, cur, &cur_ty);
                    self.emit(assign(name(k.clone()), name(k.clone()).plus(1)));
                }
                None => self.store(&out.clone().index(vec![name(idx)]), elem_ty, cur, &cur_ty),
            },
            Terminal::ForEach(cb) => {
                self.apply_cb(cb, vec![(cur, cur_ty), ix()], false);
            }
            Terminal::Reduce { acc, acc_ty, cb } => {
                let v = self.apply_cb(cb, vec![(acc.clone(), acc_ty.clone()), (cur, cur_ty), ix()], true);
                self.store(acc, acc_ty, v, acc_ty);
            }
            Terminal::Any { flag, cb, every } => {
                let c = self.apply_cb(cb, vec![(cur, cur_ty), ix()], true);
                let (c, value) = if *every { (c.not(), "False") } else { (c, "True") };
                self.emit(if_then(c, vec![assign(flag.clone(), lit(value)), VStmt::Exit("For")]));
            }
            Terminal::Find { out, ty, cb, index } => {
                // The element may be an expression of an earlier `map`: evaluate it once.
                let cur = if *index { cur } else { self.spill(cur, &cur_ty) };
                let c = self.apply_cb(cb, vec![(cur.clone(), cur_ty.clone()), ix()], true);
                let (mut body, _) = self.capture(|lw| {
                    if *index {
                        lw.emit(assign(out.clone(), name(idx)));
                    } else {
                        lw.store(out, ty, cur, &cur_ty);
                    }
                });
                body.push(VStmt::Exit("For"));
                self.emit(if_then(c, body));
            }
        }
    }

    /// True when the callback / source expressions refer to the VBA variable `dest`.
    fn ast_mentions(&self, e: &Expr, dest: &str) -> bool {
        let info = self.info();
        any_expr_in(e, &mut |x| match (&x.kind, info.res.get(&x.id)) {
            (ExprKind::Name(_), Some(Res::Local(l))) => match self.frame().locals.get(l) {
                Some(Binding::Var(n)) => n == dest,
                Some(Binding::Expr(v)) => mentions(v, dest),
                None => false,
            },
            _ => false,
        })
    }

    /// Array-producing members (`map`, `filter`, `slice`, `concat`, `keys`, `values`) built into `dest`.
    pub(crate) fn array_intrinsic(&mut self, e: &Expr, intr: Intr, recv: &Expr, args: &[Expr], dest: &VExpr, ty: &Type) -> bool {
        let dest_name = dest.print();
        let elem_ty = ty.elem().cloned().unwrap_or(Type::Any);
        match intr {
            Intr::ArrMap | Intr::ArrFilter => {
                let (mut stages, source) = self.collect_chain(recv);
                let Some(arg) = args.first() else { return false };
                stages.push(Stage { is_filter: intr == Intr::ArrFilter, arg, out_ty: elem_ty.clone() });
                let src_ty = self.ty(source);
                let src = self.stable_array(source);
                let aliased = mentions(&src, &dest_name) || stages.iter().any(|s| self.ast_mentions(s.arg, &dest_name));
                let out = if aliased { VExpr::Name(self.temp(ty)) } else { dest.clone() };
                let cbs: Vec<Callback> = stages.iter().map(|s| self.prepare_cb(s.arg)).collect();
                let has_filter = stages.iter().any(|s| s.is_filter);
                let idx = self.declare_raw("ras_i", "Long", false);
                let counter = has_filter.then(|| self.declare_raw("ras_n", "Long", false));
                if let Some(k) = &counter {
                    self.emit(assign(name(k.clone()), long(0)));
                }
                let term = Terminal::Collect { out: out.clone(), counter: counter.clone(), elem_ty };
                let src_elem = src_ty.elem().cloned().unwrap_or(Type::Any);
                self.pc.loop_depth += 1;
                let (body, _) = self.capture(|lw| lw.run_stages(&stages, &cbs, 0, src.clone().index(vec![name(idx.clone())]), src_elem, &idx, &term));
                self.pc.loop_depth -= 1;
                let ub = call("UBound", vec![src.clone()]);
                let filled = vec![
                    VStmt::ReDim { preserve: false, target: out.clone(), dims: vec![(long(0), ub.clone())] },
                    VStmt::For { var: idx, from: long(0), to: ub, step: None, body },
                ];
                let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
                let non_empty = bin(">", call("RAS_Len", vec![src]), long(0));
                match counter {
                    None => self.emit(if_else(non_empty, filled, clear)),
                    Some(k) => {
                        self.emit(if_then(non_empty, filled));
                        // Trim to the number of elements that passed the filters.
                        self.emit(VStmt::If {
                            cond: bin("=", name(k.clone()), long(0)),
                            then: clear,
                            elifs: vec![(
                                bin("<=", name(k.clone()), call("UBound", vec![out.clone()])),
                                vec![VStmt::ReDim { preserve: true, target: out.clone(), dims: vec![(long(0), name(k).plus(-1))] }],
                            )],
                            els: vec![],
                        });
                    }
                }
                if aliased {
                    self.emit(assign(dest.clone(), out));
                }
                self.note(e.span, true, "fusion", format!("配列メソッドの連鎖（{} 段）を 1 本のループに融合しました", stages.len()));
                true
            }
            Intr::ArrSlice => {
                let src = self.stable_array(recv);
                let aliased = mentions(&src, &dest_name);
                let out = if aliased { VExpr::Name(self.temp(ty)) } else { dest.clone() };
                let n = VExpr::Name(self.temp(&Type::Int));
                self.emit(assign(n.clone(), call("RAS_Len", vec![src.clone()])));
                let start = VExpr::Name(self.temp(&Type::Int));
                let count = VExpr::Name(self.temp(&Type::Int));
                let a = match args.first() {
                    Some(x) => self.expr(x),
                    None => long(0),
                };
                self.emit(assign(start.clone(), call("RAS_ClampIndex", vec![a, n.clone()])));
                let end = match args.get(1) {
                    Some(x) => {
                        let b = self.expr(x);
                        call("RAS_ClampIndex", vec![b, n])
                    }
                    None => n,
                };
                self.emit(assign(count.clone(), bin("-", end, start.clone())));
                let idx = self.declare_raw("ras_i", "Long", false);
                let (copy, _) = self.capture(|lw| lw.store(&out.clone().index(vec![name(idx.clone())]), &elem_ty, src.index(vec![bin("+", start, name(idx.clone()))]), &elem_ty));
                let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
                self.emit(if_else(
                    bin(">", count.clone(), long(0)),
                    vec![
                        VStmt::ReDim { preserve: false, target: out.clone(), dims: vec![(long(0), count.clone().plus(-1))] },
                        VStmt::For { var: idx, from: long(0), to: count.plus(-1), step: None, body: copy },
                    ],
                    clear,
                ));
                if aliased {
                    self.emit(assign(dest.clone(), out));
                }
                true
            }
            Intr::ArrConcat => {
                let a = self.stable_array(recv);
                let Some(other) = args.first() else { return false };
                let b = self.stable_array(other);
                let aliased = mentions(&a, &dest_name) || mentions(&b, &dest_name);
                let out = if aliased { VExpr::Name(self.temp(ty)) } else { dest.clone() };
                let na = VExpr::Name(self.temp(&Type::Int));
                let nb = VExpr::Name(self.temp(&Type::Int));
                self.emit(assign(na.clone(), call("RAS_Len", vec![a.clone()])));
                self.emit(assign(nb.clone(), call("RAS_Len", vec![b.clone()])));
                let idx = self.declare_raw("ras_i", "Long", false);
                let (copy_a, _) = self.capture(|lw| lw.store(&out.clone().index(vec![name(idx.clone())]), &elem_ty, a.index(vec![name(idx.clone())]), &elem_ty));
                let (copy_b, _) = self.capture(|lw| lw.store(&out.clone().index(vec![bin("+", na.clone(), name(idx.clone()))]), &elem_ty, b.index(vec![name(idx.clone())]), &elem_ty));
                let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
                let total = bin("+", na.clone(), nb.clone());
                self.emit(if_else(
                    bin(">", total.clone(), long(0)),
                    vec![
                        VStmt::ReDim { preserve: false, target: out.clone(), dims: vec![(long(0), total.plus(-1))] },
                        VStmt::For { var: idx.clone(), from: long(0), to: na.plus(-1), step: None, body: copy_a },
                        VStmt::For { var: idx, from: long(0), to: nb.plus(-1), step: None, body: copy_b },
                    ],
                    clear,
                ));
                if aliased {
                    self.emit(assign(dest.clone(), out));
                }
                true
            }
            Intr::MapKeys | Intr::MapValues => {
                let d = self.expr(recv);
                let d = self.spill(d, &Type::Object);
                let k = self.declare_raw("ras_n", "Long", false);
                let item = self.declare_raw("ras_v", "Variant", false);
                let (mut body, _) = self.capture(|lw| lw.store(&dest.clone().index(vec![name(k.clone())]), &elem_ty, name(item.clone()), &Type::Any));
                body.push(assign(name(k.clone()), name(k.clone()).plus(1)));
                let (clear, _) = self.capture(|lw| lw.clear_array(dest, ty));
                let coll = d.clone().member(if intr == Intr::MapKeys { "Keys" } else { "Items" });
                self.emit(if_else(
                    bin(">", d.clone().member("Count"), long(0)),
                    vec![
                        VStmt::ReDim { preserve: false, target: dest.clone(), dims: vec![(long(0), d.member("Count").plus(-1))] },
                        assign(name(k), long(0)),
                        VStmt::ForEach { var: item, coll, body },
                    ],
                    clear,
                ));
                true
            }
            _ => false,
        }
    }

    /// `forEach` / `reduce` / `some` / `every` / `find` / `findIndex`, fused with preceding `map` / `filter` stages.
    pub(crate) fn scalar_fusion(&mut self, e: &Expr, intr: Intr, recv: &Expr, args: &[Expr], want: bool) -> Option<VExpr> {
        let (stages, source) = self.collect_chain(recv);
        let src_ty = self.ty(source);
        let src = self.stable_array(source);
        let cbs: Vec<Callback> = stages.iter().map(|s| self.prepare_cb(s.arg)).collect();
        let cb = self.prepare_cb(args.first()?);
        let result_ty = self.ty(e);
        let (term, result) = match intr {
            Intr::ArrForEach => (Terminal::ForEach(&cb), None),
            Intr::ArrReduce => {
                let acc = VExpr::Name(self.temp(&result_ty));
                self.expr_into(args.get(1)?, &acc, &result_ty);
                (Terminal::Reduce { acc: acc.clone(), acc_ty: result_ty.clone(), cb: &cb }, Some(acc))
            }
            Intr::ArrSome | Intr::ArrEvery => {
                let flag = VExpr::Name(self.temp(&Type::Bool));
                let every = intr == Intr::ArrEvery;
                self.emit(assign(flag.clone(), lit(if every { "True" } else { "False" })));
                (Terminal::Any { flag: flag.clone(), cb: &cb, every }, Some(flag))
            }
            Intr::ArrFind => {
                let out = VExpr::Name(self.temp(&result_ty));
                self.reset_var(&out, &result_ty);
                (Terminal::Find { out: out.clone(), ty: result_ty.clone(), cb: &cb, index: false }, Some(out))
            }
            Intr::ArrFindIndex => {
                let out = VExpr::Name(self.temp(&Type::Int));
                self.emit(assign(out.clone(), lit("-1&")));
                (Terminal::Find { out: out.clone(), ty: Type::Int, cb: &cb, index: true }, Some(out))
            }
            _ => return None,
        };
        let idx = self.declare_raw("ras_i", "Long", false);
        let src_elem = src_ty.elem().cloned().unwrap_or(Type::Any);
        self.pc.loop_depth += 1;
        let (body, _) = self.capture(|lw| lw.run_stages(&stages, &cbs, 0, src.clone().index(vec![name(idx.clone())]), src_elem, &idx, &term));
        self.pc.loop_depth -= 1;
        self.emit(VStmt::For { var: idx, from: long(0), to: call("RAS_Len", vec![src]).plus(-1), step: None, body });
        self.note(e.span, true, "fusion", "コールバックをループに展開しました");
        if want {
            Some(result.unwrap_or_else(|| lit("Empty")))
        } else {
            None
        }
    }

    /// A lambda that escapes becomes a class module with an `Apply` method (spec §4.3).
    pub(crate) fn closure_value(&mut self, e: &Expr, arrow: &ast::ArrowFn) -> VExpr {
        let info = self.info();
        self.lambda_n += 1;
        let cname = format!("RAS_Lambda_{:03}", self.lambda_n);
        let captures = info.lambda_captures.get(&e.id).cloned().unwrap_or_default();
        let params = info.lambda_params.get(&e.id).cloned().unwrap_or_default();
        let ret_ty = match self.ty(e) {
            Type::Func(f) => f.ret,
            _ => Type::Any,
        };
        let fn_id = self.frame().fn_id;
        let mut caps: Vec<(LocalId, String, Type, VExpr)> = Vec::new();
        for &l in &captures {
            // Locals of the lambda itself are not captures.
            if params.contains(&l) || !self.frame().locals.contains_key(&l) {
                continue;
            }
            let lname = fn_id.and_then(|f| self.prog.fns[f].locals.get(l)).map(|d| d.name.clone()).unwrap_or_else(|| format!("v{l}"));
            let lname = if lname == "this" { "this_".to_string() } else { lname };
            let mut field = format!("Captured_{}", names::valid_ident(&lname));
            while caps.iter().any(|c| c.1.eq_ignore_ascii_case(&field)) {
                field.push('_');
            }
            caps.push((l, field, self.local_ty(l), self.local_expr(l)));
        }
        // Constants of a script's `Main` that became locals are copied into the closure as well.
        let mut demoted_caps: Vec<(VarId, String, Type, VExpr)> = Vec::new();
        {
            let mut visit = |x: &Expr| {
                if let Some(Res::Global(v)) = info.res.get(&x.id) {
                    if let Some(n) = self.pc.demoted.get(v) {
                        if !demoted_caps.iter().any(|c| c.0 == *v) {
                            let field = format!("Captured_{}", names::valid_ident(&self.prog.vars[*v].name));
                            demoted_caps.push((*v, field, self.prog.vars[*v].ty.clone(), VExpr::Name(n.clone())));
                        }
                    }
                }
                false
            };
            match &arrow.body {
                ArrowBody::Expr(b) => {
                    any_expr_in(b, &mut visit);
                }
                ArrowBody::Block(b) => {
                    for st in &b.stmts {
                        crate::astutil::any_expr(st, &mut visit);
                    }
                }
            }
        }
        let outer_this = self.frame().this.clone();
        let body_uses_this = outer_this.is_some()
            && match &arrow.body {
                ArrowBody::Expr(b) => any_expr_in(b, &mut |x| matches!(x.kind, ExprKind::This)),
                ArrowBody::Block(b) => b.stmts.iter().any(|s| crate::astutil::any_expr(s, &mut |x| matches!(x.kind, ExprKind::This))),
            };

        let saved = std::mem::take(&mut self.pc);
        let outer = saved.frames.last().unwrap();
        self.pc.emit_module = saved.emit_module;
        self.pc.in_closure = true;
        self.pc.exit_kw = "Function";
        self.pc.proc_name = "Apply".into();
        self.pc.ret_ty = ret_ty.clone();
        self.pc.display_name = saved.display_name.clone();
        self.pc.blocks.push(Vec::new());
        self.pc.names.insert("apply".into());
        let mut frame = Frame { module: outer.module, fn_id: outer.fn_id, env: outer.env.clone(), locals: HashMap::new(), this: body_uses_this.then(|| lit("Me").member("Captured_Me")), inl: None };
        for (l, field, _, _) in &caps {
            self.pc.names.insert(field.to_lowercase());
            frame.locals.insert(*l, Binding::Expr(lit("Me").member(field.clone())));
        }
        for (v, field, _, _) in &demoted_caps {
            self.pc.names.insert(field.to_lowercase());
            self.pc.demoted.insert(*v, format!("Me.{field}"));
        }
        let mut vparams = Vec::new();
        for (i, &l) in params.iter().enumerate() {
            let n = self.unique(&names::local_ident(&arrow.params[i].name.text));
            frame.locals.insert(l, Binding::Var(n.clone()));
            vparams.push(VParam { name: n, byref: false, ty: "Variant".into(), is_array: false, optional_default: None });
        }
        self.pc.frames.push(frame);
        self.pc.rets.push(RetMode::Proc);
        match &arrow.body {
            ArrowBody::Expr(b) => {
                if ret_ty == Type::Void {
                    self.expr_stmt(b);
                } else {
                    let bt = self.ty(b);
                    let v = self.expr(b);
                    self.store(&name("Apply"), &ret_ty, v, &bt);
                }
            }
            ArrowBody::Block(b) => {
                if needs_handler(&b.stmts) {
                    self.err(arrow.span, "E0640", "変数に代入される・引数として渡されるアロー関数の中では `try` / `using` を使えません。通常の関数に切り出してください");
                }
                self.stmts(&b.stmts, true);
            }
        }
        let body = self.pc.blocks.pop().unwrap();
        let locals = std::mem::take(&mut self.pc.locals);
        self.pc = saved;

        let ci = self.vmodule(&cname, VModuleKind::Class);
        let src_path = self.prog.modules[self.src_module()].rel_path.clone();
        self.out[ci].sources.push(src_path);
        for (_, field, ty, _) in &caps {
            let (vt, arr) = self.vtype(ty);
            // Public array fields are not allowed in class modules.
            let vt = if arr { "Variant".to_string() } else { vt };
            self.out[ci].decls.push(format!("Public {field} As {vt}"));
        }
        for (_, field, ty, _) in &demoted_caps {
            let (vt, arr) = self.vtype(ty);
            let vt = if arr { "Variant".to_string() } else { vt };
            self.out[ci].decls.push(format!("Public {field} As {vt}"));
        }
        if body_uses_this {
            self.out[ci].decls.push("Public Captured_Me As Object".into());
        }
        self.out[ci].procs.push(VProc { kind: ProcKind::Function, public: true, name: "Apply".into(), params: vparams, ret: Some("Variant".into()), locals, body, leading_comments: vec![], attributes: vec![], source: Some(0) });

        let t = VExpr::Name(self.temp(&Type::Object));
        self.emit(set(t.clone(), lit(format!("New {cname}"))));
        for (_, field, ty, outer_value) in caps {
            self.store(&t.clone().member(field), &ty, outer_value, &ty);
        }
        for (_, field, ty, outer_value) in demoted_caps {
            self.store(&t.clone().member(field), &ty, outer_value, &ty);
        }
        if body_uses_this {
            self.emit(set(t.clone().member("Captured_Me"), outer_this.unwrap()));
        }
        // spec §4.1: tell the user that this lambda is slow.
        self.diag(Severity::Info, self.src_module(), arrow.span, "I0641", format!("このアロー関数はクラス（{cname}）に変換されるため、呼び出しが遅くなります。配列メソッドの引数に直接書くとループに展開されます"));
        t
    }
}
