//! Lambdas: loop fusion for array methods (spec §4.1 rule 1) and closure
//! classes for lambdas that escape (rule 2).

use std::collections::HashMap;

use ras_syntax::ast::{self, ArrowBody, Expr, ExprKind};

use super::expr::{mentions, simple};
use super::stmt::needs_handler;
use super::{Binding, Frame, Lowerer, RetMode};
use crate::astutil::any_expr_in;
use crate::diag::Severity;
use crate::intrinsics::Intr;
use crate::model::*;
use crate::names;
use crate::vba::*;

/// How an end of a copied range is given (spec §4.5).
pub(crate) enum RangeEnd {
    /// `slice(i)`: a negative value counts from the end.
    Js(VExpr),
    /// `xs[i..]`: counted from the start, kept within the array.
    Abs(VExpr),
    /// `xs[^k..]`.
    FromEnd(VExpr),
    /// Left out at the end: the length.
    Len,
}

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
    Collect {
        out: VExpr,
        counter: Option<String>,
        elem_ty: Type,
    },
    ForEach(&'a Callback<'a>),
    Reduce {
        acc: VExpr,
        acc_ty: Type,
        cb: &'a Callback<'a>,
    },
    Any {
        flag: VExpr,
        cb: &'a Callback<'a>,
        every: bool,
    },
    Find {
        out: VExpr,
        ty: Type,
        cb: &'a Callback<'a>,
        index: bool,
    },
}

/// Which counter each callback of a fused chain receives as its index. As in JavaScript, that is
/// the position in the array the method is called on: after a `filter`, positions count the
/// elements that passed it, not the elements of the source.
struct IndexPlan {
    /// Something reads the position in the source.
    source: bool,
    /// Per stage, for a `filter` whose result is indexed by a later callback: the counter of the
    /// elements that passed, and whether the filter itself advances it (`Collect` advances its own).
    passed: Vec<Option<(String, bool)>>,
}

impl Callback<'_> {
    /// Does the callback take the parameter at `pos` (the element index)?
    fn takes(&self, pos: usize) -> bool {
        match self {
            Callback::Arrow(_, a) => a.params.len() > pos,
            Callback::Obj(_, arity) => *arity > pos,
        }
    }
}

enum Source {
    Array {
        src: VExpr,
        ty: Type,
    },
    /// Host collection, the `For Each` control variable and the type of its members.
    Coll {
        coll: VExpr,
        item: String,
        elem_ty: Type,
    },
}

impl Source {
    fn mentions(&self, name: &str) -> bool {
        match self {
            Source::Array { src, .. } => mentions(src, name),
            Source::Coll { coll, .. } => mentions(coll, name),
        }
    }

    /// Index of the last element (`UBound(a)` / `coll.Count - 1`).
    fn upper(&self) -> VExpr {
        match self {
            Source::Array { src, .. } => call("UBound", vec![src.clone()]),
            Source::Coll { coll, .. } => coll.clone().member("Count").plus(-1),
        }
    }
}

fn simple_elem(v: &VExpr) -> bool {
    match v {
        VExpr::Lit(_) | VExpr::Name(_) => true,
        VExpr::Call(f, args) => {
            matches!(**f, VExpr::Name(_))
                && args
                    .iter()
                    .all(|a| matches!(a, VExpr::Lit(_) | VExpr::Name(_)))
        }
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
            let ExprKind::Call { callee, args, .. } = &c.kind else {
                break;
            };
            let Some(CallTarget::Intrinsic(i @ (Intr::ArrMap | Intr::ArrFilter))) =
                info.calls.get(&c.id).map(|x| &x.target)
            else {
                break;
            };
            let ExprKind::Member { object, .. } = &callee.peel().kind else {
                break;
            };
            let Some(arg) = args.first() else { break };
            stages.push(Stage {
                is_filter: *i == Intr::ArrFilter,
                arg,
                out_ty: self.ty(c).elem().cloned().unwrap_or(Type::Any),
            });
            cur = object;
        }
        stages.reverse();
        (stages, cur)
    }

    /// What a fused loop runs over: an array (`For i = 0 To UBound(a)`) or a host collection
    /// such as `worksheets` (`For Each`, spec §4.4).
    fn fusion_source(&mut self, source: &Expr) -> Source {
        let ty = self.ty(source);
        match self.prog.collection_elem(&ty) {
            Some(elem_ty) if !ty.is_array() => {
                let c = self.expr(source);
                let coll = self.spill_collection(c, &ty);
                // The control variable of `For Each` must be a Variant or an object.
                let item = if elem_ty.is_object() {
                    self.declare("ras_e", &elem_ty)
                } else {
                    self.declare_raw("ras_v", "Variant", false)
                };
                Source::Coll {
                    coll,
                    item,
                    elem_ty,
                }
            }
            _ => Source::Array {
                src: self.stable_array(source),
                ty,
            },
        }
    }

    /// A collection expression that is evaluated several times (`.Count`, `For Each`): a global
    /// such as `Worksheets` is used as it is, anything else is evaluated once.
    fn spill_collection(&mut self, c: VExpr, ty: &Type) -> VExpr {
        if matches!(c, VExpr::Name(_)) {
            c
        } else {
            self.spill(c, ty)
        }
    }

    /// Decides the index counters of a chain. `term_reads` tells whether the terminal reads the index
    /// of the array it runs on; `collect` is the output counter of a filtering `map` / `filter` chain,
    /// which is at the same time the position after the last `filter`.
    fn plan_indexes(
        &mut self,
        stages: &[Stage],
        cbs: &[Callback],
        term_reads: bool,
        collect: Option<&str>,
    ) -> IndexPlan {
        // Segment 0 is the source; every `filter` starts a new segment.
        let mut reads = vec![false];
        for (st, cb) in stages.iter().zip(cbs) {
            *reads.last_mut().unwrap() |= cb.takes(1);
            if st.is_filter {
                reads.push(false);
            }
        }
        *reads.last_mut().unwrap() |= term_reads;
        let last_filter = stages.iter().rposition(|s| s.is_filter);
        let mut passed = Vec::with_capacity(stages.len());
        let mut segment = 0;
        for (i, st) in stages.iter().enumerate() {
            if !st.is_filter {
                passed.push(None);
                continue;
            }
            segment += 1;
            passed.push(match collect {
                Some(k) if last_filter == Some(i) => Some((k.to_string(), false)),
                _ if reads[segment] => {
                    let k = self.declare_raw("ras_k", "Long", false);
                    self.emit(assign(name(k.clone()), long(0)));
                    Some((k, true))
                }
                _ => None,
            });
        }
        IndexPlan {
            source: reads[0],
            passed,
        }
    }

    /// The loop counter. A `For Each` loop only counts when something reads the index.
    fn fusion_index(&mut self, source: &Source, needed: bool) -> String {
        if needed || matches!(source, Source::Array { .. }) {
            self.declare_raw("ras_i", "Long", false)
        } else {
            String::new()
        }
    }

    /// Emits the loop over `source`: `body` receives the current element and its type. `idx` counts from 0
    /// (empty: a `For Each` loop that does not count).
    fn fused_loop(
        &mut self,
        source: &Source,
        idx: &str,
        body: impl FnOnce(&mut Self, VExpr, Type),
    ) {
        self.pc.loop_depth += 1;
        match source {
            Source::Array { src, ty } => {
                let elem_ty = ty.elem().cloned().unwrap_or(Type::Any);
                let (b, _) =
                    self.capture(|lw| body(lw, src.clone().index(vec![name(idx)]), elem_ty));
                self.pc.loop_depth -= 1;
                let len = self.len_of(src.clone(), ty);
                self.emit(VStmt::For {
                    var: idx.to_string(),
                    from: long(0),
                    to: len.plus(-1),
                    step: None,
                    body: b,
                });
            }
            Source::Coll {
                coll,
                item,
                elem_ty,
            } => {
                let (mut b, _) = self.capture(|lw| {
                    if elem_ty.is_object() || elem_ty.is_any() {
                        body(lw, name(item.clone()), elem_ty.clone());
                    } else {
                        // A value element (`FileDialogSelectedItems`) arrives in a Variant.
                        let typed = VExpr::Name(lw.temp(elem_ty));
                        lw.store(&typed, elem_ty, name(item.clone()), &Type::Any);
                        body(lw, typed, elem_ty.clone());
                    }
                });
                self.pc.loop_depth -= 1;
                if !idx.is_empty() {
                    b.push(assign(name(idx), name(idx).plus(1)));
                    self.emit(assign(name(idx), long(0)));
                }
                self.emit(VStmt::ForEach {
                    var: item.clone(),
                    coll: coll.clone(),
                    body: b,
                });
            }
        }
    }

    /// Copies the members of the host collection `coll` into the array `dest`, in `For Each` order.
    fn collection_into(&mut self, coll: &Expr, dest: &VExpr, arr_ty: &Type) {
        let elem_ty = arr_ty.elem().cloned().unwrap_or(Type::Any);
        let coll_ty = self.ty(coll);
        let c = self.expr(coll);
        let c = self.spill_collection(c, &coll_ty);
        let k = self.declare_raw("ras_n", "Long", false);
        // The control variable of `For Each` must be a Variant or an object.
        let (item, item_ty) = if elem_ty.is_object() {
            (self.declare("ras_e", &elem_ty), elem_ty.clone())
        } else {
            (self.declare_raw("ras_v", "Variant", false), Type::Any)
        };
        let (mut body, _) = self.capture(|lw| {
            lw.store(
                &dest.clone().index(vec![name(k.clone())]),
                &elem_ty,
                name(item.clone()),
                &item_ty,
            )
        });
        body.push(assign(name(k.clone()), name(k.clone()).plus(1)));
        let (clear, _) = self.capture(|lw| lw.clear_array(dest, arr_ty));
        let count = c.clone().member("Count");
        self.emit(if_else(
            bin(">", count.clone(), long(0)),
            vec![
                VStmt::ReDim {
                    preserve: false,
                    target: dest.clone(),
                    dims: vec![(long(0), count.plus(-1))],
                },
                assign(name(k), long(0)),
                VStmt::ForEach {
                    var: item,
                    coll: c,
                    body,
                },
            ],
            clear,
        ));
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
                    let Some((v, vty)) = vals.get(i) else {
                        continue;
                    };
                    let reassigned = fn_id
                        .and_then(|f| self.prog.fns[f].locals.get(l))
                        .is_some_and(|d| d.reassigned);
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
                        let var =
                            (ret_ty != Type::Void && want).then(|| self.declare("ras_r", &ret_ty));
                        let label = self.label("RAS_LambdaEnd");
                        let tries_base = self.pc.tries.len();
                        self.pc.rets.push(RetMode::Inline {
                            var: var.clone(),
                            ty: ret_ty,
                            label: label.clone(),
                            used: false,
                            tries_base,
                        });
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

    /// Runs stage `i` and what follows on the element `cur`. `idx` is the position of `cur` in the
    /// array stage `i` works on (the source, or the result of the preceding `filter`).
    #[allow(clippy::too_many_arguments)]
    fn run_stages(
        &mut self,
        stages: &[Stage],
        cbs: &[Callback],
        i: usize,
        cur: VExpr,
        cur_ty: Type,
        idx: &str,
        plan: &IndexPlan,
        term: &Terminal,
    ) {
        let ix = || (name(idx), Type::Int);
        if i < stages.len() {
            let st = &stages[i];
            if st.is_filter {
                let c = self.apply_cb(&cbs[i], vec![(cur.clone(), cur_ty.clone()), ix()], true);
                // What follows sees the position among the elements that passed.
                let passed = plan.passed[i].as_ref();
                let inner = passed.map_or(idx, |p| p.0.as_str());
                let (mut body, _) = self.capture(|lw| {
                    lw.run_stages(stages, cbs, i + 1, cur, cur_ty, inner, plan, term)
                });
                if let Some((k, true)) = passed {
                    body.push(assign(name(k.clone()), name(k.clone()).plus(1)));
                }
                self.emit(if_then(c, body));
            } else {
                let v = self.apply_cb(&cbs[i], vec![(cur, cur_ty), ix()], true);
                self.run_stages(stages, cbs, i + 1, v, st.out_ty.clone(), idx, plan, term);
            }
            return;
        }
        match term {
            Terminal::Collect {
                out,
                counter,
                elem_ty,
            } => match counter {
                Some(k) => {
                    self.store(
                        &out.clone().index(vec![name(k.clone())]),
                        elem_ty,
                        cur,
                        &cur_ty,
                    );
                    self.emit(assign(name(k.clone()), name(k.clone()).plus(1)));
                }
                None => self.store(&out.clone().index(vec![name(idx)]), elem_ty, cur, &cur_ty),
            },
            Terminal::ForEach(cb) => {
                self.apply_cb(cb, vec![(cur, cur_ty), ix()], false);
            }
            Terminal::Reduce { acc, acc_ty, cb } => {
                let v = self.apply_cb(
                    cb,
                    vec![(acc.clone(), acc_ty.clone()), (cur, cur_ty), ix()],
                    true,
                );
                self.store(acc, acc_ty, v, acc_ty);
            }
            Terminal::Any { flag, cb, every } => {
                let c = self.apply_cb(cb, vec![(cur, cur_ty), ix()], true);
                let (c, value) = if *every {
                    (c.not(), "False")
                } else {
                    (c, "True")
                };
                self.emit(if_then(
                    c,
                    vec![assign(flag.clone(), lit(value)), VStmt::Exit("For")],
                ));
            }
            Terminal::Find { out, ty, cb, index } => {
                // The element may be an expression of an earlier `map`: evaluate it once.
                let cur = if *index {
                    cur
                } else {
                    self.spill(cur, &cur_ty)
                };
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
    pub(crate) fn ast_mentions(&self, e: &Expr, dest: &str) -> bool {
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
    pub(crate) fn array_intrinsic(
        &mut self,
        e: &Expr,
        intr: Intr,
        recv: &Expr,
        args: &[Expr],
        dest: &VExpr,
        ty: &Type,
    ) -> bool {
        let dest_name = dest.print();
        let elem_ty = ty.elem().cloned().unwrap_or(Type::Any);
        match intr {
            Intr::ArrMap | Intr::ArrFilter => {
                let (mut stages, source) = self.collect_chain(recv);
                let Some(arg) = args.first() else {
                    return false;
                };
                stages.push(Stage {
                    is_filter: intr == Intr::ArrFilter,
                    arg,
                    out_ty: elem_ty.clone(),
                });
                let src = self.fusion_source(source);
                let aliased = src.mentions(&dest_name)
                    || stages.iter().any(|s| self.ast_mentions(s.arg, &dest_name));
                let out = if aliased {
                    VExpr::Name(self.temp(ty))
                } else {
                    dest.clone()
                };
                let cbs: Vec<Callback> = stages.iter().map(|s| self.prepare_cb(s.arg)).collect();
                let has_filter = stages.iter().any(|s| s.is_filter);
                let counter = has_filter.then(|| self.declare_raw("ras_n", "Long", false));
                if let Some(k) = &counter {
                    self.emit(assign(name(k.clone()), long(0)));
                }
                // Without a filter the output is written at the position of the source.
                let plan = self.plan_indexes(&stages, &cbs, !has_filter, counter.as_deref());
                let idx = self.fusion_index(&src, plan.source);
                let term = Terminal::Collect {
                    out: out.clone(),
                    counter: counter.clone(),
                    elem_ty,
                };
                let ub = src.upper();
                let (filled, _) = self.capture(|lw| {
                    lw.emit(VStmt::ReDim {
                        preserve: false,
                        target: out.clone(),
                        dims: vec![(long(0), ub.clone())],
                    });
                    match &src {
                        // The length is known to be positive here: `UBound` needs no guard.
                        Source::Array { src: a, ty: aty } => {
                            let elem = aty.elem().cloned().unwrap_or(Type::Any);
                            lw.pc.loop_depth += 1;
                            let (body, _) = lw.capture(|lw| {
                                lw.run_stages(
                                    &stages,
                                    &cbs,
                                    0,
                                    a.clone().index(vec![name(idx.clone())]),
                                    elem,
                                    &idx,
                                    &plan,
                                    &term,
                                )
                            });
                            lw.pc.loop_depth -= 1;
                            lw.emit(VStmt::For {
                                var: idx.clone(),
                                from: long(0),
                                to: ub.clone(),
                                step: None,
                                body,
                            });
                        }
                        Source::Coll { .. } => lw.fused_loop(&src, &idx, |lw, cur, cur_ty| {
                            lw.run_stages(&stages, &cbs, 0, cur, cur_ty, &idx, &plan, &term)
                        }),
                    }
                });
                let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
                let non_empty = match &src {
                    Source::Array { src: a, ty: aty } => {
                        bin(">", self.len_of(a.clone(), aty), long(0))
                    }
                    Source::Coll { coll, .. } => bin(">", coll.clone().member("Count"), long(0)),
                };
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
                                vec![VStmt::ReDim {
                                    preserve: true,
                                    target: out.clone(),
                                    dims: vec![(long(0), name(k).plus(-1))],
                                }],
                            )],
                            els: vec![],
                        });
                    }
                }
                if aliased {
                    self.emit(assign(dest.clone(), out));
                }
                self.note(
                    e.span,
                    true,
                    "fusion",
                    format!(
                        "配列メソッドの連鎖（{} 段）を 1 本のループに融合しました",
                        stages.len()
                    ),
                );
                true
            }
            Intr::CollToArray => {
                self.collection_into(recv, dest, ty);
                true
            }
            Intr::ArrSlice => {
                let src = self.stable_array(recv);
                let start = match args.first() {
                    Some(x) => RangeEnd::Js(self.expr(x)),
                    None => RangeEnd::Abs(long(0)),
                };
                let end = match args.get(1) {
                    Some(x) => RangeEnd::Js(self.expr(x)),
                    None => RangeEnd::Len,
                };
                self.copy_range(src, ty, start, end, dest);
                true
            }
            Intr::ArrStep => {
                let src = self.stable_array(recv);
                let Some(k) = args.first() else {
                    return false;
                };
                let k = self.expr(k);
                let k = self.spill(k, &Type::Int);
                let len = self.len_of(src.clone(), ty);
                let count = call("RAS_StepCount", vec![len, k.clone()]);
                self.copy_mapped(src, ty, dest, count, move |i| bin("*", i, k.clone()));
                true
            }
            Intr::ArrToReversed => {
                let src = self.stable_array(recv);
                let n = VExpr::Name(self.temp(&Type::Int));
                let len = self.len_of(src.clone(), ty);
                self.emit(assign(n.clone(), len));
                let last = n.clone().plus(-1);
                self.copy_mapped(src, ty, dest, n, move |i| bin("-", last.clone(), i));
                true
            }
            Intr::ArrConcat => {
                let a = self.stable_array(recv);
                let Some(other) = args.first() else {
                    return false;
                };
                let b = self.stable_array(other);
                let aliased = mentions(&a, &dest_name) || mentions(&b, &dest_name);
                let out = if aliased {
                    VExpr::Name(self.temp(ty))
                } else {
                    dest.clone()
                };
                let na = VExpr::Name(self.temp(&Type::Int));
                let nb = VExpr::Name(self.temp(&Type::Int));
                let (la, lb) = (self.len_of(a.clone(), ty), self.len_of(b.clone(), ty));
                self.emit(assign(na.clone(), la));
                self.emit(assign(nb.clone(), lb));
                let idx = self.declare_raw("ras_i", "Long", false);
                let (copy_a, _) = self.capture(|lw| {
                    lw.store(
                        &out.clone().index(vec![name(idx.clone())]),
                        &elem_ty,
                        a.index(vec![name(idx.clone())]),
                        &elem_ty,
                    )
                });
                let (copy_b, _) = self.capture(|lw| {
                    lw.store(
                        &out.clone()
                            .index(vec![bin("+", na.clone(), name(idx.clone()))]),
                        &elem_ty,
                        b.index(vec![name(idx.clone())]),
                        &elem_ty,
                    )
                });
                let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
                let total = bin("+", na.clone(), nb.clone());
                self.emit(if_else(
                    bin(">", total.clone(), long(0)),
                    vec![
                        VStmt::ReDim {
                            preserve: false,
                            target: out.clone(),
                            dims: vec![(long(0), total.plus(-1))],
                        },
                        VStmt::For {
                            var: idx.clone(),
                            from: long(0),
                            to: na.plus(-1),
                            step: None,
                            body: copy_a,
                        },
                        VStmt::For {
                            var: idx,
                            from: long(0),
                            to: nb.plus(-1),
                            step: None,
                            body: copy_b,
                        },
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
                let (mut body, _) = self.capture(|lw| {
                    lw.store(
                        &dest.clone().index(vec![name(k.clone())]),
                        &elem_ty,
                        name(item.clone()),
                        &Type::Any,
                    )
                });
                body.push(assign(name(k.clone()), name(k.clone()).plus(1)));
                let (clear, _) = self.capture(|lw| lw.clear_array(dest, ty));
                let coll = d.clone().member(if intr == Intr::MapKeys {
                    "Keys"
                } else {
                    "Items"
                });
                self.emit(if_else(
                    bin(">", d.clone().member("Count"), long(0)),
                    vec![
                        VStmt::ReDim {
                            preserve: false,
                            target: dest.clone(),
                            dims: vec![(long(0), d.member("Count").plus(-1))],
                        },
                        assign(name(k), long(0)),
                        VStmt::ForEach {
                            var: item,
                            coll,
                            body,
                        },
                    ],
                    clear,
                ));
                true
            }
            _ => false,
        }
    }

    /// `forEach` / `reduce` / `some` / `every` / `find` / `findIndex`, fused with preceding `map` / `filter` stages.
    /// Copies `src[start..end]` (spec §4.5) into `dest`.
    pub(crate) fn copy_range(
        &mut self,
        src: VExpr,
        ty: &Type,
        start: RangeEnd,
        end: RangeEnd,
        dest: &VExpr,
    ) {
        let n = VExpr::Name(self.temp(&Type::Int));
        let len = self.len_of(src.clone(), ty);
        self.emit(assign(n.clone(), len));
        let bound = |e: RangeEnd, n: &VExpr| match e {
            RangeEnd::Js(v) => call("RAS_ClampIndex", vec![v, n.clone()]),
            RangeEnd::Abs(v) => match v.as_long() {
                Some(0) => long(0),
                _ => call("RAS_ClampRange", vec![v, n.clone()]),
            },
            RangeEnd::FromEnd(k) => match k.as_long() {
                Some(0) => n.clone(),
                _ => call("RAS_ClampRange", vec![bin("-", n.clone(), k), n.clone()]),
            },
            RangeEnd::Len => n.clone(),
        };
        let first = bound(start, &n);
        let first = if simple(&first) {
            first
        } else {
            let t = VExpr::Name(self.temp(&Type::Int));
            self.emit(assign(t.clone(), first));
            t
        };
        let count = VExpr::Name(self.temp(&Type::Int));
        let last = bound(end, &n);
        let count_v = match first.as_long() {
            Some(0) => last,
            _ => bin("-", last, first.clone()),
        };
        self.emit(assign(count.clone(), count_v));
        let offset = first;
        self.copy_mapped(src, ty, dest, count, move |i| match offset.as_long() {
            Some(0) => i,
            _ => bin("+", offset.clone(), i),
        });
    }

    /// `dest(i) = src(at(i))` for `i` in `0..count` (`dest` is left empty when `count <= 0`).
    pub(crate) fn copy_mapped(
        &mut self,
        src: VExpr,
        ty: &Type,
        dest: &VExpr,
        count: VExpr,
        at: impl Fn(VExpr) -> VExpr,
    ) {
        let elem_ty = ty.elem().cloned().unwrap_or(Type::Any);
        let aliased = mentions(&src, &dest.print());
        let out = if aliased {
            VExpr::Name(self.temp(ty))
        } else {
            dest.clone()
        };
        let count = if simple(&count) {
            count
        } else {
            let t = VExpr::Name(self.temp(&Type::Int));
            self.emit(assign(t.clone(), count));
            t
        };
        let idx = self.declare_raw("ras_i", "Long", false);
        let (copy, _) = self.capture(|lw| {
            lw.store(
                &out.clone().index(vec![name(idx.clone())]),
                &elem_ty,
                src.index(vec![at(name(idx.clone()))]),
                &elem_ty,
            )
        });
        let (clear, _) = self.capture(|lw| lw.clear_array(&out, ty));
        self.emit(if_else(
            bin(">", count.clone(), long(0)),
            vec![
                VStmt::ReDim {
                    preserve: false,
                    target: out.clone(),
                    dims: vec![(long(0), count.clone().plus(-1))],
                },
                VStmt::For {
                    var: idx,
                    from: long(0),
                    to: count.plus(-1),
                    step: None,
                    body: copy,
                },
            ],
            clear,
        ));
        if aliased {
            self.emit(assign(dest.clone(), out));
        }
    }

    pub(crate) fn scalar_fusion(
        &mut self,
        e: &Expr,
        intr: Intr,
        recv: &Expr,
        args: &[Expr],
        want: bool,
    ) -> Option<VExpr> {
        let (stages, source) = self.collect_chain(recv);
        let src = self.fusion_source(source);
        let cbs: Vec<Callback> = stages.iter().map(|s| self.prepare_cb(s.arg)).collect();
        let cb = self.prepare_cb(args.first()?);
        let result_ty = self.ty(e);
        let (term, result) = match intr {
            Intr::ArrForEach => (Terminal::ForEach(&cb), None),
            Intr::ArrReduce => {
                let acc = VExpr::Name(self.temp(&result_ty));
                self.expr_into(args.get(1)?, &acc, &result_ty);
                (
                    Terminal::Reduce {
                        acc: acc.clone(),
                        acc_ty: result_ty.clone(),
                        cb: &cb,
                    },
                    Some(acc),
                )
            }
            Intr::ArrSome | Intr::ArrEvery => {
                let flag = VExpr::Name(self.temp(&Type::Bool));
                let every = intr == Intr::ArrEvery;
                self.emit(assign(
                    flag.clone(),
                    lit(if every { "True" } else { "False" }),
                ));
                (
                    Terminal::Any {
                        flag: flag.clone(),
                        cb: &cb,
                        every,
                    },
                    Some(flag),
                )
            }
            Intr::ArrFind => {
                let out = VExpr::Name(self.temp(&result_ty));
                self.reset_var(&out, &result_ty);
                (
                    Terminal::Find {
                        out: out.clone(),
                        ty: result_ty.clone(),
                        cb: &cb,
                        index: false,
                    },
                    Some(out),
                )
            }
            Intr::ArrFindIndex => {
                let out = VExpr::Name(self.temp(&Type::Int));
                self.emit(assign(out.clone(), lit("-1&")));
                (
                    Terminal::Find {
                        out: out.clone(),
                        ty: Type::Int,
                        cb: &cb,
                        index: true,
                    },
                    Some(out),
                )
            }
            _ => return None,
        };
        let term_reads =
            intr == Intr::ArrFindIndex || cb.takes(if intr == Intr::ArrReduce { 2 } else { 1 });
        let plan = self.plan_indexes(&stages, &cbs, term_reads, None);
        let idx = self.fusion_index(&src, plan.source);
        self.fused_loop(&src, &idx, |lw, cur, cur_ty| {
            lw.run_stages(&stages, &cbs, 0, cur, cur_ty, &idx, &plan, &term)
        });
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
            let lname = fn_id
                .and_then(|f| self.prog.fns[f].locals.get(l))
                .map(|d| d.name.clone())
                .unwrap_or_else(|| format!("v{l}"));
            let lname = if lname == "this" {
                "this_".to_string()
            } else {
                lname
            };
            let mut field = format!("Captured_{}", names::valid_ident(&lname));
            while caps.iter().any(|c| c.1.eq_ignore_ascii_case(&field)) {
                field.push('_');
            }
            caps.push((l, field, self.local_ty(l), self.local_expr(l)));
        }
        // The closure is a class module called through `Object`: a `Type` cannot cross that boundary.
        let param_struct = params.iter().any(|&l| self.local_ty(l).has_struct());
        if param_struct || ret_ty.has_struct() || caps.iter().any(|c| c.2.has_struct()) {
            self.err(
                e.span,
                "E0395",
                "変数に入れる・引数として渡すアロー関数では、構造体を引数・戻り値にしたり外側の構造体変数を参照したりできません（クラスに変換されるためです）。配列メソッドの引数に直接書くか、通常の関数にしてください",
            );
        }
        // Constants of a script's `Main` that became locals are copied into the closure as well.
        let mut demoted_caps: Vec<(VarId, String, Type, VExpr)> = Vec::new();
        {
            let mut visit = |x: &Expr| {
                if let Some(Res::Global(v)) = info.res.get(&x.id) {
                    if let Some(n) = self.pc.demoted.get(v) {
                        if !demoted_caps.iter().any(|c| c.0 == *v) {
                            let field = format!(
                                "Captured_{}",
                                names::valid_ident(&self.prog.vars[*v].name)
                            );
                            demoted_caps.push((
                                *v,
                                field,
                                self.prog.vars[*v].ty.clone(),
                                VExpr::Name(n.clone()),
                            ));
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
        // In the code of a form, of a sheet or of the workbook, `this` is the module itself. A
        // closure of a form always knows its form: the functions and variables of the
        // code-behind are reached through it (`fn_callee`, `var_ref`).
        let object_module = self.prog.modules[self.frame().module]
            .form
            .as_ref()
            .map(|f| f.kind);
        let outer_this = self.frame().this.clone().or_else(|| {
            (object_module.is_some() && !self.pc.in_closure && !self.pc.in_class).then(|| lit("Me"))
        });
        let body_uses_this = outer_this.is_some()
            && (object_module == Some(ObjectKind::Form)
                || match &arrow.body {
                    ArrowBody::Expr(b) => any_expr_in(b, &mut |x| matches!(x.kind, ExprKind::This)),
                    ArrowBody::Block(b) => b.stmts.iter().any(|s| {
                        crate::astutil::any_expr(s, &mut |x| matches!(x.kind, ExprKind::This))
                    }),
                });

        let saved = std::mem::take(&mut self.pc);
        let outer = saved.frames.last().unwrap();
        self.pc.emit_module = saved.emit_module;
        self.pc.in_closure = true;
        self.pc.closure_span = Some(arrow.span);
        self.pc.exit_kw = "Function";
        self.pc.proc_name = "Apply".into();
        self.pc.ret_ty = ret_ty.clone();
        self.pc.display_name = saved.display_name.clone();
        self.pc.blocks.push(Vec::new());
        self.pc.names.insert("apply".into());
        let mut frame = Frame {
            module: outer.module,
            fn_id: outer.fn_id,
            env: outer.env.clone(),
            locals: HashMap::new(),
            this: body_uses_this.then(|| lit("Me").member("Captured_Me")),
            inl: None,
        };
        for (l, field, _, _) in &caps {
            self.pc.names.insert(field.to_lowercase());
            frame
                .locals
                .insert(*l, Binding::Expr(lit("Me").member(field.clone())));
        }
        for (v, field, _, _) in &demoted_caps {
            self.pc.names.insert(field.to_lowercase());
            self.pc.demoted.insert(*v, format!("Me.{field}"));
        }
        let mut vparams = Vec::new();
        for (i, &l) in params.iter().enumerate() {
            let n = self.unique(&self.local_ident(&arrow.params[i].name.text));
            frame.locals.insert(l, Binding::Var(n.clone()));
            vparams.push(VParam {
                name: n,
                byref: false,
                ty: "Variant".into(),
                is_array: false,
                optional_default: None,
            });
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
            self.out[ci]
                .decls
                .push("Public Captured_Me As Object".into());
        }
        self.out[ci].procs.push(VProc {
            kind: ProcKind::Function,
            public: true,
            friend: false,
            name: "Apply".into(),
            params: vparams,
            ret: Some("Variant".into()),
            locals,
            body,
            leading_comments: vec![],
            attributes: vec![],
            source: Some(0),
        });

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
        // spec §4.1: tell the user that this lambda is slow (not the handler of an event: it
        // has to be an object).
        if self.quiet_closures {
            return t;
        }
        self.diag(Severity::Info, self.src_module(), arrow.span, "I0641", format!("このアロー関数はクラス（{cname}）に変換されるため、呼び出しが遅くなります。配列メソッドの引数に直接書くとループに展開されます"));
        t
    }
}
