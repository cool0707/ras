//! Range ⇔ 2-D array bulk transfer (spec §9): loops over `ws.Cells(r, c).Value`
//! become one read of / one write to a rectangular Range.

use std::collections::{HashMap, HashSet};

use ras_syntax::ast::{self, BinaryOp, Expr, ExprId, ExprKind, Stmt, StmtKind};

use super::{Binding, Lowerer};
use crate::astutil::{any_stmt, walk_stmt};
use crate::model::*;
use crate::vba::*;

#[derive(Clone)]
pub(crate) struct BulkSlot {
    /// Buffer element that replaces the cell access.
    pub element: VExpr,
}

#[derive(Clone, PartialEq, Debug)]
enum Axis {
    /// Induction variable of one of the loops.
    Var(LocalId),
    /// Literal row / column number.
    Lit,
    /// Loop-invariant variable (same one for every access of the group).
    Base(String),
}

struct Access<'a> {
    id: ExprId,
    sheet: String,
    sheet_expr: &'a Expr,
    prop: String,
    row: (Axis, i64, Option<&'a Expr>),
    col: (Axis, i64, Option<&'a Expr>),
    write: bool,
}

struct LoopInfo<'a> {
    local: LocalId,
    name: &'a str,
    from: &'a Expr,
    bound: &'a Expr,
    inclusive: bool,
}

struct Group {
    write: bool,
    sheet: VExpr,
    prop: String,
    buf: String,
    rlo: VExpr,
    rhi: VExpr,
    clo: VExpr,
    chi: VExpr,
    slots: Vec<(ExprId, VExpr)>,
}

pub(crate) struct BulkPlan {
    groups: Vec<Group>,
    guard: Option<VExpr>,
}

const PROPS: &[&str] = &["Value", "Value2", "Formula", "FormulaR1C1"];

fn block_stmts(s: &Stmt) -> &[Stmt] {
    match &s.kind {
        StmtKind::Block(b) => &b.stmts,
        _ => std::slice::from_ref(s),
    }
}

fn diff(hi: VExpr, lo: &VExpr) -> VExpr {
    match (hi.as_long(), lo.as_long()) {
        (Some(h), Some(l)) => long(h - l),
        (_, Some(l)) => hi.plus(-l),
        _ => bin("-", hi, lo.clone()),
    }
}

impl<'p> Lowerer<'p> {
    fn refuse_bulk(&mut self, s: &Stmt, why: &str) -> Option<BulkPlan> {
        self.note(ast::Span::at(s.span.start(), 3.into()), false, "bulk", format!("セルアクセスを一括転送に変換しません: {why}"));
        None
    }

    /// `X.Cells(row, col).Prop` → (sheet expr, row, col, prop)
    fn cell_access<'a>(&self, e: &'a Expr) -> Option<(&'a Expr, &'a Expr, &'a Expr, &'p str)> {
        let info = self.info();
        let ExprKind::Member { object, .. } = &e.kind else { return None };
        // The property by its type-library name (`value` in the source is `Value`).
        let Some(Res::Field(fc, fi)) = info.res.get(&e.id) else { return None };
        let prog: &'p Program = self.prog;
        let prop = prog.classes[*fc].fields[*fi].vba_name();
        if prog.classes[*fc].kind != ClassKind::Host || !PROPS.contains(&prop) {
            return None;
        }
        let ExprKind::Call { callee, args, .. } = &object.peel().kind else { return None };
        let Some(CallTarget::Host(f)) = info.calls.get(&object.peel().id).map(|c| &c.target) else { return None };
        let def = &self.prog.fns[*f];
        let FnKind::Host { class: Some(c) } = def.kind else { return None };
        if def.vba_name() != "Cells" || self.prog.classes[c].name != "Worksheet" || args.len() != 2 {
            return None;
        }
        let ExprKind::Member { object: sheet, .. } = &callee.peel().kind else { return None };
        Some((sheet, &args[0], &args[1], prop))
    }

    fn affine<'a>(&self, e: &'a Expr, loops: &[LoopInfo], whole: &Stmt) -> Option<(Axis, i64, Option<&'a Expr>)> {
        let p = e.peel();
        match &p.kind {
            ExprKind::Int { value, .. } => Some((Axis::Lit, i64::try_from(*value).ok()?, None)),
            ExprKind::Name(id) => {
                let l = self.local_of(p);
                if let Some(l) = l {
                    if loops.iter().any(|x| x.local == l) {
                        return Some((Axis::Var(l), 0, None));
                    }
                }
                self.is_invariant(p, whole).then(|| (Axis::Base(id.text.clone()), 0, Some(p)))
            }
            ExprKind::Binary { op: op @ (BinaryOp::Add | BinaryOp::Sub), lhs, rhs } => {
                let k = |x: &Expr| match &x.peel().kind {
                    ExprKind::Int { value, .. } => i64::try_from(*value).ok(),
                    _ => None,
                };
                if let Some(c) = k(rhs) {
                    let (axis, off, base) = self.affine(lhs, loops, whole)?;
                    if axis == Axis::Lit {
                        return None;
                    }
                    return Some((axis, if *op == BinaryOp::Add { off + c } else { off - c }, base));
                }
                if let (BinaryOp::Add, Some(c)) = (op, k(lhs)) {
                    let (axis, off, base) = self.affine(rhs, loops, whole)?;
                    if axis == Axis::Lit {
                        return None;
                    }
                    return Some((axis, off + c, base));
                }
                None
            }
            _ => None,
        }
    }

    /// A small function that the loop will inline and that makes no calls of its own.
    fn expands_harmlessly(&self, f: FnId) -> bool {
        let d = &self.prog.fns[f];
        self.opt >= 1
            && matches!(d.kind, FnKind::Free)
            && d.decl.as_ref().is_some_and(|x| x.body.is_some())
            && d.inline != ast::InlineMode::NoInline
            && d.inline_blockers.is_empty()
            && d.calls.is_empty()
            && !self.pc.inline_stack.contains(&f)
            && (d.inline == ast::InlineMode::Inline || d.cost <= self.cfg.inline_threshold())
    }

    /// Analyses loop `s`; returns a plan when every condition of spec §9.3 holds.
    pub(crate) fn bulk_plan(&mut self, s: &Stmt) -> Option<BulkPlan> {
        if self.opt == 0 {
            return None;
        }
        let StmtKind::For { decorators, init, cond, update, body } = &s.kind else { return None };
        if decorators.iter().any(|d| d.name.text == "nobulk") {
            return None;
        }
        let force = decorators.iter().any(|d| d.name.text == "bulk");
        let info = self.info();

        // 1. candidate accesses
        let mut found: Vec<(&Expr, &Expr, &Expr, &Expr, &str)> = Vec::new();
        let mut write_targets: HashSet<ExprId> = HashSet::new();
        let mut compound_targets: HashSet<ExprId> = HashSet::new();
        walk_stmt(s, &mut |_| true, &mut |e| {
            if let ExprKind::Assign { op, target, .. } = &e.kind {
                if op.is_some() {
                    compound_targets.insert(target.peel().id);
                } else {
                    write_targets.insert(target.peel().id);
                }
            }
            if let ExprKind::IncDec { operand, .. } = &e.kind {
                compound_targets.insert(operand.peel().id);
            }
            if let Some((sheet, r, c, p)) = self.cell_access(e) {
                found.push((e, sheet, r, c, p));
            }
            true
        });
        if found.is_empty() {
            return None;
        }

        // 2. loop nest: `for (let r = a; r <= b; r++)`, optionally with one inner loop of the same shape
        let ok_shape = |sh: &super::stmt::ForShape| sh.decl.is_some() && sh.step == 1 && matches!(sh.op, BinaryOp::Lt | BinaryOp::Le);
        let outer = match self.for_shape(init, cond.as_ref(), update, body) {
            Some(sh) if ok_shape(&sh) => sh,
            _ => return self.refuse_bulk(s, "ループが `for (let r = a; r <= b; r++)` の形ではないか、範囲がループ内で変化します"),
        };
        let mut loops = vec![LoopInfo { local: outer.local, name: &outer.decl.unwrap().name.text, from: outer.from, bound: outer.bound, inclusive: outer.op == BinaryOp::Le }];
        let mut unconditional: HashSet<ExprId> = HashSet::new();
        let collect_direct = |stmts: &[Stmt], set: &mut HashSet<ExprId>| {
            for st in stmts {
                if let StmtKind::Expr(e) = &st.kind {
                    if let ExprKind::Assign { op: None, target, .. } = &e.peel().kind {
                        set.insert(target.peel().id);
                    }
                }
            }
        };
        collect_direct(block_stmts(body), &mut unconditional);
        for st in block_stmts(body) {
            let StmtKind::For { init: i2, cond: c2, update: u2, body: b2, .. } = &st.kind else { continue };
            if let Some(sh) = self.for_shape(i2, c2.as_ref(), u2, b2) {
                let depends_on_outer = [sh.from, sh.bound].iter().any(|x| crate::astutil::any_expr_in(x, &mut |y| self.local_of(y) == Some(outer.local)));
                if ok_shape(&sh) && !depends_on_outer && self.is_invariant(sh.from, body) && self.is_invariant(sh.bound, body) {
                    loops.push(LoopInfo { local: sh.local, name: &sh.decl.unwrap().name.text, from: sh.from, bound: sh.bound, inclusive: sh.op == BinaryOp::Le });
                    collect_direct(block_stmts(b2), &mut unconditional);
                    break;
                }
            }
        }

        // 3. classify the accesses
        let mut accesses: Vec<Access> = Vec::new();
        let mut cells_calls: HashSet<ExprId> = HashSet::new();
        let mut sheet_uses: HashSet<ExprId> = HashSet::new();
        for (e, sheet, r, c, prop) in found {
            // Condition 1: the sheet must not change inside the loop.
            let sheet_name = match (&sheet.peel().kind, info.res.get(&sheet.peel().id)) {
                (ExprKind::Name(id), Some(Res::Local(l))) if !self.modifies_local(s, *l) => id.text.clone(),
                (ExprKind::Name(id), Some(Res::Global(v))) if self.prog.vars[*v].declared || self.prog.vars[*v].is_const => id.text.clone(),
                _ => return self.refuse_bulk(s, "対象シートがループ内で不変と判定できません（`const` 変数かコード名を使ってください）"),
            };
            if compound_targets.contains(&e.id) {
                return self.refuse_bulk(s, "同じセルを読んで書き戻しています（`+=` / `++`）");
            }
            // Condition 2: indices are `induction variable ± constant` or invariant.
            let (Some(row), Some(col)) = (self.affine(r, &loops, s), self.affine(c, &loops, s)) else {
                return self.refuse_bulk(s, "行・列の添字が「誘導変数 ± 定数」の形ではありません");
            };
            if let ExprKind::Member { object, .. } = &e.kind {
                cells_calls.insert(object.peel().id);
            }
            sheet_uses.insert(sheet.peel().id);
            accesses.push(Access { id: e.id, sheet: sheet_name, sheet_expr: sheet, prop: prop.to_string(), row, col, write: write_targets.contains(&e.id) });
        }

        // Condition 4: nothing else may touch the sheet.
        if !force {
            let sheets: HashSet<&str> = accesses.iter().map(|a| a.sheet.as_str()).collect();
            let mut other_use = false;
            let mut unknown_call = false;
            walk_stmt(s, &mut |_| true, &mut |e| {
                match &e.kind {
                    ExprKind::Name(id) if sheets.contains(id.text.as_str()) && !sheet_uses.contains(&e.id) => {
                        if matches!(info.res.get(&e.id), Some(Res::Local(_) | Res::Global(_))) {
                            other_use = true;
                        }
                    }
                    ExprKind::Call { .. } if !cells_calls.contains(&e.id) => match info.calls.get(&e.id).map(|c| &c.target) {
                        Some(CallTarget::Intrinsic(_)) => {}
                        // Calls that will be expanded in place and call nothing themselves cannot touch the sheet.
                        Some(CallTarget::Fn(f) | CallTarget::Ext(f)) if self.expands_harmlessly(*f) => {}
                        _ => unknown_call = true,
                    },
                    ExprKind::New { .. } => unknown_call = true,
                    _ => {}
                }
                true
            });
            if other_use {
                return self.refuse_bulk(s, "ループ内で同じシートに対する他の操作があります");
            }
            if unknown_call {
                return self.refuse_bulk(s, "ループ内にシートへ触れうる関数呼び出しがあります（`@bulk` を付けると無視して変換します）");
            }
        }
        // Condition 6: no reading back of written sheets.
        let written: HashSet<&str> = accesses.iter().filter(|a| a.write).map(|a| a.sheet.as_str()).collect();
        if accesses.iter().any(|a| !a.write && written.contains(a.sheet.as_str())) {
            return self.refuse_bulk(s, "書き込むシートを同じループ内で読み出しています");
        }
        let has_write = !written.is_empty();
        if has_write {
            // Conditions 5 and 7.
            if any_stmt(s, &mut |x| matches!(x.kind, StmtKind::Break(_) | StmtKind::Continue(_) | StmtKind::Return(_) | StmtKind::Throw(_))) {
                return self.refuse_bulk(s, "書き込みループに途中脱出（break / continue / return / throw）があります");
            }
            if accesses.iter().any(|a| a.write && !unconditional.contains(&a.id)) {
                return self.refuse_bulk(s, "条件付きの書き込みがあります（矩形内の全セルに無条件で書き込む必要があります）");
            }
        }

        // 4. group by (sheet, property, direction) and work out the rectangle
        let mut order: Vec<(String, String, bool)> = Vec::new();
        let mut by_key: HashMap<(String, String, bool), Vec<usize>> = HashMap::new();
        for (i, a) in accesses.iter().enumerate() {
            let key = (a.sheet.clone(), a.prop.clone(), a.write);
            if !by_key.contains_key(&key) {
                order.push(key.clone());
            }
            by_key.entry(key).or_default().push(i);
        }
        struct AxisPlan {
            axis: Axis,
            kmin: i64,
            kmax: i64,
        }
        let mut plans: Vec<(AxisPlan, AxisPlan)> = Vec::new();
        for key in &order {
            let idxs = &by_key[key];
            let mut axes = Vec::new();
            for pick in [0usize, 1] {
                let get = |a: &Access| if pick == 0 { (a.row.0.clone(), a.row.1) } else { (a.col.0.clone(), a.col.1) };
                let (first_axis, _) = get(&accesses[idxs[0]]);
                if idxs.iter().any(|&i| get(&accesses[i]).0 != first_axis) {
                    return self.refuse_bulk(s, "同じシートへのアクセスで行・列の添字の形が揃っていません");
                }
                let offs: Vec<i64> = idxs.iter().map(|&i| get(&accesses[i]).1).collect();
                let (kmin, kmax) = (*offs.iter().min().unwrap(), *offs.iter().max().unwrap());
                if key.2 {
                    // Writes must cover the rectangle completely.
                    let covered = match &first_axis {
                        Axis::Var(_) => kmin == kmax,
                        _ => (kmin..=kmax).all(|k| offs.contains(&k)),
                    };
                    if !covered {
                        return self.refuse_bulk(s, "書き込みが矩形領域の全セルを覆っていません");
                    }
                }
                if kmax - kmin > 16_384 {
                    return self.refuse_bulk(s, "対象範囲が広すぎます");
                }
                axes.push(AxisPlan { axis: first_axis, kmin, kmax });
            }
            let col = axes.pop().unwrap();
            let row = axes.pop().unwrap();
            plans.push((row, col));
        }

        // 5. everything checks out: bind the loop variables and build the VBA pieces
        let mut var_names: HashMap<LocalId, (VExpr, VExpr, VExpr)> = HashMap::new(); // local → (var, from, last)
        let mut guard: Option<VExpr> = None;
        for l in &loops {
            let ty = self.local_ty(l.local);
            let n = self.local_name(l.name, &ty);
            self.bind(l.local, Binding::Var(n.clone()));
            let from = self.expr(l.from);
            let bound = self.expr(l.bound);
            let last = if l.inclusive { bound.clone() } else { bound.clone().plus(-1) };
            // The transfer only happens when the loops run at least once (spec §9.1).
            let c = match (bound.as_long(), from.as_long()) {
                (Some(b), Some(f)) => {
                    if (l.inclusive && b < f) || (!l.inclusive && b <= f) {
                        return None;
                    }
                    None
                }
                _ => Some(bin(if l.inclusive { ">=" } else { ">" }, bound, from.clone())),
            };
            if let Some(c) = c {
                guard = Some(match guard {
                    None => c,
                    Some(g) => bin("And", g, c),
                });
            }
            var_names.insert(l.local, (name(n), from, last));
        }
        let mut groups = Vec::new();
        for (key, (row, col)) in order.iter().zip(plans) {
            let idxs = &by_key[key];
            let first = &accesses[idxs[0]];
            let sheet = self.expr(first.sheet_expr);
            let range = |lw: &mut Self, p: &AxisPlan, base: Option<&Expr>| -> (VExpr, VExpr) {
                match &p.axis {
                    Axis::Var(l) => {
                        let (_, from, last) = &var_names[l];
                        (from.clone().plus(p.kmin), last.clone().plus(p.kmax))
                    }
                    Axis::Lit => (long(p.kmin), long(p.kmax)),
                    Axis::Base(_) => {
                        let b = lw.expr(base.unwrap());
                        (b.clone().plus(p.kmin), b.plus(p.kmax))
                    }
                }
            };
            let (rlo, rhi) = range(self, &row, first.row.2);
            let (clo, chi) = range(self, &col, first.col.2);
            self.pc.buf_n += 1;
            let buf = if key.2 {
                self.declare_raw(&format!("ras_buf{}", self.pc.buf_n), "Variant", true)
            } else {
                self.declare_raw(&format!("ras_src{}", self.pc.buf_n), "Variant", false)
            };
            let index = |p: &AxisPlan, off: i64| -> VExpr {
                match &p.axis {
                    Axis::Var(l) => {
                        let (var, from, _) = &var_names[l];
                        diff(var.clone(), from).plus(off - p.kmin)
                    }
                    _ => long(off - p.kmin),
                }
            };
            let slots = idxs
                .iter()
                .map(|&i| {
                    let a = &accesses[i];
                    (a.id, name(buf.clone()).index(vec![index(&row, a.row.1), index(&col, a.col.1)]))
                })
                .collect();
            groups.push(Group { write: key.2, sheet, prop: key.1.clone(), buf, rlo, rhi, clo, chi, slots });
        }
        Some(BulkPlan { groups, guard })
    }

    pub(crate) fn emit_bulk(&mut self, s: &Stmt, plan: BulkPlan, label: Option<String>) {
        let range_of = |g: &Group| {
            let cells = |r: &VExpr, c: &VExpr| g.sheet.clone().member("Cells").index(vec![r.clone(), c.clone()]);
            g.sheet.clone().member("Range").index(vec![cells(&g.rlo, &g.clo), cells(&g.rhi, &g.chi)])
        };
        let (inner, _) = self.capture(|lw| {
            for g in &plan.groups {
                if g.write {
                    lw.emit(VStmt::ReDim { preserve: false, target: name(g.buf.clone()), dims: vec![(long(0), diff(g.rhi.clone(), &g.rlo)), (long(0), diff(g.chi.clone(), &g.clo))] });
                } else {
                    let value = if g.prop == "Value" { call("RAS_Read2D", vec![range_of(g)]) } else { call("RAS_Rebase", vec![range_of(g).member(g.prop.clone())]) };
                    lw.emit(assign(name(g.buf.clone()), value));
                }
                for (id, element) in &g.slots {
                    lw.pc.bulk.insert(*id, BulkSlot { element: element.clone() });
                }
            }
            lw.pc.bulk_depth = lw.pc.frames.len();
            lw.pc.suppress_yield = true;
            lw.for_stmt_plain(s, label);
            lw.pc.suppress_yield = false;
            lw.pc.bulk.clear();
            for g in plan.groups.iter().filter(|g| g.write) {
                lw.emit(assign(range_of(g).member(g.prop.clone()), name(g.buf.clone())));
            }
        });
        match plan.guard {
            Some(c) => self.emit(if_then(c, inner)),
            None => {
                for st in inner {
                    self.emit(st);
                }
            }
        }
        let reads = plan.groups.iter().filter(|g| !g.write).count();
        let writes = plan.groups.len() - reads;
        self.note(ast::Span::at(s.span.start(), 3.into()), true, "bulk", format!("セルアクセスを Range ⇔ 配列の一括転送に変換しました（読み込み {reads} 件、書き込み {writes} 件）"));
    }
}
