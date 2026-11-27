//! Statement code generation: one JavaScript statement per RAS statement.

use ras_syntax::ast::{self, Expr, ExprKind, Stmt, StmtKind, VarKind};

use super::{js_ident, Gen};
use crate::intrinsics::Intr;
use crate::model::*;

impl<'p> Gen<'p> {
    pub(crate) fn stmts(&mut self, stmts: &[Stmt]) {
        for (i, s) in stmts.iter().enumerate() {
            if let StmtKind::Var(v) = &s.kind {
                if v.kind == VarKind::Using {
                    // `using` covers the rest of the block (spec §5.3).
                    self.using_stmt(s, v, &stmts[i + 1..]);
                    return;
                }
            }
            self.stmt(s);
        }
    }

    /// The body of a compound statement, one level deeper.
    fn body(&mut self, s: &Stmt) {
        self.indent += 1;
        match &s.kind {
            StmtKind::Block(b) => self.stmts(&b.stmts),
            _ => self.stmt(s),
        }
        self.indent -= 1;
    }

    /// A simple statement. Inside `try { } catch { resume; }` it ignores its own errors.
    fn simple(&mut self, s: &Stmt, text: String) {
        if self.resume > 0 {
            self.wm(s.span, format!("try {{ {text} }} catch {{}}"));
        } else {
            self.wm(s.span, text);
        }
    }

    pub(crate) fn stmt(&mut self, s: &Stmt) {
        match &s.kind {
            StmtKind::Empty | StmtKind::Error | StmtKind::Resume | StmtKind::Fn(_) => {}
            // Pauses only while a debugger is attached (DevTools, VS Code); no-op otherwise.
            StmtKind::Debugger => self.wm(s.span, "debugger;"),
            StmtKind::Var(v) => self.var_stmt(s, v),
            StmtKind::Expr(e) => {
                let text = self.effect(e.peel());
                self.simple(s, format!("{text};"));
            }
            StmtKind::Block(b) => {
                self.w("{");
                self.indent += 1;
                self.stmts(&b.stmts);
                self.indent -= 1;
                self.w("}");
            }
            StmtKind::If { cond, then, els } => {
                let c = self.expr(cond);
                self.wm(s.span, format!("if ({c}) {{"));
                self.body(then);
                let mut next = els.as_deref();
                // `else if` chains stay flat.
                while let Some(e) = next {
                    match &e.kind {
                        StmtKind::If { cond, then, els } => {
                            let c = self.expr(cond);
                            self.wm(e.span, format!("}} else if ({c}) {{"));
                            self.body(then);
                            next = els.as_deref();
                        }
                        _ => {
                            self.w("} else {");
                            self.body(e);
                            next = None;
                        }
                    }
                }
                self.w("}");
            }
            StmtKind::For { .. } => self.for_stmt(s, ""),
            StmtKind::ForOf { .. } => self.for_of_stmt(s, ""),
            StmtKind::While { cond, body, .. } => self.while_stmt(s, cond, body, ""),
            StmtKind::DoWhile { body, cond, .. } => self.do_while_stmt(s, body, cond, ""),
            StmtKind::Labeled { label, body } => {
                let prefix = format!("{}: ", js_ident(&label.text));
                match &body.kind {
                    StmtKind::For { .. } => self.for_stmt(body, &prefix),
                    StmtKind::ForOf { .. } => self.for_of_stmt(body, &prefix),
                    StmtKind::While { cond, body: b, .. } => {
                        self.while_stmt(body, cond, b, &prefix)
                    }
                    StmtKind::DoWhile { body: b, cond, .. } => {
                        self.do_while_stmt(body, b, cond, &prefix)
                    }
                    _ => {
                        self.w(format!("{prefix}{{"));
                        self.body(body);
                        self.w("}");
                    }
                }
            }
            StmtKind::Switch { disc, cases } => {
                let d = self.expr(disc);
                self.wm(s.span, format!("switch ({d}) {{"));
                self.indent += 1;
                for c in cases {
                    match &c.test {
                        Some(t) => {
                            let v = self.expr(t);
                            self.w(format!("case {v}: {{"));
                        }
                        None => self.w("default: {"),
                    }
                    self.indent += 1;
                    self.stmts(&c.body);
                    self.indent -= 1;
                    self.w("}");
                }
                self.indent -= 1;
                self.w("}");
            }
            StmtKind::Try(t) => self.try_stmt(s, t),
            StmtKind::Throw(e) => {
                let text = match &e.peel().kind {
                    ExprKind::New { .. } => format!("throw {};", self.expr(e)),
                    _ if self.ty(e) == Type::Str => {
                        format!("throw $rt.makeError({});", self.expr(e))
                    }
                    _ => format!("throw {};", self.expr(e)),
                };
                self.simple(s, text);
            }
            StmtKind::Return(value) => {
                let text = match (value, self.ret_ty.clone()) {
                    (None, _) => "return;".to_string(),
                    (Some(v), Some(Type::Void)) => format!("{}; return;", self.effect(v.peel())),
                    (Some(v), Some(ty)) => format!("return {};", self.coerce_store(v, &ty)),
                    (Some(v), None) => format!("return {};", self.expr(v)),
                };
                self.wm(s.span, text);
            }
            StmtKind::Break(label) => self.wm(
                s.span,
                format!(
                    "break{};",
                    label
                        .as_ref()
                        .map(|l| format!(" {}", js_ident(&l.text)))
                        .unwrap_or_default()
                ),
            ),
            StmtKind::Continue(label) => self.wm(
                s.span,
                format!(
                    "continue{};",
                    label
                        .as_ref()
                        .map(|l| format!(" {}", js_ident(&l.text)))
                        .unwrap_or_default()
                ),
            ),
        }
    }

    /// `let x = init` for a local declaration, or `x = init` for a top-level variable (declared at module level).
    fn var_decl(&mut self, v: &ast::VarDecl) -> (String, String) {
        let offset = u32::from(v.name.span.start());
        let info = self.info();
        if let Some(&var) = info.decl_globals.get(&offset) {
            let ty = self.prog.vars[var].ty.clone();
            let target = self.global_ref(var);
            return match &v.init {
                Some(init) => (
                    String::new(),
                    format!("{target} = {}", self.coerce_store(init, &ty)),
                ),
                None => (String::new(), String::new()),
            };
        }
        let Some(&local) = info.decl_locals.get(&offset) else {
            return (String::new(), String::new());
        };
        let ty = self.local_ty(local);
        let name = self.local_name(local);
        let kw = if v.kind == VarKind::Let {
            "let"
        } else {
            "const"
        };
        let value = match &v.init {
            Some(init) => self.coerce_store(init, &ty),
            None => self.zero(&ty),
        };
        (format!("{kw} "), format!("{name} = {value}"))
    }

    fn var_stmt(&mut self, s: &Stmt, v: &ast::VarDecl) {
        let (kw, assign) = self.var_decl(v);
        if assign.is_empty() {
            return;
        }
        if self.resume > 0 && !kw.is_empty() {
            // The declaration must survive a failing initialiser.
            let name = assign.split(" = ").next().unwrap_or_default().to_string();
            let offset = u32::from(v.name.span.start());
            let ty = self
                .info()
                .decl_locals
                .get(&offset)
                .map(|&l| self.local_ty(l))
                .unwrap_or(Type::Any);
            let z = self.zero(&ty);
            self.w(format!("let {name} = {z};"));
            self.wm(s.span, format!("try {{ {assign}; }} catch {{}}"));
        } else {
            self.simple(s, format!("{kw}{assign};"));
        }
    }

    fn for_stmt(&mut self, s: &Stmt, label: &str) {
        let StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } = &s.kind
        else {
            return;
        };
        // Several declarations in the header: hoist them into an enclosing block.
        let hoist = init.len() > 1
            || init
                .iter()
                .any(|i| !matches!(i.kind, StmtKind::Var(_) | StmtKind::Expr(_)));
        if hoist {
            self.w("{");
            self.indent += 1;
            for i in init {
                self.stmt(i);
            }
        }
        let head_init = match init.first().filter(|_| !hoist).map(|i| &i.kind) {
            Some(StmtKind::Var(v)) => {
                let (kw, assign) = self.var_decl(v);
                format!("{}{assign}", if kw.is_empty() { "" } else { "let " })
            }
            Some(StmtKind::Expr(e)) => self.effect(e.peel()),
            _ => String::new(),
        };
        let c = cond.as_ref().map(|c| self.expr(c)).unwrap_or_default();
        let u: Vec<String> = update.iter().map(|x| self.effect(x.peel())).collect();
        self.wm(
            s.span,
            format!("{label}for ({head_init}; {c}; {}) {{", u.join(", ")),
        );
        self.body(body);
        self.w("}");
        if hoist {
            self.indent -= 1;
            self.w("}");
        }
    }

    fn for_of_stmt(&mut self, s: &Stmt, label: &str) {
        let StmtKind::ForOf {
            kind,
            name,
            iter,
            body,
            ..
        } = &s.kind
        else {
            return;
        };
        let info = self.info();
        let Some(&local) = info.decl_locals.get(&u32::from(name.span.start())) else {
            return;
        };
        let var_ty = self.local_ty(local);
        let var = self.local_name(local);
        let iter_ty = self.ty(iter);
        let kw = if *kind == VarKind::Let {
            "let"
        } else {
            "const"
        };
        // `for (const i of a..b)` (spec §4.5): the end is evaluated once.
        if let ExprKind::Range { start, end } = &iter.kind {
            let a = match start.as_deref() {
                Some(x) => self.expr(x),
                None => "0".into(),
            };
            let b = match end.as_deref() {
                Some(x) => self.expr(x),
                None => "0".into(),
            };
            let converted = self.coerce("$i".into(), &iter_ty, &var_ty);
            self.wm(
                s.span,
                format!("{label}for (let $i = {a}, $n = {b}; $i < $n; $i++) {{"),
            );
            self.w(format!("    {kw} {var} = {converted};"));
            self.body(body);
            self.w("}");
            return;
        }
        let is_map_iter = matches!(
            info.calls.get(&iter.peel().id).map(|c| &c.target),
            Some(CallTarget::Intrinsic(Intr::MapKeys | Intr::MapValues))
        );
        let it = self.expr(iter);
        let (coll, elem_ty) = if iter_ty.is_array() || is_map_iter {
            let elem = match &iter_ty {
                Type::Array(inner) => (**inner).clone(),
                _ => Type::Any,
            };
            (it, elem)
        } else {
            // Host collection (`For Each`).
            (format!("$rt.iter({it})"), Type::Any)
        };
        let converted = self.coerce("$item".into(), &elem_ty, &var_ty);
        // The loop variable of a struct array is a copy of the element (a value type).
        let converted = if matches!(var_ty, Type::Struct(_)) {
            format!("$rt.clone({converted})")
        } else {
            converted
        };
        if converted == "$item" {
            self.wm(s.span, format!("{label}for ({kw} {var} of {coll}) {{"));
        } else {
            self.wm(s.span, format!("{label}for (const $item of {coll}) {{"));
            self.w(format!("    {kw} {var} = {converted};"));
        }
        self.body(body);
        self.w("}");
    }

    fn while_stmt(&mut self, s: &Stmt, cond: &Expr, body: &Stmt, label: &str) {
        let c = self.expr(cond);
        self.wm(s.span, format!("{label}while ({c}) {{"));
        self.body(body);
        self.w("}");
    }

    fn do_while_stmt(&mut self, s: &Stmt, body: &Stmt, cond: &Expr, label: &str) {
        self.wm(s.span, format!("{label}do {{"));
        self.body(body);
        let c = self.expr(cond);
        self.wm(cond.span, format!("}} while ({c});"));
    }

    fn try_stmt(&mut self, s: &Stmt, t: &ast::TryStmt) {
        if t.catch.as_ref().is_some_and(|c| c.is_resume_block()) {
            // spec §5.2 (`On Error Resume Next`): a failing statement is skipped, the next one runs.
            self.w("{");
            self.indent += 1;
            self.resume += 1;
            self.stmts(&t.block.stmts);
            self.resume -= 1;
            self.indent -= 1;
            self.w("}");
            if let Some(f) = &t.finally {
                self.w("{");
                self.indent += 1;
                self.stmts(&f.stmts);
                self.indent -= 1;
                self.w("}");
            }
            return;
        }
        // Statements inside a real `try` raise normally again.
        let saved = std::mem::take(&mut self.resume);
        self.wm(s.span, "try {");
        self.indent += 1;
        self.stmts(&t.block.stmts);
        self.indent -= 1;
        if let Some(c) = &t.catch {
            match &c.param {
                Some(p) => {
                    let name = self
                        .info()
                        .decl_locals
                        .get(&u32::from(p.span.start()))
                        .map(|&l| self.local_name(l))
                        .unwrap_or_else(|| js_ident(&p.text));
                    self.w("} catch ($e) {");
                    self.wm(c.span, format!("    const {name} = $rt.caught($e);"));
                }
                None => self.w("} catch {"),
            }
            self.indent += 1;
            self.stmts(&c.body.stmts);
            self.indent -= 1;
        }
        if let Some(f) = &t.finally {
            self.w("} finally {");
            self.indent += 1;
            self.stmts(&f.stmts);
            self.indent -= 1;
        }
        if t.catch.is_none() && t.finally.is_none() {
            self.w("} finally {");
        }
        self.w("}");
        self.resume = saved;
    }

    fn using_stmt(&mut self, s: &Stmt, v: &ast::VarDecl, rest: &[Stmt]) {
        let Some(init) = &v.init else { return };
        let info = self.info();
        let is_fast = matches!(
            info.calls.get(&init.peel().id).map(|c| &c.target),
            Some(CallTarget::Intrinsic(Intr::FastMode))
        );
        self.w("{");
        self.indent += 1;
        let cleanup = if is_fast {
            let mask = crate::lower::expr::fast_mode_mask(init);
            self.wm(s.span, format!("$rt.suspend({mask});"));
            format!("$rt.resume({mask});")
        } else {
            let Some(&local) = info.decl_locals.get(&u32::from(v.name.span.start())) else {
                self.indent -= 1;
                self.w("}");
                return;
            };
            let name = self.local_name(local);
            let ty = self.local_ty(local);
            let value = self.coerce_store(init, &ty);
            self.wm(s.span, format!("const {name} = {value};"));
            format!("if ({name} != null) {name}.dispose();")
        };
        let saved = std::mem::take(&mut self.resume);
        self.w("try {");
        self.indent += 1;
        self.stmts(rest);
        self.indent -= 1;
        self.w("} finally {");
        self.w(format!("    {cleanup}"));
        self.w("}");
        self.resume = saved;
        self.indent -= 1;
        self.w("}");
    }
}
