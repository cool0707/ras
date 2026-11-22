//! Type checking of function bodies, plus the lint rules of spec §6.

use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use ras_syntax::ast::{
    self, BinaryOp, Expr, ExprId, ExprKind, Ident, PassMode, Span, Stmt, StmtKind, UnaryOp,
    Visibility,
};

use crate::collect::{const_int, resolve_type};
use crate::diag::{Diagnostic, Severity};
use crate::intrinsics::{self, Intr};
use crate::model::*;

#[derive(Default)]
struct Scope {
    names: HashMap<String, LocalId>,
    /// Set on the scope that holds an arrow function's parameters.
    lambda: Option<ExprId>,
}

#[derive(Default)]
struct LambdaRet {
    expected: Option<Type>,
    inferred: Option<Type>,
}

#[derive(Default)]
struct FnCtx {
    /// Depth of arrow functions that are the handler of an `on` call (spec §12.2.3): they
    /// may assign to module variables.
    event_handlers: u32,
    fn_id: Option<FnId>,
    class: Option<ClassId>,
    is_static: bool,
    this_local: Option<LocalId>,
    type_params: Vec<TypeParamDef>,
    locals: Vec<LocalDef>,
    scopes: Vec<Scope>,
    ret: Option<Type>,
    inferred: Option<Type>,
    lambda_rets: Vec<LambdaRet>,
    loops: Vec<Option<String>>,
    switch_depth: usize,
    calls: Vec<FnId>,
    var_refs: Vec<VarId>,
    has_loops: bool,
    blockers: Vec<&'static str>,
    cost: u32,
    row_views: HashMap<ExprId, Span>,
    /// `Main` / module initializer: top-level `let` are module variables.
    is_toplevel: bool,
    captures: Vec<(LocalId, Span)>,
    /// Inside the arguments of `ws.Range(...)`: unqualified `Cells` / `Range` are auto-qualified.
    qualify_depth: usize,
    in_field_init: bool,
    /// Locals whose type is narrowed by `instanceof` / `typeof` here (innermost last).
    narrowed: Vec<(LocalId, Type)>,
    /// `x instanceof T` → the resolved `T`.
    instanceof_types: HashMap<ExprId, Type>,
    /// Method of a union-typed receiver (callee member expression) → the methods of the other members.
    union_calls: HashMap<ExprId, Vec<(FnId, HashMap<String, Type>)>>,
    /// Property of a union-typed receiver → (every member's is writable, the type they share).
    union_props: HashMap<ExprId, (bool, Option<Type>)>,
}

/// Narrowed types of locals that hold when a condition is true (or false).
type Facts = Vec<(LocalId, Type)>;

/// A member found on one member type of a union.
enum UnionMember {
    Value {
        res: Res,
        ty: Type,
        writable: bool,
        owner: ClassId,
        public: bool,
    },
    Method {
        f: FnId,
        env: HashMap<String, Type>,
        owner: ClassId,
        public: bool,
    },
}

pub(crate) struct Checker<'a> {
    pub prog: &'a mut Program,
    module: ModuleId,
    fx: FnCtx,
    in_progress: HashSet<FnId>,
}

/// How a name or member can be used.
enum Access {
    /// Has a value of this type (variable, field, property, enum member).
    Value(Type),
    /// A function or method; `recv` tells how the receiver is passed.
    Func {
        f: FnId,
        env: HashMap<String, Type>,
        recv: Recv,
    },
    Intrinsic(Intr),
    Dynamic,
    Error,
}

#[derive(Clone, Copy, PartialEq)]
enum Recv {
    None,
    Instance,
    Extension,
}

pub fn check(prog: &mut Program) {
    let mut ck = Checker {
        prog,
        module: 0,
        fx: FnCtx::default(),
        in_progress: HashSet::new(),
    };
    for f in 0..ck.prog.fns.len() {
        if matches!(ck.prog.fns[f].kind, FnKind::Main | FnKind::ModInit) {
            ck.ensure_fn(f);
        }
    }
    for c in 0..ck.prog.classes.len() {
        ck.check_class_fields(c);
    }
    for f in 0..ck.prog.fns.len() {
        ck.ensure_fn(f);
    }
    for c in 0..ck.prog.classes.len() {
        ck.check_implements(c);
        ck.check_struct_members(c);
    }
    ck.lints();
}

fn has_named(args: &[Expr]) -> bool {
    args.iter()
        .any(|a| matches!(a.kind, ExprKind::NamedArg { .. }))
}

/// The value of an argument, without its name.
fn named_value(a: &Expr) -> &Expr {
    match &a.kind {
        ExprKind::NamedArg { value, .. } => value,
        _ => a,
    }
}

fn is_lambda(e: &Expr) -> bool {
    matches!(e.peel().kind, ExprKind::Arrow(_))
}

impl Checker<'_> {
    // ------------------------------------------------------------------ plumbing

    fn diag(
        &mut self,
        severity: Severity,
        span: Span,
        code: &'static str,
        message: impl Into<String>,
    ) {
        // Tests only run as JavaScript, where an arrow function is a closure: the hooks and
        // tests of a `describe` share its variables (spec §14.12.3).
        if matches!(code, "E0320" | "E0322") && self.prog.modules[self.module].is_test() {
            return;
        }
        self.prog.diagnostics.push(Diagnostic {
            severity,
            code,
            message: message.into(),
            module: Some(self.module),
            span,
        });
    }

    fn err(&mut self, span: Span, code: &'static str, message: impl Into<String>) {
        self.diag(Severity::Error, span, code, message);
    }

    fn warn(&mut self, span: Span, code: &'static str, message: impl Into<String>) {
        self.diag(Severity::Warning, span, code, message);
    }

    fn show(&self, t: &Type) -> String {
        self.prog.show(t)
    }

    fn info(&mut self) -> &mut TypeInfo {
        &mut self.prog.modules[self.module].info
    }

    fn set_res(&mut self, id: ExprId, res: Res) {
        self.info().res.insert(id, res);
    }

    // ------------------------------------------------------------------ type relations

    pub(crate) fn assignable(&self, from: &Type, to: &Type) -> bool {
        use Type::*;
        if from == to {
            return true;
        }
        match (from, to) {
            (Error, _) | (_, Error) => true,
            (Void, _) | (_, Void) | (Ns(_), _) | (_, Ns(_)) => false,
            // Structs (and arrays of them) cannot be stored in a Variant.
            // (A test only runs as JavaScript, where a struct is an object like any other.)
            (f, Any) | (Any, f) if f.has_struct() => self.prog.modules[self.module].is_test(),
            (_, Any) | (Any, _) => true,
            // `A | B` fits where each of its members fits; a value fits a union when it fits one of the members.
            (Union(fs), _) => fs.iter().all(|f| self.assignable(f, to)),
            (Null, Union(ts)) => ts.iter().any(|t| t.is_object()),
            (_, Union(ts)) => ts.iter().any(|t| self.assignable(from, t)),
            (Null, t) => {
                t.is_object()
                    || matches!(t, Param(p) if self.constraint_of(p) == Constraint::ObjectType)
            }
            (Byte, Short | Int | Long | Float | Double | Currency | Decimal) => true,
            (Short, Int | Long | Float | Double | Currency | Decimal) => true,
            (Int, Long | Float | Double | Currency | Decimal | Ptr) => true,
            (Long, Float | Double | Decimal) => true,
            (Float, Double) => true,
            (Enum(_), Int | Long | Double | Float) => true,
            (Byte | Short | Int, Enum(_)) => true,
            (Str, WStr | AStr) => true,
            // The element type of a host collection (`Excel.Sheets<T>`) is a typing aid: the COM interface is the same.
            (Class(a, _), Class(b, _))
                if a == b && self.prog.classes[*a].kind == ClassKind::Host =>
            {
                true
            }
            (Class(a, aa), Class(b, bb)) => {
                (a == b && aa == bb) || (a != b && self.prog.class_conforms(*a, *b))
            }
            (Class(..) | Map(..) | Func(_), Object) => true,
            (Object, Class(..) | Map(..) | Func(_)) => true,
            (Array(a), Array(b)) => a.is_error() || b.is_error(),
            (Func(a), Func(b)) => {
                a.params.len() <= b.params.len()
                    && a.params
                        .iter()
                        .zip(b.params.iter())
                        .all(|(x, y)| self.assignable(y, x))
                    && (b.ret == Void || self.assignable(&a.ret, &b.ret))
            }
            _ => false,
        }
    }

    fn constraint_of(&self, param: &str) -> Constraint {
        self.fx
            .type_params
            .iter()
            .find(|p| p.name == param)
            .map(|p| p.constraint.clone())
            .unwrap_or(Constraint::None)
    }

    fn satisfies(&self, ty: &Type, c: &Constraint) -> bool {
        if ty.is_error() {
            return true;
        }
        match c {
            Constraint::None => true,
            Constraint::Comparable => {
                (ty.is_numeric() || matches!(ty, Type::Str | Type::Date))
                    || matches!(ty, Type::Param(p) if matches!(self.constraint_of(p), Constraint::Comparable | Constraint::Numeric))
            }
            Constraint::Numeric => {
                ty.is_numeric()
                    || matches!(ty, Type::Param(p) if self.constraint_of(p) == Constraint::Numeric)
            }
            Constraint::ObjectType => {
                (ty.is_object() && *ty != Type::Null)
                    || matches!(ty, Type::Param(p) if self.constraint_of(p) == Constraint::ObjectType)
            }
            Constraint::Type(t) => self.assignable(ty, t),
        }
    }

    fn is_numeric(&self, t: &Type) -> bool {
        t.is_numeric()
            || matches!(t, Type::Any | Type::Error)
            || matches!(t, Type::Param(p) if self.constraint_of(p) == Constraint::Numeric)
    }

    /// C#-style binary numeric promotion (spec §2.3). `None` when the mix needs an explicit cast.
    fn promote(&self, a: &Type, b: &Type) -> Option<Type> {
        use Type::*;
        if a.is_error() || b.is_error() {
            return Some(Error);
        }
        if *a == Any || *b == Any {
            return Some(Any);
        }
        if let (Param(x), Param(y)) = (a, b) {
            return (x == y).then(|| a.clone());
        }
        if let Param(_) = a {
            return self.assignable(b, a).then(|| a.clone());
        }
        if let Param(_) = b {
            return self.assignable(a, b).then(|| b.clone());
        }
        match (a, b) {
            (Currency, Currency) => return Some(Currency),
            (Decimal, Decimal) => return Some(Decimal),
            (Currency, x) | (x, Currency) => return x.is_integer().then_some(Currency),
            (Decimal, x) | (x, Decimal) => return x.is_integer().then_some(Decimal),
            (Ptr, x) | (x, Ptr) => return (x.is_integer() || *x == Ptr).then_some(Ptr),
            (Date, x) | (x, Date) => return (x.is_numeric() || *x == Date).then_some(Date),
            _ => {}
        }
        let (ra, rb) = (a.rank()?, b.rank()?);
        Some(match ra.max(rb) {
            1 => Int,
            2 => Long,
            3 => Float,
            _ => Double,
        })
    }

    fn unify(
        &self,
        pattern: &Type,
        actual: &Type,
        env: &mut HashMap<String, Type>,
        tparams: &[TypeParamDef],
    ) {
        match (pattern, actual) {
            (Type::Param(n), a) if tparams.iter().any(|p| &p.name == n) => {
                if matches!(a, Type::Null | Type::Error | Type::Void | Type::Ns(_)) {
                    return;
                }
                match env.get(n) {
                    None => {
                        env.insert(n.clone(), a.clone());
                    }
                    Some(bound) => {
                        if !self.assignable(a, bound) && self.assignable(bound, a) {
                            env.insert(n.clone(), a.clone());
                        }
                    }
                }
            }
            (Type::Array(p), Type::Array(a)) => self.unify(p, a, env, tparams),
            (Type::Class(c1, pa), Type::Class(c2, aa)) if c1 == c2 => {
                for (p, a) in pa.iter().zip(aa.iter()) {
                    self.unify(p, a, env, tparams);
                }
            }
            (Type::Map(pk, pv), Type::Map(ak, av)) => {
                self.unify(pk, ak, env, tparams);
                self.unify(pv, av, env, tparams);
            }
            (Type::Func(pf), Type::Func(af)) => {
                for (p, a) in pf.params.iter().zip(af.params.iter()) {
                    self.unify(p, a, env, tparams);
                }
                self.unify(&pf.ret, &af.ret, env, tparams);
            }
            _ => {}
        }
    }

    /// Reports an error unless a value of type `from` (expression `e`) can be stored into `to`.
    fn coerce(&mut self, e: &Expr, from: &Type, to: &Type) {
        if self.assignable(from, to) || self.implicit_narrowing(e.span, from, to) {
            return;
        }
        // Integer constants may be stored into narrower integer types when they fit (spec §2.2).
        if let Some(v) = const_int(e) {
            let fits = match to {
                Type::Byte => (0..=255).contains(&v),
                Type::Short => (-32768..=32767).contains(&v),
                Type::Int => (i32::MIN as i128..=i32::MAX as i128).contains(&v),
                _ => false,
            };
            if fits && from.is_integer() {
                return;
            }
        }
        let hint = if from.is_numeric() && to.is_numeric() {
            format!(
                "（縮小変換です。`as {}` で明示的に変換してください）",
                self.show(to)
            )
        } else if matches!(to, Type::Str)
            && (from.is_numeric() || matches!(from, Type::Bool | Type::Date))
        {
            "（`.toString()` かテンプレート文字列 `${x}` で文字列に変換してください）".to_string()
        } else if from.has_struct() && matches!(to, Type::Any | Type::Object) {
            "（構造体は `any` に入れられません。フィールドを個別に渡すか、参照が必要なら `class` を使ってください）".to_string()
        } else if matches!((from, to), (Type::Struct(_), Type::Struct(_))) {
            "（別の構造体型です。リテラルなら `const p: T = { ... }` のように型注釈を付け、そうでなければフィールドを個別に代入してください）".to_string()
        } else {
            String::new()
        };
        self.err(
            e.span,
            "E0301",
            format!(
                "型 `{}` を `{}` に代入できません{hint}",
                self.show(from),
                self.show(to)
            ),
        );
    }

    /// A union stored where only some of its members fit (`const ws: Excel.Worksheet = activeSheet`):
    /// allowed like VBA's `Set ws = ActiveSheet`, which checks at run time, with a warning (spec §2.6.2).
    fn implicit_narrowing(&mut self, span: Span, from: &Type, to: &Type) -> bool {
        let Some(ms) = from.union_members() else {
            return false;
        };
        if to.has_struct() || matches!(to, Type::Void | Type::Ns(_)) {
            return false;
        }
        let rest: Vec<String> = ms
            .iter()
            .filter(|m| !self.assignable(m, to))
            .map(|m| self.show(m))
            .collect();
        if rest.len() == ms.len() {
            return false;
        }
        let what = if to.is_object() {
            "実行時エラー（型が一致しません）になります"
        } else {
            "VBA の暗黙の型変換になり、変換できなければ実行時エラーになります"
        };
        let (shown_from, shown_to) = (self.show(from), self.show(to));
        self.warn(
            span,
            "W0306",
            format!("`{shown_from}` を `{shown_to}` として扱います。値が `{}` のときは{what}（`instanceof` / `typeof` で絞り込むか、`as {shown_to}` で変換してください）", rest.join("` / `")),
        );
        true
    }

    fn unify_branches(&mut self, span: Span, a: &Type, b: &Type) -> Type {
        if self.assignable(b, a) && *a != Type::Null {
            a.clone()
        } else if self.assignable(a, b) {
            b.clone()
        } else if let Some(p) = self.promote(a, b) {
            p
        } else {
            let unionable = |t: &Type| {
                t.is_value()
                    && !t.has_struct()
                    && !t.contains_param()
                    && !matches!(t, Type::Null | Type::WStr | Type::AStr)
            };
            let hint = if unionable(a) && unionable(b) {
                format!(
                    "（どちらの型も受けるには `{}` のようにユニオン型の型注釈を付けてください）",
                    self.show(&Type::union([a.clone(), b.clone()]))
                )
            } else {
                String::new()
            };
            self.err(
                span,
                "E0302",
                format!(
                    "型が一致しません: `{}` と `{}`{hint}",
                    self.show(a),
                    self.show(b)
                ),
            );
            Type::Error
        }
    }

    // ------------------------------------------------------------------ scopes

    fn push_scope(&mut self) {
        self.fx.scopes.push(Scope::default());
    }

    fn pop_scope(&mut self) {
        self.fx.scopes.pop();
    }

    fn declare_local(&mut self, name: &Ident, ty: Type, kind: LocalKind) -> LocalId {
        let id = self.fx.locals.len();
        self.fx.locals.push(LocalDef {
            name: name.text.clone(),
            ty,
            kind,
            byref: false,
            reassigned: false,
            captured: false,
            span: name.span,
            used: false,
        });
        if name.text.is_empty() {
            return id;
        }
        let dup = self
            .fx
            .scopes
            .last()
            .is_some_and(|s| s.names.contains_key(&name.text));
        if dup {
            self.err(
                name.span,
                "E0101",
                format!("`{}` はこのスコープで既に宣言されています", name.text),
            );
        }
        if let Some(s) = self.fx.scopes.last_mut() {
            s.names.insert(name.text.clone(), id);
        }
        let offset = u32::from(name.span.start());
        self.info().decl_locals.insert(offset, id);
        id
    }

    /// Finds a local; also returns the arrow functions whose boundary the reference crosses.
    fn lookup_local(&self, name: &str) -> Option<(LocalId, Vec<ExprId>)> {
        let mut crossed = Vec::new();
        for scope in self.fx.scopes.iter().rev() {
            if let Some(&id) = scope.names.get(name) {
                return Some((id, crossed));
            }
            if let Some(l) = scope.lambda {
                crossed.push(l);
            }
        }
        None
    }

    fn lookup_sym(&self, name: &str) -> Option<Sym> {
        self.prog.modules[self.module]
            .scope
            .get(name)
            .or_else(|| self.prog.globals.get(name))
            .cloned()
    }

    fn current_lambda(&self) -> Option<ExprId> {
        self.fx.scopes.iter().rev().find_map(|s| s.lambda)
    }

    // ------------------------------------------------------------------ functions

    pub(crate) fn ensure_fn(&mut self, f: FnId) {
        if self.prog.fns[f].checked || self.in_progress.contains(&f) {
            return;
        }
        let Some(decl) = self.prog.fns[f].decl.clone() else {
            self.prog.fns[f].checked = true;
            return;
        };
        let Some(body) = &decl.body else {
            self.prog.fns[f].checked = true;
            return;
        };
        self.in_progress.insert(f);
        let saved_module = self.module;
        let saved_fx = std::mem::take(&mut self.fx);
        let def = self.prog.fns[f].clone();
        self.module = def.module;
        let (class, is_static) = match def.kind {
            FnKind::Method { class, is_static }
            | FnKind::Getter { class, is_static }
            | FnKind::Setter { class, is_static } => (Some(class), is_static),
            FnKind::Ctor { class } => (Some(class), false),
            // A form's code-behind: `this` is the form (spec §14.11.2).
            FnKind::Free | FnKind::ModInit if !def.is_extension => (
                self.prog.modules[def.module]
                    .form
                    .as_ref()
                    .and_then(|f| f.class),
                false,
            ),
            _ => (None, false),
        };
        let mut type_params = class
            .map(|c| self.prog.classes[c].type_params.clone())
            .unwrap_or_default();
        type_params.extend(def.type_params.iter().cloned());
        self.fx = FnCtx {
            fn_id: Some(f),
            class,
            is_static,
            type_params,
            locals: def.locals.clone(),
            ret: def.ret.clone(),
            is_toplevel: matches!(def.kind, FnKind::Main | FnKind::ModInit),
            ..Default::default()
        };
        self.push_scope();
        for (i, p) in decl.params.iter().enumerate() {
            if p.is_this {
                self.fx.this_local = Some(i);
            } else if !p.name.text.is_empty()
                && self.fx.scopes[0]
                    .names
                    .insert(p.name.text.clone(), i)
                    .is_some()
            {
                self.err(
                    p.name.span,
                    "E0101",
                    format!("引数 `{}` が重複しています", p.name.text),
                );
            }
            let offset = u32::from(p.name.span.start());
            self.info().decl_locals.insert(offset, i);
        }
        for (i, p) in decl.params.iter().enumerate() {
            if let Some(d) = &p.default {
                let pt = def.params[i].ty.clone();
                let t = self.value(d, Some(&pt));
                self.coerce(d, &t, &pt);
            }
        }
        self.stmts(&body.stmts);
        self.pop_scope();

        let ret = match &def.ret {
            Some(t) => t.clone(),
            None => self.fx.inferred.clone().unwrap_or(Type::Void),
        };
        if ret != Type::Void
            && !ret.is_error()
            && !matches!(def.kind, FnKind::Ctor { .. })
            && !always_returns(&body.stmts)
        {
            self.warn(
                def.name_span,
                "W0310",
                format!(
                    "`{}`: 値を返さずに終了する経路があります（その場合は型の既定値が返ります）",
                    def.name
                ),
            );
        }
        for (local, span) in std::mem::take(&mut self.fx.captures) {
            let l = &self.fx.locals[local];
            if l.reassigned && l.kind != LocalKind::Const {
                let name = l.name.clone();
                self.err(span, "E0320", format!("アロー関数から参照できるのは再代入されない変数だけです（`{name}` は再代入されています）"));
            }
        }
        for (_, span) in std::mem::take(&mut self.fx.row_views) {
            self.err(span, "E0321", "2次元配列は `a[行][列]` の形で要素を参照してください（行だけを取り出すことはできません）");
        }
        let fx = std::mem::replace(&mut self.fx, saved_fx);
        let d = &mut self.prog.fns[f];
        d.locals = fx.locals;
        d.ret = Some(ret);
        d.calls = fx.calls;
        d.var_refs = fx.var_refs;
        d.has_loops = fx.has_loops;
        d.inline_blockers = fx.blockers;
        d.cost = fx.cost;
        d.checked = true;
        self.module = saved_module;
        self.in_progress.remove(&f);
    }

    /// Return type of a function, checking its body first when it has to be inferred.
    fn fn_ret(&mut self, f: FnId, span: Span) -> Type {
        if self.prog.fns[f].ret.is_none() {
            if self.in_progress.contains(&f) {
                let name = self.prog.fns[f].name.clone();
                self.err(
                    span,
                    "E0330",
                    format!("再帰する関数 `{name}` には戻り値の型注釈が必要です"),
                );
                return Type::Error;
            }
            self.ensure_fn(f);
        }
        self.prog.fns[f].ret_ty()
    }

    fn var_type(&mut self, v: VarId, span: Span) -> Type {
        if self.prog.vars[v].ty == Type::Void {
            // Inferred from the initializer, which runs in the module's Main / initializer.
            let m = self.prog.vars[v].module;
            let init = self.prog.modules[m]
                .main_fn
                .or(self.prog.modules[m].init_fn);
            match init {
                Some(f) if !self.in_progress.contains(&f) => self.ensure_fn(f),
                _ => {}
            }
            if self.prog.vars[v].ty == Type::Void {
                let name = self.prog.vars[v].name.clone();
                self.err(span, "E0331", format!("`{name}` の型を推論できません（初期化より前に参照されています）。型注釈を付けてください"));
                self.prog.vars[v].ty = Type::Error;
            }
        }
        self.prog.vars[v].ty.clone()
    }

    fn check_class_fields(&mut self, c: ClassId) {
        let class = self.prog.classes[c].clone();
        if class.kind != ClassKind::Class {
            return;
        }
        self.module = class.module;
        for (i, field) in class.fields.iter().enumerate() {
            let Some(init) = &field.init else { continue };
            self.fx = FnCtx {
                class: Some(c),
                is_static: field.is_static,
                type_params: class.type_params.clone(),
                in_field_init: true,
                ..Default::default()
            };
            self.push_scope();
            let expect = (field.ty != Type::Void).then(|| field.ty.clone());
            let t = self.value(init, expect.as_ref());
            match expect {
                Some(e) => self.coerce(init, &t, &e),
                None => {
                    if t == Type::Null {
                        self.err(
                            init.span,
                            "E0132",
                            "`null` で初期化するフィールドには型注釈が必要です",
                        );
                    }
                    self.prog.classes[c].fields[i].ty =
                        if t == Type::Null { Type::Error } else { t };
                }
            }
            self.fx = FnCtx::default();
        }
    }

    /// VBA class modules cannot expose a `Type` from a standard module: struct-typed
    /// fields must be private, and interfaces cannot mention structs at all.
    fn check_struct_members(&mut self, c: ClassId) {
        let class = self.prog.classes[c].clone();
        self.module = class.module;
        match class.kind {
            ClassKind::Class => {
                for f in &class.fields {
                    if f.ty.has_struct() && f.vis == ast::Visibility::Public && !f.is_static {
                        self.err(f.span, "E0394", format!("構造体型のフィールド `{}` は `private` にしてください（VBA のクラスは構造体のフィールドを公開できません。メソッドで受け渡してください）", f.name));
                    }
                }
            }
            ClassKind::Interface => {
                for &m in &class.methods {
                    let d = &self.prog.fns[m];
                    if d.params.iter().any(|p| p.ty.has_struct()) || d.ret_ty().has_struct() {
                        let (span, name) = (d.name_span, d.name.clone());
                        self.err(span, "E0394", format!("インターフェイスのメソッド `{name}` の引数・戻り値に構造体は使えません（VBA の制限です）"));
                    }
                }
                for f in &class.fields {
                    if f.ty.has_struct() {
                        self.err(f.span, "E0394", format!("インターフェイスのフィールド `{}` に構造体は使えません（VBA の制限です）", f.name));
                    }
                }
            }
            ClassKind::Host => {}
        }
    }

    fn check_implements(&mut self, c: ClassId) {
        let class = self.prog.classes[c].clone();
        if class.kind != ClassKind::Class {
            return;
        }
        self.module = class.module;
        for iface in &class.implements {
            let Type::Class(i, iargs) = iface else {
                continue;
            };
            let env = self.prog.class_env(*i, iargs);
            let idef = self.prog.classes[*i].clone();
            for &m in &idef.methods {
                let want = self.prog.fns[m].clone();
                let found = class.methods.iter().copied().find(|&f| {
                    let d = &self.prog.fns[f];
                    d.name == want.name
                        && matches!(
                            d.kind,
                            FnKind::Method {
                                is_static: false,
                                ..
                            }
                        )
                });
                match found {
                    None => self.err(class.name_span, "E0340", format!("クラス `{}` はインターフェイス `{}` のメソッド `{}` を実装していません", class.name, idef.name, want.name)),
                    Some(f) => {
                        let have = self.prog.fns[f].clone();
                        let same = have.params.len() == want.params.len()
                            && have.params.iter().zip(want.params.iter()).all(|(a, b)| a.ty == b.ty.subst(&env) && a.mode == b.mode)
                            && have.ret_ty() == want.ret_ty().subst(&env);
                        if !same {
                            self.err(have.name_span, "E0341", format!("`{}` のシグネチャがインターフェイス `{}` と一致しません", have.name, idef.name));
                        }
                        if have.vis != Visibility::Public {
                            self.err(have.name_span, "E0341", format!("インターフェイスを実装するメソッド `{}` は public である必要があります", have.name));
                        }
                    }
                }
            }
            for field in &idef.fields {
                let ok = class.fields.iter().any(|f| {
                    f.name == field.name
                        && f.ty == field.ty.subst(&env)
                        && f.vis == Visibility::Public
                }) || class.methods.iter().any(|&f| {
                    self.prog.fns[f].name == field.name
                        && matches!(self.prog.fns[f].kind, FnKind::Getter { .. })
                });
                if !ok {
                    self.err(class.name_span, "E0340", format!("クラス `{}` はインターフェイス `{}` のプロパティ `{}` を実装していません", class.name, idef.name, field.name));
                }
            }
        }
    }

    // ------------------------------------------------------------------ statements

    fn block(&mut self, b: &ast::Block) {
        self.push_scope();
        self.stmts(&b.stmts);
        self.pop_scope();
    }

    /// A statement list. After `if (!(x instanceof T)) return;` the rest of the list sees `x` as `T`.
    fn stmts(&mut self, stmts: &[Stmt]) {
        let mark = self.fx.narrowed.len();
        for (i, s) in stmts.iter().enumerate() {
            self.stmt(s);
            if let StmtKind::If { cond, then, els } = &s.kind {
                let then_exits = stmt_exits(then);
                let else_exits = els.as_ref().is_some_and(|e| stmt_exits(e));
                if then_exits != else_exits {
                    let (t, f) = self.guard_facts(cond, &Vec::new());
                    let facts = if then_exits { f } else { t };
                    let rest = &stmts[i + 1..];
                    let facts = self.stable_facts(facts, &mut |name| {
                        rest.iter().any(|r| assigns_in_stmt(r, name))
                    });
                    self.fx.narrowed.extend(facts);
                }
            }
        }
        self.fx.narrowed.truncate(mark);
    }

    /// Runs `f` with `facts` in effect (those whose local `region` does not assign).
    fn with_facts<R>(
        &mut self,
        facts: Facts,
        assigned: &mut dyn FnMut(&str) -> bool,
        f: impl FnOnce(&mut Self) -> R,
    ) -> R {
        let facts = self.stable_facts(facts, assigned);
        let mark = self.fx.narrowed.len();
        self.fx.narrowed.extend(facts);
        let r = f(self);
        self.fx.narrowed.truncate(mark);
        r
    }

    /// Drops the facts about locals that are assigned in the region they would apply to.
    fn stable_facts(&self, facts: Facts, assigned: &mut dyn FnMut(&str) -> bool) -> Facts {
        facts
            .into_iter()
            .filter(|(l, _)| !assigned(&self.fx.locals[*l].name))
            .collect()
    }

    // ------------------------------------------------------------------ narrowing (spec §2.6.4)

    /// The current type of a local: its declared type unless narrowed here.
    fn local_ty(&self, l: LocalId) -> Type {
        self.fx
            .narrowed
            .iter()
            .rev()
            .find(|(x, _)| *x == l)
            .map(|(_, t)| t.clone())
            .unwrap_or_else(|| self.fx.locals[l].ty.clone())
    }

    fn fact_ty(&self, env: &Facts, l: LocalId) -> Type {
        env.iter()
            .rev()
            .find(|(x, _)| *x == l)
            .map(|(_, t)| t.clone())
            .unwrap_or_else(|| self.local_ty(l))
    }

    /// A local variable or parameter that a type guard can narrow.
    fn guarded_local(&self, e: &Expr) -> Option<LocalId> {
        let inner = e.peel();
        if !matches!(inner.kind, ExprKind::Name(_)) {
            return None;
        }
        match self.prog.modules[self.module].info.res.get(&inner.id) {
            Some(Res::Local(l))
                if self
                    .fx
                    .locals
                    .get(*l)
                    .is_some_and(|d| d.kind != LocalKind::This) =>
            {
                Some(*l)
            }
            _ => None,
        }
    }

    /// What `cond` being true / false tells about the types of locals (`env`: facts already known).
    fn guard_facts(&self, cond: &Expr, env: &Facts) -> (Facts, Facts) {
        let none = || (Vec::new(), Vec::new());
        match &cond.kind {
            ExprKind::Paren(inner) => self.guard_facts(inner, env),
            ExprKind::Unary {
                op: UnaryOp::Not,
                operand,
            } => {
                let (t, f) = self.guard_facts(operand, env);
                (f, t)
            }
            ExprKind::Binary {
                op: BinaryOp::And,
                lhs,
                rhs,
            } => {
                let (ta, _) = self.guard_facts(lhs, env);
                let env2: Facts = env.iter().cloned().chain(ta.iter().cloned()).collect();
                let (tb, _) = self.guard_facts(rhs, &env2);
                (ta.into_iter().chain(tb).collect(), Vec::new())
            }
            ExprKind::Binary {
                op: BinaryOp::Or,
                lhs,
                rhs,
            } => {
                let (ta, fa) = self.guard_facts(lhs, env);
                let env2: Facts = env.iter().cloned().chain(fa.iter().cloned()).collect();
                let (tb, fb) = self.guard_facts(rhs, &env2);
                // True when `lhs` is true, or when `lhs` is false and `rhs` is true.
                let mut locals: Vec<LocalId> = Vec::new();
                for (l, _) in ta.iter().chain(tb.iter()) {
                    if !locals.contains(l) {
                        locals.push(*l);
                    }
                }
                let t = locals
                    .into_iter()
                    .map(|l| {
                        let a = ta
                            .iter()
                            .rev()
                            .find(|(x, _)| *x == l)
                            .map(|(_, t)| t.clone())
                            .unwrap_or_else(|| self.fact_ty(env, l));
                        let b = tb
                            .iter()
                            .rev()
                            .find(|(x, _)| *x == l)
                            .map(|(_, t)| t.clone())
                            .unwrap_or_else(|| self.fact_ty(&env2, l));
                        (l, Type::union([a, b]))
                    })
                    .collect();
                (t, fa.into_iter().chain(fb).collect())
            }
            ExprKind::InstanceOf { expr, .. } => {
                let (Some(l), Some(target)) = (
                    self.guarded_local(expr),
                    self.fx.instanceof_types.get(&cond.id),
                ) else {
                    return none();
                };
                if !matches!(target, Type::Class(..)) {
                    return none();
                }
                let cur = self.fact_ty(env, l);
                let t = vec![(l, self.narrow_to(&cur, target))];
                let f = self
                    .narrow_away(&cur, |m| self.is_subtype(m, target))
                    .map(|ty| vec![(l, ty)])
                    .unwrap_or_default();
                (t, f)
            }
            ExprKind::Binary {
                op: op @ (BinaryOp::Eq | BinaryOp::Ne),
                lhs,
                rhs,
            } => {
                let guard = |a: &Expr, b: &Expr| match (&a.peel().kind, &b.peel().kind) {
                    (
                        ExprKind::Unary {
                            op: UnaryOp::TypeOf,
                            operand,
                        },
                        ExprKind::Str(name),
                    ) => self.guarded_local(operand).map(|l| (l, name.clone())),
                    _ => None,
                };
                let Some((l, name)) = guard(lhs, rhs).or_else(|| guard(rhs, lhs)) else {
                    return none();
                };
                let cur = self.fact_ty(env, l);
                let t = self
                    .narrow_typeof(&cur, &name)
                    .map(|ty| vec![(l, ty)])
                    .unwrap_or_default();
                let f = self
                    .narrow_away(&cur, |m| {
                        self.type_name_of(m)
                            .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                    })
                    .map(|ty| vec![(l, ty)])
                    .unwrap_or_default();
                if *op == BinaryOp::Eq {
                    (t, f)
                } else {
                    (f, t)
                }
            }
            _ => none(),
        }
    }

    /// `a` is always an instance of `b`.
    fn is_subtype(&self, a: &Type, b: &Type) -> bool {
        match (a, b) {
            (_, Type::Object) => a.is_object(),
            (Type::Class(x, xa), Type::Class(y, ya)) => {
                if x == y {
                    xa == ya || self.prog.classes[*x].kind == ClassKind::Host
                } else {
                    self.prog.class_conforms(*x, *y)
                }
            }
            _ => a == b,
        }
    }

    /// The type of a value of type `cur` for which `instanceof target` is true.
    fn narrow_to(&self, cur: &Type, target: &Type) -> Type {
        let members: Vec<Type> = match cur {
            Type::Union(ms) => ms.clone(),
            t => vec![t.clone()],
        };
        let mut kept = Vec::new();
        for m in members {
            if self.is_subtype(&m, target) {
                kept.push(m);
            } else if self.is_subtype(target, &m) || matches!(m, Type::Any) {
                kept.push(target.clone());
            }
        }
        if kept.is_empty() {
            target.clone()
        } else {
            Type::union(kept)
        }
    }

    /// A union without the members for which `drop` holds (`None` unless something is dropped and something remains).
    fn narrow_away(&self, cur: &Type, drop: impl Fn(&Type) -> bool) -> Option<Type> {
        let ms = cur.union_members()?;
        let rest: Vec<Type> = ms.iter().filter(|m| !drop(m)).cloned().collect();
        (rest.len() < ms.len() && !rest.is_empty()).then(|| Type::union(rest))
    }

    /// `typeof x == "String"`: the members of `cur` whose values have that `TypeName`.
    fn narrow_typeof(&self, cur: &Type, name: &str) -> Option<Type> {
        match cur {
            Type::Union(ms) => {
                let kept: Vec<Type> = ms
                    .iter()
                    .filter(|m| {
                        self.type_name_of(m)
                            .is_some_and(|n| n.eq_ignore_ascii_case(name))
                    })
                    .cloned()
                    .collect();
                (!kept.is_empty()).then(|| Type::union(kept))
            }
            Type::Any => type_of_type_name(name),
            _ => None,
        }
    }

    /// What VBA's `TypeName` (RAS `typeof`) returns for a value of type `t`, when it is fixed.
    fn type_name_of(&self, t: &Type) -> Option<String> {
        Some(match t {
            // Only a host type's name is fixed: a RAS class may be renamed in VBA (`Circle` → `Circle_`).
            Type::Class(c, _) if self.prog.classes[*c].kind == ClassKind::Host => {
                self.prog.classes[*c].name.clone()
            }
            Type::Class(..) => return None,
            Type::Map(..) => "Dictionary".into(),
            Type::Enum(_) => "Long".into(),
            Type::Decimal => "Decimal".into(),
            // `TypeName` of an array: the element type and `()` (`"Variant()"`, `"String()"`).
            Type::Array(e) if !e.is_array() => {
                let elem = match &**e {
                    Type::Any | Type::Decimal | Type::Error | Type::Union(_) => {
                        "Variant".to_string()
                    }
                    Type::Object | Type::Map(..) | Type::Func(_) => "Object".to_string(),
                    other => self.type_name_of(other)?,
                };
                format!("{elem}()")
            }
            other => crate::query::vba_type(other)
                .filter(|n| !matches!(*n, "Variant" | "Object" | "LongPtr"))?
                .to_string(),
        })
    }

    fn scoped_stmt(&mut self, s: &Stmt) {
        match &s.kind {
            StmtKind::Block(b) => self.block(b),
            _ => {
                self.push_scope();
                self.stmt(s);
                self.pop_scope();
            }
        }
    }

    fn cond(&mut self, e: &Expr) {
        let t = self.value(e, Some(&Type::Bool));
        if !matches!(t, Type::Bool | Type::Any | Type::Error) {
            self.err(e.span, "E0303", format!("条件式は `boolean` である必要があります（`{}` です）。比較演算子で明示してください", self.show(&t)));
        }
    }

    fn loop_decorators(&mut self, decorators: &[ast::Decorator]) {
        for d in decorators {
            if !matches!(
                d.name.text.as_str(),
                "bulk" | "nobulk" | "doevents" | "nodoevents"
            ) {
                self.warn(
                    d.span,
                    "W0110",
                    format!("不明なループデコレータ `@{}` は無視されます", d.name.text),
                );
            }
        }
    }

    fn stmt(&mut self, s: &Stmt) {
        self.fx.cost += 1;
        match &s.kind {
            StmtKind::Empty | StmtKind::Error => {}
            StmtKind::Var(v) => self.var_stmt(v),
            StmtKind::Expr(e) => {
                let t = self.expr(e, None);
                let effectful = matches!(
                    e.peel().kind,
                    ExprKind::Call { .. }
                        | ExprKind::Assign { .. }
                        | ExprKind::IncDec { .. }
                        | ExprKind::New { .. }
                        | ExprKind::Error
                );
                if !effectful && !t.is_error() {
                    self.warn(e.span, "W0311", "この式の値は使われていません");
                }
            }
            StmtKind::Block(b) => self.block(b),
            StmtKind::If { cond, then, els } => {
                self.cond(cond);
                let (t, f) = self.guard_facts(cond, &Vec::new());
                self.with_facts(t, &mut |n| assigns_in_stmt(then, n), |ck| {
                    ck.scoped_stmt(then)
                });
                if let Some(e) = els {
                    self.with_facts(f, &mut |n| assigns_in_stmt(e, n), |ck| ck.scoped_stmt(e));
                }
            }
            StmtKind::For {
                decorators,
                init,
                cond,
                update,
                body,
            } => {
                self.loop_decorators(decorators);
                self.fx.has_loops = true;
                self.push_scope();
                for i in init {
                    self.stmt(i);
                }
                if let Some(c) = cond {
                    self.cond(c);
                }
                for u in update {
                    self.expr(u, None);
                }
                let facts = cond
                    .as_ref()
                    .map(|c| self.guard_facts(c, &Vec::new()).0)
                    .unwrap_or_default();
                self.fx.loops.push(None);
                self.with_facts(
                    facts,
                    &mut |n| {
                        assigns_in_stmt(body, n) || update.iter().any(|u| assigns_in_expr(u, n))
                    },
                    |ck| ck.scoped_stmt(body),
                );
                self.fx.loops.pop();
                self.pop_scope();
            }
            StmtKind::ForOf {
                decorators,
                kind,
                name,
                ty,
                iter,
                body,
            } => {
                self.loop_decorators(decorators);
                self.fx.has_loops = true;
                let it = self.value(iter, None);
                let elem = match &it {
                    Type::Array(_) if it.dims() == 2 => {
                        self.err(iter.span, "E0304", "2次元配列は `for...of` で列挙できません。添字の `for` を使ってください");
                        Type::Error
                    }
                    Type::Array(e) => (**e).clone(),
                    Type::Any | Type::Object => Type::Any,
                    Type::Class(c, args) if self.prog.classes[*c].kind == ClassKind::Host => {
                        match self.prog.classes[*c].call_sig {
                            Some(sig) => {
                                let env = self.prog.class_env(*c, args);
                                self.prog.fns[sig].ret_ty().subst(&env)
                            }
                            None => Type::Any,
                        }
                    }
                    Type::Error => Type::Error,
                    other => {
                        let shown = self.show(other);
                        let hint = if matches!(other, Type::Map(..)) {
                            "（`map.keys()` を列挙してください）"
                        } else {
                            ""
                        };
                        self.err(
                            iter.span,
                            "E0304",
                            format!("`{shown}` は `for...of` で列挙できません{hint}"),
                        );
                        Type::Error
                    }
                };
                let var_ty = match ty {
                    Some(t) => {
                        let tps = self.fx.type_params.clone();
                        let declared = resolve_type(self.prog, self.module, &tps, t);
                        if !self.assignable(&elem, &declared)
                            && !self.implicit_narrowing(t.span(), &elem, &declared)
                        {
                            self.err(
                                t.span(),
                                "E0301",
                                format!(
                                    "要素の型 `{}` を `{}` に代入できません",
                                    self.show(&elem),
                                    self.show(&declared)
                                ),
                            );
                        }
                        declared
                    }
                    None => elem,
                };
                self.push_scope();
                let _ = kind;
                self.declare_local(name, var_ty, LocalKind::Iter);
                self.fx.loops.push(None);
                self.scoped_stmt(body);
                self.fx.loops.pop();
                self.pop_scope();
            }
            StmtKind::While {
                decorators,
                cond,
                body,
            } => {
                self.loop_decorators(decorators);
                self.fx.has_loops = true;
                self.cond(cond);
                let (t, _) = self.guard_facts(cond, &Vec::new());
                self.fx.loops.push(None);
                self.with_facts(t, &mut |n| assigns_in_stmt(body, n), |ck| {
                    ck.scoped_stmt(body)
                });
                self.fx.loops.pop();
            }
            StmtKind::DoWhile {
                decorators,
                body,
                cond,
            } => {
                self.loop_decorators(decorators);
                self.fx.has_loops = true;
                self.fx.loops.push(None);
                self.scoped_stmt(body);
                self.fx.loops.pop();
                self.cond(cond);
            }
            StmtKind::Switch { disc, cases } => {
                let dt = self.value(disc, None);
                self.fx.switch_depth += 1;
                let mut seen_default = false;
                for c in cases {
                    match &c.test {
                        Some(t) => {
                            let tt = self.value(t, Some(&dt));
                            if !self.comparable(&dt, &tt, true) {
                                self.err(t.span, "E0305", format!("`case` の値（`{}`）を `switch` の値（`{}`）と比較できません", self.show(&tt), self.show(&dt)));
                            }
                        }
                        None => {
                            if seen_default {
                                self.err(c.span, "E0306", "`default` は 1 つだけ書けます");
                            }
                            seen_default = true;
                        }
                    }
                    self.push_scope();
                    self.stmts(&c.body);
                    self.pop_scope();
                }
                self.fx.switch_depth -= 1;
            }
            StmtKind::Try(t) => {
                self.block(&t.block);
                if let Some(c) = &t.catch {
                    if c.contains_resume() {
                        // spec §5.2 / §6 rule 4: `catch { resume; }` must contain nothing else.
                        if !c.is_resume_block() {
                            self.err(
                                c.span,
                                "E0350",
                                "`resume;` を書いた `catch` ブロックには他の文を書けません",
                            );
                        }
                        if c.param.is_some() {
                            self.err(
                                c.span,
                                "E0350",
                                "`catch { resume; }` ではエラー変数を受け取れません",
                            );
                        }
                        if t.finally.is_some() {
                            self.err(
                                c.span,
                                "E0350",
                                "`catch { resume; }` に `finally` は付けられません",
                            );
                        }
                        self.fx.blockers.push("resume");
                    } else {
                        self.push_scope();
                        if let Some(p) = &c.param {
                            let ty = self.error_type();
                            self.declare_local(p, ty, LocalKind::Catch);
                        }
                        self.stmts(&c.body.stmts);
                        self.pop_scope();
                    }
                }
                if let Some(f) = &t.finally {
                    self.block(f);
                }
            }
            StmtKind::Throw(e) => {
                let t = self.value(e, None);
                let err_ty = self.error_type();
                if !(t == err_ty || matches!(t, Type::Str | Type::Any | Type::Error)) {
                    self.err(e.span, "E0351", "`throw` できるのは `new Error(\"...\")`・`catch` で受け取ったエラー・文字列のいずれかです");
                }
            }
            StmtKind::Return(value) => {
                if self.fx.in_field_init {
                    return;
                }
                if let Some(lr) = self.fx.lambda_rets.last() {
                    let expected = lr.expected.clone();
                    let t = match value {
                        Some(v) => self.value(v, expected.as_ref()),
                        None => Type::Void,
                    };
                    match (&expected, value) {
                        (Some(exp), Some(v)) if *exp != Type::Void => self.coerce(v, &t, exp),
                        _ => {}
                    }
                    if expected.is_none() {
                        let prev = self.fx.lambda_rets.last().unwrap().inferred.clone();
                        let merged = match prev {
                            Some(p) => self.unify_branches(s.span, &p, &t),
                            None => t,
                        };
                        self.fx.lambda_rets.last_mut().unwrap().inferred = Some(merged);
                    }
                    return;
                }
                let declared = self.fx.ret.clone();
                match (declared, value) {
                    (Some(Type::Void), Some(v)) => {
                        let t = self.expr(v, None);
                        if t != Type::Void && !t.is_error() {
                            self.err(
                                v.span,
                                "E0307",
                                "戻り値のない関数（`void`）から値は返せません",
                            );
                        }
                    }
                    (Some(d), Some(v)) => {
                        let t = self.value(v, Some(&d));
                        self.coerce(v, &t, &d);
                    }
                    (Some(d), None) => {
                        if d != Type::Void && !d.is_error() {
                            self.err(
                                s.span,
                                "E0307",
                                format!("`{}` 型の値を返す必要があります", self.show(&d)),
                            );
                        }
                    }
                    (None, Some(v)) => {
                        let t = self.value(v, None);
                        let merged = match self.fx.inferred.clone() {
                            Some(p) => self.unify_branches(v.span, &p, &t),
                            None => {
                                if t == Type::Null {
                                    self.err(
                                        v.span,
                                        "E0308",
                                        "`null` を返す関数には戻り値の型注釈が必要です",
                                    );
                                    Type::Error
                                } else {
                                    t
                                }
                            }
                        };
                        self.fx.inferred = Some(merged);
                    }
                    (None, None) => {
                        if self.fx.inferred.is_none() {
                            self.fx.inferred = Some(Type::Void);
                        }
                    }
                }
            }
            StmtKind::Break(label) => {
                if self.current_lambda().is_some() {
                    self.err(
                        s.span,
                        "E0309",
                        "アロー関数の中から外側のループを `break` できません",
                    );
                } else if let Some(l) = label {
                    if !self
                        .fx
                        .loops
                        .iter()
                        .any(|x| x.as_deref() == Some(l.text.as_str()))
                    {
                        self.err(
                            l.span,
                            "E0309",
                            format!("ラベル `{}` の付いたループがありません", l.text),
                        );
                    }
                } else if self.fx.loops.is_empty() && self.fx.switch_depth == 0 {
                    self.err(
                        s.span,
                        "E0309",
                        "`break` はループか `switch` の中でのみ使えます",
                    );
                }
            }
            StmtKind::Continue(label) => {
                if self.current_lambda().is_some() {
                    self.err(
                        s.span,
                        "E0309",
                        "アロー関数の中から外側のループを `continue` できません",
                    );
                } else if let Some(l) = label {
                    if !self
                        .fx
                        .loops
                        .iter()
                        .any(|x| x.as_deref() == Some(l.text.as_str()))
                    {
                        self.err(
                            l.span,
                            "E0309",
                            format!("ラベル `{}` の付いたループがありません", l.text),
                        );
                    }
                } else if self.fx.loops.is_empty() {
                    self.err(s.span, "E0309", "`continue` はループの中でのみ使えます");
                }
            }
            StmtKind::Labeled { label, body } => {
                let is_loop = matches!(
                    body.kind,
                    StmtKind::For { .. }
                        | StmtKind::ForOf { .. }
                        | StmtKind::While { .. }
                        | StmtKind::DoWhile { .. }
                );
                if !is_loop {
                    self.err(label.span, "E0309", "ラベルはループにのみ付けられます");
                    self.stmt(body);
                    return;
                }
                if self
                    .fx
                    .loops
                    .iter()
                    .any(|x| x.as_deref() == Some(label.text.as_str()))
                {
                    self.err(
                        label.span,
                        "E0101",
                        format!("ラベル `{}` は既に使われています", label.text),
                    );
                }
                // The loop statement pushes its own (unlabelled) entry; label it afterwards.
                self.labeled_loop(label, body);
            }
            StmtKind::Resume => {
                self.err(
                    s.span,
                    "E0350",
                    "`resume;` は `catch { resume; }` の形でのみ使えます",
                );
            }
            StmtKind::Fn(f) => {
                self.err(f.name.span, "E0352", "関数の中に関数は宣言できません。アロー関数を変数に代入するか、モジュールの最上位に宣言してください");
            }
        }
    }

    fn labeled_loop(&mut self, label: &Ident, body: &Stmt) {
        // Re-implements the loop push so that the entry carries the label.
        let depth = self.fx.loops.len();
        self.fx.loops.push(Some(label.text.clone()));
        // Mark: the nested statement will push another entry; keep ours below it.
        self.stmt(body);
        self.fx.loops.truncate(depth);
    }

    fn error_type(&mut self) -> Type {
        match self.prog.globals.get("RAS_Error") {
            Some(Sym::Class(c)) => Type::Class(*c, Vec::new()),
            _ => Type::Any,
        }
    }

    fn var_stmt(&mut self, v: &ast::VarDecl) {
        let tps = self.fx.type_params.clone();
        let declared =
            v.ty.as_ref()
                .map(|t| resolve_type(self.prog, self.module, &tps, t));
        if declared == Some(Type::Void) {
            self.err(v.span, "E0360", "`void` 型の変数は宣言できません");
        }
        // Top-level variables of the module are globals initialised inside Main / the module initializer.
        let offset = u32::from(v.name.span.start());
        let global = if self.fx.is_toplevel && self.fx.scopes.len() == 1 {
            match self.prog.modules[self.module].scope.get(&v.name.text) {
                Some(Sym::Var(id)) if self.prog.vars[*id].name_span == v.name.span => Some(*id),
                _ => None,
            }
        } else {
            None
        };
        let init_ty = match &v.init {
            Some(init) => {
                if v.kind == ast::VarKind::Using {
                    self.fx.blockers.push("using");
                }
                let t = self.value(init, declared.as_ref());
                if let Some(d) = &declared {
                    self.coerce(init, &t, d);
                }
                Some(t)
            }
            None => None,
        };
        let ty = match (declared, init_ty) {
            (Some(d), _) => d,
            (None, Some(Type::Null)) => {
                self.err(
                    v.name.span,
                    "E0132",
                    "`null` で初期化する変数には型注釈が必要です",
                );
                Type::Error
            }
            (None, Some(t)) => t,
            (None, None) => {
                if global.is_none() {
                    self.err(v.name.span, "E0132", "型注釈か初期化式が必要です");
                }
                Type::Error
            }
        };
        if v.kind == ast::VarKind::Const && v.init.is_none() && global.is_none() {
            self.err(v.name.span, "E0133", "`const` には初期化式が必要です");
        }
        if v.kind == ast::VarKind::Using {
            let ok = match &ty {
                Type::Class(c, _) => !self.prog.find_methods(*c, "dispose").is_empty(),
                Type::Int => v.init.as_ref().is_some_and(|i| {
                    matches!(
                        self.info().calls.get(&i.peel().id).map(|c| &c.target),
                        Some(CallTarget::Intrinsic(Intr::FastMode))
                    )
                }),
                Type::Error | Type::Any | Type::Object => true,
                _ => false,
            };
            if !ok {
                self.err(v.span, "E0361", "`using` で束縛できるのは `dispose()` メソッドを持つオブジェクトか `fastMode()` です");
            }
        }
        match global {
            Some(id) => {
                if self.prog.vars[id].ty == Type::Void || self.prog.vars[id].ty.is_error() {
                    self.prog.vars[id].ty = ty;
                }
                if let Some(init) = &v.init {
                    self.prog.vars[id].const_value = if v.kind == ast::VarKind::Const {
                        const_value(init)
                    } else {
                        None
                    };
                }
                self.info().decl_globals.insert(offset, id);
            }
            None => {
                let kind = match v.kind {
                    ast::VarKind::Let => LocalKind::Let,
                    ast::VarKind::Const => LocalKind::Const,
                    ast::VarKind::Using => LocalKind::Using,
                };
                self.declare_local(&v.name, ty, kind);
            }
        }
    }

    // ------------------------------------------------------------------ expressions

    /// Checks an expression that must produce a value.
    fn value(&mut self, e: &Expr, expect: Option<&Type>) -> Type {
        let t = self.expr(e, expect);
        match &t {
            Type::Void => {
                self.err(e.span, "E0370", "この式は値を返しません（`void`）");
                Type::Error
            }
            Type::Ns(_) => {
                self.err(
                    e.span,
                    "E0370",
                    "ここには値が必要です（モジュール・名前空間・型は値として使えません）",
                );
                Type::Error
            }
            _ => t,
        }
    }

    fn expr(&mut self, e: &Expr, expect: Option<&Type>) -> Type {
        self.fx.cost += 1;
        let t = self.expr_inner(e, expect);
        let info = self.info();
        if let Some(slot) = info.expr_types.get_mut(e.id as usize) {
            *slot = t.clone();
        }
        t
    }

    fn comparable(&self, a: &Type, b: &Type, equality: bool) -> bool {
        use Type::*;
        if a.is_error() || b.is_error() || *a == Any || *b == Any {
            return true;
        }
        // A union equals a value when one of its members can; ordering needs every member to be ordered.
        if let Some(ms) = a.union_members() {
            return if equality {
                ms.iter().any(|m| self.comparable(m, b, true))
            } else {
                ms.iter().all(|m| self.comparable(m, b, false))
            };
        }
        if let Some(ms) = b.union_members() {
            return if equality {
                ms.iter().any(|m| self.comparable(a, m, true))
            } else {
                ms.iter().all(|m| self.comparable(a, m, false))
            };
        }
        if a.is_numeric() && b.is_numeric() {
            return true;
        }
        match (a, b) {
            (Str, Str) | (Date, Date) => true,
            (Bool, Bool) => equality,
            (Param(x), Param(y)) => {
                x == y
                    && (equality
                        || matches!(
                            self.constraint_of(x),
                            Constraint::Comparable | Constraint::Numeric
                        ))
            }
            (Param(x), other) | (other, Param(x)) => match self.constraint_of(x) {
                Constraint::Numeric => other.is_numeric(),
                Constraint::ObjectType => equality && (other.is_object()),
                _ => false,
            },
            _ if a.is_object() && b.is_object() => equality,
            _ => false,
        }
    }

    fn expr_inner(&mut self, e: &Expr, expect: Option<&Type>) -> Type {
        match &e.kind {
            ExprKind::Error => Type::Error,
            ExprKind::Int { value, suffix, .. } => {
                if *suffix == ast::NumSuffix::L {
                    if i64::try_from(*value).is_err() {
                        self.err(
                            e.span,
                            "E0380",
                            "整数リテラルが `long` の範囲を超えています",
                        );
                    }
                    Type::Long
                } else if i32::try_from(*value).is_ok() {
                    Type::Int
                } else if i64::try_from(*value).is_ok() {
                    Type::Long
                } else {
                    self.err(
                        e.span,
                        "E0380",
                        "整数リテラルが `long` の範囲を超えています",
                    );
                    Type::Error
                }
            }
            ExprKind::Float { suffix, .. } => match suffix {
                ast::NumSuffix::F => Type::Float,
                ast::NumSuffix::M => Type::Decimal,
                _ => Type::Double,
            },
            ExprKind::Str(_) => Type::Str,
            ExprKind::Template(parts) => {
                for p in parts {
                    if let ast::TemplatePart::Expr(x) = p {
                        let t = self.value(x, None);
                        if !self.stringifiable(&t) {
                            self.err(
                                x.span,
                                "E0381",
                                format!("`{}` は文字列に埋め込めません", self.show(&t)),
                            );
                        }
                    }
                }
                Type::Str
            }
            ExprKind::Bool(_) => Type::Bool,
            ExprKind::Null | ExprKind::Undefined => Type::Null,
            ExprKind::Name(id) => match self.name_access(e, id) {
                Access::Value(t) => t,
                Access::Func { f, .. } => {
                    let def = &self.prog.fns[f];
                    if matches!(def.kind, FnKind::Host { .. })
                        && def.required_params() == 0
                        && def.ret_ty() != Type::Void
                    {
                        // Host functions without required arguments read like properties (`Now`, `Cells`).
                        def.ret_ty()
                    } else {
                        self.err(e.span, "E0382", format!("関数 `{}` を値として使うことはできません。呼び出すか、アロー関数 `(x) => {}(x)` で包んでください", id.text, id.text));
                        Type::Error
                    }
                }
                Access::Intrinsic(i) => {
                    if i.is_property() {
                        self.intrinsic_call(e, i, None, None, &[], expect)
                    } else {
                        self.err(
                            e.span,
                            "E0382",
                            format!("組み込み関数 `{}` は呼び出して使ってください", id.text),
                        );
                        Type::Error
                    }
                }
                Access::Dynamic => Type::Any,
                Access::Error => Type::Error,
            },
            ExprKind::This => {
                if let Some(l) = self.fx.this_local {
                    self.set_res(e.id, Res::Local(l));
                    if let Some(lam) = self.current_lambda() {
                        self.capture(l, &[lam], e.span);
                    }
                    return self.fx.locals[l].ty.clone();
                }
                match self.fx.class {
                    Some(c) if !self.fx.is_static => {
                        let args = self.prog.classes[c]
                            .type_params
                            .iter()
                            .map(|p| Type::Param(p.name.clone()))
                            .collect();
                        Type::Class(c, args)
                    }
                    _ => {
                        self.err(e.span, "E0383", "`this` はクラスのインスタンスメソッド、拡張メソッド、フォームのコードビハインドの中でのみ使えます");
                        Type::Error
                    }
                }
            }
            ExprKind::Super => {
                self.err(
                    e.span,
                    "E0140",
                    "`super` には対応していません（クラスの継承は未対応です）",
                );
                Type::Error
            }
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => self.expr(inner, expect),
            ExprKind::Array(items) if matches!(expect, Some(Type::Struct(s)) if self.prog.structs[*s].kind == StructKind::Tuple) =>
            {
                let Some(Type::Struct(s)) = expect else {
                    unreachable!()
                };
                let fields = self.prog.structs[*s].fields.clone();
                if items.len() != fields.len() {
                    self.err(
                        e.span,
                        "E0392",
                        format!(
                            "`{}` には {} 個の要素が必要です（{} 個あります）",
                            self.show(&Type::Struct(*s)),
                            fields.len(),
                            items.len()
                        ),
                    );
                }
                for (i, it) in items.iter().enumerate() {
                    let want = fields.get(i).map(|f| f.1.clone());
                    let t = self.value(it, want.as_ref());
                    if let Some(w) = &want {
                        self.coerce(it, &t, w);
                    }
                }
                Type::Struct(*s)
            }
            ExprKind::Array(items) => {
                let elem_expect = match expect {
                    Some(Type::Array(el)) if !matches!(**el, Type::Array(_)) => {
                        Some((**el).clone())
                    }
                    _ => None,
                };
                if let Some(Type::Array(el)) = expect {
                    if matches!(**el, Type::Array(_)) && !items.is_empty() {
                        self.err(e.span, "E0384", "2次元配列は配列リテラルでは作れません。`new Array<T>(行数, 列数)` を使ってください");
                    }
                }
                let mut elem: Option<Type> = elem_expect.clone();
                for it in items {
                    let t = self.value(it, elem_expect.as_ref());
                    match &elem_expect {
                        Some(want) => self.coerce(it, &t, want),
                        None => {
                            elem = Some(match elem.take() {
                                None => t,
                                Some(prev) => self.unify_branches(it.span, &prev, &t),
                            });
                        }
                    }
                }
                match elem {
                    Some(Type::Null) | None => match expect {
                        Some(t @ Type::Array(_)) => t.clone(),
                        Some(Type::Any) => Type::array(Type::Any),
                        _ => {
                            self.err(e.span, "E0385", "配列リテラルの要素の型を決められません。`let a: int[] = []` のように型注釈を付けてください");
                            Type::Error
                        }
                    },
                    Some(t) => Type::array(t),
                }
            }
            ExprKind::Object(props) => self.struct_literal(e, props, expect),
            ExprKind::Arrow(arrow) => self.arrow(e, arrow, expect),
            ExprKind::Unary { op, operand } => {
                let t = self.value(operand, None);
                match op {
                    UnaryOp::Not => {
                        if !matches!(t, Type::Bool | Type::Any | Type::Error) {
                            self.err(
                                operand.span,
                                "E0303",
                                format!(
                                    "`!` は `boolean` にのみ使えます（`{}` です）",
                                    self.show(&t)
                                ),
                            );
                        }
                        Type::Bool
                    }
                    UnaryOp::Neg | UnaryOp::Plus => {
                        if !self.is_numeric(&t) {
                            self.err(
                                operand.span,
                                "E0387",
                                format!("数値が必要です（`{}` です）", self.show(&t)),
                            );
                            return Type::Error;
                        }
                        if let (UnaryOp::Neg, Some(v)) = (op, const_int(e)) {
                            // `-2147483648` is an int even though the digits alone are not.
                            if i32::try_from(v).is_ok()
                                && t == Type::Long
                                && !matches!(
                                    operand.peel().kind,
                                    ExprKind::Int {
                                        suffix: ast::NumSuffix::L,
                                        ..
                                    }
                                )
                            {
                                return Type::Int;
                            }
                        }
                        match t {
                            Type::Byte | Type::Short | Type::Enum(_) => Type::Int,
                            other => other,
                        }
                    }
                    UnaryOp::BitNot => {
                        if !(t.is_integer() || matches!(t, Type::Any | Type::Error | Type::Enum(_)))
                        {
                            self.err(operand.span, "E0387", "`~` は整数にのみ使えます");
                            return Type::Error;
                        }
                        match t {
                            Type::Byte | Type::Short | Type::Enum(_) => Type::Int,
                            other => other,
                        }
                    }
                    UnaryOp::TypeOf => Type::Str,
                }
            }
            ExprKind::IncDec { operand, .. } => {
                let t = self.lvalue(operand);
                if !self.is_numeric(&t) {
                    self.err(
                        operand.span,
                        "E0387",
                        format!(
                            "`++` / `--` は数値にのみ使えます（`{}` です）",
                            self.show(&t)
                        ),
                    );
                    return Type::Error;
                }
                t
            }
            ExprKind::Binary { op, lhs, rhs } => self.binary(e, *op, lhs, rhs),
            ExprKind::Assign { op, target, value } => {
                let tt = self.lvalue(target);
                match op {
                    None => {
                        let vt = self.value(value, Some(&tt));
                        self.coerce(value, &vt, &tt);
                    }
                    Some(op) => {
                        let vt = self.value(value, None);
                        let rt = self.binary_types(e.span, *op, &tt, &vt, target, value);
                        if !rt.is_error() && !self.assignable(&rt, &tt) {
                            self.err(
                                e.span,
                                "E0301",
                                format!(
                                    "`{}=` の結果（`{}`）を `{}` に代入できません",
                                    op.symbol(),
                                    self.show(&rt),
                                    self.show(&tt)
                                ),
                            );
                        }
                    }
                }
                tt
            }
            ExprKind::Ternary { cond, then, els } => {
                self.cond(cond);
                let (t, f) = self.guard_facts(cond, &Vec::new());
                let a = self.with_facts(t, &mut |n| assigns_in_expr(then, n), |ck| {
                    ck.value(then, expect)
                });
                let b = self.with_facts(f, &mut |n| assigns_in_expr(els, n), |ck| {
                    ck.value(els, expect.or(Some(&a)))
                });
                match expect {
                    Some(want)
                        if self.assignable(&a, want)
                            && self.assignable(&b, want)
                            && *want != Type::Any =>
                    {
                        want.clone()
                    }
                    _ => self.unify_branches(e.span, &a, &b),
                }
            }
            ExprKind::Call {
                callee,
                type_args,
                args,
            } => self.call(e, callee, type_args, args, expect),
            ExprKind::New { ty, args } => self.new_expr(e, ty, args),
            ExprKind::NamedArg { value, .. } => {
                self.err(e.span, "E0411", "名前付き引数は、宣言された関数・メソッド・コンストラクタの呼び出しにだけ書けます");
                self.expr(value, None);
                Type::Error
            }
            ExprKind::Member {
                object,
                name,
                optional,
            } => {
                if *optional {
                    self.err(e.span, "E0388", "オプショナルチェーン（`?.`）には対応していません。`!= null` で判定してください");
                }
                let ot = self.expr(object, None);
                match self.member_access(e, object, &ot, name) {
                    Access::Value(t) => t,
                    Access::Func { f, env, recv } => {
                        let def = &self.prog.fns[f];
                        let host = matches!(def.kind, FnKind::Host { .. });
                        let union_ok = self.fx.union_calls.get(&e.id).is_none_or(|o| {
                            o.iter()
                                .all(|(g, _)| self.prog.fns[*g].required_params() == 0)
                        });
                        if host
                            && def.required_params() == 0
                            && def.ret_ty() != Type::Void
                            && recv != Recv::Extension
                            && union_ok
                        {
                            let ret = def.ret_ty().subst(&env);
                            self.union_call_ret(e.id, ret)
                        } else {
                            self.err(
                                name.span,
                                "E0382",
                                format!(
                                    "メソッド `{}` は呼び出して使ってください（`()` が必要です）",
                                    name.text
                                ),
                            );
                            Type::Error
                        }
                    }
                    Access::Intrinsic(i) => {
                        if i.is_property() {
                            self.intrinsic_call(e, i, Some(object), Some(&ot), &[], expect)
                        } else {
                            self.err(
                                name.span,
                                "E0382",
                                format!(
                                    "メソッド `{}` は呼び出して使ってください（`()` が必要です）",
                                    name.text
                                ),
                            );
                            Type::Error
                        }
                    }
                    Access::Dynamic => Type::Any,
                    Access::Error => Type::Error,
                }
            }
            ExprKind::Index { object, index } => {
                let ot = self.value(object, None);
                match &ot {
                    Type::Array(el) => {
                        let it = self.value(index, Some(&Type::Int));
                        if !(it.is_integer()
                            || matches!(it, Type::Any | Type::Error | Type::Enum(_)))
                        {
                            self.err(
                                index.span,
                                "E0389",
                                format!(
                                    "配列の添字は整数である必要があります（`{}` です）",
                                    self.show(&it)
                                ),
                            );
                        }
                        if matches!(**el, Type::Array(_)) {
                            // `a[r]` of a 2-D array is only valid as the object of another index.
                            self.fx.row_views.insert(e.id, e.span);
                        }
                        self.fx.row_views.remove(&object.peel().id);
                        (**el).clone()
                    }
                    Type::Str => {
                        let it = self.value(index, Some(&Type::Int));
                        if !(it.is_integer() || matches!(it, Type::Any | Type::Error)) {
                            self.err(
                                index.span,
                                "E0389",
                                "文字列の添字は整数である必要があります",
                            );
                        }
                        Type::Str
                    }
                    Type::Map(k, v) => {
                        let it = self.value(index, Some(k));
                        self.coerce(index, &it, k);
                        (**v).clone()
                    }
                    Type::Struct(s) if self.prog.structs[*s].kind == StructKind::Tuple => {
                        self.value(index, Some(&Type::Int));
                        let fields = &self.prog.structs[*s].fields;
                        match &index.peel().kind {
                            ExprKind::Int { value, .. }
                                if (*value as usize) < fields.len() && *value >= 0 =>
                            {
                                fields[*value as usize].1.clone()
                            }
                            ExprKind::Int { .. } => {
                                let (n, shown) = (fields.len(), self.show(&ot));
                                self.err(
                                    index.span,
                                    "E0389",
                                    format!("`{shown}` の要素は {n} 個です"),
                                );
                                Type::Error
                            }
                            _ => {
                                self.err(
                                    index.span,
                                    "E0389",
                                    "タプルの添字は整数リテラル（`t[0]`）で書いてください",
                                );
                                Type::Error
                            }
                        }
                    }
                    Type::Class(c, args) if self.prog.classes[*c].call_sig.is_some() => {
                        let sig = self.prog.classes[*c].call_sig.unwrap();
                        let env = self.prog.class_env(*c, args);
                        self.value(index, None);
                        self.info().calls.insert(
                            e.id,
                            CallInfo {
                                target: CallTarget::DefaultMember(sig),
                                type_args: Vec::new(),
                            },
                        );
                        self.prog.fns[sig].ret_ty().subst(&env)
                    }
                    Type::Any | Type::Object => {
                        self.value(index, None);
                        Type::Any
                    }
                    Type::Error => {
                        self.value(index, None);
                        Type::Error
                    }
                    other => {
                        let shown = self.show(other);
                        self.value(index, None);
                        self.err(e.span, "E0389", format!("`{shown}` は添字で参照できません"));
                        Type::Error
                    }
                }
            }
            ExprKind::As { expr, ty } => {
                let tps = self.fx.type_params.clone();
                let target = resolve_type(self.prog, self.module, &tps, ty);
                let from = self.value(expr, Some(&target));
                if !self.castable(&from, &target) {
                    self.err(
                        e.span,
                        "E0390",
                        format!(
                            "`{}` を `{}` に変換できません",
                            self.show(&from),
                            self.show(&target)
                        ),
                    );
                }
                target
            }
            ExprKind::InstanceOf { expr, ty } => {
                let from = self.value(expr, None);
                let tps = self.fx.type_params.clone();
                let target = resolve_type(self.prog, self.module, &tps, ty);
                self.fx.instanceof_types.insert(e.id, target.clone());
                if !(from.is_object()
                    || from.is_mixed_union()
                    || matches!(from, Type::Any | Type::Error))
                {
                    self.err(
                        expr.span,
                        "E0391",
                        "`instanceof` の左辺はオブジェクトである必要があります",
                    );
                }
                if !matches!(target, Type::Class(..) | Type::Error) {
                    self.err(
                        ty.span(),
                        "E0391",
                        "`instanceof` の右辺はクラスかインターフェイスである必要があります",
                    );
                }
                Type::Bool
            }
        }
    }

    /// `{ a: 1, b: "x" }`: a value of the expected struct type, or of the anonymous
    /// struct with these fields when nothing more specific is expected.
    fn struct_literal(
        &mut self,
        e: &Expr,
        props: &[ast::ObjectProp],
        expect: Option<&Type>,
    ) -> Type {
        let mut ok = true;
        for (i, p) in props.iter().enumerate() {
            if props[..i]
                .iter()
                .any(|q| q.key.text.eq_ignore_ascii_case(&p.key.text))
            {
                self.err(
                    p.key.span,
                    "E0392",
                    format!("フィールド `{}` が重複しています", p.key.text),
                );
                ok = false;
            }
        }
        if props.is_empty() {
            self.err(e.span, "E0392", "空のオブジェクトリテラルは使えません");
            return Type::Error;
        }
        let target = match expect {
            Some(Type::Struct(s)) if self.prog.structs[*s].kind != StructKind::Tuple => Some(*s),
            _ => None,
        };
        match target {
            Some(s) => {
                let def = self.prog.structs[s].clone();
                for p in props {
                    match def.fields.iter().find(|f| f.0 == p.key.text) {
                        Some((_, fty)) => {
                            let t = self.value(&p.value, Some(fty));
                            self.coerce(&p.value, &t, fty);
                        }
                        None => {
                            self.value(&p.value, None);
                            self.err(
                                p.key.span,
                                "E0202",
                                format!(
                                    "`{}` にフィールド `{}` はありません",
                                    self.show(&Type::Struct(s)),
                                    p.key.text
                                ),
                            );
                        }
                    }
                }
                let missing: Vec<String> = def
                    .fields
                    .iter()
                    .zip(&def.optional)
                    .filter(|(f, opt)| !**opt && !props.iter().any(|p| p.key.text == f.0))
                    .map(|(f, _)| format!("`{}`", f.0))
                    .collect();
                if !missing.is_empty() {
                    self.err(e.span, "E0392", format!("フィールド {} がありません（省略できるのは `name?: T` と宣言したフィールドだけです）", missing.join("、")));
                }
                Type::Struct(s)
            }
            None => {
                let mut fields = Vec::new();
                for p in props {
                    let t = self.value(&p.value, None);
                    match &t {
                        Type::Null => {
                            self.err(p.value.span, "E0392", "`null` のフィールドの型を決められません。`struct` を宣言して型注釈を付けてください");
                            ok = false;
                        }
                        Type::Error => ok = false,
                        t if t.contains_param() => {
                            self.err(p.value.span, "E0392", "型パラメータを含む値は無名のオブジェクトにできません。`struct` を宣言してください");
                            ok = false;
                        }
                        _ => {}
                    }
                    fields.push((p.key.text.clone(), t, false));
                }
                if !ok {
                    return Type::Error;
                }
                Type::Struct(
                    self.prog
                        .intern_struct(self.module, StructKind::Object, fields),
                )
            }
        }
    }

    fn stringifiable(&self, t: &Type) -> bool {
        if let Some(ms) = t.union_members() {
            return ms.iter().all(|m| self.stringifiable(m));
        }
        t.is_numeric()
            || matches!(
                t,
                Type::Str | Type::Bool | Type::Date | Type::Any | Type::Error | Type::Param(_)
            )
    }

    fn castable(&self, from: &Type, to: &Type) -> bool {
        use Type::*;
        if self.assignable(from, to) || self.assignable(to, from) {
            return true;
        }
        if let Some(fs) = from.union_members() {
            return fs.iter().any(|f| self.castable(f, to));
        }
        if let Some(ts) = to.union_members() {
            return ts.iter().any(|t| self.castable(from, t));
        }
        match (from, to) {
            (a, b) if a.is_numeric() && b.is_numeric() => true,
            (Bool, b) if b.is_numeric() => true,
            (Date, Double) | (Double, Date) | (Date, Float) | (Float, Date) => true,
            (Class(..), Class(..)) => true,
            (Param(_), _) | (_, Param(_)) => true,
            _ => false,
        }
    }

    fn capture(&mut self, local: LocalId, lambdas: &[ExprId], span: Span) {
        self.fx.locals[local].captured = true;
        self.fx.captures.push((local, span));
        for lam in lambdas {
            let list = self.info().lambda_captures.entry(*lam).or_default();
            if !list.contains(&local) {
                list.push(local);
            }
        }
    }

    fn name_access(&mut self, e: &Expr, id: &Ident) -> Access {
        if let Some((local, crossed)) = self.lookup_local(&id.text) {
            self.fx.locals[local].used = true;
            if !crossed.is_empty() {
                self.capture(local, &crossed, e.span);
            }
            self.set_res(e.id, Res::Local(local));
            return Access::Value(self.local_ty(local));
        }
        match self.lookup_sym(&id.text) {
            Some(Sym::Var(v)) => {
                let ty = self.var_type(v, e.span);
                let owner = self.prog.vars[v].module;
                // Constants may also be read from arrow functions (they are copied when a closure is created).
                let in_own_main = self.fx.is_toplevel
                    && owner == self.module
                    && (self.current_lambda().is_none() || self.prog.vars[v].is_const);
                if !in_own_main {
                    self.prog.vars[v].used_outside_main = true;
                }
                if !self.fx.var_refs.contains(&v) {
                    self.fx.var_refs.push(v);
                }
                if self.prog.vars[v].declared {
                    let host = self.prog.vars[v].vba_name().to_string();
                    self.lint_host_global(e, &host, &id.text);
                }
                self.set_res(e.id, Res::Global(v));
                Access::Value(ty)
            }
            Some(Sym::Fn(f)) => {
                self.set_res(e.id, Res::Fn(f));
                if matches!(self.prog.fns[f].kind, FnKind::Host { .. }) {
                    let host = self.prog.fns[f].vba_name().to_string();
                    self.lint_host_global(e, &host, &id.text);
                }
                Access::Func {
                    f,
                    env: HashMap::new(),
                    recv: Recv::None,
                }
            }
            Some(Sym::Class(c)) => {
                self.set_res(e.id, Res::Class(c));
                Access::Value(Type::Ns(NsRef::Class(c)))
            }
            Some(Sym::Enum(en)) => {
                self.set_res(e.id, Res::Enum(en));
                Access::Value(Type::Ns(NsRef::Enum(en)))
            }
            Some(Sym::EnumMember(en, i)) => {
                self.set_res(e.id, Res::EnumMember(en, i));
                Access::Value(Type::Enum(en))
            }
            Some(Sym::Module(m)) => {
                self.set_res(e.id, Res::Module(m));
                Access::Value(Type::Ns(NsRef::Module(m)))
            }
            Some(Sym::Namespace(n)) => {
                self.set_res(e.id, Res::Namespace(n.clone()));
                Access::Value(Type::Ns(NsRef::Host(n)))
            }
            Some(Sym::Struct(_) | Sym::Alias(_)) => {
                self.err(
                    e.span,
                    "E0370",
                    format!("`{}` は型です。値として使うことはできません", id.text),
                );
                Access::Error
            }
            None => {
                if self.prog.namespaces.contains_key(&id.text) {
                    self.set_res(e.id, Res::Namespace(id.text.clone()));
                    return Access::Value(Type::Ns(NsRef::Host(id.text.clone())));
                }
                if let Some(ns) = intrinsics::BUILTIN_NAMESPACES
                    .iter()
                    .find(|n| **n == id.text)
                {
                    self.set_res(e.id, Res::Builtin(ns));
                    return Access::Value(Type::Ns(NsRef::Builtin(ns)));
                }
                if let Some(i) = intrinsics::global_fn(&id.text) {
                    self.set_res(e.id, Res::Intrinsic(i));
                    return Access::Intrinsic(i);
                }
                // VBA does not distinguish case: `Cells` is the bundled `cells`.
                if let Some(js) = self
                    .prog
                    .host_global_names
                    .get(&id.text.to_lowercase())
                    .cloned()
                {
                    if js != id.text && self.lookup_sym(&js).is_some_and(|s| self.is_host_sym(&s)) {
                        self.warn_host_case(id, &js);
                        return self.name_access(
                            e,
                            &Ident {
                                text: js,
                                span: id.span,
                            },
                        );
                    }
                }
                if id.text == "Error" {
                    self.err(
                        e.span,
                        "E0201",
                        "`Error` は `throw new Error(\"...\")` の形で使ってください",
                    );
                } else if !id.text.is_empty() {
                    self.err(
                        e.span,
                        "E0201",
                        format!("名前 `{}` が見つかりません", id.text),
                    );
                }
                Access::Error
            }
        }
    }

    fn is_host_sym(&self, s: &Sym) -> bool {
        match s {
            Sym::Var(v) => self.prog.vars[*v].declared,
            Sym::Fn(f) => matches!(self.prog.fns[*f].kind, FnKind::Host { .. }),
            _ => false,
        }
    }

    /// A bundled host name written with another casing (usually the type library's `Cells`).
    /// Whether the controls on the sheets are known (the project's workbook was read): a member
    /// a sheet does not have is then an error, not a late-bound call (spec §12.2.1).
    fn knows_document_members(&self, _module: ModuleId) -> bool {
        self.prog.config.sheet_controls_known
    }

    fn warn_host_case(&mut self, written: &Ident, js: &str) {
        self.warn(written.span, "W0304", format!("`{}` は RAS では `{js}` と書きます（Excel / VBA のメンバーは camelCase で参照します）", written.text));
        self.prog.fixes.push(crate::diag::Fix {
            module: self.module,
            span: written.span,
            replacement: js.to_string(),
            code: "W0304",
        });
    }

    /// The RAS name of the member of host class `c` that VBA would find for `name` (ignoring case).
    fn host_member_ci(&self, c: ClassId, name: &str) -> Option<String> {
        let def = &self.prog.classes[c];
        if def.kind != ClassKind::Host {
            return None;
        }
        let hit =
            |n: &str, host: &str| n.eq_ignore_ascii_case(name) || host.eq_ignore_ascii_case(name);
        if let Some(f) = def.fields.iter().find(|f| hit(&f.name, f.vba_name())) {
            return Some(f.name.clone());
        }
        if let Some(&m) = def
            .methods
            .iter()
            .find(|&&m| hit(&self.prog.fns[m].name, self.prog.fns[m].vba_name()))
        {
            return Some(self.prog.fns[m].name.clone());
        }
        def.extends.iter().find_map(|b| match b {
            Type::Class(b, _) => self.host_member_ci(*b, name),
            _ => None,
        })
    }

    /// Lint rules 9 and 10 of spec §6. `name` is the type library's spelling, `shown` the one in the source.
    fn lint_host_global(&mut self, e: &Expr, name: &str, shown: &str) {
        let shown = self
            .prog
            .host_global_names
            .get(&shown.to_lowercase())
            .cloned()
            .unwrap_or_else(|| shown.to_string());
        match name {
            "Range" | "Cells" | "Rows" | "Columns" | "ActiveSheet" | "ActiveCell"
            | "ActiveWorkbook" => {
                if self.fx.qualify_depth > 0 && matches!(name, "Range" | "Cells") {
                    self.info().auto_qualified.insert(e.id);
                } else {
                    self.warn(e.span, "W0401", format!("修飾されていない `{shown}` はアクティブなシート／ブックに依存します。`ws.{shown}` のようにシートを明示してください"));
                }
            }
            "Selection" => self.diag(
                Severity::Info,
                e.span,
                "I0402",
                format!("`{shown}` への依存は遅く不安定です。対象の Range を直接操作してください"),
            ),
            _ => {}
        }
    }

    fn class_member(
        &mut self,
        e: &Expr,
        c: ClassId,
        targs: &[Type],
        name: &Ident,
        static_access: bool,
    ) -> Option<Access> {
        let env = self.prog.class_env(c, targs);
        if let Some((owner, idx)) = self.prog.find_field(c, &name.text) {
            let field = self.prog.classes[owner].fields[idx].clone();
            if field.is_static == static_access {
                if field.vis != Visibility::Public && self.fx.class != Some(owner) {
                    self.err(
                        name.span,
                        "E0392",
                        format!(
                            "`{}` は private なのでクラスの外からは参照できません",
                            name.text
                        ),
                    );
                }
                self.set_res(e.id, Res::Field(owner, idx));
                let ty = if field.ty == Type::Void {
                    Type::Error
                } else {
                    field.ty.subst(&env)
                };
                return Some(Access::Value(ty));
            }
        }
        let methods = self.prog.find_methods(c, &name.text);
        let mut getter = None;
        let mut setter = None;
        for &m in &methods {
            let def = &self.prog.fns[m];
            let (is_static, kind) = match def.kind {
                FnKind::Method { is_static, .. } => (is_static, 0),
                FnKind::Getter { is_static, .. } => (is_static, 1),
                FnKind::Setter { is_static, .. } => (is_static, 2),
                _ => (false, 0),
            };
            if is_static != static_access {
                continue;
            }
            if def.vis != Visibility::Public && self.fx.class != Some(c) {
                self.err(
                    name.span,
                    "E0392",
                    format!(
                        "`{}` は private なのでクラスの外からは参照できません",
                        name.text
                    ),
                );
            }
            match kind {
                1 => getter = Some(m),
                2 => setter = Some(m),
                _ => {
                    self.set_res(e.id, Res::Method(m));
                    let recv = if static_access {
                        Recv::None
                    } else {
                        Recv::Instance
                    };
                    return Some(Access::Func { f: m, env, recv });
                }
            }
        }
        if getter.is_some() || setter.is_some() {
            self.set_res(
                e.id,
                Res::Property {
                    get: getter,
                    set: setter,
                },
            );
            let ty = match (getter, setter) {
                (Some(g), _) => self.fn_ret(g, name.span).subst(&env),
                (None, Some(s)) => self.prog.fns[s]
                    .params
                    .first()
                    .map(|p| p.ty.subst(&env))
                    .unwrap_or(Type::Error),
                _ => Type::Error,
            };
            return Some(Access::Value(ty));
        }
        None
    }

    /// Finds the best extension method `name` applicable to a receiver of type `recv` (spec §3.2.3).
    fn find_ext(&mut self, recv: &Type, name: &Ident) -> Option<(FnId, HashMap<String, Type>)> {
        let candidates: Vec<FnId> = self.prog.modules[self.module]
            .ext_methods
            .iter()
            .copied()
            .filter(|&f| self.prog.fns[f].name == name.text)
            .collect();
        let mut best: Vec<(u8, FnId, HashMap<String, Type>)> = Vec::new();
        for f in candidates {
            let def = &self.prog.fns[f];
            let Some(p0) = def.params.first() else {
                continue;
            };
            let mut env = HashMap::new();
            self.unify(&p0.ty, recv, &mut env, &def.type_params);
            let want = p0.ty.subst(&env);
            let score = if want == *recv {
                if def.type_params.is_empty() {
                    0
                } else {
                    1
                }
            } else if self.assignable(recv, &want) && !matches!(want, Type::Param(_)) {
                2
            } else {
                continue;
            };
            best.push((score, f, env));
        }
        best.sort_by_key(|b| b.0);
        match best.len() {
            0 => None,
            1 => best.pop().map(|(_, f, env)| (f, env)),
            _ => {
                if best[0].0 == best[1].0 {
                    self.err(
                        name.span,
                        "E0393",
                        format!(
                            "拡張メソッド `{}` の候補が複数あり、どれを呼ぶか決められません",
                            name.text
                        ),
                    );
                }
                let (_, f, env) = best.swap_remove(0);
                Some((f, env))
            }
        }
    }

    fn ext_access(&mut self, e: &Expr, recv: &Type, name: &Ident) -> Option<Access> {
        let (f, env) = self.find_ext(recv, name)?;
        self.set_res(e.id, Res::ExtMethod(f));
        Some(Access::Func {
            f,
            env,
            recv: Recv::Extension,
        })
    }

    fn member_access(&mut self, e: &Expr, object: &Expr, ot: &Type, name: &Ident) -> Access {
        if name.text.is_empty() {
            return Access::Error;
        }
        match ot {
            Type::Error => Access::Error,
            Type::Ns(NsRef::Module(m)) => {
                match self.prog.modules[*m].exports.get(&name.text).cloned() {
                    Some(Sym::Fn(f)) => {
                        self.set_res(e.id, Res::Fn(f));
                        Access::Func {
                            f,
                            env: HashMap::new(),
                            recv: Recv::None,
                        }
                    }
                    Some(Sym::Var(v)) => {
                        let ty = self.var_type(v, e.span);
                        self.prog.vars[v].used_outside_main = true;
                        self.set_res(e.id, Res::Global(v));
                        Access::Value(ty)
                    }
                    Some(Sym::Class(c)) => {
                        self.set_res(e.id, Res::Class(c));
                        Access::Value(Type::Ns(NsRef::Class(c)))
                    }
                    Some(Sym::Enum(en)) => {
                        self.set_res(e.id, Res::Enum(en));
                        Access::Value(Type::Ns(NsRef::Enum(en)))
                    }
                    _ => {
                        let key = self.prog.modules[*m].key.clone();
                        self.err(
                            name.span,
                            "E0123",
                            format!(
                                "モジュール `{key}` に `export` された `{}` はありません",
                                name.text
                            ),
                        );
                        Access::Error
                    }
                }
            }
            Type::Ns(NsRef::Host(ns)) => {
                if ns == "Excel" && name.text == "match" {
                    self.set_res(e.id, Res::Intrinsic(Intr::ExcelMatch));
                    return Access::Intrinsic(Intr::ExcelMatch);
                }
                match self
                    .prog
                    .namespaces
                    .get(ns)
                    .and_then(|n| n.get(&name.text))
                    .cloned()
                {
                    Some(Sym::Class(c)) => {
                        self.set_res(e.id, Res::Class(c));
                        Access::Value(Type::Ns(NsRef::Class(c)))
                    }
                    Some(Sym::Enum(en)) => {
                        self.set_res(e.id, Res::Enum(en));
                        Access::Value(Type::Ns(NsRef::Enum(en)))
                    }
                    Some(Sym::Fn(f)) => {
                        self.set_res(e.id, Res::Fn(f));
                        Access::Func {
                            f,
                            env: HashMap::new(),
                            recv: Recv::None,
                        }
                    }
                    Some(Sym::Var(v)) => {
                        self.set_res(e.id, Res::Global(v));
                        Access::Value(self.prog.vars[v].ty.clone())
                    }
                    _ => {
                        let js = crate::names::js_name(&name.text);
                        if js != name.text
                            && matches!(
                                self.prog.namespaces.get(ns).and_then(|n| n.get(&js)),
                                Some(Sym::Fn(_) | Sym::Var(_))
                            )
                        {
                            self.warn_host_case(name, &js);
                            return self.member_access(
                                e,
                                object,
                                ot,
                                &Ident {
                                    text: js,
                                    span: name.span,
                                },
                            );
                        }
                        self.err(
                            name.span,
                            "E0202",
                            format!("名前空間 `{ns}` に `{}` はありません", name.text),
                        );
                        Access::Error
                    }
                }
            }
            Type::Ns(NsRef::Enum(en)) => match self.prog.enums[*en]
                .members
                .iter()
                .position(|m| m.name == name.text)
            {
                Some(i) => {
                    self.set_res(e.id, Res::EnumMember(*en, i));
                    Access::Value(Type::Enum(*en))
                }
                None => {
                    let ename = self.prog.enums[*en].name.clone();
                    self.err(
                        name.span,
                        "E0202",
                        format!("列挙型 `{ename}` に `{}` はありません", name.text),
                    );
                    Access::Error
                }
            },
            Type::Ns(NsRef::Class(c)) => match self.class_member(e, *c, &[], name, true) {
                Some(a) => a,
                None => {
                    let cname = self.prog.classes[*c].name.clone();
                    self.err(
                        name.span,
                        "E0202",
                        format!(
                            "クラス `{cname}` に static メンバー `{}` はありません",
                            name.text
                        ),
                    );
                    Access::Error
                }
            },
            Type::Ns(NsRef::Builtin(ns)) => match intrinsics::builtin_ns_member(ns, &name.text) {
                Some(i) => {
                    self.set_res(e.id, Res::Intrinsic(i));
                    Access::Intrinsic(i)
                }
                None => {
                    self.err(
                        name.span,
                        "E0202",
                        format!("`{ns}` に `{}` はありません", name.text),
                    );
                    Access::Error
                }
            },
            Type::Class(c, targs) => {
                // 1. members of the type win over extension methods (spec §3.2.3 rule 1)
                if let Some(a) = self.class_member(e, *c, targs, name, false) {
                    return a;
                }
                if let Some(a) = self.ext_access(e, ot, name) {
                    return a;
                }
                if let Some(i) = intrinsics::host_member(self.prog, ot, &name.text) {
                    self.set_res(e.id, Res::Intrinsic(i));
                    return Access::Intrinsic(i);
                }
                // Array-style methods of host collections (spec §4.4).
                if self.prog.collection_elem(ot).is_some() {
                    if let Some(i) = intrinsics::collection_member(&name.text) {
                        self.set_res(e.id, Res::Intrinsic(i));
                        return Access::Intrinsic(i);
                    }
                }
                if let Some(js) = self
                    .host_member_ci(*c, &name.text)
                    .filter(|js| *js != name.text)
                {
                    self.warn_host_case(name, &js);
                    return self.member_access(
                        e,
                        object,
                        ot,
                        &Ident {
                            text: js,
                            span: name.span,
                        },
                    );
                }
                let def = &self.prog.classes[*c];
                // A form's members are all known: its controls, the UserForm members, its exports.
                let object = self.prog.modules[def.module]
                    .form
                    .as_ref()
                    .filter(|f| f.class == Some(*c))
                    .map(|f| f.kind);
                if object == Some(ObjectKind::Form) {
                    let shown = self.prog.class_display(*c);
                    self.err(name.span, "E0202", format!("フォーム `{shown}` にコントロールやメンバー `{}` はありません（コントロールは forms/{shown}.form.json に、公開する関数は export して書きます）", name.text));
                    Access::Error
                } else if object.is_some() && self.knows_document_members(def.module) {
                    // A sheet's members are all known too: those of the worksheet, the controls
                    // on it and the exports of its file (spec §12.2).
                    let shown = self.prog.class_display(*c);
                    let what = if object == Some(ObjectKind::Workbook) {
                        "ブック".to_string()
                    } else {
                        format!("シート `{shown}` ")
                    };
                    self.err(name.span, "E0202", format!("{what}にメンバーやコントロール `{}` はありません（公開する関数は objects/{shown}.ras に export して書きます。Excel で足したコントロールは、ブックを保存すると使えます）", name.text));
                    Access::Error
                } else if def.kind == ClassKind::Host {
                    let shown = self.prog.class_display(*c);
                    self.warn(name.span, "W0301", format!("`{shown}` の型定義に `{}` はありません。遅延バインディングとして扱います", name.text));
                    self.set_res(e.id, Res::Dynamic);
                    Access::Dynamic
                } else {
                    let shown = self.prog.class_display(*c);
                    self.err(
                        name.span,
                        "E0202",
                        format!("`{shown}` にメンバー `{}` はありません", name.text),
                    );
                    Access::Error
                }
            }
            Type::Union(members) => self.union_member_access(e, object, ot, members, name),
            Type::Struct(s) => match self.prog.structs[*s]
                .fields
                .iter()
                .find(|f| f.0 == name.text)
                .map(|f| f.1.clone())
            {
                Some(ty) => {
                    self.set_res(e.id, Res::Dynamic);
                    Access::Value(ty)
                }
                None => {
                    self.err(
                        name.span,
                        "E0202",
                        format!("構造体にフィールド `{}` はありません", name.text),
                    );
                    Access::Error
                }
            },
            Type::Any | Type::Object => {
                // spec §3.2.3 rule 5
                let has_ext = self.prog.modules[self.module]
                    .ext_methods
                    .iter()
                    .any(|&f| self.prog.fns[f].name == name.text);
                if has_ext {
                    self.warn(name.span, "W0302", format!("レシーバが `{}` 型のため、拡張メソッド `{}` は使われず遅延バインディングの呼び出しになります", self.show(ot), name.text));
                }
                self.set_res(e.id, Res::Dynamic);
                Access::Dynamic
            }
            Type::Param(p) => {
                if let Constraint::Type(Type::Class(c, targs)) = self.constraint_of(p) {
                    if let Some(a) = self.class_member(e, c, &targs, name, false) {
                        return a;
                    }
                }
                if let Some(a) = self.ext_access(e, ot, name) {
                    return a;
                }
                self.err(name.span, "E0202", format!("型パラメータ `{p}` にメンバー `{}` はありません（`{p} extends インターフェイス` で制約を付けてください）", name.text));
                Access::Error
            }
            _ => {
                // user extension methods take priority over built-in members (spec §3.2.4)
                let builtin = match ot {
                    // `controls.on("Click", …)` on an array of objects (spec §12.2.3).
                    Type::Array(_)
                        if ot.dims() == 1
                            && intrinsics::EVENT_MEMBERS.contains(&name.text.as_str())
                            && matches!(ot.elem(), Some(Type::Class(..))) =>
                    {
                        Some(Intr::EventOn)
                    }
                    Type::Array(_) => intrinsics::array_member(&name.text, ot.dims()),
                    Type::Str => intrinsics::string_member(&name.text),
                    Type::Map(..) => intrinsics::map_member(&name.text),
                    Type::Bool => (name.text == "toString").then_some(Intr::BoolToString),
                    Type::Date => (name.text == "toString").then_some(Intr::DateToString),
                    t if t.is_numeric() => match name.text.as_str() {
                        "toString" => Some(Intr::NumToString),
                        "toFixed" => Some(Intr::NumToFixed),
                        _ => None,
                    },
                    _ => None,
                };
                if let Some(a) = self.ext_access(e, ot, name) {
                    if builtin.is_some() {
                        self.warn(name.span, "W0303", format!("ユーザー定義の拡張メソッド `{}` が標準ライブラリの同名メソッドより優先されます", name.text));
                    }
                    return a;
                }
                match builtin {
                    Some(i) => {
                        self.set_res(e.id, Res::Intrinsic(i));
                        let _ = object;
                        Access::Intrinsic(i)
                    }
                    None => {
                        self.err(
                            name.span,
                            "E0202",
                            format!(
                                "`{}` にメンバー `{}` はありません",
                                self.show(ot),
                                name.text
                            ),
                        );
                        Access::Error
                    }
                }
            }
        }
    }

    /// Instance member `name` of a class (or host interface), without reporting anything.
    fn lookup_instance_member(
        &mut self,
        c: ClassId,
        targs: &[Type],
        name: &str,
    ) -> Option<UnionMember> {
        let env = self.prog.class_env(c, targs);
        if let Some((owner, idx)) = self.prog.find_field(c, name) {
            let field = &self.prog.classes[owner].fields[idx];
            if !field.is_static {
                let ty = if field.ty == Type::Void {
                    Type::Error
                } else {
                    field.ty.subst(&env)
                };
                return Some(UnionMember::Value {
                    res: Res::Field(owner, idx),
                    ty,
                    writable: !field.readonly,
                    owner,
                    public: field.vis == Visibility::Public,
                });
            }
        }
        let mut getter = None;
        let mut setter = None;
        for m in self.prog.find_methods(c, name) {
            let def = &self.prog.fns[m];
            let public = def.vis == Visibility::Public;
            match def.kind {
                FnKind::Getter {
                    is_static: false, ..
                } => getter = Some(m),
                FnKind::Setter {
                    is_static: false, ..
                } => setter = Some(m),
                FnKind::Method {
                    is_static: true, ..
                }
                | FnKind::Getter { .. }
                | FnKind::Setter { .. } => {}
                _ => {
                    return Some(UnionMember::Method {
                        f: m,
                        env,
                        owner: c,
                        public,
                    })
                }
            }
        }
        let f = getter.or(setter)?;
        let public = self.prog.fns[f].vis == Visibility::Public;
        let ty = match (getter, setter) {
            (Some(g), _) => self.fn_ret(g, Span::default()).subst(&env),
            (None, Some(st)) => self.prog.fns[st]
                .params
                .first()
                .map(|p| p.ty.subst(&env))
                .unwrap_or(Type::Error),
            _ => Type::Error,
        };
        Some(UnionMember::Value {
            res: Res::Property {
                get: getter,
                set: setter,
            },
            ty,
            writable: setter.is_some(),
            owner: c,
            public,
        })
    }

    /// `x.name` where `x` is a union: a member that every member type has (spec §2.6.3).
    fn union_member_access(
        &mut self,
        e: &Expr,
        object: &Expr,
        ot: &Type,
        members: &[Type],
        name: &Ident,
    ) -> Access {
        let shown = self.show(ot);
        let classes: Vec<(ClassId, Vec<Type>)> = members
            .iter()
            .filter_map(|m| match m {
                Type::Class(c, args) => Some((*c, args.clone())),
                _ => None,
            })
            .collect();
        if classes.len() != members.len() {
            if let Some(a) = self.ext_access(e, ot, name) {
                return a;
            }
            self.err(name.span, "E0202", format!("ユニオン型 `{shown}` のメンバー `{}` は参照できません。`instanceof` / `typeof` で型を絞り込んでください", name.text));
            return Access::Error;
        }
        let mut found = Vec::new();
        let mut missing = Vec::new();
        for (c, targs) in &classes {
            match self.lookup_instance_member(*c, targs, &name.text) {
                Some(m) => found.push(m),
                None => missing.push(*c),
            }
        }
        if !missing.is_empty() {
            if let Some(a) = self.ext_access(e, ot, name) {
                return a;
            }
            // VBA spelling of a host member (`s.Name`).
            let js = classes
                .iter()
                .find_map(|(c, _)| self.host_member_ci(*c, &name.text))
                .filter(|js| *js != name.text);
            if let Some(js) = js {
                self.warn_host_case(name, &js);
                return self.member_access(
                    e,
                    object,
                    ot,
                    &Ident {
                        text: js,
                        span: name.span,
                    },
                );
            }
            let names: Vec<String> = missing
                .iter()
                .map(|c| self.prog.class_display(*c))
                .collect();
            let hosts_only = missing.iter().all(|c| {
                let d = &self.prog.classes[*c];
                d.kind == ClassKind::Host
                    && !self.prog.modules[d.module]
                        .form
                        .as_ref()
                        .is_some_and(|f| f.class == Some(*c))
            });
            if hosts_only && found.is_empty() {
                self.warn(
                    name.span,
                    "W0301",
                    format!(
                        "`{}` の型定義に `{}` はありません。遅延バインディングとして扱います",
                        names.join("` / `"),
                        name.text
                    ),
                );
                self.set_res(e.id, Res::Dynamic);
                return Access::Dynamic;
            }
            if hosts_only {
                // `activeSheet.cells(1, 1)`: typed by the members that have it, called by late binding.
                self.warn(
                    name.span,
                    "W0305",
                    format!("`{}` は `{}` にないため、値が `{}` のときは実行時エラーになります（ユニオン型 `{shown}` のうち一部の型のメンバーです。`instanceof` で型を絞り込んでください）", name.text, names.join("` / `"), names.join("` / `")),
                );
            } else {
                self.err(
                name.span,
                "E0202",
                    format!("`{}` に `{}` はありません（ユニオン型 `{shown}` では、すべての型に共通するメンバーだけを参照できます。`instanceof` で型を絞り込んでください）", names.join("` / `"), name.text),
                );
                return Access::Error;
            }
        }
        for m in &found {
            let (owner, public) = match m {
                UnionMember::Value { owner, public, .. }
                | UnionMember::Method { owner, public, .. } => (*owner, *public),
            };
            if !public && self.fx.class != Some(owner) {
                self.err(
                    name.span,
                    "E0392",
                    format!(
                        "`{}` は private なのでクラスの外からは参照できません",
                        name.text
                    ),
                );
                return Access::Error;
            }
        }
        if found.iter().all(|m| matches!(m, UnionMember::Value { .. })) {
            let mut res = None;
            let mut tys = Vec::new();
            let mut writable = true;
            for m in found {
                if let UnionMember::Value {
                    res: r,
                    ty,
                    writable: w,
                    ..
                } = m
                {
                    res.get_or_insert(r);
                    tys.push(ty);
                    writable &= w;
                }
            }
            let same = tys.windows(2).all(|w| w[0] == w[1]).then(|| tys[0].clone());
            let ty = match &same {
                Some(t) => t.clone(),
                None if tys
                    .iter()
                    .any(|t| matches!(t, Type::Void | Type::Ns(_)) || t.has_struct()) =>
                {
                    self.err(name.span, "E0202", format!("`{}` の型が型ごとに異なるため、ユニオン型 `{shown}` からは参照できません", name.text));
                    return Access::Error;
                }
                None => Type::union(tys),
            };
            self.set_res(e.id, res.unwrap());
            self.fx.union_props.insert(e.id, (writable, same));
            return Access::Value(ty);
        }
        if found
            .iter()
            .all(|m| matches!(m, UnionMember::Method { .. }))
        {
            let mut methods: Vec<(FnId, HashMap<String, Type>)> = found
                .into_iter()
                .filter_map(|m| match m {
                    UnionMember::Method { f, env, .. } => Some((f, env)),
                    _ => None,
                })
                .collect();
            if methods
                .iter()
                .any(|(f, _)| !self.prog.fns[*f].type_params.is_empty())
            {
                self.err(name.span, "E0202", format!("ジェネリックなメソッド `{}` はユニオン型 `{shown}` からは呼び出せません。`instanceof` で型を絞り込んでください", name.text));
                return Access::Error;
            }
            // The signature with the fewest parameters checks the call; the others are compared with it.
            let primary = methods
                .iter()
                .enumerate()
                .min_by_key(|(_, (f, _))| self.prog.fns[*f].params.len())
                .map(|(i, _)| i)
                .unwrap_or(0);
            let (f, env) = methods.remove(primary);
            self.set_res(e.id, Res::Method(f));
            self.fx.union_calls.insert(e.id, methods);
            return Access::Func {
                f,
                env,
                recv: Recv::Instance,
            };
        }
        self.err(name.span, "E0202", format!("`{}` は型によってプロパティとメソッドが異なるため、ユニオン型 `{shown}` からは参照できません", name.text));
        Access::Error
    }

    /// Return type of a method reached through a union (the union of the members' return types).
    fn union_call_ret(&self, callee: ExprId, ret: Type) -> Type {
        let Some(others) = self.fx.union_calls.get(&callee) else {
            return ret;
        };
        let mut rets = vec![ret];
        rets.extend(
            others
                .iter()
                .map(|(g, env)| self.prog.fns[*g].ret_ty().subst(env)),
        );
        if rets.contains(&Type::Void) {
            Type::Void
        } else {
            Type::union(rets)
        }
    }

    /// Arguments of a call through a union-typed receiver, against the methods of the other members.
    fn check_union_call(
        &mut self,
        call: &Expr,
        callee: &Expr,
        name: &str,
        named: bool,
        args: &[Expr],
    ) {
        let Some(others) = self.fx.union_calls.get(&callee.id).cloned() else {
            return;
        };
        let primary = match self.info().res.get(&callee.id) {
            Some(Res::Method(f)) => Some(*f),
            _ => None,
        };
        for (g, env) in others {
            let d = self.prog.fns[g].clone();
            let owner = match d.kind {
                FnKind::Method { class, .. }
                | FnKind::Abstract { class }
                | FnKind::Host { class: Some(class) } => self.prog.class_display(class),
                _ => d.name.clone(),
            };
            let rest = d.params.last().is_some_and(|p| p.rest);
            let (min, max) = (d.required_params(), d.params.len());
            if args.len() < min || (args.len() > max && !rest) {
                let count = if rest {
                    format!("{min} 個以上")
                } else if min == max {
                    format!("{min} 個")
                } else {
                    format!("{min}〜{max} 個")
                };
                self.err(call.span, "E0402", format!("`{owner}.{name}` の引数は {count}です（ユニオン型のすべての型で呼び出せる必要があります）"));
                continue;
            }
            if named {
                if let Some(p) = primary {
                    let pp = &self.prog.fns[p].params;
                    let same = (0..args.len()).all(|i| match (pp.get(i), d.params.get(i)) {
                        (Some(a), Some(b)) => a.name.eq_ignore_ascii_case(&b.name),
                        _ => false,
                    });
                    if !same {
                        self.err(call.span, "E0411", format!("`{owner}.{name}` とは引数の名前・順序が異なるため、名前付き引数は使えません"));
                        continue;
                    }
                }
            }
            for (i, a) in args.iter().enumerate() {
                if matches!(a.kind, ExprKind::Undefined) {
                    continue;
                }
                let Some(p) = d
                    .params
                    .get(i)
                    .or(if rest { d.params.last() } else { None })
                else {
                    continue;
                };
                let want = if p.rest && i + 1 >= d.params.len() {
                    p.ty.elem().cloned().unwrap_or(Type::Any).subst(&env)
                } else {
                    p.ty.subst(&env)
                };
                let at = self.info().ty(a.id);
                let ok = if p.mode == PassMode::ByRef {
                    at == want || want == Type::Any || at.is_error()
                } else {
                    self.assignable(&at, &want)
                };
                if !ok {
                    self.err(a.span, "E0301", format!("`{owner}.{name}` の引数 `{}` には `{}` 型が必要です（`{}` は渡せません）", p.name, self.show(&want), self.show(&at)));
                }
            }
        }
    }

    /// Checks an assignment target and returns its type.
    fn lvalue(&mut self, target: &Expr) -> Type {
        let t = self.expr(target, None);
        let inner = target.peel();
        match &inner.kind {
            ExprKind::Name(id) => {
                let res = self.info().res.get(&inner.id).cloned();
                match res {
                    Some(Res::Local(l)) => {
                        let (kind, crossed) = (
                            self.fx.locals[l].kind.clone(),
                            self.lookup_local(&id.text).map(|x| x.1).unwrap_or_default(),
                        );
                        if !crossed.is_empty() {
                            self.err(target.span, "E0322", format!("アロー関数の中から外側の変数 `{}` は書き換えられません。`for...of` を使ってください", id.text));
                        }
                        match kind {
                            LocalKind::Const | LocalKind::Using | LocalKind::Catch => {
                                self.err(target.span, "E0323", format!("`{}` は `const` のため再代入できません", id.text))
                            }
                            LocalKind::Param if matches!(self.fx.locals[l].ty, Type::Array(_)) && !self.fx.locals[l].byref => self.err(
                                target.span,
                                "E0324",
                                format!("配列の引数 `{}` への再代入は呼び出し元の配列も置き換えます。`byref` を明示するか、ローカル変数にコピーしてください", id.text),
                            ),
                            _ => {}
                        }
                        self.fx.locals[l].reassigned = true;
                    }
                    Some(Res::Global(v)) => {
                        let var = &self.prog.vars[v];
                        if var.is_const || var.declared {
                            self.err(
                                target.span,
                                "E0323",
                                format!("`{}` は再代入できません", id.text),
                            );
                        }
                        // (The handler of an event keeps its state there: spec §12.2.3.)
                        if self.current_lambda().is_some() && self.fx.event_handlers == 0 {
                            self.err(
                                target.span,
                                "E0322",
                                format!(
                                    "アロー関数の中からモジュール変数 `{}` は書き換えられません",
                                    id.text
                                ),
                            );
                        }
                        self.prog.vars[v].reassigned = true;
                    }
                    Some(_) if !t.is_error() => {
                        self.err(target.span, "E0325", "ここには代入できません")
                    }
                    _ => {}
                }
            }
            ExprKind::This => match self.fx.this_local {
                Some(l) if self.fx.locals[l].byref => self.fx.locals[l].reassigned = true,
                _ => self.err(
                    target.span,
                    "E0325",
                    "`this` に代入するには `byref this: T` と宣言してください",
                ),
            },
            ExprKind::Member { name, object, .. } => {
                // A struct returned by a call is a copy: writing to its field would be lost.
                let mut root: &Expr = object.peel();
                let mut through_struct = matches!(self.info().ty(root.id), Type::Struct(_));
                while let ExprKind::Member { object: o, .. } = &root.kind {
                    if !matches!(self.info().ty(o.id), Type::Struct(_)) {
                        break;
                    }
                    root = o.peel();
                    through_struct = true;
                }
                if through_struct && matches!(root.kind, ExprKind::Call { .. }) {
                    self.err(target.span, "E0325", "関数やメソッドが返した構造体はコピーなので、フィールドに代入しても元は変わりません。変数に受けて書き換え、必要なら設定し直してください");
                }
                if let Some((writable, same)) = self.fx.union_props.get(&inner.id).cloned() {
                    if !writable {
                        self.err(
                            name.span,
                            "E0323",
                            format!(
                                "`{}` はユニオン型のいずれかの型で読み取り専用です",
                                name.text
                            ),
                        );
                    } else if same.is_none() {
                        self.err(name.span, "E0325", format!("`{}` の型がユニオン型の型ごとに異なるため代入できません。`instanceof` で型を絞り込んでください", name.text));
                    }
                    return t;
                }
                let res = self.info().res.get(&inner.id).cloned();
                match res {
                    Some(Res::Field(c, i)) => {
                        let f = &self.prog.classes[c].fields[i];
                        let in_ctor = self.fx.class == Some(c)
                            && self.fx.fn_id.is_some_and(|f| {
                                matches!(self.prog.fns[f].kind, FnKind::Ctor { .. })
                            });
                        if f.readonly && !in_ctor {
                            self.err(
                                name.span,
                                "E0323",
                                format!("`{}` は読み取り専用です", name.text),
                            );
                        }
                    }
                    Some(Res::Property { set: None, .. }) => self.err(
                        name.span,
                        "E0323",
                        format!("`{}` には setter がありません", name.text),
                    ),
                    Some(Res::Property { .. } | Res::Dynamic | Res::Global(_)) => {}
                    Some(_) if !t.is_error() => {
                        self.err(target.span, "E0325", "ここには代入できません")
                    }
                    _ => {}
                }
            }
            ExprKind::Index { object, .. } => {
                let ot = self.info().ty(object.id);
                if ot == Type::Str {
                    self.err(
                        target.span,
                        "E0325",
                        "文字列の 1 文字を書き換えることはできません",
                    );
                }
                // Writing an element of a captured array from a lambda mutates outer state.
                if let ExprKind::Name(id) = &object.peel().kind {
                    if let Some((l, crossed)) = self.lookup_local(&id.text) {
                        if !crossed.is_empty() {
                            self.err(
                                target.span,
                                "E0322",
                                format!(
                                    "アロー関数の中から外側の配列 `{}` は書き換えられません",
                                    id.text
                                ),
                            );
                        }
                        self.fx.locals[l].reassigned |= false;
                    }
                }
            }
            ExprKind::Error => {}
            _ => {
                if !t.is_error() {
                    self.err(target.span, "E0325", "ここには代入できません");
                }
            }
        }
        t
    }

    fn binary(&mut self, e: &Expr, op: BinaryOp, lhs: &Expr, rhs: &Expr) -> Type {
        match op {
            BinaryOp::And | BinaryOp::Or => {
                self.cond(lhs);
                let (t, f) = self.guard_facts(lhs, &Vec::new());
                let facts = if op == BinaryOp::And { t } else { f };
                self.with_facts(facts, &mut |n| assigns_in_expr(rhs, n), |ck| ck.cond(rhs));
                Type::Bool
            }
            BinaryOp::Coalesce => {
                let a = self.value(lhs, None);
                let b = self.value(rhs, Some(&a));
                if !(a.is_object() || matches!(a, Type::Any | Type::Error)) {
                    self.err(lhs.span, "E0394", format!("`??` の左辺はオブジェクト型か `any` である必要があります（`{}` は null になりません）", self.show(&a)));
                    return Type::Error;
                }
                if a == Type::Any || a == Type::Null {
                    return if b == Type::Null { Type::Any } else { b };
                }
                self.coerce(rhs, &b, &a);
                a
            }
            _ => {
                let a = self.value(lhs, None);
                let b = self.value(rhs, None);
                if matches!(op, BinaryOp::Eq | BinaryOp::Ne) {
                    self.lint_typeof_name(lhs, rhs);
                }
                self.binary_types(e.span, op, &a, &b, lhs, rhs)
            }
        }
    }

    /// `typeof x == "string"`: RAS's `typeof` is VBA's `TypeName` (`"String"`, `"Long"`, `"Worksheet"` ...).
    fn lint_typeof_name(&mut self, lhs: &Expr, rhs: &Expr) {
        let (lit, other) = match (&lhs.peel().kind, &rhs.peel().kind) {
            (
                ExprKind::Unary {
                    op: UnaryOp::TypeOf,
                    ..
                },
                ExprKind::Str(s),
            ) => (rhs, s),
            (
                ExprKind::Str(s),
                ExprKind::Unary {
                    op: UnaryOp::TypeOf,
                    ..
                },
            ) => (lhs, s),
            _ => return,
        };
        let hint = match other.as_str() {
            "string" => "\"String\"",
            "number" => "\"Long\" / \"Double\" など",
            "boolean" => "\"Boolean\"",
            "object" => "\"Object\" かクラス名（`instanceof` を使うと型が絞り込まれます）",
            "undefined" => "\"Empty\"",
            "function" | "bigint" | "symbol" => "VBA の型名",
            _ => return,
        };
        self.warn(lit.span, "W0312", format!("`typeof` は VBA の `TypeName` と同じ型名を返すため、`\"{other}\"` とは一致しません。{hint} と比較してください"));
    }

    fn binary_types(
        &mut self,
        span: Span,
        op: BinaryOp,
        a: &Type,
        b: &Type,
        lhs: &Expr,
        rhs: &Expr,
    ) -> Type {
        if a.is_error() || b.is_error() {
            return if op.is_comparison() {
                Type::Bool
            } else {
                Type::Error
            };
        }
        match op {
            BinaryOp::Add if *a == Type::Str || *b == Type::Str => {
                let other = if *a == Type::Str { b } else { a };
                if !self.stringifiable(other) {
                    self.err(
                        span,
                        "E0381",
                        format!("`{}` は文字列と連結できません", self.show(other)),
                    );
                }
                Type::Str
            }
            BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div | BinaryOp::Rem => {
                if (!self.is_numeric(a) || !self.is_numeric(b))
                    && !(matches!(a, Type::Date) || matches!(b, Type::Date))
                {
                    self.err(
                        span,
                        "E0387",
                        format!(
                            "演算子 `{}` は `{}` と `{}` には使えません",
                            op.symbol(),
                            self.show(a),
                            self.show(b)
                        ),
                    );
                    return Type::Error;
                }
                if op == BinaryOp::Sub && *a == Type::Date && *b == Type::Date {
                    return Type::Double;
                }
                match self.promote(a, b) {
                    Some(t) => {
                        self.check_const_overflow(span, op, &t, lhs, rhs);
                        t
                    }
                    None => {
                        self.err(
                            span,
                            "E0395",
                            format!(
                                "`{}` と `{}` の混在した演算には明示的な型変換（`as`）が必要です",
                                self.show(a),
                                self.show(b)
                            ),
                        );
                        Type::Error
                    }
                }
            }
            BinaryOp::Pow => {
                if !self.is_numeric(a) || !self.is_numeric(b) {
                    self.err(span, "E0387", "`**` は数値にのみ使えます");
                }
                Type::Double
            }
            BinaryOp::Eq | BinaryOp::Ne => {
                if !self.comparable(a, b, true) {
                    self.err(
                        span,
                        "E0396",
                        format!("`{}` と `{}` は比較できません", self.show(a), self.show(b)),
                    );
                }
                Type::Bool
            }
            BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
                if !self.comparable(a, b, false) {
                    self.err(
                        span,
                        "E0396",
                        format!(
                            "`{}` と `{}` は大小比較できません",
                            self.show(a),
                            self.show(b)
                        ),
                    );
                }
                Type::Bool
            }
            BinaryOp::BitAnd
            | BinaryOp::BitOr
            | BinaryOp::BitXor
            | BinaryOp::Shl
            | BinaryOp::Shr => {
                let int_like = |t: &Type| t.is_integer() || matches!(t, Type::Enum(_) | Type::Any);
                if !int_like(a) || !int_like(b) {
                    self.err(
                        span,
                        "E0387",
                        format!("ビット演算子 `{}` は整数にのみ使えます", op.symbol()),
                    );
                    return Type::Error;
                }
                if matches!(op, BinaryOp::Shl | BinaryOp::Shr) {
                    return if *a == Type::Long {
                        Type::Long
                    } else {
                        Type::Int
                    };
                }
                self.promote(a, b).unwrap_or(Type::Int)
            }
            BinaryOp::And | BinaryOp::Or | BinaryOp::Coalesce => Type::Bool,
        }
    }

    /// Constant integer expressions that overflow are compile errors (spec §2.3).
    fn check_const_overflow(
        &mut self,
        span: Span,
        op: BinaryOp,
        ty: &Type,
        lhs: &Expr,
        rhs: &Expr,
    ) {
        let (Some(a), Some(b)) = (const_int(lhs), const_int(rhs)) else {
            return;
        };
        let r = match op {
            BinaryOp::Add => a.checked_add(b),
            BinaryOp::Sub => a.checked_sub(b),
            BinaryOp::Mul => a.checked_mul(b),
            BinaryOp::Div | BinaryOp::Rem if b == 0 => {
                self.err(span, "E0397", "定数式でのゼロ除算です");
                return;
            }
            _ => return,
        };
        let fits = match (r, ty) {
            (Some(v), Type::Int) => i32::try_from(v).is_ok(),
            (Some(v), Type::Long) => i64::try_from(v).is_ok(),
            (None, _) => false,
            _ => true,
        };
        if !fits {
            self.err(span, "E0397", format!("定数式が `{}` の範囲を超えます（オーバーフロー）。片方を `long`（`123L`）か `double` にしてください", self.show(ty)));
        }
    }

    // ------------------------------------------------------------------ lambdas

    fn arrow(&mut self, e: &Expr, arrow: &ast::ArrowFn, expect: Option<&Type>) -> Type {
        if self.fx.in_field_init {
            self.err(e.span, "E0398", "フィールドの初期化式ではアロー関数を使えません。コンストラクタで初期化してください");
        }
        let expected = match expect {
            Some(Type::Func(f)) => Some((**f).clone()),
            _ => None,
        };
        let tps = self.fx.type_params.clone();
        self.fx.scopes.push(Scope {
            names: HashMap::new(),
            lambda: Some(e.id),
        });
        let mut param_types = Vec::new();
        let mut locals = Vec::new();
        for (i, p) in arrow.params.iter().enumerate() {
            let ty = match &p.ty {
                Some(t) => resolve_type(self.prog, self.module, &tps, t),
                None => match expected.as_ref().and_then(|f| f.params.get(i)) {
                    Some(t)
                        if !t.contains_param()
                            || self
                                .fx
                                .type_params
                                .iter()
                                .any(|tp| matches!(t, Type::Param(n) if *n == tp.name)) =>
                    {
                        t.clone()
                    }
                    Some(t) if !t.contains_param() => t.clone(),
                    _ => {
                        self.err(p.name.span, "E0150", format!("引数 `{}` の型を推論できません。`({}: 型) => ...` のように型注釈を付けてください", p.name.text, p.name.text));
                        Type::Error
                    }
                },
            };
            if p.mode == PassMode::ByRef {
                self.err(p.span, "E0399", "アロー関数の引数に `byref` は使えません");
            }
            let l = self.declare_local(&p.name, ty.clone(), LocalKind::LambdaParam);
            locals.push(l);
            param_types.push(ty);
        }
        if let Some(f) = &expected {
            if arrow.params.len() > f.params.len() {
                self.err(
                    arrow.span,
                    "E0400",
                    format!("アロー関数の引数が多すぎます（最大 {} 個）", f.params.len()),
                );
            }
        }
        self.info().lambda_params.insert(e.id, locals);
        self.info().lambda_captures.entry(e.id).or_default();
        let declared_ret = arrow
            .ret
            .as_ref()
            .map(|t| resolve_type(self.prog, self.module, &tps, t));
        let expected_ret = declared_ret.clone().or_else(|| {
            expected.as_ref().map(|f| f.ret.clone()).filter(|r| {
                !r.contains_param()
                    || self
                        .fx
                        .type_params
                        .iter()
                        .any(|tp| matches!(r, Type::Param(n) if *n == tp.name))
            })
        });
        let saved_loops = std::mem::take(&mut self.fx.loops);
        let saved_switch = std::mem::replace(&mut self.fx.switch_depth, 0);
        let ret = match &arrow.body {
            ast::ArrowBody::Expr(body) => {
                let t = self.expr(body, expected_ret.as_ref());
                match &expected_ret {
                    Some(want) if *want != Type::Void => {
                        if t == Type::Void {
                            self.err(body.span, "E0370", "この式は値を返しません（`void`）");
                        } else {
                            self.coerce(body, &t, want);
                        }
                        want.clone()
                    }
                    Some(want) => want.clone(),
                    None => {
                        if matches!(t, Type::Ns(_)) {
                            self.err(body.span, "E0370", "ここには値が必要です");
                            Type::Error
                        } else {
                            t
                        }
                    }
                }
            }
            ast::ArrowBody::Block(b) => {
                self.fx.lambda_rets.push(LambdaRet {
                    expected: expected_ret.clone(),
                    inferred: None,
                });
                self.block(b);
                let lr = self.fx.lambda_rets.pop().unwrap();
                expected_ret.clone().or(lr.inferred).unwrap_or(Type::Void)
            }
        };
        self.fx.loops = saved_loops;
        self.fx.switch_depth = saved_switch;
        self.fx.scopes.pop();
        Type::Func(Box::new(FuncType {
            params: param_types,
            ret,
        }))
    }

    // ------------------------------------------------------------------ calls

    fn call(
        &mut self,
        e: &Expr,
        callee: &Expr,
        type_args: &[ast::TypeRef],
        args: &[Expr],
        expect: Option<&Type>,
    ) -> Type {
        let orig_args = args;
        let c = callee.peel();
        let (access, recv): (Access, Option<(&Expr, Type)>) = match &c.kind {
            ExprKind::Name(id) => {
                let a = self.name_access(c, id);
                if let Access::Value(t) = &a {
                    let t = t.clone();
                    if let Some(slot) = self.info().expr_types.get_mut(c.id as usize) {
                        *slot = t;
                    }
                }
                (a, None)
            }
            ExprKind::Member {
                object,
                name,
                optional,
            } => {
                if *optional {
                    self.err(
                        c.span,
                        "E0388",
                        "オプショナルチェーン（`?.`）には対応していません",
                    );
                }
                let ot = self.expr(object, None);
                let a = self.member_access(c, object, &ot, name);
                if let Access::Value(t) = &a {
                    let t = t.clone();
                    if let Some(slot) = self.info().expr_types.get_mut(c.id as usize) {
                        *slot = t;
                    }
                }
                (a, Some((&**object, ot)))
            }
            _ => (Access::Value(self.value(callee, None)), None),
        };
        if !type_args.is_empty() && !matches!(access, Access::Func { .. }) {
            self.err(
                e.span,
                "E0401",
                "型引数はジェネリック関数の呼び出しにのみ指定できます",
            );
        }
        let positional;
        let args: &[Expr] = if has_named(args) {
            let sig = match &access {
                Access::Func { f, recv, .. } => Some((*f, usize::from(*recv == Recv::Extension))),
                Access::Value(Type::Class(cid, _)) => {
                    self.prog.classes[*cid].call_sig.map(|f| (f, 0))
                }
                _ => None,
            };
            match sig {
                Some((f, skip)) => match self.reorder_named(e, f, skip, args) {
                    Some(v) => {
                        positional = v;
                        &positional
                    }
                    None => return Type::Error,
                },
                None => {
                    if !matches!(access, Access::Error | Access::Value(Type::Error)) {
                        self.err(e.span, "E0411", "名前付き引数は、宣言された関数・メソッド・コンストラクタの呼び出しにだけ書けます");
                    }
                    for a in args {
                        self.expr(named_value(a), None);
                    }
                    return Type::Error;
                }
            }
        } else {
            args
        };
        match access {
            Access::Error => {
                for a in args {
                    self.expr(a, None);
                }
                Type::Error
            }
            Access::Dynamic => {
                for a in args {
                    if !matches!(a.kind, ExprKind::Undefined) {
                        self.value(a, None);
                    }
                }
                self.info().calls.insert(
                    e.id,
                    CallInfo {
                        target: CallTarget::Dynamic,
                        type_args: Vec::new(),
                    },
                );
                Type::Any
            }
            Access::Intrinsic(i) => {
                self.info().calls.insert(
                    e.id,
                    CallInfo {
                        target: CallTarget::Intrinsic(i),
                        type_args: Vec::new(),
                    },
                );
                let (robj, rty) = match &recv {
                    Some((o, t)) if i == Intr::EventOn => (Some(*o), Some(t.clone())),
                    // A host collection is seen as an array of its members.
                    Some((o, t)) => (
                        Some(*o),
                        Some(
                            self.prog
                                .collection_elem(t)
                                .map(Type::array)
                                .unwrap_or_else(|| t.clone()),
                        ),
                    ),
                    None => (None, None),
                };
                self.intrinsic_call(e, i, robj, rty.as_ref(), args, expect)
            }
            Access::Func { f, env, recv: rk } => {
                let def_kind = self.prog.fns[f].kind.clone();
                let target = match (&def_kind, rk) {
                    (_, Recv::Extension) => CallTarget::Ext(f),
                    (FnKind::Host { .. }, _) => CallTarget::Host(f),
                    (_, Recv::Instance) => CallTarget::Method(f),
                    _ => CallTarget::Fn(f),
                };
                let qualify = matches!(&def_kind, FnKind::Host { class: Some(_) })
                    && self.prog.fns[f].vba_name() == "Range";
                if qualify {
                    self.fx.qualify_depth += 1;
                }
                let recv_arg = if rk == Recv::Extension { recv } else { None };
                let (ret, targs) = self.check_fn_call(e, f, env, recv_arg, type_args, args);
                let ret = if self.fx.union_calls.contains_key(&c.id) {
                    let name = match &c.kind {
                        ExprKind::Member { name, .. } => name.text.clone(),
                        _ => String::new(),
                    };
                    let named = !std::ptr::eq(args.as_ptr(), orig_args.as_ptr());
                    self.check_union_call(e, c, &name, named, args);
                    self.union_call_ret(c.id, ret)
                } else {
                    ret
                };
                if qualify {
                    self.fx.qualify_depth -= 1;
                }
                if let Some(cur) = self.fx.fn_id {
                    let _ = cur;
                    if !self.fx.calls.contains(&f) {
                        self.fx.calls.push(f);
                    }
                }
                self.lint_member_call(c);
                self.info().calls.insert(
                    e.id,
                    CallInfo {
                        target,
                        type_args: targs,
                    },
                );
                ret
            }
            Access::Value(t) => match &t {
                Type::Func(ft) => {
                    if args.len() != ft.params.len() {
                        self.err(
                            e.span,
                            "E0402",
                            format!(
                                "引数の数が合いません（{} 個必要ですが {} 個渡されています）",
                                ft.params.len(),
                                args.len()
                            ),
                        );
                    }
                    for (i, a) in args.iter().enumerate() {
                        let pt = ft.params.get(i).cloned();
                        let at = self.value(a, pt.as_ref());
                        if let Some(pt) = pt {
                            self.coerce(a, &at, &pt);
                        }
                    }
                    self.info().calls.insert(
                        e.id,
                        CallInfo {
                            target: CallTarget::Closure,
                            type_args: Vec::new(),
                        },
                    );
                    ft.ret.clone()
                }
                Type::Class(cid, cargs) if self.prog.classes[*cid].call_sig.is_some() => {
                    let sig = self.prog.classes[*cid].call_sig.unwrap();
                    let env = self.prog.class_env(*cid, cargs);
                    let (ret, _) = self.check_fn_call(e, sig, env, None, &[], args);
                    self.info().calls.insert(
                        e.id,
                        CallInfo {
                            target: CallTarget::DefaultMember(sig),
                            type_args: Vec::new(),
                        },
                    );
                    ret
                }
                Type::Any | Type::Object => {
                    for a in args {
                        self.value(a, None);
                    }
                    self.info().calls.insert(
                        e.id,
                        CallInfo {
                            target: CallTarget::Dynamic,
                            type_args: Vec::new(),
                        },
                    );
                    Type::Any
                }
                Type::Ns(NsRef::Class(cid)) => {
                    let name = self.prog.classes[*cid].name.clone();
                    for a in args {
                        self.expr(a, None);
                    }
                    self.err(
                        e.span,
                        "E0403",
                        format!(
                            "クラスのインスタンスを作るには `new {name}(...)` と書いてください"
                        ),
                    );
                    Type::Error
                }
                Type::Error => {
                    for a in args {
                        self.expr(a, None);
                    }
                    Type::Error
                }
                other => {
                    let shown = self.show(other);
                    for a in args {
                        self.expr(a, None);
                    }
                    self.err(callee.span, "E0403", format!("`{shown}` は呼び出せません"));
                    Type::Error
                }
            },
        }
    }

    /// Lint rule 10: `.Select` / `.Activate`.
    fn lint_member_call(&mut self, callee: &Expr) {
        if let ExprKind::Member { name, .. } = &callee.kind {
            let Some(Res::Method(m)) = self.info().res.get(&callee.id).cloned() else {
                return;
            };
            let host = self.prog.fns[m].vba_name().to_string();
            if matches!(host.as_str(), "Select" | "Activate") {
                self.diag(
                    Severity::Info,
                    name.span,
                    "I0402",
                    format!(
                        "`.{}` は遅く不安定です。対象のオブジェクトを直接操作してください",
                        name.text
                    ),
                );
            }
        }
    }

    /// Checks the arguments of a call to `f`; returns the return type and the type arguments.
    fn check_fn_call(
        &mut self,
        call: &Expr,
        f: FnId,
        mut env: HashMap<String, Type>,
        recv: Option<(&Expr, Type)>,
        explicit: &[ast::TypeRef],
        args: &[Expr],
    ) -> (Type, Vec<Type>) {
        let def = self.prog.fns[f].clone();
        let tparams = def.type_params.clone();
        if !explicit.is_empty() {
            if explicit.len() != tparams.len() {
                self.err(
                    call.span,
                    "E0401",
                    format!(
                        "`{}` の型引数は {} 個です（{} 個指定されています）",
                        def.name,
                        tparams.len(),
                        explicit.len()
                    ),
                );
            }
            let tps = self.fx.type_params.clone();
            for (tp, t) in tparams.iter().zip(explicit.iter()) {
                let ty = resolve_type(self.prog, self.module, &tps, t);
                env.insert(tp.name.clone(), ty);
            }
        }
        let mut params: &[ParamDef] = &def.params;
        if let Some((robj, rty)) = &recv {
            let p0 = &params[0];
            self.unify(&p0.ty, rty, &mut env, &tparams);
            if p0.mode == PassMode::ByRef {
                self.byref_arg(robj, rty, &p0.ty.subst(&env), "レシーバ");
            }
            params = &params[1..];
        }
        let required = params
            .iter()
            .filter(|p| !p.optional && p.default.is_none())
            .count();
        if args.len() < required || args.len() > params.len() {
            let want = if required == params.len() {
                format!("{required}")
            } else {
                format!("{required}〜{}", params.len())
            };
            let shown = if def.name.is_empty() {
                "既定メンバー".to_string()
            } else {
                format!("`{}`", def.name)
            };
            self.err(
                call.span,
                "E0402",
                format!(
                    "{shown} の引数の数が合いません（{want} 個必要ですが {} 個渡されています）",
                    args.len()
                ),
            );
        }
        let mut arg_types: Vec<Option<Type>> = vec![None; args.len()];
        // pass 1: everything except lambdas, so that their parameter types can be inferred
        for (i, a) in args.iter().enumerate() {
            let Some(p) = params.get(i) else {
                self.expr(a, None);
                continue;
            };
            if matches!(a.kind, ExprKind::Undefined) && (p.optional || p.default.is_some()) {
                self.expr(a, None);
                continue;
            }
            if is_lambda(a) {
                continue;
            }
            let pt = p.ty.subst(&env);
            let has_unbound = self.has_unbound(&pt, &tparams, &env);
            let at = self.value(a, if has_unbound { None } else { Some(&pt) });
            self.unify(&p.ty, &at, &mut env, &tparams);
            arg_types[i] = Some(at);
        }
        // pass 2: lambdas
        for (i, a) in args.iter().enumerate() {
            let Some(p) = params.get(i) else { continue };
            if !is_lambda(a) {
                continue;
            }
            let pt = self.subst_bound(&p.ty, &tparams, &env);
            let at = self.value(a, Some(&pt));
            self.unify(&p.ty, &at, &mut env, &tparams);
            arg_types[i] = Some(at);
        }
        let mut targs = Vec::new();
        for tp in &tparams {
            match env.get(&tp.name) {
                Some(t) => {
                    if !self.satisfies(t, &tp.constraint) {
                        let c = match &tp.constraint {
                            Constraint::Comparable => "Comparable".to_string(),
                            Constraint::Numeric => "Numeric".to_string(),
                            Constraint::ObjectType => "ObjectType".to_string(),
                            Constraint::Type(t) => self.show(t),
                            Constraint::None => String::new(),
                        };
                        self.err(
                            call.span,
                            "E0404",
                            format!(
                                "型引数 `{} = {}` は制約 `{c}` を満たしません",
                                tp.name,
                                self.show(t)
                            ),
                        );
                    }
                    targs.push(t.clone());
                }
                // `expect(null)`: nothing to infer from.
                None if self.prog.modules[def.module].key == crate::collect::TEST_DEFS_MODULE => {
                    env.insert(tp.name.clone(), Type::Any);
                    targs.push(Type::Any);
                }
                None => {
                    self.err(call.span, "E0405", format!("`{}` の型引数 `{}` を推論できません。`{}<型>(...)` のように明示してください", def.name, tp.name, def.name));
                    env.insert(tp.name.clone(), Type::Error);
                    targs.push(Type::Error);
                }
            }
        }
        if let Some((robj, rty)) = &recv {
            let want = def.params[0].ty.subst(&env);
            if !self.assignable(rty, &want) {
                self.err(
                    robj.span,
                    "E0301",
                    format!(
                        "レシーバの型 `{}` は `{}` ではありません",
                        self.show(rty),
                        self.show(&want)
                    ),
                );
            }
        }
        for (i, a) in args.iter().enumerate() {
            let (Some(p), Some(at)) = (params.get(i), &arg_types[i]) else {
                continue;
            };
            let want = p.ty.subst(&env);
            if p.mode == PassMode::ByRef {
                self.byref_arg(a, at, &want, &format!("`byref` 引数 `{}`", p.name));
            } else if matches!(want, Type::Array(_)) && matches!(at, Type::Array(_)) {
                // VBA passes arrays by reference: the element types must match exactly.
                if *at != want && !at.is_error() && !want.is_error() {
                    self.err(a.span, "E0301", format!("配列 `{}` を `{}` の引数に渡せません（要素の型が一致している必要があります）", self.show(at), self.show(&want)));
                }
            } else {
                self.coerce(a, at, &want);
            }
        }
        let ret = self.fn_ret(f, call.span).subst(&env);
        // `createObject("Scripting.FileSystemObject")`: the type of what a ProgID written out creates.
        if def.host_name.as_deref() == Some("CreateObject")
            && self.prog.modules[def.module]
                .rel_path
                .starts_with("<stdlib>/")
        {
            if let [arg] = args {
                if let ExprKind::Str(id) = &arg.peel().kind {
                    if let Some(c) = self.prog.class_of_prog_id(id) {
                        return (Type::Class(c, Vec::new()), targs);
                    }
                }
            }
        }
        (ret, targs)
    }

    fn has_unbound(&self, t: &Type, tparams: &[TypeParamDef], env: &HashMap<String, Type>) -> bool {
        match t {
            Type::Param(n) => tparams.iter().any(|p| &p.name == n) && !env.contains_key(n),
            Type::Array(e) => self.has_unbound(e, tparams, env),
            Type::Class(_, args) => args.iter().any(|a| self.has_unbound(a, tparams, env)),
            Type::Map(k, v) => {
                self.has_unbound(k, tparams, env) || self.has_unbound(v, tparams, env)
            }
            Type::Func(f) => {
                f.params.iter().any(|p| self.has_unbound(p, tparams, env))
                    || self.has_unbound(&f.ret, tparams, env)
            }
            _ => false,
        }
    }

    /// Substitutes bound type parameters; unbound ones of the callee become `Error` (to be inferred).
    fn subst_bound(&self, t: &Type, tparams: &[TypeParamDef], env: &HashMap<String, Type>) -> Type {
        match t {
            Type::Param(n) if tparams.iter().any(|p| &p.name == n) => {
                env.get(n).cloned().unwrap_or(Type::Param(format!("?{n}")))
            }
            Type::Array(e) => Type::Array(Box::new(self.subst_bound(e, tparams, env))),
            Type::Class(c, args) => Type::Class(
                *c,
                args.iter()
                    .map(|a| self.subst_bound(a, tparams, env))
                    .collect(),
            ),
            Type::Map(k, v) => Type::Map(
                Box::new(self.subst_bound(k, tparams, env)),
                Box::new(self.subst_bound(v, tparams, env)),
            ),
            Type::Func(f) => Type::Func(Box::new(FuncType {
                params: f
                    .params
                    .iter()
                    .map(|p| self.subst_bound(p, tparams, env))
                    .collect(),
                ret: self.subst_bound(&f.ret, tparams, env),
            })),
            other => other.subst(env),
        }
    }

    /// spec §6 rule 5: `byref` needs an assignable variable of exactly the parameter's type.
    fn byref_arg(&mut self, a: &Expr, at: &Type, want: &Type, what: &str) {
        let inner = a.peel();
        let is_lvalue = match &inner.kind {
            ExprKind::Name(_) => matches!(
                self.info().res.get(&inner.id),
                Some(Res::Local(_) | Res::Global(_))
            ),
            ExprKind::Member { .. } => matches!(
                self.info().res.get(&inner.id),
                Some(Res::Field(..) | Res::Dynamic | Res::Global(_))
            ),
            ExprKind::Index { .. } | ExprKind::This => true,
            ExprKind::Error => true,
            _ => false,
        };
        if !is_lvalue {
            self.err(
                a.span,
                "E0406",
                format!("{what} にはリテラルや式ではなく変数を渡してください（参照渡しです）"),
            );
            return;
        }
        if let Some(l) = self.guarded_local(a) {
            if self.local_ty(l) != self.fx.locals[l].ty {
                let name = self.fx.locals[l].name.clone();
                self.err(a.span, "E0415", format!("型が絞り込まれている `{name}` は {what} に渡せません（参照渡しで書き換えられると絞り込みが成り立たなくなります）。別の変数に代入して渡してください"));
                return;
            }
        }
        if at != want && !at.is_error() && !want.is_error() && *want != Type::Any {
            self.err(a.span, "E0407", format!("{what} には `{}` 型の変数が必要です（`{}` は渡せません。参照渡しでは型が完全に一致している必要があります）", self.show(want), self.show(at)));
        }
        // The callee may assign through the reference.
        if let ExprKind::Name(id) = &inner.kind {
            match self.info().res.get(&inner.id).cloned() {
                Some(Res::Local(l)) => {
                    let (kind, crossed) = (
                        self.fx.locals[l].kind.clone(),
                        self.lookup_local(&id.text).map(|x| x.1).unwrap_or_default(),
                    );
                    if kind == LocalKind::Const {
                        self.err(
                            a.span,
                            "E0323",
                            format!("`const` の `{}` は `byref` 引数に渡せません", id.text),
                        );
                    }
                    if !crossed.is_empty() {
                        self.err(a.span, "E0322", format!("アロー関数の中から外側の変数 `{}` を `byref` で渡すことはできません", id.text));
                    }
                    self.fx.locals[l].reassigned = true;
                }
                Some(Res::Global(v)) => {
                    if self.prog.vars[v].is_const {
                        self.err(
                            a.span,
                            "E0323",
                            format!("`const` の `{}` は `byref` 引数に渡せません", id.text),
                        );
                    }
                    self.prog.vars[v].reassigned = true;
                }
                _ => {}
            }
        }
    }

    fn new_expr(&mut self, e: &Expr, ty: &ast::TypeRef, args: &[Expr]) -> Type {
        if !has_named(args) {
            return self.new_expr_inner(e, ty, args);
        }
        let builtin = matches!(ty, ast::TypeRef::Named { path, .. } if path.len() == 1 && matches!(path[0].text.as_str(), "Error" | "Map" | "Array"));
        let resolved = if builtin {
            None
        } else {
            // Resolved again (with its diagnostics) by `new_expr_inner`.
            let n = self.prog.diagnostics.len();
            let tps = self.fx.type_params.clone();
            let t = resolve_type(self.prog, self.module, &tps, ty);
            self.prog.diagnostics.truncate(n);
            Some(t)
        };
        let ctor = match &resolved {
            Some(Type::Class(c, _)) if self.prog.classes[*c].kind == ClassKind::Class => {
                self.prog.classes[*c].ctor
            }
            _ => None,
        };
        match ctor {
            Some(f) => match self.reorder_named(e, f, 0, args) {
                Some(v) => self.new_expr_inner(e, ty, &v),
                None => Type::Error,
            },
            None => {
                if !matches!(resolved, Some(Type::Error)) {
                    self.err(e.span, "E0411", "名前付き引数は、宣言された関数・メソッド・コンストラクタの呼び出しにだけ書けます");
                }
                let values: Vec<Expr> = args.iter().map(|a| named_value(a).clone()).collect();
                self.new_expr_inner(e, ty, &values);
                Type::Error
            }
        }
    }

    /// Puts the arguments of a call with named arguments into parameter order (spec §3.3).
    /// `skip` leading parameters are not written in the argument list (the receiver of an
    /// extension method). Returns `None` after reporting an error; every value is then checked.
    fn reorder_named(
        &mut self,
        call: &Expr,
        f: FnId,
        skip: usize,
        args: &[Expr],
    ) -> Option<Rc<[Expr]>> {
        let def = &self.prog.fns[f];
        let params: Vec<ParamDef> = def.params.iter().skip(skip).cloned().collect();
        let shown = if def.name.is_empty() {
            "既定メンバー".to_string()
        } else {
            format!("`{}`", def.name)
        };
        let mut slots: Vec<Option<Expr>> = Vec::new();
        let mut hole = None;
        let mut ok = true;
        let mut named_seen = false;
        for (i, a) in args.iter().enumerate() {
            let ExprKind::NamedArg {
                name,
                value,
                hole: h,
            } = &a.kind
            else {
                if named_seen {
                    self.err(
                        a.span,
                        "E0413",
                        "名前付き引数の後ろに、名前のない引数は書けません",
                    );
                    ok = false;
                } else {
                    slots.resize(slots.len().max(i + 1), None);
                    slots[i] = Some(a.clone());
                }
                continue;
            };
            named_seen = true;
            hole.get_or_insert(*h);
            match params.iter().position(|p| p.name == name.text) {
                Some(j) if params[j].rest => {
                    self.err(
                        name.span,
                        "E0412",
                        format!("可変長引数 `...{}` には名前を付けて渡せません", name.text),
                    );
                    ok = false;
                }
                Some(j) => {
                    slots.resize(slots.len().max(j + 1), None);
                    if slots[j].is_some() {
                        self.err(
                            name.span,
                            "E0413",
                            format!("引数 `{}` が二重に指定されています", name.text),
                        );
                        ok = false;
                    } else {
                        slots[j] = Some((**value).clone());
                    }
                }
                None => {
                    let msg = match params
                        .iter()
                        .find(|p| p.name.eq_ignore_ascii_case(&name.text))
                    {
                        Some(p) => format!(
                            "{shown} に引数 `{}` はありません。`{}` ではありませんか",
                            name.text, p.name
                        ),
                        None if params.is_empty() => format!("{shown} は引数を取りません"),
                        None => {
                            let list = params
                                .iter()
                                .map(|p| format!("`{}`", p.name))
                                .collect::<Vec<_>>()
                                .join(", ");
                            format!(
                                "{shown} に引数 `{}` はありません（引数: {list}）",
                                name.text
                            )
                        }
                    };
                    self.err(name.span, "E0412", msg);
                    ok = false;
                }
            }
        }
        if ok {
            let missing: Vec<String> = params
                .iter()
                .enumerate()
                .filter(|(j, p)| {
                    !p.optional
                        && p.default.is_none()
                        && !p.rest
                        && slots.get(*j).is_none_or(|s| s.is_none())
                })
                .map(|(_, p)| format!("`{}`", p.name))
                .collect();
            if !missing.is_empty() {
                self.err(
                    call.span,
                    "E0414",
                    format!(
                        "{shown} の必須の引数 {} が指定されていません",
                        missing.join(", ")
                    ),
                );
                ok = false;
            }
        }
        if !ok {
            for a in args {
                self.expr(named_value(a), None);
            }
            return None;
        }
        let hole = hole.expect("a named argument");
        let span = Span::empty(call.span.end());
        let v: Rc<[Expr]> = slots
            .into_iter()
            .map(|s| {
                s.unwrap_or(Expr {
                    id: hole,
                    kind: ExprKind::Undefined,
                    span,
                })
            })
            .collect();
        self.info().positional_args.insert(call.id, v.clone());
        Some(v)
    }

    fn new_expr_inner(&mut self, e: &Expr, ty: &ast::TypeRef, args: &[Expr]) -> Type {
        // Built-in constructors.
        if let ast::TypeRef::Named {
            path, args: targs, ..
        } = ty
        {
            if path.len() == 1 {
                match path[0].text.as_str() {
                    "Error" => {
                        if args.is_empty() || args.len() > 2 {
                            self.err(
                                e.span,
                                "E0402",
                                "`new Error(message, code?)` の形で使ってください",
                            );
                        }
                        for (i, a) in args.iter().enumerate() {
                            let want = if i == 0 { Type::Str } else { Type::Int };
                            let t = self.value(a, Some(&want));
                            self.coerce(a, &t, &want);
                        }
                        return self.error_type();
                    }
                    "Array" => {
                        let tps = self.fx.type_params.clone();
                        let elem = match targs.first() {
                            Some(t) if targs.len() == 1 => {
                                resolve_type(self.prog, self.module, &tps, t)
                            }
                            _ => {
                                self.err(
                                    e.span,
                                    "E0401",
                                    "`new Array<T>(長さ)` のように要素の型を指定してください",
                                );
                                Type::Error
                            }
                        };
                        if args.is_empty() || args.len() > 2 {
                            self.err(e.span, "E0402", "`new Array<T>(長さ)` または `new Array<T>(行数, 列数)` の形で使ってください");
                        }
                        for a in args {
                            let t = self.value(a, Some(&Type::Int));
                            if !(t.is_integer() || matches!(t, Type::Any | Type::Error)) {
                                self.err(a.span, "E0389", "配列の長さは整数である必要があります");
                            }
                        }
                        return if args.len() == 2 {
                            Type::array(Type::array(elem))
                        } else {
                            Type::array(elem)
                        };
                    }
                    "Map" => {
                        let tps = self.fx.type_params.clone();
                        if targs.len() != 2 {
                            self.err(
                                e.span,
                                "E0401",
                                "`new Map<K, V>()` のようにキーと値の型を指定してください",
                            );
                            return Type::Error;
                        }
                        let k = resolve_type(self.prog, self.module, &tps, &targs[0]);
                        let v = resolve_type(self.prog, self.module, &tps, &targs[1]);
                        if !args.is_empty() {
                            self.err(e.span, "E0402", "`new Map<K, V>()` は引数を取りません");
                        }
                        if !(k.is_numeric()
                            || matches!(
                                k,
                                Type::Str | Type::Any | Type::Date | Type::Error | Type::Param(_)
                            ))
                        {
                            self.err(
                                targs[0].span(),
                                "E0408",
                                "`Map` のキーには数値・文字列・日付のいずれかを使ってください",
                            );
                        }
                        // A map of structs is an instance of the bundled class (see `resolve_type`).
                        if let (Type::Struct(_), Some(c)) =
                            (&v, crate::collect::struct_map_class(self.prog))
                        {
                            let targs = vec![k, v];
                            self.info().calls.insert(
                                e.id,
                                CallInfo {
                                    target: CallTarget::Ctor(c),
                                    type_args: targs.clone(),
                                },
                            );
                            return Type::Class(c, targs);
                        }
                        return Type::Map(Box::new(k), Box::new(v));
                    }
                    _ => {}
                }
            }
        }
        let tps = self.fx.type_params.clone();
        let t = resolve_type(self.prog, self.module, &tps, ty);
        match &t {
            Type::Class(c, targs) => {
                let def = self.prog.classes[*c].clone();
                match def.kind {
                    ClassKind::Interface => {
                        self.err(
                            e.span,
                            "E0409",
                            format!("インターフェイス `{}` のインスタンスは作れません", def.name),
                        );
                    }
                    ClassKind::Host => {
                        if def.namespace.as_deref() != Some("VBA")
                            && crate::model::prog_id(def.namespace.as_deref(), &def.name).is_none()
                        {
                            self.err(
                                e.span,
                                "E0409",
                                format!(
                                    "`{}` は `new` で作成できません",
                                    self.prog.class_display(*c)
                                ),
                            );
                        }
                        for a in args {
                            self.value(a, None);
                        }
                        self.info().calls.insert(
                            e.id,
                            CallInfo {
                                target: CallTarget::Ctor(*c),
                                type_args: targs.clone(),
                            },
                        );
                        return t;
                    }
                    ClassKind::Class => {}
                }
                match def.ctor {
                    Some(ctor) => {
                        let env = self.prog.class_env(*c, targs);
                        self.check_fn_call(e, ctor, env, None, &[], args);
                        if !self.fx.calls.contains(&ctor) {
                            self.fx.calls.push(ctor);
                        }
                    }
                    None => {
                        if !args.is_empty() {
                            self.err(
                                e.span,
                                "E0402",
                                format!(
                                    "クラス `{}` にはコンストラクタがないため、引数は渡せません",
                                    def.name
                                ),
                            );
                        }
                        for a in args {
                            self.expr(a, None);
                        }
                    }
                }
                self.info().calls.insert(
                    e.id,
                    CallInfo {
                        target: CallTarget::Ctor(*c),
                        type_args: targs.clone(),
                    },
                );
                t
            }
            Type::Error => {
                for a in args {
                    self.expr(a, None);
                }
                Type::Error
            }
            other => {
                let shown = self.show(other);
                for a in args {
                    self.expr(a, None);
                }
                self.err(
                    e.span,
                    "E0409",
                    format!("`{shown}` は `new` で作成できません"),
                );
                Type::Error
            }
        }
    }

    // ------------------------------------------------------------------ intrinsics

    fn arg_count(&mut self, call: &Expr, args: &[Expr], min: usize, max: usize) -> bool {
        if args.len() < min || args.len() > max {
            let want = if min == max {
                format!("{min}")
            } else {
                format!("{min}〜{max}")
            };
            self.err(
                call.span,
                "E0402",
                format!(
                    "引数の数が合いません（{want} 個必要ですが {} 個渡されています）",
                    args.len()
                ),
            );
            for a in args {
                self.expr(a, None);
            }
            return false;
        }
        true
    }

    /// Checks `args[i]` (when present) against `want`.
    fn arg(&mut self, args: &[Expr], i: usize, want: &Type) -> Type {
        match args.get(i) {
            Some(a) => {
                let t = self.value(a, Some(want));
                self.coerce(a, &t, want);
                t
            }
            None => Type::Error,
        }
    }

    /// Checks a callback argument; returns its return type.
    fn lambda_arg(
        &mut self,
        args: &[Expr],
        i: usize,
        params: Vec<Type>,
        ret: Option<Type>,
    ) -> Type {
        let Some(a) = args.get(i) else {
            return Type::Error;
        };
        let unknown = Type::Param("?ret".into());
        let expect = Type::Func(Box::new(FuncType {
            params: params.clone(),
            ret: ret.clone().unwrap_or(unknown),
        }));
        let t = self.value(a, Some(&expect));
        match t {
            Type::Func(f) => {
                if f.params.len() > params.len() {
                    self.err(
                        a.span,
                        "E0400",
                        format!("コールバックの引数が多すぎます（最大 {} 個）", params.len()),
                    );
                }
                for (have, want) in f.params.iter().zip(params.iter()) {
                    if !self.assignable(want, have) {
                        self.err(
                            a.span,
                            "E0301",
                            format!(
                                "コールバックの引数の型 `{}` に `{}` を渡せません",
                                self.show(have),
                                self.show(want)
                            ),
                        );
                    }
                }
                if let Some(r) = &ret {
                    if *r != Type::Void && !self.assignable(&f.ret, r) {
                        self.err(
                            a.span,
                            "E0301",
                            format!(
                                "コールバックは `{}` を返す必要があります（`{}` を返しています）",
                                self.show(r),
                                self.show(&f.ret)
                            ),
                        );
                    }
                }
                f.ret
            }
            Type::Error => Type::Error,
            other => {
                let shown = self.show(&other);
                self.err(
                    a.span,
                    "E0301",
                    format!("ここには関数が必要です（`{shown}` です）"),
                );
                Type::Error
            }
        }
    }

    /// The RAS type of a VBA type an event declares a parameter with (`MSForms.ReturnBoolean`,
    /// `Range`, `Integer`); `any` for a class the type definitions do not have.
    fn event_param_type(&self, vba: &str) -> Type {
        match vba {
            "Integer" => Type::Short,
            "Long" => Type::Int,
            "Single" => Type::Float,
            "Double" => Type::Double,
            "String" => Type::Str,
            "Boolean" => Type::Bool,
            "Object" => Type::Object,
            "Variant" => Type::Any,
            other => {
                let (ns, name) = other.split_once('.').unwrap_or(("Excel", other));
                match self.prog.namespaces.get(ns).and_then(|n| n.get(name)) {
                    Some(Sym::Class(c)) => Type::Class(*c, Vec::new()),
                    Some(Sym::Enum(e)) => Type::Enum(*e),
                    _ => Type::Any,
                }
            }
        }
    }

    /// The function `e` names, when it is the name of a function of the project (`refresh`,
    /// `report.refresh`, `Sheet1.refresh`) and not a value.
    fn function_name(&mut self, e: &Expr) -> Option<FnId> {
        let p = e.peel();
        let access = match &p.kind {
            ExprKind::Name(id)
                if self.lookup_local(&id.text).is_none()
                    && matches!(self.lookup_sym(&id.text), Some(Sym::Fn(_))) =>
            {
                self.name_access(p, id)
            }
            ExprKind::Member {
                object,
                name,
                optional: false,
            } => {
                let o = object.peel();
                let ExprKind::Name(id) = &o.kind else {
                    return None;
                };
                if self.lookup_local(&id.text).is_some()
                    || !matches!(
                        self.lookup_sym(&id.text),
                        Some(Sym::Module(_) | Sym::Var(_))
                    )
                {
                    return None;
                }
                let ot = self.expr(object, None);
                let is_fn = match &ot {
                    Type::Ns(NsRef::Module(m)) => matches!(
                        self.prog.modules[*m].exports.get(&name.text),
                        Some(Sym::Fn(_))
                    ),
                    Type::Class(c, _) => {
                        self.prog.find_field(*c, &name.text).is_none()
                            && self
                                .prog
                                .find_methods(*c, &name.text)
                                .iter()
                                .any(|&f| matches!(self.prog.fns[f].kind, FnKind::Free))
                    }
                    _ => false,
                };
                if !is_fn {
                    return None;
                }
                self.member_access(p, object, &ot, name)
            }
            _ => return None,
        };
        match access {
            Access::Func { f, .. } if matches!(self.prog.fns[f].kind, FnKind::Free) => Some(f),
            _ => None,
        }
    }

    /// Checks the handler of an `on` call that is wrong already, without asking for the types
    /// of the parameters of an arrow function.
    fn skip_handler(&mut self, handler: &Expr) {
        if is_lambda(handler) {
            let any = Type::Func(Box::new(FuncType {
                params: vec![Type::Any; 16],
                ret: Type::Param("?ret".into()),
            }));
            self.fx.event_handlers += 1;
            self.expr(handler, Some(&any));
            self.fx.event_handlers -= 1;
        } else if self.function_name(handler).is_none() {
            self.expr(handler, None);
        }
    }

    /// `target.on("<Event>", handler)` / `target.addEventListener(…)` (spec §12.2.3): the
    /// handler is called for the events of the object from now on. An object with events is
    /// listened to by a generated class (VBA's `WithEvents`); one with an `OnAction` property
    /// gets the name of the macro.
    fn event_on(&mut self, call: &Expr, recv: Option<&Type>, args: &[Expr]) -> Type {
        let listener = match self.prog.globals.get("RAS_Listener") {
            Some(Sym::Class(c)) => Type::Class(*c, Vec::new()),
            _ => Type::Object,
        };
        if !self.arg_count(call, args, 2, 2) {
            return Type::Error;
        }
        let event = match &args[0].peel().kind {
            ExprKind::Str(s) => Some(s.clone()),
            _ => None,
        };
        self.expr(&args[0], Some(&Type::Str));
        let Some(event) = event else {
            self.err(args[0].span, "E0616", "`on` のイベントの名前は、文字列リテラルで書きます（`on(\"Click\", …)`。イベントごとに受け取るクラスを生成するため、実行時に決まる名前は使えません）");
            self.skip_handler(&args[1]);
            return Type::Error;
        };
        // An array or a collection of objects: each member gets the handler.
        let any_control = |prog: &Program, c: ClassId| {
            prog.classes[c].namespace.as_deref() == Some("MSForms")
                && prog.classes[c].name == "Control"
        };
        let (class, each) = match recv {
            Some(t @ Type::Array(_)) => match t.elem() {
                Some(Type::Class(c, _)) => (*c, true),
                _ => {
                    self.skip_handler(&args[1]);
                    return Type::Error;
                }
            },
            Some(t @ Type::Class(c, _)) => match self.prog.collection_elem(t) {
                Some(Type::Class(m, _))
                    if crate::lower::listen_target(self.prog, *c).is_none()
                        && (any_control(self.prog, m)
                            || crate::lower::listen_target(self.prog, m).is_some()) =>
                {
                    (m, true)
                }
                _ => (*c, false),
            },
            _ => {
                self.skip_handler(&args[1]);
                return Type::Error;
            }
        };
        let shown = self.prog.class_display(class);
        // A `MSForms.Control` is any control: the events that controls have, with the
        // arguments that all of those controls pass.
        let dynamic = any_control(self.prog, class);
        let mut control_classes: Vec<(ClassId, Vec<(String, Vec<(String, String, bool)>)>)> =
            Vec::new();
        let target = if dynamic {
            let mut ids: Vec<ClassId> = self
                .prog
                .namespaces
                .get("MSForms")
                .map(|n| {
                    n.values()
                        .filter_map(|s| {
                            if let Sym::Class(c) = s {
                                Some(*c)
                            } else {
                                None
                            }
                        })
                        .collect()
                })
                .unwrap_or_default();
            ids.sort_by(|a, b| self.prog.classes[*a].name.cmp(&self.prog.classes[*b].name));
            let mut events: Vec<(String, Vec<(String, String, bool)>)> = Vec::new();
            for c in ids {
                let Some(crate::lower::ListenTarget::Events(_, own)) =
                    crate::lower::listen_target(self.prog, c)
                else {
                    continue;
                };
                for (name, params) in &own {
                    match events.iter_mut().find(|e| e.0.eq_ignore_ascii_case(name)) {
                        Some(e) => {
                            let common =
                                e.1.iter()
                                    .zip(params.iter())
                                    .take_while(|(a, b)| a.1 == b.1 && a.2 == b.2)
                                    .count();
                            e.1.truncate(common);
                        }
                        None => events.push((name.clone(), params.clone())),
                    }
                }
                control_classes.push((c, own));
            }
            events.sort_by(|a, b| a.0.cmp(&b.0));
            Some(crate::lower::ListenTarget::Events(class, events))
        } else {
            crate::lower::listen_target(self.prog, class)
        };
        match target {
            None => {
                let def = &self.prog.classes[class];
                let hint = if self.prog.modules[def.module].is_form() {
                    "フォーム自身のイベントは、コードビハインドの `@on(\"UserForm.<イベント>\")` で受け取ります".to_string()
                } else {
                    "`on` を使えるのは、イベントを持つオブジェクト（`MSForms.CommandButton` などのコントロール、`Excel.Application`・`Excel.Workbook`・`Excel.Worksheet`・`Excel.Chart`・`Excel.QueryTable`）と、`onAction` を持つオブジェクト（`Excel.Shape`）です".to_string()
                };
                self.err(
                    call.span,
                    "E0616",
                    format!("`{shown}` にイベントハンドラは付けられません。{hint}"),
                );
                self.skip_handler(&args[1]);
                Type::Error
            }
            Some(crate::lower::ListenTarget::OnAction) => {
                if !event.eq_ignore_ascii_case("Click") {
                    self.err(args[0].span, "E0611", format!("`{shown}` で受け取れるイベントは `Click` だけです（`OnAction` に登録したマクロが、クリックされたときに呼ばれます）"));
                }
                match self.function_name(&args[1]) {
                    Some(f) => {
                        let d = &self.prog.fns[f];
                        let home = &self.prog.modules[d.module];
                        if !d.is_macro_entry()
                            || home.is_form()
                            || home.kind != ModuleKind::User
                            || home.is_test()
                        {
                            let name = d.name.clone();
                            self.err(args[1].span, "E0617", format!("関数 `{name}` は `OnAction` に登録できません。登録できるのは、Excel がマクロとして呼べる関数（`src/` か `objects/` のファイルの、引数も戻り値もない `export function`、または `@macro` を付けた関数）です"));
                        }
                        let sig = Type::Func(Box::new(FuncType {
                            params: Vec::new(),
                            ret: Type::Void,
                        }));
                        if let Some(slot) =
                            self.info().expr_types.get_mut(args[1].peel().id as usize)
                        {
                            *slot = sig;
                        }
                        self.info().event_binds.insert(
                            call.id,
                            EventBind {
                                kind: EventBindKind::OnAction { each },
                                handler_fn: Some(f),
                            },
                        );
                    }
                    None => {
                        self.skip_handler(&args[1]);
                        self.err(args[1].span, "E0617", format!("`{shown}` の `on` には、マクロにできる関数の名前を渡します（`shape.on(\"Click\", refresh)`）。Excel はマクロの名前をブックに保存するので、アロー関数や変数に入れた関数は登録できません"));
                    }
                }
                Type::Void
            }
            Some(crate::lower::ListenTarget::Events(event_class, events)) => {
                let Some((name, params)) = events
                    .iter()
                    .find(|e| e.0.eq_ignore_ascii_case(&event))
                    .cloned()
                else {
                    let form_only = ["Enter", "Exit", "BeforeUpdate", "AfterUpdate"]
                        .iter()
                        .any(|x| x.eq_ignore_ascii_case(&event))
                        && self.prog.classes[event_class].namespace.as_deref() == Some("MSForms");
                    let hint = if form_only {
                        format!("`{event}` はフォームが起こすイベントで、`on` では受け取れません。コードビハインドの `@on(\"<コントロール名>.{event}\")` で受け取ります")
                    } else {
                        format!(
                            "使えるイベント: {}",
                            events
                                .iter()
                                .map(|e| e.0.as_str())
                                .collect::<Vec<_>>()
                                .join(", ")
                        )
                    };
                    self.err(
                        args[0].span,
                        "E0611",
                        format!("`{shown}` にイベント `{event}` はありません（{hint}）"),
                    );
                    self.skip_handler(&args[1]);
                    return if each { Type::Void } else { listener };
                };
                // Every member is listened to: there is not one listener to give back.
                let listener = if each { Type::Void } else { listener };
                let classes: Vec<ClassId> = if dynamic {
                    control_classes
                        .iter()
                        .filter(|(_, own)| own.iter().any(|e| e.0.eq_ignore_ascii_case(&name)))
                        .map(|(c, _)| *c)
                        .collect()
                } else {
                    vec![event_class]
                };
                let mut param_types: Vec<Type> =
                    params.iter().map(|p| self.event_param_type(&p.1)).collect();
                // The handler of the members of an array or a collection is told which member
                // the event is of: it comes first, before the arguments of the event.
                if each {
                    param_types.insert(0, Type::Class(class, Vec::new()));
                }
                let limit = param_types.len();
                let (argc, handler_fn) = match self.function_name(&args[1]) {
                    Some(f) => {
                        self.ensure_fn(f);
                        let d = &self.prog.fns[f];
                        let fname = d.name.clone();
                        let have: Vec<(Type, bool)> = d
                            .params
                            .iter()
                            .map(|p| (p.ty.clone(), p.mode == PassMode::ByRef))
                            .collect();
                        let generic = d.is_generic();
                        let test = self.prog.modules[d.module].is_test()
                            && !self.prog.modules[self.module].is_test();
                        if generic || test {
                            self.err(args[1].span, "E0617", format!("関数 `{fname}` はイベントハンドラにできません（ジェネリック関数・テストの関数は使えません）"));
                        } else if have.len() > limit {
                            self.err(args[1].span, "E0612", format!("イベント `{name}` のハンドラに渡される引数は {limit} 個ですが、ハンドラ `{fname}` は {} 個受け取ります", have.len()));
                        } else {
                            for (i, (ty, byref)) in have.iter().enumerate() {
                                if *byref {
                                    self.err(args[1].span, "E0617", format!("`on` に渡す関数 `{fname}` の引数に `byref` は使えません（イベントの引数は値で渡されます。取り消すには `false` を返します）"));
                                } else if !self.assignable(&param_types[i], ty) {
                                    self.err(args[1].span, "E0301", format!("イベント `{name}` の {} 番目の引数 `{}` を、ハンドラ `{fname}` の引数 `{}` に渡せません", i + 1, self.show(&param_types[i]), self.show(ty)));
                                }
                            }
                        }
                        let sig = Type::Func(Box::new(FuncType {
                            params: have.iter().map(|h| h.0.clone()).collect(),
                            ret: self.prog.fns[f].ret_ty(),
                        }));
                        if let Some(slot) =
                            self.info().expr_types.get_mut(args[1].peel().id as usize)
                        {
                            *slot = sig;
                        }
                        if !self.fx.calls.contains(&f) {
                            self.fx.calls.push(f);
                        }
                        (have.len().min(limit), Some(f))
                    }
                    None => {
                        // (Said once, in the words of events.)
                        if let ExprKind::Arrow(arrow) = &args[1].peel().kind {
                            if arrow.params.len() > limit {
                                self.err(args[1].span, "E0612", format!("イベント `{name}` のハンドラに渡される引数は {limit} 個ですが、ハンドラは {} 個受け取ります", arrow.params.len()));
                                self.skip_handler(&args[1]);
                                return listener;
                            }
                        }
                        self.fx.event_handlers += 1;
                        self.lambda_arg(args, 1, param_types, None);
                        self.fx.event_handlers -= 1;
                        let argc = match self.info().ty(args[1].id) {
                            Type::Func(f) => f.params.len(),
                            _ => 0,
                        };
                        (argc.min(limit), None)
                    }
                };
                self.info().event_binds.insert(
                    call.id,
                    EventBind {
                        kind: EventBindKind::Listener {
                            classes,
                            dynamic,
                            each,
                            event: name,
                            params,
                            argc,
                        },
                        handler_fn,
                    },
                );
                listener
            }
        }
    }

    fn intrinsic_call(
        &mut self,
        call: &Expr,
        intr: Intr,
        recv_expr: Option<&Expr>,
        recv: Option<&Type>,
        args: &[Expr],
        expect: Option<&Type>,
    ) -> Type {
        use Intr::*;
        let _ = expect;
        let elem = recv.and_then(|t| t.elem().cloned()).unwrap_or(Type::Error);
        let arr = || Type::array(elem.clone());
        if recv.is_some_and(|t| t.is_array()) && elem.has_struct() {
            // These go through run-time helpers that take the array as a Variant.
            if matches!(
                intr,
                ArrIndexOf | ArrIncludes | ArrJoin | ArrSort | ArrReverse
            ) {
                self.err(call.span, "E0393", "構造体の配列では `indexOf` / `includes` / `join` / `sort` / `reverse` を使えません。ループか `findIndex` などのコールバック付きメソッドを使ってください");
                for a in args {
                    self.expr(a, None);
                }
                return Type::Error;
            }
            if intr.takes_lambda() && args.first().is_some_and(|a| !is_lambda(a)) {
                self.err(args[0].span, "E0393", "構造体の配列のコールバックには、アロー関数を直接書いてください（変数に入れた関数は使えません）");
            }
        }
        // Mutating array methods need an assignable receiver.
        if matches!(intr, ArrPush | ArrPop | ArrSort | ArrReverse | ArrClear) {
            if let Some(r) = recv_expr {
                self.mutable_array_receiver(r);
            }
        }
        match intr {
            ArrLength | ArrRows | ArrCols | StrLength | MapSize => Type::Int,
            CollToArray => {
                self.arg_count(call, args, 0, 0);
                arr()
            }
            ArrPush => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &elem);
                }
                Type::Void
            }
            ArrPop => {
                self.arg_count(call, args, 0, 0);
                elem
            }
            ArrMap => {
                if !self.arg_count(call, args, 1, 1) {
                    return Type::Error;
                }
                let r = self.lambda_arg(args, 0, vec![elem, Type::Int], None);
                if r == Type::Void {
                    self.err(
                        args[0].span,
                        "E0370",
                        "`map` のコールバックは値を返す必要があります",
                    );
                    return Type::Error;
                }
                Type::array(r)
            }
            ArrFilter => {
                if self.arg_count(call, args, 1, 1) {
                    self.lambda_arg(args, 0, vec![elem.clone(), Type::Int], Some(Type::Bool));
                }
                arr()
            }
            ArrForEach => {
                if self.arg_count(call, args, 1, 1) {
                    self.lambda_arg(args, 0, vec![elem, Type::Int], Some(Type::Void));
                }
                Type::Void
            }
            ArrReduce => {
                if !self.arg_count(call, args, 2, 2) {
                    return Type::Error;
                }
                let acc = self.value(&args[1], None);
                let acc = if acc == Type::Null { Type::Any } else { acc };
                self.lambda_arg(
                    args,
                    0,
                    vec![acc.clone(), elem, Type::Int],
                    Some(acc.clone()),
                );
                acc
            }
            ArrSome | ArrEvery => {
                if self.arg_count(call, args, 1, 1) {
                    self.lambda_arg(args, 0, vec![elem, Type::Int], Some(Type::Bool));
                }
                Type::Bool
            }
            ArrFind => {
                if self.arg_count(call, args, 1, 1) {
                    self.lambda_arg(args, 0, vec![elem.clone(), Type::Int], Some(Type::Bool));
                }
                elem
            }
            ArrFindIndex => {
                if self.arg_count(call, args, 1, 1) {
                    self.lambda_arg(args, 0, vec![elem, Type::Int], Some(Type::Bool));
                }
                Type::Int
            }
            ArrIndexOf => {
                if self.arg_count(call, args, 1, 2) {
                    self.arg(args, 0, &elem);
                    self.arg(args, 1, &Type::Int);
                }
                Type::Int
            }
            ArrIncludes => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &elem);
                }
                Type::Bool
            }
            ArrJoin => {
                if self.arg_count(call, args, 0, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                if !self.stringifiable(&elem) {
                    self.err(
                        call.span,
                        "E0381",
                        "`join` は数値・文字列などの配列にのみ使えます",
                    );
                }
                Type::Str
            }
            ArrSlice => {
                if self.arg_count(call, args, 0, 2) {
                    self.arg(args, 0, &Type::Int);
                    self.arg(args, 1, &Type::Int);
                }
                arr()
            }
            ArrSort => {
                if self.arg_count(call, args, 0, 1) && !args.is_empty() {
                    self.lambda_arg(
                        args,
                        0,
                        vec![elem.clone(), elem.clone()],
                        Some(Type::Double),
                    );
                } else if !self.satisfies(&elem, &Constraint::Comparable) {
                    self.err(call.span, "E0404", "この配列の要素は大小比較できません。`sort((a, b) => ...)` のように比較関数を渡してください");
                }
                Type::Void
            }
            ArrReverse | ArrClear => {
                self.arg_count(call, args, 0, 0);
                Type::Void
            }
            ArrConcat => {
                if self.arg_count(call, args, 1, 1) {
                    let t = self.value(&args[0], Some(&arr()));
                    if t != arr() && !t.is_error() {
                        self.err(
                            args[0].span,
                            "E0301",
                            "`concat` には同じ要素型の配列を渡してください",
                        );
                    }
                }
                arr()
            }
            StrTrim | StrTrimStart | StrTrimEnd | StrUpper | StrLower | StrToString => {
                self.arg_count(call, args, 0, 0);
                Type::Str
            }
            StrSubstring | StrSlice | StrSubstr => {
                if self.arg_count(call, args, 1, 2) {
                    self.arg(args, 0, &Type::Int);
                    self.arg(args, 1, &Type::Int);
                }
                Type::Str
            }
            StrIndexOf | StrLastIndexOf => {
                if self.arg_count(call, args, 1, 2) {
                    self.arg(args, 0, &Type::Str);
                    self.arg(args, 1, &Type::Int);
                }
                Type::Int
            }
            StrIncludes | StrStartsWith | StrEndsWith => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                Type::Bool
            }
            StrPadStart | StrPadEnd => {
                if self.arg_count(call, args, 1, 2) {
                    self.arg(args, 0, &Type::Int);
                    self.arg(args, 1, &Type::Str);
                }
                Type::Str
            }
            StrRepeat => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Int);
                }
                Type::Str
            }
            StrReplace | StrReplaceAll => {
                if self.arg_count(call, args, 2, 2) {
                    self.arg(args, 0, &Type::Str);
                    self.arg(args, 1, &Type::Str);
                }
                Type::Str
            }
            StrSplit => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                Type::array(Type::Str)
            }
            StrCharAt => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Int);
                }
                Type::Str
            }
            StrCharCodeAt => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Int);
                }
                Type::Int
            }
            NumToString | BoolToString | DateToString => {
                self.arg_count(call, args, 0, 0);
                Type::Str
            }
            NumToFixed => {
                if self.arg_count(call, args, 0, 1) {
                    self.arg(args, 0, &Type::Int);
                }
                Type::Str
            }
            MapGet | MapHas | MapDelete => {
                let (k, v) = match recv {
                    Some(Type::Map(k, v)) => ((**k).clone(), (**v).clone()),
                    _ => (Type::Error, Type::Error),
                };
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &k);
                }
                match intr {
                    MapGet => v,
                    MapHas => Type::Bool,
                    _ => Type::Void,
                }
            }
            MapSet => {
                let (k, v) = match recv {
                    Some(Type::Map(k, v)) => ((**k).clone(), (**v).clone()),
                    _ => (Type::Error, Type::Error),
                };
                if self.arg_count(call, args, 2, 2) {
                    self.arg(args, 0, &k);
                    self.arg(args, 1, &v);
                }
                Type::Void
            }
            MapKeys | MapValues => {
                self.arg_count(call, args, 0, 0);
                match recv {
                    Some(Type::Map(k, v)) => Type::array(if intr == MapKeys {
                        (**k).clone()
                    } else {
                        (**v).clone()
                    }),
                    _ => Type::Error,
                }
            }
            MapClear => {
                self.arg_count(call, args, 0, 0);
                Type::Void
            }
            MathAbs => {
                if !self.arg_count(call, args, 1, 1) {
                    return Type::Error;
                }
                let t = self.value(&args[0], None);
                if !self.is_numeric(&t) {
                    self.err(args[0].span, "E0387", "数値が必要です");
                    return Type::Error;
                }
                match t {
                    Type::Byte | Type::Short | Type::Enum(_) => Type::Int,
                    other => other,
                }
            }
            MathFloor | MathCeil | MathTrunc | MathSqrt | MathSin | MathCos | MathTan
            | MathAtan | MathLog | MathExp => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Double);
                }
                Type::Double
            }
            MathRound | MathBankersRound => {
                if self.arg_count(call, args, 1, 2) {
                    self.arg(args, 0, &Type::Double);
                    self.arg(args, 1, &Type::Int);
                }
                Type::Double
            }
            MathPow => {
                if self.arg_count(call, args, 2, 2) {
                    self.arg(args, 0, &Type::Double);
                    self.arg(args, 1, &Type::Double);
                }
                Type::Double
            }
            MathMin | MathMax => {
                if args.len() < 2 {
                    self.err(call.span, "E0402", "2 つ以上の引数が必要です");
                    for a in args {
                        self.expr(a, None);
                    }
                    return Type::Error;
                }
                let mut ty: Option<Type> = None;
                for a in args {
                    let t = self.value(a, None);
                    if !self.is_numeric(&t) && !matches!(t, Type::Date) {
                        self.err(a.span, "E0387", "数値が必要です");
                        continue;
                    }
                    ty = Some(match ty {
                        None => t,
                        Some(p) => match self.promote(&p, &t) {
                            Some(x) => x,
                            None => {
                                self.err(
                                    a.span,
                                    "E0395",
                                    "型の混在した比較には明示的な型変換（`as`）が必要です",
                                );
                                Type::Error
                            }
                        },
                    });
                }
                ty.unwrap_or(Type::Error)
            }
            MathRandom => {
                self.arg_count(call, args, 0, 0);
                Type::Double
            }
            MathSign => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Double);
                }
                Type::Int
            }
            MathPi => Type::Double,
            ParseInt => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                Type::Int
            }
            ParseFloat => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                Type::Double
            }
            ToStr => {
                if self.arg_count(call, args, 1, 1) {
                    let t = self.value(&args[0], None);
                    if !self.stringifiable(&t) {
                        self.err(
                            args[0].span,
                            "E0381",
                            format!("`{}` は文字列に変換できません", self.show(&t)),
                        );
                    }
                }
                Type::Str
            }
            ConsoleLog => {
                for a in args {
                    let t = self.value(a, None);
                    if t.has_struct() {
                        self.err(
                            a.span,
                            "E0301",
                            format!(
                                "`{}` はそのまま出力できません。フィールドを個別に渡してください",
                                self.show(&t)
                            ),
                        );
                    }
                }
                Type::Void
            }
            IsNullish => Type::Bool,
            FastMode => {
                if self.arg_count(call, args, 0, 1) {
                    if let Some(a) = args.first() {
                        match &a.kind {
                            ExprKind::Object(props) => {
                                for p in props {
                                    let known = matches!(p.key.text.as_str(), "screen" | "events" | "calc" | "alerts" | "statusBar" | "cursor" | "interactive");
                                    if !known {
                                        self.err(p.key.span, "E0410", format!("`fastMode` に `{}` というオプションはありません（screen / events / calc / alerts / statusBar / cursor / interactive）", p.key.text));
                                    }
                                    if !matches!(p.value.kind, ExprKind::Bool(_)) {
                                        self.err(p.value.span, "E0410", "`fastMode` のオプションには `true` か `false` を直接書いてください");
                                    }
                                }
                            }
                            _ => self.err(a.span, "E0410", "`fastMode({ calc: false })` のようにオブジェクトリテラルで指定してください"),
                        }
                    }
                }
                Type::Int
            }
            SuspendScreen | ResumeScreen | SuspendEvents | ResumeEvents | SuspendCalc
            | ResumeCalc | SuspendAlerts | ResumeAlerts | SuspendAll | ResumeAll | YieldNow => {
                self.arg_count(call, args, 0, 0);
                Type::Void
            }
            Status => {
                if self.arg_count(call, args, 0, 1) {
                    self.arg(args, 0, &Type::Str);
                }
                Type::Void
            }
            Sleep => {
                if self.arg_count(call, args, 1, 1) {
                    self.arg(args, 0, &Type::Int);
                }
                Type::Void
            }
            WaitUntil => {
                if self.arg_count(call, args, 1, 2) {
                    self.lambda_arg(args, 0, vec![], Some(Type::Bool));
                    self.arg(args, 1, &Type::Int);
                }
                Type::Bool
            }
            PlatformIs64 | PlatformIsVba7 => Type::Bool,
            EventOn => self.event_on(call, recv, args),
            RangeToArray | SheetUsedTable => {
                self.arg_count(call, args, 0, 0);
                Type::array(Type::array(Type::Any))
            }
            RangeToColumn | RangeToRow => {
                self.arg_count(call, args, 0, 0);
                Type::array(Type::Any)
            }
            RangeFromArray => {
                if self.arg_count(call, args, 1, 1) {
                    let t = self.value(&args[0], None);
                    if !matches!(t, Type::Array(_) | Type::Any | Type::Error) {
                        self.err(
                            args[0].span,
                            "E0301",
                            "`fromArray` には配列を渡してください",
                        );
                    }
                }
                Type::Void
            }
            ExcelMatch => {
                if self.arg_count(call, args, 2, 3) {
                    self.value(&args[0], None);
                    self.value(&args[1], None);
                    self.arg(args, 2, &Type::Int);
                }
                Type::Any
            }
        }
    }

    fn mutable_array_receiver(&mut self, recv: &Expr) {
        let inner = recv.peel();
        if let ExprKind::Name(id) = &inner.kind {
            if let Some((l, crossed)) = self.lookup_local(&id.text) {
                if !crossed.is_empty() {
                    self.err(
                        recv.span,
                        "E0322",
                        format!(
                            "アロー関数の中から外側の配列 `{}` は書き換えられません",
                            id.text
                        ),
                    );
                }
                let _ = l;
            }
            if let Some(Res::Global(v)) = self.info().res.get(&inner.id).cloned() {
                self.prog.vars[v].reassigned = true;
            }
        } else if !matches!(
            inner.kind,
            ExprKind::Member { .. } | ExprKind::This | ExprKind::Index { .. } | ExprKind::Error
        ) {
            self.err(
                recv.span,
                "E0325",
                "配列を書き換えるメソッドは、変数かフィールドに対して呼んでください",
            );
        }
    }

    // ------------------------------------------------------------------ lints (spec §6)

    fn lints(&mut self) {
        self.lint_inline_recursion();
        self.lint_identifier_case();
        self.lint_docs();
    }

    /// Rule 8: `inline` functions must not be (mutually) recursive.
    fn lint_inline_recursion(&mut self) {
        for f in 0..self.prog.fns.len() {
            if self.prog.fns[f].inline != ast::InlineMode::Inline {
                continue;
            }
            let mut seen = HashSet::new();
            let mut stack: Vec<FnId> = self.prog.fns[f].calls.clone();
            let mut recursive = false;
            while let Some(g) = stack.pop() {
                if g == f {
                    recursive = true;
                    break;
                }
                if seen.insert(g) {
                    stack.extend(self.prog.fns[g].calls.iter().copied());
                }
            }
            if recursive {
                let d = &self.prog.fns[f];
                let (m, span, name) = (d.module, d.name_span, d.name.clone());
                self.module = m;
                self.err(span, "E0420", format!("`inline` 関数 `{name}` は（直接または間接に）再帰しているため展開できません"));
            }
        }
    }

    /// Rule 6: identifiers that differ only by case collide in VBA.
    fn lint_identifier_case(&mut self) {
        let mut seen: HashMap<String, String> = HashMap::new();
        let mut items: Vec<(ModuleId, Span, String)> = Vec::new();
        for m in &self.prog.modules {
            if m.kind != ModuleKind::User {
                continue;
            }
            for &f in &m.fns {
                let d = &self.prog.fns[f];
                if !matches!(d.kind, FnKind::Main | FnKind::ModInit)
                    && !d.name.is_empty()
                    && d.name != "constructor"
                {
                    items.push((m.id, d.name_span, d.name.clone()));
                }
                for l in &d.locals {
                    if !l.name.is_empty() && l.kind != LocalKind::This {
                        items.push((m.id, l.span, l.name.clone()));
                    }
                }
            }
            for &c in &m.classes {
                let d = &self.prog.classes[c];
                // The form of a code-behind: its controls are named in the designer.
                if d.kind == ClassKind::Host {
                    continue;
                }
                items.push((m.id, d.name_span, d.name.clone()));
                for fld in &d.fields {
                    items.push((m.id, fld.span, fld.name.clone()));
                }
            }
            for &v in &m.vars {
                let d = &self.prog.vars[v];
                items.push((m.id, d.name_span, d.name.clone()));
            }
            for &e in &m.enums {
                let d = &self.prog.enums[e];
                items.push((m.id, d.name_span, d.name.clone()));
            }
        }
        let mut reported: HashSet<String> = HashSet::new();
        for (m, span, name) in items {
            let lower = name.to_lowercase();
            match seen.get(&lower) {
                Some(first) if *first != name => {
                    if reported.insert(format!("{lower}/{name}")) {
                        self.module = m;
                        let first = first.clone();
                        self.err(span, "E0430", format!("`{name}` と `{first}` は大文字小文字だけが異なります。VBA は大文字小文字を区別しないため共存できません"));
                    }
                }
                Some(_) => {}
                None => {
                    seen.insert(lower, name);
                }
            }
        }
    }

    /// `[lint] require-docs` and `@param` name checks (spec §16.1.2).
    fn lint_docs(&mut self) {
        let require = self.prog.config.lint.require_docs;
        for f in 0..self.prog.fns.len() {
            let d = &self.prog.fns[f];
            if self.prog.modules[d.module].kind != ModuleKind::User
                || matches!(d.kind, FnKind::Main | FnKind::ModInit)
            {
                continue;
            }
            let (m, span, name, exported) = (d.module, d.name_span, d.name.clone(), d.exported);
            match d.doc.clone() {
                Some(doc) => {
                    let params: Vec<String> = d.params.iter().map(|p| p.name.clone()).collect();
                    for line in doc.lines() {
                        let Some(rest) = line.trim().strip_prefix("@param") else {
                            continue;
                        };
                        let mut rest = rest.trim_start();
                        // JSDoc type annotations (`{string}`) are ignored.
                        if rest.starts_with('{') {
                            rest = rest.split_once('}').map(|x| x.1).unwrap_or("").trim_start();
                        }
                        let pname = rest.split_whitespace().next().unwrap_or("");
                        if !pname.is_empty() && !params.iter().any(|p| p == pname) {
                            self.module = m;
                            self.warn(span, "W0440", format!("`@param {pname}`: `{name}` に `{pname}` という引数はありません"));
                        }
                    }
                }
                None if require && exported => {
                    self.module = m;
                    self.warn(
                        span,
                        "W0441",
                        format!("`export` された `{name}` にドキュメントコメントがありません"),
                    );
                }
                None => {}
            }
        }
    }
}

/// Compile-time constant value of an initializer (for VBA `Const`).
pub fn const_value(e: &Expr) -> Option<ConstVal> {
    match &e.peel().kind {
        ExprKind::Int { .. } | ExprKind::Unary { .. } | ExprKind::Binary { .. }
            if const_int(e).is_some() =>
        {
            const_int(e).map(ConstVal::Int)
        }
        ExprKind::Float { text, suffix } if *suffix != ast::NumSuffix::M => {
            text.parse().ok().map(ConstVal::Float)
        }
        ExprKind::Unary {
            op: UnaryOp::Neg,
            operand,
        } => match const_value(operand) {
            Some(ConstVal::Float(f)) => Some(ConstVal::Float(-f)),
            _ => None,
        },
        ExprKind::Str(s) => Some(ConstVal::Str(s.clone())),
        ExprKind::Bool(b) => Some(ConstVal::Bool(*b)),
        _ => None,
    }
}

/// Conservative "all paths return or throw" analysis.
fn always_returns(stmts: &[Stmt]) -> bool {
    stmts.iter().any(stmt_returns)
}

fn stmt_returns(s: &Stmt) -> bool {
    match &s.kind {
        StmtKind::Return(_) | StmtKind::Throw(_) => true,
        StmtKind::Block(b) => always_returns(&b.stmts),
        StmtKind::If {
            then,
            els: Some(els),
            ..
        } => stmt_returns(then) && stmt_returns(els),
        StmtKind::Try(t) => {
            t.finally.as_ref().is_some_and(|f| always_returns(&f.stmts))
                || (always_returns(&t.block.stmts)
                    && t.catch
                        .as_ref()
                        .is_none_or(|c| always_returns(&c.body.stmts)))
        }
        StmtKind::Switch { cases, .. } => {
            cases.iter().any(|c| c.test.is_none())
                && cases
                    .iter()
                    .all(|c| always_returns(&c.body) || c.body.is_empty())
                && cases.last().is_some_and(|c| always_returns(&c.body))
        }
        StmtKind::While { cond, body, .. } => {
            matches!(cond.peel().kind, ExprKind::Bool(true)) && !contains_break(body)
        }
        StmtKind::For {
            cond: None, body, ..
        } => !contains_break(body),
        StmtKind::DoWhile { body, .. } => stmt_returns(body),
        StmtKind::Labeled { body, .. } => stmt_returns(body),
        _ => false,
    }
}

/// Control never reaches the statement after `s`.
fn stmt_exits(s: &Stmt) -> bool {
    match &s.kind {
        StmtKind::Return(_) | StmtKind::Throw(_) | StmtKind::Break(_) | StmtKind::Continue(_) => {
            true
        }
        StmtKind::Block(b) => b.stmts.iter().any(stmt_exits),
        StmtKind::If {
            then,
            els: Some(els),
            ..
        } => stmt_exits(then) && stmt_exits(els),
        _ => stmt_returns(s),
    }
}

/// `target` (of an assignment) writes the variable `name`.
fn assigns_target(target: &Expr, name: &str) -> bool {
    match &target.peel().kind {
        ExprKind::Name(id) => id.text == name,
        ExprKind::Array(items) => items.iter().any(|i| assigns_target(i, name)),
        _ => false,
    }
}

/// `e` contains an assignment to the variable `name`.
fn assigns_in_expr(e: &Expr, name: &str) -> bool {
    crate::astutil::any_expr_in(e, &mut |x| match &x.kind {
        ExprKind::Assign { target, .. } => assigns_target(target, name),
        ExprKind::IncDec { operand, .. } => assigns_target(operand, name),
        _ => false,
    })
}

fn assigns_in_stmt(s: &Stmt, name: &str) -> bool {
    crate::astutil::any_expr(s, &mut |x| match &x.kind {
        ExprKind::Assign { target, .. } => assigns_target(target, name),
        ExprKind::IncDec { operand, .. } => assigns_target(operand, name),
        _ => false,
    })
}

/// The type whose values `TypeName` reports as `name` (`"String"` → `string`, `"Variant()"` → `any[]`).
fn type_of_type_name(name: &str) -> Option<Type> {
    if let Some(elem) = name.strip_suffix("()") {
        return Some(Type::array(if elem == "Variant" {
            Type::Any
        } else {
            type_of_type_name(elem)?
        }));
    }
    Some(match name {
        "String" => Type::Str,
        "Boolean" => Type::Bool,
        "Byte" => Type::Byte,
        "Integer" => Type::Short,
        "Long" => Type::Int,
        "LongLong" => Type::Long,
        "Single" => Type::Float,
        "Double" => Type::Double,
        "Currency" => Type::Currency,
        "Decimal" => Type::Decimal,
        "Date" => Type::Date,
        _ => return None,
    })
}

fn contains_break(s: &Stmt) -> bool {
    match &s.kind {
        StmtKind::Break(_) => true,
        StmtKind::Block(b) => b.stmts.iter().any(contains_break),
        StmtKind::If { then, els, .. } => {
            contains_break(then) || els.as_ref().is_some_and(|e| contains_break(e))
        }
        StmtKind::Try(t) => {
            t.block.stmts.iter().any(contains_break)
                || t.catch
                    .as_ref()
                    .is_some_and(|c| c.body.stmts.iter().any(contains_break))
                || t.finally
                    .as_ref()
                    .is_some_and(|f| f.stmts.iter().any(contains_break))
        }
        StmtKind::Labeled { body, .. } => contains_break(body),
        // `break` inside a nested loop / switch belongs to that construct.
        _ => false,
    }
}

#[allow(dead_code)]
fn _unused(_: Rc<()>) {}
