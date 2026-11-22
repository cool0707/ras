//! Expression lowering.

use std::collections::HashMap;

use ras_syntax::ast::{self, BinaryOp, Expr, ExprKind, InlineMode, NumSuffix, PassMode, UnaryOp};

use super::stmt::{flatten_if, needs_handler};
use super::{Binding, Frame, Lowerer, RetMode};
use crate::diag::Severity;
use crate::intrinsics::Intr;
use crate::model::*;
use crate::names;
use crate::vba::*;

/// Formats a `Double` literal (always with a type suffix or exponent, spec §2.2).
pub(crate) fn double_lit(f: f64) -> String {
    if !f.is_finite() {
        return "0#".into();
    }
    if f == f.trunc() && f.abs() < 1e15 {
        return format!("{f:.0}#");
    }
    let s = format!("{f}");
    let a = f.abs();
    if s.contains('e') || (a != 0.0 && !(1e-5..1e15).contains(&a)) {
        // An exponent form is a Double by itself.
        let e = format!("{f:E}");
        match e.split_once('E') {
            Some((m, x)) => {
                let m = if m.contains('.') {
                    m.to_string()
                } else {
                    format!("{m}.0")
                };
                let x = if x.starts_with('-') {
                    x.to_string()
                } else {
                    format!("+{x}")
                };
                format!("{m}E{x}")
            }
            None => format!("{s}#"),
        }
    } else {
        format!("{s}#")
    }
}

fn int_lit(value: i128, ty: &Type, hex: bool) -> VExpr {
    let suffix = if *ty == Type::Long { "^" } else { "&" };
    if value == i32::MIN as i128 && *ty != Type::Long {
        // `-2147483648&` does not compile: the positive part overflows first.
        return lit("(-2147483647& - 1&)");
    }
    if hex && value >= 0 {
        return lit(format!("&H{value:X}{suffix}"));
    }
    lit(format!("{value}{suffix}"))
}

/// `CDec` expression for a decimal literal, independent of the locale's decimal separator.
fn decimal_lit(text: &str) -> VExpr {
    let t = text.to_ascii_lowercase();
    if t.contains('e') {
        return call("CDec", vec![lit(double_lit(t.parse().unwrap_or(0.0)))]);
    }
    match t.split_once('.') {
        Some((int, frac)) => {
            let digits = format!("{int}{frac}");
            let digits = digits.trim_start_matches('0');
            let digits = if digits.is_empty() { "0" } else { digits };
            let scale = format!("1{}", "0".repeat(frac.len()));
            lit(format!("(CDec(\"{digits}\") / CDec(\"{scale}\"))"))
        }
        None => call("CDec", vec![lit(format!("\"{t}\""))]),
    }
}

/// Text of a parameter default when it is a plain literal (usable after `Optional ... =`).
pub(crate) fn literal_default(d: &Expr, ty: &Type) -> Option<String> {
    match &d.peel().kind {
        ExprKind::Int { value, .. } => Some(match ty {
            Type::Double => format!("{value}#"),
            Type::Float => format!("{value}!"),
            Type::Long => format!("{value}^"),
            Type::Byte | Type::Short => format!("{value}"),
            _ => i32::try_from(*value).ok().map(|v| format!("{v}&"))?,
        }),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            operand,
        } => match &operand.peel().kind {
            ExprKind::Int { value, .. } if *value < i32::MAX as i128 => Some(format!(
                "-{value}{}",
                if *ty == Type::Double { "#" } else { "&" }
            )),
            ExprKind::Float {
                text,
                suffix: NumSuffix::None | NumSuffix::D,
            } => Some(double_lit(-text.parse::<f64>().ok()?)),
            _ => None,
        },
        ExprKind::Float {
            text,
            suffix: NumSuffix::None | NumSuffix::D,
        } => Some(double_lit(text.parse().ok()?)),
        ExprKind::Str(s) if s.chars().all(|c| !c.is_control() && c.is_ascii()) => {
            Some(format!("\"{}\"", s.replace('"', "\"\"")))
        }
        ExprKind::Bool(b) => Some(if *b { "True".into() } else { "False".into() }),
        _ => None,
    }
}

/// Bit mask of `fastMode({...})`: screen=1 events=2 calc=4 alerts=8 statusBar=16 cursor=32 interactive=64.
pub(crate) fn fast_mode_mask(call_expr: &Expr) -> i64 {
    let mut mask = 63;
    if let ExprKind::Call { args, .. } = &call_expr.peel().kind {
        if let Some(ExprKind::Object(props)) = args.first().map(|a| &a.kind) {
            for p in props {
                let bit = match p.key.text.as_str() {
                    "screen" => 1,
                    "events" => 2,
                    "calc" => 4,
                    "alerts" => 8,
                    "statusBar" => 16,
                    "cursor" => 32,
                    "interactive" => 64,
                    _ => 0,
                };
                match (&p.value.kind, bit) {
                    // `interactive: false` *adds* blocking of user input (spec §10.1).
                    (ExprKind::Bool(false), 64) => mask |= 64,
                    (ExprKind::Bool(true), 64) => mask &= !64,
                    (ExprKind::Bool(false), b) => mask &= !b,
                    (ExprKind::Bool(true), b) => mask |= b,
                    _ => {}
                }
            }
        }
    }
    mask
}

/// VBA functions that have a faster `$` variant returning `String` (spec §8.5).
const DOLLAR_FUNCS: &[&str] = &[
    "Left", "Right", "Mid", "Trim", "LTrim", "RTrim", "UCase", "LCase", "Space", "String", "Chr",
    "ChrW", "Hex", "Oct", "Str", "Format", "Environ", "Dir", "CurDir",
];

fn simple(v: &VExpr) -> bool {
    match v {
        VExpr::Lit(_) | VExpr::Name(_) => true,
        VExpr::Neg(x) => simple(x),
        _ => false,
    }
}

impl<'p> Lowerer<'p> {
    pub(crate) fn str_lit(&self, s: &str) -> VExpr {
        let mut parts: Vec<VExpr> = Vec::new();
        let mut cur = String::new();
        let flush = |cur: &mut String, parts: &mut Vec<VExpr>| {
            if !cur.is_empty() {
                parts.push(lit(format!("\"{}\"", cur.replace('"', "\"\""))));
                cur.clear();
            }
        };
        let mut chars = s.chars().peekable();
        while let Some(c) = chars.next() {
            let special = match c {
                '\r' if chars.peek() == Some(&'\n') => {
                    chars.next();
                    Some(lit("vbCrLf"))
                }
                '\r' => Some(lit("vbCr")),
                '\n' => Some(lit("vbLf")),
                '\t' => Some(lit("vbTab")),
                '\0' => Some(lit("vbNullChar")),
                c if c.is_control() => Some(call("ChrW", vec![long(c as i64)])),
                // Characters the target code page cannot store become ChrW (spec §13.7).
                c if !c.is_ascii() && !self.encodable(c.encode_utf8(&mut [0; 4])) => {
                    let mut units = [0u16; 2];
                    let enc = c.encode_utf16(&mut units);
                    let mut e = call("ChrW", vec![lit(format!("&H{:X}&", enc[0]))]);
                    if enc.len() == 2 {
                        e = bin("&", e, call("ChrW", vec![lit(format!("&H{:X}&", enc[1]))]));
                    }
                    Some(e)
                }
                _ => None,
            };
            match special {
                Some(e) => {
                    flush(&mut cur, &mut parts);
                    parts.push(e);
                }
                None => {
                    cur.push(c);
                    // Keep literals well below VBA's 1023-character line limit (spec §13.8).
                    if cur.len() > 400 {
                        flush(&mut cur, &mut parts);
                    }
                }
            }
        }
        flush(&mut cur, &mut parts);
        let mut it = parts.into_iter();
        match it.next() {
            None => lit("\"\""),
            Some(first) => it.fold(first, |acc, p| bin("&", acc, p)),
        }
    }

    fn null_for(&mut self, ty: &Type) -> VExpr {
        if ty.is_object() || ty.is_mixed_union() {
            lit("Nothing")
        } else {
            self.default_value(ty).unwrap_or_else(|| lit("Empty"))
        }
    }

    /// Stores `v` in a fresh temporary and returns it.
    pub(crate) fn spill(&mut self, v: VExpr, ty: &Type) -> VExpr {
        if simple(&v) {
            return v;
        }
        let t = VExpr::Name(self.temp(ty));
        self.store(&t, ty, v, ty);
        t
    }

    /// Lowers operands left to right. When a later operand needs statements, earlier
    /// operands are saved in temporaries first so that evaluation order is kept (spec §8.3).
    pub(crate) fn operands(&mut self, es: &[&Expr]) -> Vec<VExpr> {
        let mut out: Vec<VExpr> = Vec::with_capacity(es.len());
        for (i, e) in es.iter().enumerate() {
            let mark = self.out_len();
            let v = self.expr(e);
            if self.out_len() > mark {
                let mut at = mark;
                for j in 0..i {
                    if simple(&out[j]) && !matches!(out[j], VExpr::Name(_)) {
                        continue;
                    }
                    if let VExpr::Name(n) = &out[j] {
                        if n.starts_with("ras_t") {
                            continue;
                        }
                    }
                    let ty = self.ty(es[j]);
                    if ty == Type::Null || !ty.is_value() {
                        continue;
                    }
                    let prev = std::mem::replace(&mut out[j], VExpr::Missing);
                    let t = VExpr::Name(self.temp(&ty));
                    let (stmts, _) = self.capture(|lw| lw.store(&t, &ty, prev, &ty));
                    let block = self.pc.blocks.last_mut().unwrap();
                    for s in stmts {
                        block.insert(at, s);
                        at += 1;
                    }
                    out[j] = t;
                }
            }
            out.push(v);
        }
        out
    }

    pub(crate) fn cond(&mut self, e: &Expr) -> VExpr {
        self.expr(e)
    }

    /// How a value of type `vty` is assigned to a target of type `tty`.
    pub(crate) fn store(&mut self, target: &VExpr, tty: &Type, value: VExpr, vty: &Type) {
        if *vty == Type::Null {
            if tty.is_object() || tty.is_mixed_union() {
                self.emit(set(target.clone(), lit("Nothing")));
            } else if let Some(d) = self.default_value(tty) {
                self.emit(assign(target.clone(), d));
            }
            return;
        }
        if tty.is_array() || vty.is_array() {
            self.emit(assign(target.clone(), value));
            return;
        }
        if *tty == Type::Decimal && !matches!(vty, Type::Decimal | Type::Any | Type::Error) {
            self.emit(assign(target.clone(), call("CDec", vec![value])));
            return;
        }
        // A value stored in a union keeps the member type it is converted to (`int` into `string | double`
        // becomes a Double), so that `typeof` sees that member (spec §2.6.2).
        let value = self.to_union_member(value, vty, tty);
        let v_obj = vty.is_object();
        let v_dyn = vty.is_dynamic();
        let t_dyn = tty.is_dynamic();
        if v_obj || (v_dyn && tty.is_object()) {
            self.emit(set(target.clone(), value));
        } else if v_dyn && t_dyn {
            // The value may or may not be an object: decide at run time (spec §2.3).
            let direct = match target {
                VExpr::Name(n) => *n != self.pc.proc_name,
                VExpr::Call(f, _) => matches!(**f, VExpr::Name(_)),
                _ => false,
            };
            if direct {
                self.emit(call_stmt("RAS_Assign", vec![target.clone(), value]));
            } else {
                let t = VExpr::Name(self.temp(&Type::Any));
                self.emit(call_stmt("RAS_Assign", vec![t.clone(), value]));
                self.emit(if_else(
                    call("IsObject", vec![t.clone()]),
                    vec![set(target.clone(), t.clone())],
                    vec![assign(target.clone(), t)],
                ));
            }
        } else {
            self.emit(assign(target.clone(), value));
        }
    }

    /// Converts a value that is not one of the members of the value union `tty` to the member it widens to.
    pub(crate) fn to_union_member(&mut self, v: VExpr, vty: &Type, tty: &Type) -> VExpr {
        let Some(ms) = tty.union_members() else {
            return v;
        };
        if vty.union_members().is_some()
            || vty.is_dynamic()
            || vty.is_object()
            || ms.contains(vty)
            || !tty.is_value_union()
        {
            return v;
        }
        let widened = |m: &Type| {
            matches!(
                (vty, m),
                (
                    Type::Byte,
                    Type::Short
                        | Type::Int
                        | Type::Long
                        | Type::Float
                        | Type::Double
                        | Type::Currency
                        | Type::Decimal
                ) | (
                    Type::Short,
                    Type::Int
                        | Type::Long
                        | Type::Float
                        | Type::Double
                        | Type::Currency
                        | Type::Decimal
                ) | (
                    Type::Int | Type::Enum(_),
                    Type::Long
                        | Type::Float
                        | Type::Double
                        | Type::Currency
                        | Type::Decimal
                        | Type::Ptr
                ) | (Type::Long, Type::Float | Type::Double | Type::Decimal)
                    | (Type::Float, Type::Double)
                    | (Type::Byte | Type::Short, Type::Enum(_))
                    | (Type::Enum(_), Type::Int)
            )
        };
        match ms.iter().find(|m| widened(m)) {
            Some(m) => {
                let m = m.clone();
                self.convert(v, vty, &m)
            }
            None => v,
        }
    }

    pub(crate) fn clear_array(&mut self, target: &VExpr, ty: &Type) {
        if self.vtype(ty).1 {
            self.emit(VStmt::Erase(target.clone()));
        } else {
            self.emit(assign(target.clone(), lit("Empty")));
        }
    }

    pub(crate) fn reset_var(&mut self, target: &VExpr, ty: &Type) {
        if ty.is_array() {
            self.clear_array(target, ty);
        } else if let Type::Struct(s) = ty {
            // A `Type` has no literal default: reset field by field.
            let def = &self.prog.structs[*s];
            for (fname, fty) in &def.fields {
                self.reset_var(&target.clone().member(names::member_ident(fname)), fty);
            }
        } else if let Some(d) = self.default_value(ty) {
            if ty.is_object() {
                self.emit(set(target.clone(), d));
            } else {
                self.emit(assign(target.clone(), d));
            }
        }
    }

    /// An array expression that can be used several times (a variable or field; otherwise a temporary copy).
    pub(crate) fn stable_array(&mut self, e: &Expr) -> VExpr {
        let ty = self.ty(e);
        let is_place = match &e.peel().kind {
            ExprKind::Name(_) | ExprKind::This => true,
            ExprKind::Member { .. } => {
                matches!(self.info().res.get(&e.peel().id), Some(Res::Field(..)))
            }
            _ => false,
        };
        let v = self.expr(e);
        if is_place {
            v
        } else {
            let t = VExpr::Name(self.temp(&ty));
            self.emit(assign(t.clone(), v));
            t
        }
    }

    fn to_string(&mut self, v: VExpr, ty: &Type) -> VExpr {
        match ty {
            Type::Str => v,
            Type::Byte
            | Type::Short
            | Type::Int
            | Type::Long
            | Type::Enum(_)
            | Type::Ptr
            | Type::Date => call("CStr", vec![v]),
            Type::Float | Type::Double | Type::Currency | Type::Decimal => {
                call("RAS_NumStr", vec![v])
            }
            Type::Bool => call("RAS_BoolStr", vec![v]),
            _ => call("RAS_Str", vec![v]),
        }
    }

    fn convert(&mut self, v: VExpr, from: &Type, to: &Type) -> VExpr {
        use Type::*;
        if from == to || matches!(to, Any | Error) {
            return v;
        }
        let fix = |v: VExpr| {
            if from.is_integer() || matches!(from, Enum(_) | Bool) {
                v
            } else {
                call("Fix", vec![v])
            }
        };
        // `True` is -1 in VBA: booleans convert to 1 / 0 (spec §13.2).
        let v = if *from == Bool && to.is_numeric() {
            VExpr::Neg(Box::new(call("CLng", vec![v])))
        } else {
            v
        };
        match to {
            Int | Enum(_) => {
                if matches!(from, Bool | Enum(_) | Byte | Short) {
                    if *from == Bool {
                        v
                    } else {
                        call("CLng", vec![v])
                    }
                } else {
                    // CLng rounds half to even; RAS truncates toward zero like C# (spec §2.3).
                    call("CLng", vec![fix(v)])
                }
            }
            Short => call("CInt", vec![fix(v)]),
            Byte => call("CByte", vec![fix(v)]),
            Long => call("CLngLng", vec![fix(v)]),
            Double => call("CDbl", vec![v]),
            Float => call("CSng", vec![v]),
            Currency => call("CCur", vec![v]),
            Decimal => call("CDec", vec![v]),
            Date => call("CDate", vec![v]),
            Ptr => call("CLngPtr", vec![v]),
            Str => self.to_string(v, from),
            Bool => call("CBool", vec![v]),
            _ => v,
        }
    }

    // ------------------------------------------------------------------ expressions

    pub(crate) fn expr(&mut self, e: &Expr) -> VExpr {
        if let Some(slot) = self.bulk_slot(e) {
            return slot;
        }
        match &e.kind {
            ExprKind::Error | ExprKind::Super => lit("Empty"),
            ExprKind::Object(_) => match self.ty(e) {
                ty @ Type::Struct(_) => {
                    let t = VExpr::Name(self.temp(&ty));
                    let once = self.runs_once();
                    self.struct_lit_fill(e, &t, once);
                    t
                }
                _ => lit("Empty"),
            },
            ExprKind::Array(_) if matches!(self.ty(e), Type::Struct(_)) => {
                let ty = self.ty(e);
                let t = VExpr::Name(self.temp(&ty));
                let once = self.runs_once();
                self.struct_lit_fill(e, &t, once);
                t
            }
            ExprKind::Int { value, hex, .. } => int_lit(*value, &self.ty(e), *hex),
            ExprKind::Float { text, suffix } => match suffix {
                NumSuffix::F => {
                    let f: f64 = text.parse().unwrap_or(0.0);
                    lit(if f == f.trunc() && f.abs() < 1e7 {
                        format!("{f:.0}!")
                    } else {
                        format!("{f}!")
                    })
                }
                NumSuffix::M => decimal_lit(text),
                _ => lit(double_lit(text.parse().unwrap_or(0.0))),
            },
            ExprKind::Str(s) => self.str_lit(s),
            ExprKind::Template(parts) => {
                let mut acc: Option<VExpr> = None;
                for p in parts {
                    let v = match p {
                        ast::TemplatePart::Str(s) => self.str_lit(s),
                        ast::TemplatePart::Expr(x) => {
                            let ty = self.ty(x);
                            let v = self.expr(x);
                            self.to_string(v, &ty)
                        }
                    };
                    acc = Some(match acc {
                        None => v,
                        Some(a) => bin("&", a, v),
                    });
                }
                acc.unwrap_or_else(|| lit("\"\""))
            }
            ExprKind::Bool(b) => lit(if *b { "True" } else { "False" }),
            ExprKind::Null | ExprKind::Undefined => lit("Nothing"),
            ExprKind::Name(id) => self.name_expr(e, &id.text),
            ExprKind::This => match self.info().res.get(&e.id) {
                Some(Res::Local(l)) => self.local_expr(*l),
                _ => self.frame().this.clone().unwrap_or_else(|| lit("Me")),
            },
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => self.expr(inner),
            ExprKind::Array(_) => self.array_value(e),
            ExprKind::Arrow(a) => self.closure_value(e, a),
            ExprKind::Unary { op, operand } => {
                if let (UnaryOp::Neg, ExprKind::Int { value, .. }) = (op, &operand.peel().kind) {
                    return int_lit(-*value, &self.ty(e), false);
                }
                let v = self.expr(operand);
                match op {
                    UnaryOp::Not => v.not(),
                    UnaryOp::Neg => match v {
                        VExpr::Lit(s)
                            if !s.starts_with('-')
                                && !s.starts_with('(')
                                && !s.starts_with('"')
                                && !s.starts_with('&') =>
                        {
                            lit(format!("-{s}"))
                        }
                        other => VExpr::Neg(Box::new(other)),
                    },
                    UnaryOp::Plus => v,
                    UnaryOp::BitNot => VExpr::Not(Box::new(v)),
                    UnaryOp::TypeOf => call("TypeName", vec![v]),
                }
            }
            ExprKind::IncDec {
                inc,
                prefix,
                operand,
            } => {
                let (t, ty) = self.lvalue_ex(operand, true);
                let one = self.one_for(&ty);
                let next = bin(if *inc { "+" } else { "-" }, t.clone(), one);
                if *prefix {
                    self.emit(assign(t.clone(), next));
                    t
                } else {
                    let old = VExpr::Name(self.temp(&ty));
                    self.emit(assign(old.clone(), t.clone()));
                    self.emit(assign(t, next));
                    old
                }
            }
            ExprKind::Binary { op, lhs, rhs } => self.binary(e, *op, lhs, rhs),
            ExprKind::Assign { op, target, value } => self.assign_expr(*op, target, value),
            ExprKind::Ternary { .. } => {
                let ty = self.ty(e);
                let t = VExpr::Name(self.temp(&ty));
                self.expr_into(e, &t, &ty);
                t
            }
            ExprKind::Call { .. } => self.call(e, true).unwrap_or_else(|| lit("Empty")),
            ExprKind::New { ty, args } => self.new_expr(e, ty, self.info().call_args(e.id, args)),
            // Reordered into the call's arguments by the checker (spec §3.3).
            ExprKind::NamedArg { value, .. } => self.expr(value),
            ExprKind::Member { object, name, .. } => self.member_expr(e, object, &name.text),
            ExprKind::Index { object, index } => self.index_expr(e, object, index),
            ExprKind::As { expr, .. } => {
                let from = self.ty(expr);
                let to = self.ty(e);
                if from == Type::Null {
                    return self.null_for(&to);
                }
                let v = self.expr(expr);
                self.convert(v, &from, &to)
            }
            ExprKind::InstanceOf { expr, ty } => {
                let v = self.expr(expr);
                let tps: Vec<TypeParamDef> = Vec::new();
                let _ = tps;
                let target = self.resolve_static_type(ty);
                let cname = match &target {
                    Type::Class(c, args) => self.class_name(*c, args),
                    _ => "Object".into(),
                };
                lit(format!("(TypeOf {} Is {cname})", v.print()))
            }
        }
    }

    /// Resolves a type written in an expression (`x instanceof T`) in the current instantiation.
    fn resolve_static_type(&mut self, t: &ast::TypeRef) -> Type {
        let ast::TypeRef::Named { path, args, .. } = t else {
            return Type::Error;
        };
        let m = self.src_module();
        let sym = if path.len() == 1 {
            if let Some(b) = self.frame().env.get(&path[0].text) {
                return b.clone();
            }
            self.prog.modules[m]
                .scope
                .get(&path[0].text)
                .or_else(|| self.prog.globals.get(&path[0].text))
                .cloned()
        } else {
            match self.prog.modules[m].scope.get(&path[0].text) {
                Some(Sym::Module(t)) => self.prog.modules[*t].exports.get(&path[1].text).cloned(),
                _ => self
                    .prog
                    .namespaces
                    .get(&path[0].text)
                    .and_then(|n| n.get(&path[1].text))
                    .cloned(),
            }
        };
        match sym {
            Some(Sym::Class(c)) => {
                let targs = args.iter().map(|a| self.resolve_static_type(a)).collect();
                Type::Class(c, targs)
            }
            _ => Type::Error,
        }
    }

    fn one_for(&self, ty: &Type) -> VExpr {
        lit(match ty {
            Type::Long => "1^",
            Type::Double => "1#",
            Type::Float => "1!",
            Type::Byte | Type::Short => "1",
            _ => "1&",
        })
    }

    fn name_expr(&mut self, e: &Expr, text: &str) -> VExpr {
        match self.info().res.get(&e.id) {
            Some(Res::Local(l)) => self.local_expr(*l),
            Some(Res::Global(v)) => match self.pc.demoted.get(v) {
                Some(n) => VExpr::Name(n.clone()),
                None => self.var_ref(*v),
            },
            Some(Res::Fn(f)) => {
                // Host function read like a property (`Now`, `Cells`).
                let n = VExpr::Name(self.prog.fns[*f].vba_name().to_string());
                match (
                    &self.pc.qualifier,
                    self.info().auto_qualified.contains(&e.id),
                ) {
                    (Some(q), true) => q.clone().member(self.prog.fns[*f].vba_name().to_string()),
                    _ => n,
                }
            }
            Some(Res::EnumMember(en, i)) => self.enum_value(*en, *i),
            Some(Res::Intrinsic(i)) => self
                .intrinsic(e, *i, None, &[], true)
                .unwrap_or_else(|| lit("Empty")),
            _ => VExpr::Name(names::valid_ident(text)),
        }
    }

    fn enum_value(&self, en: EnumId, i: usize) -> VExpr {
        let d = &self.prog.enums[en];
        if d.host && !d.late_bound {
            VExpr::Name(d.members[i].name.clone())
        } else {
            // RAS enums are emitted as literals so that private enums work from any module.
            long(d.members[i].value)
        }
    }

    fn member_expr(&mut self, e: &Expr, object: &Expr, member: &str) -> VExpr {
        let info = self.info();
        let res = info.res.get(&e.id).cloned();
        let obj_is_ns = matches!(info.ty(object.id), Type::Ns(_));
        match res {
            Some(Res::Field(c, idx)) => {
                let fty = self.ty(e);
                let o = if obj_is_ns {
                    None
                } else {
                    Some(self.expr(object))
                };
                self.field_target(c, idx, &fty, o)
            }
            Some(Res::Global(v)) => self.var_ref(v),
            Some(Res::EnumMember(en, i)) => self.enum_value(en, i),
            Some(Res::Intrinsic(i)) => self
                .intrinsic(e, i, Some(object), &[], true)
                .unwrap_or_else(|| lit("Empty")),
            Some(Res::Fn(f)) => VExpr::Name(self.prog.fns[f].vba_name().to_string()),
            Some(Res::Property { get: Some(g), .. }) if obj_is_ns => {
                // static getter
                let callee = self.static_callee(g);
                VExpr::Call(Box::new(callee), vec![])
            }
            Some(Res::Method(f)) => {
                let o = self.expr(object);
                o.member(self.prog.fns[f].vba_name().to_string())
            }
            _ => {
                let o = self.expr(object);
                let own = o.is_lit("Me");
                if own {
                    // Private members cannot be reached through `Me.`.
                    VExpr::Name(names::member_ident(member))
                } else {
                    o.member(
                        self.late_member_name(object, member)
                            .unwrap_or_else(|| names::member_ident(member)),
                    )
                }
            }
        }
    }

    /// A member reached by late binding (`activeSheet.cells`): when a bundled host interface has a
    /// member of that name, it is written with the type library's spelling (`Cells`). VBA finds it
    /// either way; this keeps the VBE from showing two spellings of one name.
    fn late_member_name(&self, object: &Expr, written: &str) -> Option<String> {
        let late = match self.ty(object) {
            Type::Any | Type::Object | Type::Union(_) => true,
            Type::Class(c, _) => self.prog.classes[c].kind == ClassKind::Host,
            _ => false,
        };
        if !late {
            return None;
        }
        self.prog
            .host_member_spelling
            .get(&written.to_lowercase())
            .cloned()
    }

    fn static_callee(&mut self, f: FnId) -> VExpr {
        let name = self.fn_name(f, &[]);
        let m = self.prog.fns[f].module;
        VExpr::Name(self.prog.modules[m].vba_name.clone()).member(name)
    }

    fn index_expr(&mut self, e: &Expr, object: &Expr, index: &Expr) -> VExpr {
        let ot = self.ty(object);
        match &ot {
            Type::Array(_) => {
                // `a[r][c]` on a 2-D array is one access `a(r, c)`.
                if let ExprKind::Index {
                    object: base,
                    index: row,
                } = &object.peel().kind
                {
                    if self.ty(base).dims() == 2 {
                        let b = self.expr(base);
                        let ix = self.operands(&[row, index]);
                        return b.index(ix);
                    }
                }
                let b = self.expr(object);
                let i = self.expr(index);
                b.index(vec![i])
            }
            Type::Str => {
                let ops = self.operands(&[object, index]);
                let mut it = ops.into_iter();
                let (s, i) = (it.next().unwrap(), it.next().unwrap());
                call("Mid$", vec![s, i.plus(1), long(1)])
            }
            Type::Struct(_) => {
                let b = self.expr(object);
                b.member(tuple_field(index))
            }
            Type::Map(_, v) => {
                let ops = self.operands(&[object, index]);
                let mut it = ops.into_iter();
                let (d, k) = (it.next().unwrap(), it.next().unwrap());
                self.map_get(d, k, &self.subst(v))
            }
            _ => {
                let _ = e;
                let ops = self.operands(&[object, index]);
                let mut it = ops.into_iter();
                let (o, i) = (it.next().unwrap(), it.next().unwrap());
                o.index(vec![i])
            }
        }
    }

    fn map_get(&mut self, d: VExpr, k: VExpr, vty: &Type) -> VExpr {
        // `d(k)` would silently add a missing key.
        if vty.is_object() {
            return call("RAS_MapGetObj", vec![d, k]);
        }
        let v = call("RAS_MapGet", vec![d, k]);
        // A missing key yields the zero value of the value type, not Empty (`CStr(Empty)` is "", not "0").
        let conv = match vty {
            Type::Int | Type::Enum(_) => "CLng",
            Type::Short => "CInt",
            Type::Byte => "CByte",
            Type::Long => "CLngLng",
            Type::Double => "CDbl",
            Type::Float => "CSng",
            Type::Currency => "CCur",
            Type::Decimal => "CDec",
            Type::Date => "CDate",
            Type::Ptr => "CLngPtr",
            Type::Bool => "CBool",
            Type::Str => "CStr",
            _ => return v,
        };
        call(conv, vec![v])
    }

    /// Lowers an assignment target to a VBA l-value plus its type.
    pub(crate) fn lvalue(&mut self, e: &Expr) -> (VExpr, Type) {
        self.lvalue_ex(e, false)
    }

    /// `stable`: the target is read and written (`a[i()] += 1`), so indices and receivers
    /// with side effects are evaluated once into temporaries.
    pub(crate) fn lvalue_ex(&mut self, e: &Expr, stable: bool) -> (VExpr, Type) {
        let ty = self.ty(e);
        if let Some(slot) = self.bulk_slot(e.peel()) {
            return (slot, Type::Any);
        }
        let inner = e.peel();
        let v = match &inner.kind {
            ExprKind::Index { object, index } => {
                let ot = self.ty(object);
                if ot.is_array() {
                    if let ExprKind::Index {
                        object: base,
                        index: row,
                    } = &object.peel().kind
                    {
                        if self.ty(base).dims() == 2 {
                            let (b, _) = self.lvalue_ex(base, stable);
                            let r = self.expr(row);
                            let r = if stable { self.spill(r, &Type::Int) } else { r };
                            let c = self.expr(index);
                            let c = if stable { self.spill(c, &Type::Int) } else { c };
                            return (b.index(vec![r, c]), ty);
                        }
                    }
                    let (b, _) = self.lvalue_ex(object, stable);
                    let i = self.expr(index);
                    let i = if stable { self.spill(i, &Type::Int) } else { i };
                    b.index(vec![i])
                } else if matches!(ot, Type::Struct(_)) {
                    let (b, _) = self.lvalue_ex(object, stable);
                    b.member(tuple_field(index))
                } else {
                    let o = self.expr(object);
                    let o = if stable { self.spill(o, &ot) } else { o };
                    let it = self.ty(index);
                    let i = self.expr(index);
                    let i = if stable { self.spill(i, &it) } else { i };
                    o.index(vec![i])
                }
            }
            ExprKind::Member { object, name, .. } => {
                let res = self.info().res.get(&inner.id).cloned();
                let obj_is_ns = matches!(self.info().ty(object.id), Type::Ns(_));
                match res {
                    Some(Res::Field(c, idx)) => {
                        let o = if obj_is_ns {
                            None
                        } else {
                            let ot = self.ty(object);
                            let o = self.expr(object);
                            Some(if stable { self.spill_object(o, &ot) } else { o })
                        };
                        self.field_target(c, idx, &ty, o)
                    }
                    Some(Res::Global(v)) => self.var_ref(v),
                    _ => {
                        let ot = self.ty(object);
                        let o = self.expr(object);
                        let o = if stable { self.spill_object(o, &ot) } else { o };
                        if o.is_lit("Me") {
                            VExpr::Name(names::member_ident(&name.text))
                        } else {
                            o.member(
                                self.late_member_name(object, &name.text)
                                    .unwrap_or_else(|| names::member_ident(&name.text)),
                            )
                        }
                    }
                }
            }
            _ => self.expr(inner),
        };
        (v, ty)
    }

    /// Keeps member chains (`ws.Range("A1").Font`) as they are, but evaluates calls only once.
    fn spill_object(&mut self, o: VExpr, ty: &Type) -> VExpr {
        fn pure_chain(v: &VExpr) -> bool {
            match v {
                VExpr::Lit(_) | VExpr::Name(_) => true,
                VExpr::Member(o, _) => pure_chain(o),
                _ => false,
            }
        }
        if pure_chain(&o) {
            o
        } else {
            let t = VExpr::Name(self.temp(ty));
            self.emit(set(t.clone(), o));
            t
        }
    }

    fn arith(
        &mut self,
        op: BinaryOp,
        l: VExpr,
        lt: &Type,
        r: VExpr,
        rt: &Type,
        result: &Type,
    ) -> VExpr {
        let small = |t: &Type| matches!(t, Type::Byte | Type::Short);
        // `Integer * Integer` overflows as Integer in VBA: promote to Long (spec §2.3).
        let l = if small(lt) && small(rt) && op.is_arithmetic() {
            call("CLng", vec![l])
        } else {
            l
        };
        let int_like = |t: &Type| t.is_integer() || matches!(t, Type::Enum(_));
        match op {
            BinaryOp::Add if *result == Type::Str => {
                let a = self.to_string(l, lt);
                let b = self.to_string(r, rt);
                bin("&", a, b)
            }
            BinaryOp::Add => bin("+", l, r),
            BinaryOp::Sub => bin("-", l, r),
            BinaryOp::Mul => bin("*", l, r),
            // Integer division truncates toward zero, like C# (spec §2.3).
            BinaryOp::Div if int_like(lt) && int_like(rt) => bin("\\", l, r),
            BinaryOp::Div => bin("/", l, r),
            BinaryOp::Rem if int_like(lt) && int_like(rt) => bin("Mod", l, r),
            // VBA's Mod rounds real operands first.
            BinaryOp::Rem => call("RAS_FMod", vec![l, r]),
            BinaryOp::Pow => bin("^", l, r),
            BinaryOp::Lt => bin("<", l, r),
            BinaryOp::Le => bin("<=", l, r),
            BinaryOp::Gt => bin(">", l, r),
            BinaryOp::Ge => bin(">=", l, r),
            BinaryOp::BitAnd => bin("And", l, r),
            BinaryOp::BitOr => bin("Or", l, r),
            BinaryOp::BitXor => bin("Xor", l, r),
            BinaryOp::Shl => call("RAS_Shl", vec![l, r]),
            BinaryOp::Shr => call("RAS_Shr", vec![l, r]),
            BinaryOp::Eq => bin("=", l, r),
            BinaryOp::Ne => bin("<>", l, r),
            BinaryOp::And => bin("And", l, r),
            BinaryOp::Or | BinaryOp::Coalesce => bin("Or", l, r),
        }
    }

    /// Operands that can be evaluated eagerly without raising or side effects.
    fn is_safe_operand(&self, e: &Expr) -> bool {
        match &e.peel().kind {
            ExprKind::Int { .. }
            | ExprKind::Float { .. }
            | ExprKind::Str(_)
            | ExprKind::Bool(_)
            | ExprKind::Null => true,
            ExprKind::Name(_) => matches!(
                self.info().res.get(&e.peel().id),
                Some(Res::Local(_) | Res::EnumMember(..))
            ),
            ExprKind::Unary {
                op: UnaryOp::Not,
                operand,
            } => self.is_safe_operand(operand),
            ExprKind::Binary { op, lhs, rhs } => {
                (op.is_comparison() || matches!(op, BinaryOp::And | BinaryOp::Or))
                    && self.is_safe_operand(lhs)
                    && self.is_safe_operand(rhs)
                    && !self.ty(lhs).is_object()
                    && !matches!(self.ty(lhs), Type::Any | Type::Union(_))
                    && !matches!(self.ty(rhs), Type::Any | Type::Union(_))
            }
            _ => false,
        }
    }

    fn binary(&mut self, e: &Expr, op: BinaryOp, lhs: &Expr, rhs: &Expr) -> VExpr {
        let (lt, rt) = (self.ty(lhs), self.ty(rhs));
        match op {
            BinaryOp::And | BinaryOp::Or => {
                let l = self.expr(lhs);
                let (pre, r) = self.capture(|lw| lw.expr(rhs));
                if pre.is_empty() && self.is_safe_operand(rhs) {
                    return bin(if op == BinaryOp::And { "And" } else { "Or" }, l, r);
                }
                // VBA's And / Or evaluate both sides: short-circuit by hand (spec §13.2).
                let t = VExpr::Name(self.temp(&Type::Bool));
                self.emit(assign(t.clone(), l));
                let mut body = pre;
                body.push(assign(t.clone(), r));
                let c = if op == BinaryOp::And {
                    t.clone()
                } else {
                    t.clone().not()
                };
                self.emit(if_then(c, body));
                t
            }
            BinaryOp::Coalesce => {
                let ty = self.ty(e);
                let l = self.expr(lhs);
                let l = self.spill(l, &lt);
                let t = VExpr::Name(self.temp(&ty));
                let test = if lt.is_object() {
                    bin("Is", l.clone(), lit("Nothing"))
                } else {
                    call("RAS_IsNullish", vec![l.clone()])
                };
                let (then_b, _) = self.capture(|lw| lw.expr_into(rhs, &t, &ty));
                let (else_b, _) = self.capture(|lw| lw.store(&t, &ty, l, &lt));
                self.emit(if_else(test, then_b, else_b));
                t
            }
            BinaryOp::Eq | BinaryOp::Ne => {
                let negate = op == BinaryOp::Ne;
                let v = if lt == Type::Null && rt == Type::Null {
                    lit("True")
                } else if lt == Type::Null || rt == Type::Null {
                    let (other, oty) = if lt == Type::Null {
                        (rhs, &rt)
                    } else {
                        (lhs, &lt)
                    };
                    let v = self.expr(other);
                    if oty.is_object() {
                        bin("Is", v, lit("Nothing"))
                    } else if oty.is_dynamic() {
                        call("RAS_IsNullish", vec![v])
                    } else {
                        lit("False")
                    }
                } else {
                    let ops = self.operands(&[lhs, rhs]);
                    let mut it = ops.into_iter();
                    let (l, r) = (it.next().unwrap(), it.next().unwrap());
                    // A union may hold a value of another kind: such values are not equal (spec §2.6.2).
                    if lt.is_value_union() || rt.is_value_union() {
                        call("RAS_UnionEq", vec![l, r])
                    // Objects compare by identity (spec §13.2).
                    } else if lt.is_object() || rt.is_object() {
                        bin("Is", l, r)
                    } else {
                        bin("=", l, r)
                    }
                };
                if negate {
                    v.not()
                } else {
                    v
                }
            }
            _ => {
                let result = self.ty(e);
                let ops = self.operands(&[lhs, rhs]);
                let mut it = ops.into_iter();
                let (l, r) = (it.next().unwrap(), it.next().unwrap());
                self.arith(op, l, &lt, r, &rt, &result)
            }
        }
    }

    /// Buffer element standing in for a cell access inside a bulk-transfer loop (spec §9).
    fn bulk_slot(&self, e: &Expr) -> Option<VExpr> {
        if self.pc.bulk.is_empty() || self.pc.frames.len() != self.pc.bulk_depth {
            return None;
        }
        self.pc.bulk.get(&e.id).map(|s| s.element.clone())
    }

    fn assign_expr(&mut self, op: Option<BinaryOp>, target: &Expr, value: &Expr) -> VExpr {
        if let (None, Some(slot)) = (op, self.bulk_slot(target.peel())) {
            // Cell values are never objects: a plain element assignment.
            let v = self.expr(value);
            self.emit(assign(slot.clone(), v));
            return slot;
        }
        let (t, tty) = self.lvalue_ex(target, op.is_some());
        match op {
            None => self.expr_into(value, &t, &tty),
            Some(op) => {
                let vty = self.ty(value);
                let v = self.expr(value);
                let result = if tty == Type::Str {
                    Type::Str
                } else {
                    tty.clone()
                };
                let combined = self.arith(op, t.clone(), &tty, v, &vty, &result);
                self.store(&t, &tty, combined, &result);
            }
        }
        t
    }

    /// Evaluates `e` directly into `target` (avoids temporaries for ternaries, arrays and fused loops).
    pub(crate) fn expr_into(&mut self, e: &Expr, target: &VExpr, ty: &Type) {
        let inner = e.peel();
        match &inner.kind {
            ExprKind::Ternary { cond, then, els } => {
                // `IIf` evaluates both branches: always an If statement (spec §13.2).
                let c = self.cond(cond);
                let (tb, _) = self.capture(|lw| lw.expr_into(then, target, ty));
                let (eb, _) = self.capture(|lw| lw.expr_into(els, target, ty));
                self.emit(flatten_if(c, tb, eb));
            }
            ExprKind::Object(_) | ExprKind::Array(_)
                if matches!(self.ty(inner), Type::Struct(_)) =>
            {
                self.struct_lit_into(inner, target)
            }
            // (`f(i) = x` inside function `f` would be a recursive call, so the return variable is excluded)
            ExprKind::Array(_) | ExprKind::New { .. } | ExprKind::Call { .. }
                if self.ty(inner).is_array()
                    && matches!(target, VExpr::Name(n) if *n != self.pc.proc_name || self.pc.rets.len() > 1) =>
            {
                if !self.array_into(inner, target, ty) {
                    let v = self.expr(inner);
                    let vt = self.ty(inner);
                    self.store(target, ty, v, &vt);
                }
            }
            _ => {
                let vt = self.ty(inner);
                let v = if vt == Type::Null {
                    lit("Nothing")
                } else {
                    self.expr(inner)
                };
                self.store(target, ty, v, &vt);
            }
        }
    }

    /// Builds a struct literal (`{ a: 1 }`, tuple `[1, "x"]`) in `target`, field by field.
    /// The fields are written straight into `target` when the values cannot observe it;
    /// otherwise into a temporary that is copied afterwards.
    pub(crate) fn struct_lit_into(&mut self, e: &Expr, target: &VExpr) {
        fn root(v: &VExpr) -> Option<&str> {
            match v {
                VExpr::Name(n) => Some(n),
                VExpr::Member(o, _) => root(o),
                VExpr::Call(f, _) => root(f),
                _ => None,
            }
        }
        let values: Vec<&Expr> = match &e.kind {
            ExprKind::Object(props) => props.iter().map(|p| &p.value).collect(),
            ExprKind::Array(items) => items.iter().collect(),
            _ => Vec::new(),
        };
        let direct = match root(target) {
            // The return variable cannot be read back (`f.x` on the right would be a call).
            Some(r) if *r == self.pc.proc_name && matches!(target, VExpr::Name(_)) => {
                self.pc.rets.len() == 1
            }
            Some(r) => {
                self.pc.locals.iter().any(|l| l.0 == r)
                    && !values.iter().any(|v| self.ast_mentions(v, r))
                    && !target_index_mentions(target, r)
            }
            None => false,
        };
        if direct {
            let fresh = match target {
                VExpr::Name(n) => {
                    (*n == self.pc.proc_name && self.pc.rets.len() == 1)
                        || self.pc.fresh_target.as_deref() == Some(n.as_str())
                }
                _ => false,
            };
            self.pc.fresh_target = None;
            self.struct_lit_fill(e, target, fresh);
        } else {
            self.pc.fresh_target = None;
            let ty = self.ty(e);
            let t = VExpr::Name(self.temp(&ty));
            let once = self.runs_once();
            self.struct_lit_fill(e, &t, once);
            self.emit(assign(target.clone(), t));
        }
    }

    /// The code being lowered runs at most once per call of the procedure.
    fn runs_once(&self) -> bool {
        self.pc.loop_depth == 0 && self.frame().inl.is_none()
    }

    /// Assigns every field of the struct literal `e` to `target.field` (omitted fields are reset).
    /// `fresh`: `target` still has its zero value.
    pub(crate) fn struct_lit_fill(&mut self, e: &Expr, target: &VExpr, fresh: bool) {
        let Type::Struct(s) = self.ty(e) else { return };
        let def = &self.prog.structs[s];
        let parts: Vec<(usize, &Expr)> = match &e.kind {
            ExprKind::Object(props) => props
                .iter()
                .filter_map(|p| {
                    def.fields
                        .iter()
                        .position(|f| f.0 == p.key.text)
                        .map(|i| (i, &p.value))
                })
                .collect(),
            ExprKind::Array(items) => items.iter().enumerate().take(def.fields.len()).collect(),
            _ => Vec::new(),
        };
        // Source order, so that side effects happen as written.
        for (i, v) in &parts {
            let (fname, fty) = &def.fields[*i];
            self.expr_into(v, &target.clone().member(names::member_ident(fname)), fty);
        }
        for (i, (fname, fty)) in def.fields.iter().enumerate() {
            if !fresh && !parts.iter().any(|p| p.0 == i) {
                self.reset_var(&target.clone().member(names::member_ident(fname)), fty);
            }
        }
    }

    /// `RAS_Len(arr)`. An array of structs cannot be passed as a Variant: it is measured by
    /// a helper generated next to the struct's `Type`.
    pub(crate) fn len_of(&mut self, arr: VExpr, ty: &Type) -> VExpr {
        if let Some(Type::Struct(s)) = ty.elem() {
            self.struct_len_used.insert(*s);
            return call(format!("RAS_Len_{}", self.struct_name(*s)), vec![arr]);
        }
        call("RAS_Len", vec![arr])
    }

    fn array_value(&mut self, e: &Expr) -> VExpr {
        let ty = self.ty(e);
        let t = VExpr::Name(self.temp(&ty));
        if !self.array_into(e, &t, &ty) {
            self.clear_array(&t, &ty);
        }
        t
    }

    /// Builds the array produced by `e` in the variable `dest`. Returns false when `e` is not such an expression.
    fn array_into(&mut self, e: &Expr, dest: &VExpr, ty: &Type) -> bool {
        match &e.kind {
            ExprKind::Array(items) => {
                if items.is_empty() {
                    self.clear_array(dest, ty);
                    return true;
                }
                let refs: Vec<&Expr> = items.iter().collect();
                let vals = self.operands(&refs);
                // The elements may read the old contents of `dest`.
                let dest_name = dest.print();
                let aliased = vals.iter().any(|v| mentions(v, &dest_name));
                let real = if aliased {
                    VExpr::Name(self.temp(ty))
                } else {
                    dest.clone()
                };
                self.emit(VStmt::ReDim {
                    preserve: false,
                    target: real.clone(),
                    dims: vec![(long(0), long(items.len() as i64 - 1))],
                });
                let elem_ty = ty.elem().cloned().unwrap_or(Type::Any);
                for (i, (v, item)) in vals.into_iter().zip(items.iter()).enumerate() {
                    let vt = self.ty(item);
                    self.store(&real.clone().index(vec![long(i as i64)]), &elem_ty, v, &vt);
                }
                if aliased {
                    self.emit(assign(dest.clone(), real));
                }
                true
            }
            ExprKind::New {
                ty: ast::TypeRef::Named { path, .. },
                args,
            } if path.len() == 1 && path[0].text == "Array" => {
                let refs: Vec<&Expr> = args.iter().collect();
                let sizes = self.operands(&refs);
                let dims: Vec<(VExpr, VExpr)> = sizes
                    .iter()
                    .map(|s| (long(0), s.clone().plus(-1)))
                    .collect();
                let redim = VStmt::ReDim {
                    preserve: false,
                    target: dest.clone(),
                    dims,
                };
                let consts: Vec<Option<i64>> = sizes.iter().map(|s| s.as_long()).collect();
                if consts.iter().all(|c| matches!(c, Some(n) if *n > 0)) {
                    self.emit(redim);
                } else if consts.iter().any(|c| matches!(c, Some(n) if *n <= 0)) {
                    self.clear_array(dest, ty);
                } else {
                    // `ReDim a(0 To -1)` is an error: a zero-length array is an unallocated one.
                    let mut cond: Option<VExpr> = None;
                    for s in &sizes {
                        if s.as_long().is_none() {
                            let c = bin(">", s.clone(), long(0));
                            cond = Some(match cond {
                                None => c,
                                Some(p) => bin("And", p, c),
                            });
                        }
                    }
                    let (clear, _) = self.capture(|lw| lw.clear_array(dest, ty));
                    self.emit(if_else(cond.unwrap(), vec![redim], clear));
                }
                true
            }
            ExprKind::Call { callee, args, .. } => {
                let Some(CallTarget::Intrinsic(i)) =
                    self.info().calls.get(&e.id).map(|c| c.target.clone())
                else {
                    return false;
                };
                let ExprKind::Member { object, .. } = &callee.peel().kind else {
                    return false;
                };
                self.array_intrinsic(e, i, object, args, dest, ty)
            }
            _ => false,
        }
    }

    /// Statement-context expression: no value is needed.
    pub(crate) fn expr_stmt(&mut self, e: &Expr) {
        let inner = e.peel();
        match &inner.kind {
            ExprKind::Assign { op, target, value } => {
                self.assign_expr(*op, target, value);
            }
            ExprKind::IncDec { inc, operand, .. } => {
                let (t, ty) = self.lvalue_ex(operand, true);
                let one = self.one_for(&ty);
                self.emit(assign(t.clone(), bin(if *inc { "+" } else { "-" }, t, one)));
            }
            ExprKind::Call { .. } => {
                self.call(inner, false);
            }
            ExprKind::New { .. } => {
                let ty = self.ty(inner);
                let v = self.expr(inner);
                let t = VExpr::Name(self.temp(&ty));
                self.store(&t, &ty, v, &ty);
            }
            _ => {
                // Evaluate for side effects only.
                let _ = self.expr(inner);
            }
        }
    }

    // ------------------------------------------------------------------ calls

    fn finish_call(
        &mut self,
        target: VExpr,
        args: Vec<VExpr>,
        want: bool,
        bare_when_empty: bool,
    ) -> Option<VExpr> {
        if want {
            Some(if args.is_empty() && bare_when_empty {
                target
            } else {
                VExpr::Call(Box::new(target), args)
            })
        } else {
            self.emit(VStmt::Call { target, args });
            None
        }
    }

    /// Lowers the arguments of a call to `f` (parameters from `skip` on), filling in defaults.
    fn call_args(&mut self, f: FnId, skip: usize, args: &[Expr], targs: &[Type]) -> Vec<VExpr> {
        let def = &self.prog.fns[f];
        let env: HashMap<String, Type> = def
            .type_params
            .iter()
            .zip(targs.iter())
            .map(|(p, a)| (p.name.clone(), a.clone()))
            .collect();
        let host = matches!(def.kind, FnKind::Host { .. } | FnKind::Declare { .. });
        let params = &def.params[skip.min(def.params.len())..];
        // By-value arguments are lowered together so that evaluation order is preserved.
        let byval: Vec<&Expr> = args
            .iter()
            .enumerate()
            .filter(|(i, a)| {
                !matches!(a.kind, ExprKind::Undefined)
                    && params.get(*i).is_none_or(|p| p.mode == PassMode::ByVal)
            })
            .map(|(_, a)| a)
            .collect();
        let mut vals = self.operands(&byval).into_iter();
        let mut out = Vec::new();
        for (i, p) in params.iter().enumerate() {
            let pty = p.ty.subst(&env);
            match args
                .get(i)
                .filter(|a| !matches!(a.kind, ExprKind::Undefined))
            {
                Some(a) if p.mode == PassMode::ByRef => out.push(self.lvalue(a).0),
                Some(a) => {
                    let v = vals.next().unwrap_or(VExpr::Missing);
                    let at = self.ty(a);
                    let v = if at == Type::Null {
                        self.null_for(&pty)
                    } else {
                        self.to_union_member(v, &at, &pty)
                    };
                    // W-API strings are passed as pointers (spec §11.2).
                    out.push(if pty == Type::WStr {
                        call("StrPtr", vec![v])
                    } else {
                        v
                    });
                }
                None if host => out.push(VExpr::Missing),
                None => {
                    let v = match &p.default {
                        Some(d) => match literal_default(d, &pty) {
                            Some(text) => lit(text),
                            None => {
                                // Lower the default in the callee's own context.
                                self.pc.frames.push(Frame {
                                    module: def.module,
                                    fn_id: Some(f),
                                    env: env.clone(),
                                    locals: HashMap::new(),
                                    this: None,
                                    inl: None,
                                });
                                let v = self.expr(d);
                                self.pc.frames.pop();
                                v
                            }
                        },
                        None => self.null_for(&pty),
                    };
                    out.push(v);
                }
            }
        }
        // Extra arguments (late-bound / arity errors already reported).
        for a in args.iter().skip(params.len()) {
            let v = self.expr(a);
            out.push(v);
        }
        out
    }

    fn should_inline(&mut self, f: FnId, span: ast::Span) -> bool {
        let prog = self.prog;
        let def = &prog.fns[f];
        let Some(body) = def.decl.as_ref().and_then(|d| d.body.as_ref()) else {
            return false;
        };
        if self.opt == 0 || def.inline == InlineMode::NoInline || !matches!(def.kind, FnKind::Free)
        {
            return false;
        }
        if self.pc.inline_stack.contains(&f) {
            return false;
        }
        let explicit = def.inline == InlineMode::Inline;
        let in_loop = self.pc.loop_depth > 0;
        if !explicit && !(in_loop && self.opt >= 1) && !(self.opt >= 2 && def.cost <= 10) {
            return false;
        }
        let refuse = |lw: &mut Self, why: &str| {
            if explicit || in_loop {
                lw.note(
                    span,
                    false,
                    "inline",
                    format!("`{}` はインライン展開されません: {why}", def.name),
                );
            }
            false
        };
        if !def.inline_blockers.is_empty() || needs_handler(&body.stmts) {
            return refuse(
                self,
                "`try` / `using` を含むため、展開するとエラー状態が呼び出し元に漏れます",
            );
        }
        // Private members of another module are not visible from the expansion site (spec §8.2).
        let cross = def.module != self.pc.emit_module || self.pc.in_class || self.pc.in_closure;
        // `this` of a function of a form (a sheet) is the module: not inside a closure class.
        if self.pc.in_closure && prog.modules[def.module].form.is_some() {
            return refuse(
                self,
                "フォーム・シートのファイルの関数は、アロー関数の中には展開しません",
            );
        }
        if cross {
            let private_fn = def.calls.iter().any(|&g| {
                prog.fns[g].module == def.module
                    && !prog.fns[g].exported
                    && matches!(prog.fns[g].kind, FnKind::Free | FnKind::Declare { .. })
            });
            let private_var = def
                .var_refs
                .iter()
                .any(|&v| !prog.vars[v].exported && !prog.vars[v].declared);
            if private_fn || private_var {
                return refuse(self, "別モジュールの非公開の関数・変数を参照しています");
            }
        }
        if self.pc.stmt_count > 3000 {
            return refuse(
                self,
                "呼び出し元プロシージャが大きくなりすぎるため打ち切りました（VBA の 64KB 制限）",
            );
        }
        if explicit {
            return true;
        }
        let threshold = if in_loop {
            self.cfg.inline_threshold()
        } else {
            10
        };
        if def.cost > threshold {
            return refuse(
                self,
                &format!(
                    "本体のコスト {} が閾値 {threshold} を超えています",
                    def.cost
                ),
            );
        }
        true
    }

    /// Arguments that may be substituted textually into an inlined body (spec §8.3).
    fn is_simple_arg(&self, a: &Expr) -> bool {
        let info = self.info();
        match &a.peel().kind {
            ExprKind::Int { .. }
            | ExprKind::Float { .. }
            | ExprKind::Str(_)
            | ExprKind::Bool(_) => true,
            ExprKind::Unary {
                op: UnaryOp::Neg,
                operand,
            } => matches!(
                operand.peel().kind,
                ExprKind::Int { .. } | ExprKind::Float { .. }
            ),
            ExprKind::Name(_) | ExprKind::This => match info.res.get(&a.peel().id) {
                Some(Res::Local(_)) | Some(Res::EnumMember(..)) => true,
                Some(Res::Global(v)) => self.prog.vars[*v].is_const,
                _ => false,
            },
            ExprKind::Index { object, index } => {
                self.ty(object).dims() == 1
                    && matches!(info.res.get(&object.peel().id), Some(Res::Local(_)))
                    && matches!(&index.peel().kind, ExprKind::Int { .. } | ExprKind::Name(_))
                    && (matches!(index.peel().kind, ExprKind::Int { .. })
                        || matches!(info.res.get(&index.peel().id), Some(Res::Local(_))))
            }
            _ => false,
        }
    }

    /// Expands a call to `f` in place (spec §8).
    fn inline_call(
        &mut self,
        e: &Expr,
        f: FnId,
        targs: &[Type],
        recv: Option<&Expr>,
        args: &[Expr],
        want: bool,
    ) -> Option<VExpr> {
        let prog = self.prog;
        let def = &prog.fns[f];
        let decl = def.decl.clone().unwrap();
        let body = decl.body.as_ref().unwrap();
        self.pc.inl_n += 1;
        let n = self.pc.inl_n;
        let md = &prog.modules[def.module];
        self.emit(VStmt::Comment(format!(
            "@inline {} ({}:{})",
            def.name,
            md.rel_path,
            md.line_of(u32::from(def.name_span.start()))
        )));
        let env: HashMap<String, Type> = def
            .type_params
            .iter()
            .zip(targs.iter())
            .map(|(p, a)| (p.name.clone(), a.clone()))
            .collect();
        let all: Vec<&Expr> = recv.into_iter().chain(args.iter()).collect();
        let mut locals: HashMap<LocalId, Binding> = HashMap::new();
        for (i, p) in def.params.iter().enumerate() {
            let pty = p.ty.subst(&env);
            let reassigned = def.locals[p.local].reassigned;
            let arg = all
                .get(i)
                .copied()
                .filter(|a| !matches!(a.kind, ExprKind::Undefined));
            let binding = match arg {
                Some(a) if p.mode == PassMode::ByRef => Binding::Expr(self.lvalue(a).0),
                Some(a) => {
                    let at = self.ty(a);
                    if (self.is_simple_arg(a) && !reassigned)
                        || (pty.is_array() && matches!(a.peel().kind, ExprKind::Name(_)))
                    {
                        Binding::Expr(self.expr(a))
                    } else {
                        let t = self
                            .declare(&format!("ras_inl{n}_{}", names::valid_ident(&p.name)), &pty);
                        if at == Type::Null {
                            self.store(&VExpr::Name(t.clone()), &pty, lit("Nothing"), &Type::Null);
                        } else {
                            self.expr_into(a, &VExpr::Name(t.clone()), &pty);
                        }
                        Binding::Var(t)
                    }
                }
                None => {
                    let v = match &p.default {
                        Some(d) => match literal_default(d, &pty) {
                            Some(text) => lit(text),
                            None => {
                                self.pc.frames.push(Frame {
                                    module: def.module,
                                    fn_id: Some(f),
                                    env: env.clone(),
                                    locals: HashMap::new(),
                                    this: None,
                                    inl: None,
                                });
                                let v = self.expr(d);
                                self.pc.frames.pop();
                                v
                            }
                        },
                        None => self.null_for(&pty),
                    };
                    if reassigned || !simple(&v) {
                        let t = self
                            .declare(&format!("ras_inl{n}_{}", names::valid_ident(&p.name)), &pty);
                        self.store(&VExpr::Name(t.clone()), &pty, v, &pty);
                        Binding::Var(t)
                    } else {
                        Binding::Expr(v)
                    }
                }
            };
            locals.insert(p.local, binding);
        }
        let ret_ty = def.ret_ty().subst(&env);
        let ret_var =
            (ret_ty != Type::Void).then(|| self.declare(&format!("ras_inl{n}_ret"), &ret_ty));
        let end_label = self.unique(&format!("RAS_InlEnd_{n}"));
        let tries_base = self.pc.tries.len();
        self.pc.frames.push(Frame {
            module: def.module,
            fn_id: Some(f),
            env,
            locals,
            this: None,
            inl: Some(n),
        });
        self.pc.rets.push(RetMode::Inline {
            var: ret_var.clone(),
            ty: ret_ty.clone(),
            label: end_label.clone(),
            used: false,
            tries_base,
        });
        self.pc.inline_stack.push(f);
        self.stmts(&body.stmts, true);
        self.pc.inline_stack.pop();
        self.pc.frames.pop();
        if let Some(RetMode::Inline { used: true, .. }) = self.pc.rets.pop() {
            self.emit(VStmt::Label(end_label));
        }
        self.note(
            e.span,
            true,
            "inline",
            format!("`{}` をインライン展開しました", def.name),
        );
        if want {
            Some(ret_var.map(VExpr::Name).unwrap_or_else(|| lit("Empty")))
        } else {
            None
        }
    }

    pub(crate) fn call(&mut self, e: &Expr, want: bool) -> Option<VExpr> {
        let ExprKind::Call { callee, args, .. } = &e.kind else {
            return None;
        };
        let info = self.info();
        let args = info.call_args(e.id, args);
        let Some(ci) = info.calls.get(&e.id) else {
            return Some(lit("Empty"));
        };
        let targs: Vec<Type> = ci.type_args.iter().map(|t| self.subst(t)).collect();
        let c = callee.peel();
        let recv = match &c.kind {
            ExprKind::Member { object, .. } => Some(&**object),
            _ => None,
        };
        match ci.target.clone() {
            CallTarget::Intrinsic(i) => self.intrinsic(e, i, recv, args, want),
            CallTarget::Fn(f) => {
                if self.should_inline(f, e.span) {
                    return self.inline_call(e, f, &targs, None, args, want);
                }
                let def = &self.prog.fns[f];
                let target = match def.kind {
                    FnKind::Method { .. } | FnKind::Getter { .. } | FnKind::Setter { .. } => {
                        self.static_callee(f)
                    }
                    _ => self.fn_callee(f, &targs),
                };
                let argv = self.call_args(f, 0, args, &targs);
                self.finish_call(target, argv, want, false)
            }
            CallTarget::Ext(f) => {
                let Some(r) = recv else {
                    return Some(lit("Empty"));
                };
                if self.should_inline(f, e.span) {
                    return self.inline_call(e, f, &targs, Some(r), args, want);
                }
                let def = &self.prog.fns[f];
                // The receiver is evaluated first, exactly once (spec §3.2.2).
                let byref = def
                    .params
                    .first()
                    .is_some_and(|p| p.mode == PassMode::ByRef);
                let rv = if byref {
                    self.lvalue(r).0
                } else {
                    let rt = self.ty(r);
                    let v = self.expr(r);
                    let has_effects = args.iter().any(|a| !self.is_simple_arg(a));
                    if has_effects && !rt.is_array() {
                        self.spill(v, &rt)
                    } else {
                        v
                    }
                };
                let target = self.fn_callee(f, &targs);
                let mut argv = vec![rv];
                argv.extend(self.call_args(f, 1, args, &targs));
                self.finish_call(target, argv, want, false)
            }
            CallTarget::Method(f) => {
                let Some(r) = recv else {
                    return Some(lit("Empty"));
                };
                let o = self.expr(r);
                let name = self.fn_name(f, &[]);
                let target = if o.is_lit("Me") {
                    VExpr::Name(name)
                } else {
                    o.member(name)
                };
                let argv = self.call_args(f, 0, args, &[]);
                self.finish_call(target, argv, want, false)
            }
            CallTarget::Host(f) => {
                let def = &self.prog.fns[f];
                // An array in a host result (`getOpenFilename(..., true)`) is 1-based: RAS arrays are 0-based.
                let rebase =
                    matches!(def.ret_ty(), Type::Union(ms) if ms.iter().any(|m| m.is_array()));
                let mut fname = def.vba_name().to_string();
                let outer_qualifier = self.pc.qualifier.clone();
                let target = match (recv, &def.kind) {
                    (Some(r), FnKind::Host { class: Some(_) }) => {
                        let rt = self.ty(r);
                        let mut o = self.expr(r);
                        // `ws.Range(Cells(1, 1), Cells(2, 2))`: inner references follow the receiver (spec §13.4).
                        let info = self.info();
                        let qualifies = fname == "Range"
                            && args.iter().any(|a| {
                                crate::astutil::any_expr_in(a, &mut |x| {
                                    info.auto_qualified.contains(&x.id)
                                })
                            });
                        if qualifies {
                            o = self.spill_object(o, &rt);
                            self.pc.qualifier = Some(o.clone());
                        }
                        o.member(fname)
                    }
                    _ => {
                        if def.ret_ty() == Type::Str && DOLLAR_FUNCS.contains(&fname.as_str()) {
                            fname.push('$');
                        }
                        match (
                            &self.pc.qualifier,
                            self.info().auto_qualified.contains(&c.id),
                        ) {
                            (Some(q), true) => q.clone().member(fname),
                            _ => VExpr::Name(fname),
                        }
                    }
                };
                let argv = self.call_args(f, 0, args, &[]);
                self.pc.qualifier = outer_qualifier;
                let r = self.finish_call(target, argv, want, true);
                if rebase {
                    r.map(|v| call("RAS_Rebase1", vec![v]))
                } else {
                    r
                }
            }
            CallTarget::DefaultMember(_) => {
                let o = self.expr(callee);
                let refs: Vec<&Expr> = args.iter().collect();
                let argv = self.operands(&refs);
                self.finish_call(o, argv, want, false)
            }
            CallTarget::Closure => {
                let o = self.expr(callee);
                let refs: Vec<&Expr> = args.iter().collect();
                let argv = self.operands(&refs);
                self.finish_call(o.member("Apply"), argv, want, false)
            }
            CallTarget::Dynamic | CallTarget::Ctor(_) => {
                let target = match &c.kind {
                    ExprKind::Member { object, name, .. } => {
                        let o = self.expr(object);
                        o.member(
                            self.late_member_name(object, &name.text)
                                .unwrap_or_else(|| name.text.clone()),
                        )
                    }
                    _ => self.expr(callee),
                };
                let refs: Vec<&Expr> = args
                    .iter()
                    .filter(|a| !matches!(a.kind, ExprKind::Undefined))
                    .collect();
                let mut vals = self.operands(&refs).into_iter();
                let argv: Vec<VExpr> = args
                    .iter()
                    .map(|a| {
                        if matches!(a.kind, ExprKind::Undefined) {
                            VExpr::Missing
                        } else {
                            vals.next().unwrap_or(VExpr::Missing)
                        }
                    })
                    .collect();
                self.finish_call(target, argv, want, true)
            }
        }
    }

    fn new_expr(&mut self, e: &Expr, ty: &ast::TypeRef, args: &[Expr]) -> VExpr {
        if let ast::TypeRef::Named { path, .. } = ty {
            if path.len() == 1 {
                match path[0].text.as_str() {
                    "Error" => {
                        let refs: Vec<&Expr> = args.iter().collect();
                        let argv = self.operands(&refs);
                        return call("RAS_MakeError", argv);
                    }
                    "Map" if matches!(self.ty(e), Type::Map(..)) => {
                        return call("CreateObject", vec![lit("\"Scripting.Dictionary\"")])
                    }
                    "Array" => return self.array_value(e),
                    _ => {}
                }
            }
        }
        match self.info().calls.get(&e.id).map(|c| c.target.clone()) {
            Some(CallTarget::Ctor(c)) => {
                let def = &self.prog.classes[c];
                let targs: Vec<Type> = match self.ty(e) {
                    Type::Class(_, a) => a,
                    _ => Vec::new(),
                };
                let cname = self.class_name(c, &targs);
                if def.kind == ClassKind::Host {
                    // Without a reference to its library: by its ProgID (spec §15.3).
                    if let Some(id) = crate::model::prog_id(def.namespace.as_deref(), &def.name) {
                        return call("CreateObject", vec![lit(format!("\"{id}\""))]);
                    }
                    return lit(format!("New {cname}"));
                }
                let factory = VExpr::Name(self.prog.modules[def.module].vba_name.clone())
                    .member(format!("New_{cname}"));
                let argv = match def.ctor {
                    Some(ctor) => self.call_args(ctor, 0, args, &[]),
                    None => Vec::new(),
                };
                VExpr::Call(Box::new(factory), argv)
            }
            _ => lit("Nothing"),
        }
    }

    // ------------------------------------------------------------------ intrinsics

    fn arg_or(&mut self, args: &[Expr], i: usize, default: VExpr) -> VExpr {
        match args.get(i) {
            Some(a) => self.expr(a),
            None => default,
        }
    }

    pub(crate) fn intrinsic(
        &mut self,
        e: &Expr,
        intr: Intr,
        recv: Option<&Expr>,
        args: &[Expr],
        want: bool,
    ) -> Option<VExpr> {
        use Intr::*;
        let recv_ty = recv.map(|r| self.ty(r)).unwrap_or(Type::Error);
        // Array-producing and callback-taking members are expanded as loops.
        if matches!(
            intr,
            ArrMap
                | ArrFilter
                | ArrForEach
                | ArrReduce
                | ArrSome
                | ArrEvery
                | ArrFind
                | ArrFindIndex
                | ArrSlice
                | ArrConcat
                | MapKeys
                | MapValues
                | CollToArray
        ) {
            let r = recv?;
            let ty = self.ty(e);
            if ty.is_array() {
                let t = VExpr::Name(self.temp(&ty));
                self.array_intrinsic(e, intr, r, args, &t, &ty);
                return Some(t);
            }
            return self.scalar_fusion(e, intr, r, args, want);
        }
        let stmt = |lw: &mut Self, name: &str, argv: Vec<VExpr>| -> Option<VExpr> {
            lw.emit(call_stmt(name, argv));
            None
        };
        if intr == EventOn {
            return self.event_on(e, recv?, args, want);
        }
        let v = match intr {
            EventOn => unreachable!(),
            ArrLength => {
                let a = self.expr(recv?);
                match self.pc.push_caps.get(&a.print()) {
                    Some((len, _)) => name(len.clone()),
                    None => self.len_of(a, &recv_ty),
                }
            }
            ArrRows => call("RAS_Len", vec![self.expr(recv?)]),
            ArrCols => call("RAS_Len2", vec![self.expr(recv?)]),
            StrLength => call("Len", vec![self.expr(recv?)]),
            MapSize => self.expr(recv?).member("Count"),
            ArrPush => {
                let r = recv?;
                let (arr, _) = self.lvalue(r);
                let elem_ty = recv_ty.elem().cloned().unwrap_or(Type::Any);
                let vt = self.ty(&args[0]);
                let is_struct_lit = matches!(vt, Type::Struct(_))
                    && matches!(
                        args[0].peel().kind,
                        ExprKind::Object(_) | ExprKind::Array(_)
                    );
                if let (true, Some((len, cap))) =
                    (is_struct_lit, self.pc.push_caps.get(&arr.print()).cloned())
                {
                    // The literal is built in place: no temporary and no copy of the struct.
                    self.emit(if_then(
                        bin(">=", name(len.clone()), name(cap.clone())),
                        vec![
                            assign(
                                name(cap.clone()),
                                bin("+", bin("*", name(cap.clone()), long(2)), long(8)),
                            ),
                            VStmt::ReDim {
                                preserve: true,
                                target: arr.clone(),
                                dims: vec![(long(0), name(cap).plus(-1))],
                            },
                        ],
                    ));
                    self.struct_lit_into(args[0].peel(), &arr.index(vec![name(len.clone())]));
                    self.emit(assign(name(len.clone()), name(len).plus(1)));
                    return None;
                }
                let v = self.expr(&args[0]);
                if let Some((len, cap)) = self.pc.push_caps.get(&arr.print()).cloned() {
                    // Capacity doubling instead of `ReDim Preserve` per element (spec §13.3).
                    self.emit(if_then(
                        bin(">=", name(len.clone()), name(cap.clone())),
                        vec![
                            assign(
                                name(cap.clone()),
                                bin("+", bin("*", name(cap.clone()), long(2)), long(8)),
                            ),
                            VStmt::ReDim {
                                preserve: true,
                                target: arr.clone(),
                                dims: vec![(long(0), name(cap).plus(-1))],
                            },
                        ],
                    ));
                    self.store(&arr.index(vec![name(len.clone())]), &elem_ty, v, &vt);
                    self.emit(assign(name(len.clone()), name(len).plus(1)));
                } else {
                    let v = self.spill(v, &vt);
                    let n = VExpr::Name(self.temp(&Type::Int));
                    let len = self.len_of(arr.clone(), &recv_ty);
                    self.emit(assign(n.clone(), len));
                    self.emit(VStmt::ReDim {
                        preserve: true,
                        target: arr.clone(),
                        dims: vec![(long(0), n.clone())],
                    });
                    self.store(&arr.index(vec![n]), &elem_ty, v, &vt);
                    if self.pc.loop_depth > 0 {
                        self.note(e.span, false, "push", "このループ内の `push` は 1 要素ごとに `ReDim Preserve` します（配列がループ内で他の用途にも使われているため、容量倍増方式にできません）");
                    }
                }
                return None;
            }
            ArrPop => {
                let (arr, _) = self.lvalue(recv?);
                let elem_ty = recv_ty.elem().cloned().unwrap_or(Type::Any);
                let n = VExpr::Name(self.temp(&Type::Int));
                let out = VExpr::Name(self.temp(&elem_ty));
                let len = self.len_of(arr.clone(), &recv_ty);
                self.emit(assign(n.clone(), len.plus(-1)));
                self.reset_var(&out, &elem_ty);
                let (mut body, _) = self.capture(|lw| {
                    lw.store(&out, &elem_ty, arr.clone().index(vec![n.clone()]), &elem_ty)
                });
                let (clear, _) = self.capture(|lw| lw.clear_array(&arr, &recv_ty));
                body.push(if_else(
                    bin("=", n.clone(), long(0)),
                    clear,
                    vec![VStmt::ReDim {
                        preserve: true,
                        target: arr.clone(),
                        dims: vec![(long(0), n.clone().plus(-1))],
                    }],
                ));
                self.emit(if_then(bin(">=", n, long(0)), body));
                return want.then_some(out);
            }
            ArrIndexOf | ArrIncludes => {
                let mut all = vec![recv?];
                all.extend(args.iter());
                let mut ops = self.operands(&all);
                if intr == ArrIncludes {
                    bin(">=", call("RAS_IndexOf", ops), long(0))
                } else {
                    if ops.len() == 2 {
                        ops.push(long(0));
                    }
                    call("RAS_IndexOf", ops)
                }
            }
            ArrJoin => {
                let a = self.expr(recv?);
                let sep = self.arg_or(args, 0, lit("\",\""));
                call("RAS_Join", vec![a, sep])
            }
            ArrSort => {
                let (arr, _) = self.lvalue(recv?);
                return match args.first() {
                    Some(cmp) => {
                        let c = self.expr(cmp);
                        stmt(self, "RAS_SortBy", vec![arr, c])
                    }
                    None => stmt(self, "RAS_Sort", vec![arr]),
                };
            }
            ArrReverse => {
                let (arr, _) = self.lvalue(recv?);
                return stmt(self, "RAS_Reverse", vec![arr]);
            }
            ArrClear => {
                let (arr, _) = self.lvalue(recv?);
                self.clear_array(&arr, &recv_ty);
                return None;
            }
            StrTrim => call("Trim$", vec![self.expr(recv?)]),
            StrTrimStart => call("LTrim$", vec![self.expr(recv?)]),
            StrTrimEnd => call("RTrim$", vec![self.expr(recv?)]),
            StrUpper => call("UCase$", vec![self.expr(recv?)]),
            StrLower => call("LCase$", vec![self.expr(recv?)]),
            StrToString => self.expr(recv?),
            StrSubstring => {
                // 0-based [start, end) → Mid$(s, start + 1, end - start) (spec §13.3)
                let mut all = vec![recv?];
                all.extend(args.iter());
                let ops = self.operands(&all);
                let start = self.spill(ops[1].clone(), &Type::Int);
                match ops.get(2) {
                    Some(end) => {
                        let len = match start.as_long() {
                            Some(k) => end.clone().plus(-k),
                            None => bin("-", end.clone(), start.clone()),
                        };
                        call("Mid$", vec![ops[0].clone(), start.plus(1), len])
                    }
                    None => call("Mid$", vec![ops[0].clone(), start.plus(1)]),
                }
            }
            StrSubstr => {
                let mut all = vec![recv?];
                all.extend(args.iter());
                let ops = self.operands(&all);
                let mut a = vec![ops[0].clone(), ops[1].clone().plus(1)];
                a.extend(ops.get(2).cloned());
                call("Mid$", a)
            }
            StrSlice => {
                let mut all = vec![recv?];
                all.extend(args.iter());
                call("RAS_StrSlice", self.operands(&all))
            }
            StrIndexOf => {
                let mut all = vec![recv?];
                all.extend(args.iter());
                let ops = self.operands(&all);
                let from = ops.get(2).cloned().map(|f| f.plus(1)).unwrap_or(long(1));
                call(
                    "InStr",
                    vec![
                        from,
                        ops[0].clone(),
                        ops[1].clone(),
                        name("vbBinaryCompare"),
                    ],
                )
                .plus(-1)
            }
            StrLastIndexOf => {
                let ops = self.operands(&[recv?, &args[0]]);
                call(
                    "InStrRev",
                    vec![
                        ops[0].clone(),
                        ops[1].clone(),
                        lit("-1"),
                        name("vbBinaryCompare"),
                    ],
                )
                .plus(-1)
            }
            StrIncludes => {
                let ops = self.operands(&[recv?, &args[0]]);
                bin(
                    ">",
                    call(
                        "InStr",
                        vec![
                            long(1),
                            ops[0].clone(),
                            ops[1].clone(),
                            name("vbBinaryCompare"),
                        ],
                    ),
                    long(0),
                )
            }
            StrStartsWith | StrEndsWith => {
                let ops = self.operands(&[recv?, &args[0]]);
                let p = self.spill(ops[1].clone(), &Type::Str);
                let f = if intr == StrStartsWith {
                    "Left$"
                } else {
                    "Right$"
                };
                bin(
                    "=",
                    call(f, vec![ops[0].clone(), call("Len", vec![p.clone()])]),
                    p,
                )
            }
            StrPadStart | StrPadEnd => {
                let mut all = vec![recv?];
                all.extend(args.iter());
                let mut ops = self.operands(&all);
                if ops.len() == 2 {
                    ops.push(lit("\" \""));
                }
                call(
                    if intr == StrPadStart {
                        "RAS_PadStart"
                    } else {
                        "RAS_PadEnd"
                    },
                    ops,
                )
            }
            StrRepeat => call("RAS_Repeat", self.operands(&[recv?, &args[0]])),
            StrReplace | StrReplaceAll => {
                let mut ops = self.operands(&[recv?, &args[0], &args[1]]);
                ops.extend([
                    long(1),
                    if intr == StrReplace {
                        long(1)
                    } else {
                        lit("-1")
                    },
                    name("vbBinaryCompare"),
                ]);
                call("Replace", ops)
            }
            StrSplit => {
                let mut ops = self.operands(&[recv?, &args[0]]);
                ops.extend([lit("-1"), name("vbBinaryCompare")]);
                call("Split", ops)
            }
            StrCharAt => {
                let ops = self.operands(&[recv?, &args[0]]);
                call(
                    "Mid$",
                    vec![ops[0].clone(), ops[1].clone().plus(1), long(1)],
                )
            }
            StrCharCodeAt => {
                let ops = self.operands(&[recv?, &args[0]]);
                // AscW returns a signed Integer.
                bin(
                    "And",
                    call(
                        "AscW",
                        vec![call(
                            "Mid$",
                            vec![ops[0].clone(), ops[1].clone().plus(1), long(1)],
                        )],
                    ),
                    lit("&HFFFF&"),
                )
            }
            NumToString | BoolToString | DateToString => {
                let v = self.expr(recv?);
                self.to_string(v, &recv_ty)
            }
            NumToFixed => {
                let v = self.expr(recv?);
                let d = self.arg_or(args, 0, long(0));
                call("RAS_ToFixed", vec![v, d])
            }
            MapGet => {
                let ops = self.operands(&[recv?, &args[0]]);
                let vty = match &recv_ty {
                    Type::Map(_, v) => (**v).clone(),
                    _ => Type::Any,
                };
                self.map_get(ops[0].clone(), ops[1].clone(), &vty)
            }
            MapSet => {
                let ops = self.operands(&[recv?, &args[0], &args[1]]);
                let vty = match &recv_ty {
                    Type::Map(_, v) => (**v).clone(),
                    _ => Type::Any,
                };
                let at = self.ty(&args[1]);
                self.store(
                    &ops[0].clone().index(vec![ops[1].clone()]),
                    &vty,
                    ops[2].clone(),
                    &at,
                );
                return None;
            }
            MapHas => {
                let ops = self.operands(&[recv?, &args[0]]);
                ops[0].clone().member("Exists").index(vec![ops[1].clone()])
            }
            MapDelete => {
                let ops = self.operands(&[recv?, &args[0]]);
                let kt = self.ty(&args[0]);
                let k = self.spill(ops[1].clone(), &kt);
                self.emit(if_then(
                    ops[0].clone().member("Exists").index(vec![k.clone()]),
                    vec![VStmt::Call {
                        target: ops[0].clone().member("Remove"),
                        args: vec![k],
                    }],
                ));
                return None;
            }
            MapClear => {
                let d = self.expr(recv?);
                self.emit(VStmt::Call {
                    target: d.member("RemoveAll"),
                    args: vec![],
                });
                return None;
            }
            MathAbs => call("Abs", vec![self.expr(&args[0])]),
            MathFloor => call("Int", vec![self.expr(&args[0])]),
            MathCeil => {
                let v = self.expr(&args[0]);
                VExpr::Neg(Box::new(call("Int", vec![VExpr::Neg(Box::new(v))])))
            }
            MathTrunc => call("Fix", vec![self.expr(&args[0])]),
            // VBA's Round is banker's rounding; Math.round rounds half away from zero (spec §13.2).
            MathRound => {
                let refs: Vec<&Expr> = args.iter().collect();
                call("RAS_Round", self.operands(&refs))
            }
            MathBankersRound => {
                let refs: Vec<&Expr> = args.iter().collect();
                call("Round", self.operands(&refs))
            }
            MathSqrt => call("Sqr", vec![self.expr(&args[0])]),
            MathPow => {
                let ops = self.operands(&[&args[0], &args[1]]);
                bin("^", ops[0].clone(), ops[1].clone())
            }
            MathSin => call("Sin", vec![self.expr(&args[0])]),
            MathCos => call("Cos", vec![self.expr(&args[0])]),
            MathTan => call("Tan", vec![self.expr(&args[0])]),
            MathAtan => call("Atn", vec![self.expr(&args[0])]),
            MathLog => call("Log", vec![self.expr(&args[0])]),
            MathExp => call("Exp", vec![self.expr(&args[0])]),
            MathRandom => call("Rnd", vec![]),
            MathSign => call("Sgn", vec![self.expr(&args[0])]),
            MathPi => lit("3.14159265358979#"),
            MathMin | MathMax => {
                let ty = self.ty(e);
                let t = VExpr::Name(self.temp(&ty));
                let first = self.expr(&args[0]);
                self.emit(assign(t.clone(), first));
                for a in &args[1..] {
                    let at = self.ty(a);
                    let v = self.expr(a);
                    let v = self.spill(v, &at);
                    let cmp = bin(
                        if intr == MathMin { "<" } else { ">" },
                        v.clone(),
                        t.clone(),
                    );
                    self.emit(if_then(cmp, vec![assign(t.clone(), v)]));
                }
                t
            }
            ParseInt => call(
                "CLng",
                vec![call("Fix", vec![call("Val", vec![self.expr(&args[0])])])],
            ),
            ParseFloat => call("Val", vec![self.expr(&args[0])]),
            ToStr => {
                let at = self.ty(&args[0]);
                let v = self.expr(&args[0]);
                self.to_string(v, &at)
            }
            ConsoleLog => {
                let mut acc: Option<VExpr> = None;
                for a in args {
                    let at = self.ty(a);
                    let v = self.expr(a);
                    let s = self.to_string(v, &at);
                    acc = Some(match acc {
                        None => s,
                        Some(p) => bin("&", bin("&", p, lit("\" \"")), s),
                    });
                }
                self.emit(VStmt::Call {
                    target: name("Debug").member("Print"),
                    args: acc.into_iter().collect(),
                });
                return None;
            }
            IsNullish => lit("False"),
            FastMode => {
                self.diag(Severity::Warning, self.src_module(), e.span, "W0630", "`fastMode()` は `using _ = fastMode();` の形で使ってください（このままでは何も行われません）");
                long(fast_mode_mask(e))
            }
            SuspendScreen => return stmt(self, "RAS_Suspend", vec![long(1)]),
            ResumeScreen => return stmt(self, "RAS_Resume", vec![long(1)]),
            SuspendEvents => return stmt(self, "RAS_Suspend", vec![long(2)]),
            ResumeEvents => return stmt(self, "RAS_Resume", vec![long(2)]),
            SuspendCalc => return stmt(self, "RAS_Suspend", vec![long(4)]),
            ResumeCalc => return stmt(self, "RAS_Resume", vec![long(4)]),
            SuspendAlerts => return stmt(self, "RAS_Suspend", vec![long(8)]),
            ResumeAlerts => return stmt(self, "RAS_Resume", vec![long(8)]),
            SuspendAll => return stmt(self, "RAS_SuspendAll", vec![]),
            ResumeAll => return stmt(self, "RAS_ResumeAll", vec![]),
            Status => {
                let v = self.arg_or(args, 0, lit("vbNullString"));
                return stmt(self, "RAS_Status", vec![v]);
            }
            Sleep => {
                let v = self.expr(&args[0]);
                return stmt(self, "RAS_Sleep", vec![v]);
            }
            WaitUntil => {
                let f = self.expr(&args[0]);
                let timeout = self.arg_or(args, 1, lit("-1&"));
                let c = call("RAS_WaitUntil", vec![f, timeout]);
                if !want {
                    self.emit(VStmt::Call {
                        target: name("RAS_WaitUntil"),
                        args: match c {
                            VExpr::Call(_, a) => a,
                            _ => vec![],
                        },
                    });
                    return None;
                }
                c
            }
            YieldNow => return stmt(self, "RAS_YieldNow", vec![]),
            PlatformIs64 => call("RAS_Is64", vec![]),
            PlatformIsVba7 => call("RAS_IsVBA7", vec![]),
            RangeToArray => call("RAS_Read2D", vec![self.expr(recv?)]),
            SheetUsedTable => call("RAS_Read2D", vec![self.expr(recv?).member("UsedRange")]),
            RangeToColumn => call("RAS_ToColumn", vec![self.expr(recv?)]),
            RangeToRow => call("RAS_ToRow", vec![self.expr(recv?)]),
            RangeFromArray => {
                let ops = self.operands(&[recv?, &args[0]]);
                return stmt(self, "RAS_FromArray", ops);
            }
            ExcelMatch => {
                let refs: Vec<&Expr> = args.iter().collect();
                let mut ops = self.operands(&refs);
                if ops.len() == 2 {
                    ops.push(long(0));
                }
                call("RAS_Match", ops)
            }
            ArrMap | ArrFilter | ArrForEach | ArrReduce | ArrSome | ArrEvery | ArrFind
            | ArrFindIndex | ArrSlice | ArrConcat | MapKeys | MapValues | CollToArray => {
                unreachable!()
            }
        };
        if want {
            Some(v)
        } else {
            // Value not needed: nothing to emit for pure intrinsics.
            None
        }
    }
}

/// VBA field name of a tuple element (`t[1]` → `item1`).
pub(crate) fn tuple_field(index: &Expr) -> String {
    match &index.peel().kind {
        ExprKind::Int { value, .. } => format!("item{value}"),
        _ => "item0".into(),
    }
}

/// True when an index inside the assignment target `v` (`a(i).f`) mentions `name`.
fn target_index_mentions(v: &VExpr, name: &str) -> bool {
    match v {
        VExpr::Member(o, _) => target_index_mentions(o, name),
        VExpr::Call(f, args) => {
            target_index_mentions(f, name) || args.iter().any(|a| mentions(a, name))
        }
        _ => false,
    }
}

/// True when the printed form of `v` mentions the identifier `name`.
pub(crate) fn mentions(v: &VExpr, name: &str) -> bool {
    match v {
        VExpr::Name(n) => n == name,
        VExpr::Lit(_) | VExpr::Missing => false,
        VExpr::Member(o, _) => mentions(o, name),
        VExpr::Call(f, args) => mentions(f, name) || args.iter().any(|a| mentions(a, name)),
        VExpr::Not(x) | VExpr::Neg(x) => mentions(x, name),
        VExpr::Bin(_, l, r) => mentions(l, name) || mentions(r, name),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn doubles() {
        assert_eq!(double_lit(1.5), "1.5#");
        assert_eq!(double_lit(1000.0), "1000#");
        assert_eq!(double_lit(0.1), "0.1#");
        assert_eq!(double_lit(1e30), "1.0E+30");
        assert_eq!(double_lit(2.5e-10), "2.5E-10");
    }

    #[test]
    fn decimals_do_not_depend_on_locale() {
        assert_eq!(decimal_lit("1.5").print(), "(CDec(\"15\") / CDec(\"10\"))");
        assert_eq!(decimal_lit("12").print(), "CDec(\"12\")");
        assert_eq!(decimal_lit("0.05").print(), "(CDec(\"5\") / CDec(\"100\"))");
    }

    #[test]
    fn int_min() {
        assert_eq!(
            int_lit(-2147483648, &Type::Int, false).print(),
            "(-2147483647& - 1&)"
        );
        assert_eq!(int_lit(255, &Type::Int, true).print(), "&HFF&");
        assert_eq!(int_lit(123, &Type::Long, false).print(), "123^");
    }
}
