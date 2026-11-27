//! Expression code generation.

use ras_syntax::ast::{self, BinaryOp, Expr, ExprKind, PassMode, UnaryOp};

use super::{js_ident, js_str, Gen};
use crate::intrinsics::Intr;
use crate::model::*;

/// Something that can be assigned to.
pub(crate) enum Place {
    /// A JavaScript l-value (variable, field, array element).
    Native(String),
    /// Property of a host (COM) or late-bound object.
    Host {
        obj: String,
        name: String,
    },
    /// Default member of a host object (`coll(i) = v`).
    Item {
        obj: String,
        index: String,
    },
    MapItem {
        map: String,
        key: String,
    },
}

/// Types whose values are only known at run time (VBA `Variant`).
fn is_dyn(t: &Type) -> bool {
    matches!(
        t,
        Type::Any | Type::CellValue | Type::Error | Type::Param(_)
    ) || t.is_value_union()
}

/// The `TypeName` of the numbers a value of type `t` can hold, when there is only one.
fn numeric_type_name(t: &Type) -> Option<&'static str> {
    let members = match t {
        Type::Union(ms) => ms.as_slice(),
        other => std::slice::from_ref(other),
    };
    let mut names: Vec<&'static str> = Vec::new();
    for m in members {
        let n = match m {
            Type::Byte => "Byte",
            Type::Short => "Integer",
            Type::Int | Type::Enum(_) => "Long",
            Type::Long => "LongLong",
            Type::Float => "Single",
            Type::Double => "Double",
            Type::Currency => "Currency",
            Type::Decimal => "Decimal",
            _ => continue,
        };
        if !names.contains(&n) {
            names.push(n);
        }
    }
    (names.len() == 1).then(|| names[0])
}

fn int_like(t: &Type) -> bool {
    t.is_integer() || matches!(t, Type::Enum(_) | Type::Ptr)
}

/// An expression that names a storage location (its array / struct value must be copied on assignment).
fn is_place(e: &Expr) -> bool {
    matches!(
        e.peel().kind,
        ExprKind::Name(_) | ExprKind::This | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

impl<'p> Gen<'p> {
    fn is_host_class(&self, t: &Type) -> bool {
        matches!(t, Type::Class(c, _) if self.prog.classes[*c].kind == ClassKind::Host)
    }

    /// Members of these are reached through `$rt.get` / `$rt.call` (late binding).
    fn is_late_bound(&self, t: &Type) -> bool {
        match t {
            // The members of a union are found by name at run time unless all are RAS classes.
            Type::Union(ms) => ms.iter().any(|m| !matches!(m, Type::Class(c, _) if self.prog.classes[*c].kind != ClassKind::Host)),
            _ => {
                matches!(t, Type::Any | Type::CellValue | Type::Object | Type::Param(_))
                    || self.is_host_class(t)
            }
        }
    }

    pub(crate) fn to_string(&mut self, v: String, ty: &Type) -> String {
        match ty {
            Type::Str => v,
            Type::Byte
            | Type::Short
            | Type::Int
            | Type::Long
            | Type::Enum(_)
            | Type::Ptr
            | Type::Bool => format!("String({v})"),
            Type::Float | Type::Double | Type::Currency | Type::Decimal => {
                format!("$rt.numStr({v})")
            }
            _ => format!("$rt.str({v})"),
        }
    }

    /// Implicit conversion when a value of type `from` is stored into a slot of type `to`.
    /// Only run-time typed values need work: VBA converts a `Variant` on assignment.
    pub(crate) fn coerce(&mut self, v: String, from: &Type, to: &Type) -> String {
        if from == to || !is_dyn(from) {
            return v;
        }
        match to {
            Type::Byte => format!("$rt.cint({v}, 0, 255)"),
            Type::Short => format!("$rt.cint({v}, -32768, 32767)"),
            Type::Int | Type::Enum(_) => format!("$rt.cint({v})"),
            Type::Long | Type::Ptr => format!("$rt.clng({v})"),
            Type::Float | Type::Double | Type::Currency | Type::Decimal => format!("$rt.cnum({v})"),
            Type::Str => format!("$rt.cstr({v})"),
            Type::Bool => format!("$rt.cbool({v})"),
            Type::Date => format!("$rt.cdate({v})"),
            _ => v,
        }
    }

    /// Explicit conversion (`x as T`).
    fn convert(&mut self, v: String, from: &Type, to: &Type) -> String {
        use Type::*;
        if from == to || matches!(to, Any | CellValue | Error) {
            return v;
        }
        let exact_int = from.is_integer() || matches!(from, Enum(_));
        match to {
            // RAS truncates toward zero, like C# (spec §2.3).
            Int | Enum(_) if *from == Bool => format!("({v} ? 1 : 0)"),
            Int | Enum(_) if matches!(from, Byte | Short | Int | Enum(_)) => v,
            Int | Enum(_) => format!("$rt.toInt({v})"),
            Short => format!("$rt.toInt({v}, -32768, 32767)"),
            Byte => format!("$rt.toInt({v}, 0, 255)"),
            Long | Ptr if exact_int => v,
            Long | Ptr => format!("$rt.toLong({v})"),
            Double | Currency | Decimal if from.is_numeric() && *from != Bool => v,
            Float if exact_int => v,
            Float => format!("Math.fround($rt.cnum({v}))"),
            Double | Currency | Decimal => format!("$rt.cnum({v})"),
            Date => format!("$rt.cdate({v})"),
            Str => self.to_string(v, from),
            Bool if *from == Bool => v,
            Bool => format!("$rt.cbool({v})"),
            _ => v,
        }
    }

    /// Lowers `e` as the value stored into a slot of type `to`: converts run-time typed values
    /// and copies arrays / structs (they are value types in VBA).
    pub(crate) fn coerce_store(&mut self, e: &Expr, to: &Type) -> String {
        let from = self.ty(e);
        if from == Type::Null {
            return if to.is_object() || is_dyn(to) {
                "null".into()
            } else {
                self.zero(to)
            };
        }
        let v = self.expr(e);
        let v = self.coerce(v, &from, to);
        if (from.is_array() || matches!(from, Type::Struct(_))) && is_place(e) {
            format!("$rt.clone({v})")
        } else {
            v
        }
    }

    // ------------------------------------------------------------------ expressions

    pub(crate) fn expr(&mut self, e: &Expr) -> String {
        match &e.kind {
            // Ranges and `^n` are lowered by the index that holds them (rejected elsewhere).
            ExprKind::Error | ExprKind::Range { .. } => "undefined".into(),
            ExprKind::FromEnd(x) => self.expr(x),
            ExprKind::Super => "super".into(),
            ExprKind::Object(props) => match self.ty(e) {
                Type::Struct(sid) => {
                    let fields = self.prog.structs[sid].fields.clone();
                    let vals: Vec<String> = props
                        .iter()
                        .filter_map(|p| {
                            fields
                                .iter()
                                .find(|f| f.0 == p.key.text)
                                .map(|f| (p, f.1.clone()))
                        })
                        .map(|(p, fty)| {
                            format!("{}: {}", p.key.text, self.coerce_store(&p.value, &fty))
                        })
                        .collect();
                    format!(
                        "Object.assign(new {}(), {{ {} }})",
                        self.struct_ref(sid),
                        vals.join(", ")
                    )
                }
                _ => "null".into(),
            },
            ExprKind::Array(items) if matches!(self.ty(e), Type::Struct(_)) => {
                let Type::Struct(sid) = self.ty(e) else {
                    unreachable!()
                };
                let fields = self.prog.structs[sid].fields.clone();
                let vals: Vec<String> = items
                    .iter()
                    .zip(fields.iter())
                    .map(|(it, f)| format!("{}: {}", f.0, self.coerce_store(it, &f.1)))
                    .collect();
                format!(
                    "Object.assign(new {}(), {{ {} }})",
                    self.struct_ref(sid),
                    vals.join(", ")
                )
            }
            ExprKind::Int { value, .. } => value.to_string(),
            ExprKind::Float { text, .. } => match text.parse::<f64>() {
                Ok(f) if f.is_finite() => format!("{f:?}"),
                _ => "0".into(),
            },
            ExprKind::Str(s) => js_str(s),
            ExprKind::Template(parts) => {
                let mut out: Vec<String> = Vec::new();
                for p in parts {
                    out.push(match p {
                        ast::TemplatePart::Str(s) => js_str(s),
                        ast::TemplatePart::Expr(x) => {
                            let ty = self.ty(x);
                            let v = self.expr(x);
                            self.to_string(v, &ty)
                        }
                    });
                }
                match out.len() {
                    0 => "\"\"".into(),
                    1 if matches!(parts[0], ast::TemplatePart::Str(_)) => out.remove(0),
                    1 => format!("(\"\" + {})", out[0]),
                    _ => format!("({})", out.join(" + ")),
                }
            }
            ExprKind::Bool(b) => b.to_string(),
            ExprKind::Null => "null".into(),
            ExprKind::Undefined => "undefined".into(),
            ExprKind::Name(id) => self.name_expr(e, &id.text),
            ExprKind::This => match self.info().res.get(&e.id) {
                Some(Res::Local(l)) => self.local_js(*l),
                // In a form's code-behind, `this` is the form (its only instance, spec §14.10.7).
                _ if self.prog.modules[self.module].is_form() => "$me".into(),
                // In the file of a sheet (of `ThisWorkbook`), `this` is that sheet (workbook).
                _ if self.prog.modules[self.module].document().is_some() => format!(
                    "$rt.global({})",
                    js_str(&self.prog.modules[self.module].vba_name)
                ),
                _ => "this".into(),
            },
            ExprKind::Paren(inner) => {
                let v = self.expr(inner);
                if v.starts_with('(') && v.ends_with(')') {
                    v
                } else {
                    format!("({v})")
                }
            }
            ExprKind::NonNull(inner) => self.expr(inner),
            ExprKind::Array(items) => {
                let elem = self.ty(e).elem_of();
                let vals: Vec<String> = items.iter().map(|i| self.coerce_store(i, &elem)).collect();
                format!("[{}]", vals.join(", "))
            }
            ExprKind::Arrow(a) => self.arrow(e, a),
            ExprKind::Unary { op, operand } => {
                if let (UnaryOp::Neg, ExprKind::Int { value, .. }) = (op, &operand.peel().kind) {
                    return format!("(-{value})");
                }
                let v = self.expr(operand);
                match op {
                    UnaryOp::Not => format!("(!{v})"),
                    UnaryOp::Neg => format!("(-{v})"),
                    UnaryOp::Plus => v,
                    UnaryOp::BitNot => format!("(~{v})"),
                    // A JavaScript number does not tell `Long` from `Double`: the static type does when
                    // it has one numeric member (`string | double`).
                    UnaryOp::TypeOf => match numeric_type_name(&self.ty(operand)) {
                        Some(n) => format!("$rt.typeName({v}, {})", js_str(n)),
                        None => format!("$rt.typeName({v})"),
                    },
                }
            }
            ExprKind::IncDec { .. } | ExprKind::Assign { .. } => {
                format!("({})", self.effect_ex(e, true))
            }
            ExprKind::Binary { op, lhs, rhs } => self.binary(e, *op, lhs, rhs),
            ExprKind::Ternary { cond, then, els } => {
                let ty = self.ty(e);
                let c = self.expr(cond);
                let t = self.coerce_store(then, &ty);
                let f = self.coerce_store(els, &ty);
                format!("({c} ? {t} : {f})")
            }
            ExprKind::Call { .. } => self.call(e),
            ExprKind::New { ty, args } => self.new_expr(e, ty, self.info().call_args(e.id, args)),
            // Reordered into the call's arguments by the checker (spec §3.3).
            ExprKind::NamedArg { value, .. } => self.expr(value),
            ExprKind::Member {
                object,
                name,
                optional,
            } => self.member_expr(e, object, &name.text, *optional),
            ExprKind::Index { object, index } => self.index_expr(object, index),
            ExprKind::As { expr, .. } => {
                let (from, to) = (self.ty(expr), self.ty(e));
                if from == Type::Null {
                    return if to.is_object() || is_dyn(&to) {
                        "null".into()
                    } else {
                        self.zero(&to)
                    };
                }
                let v = self.expr(expr);
                self.convert(v, &from, &to)
            }
            ExprKind::InstanceOf { expr, ty } => {
                let v = self.expr(expr);
                match self.resolve_class(ty) {
                    Some(c) => {
                        let d = &self.prog.classes[c];
                        match d.kind {
                            ClassKind::Class => format!("({v} instanceof {})", self.class_ref(c)),
                            ClassKind::Interface => {
                                format!("$rt.implements({v}, {})", js_str(&d.name))
                            }
                            ClassKind::Host => format!("$rt.typeIs({v}, {})", js_str(&d.name)),
                        }
                    }
                    None => "false".into(),
                }
            }
        }
    }

    /// Assignments and `++` / `--`, without the surrounding parentheses a value context needs.
    pub(crate) fn effect(&mut self, e: &Expr) -> String {
        self.effect_ex(e, false)
    }

    /// `want`: the value of the expression is used (`y = x++`).
    fn effect_ex(&mut self, e: &Expr, want: bool) -> String {
        match &e.kind {
            ExprKind::Assign { op, target, value } => self.assign(*op, target, value),
            ExprKind::IncDec {
                inc,
                prefix,
                operand,
            } => {
                let ty = self.ty(operand);
                let place = self.place(operand);
                let op = if *inc { "+" } else { "-" };
                match &place {
                    Place::Native(p) if !ty.is_integer() => {
                        let sym = if *inc { "++" } else { "--" };
                        if *prefix {
                            format!("{sym}{p}")
                        } else {
                            format!("{p}{sym}")
                        }
                    }
                    // Integers overflow with a run-time error, as in VBA (spec §2.3).
                    Place::Native(p) if *prefix || !want => format!("{p} = $rt.i32({p} {op} 1)"),
                    Place::Native(p) => format!("$rt.post({p}, {p} = $rt.i32({p} {op} 1))"),
                    _ => {
                        let cur = self.place_get(&place);
                        let set = self.place_set(&place, format!("{cur} {op} 1"));
                        if *prefix || !want {
                            set
                        } else {
                            format!("$rt.post({cur}, {set})")
                        }
                    }
                }
            }
            _ => self.expr(e),
        }
    }

    fn resolve_class(&self, t: &ast::TypeRef) -> Option<ClassId> {
        let ast::TypeRef::Named { path, .. } = t else {
            return None;
        };
        let md = &self.prog.modules[self.module];
        let sym = if path.len() == 1 {
            md.scope
                .get(&path[0].text)
                .or_else(|| self.prog.globals.get(&path[0].text))
                .cloned()
        } else {
            match md.scope.get(&path[0].text) {
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
            Some(Sym::Class(c)) => Some(c),
            _ => None,
        }
    }

    fn enum_value(&self, en: EnumId, i: usize) -> String {
        let d = &self.prog.enums[en];
        let m = &d.members[i];
        // The name is kept as a comment so that the generated code stays readable in the debugger.
        if m.value < 0 {
            format!("({} /* {} */)", m.value, m.name)
        } else {
            format!("{} /* {} */", m.value, m.name)
        }
    }

    fn name_expr(&mut self, e: &Expr, text: &str) -> String {
        match self.info().res.get(&e.id).cloned() {
            Some(Res::Local(l)) => self.local_js(l),
            Some(Res::Global(v)) => self.global_ref(v),
            Some(Res::Fn(f)) => {
                if matches!(self.prog.fns[f].kind, FnKind::Host { .. }) {
                    // Host function read like a property (`Now`, `Cells`).
                    let qualified = self.info().auto_qualified.contains(&e.id);
                    self.host_call(None, f, &[], qualified)
                } else {
                    self.fn_ref(f)
                }
            }
            Some(Res::EnumMember(en, i)) => self.enum_value(en, i),
            Some(Res::Intrinsic(i)) => self.intrinsic(e, i, None, &[]),
            Some(Res::Class(c)) => self.class_ref(c),
            _ => js_ident(text),
        }
    }

    fn member_expr(&mut self, e: &Expr, object: &Expr, member: &str, optional: bool) -> String {
        let res = self.info().res.get(&e.id).cloned();
        let obj_ty = self.ty(object);
        let obj_is_ns = matches!(obj_ty, Type::Ns(_));
        let dot = if optional { "?." } else { "." };
        if matches!(obj_ty, Type::Union(_))
            && self.is_late_bound(&obj_ty)
            && matches!(
                res,
                Some(Res::Field(..) | Res::Method(_) | Res::Property { .. })
            )
        {
            let o = self.expr(object);
            return self.host_get(o, member, optional);
        }
        match res {
            Some(Res::Field(c, idx)) => {
                if self.prog.classes[c].kind == ClassKind::Host {
                    let o = self.expr(object);
                    let prog = self.prog;
                    self.host_get(o, prog.classes[c].fields[idx].emitted_name(), optional)
                } else if obj_is_ns {
                    format!("{}.{member}", self.class_ref(c))
                } else {
                    format!("{}{dot}{member}", self.expr(object))
                }
            }
            Some(Res::Global(v)) => self.global_ref(v),
            Some(Res::EnumMember(en, i)) => self.enum_value(en, i),
            Some(Res::Intrinsic(i)) => self.intrinsic(e, i, Some(object), &[]),
            Some(Res::Fn(f)) => {
                if matches!(self.prog.fns[f].kind, FnKind::Host { .. }) {
                    self.host_call(None, f, &[], false)
                } else {
                    self.fn_ref(f)
                }
            }
            Some(Res::Class(c)) => self.class_ref(c),
            Some(Res::Property { get: Some(g), .. }) if obj_is_ns => self.fn_ref(g),
            Some(Res::Method(f)) if matches!(self.prog.fns[f].kind, FnKind::Host { .. }) => {
                let o = self.expr(object);
                let prog = self.prog;
                self.host_get(o, prog.fns[f].vba_name(), optional)
            }
            _ => {
                let o = self.expr(object);
                if self.is_late_bound(&obj_ty) {
                    self.host_get(o, member, optional)
                } else {
                    format!("{o}{dot}{member}")
                }
            }
        }
    }

    fn host_get(&mut self, obj: String, member: &str, optional: bool) -> String {
        if optional {
            format!("$rt.getOpt({obj}, {})", js_str(member))
        } else {
            format!("$rt.get({obj}, {})", js_str(member))
        }
    }

    fn index_expr(&mut self, object: &Expr, index: &Expr) -> String {
        let ot = self.ty(object);
        let o = self.expr(object);
        if let (Type::Struct(_), ExprKind::Int { value, .. }) = (&ot, &index.peel().kind) {
            return format!("{o}.item{value}");
        }
        if let ExprKind::Range { start, end } = &index.kind {
            // `$rt.range(xs, start, startFromEnd, end, endFromEnd)` (spec §4.5)
            let mut argv = vec![o];
            for b in [start, end] {
                match b.as_deref() {
                    None => argv.extend(["undefined".to_string(), "false".to_string()]),
                    Some(Expr {
                        kind: ExprKind::FromEnd(k),
                        ..
                    }) => argv.extend([self.expr(k), "true".to_string()]),
                    Some(x) => argv.extend([self.expr(x), "false".to_string()]),
                }
            }
            let r = format!("$rt.range({})", argv.join(", "));
            return match ot.elem() {
                Some(Type::Struct(_)) => format!("{r}.map($rt.clone)"),
                _ => r,
            };
        }
        let i = self.expr(index);
        if matches!(index.kind, ExprKind::FromEnd(_)) {
            return match &ot {
                Type::Str => format!("$rt.charEnd({o}, {i})"),
                _ => format!("{o}[$rt.ixEnd({o}, {i})]"),
            };
        }
        match &ot {
            // An index outside the array is a run-time error in VBA, not `undefined`.
            Type::Array(_) => format!("{o}[$rt.ix({o}, {i})]"),
            Type::Str => format!("{o}.charAt({i})"),
            Type::Map(_, v) => {
                let z = self.zero(v);
                format!("$rt.mapGet({o}, {i}, {z})")
            }
            _ => format!("$rt.item({o}, [{i}])"),
        }
    }

    pub(crate) fn place(&mut self, e: &Expr) -> Place {
        let inner = e.peel();
        match &inner.kind {
            ExprKind::Index { object, index } => {
                let ot = self.ty(object);
                let o = self.expr(object);
                if let (Type::Struct(_), ExprKind::Int { value, .. }) = (&ot, &index.peel().kind) {
                    return Place::Native(format!("{o}.item{value}"));
                }
                let i = self.expr(index);
                if matches!(index.kind, ExprKind::FromEnd(_)) {
                    return Place::Native(format!("{o}[$rt.ixEnd({o}, {i})]"));
                }
                match &ot {
                    Type::Array(_) => Place::Native(format!("{o}[$rt.ix({o}, {i})]")),
                    Type::Map(..) => Place::MapItem { map: o, key: i },
                    _ => Place::Item { obj: o, index: i },
                }
            }
            ExprKind::Member { object, name, .. } => {
                let res = self.info().res.get(&inner.id).cloned();
                let obj_ty = self.ty(object);
                if matches!(obj_ty, Type::Union(_)) && self.is_late_bound(&obj_ty) {
                    return Place::Host {
                        obj: self.expr(object),
                        name: name.text.clone(),
                    };
                }
                match res {
                    Some(Res::Field(c, idx)) if self.prog.classes[c].kind == ClassKind::Host => {
                        Place::Host {
                            obj: self.expr(object),
                            name: self.prog.classes[c].fields[idx].emitted_name().to_string(),
                        }
                    }
                    Some(Res::Field(c, _)) if matches!(obj_ty, Type::Ns(_)) => {
                        Place::Native(format!("{}.{}", self.class_ref(c), name.text))
                    }
                    Some(Res::Global(v)) => Place::Native(self.global_ref(v)),
                    Some(Res::Field(..)) | Some(Res::Property { .. }) => {
                        Place::Native(format!("{}.{}", self.expr(object), name.text))
                    }
                    _ if self.is_late_bound(&obj_ty) => Place::Host {
                        obj: self.expr(object),
                        name: name.text.clone(),
                    },
                    _ => Place::Native(format!("{}.{}", self.expr(object), name.text)),
                }
            }
            _ => Place::Native(self.expr(inner)),
        }
    }

    pub(crate) fn place_get(&mut self, p: &Place) -> String {
        match p {
            Place::Native(s) => s.clone(),
            Place::Host { obj, name } => format!("$rt.get({obj}, {})", js_str(name)),
            Place::Item { obj, index } => format!("$rt.item({obj}, [{index}])"),
            Place::MapItem { map, key } => format!("$rt.mapGet({map}, {key}, null)"),
        }
    }

    pub(crate) fn place_set(&mut self, p: &Place, v: String) -> String {
        match p {
            Place::Native(s) => format!("{s} = {v}"),
            Place::Host { obj, name } => format!("$rt.put({obj}, {}, {v})", js_str(name)),
            Place::Item { obj, index } => format!("$rt.setItem({obj}, [{index}], {v})"),
            Place::MapItem { map, key } => format!("$rt.mapSet({map}, {key}, {v})"),
        }
    }

    fn assign(&mut self, op: Option<BinaryOp>, target: &Expr, value: &Expr) -> String {
        let tty = self.ty(target);
        let place = self.place(target);
        match op {
            None => {
                let v = self.coerce_store(value, &tty);
                self.place_set(&place, v)
            }
            Some(op) => {
                let vty = self.ty(value);
                let cur = self.place_get(&place);
                let v = self.expr(value);
                let result = if tty == Type::Str {
                    Type::Str
                } else {
                    tty.clone()
                };
                let combined = self.arith(op, cur, &tty, v, &vty, &result);
                self.place_set(&place, combined)
            }
        }
    }

    fn arith(
        &mut self,
        op: BinaryOp,
        l: String,
        lt: &Type,
        r: String,
        rt: &Type,
        result: &Type,
    ) -> String {
        if op == BinaryOp::Add && *result == Type::Str {
            let a = self.to_string(l, lt);
            let b = self.to_string(r, rt);
            return format!("({a} + {b})");
        }
        if (*lt == Type::Date || *rt == Type::Date) && matches!(op, BinaryOp::Add | BinaryOp::Sub) {
            let sym = if op == BinaryOp::Add { "+" } else { "-" };
            return if *result == Type::Date {
                format!("$rt.date(+{l} {sym} +{r})")
            } else {
                format!("(+{l} {sym} +{r})")
            };
        }
        let dynamic = is_dyn(lt) || is_dyn(rt);
        if dynamic
            && (op.is_arithmetic()
                || matches!(
                    op,
                    BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge
                ))
        {
            // Variant arithmetic: the operand types decide at run time.
            return match op {
                BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
                    format!("($rt.cmp({l}, {r}) {} 0)", op.symbol())
                }
                _ => format!("$rt.dyn({}, {l}, {r})", js_str(op.symbol())),
            };
        }
        let ints = int_like(lt) && int_like(rt);
        let checked = |expr: String| match result {
            Type::Byte | Type::Short | Type::Int | Type::Enum(_) => format!("$rt.i32({expr})"),
            Type::Long | Type::Ptr => format!("$rt.i64({expr})"),
            _ => format!("({expr})"),
        };
        match op {
            BinaryOp::Add => checked(format!("{l} + {r}")),
            BinaryOp::Sub => checked(format!("{l} - {r}")),
            BinaryOp::Mul => checked(format!("{l} * {r}")),
            // Integer division truncates toward zero; dividing by zero raises (spec §2.3).
            BinaryOp::Div if ints => format!("$rt.idiv({l}, {r})"),
            BinaryOp::Div => format!("$rt.div({l}, {r})"),
            BinaryOp::Rem if ints => format!("$rt.imod({l}, {r})"),
            BinaryOp::Rem => format!("$rt.fmod({l}, {r})"),
            BinaryOp::Pow => format!("$rt.pow({l}, {r})"),
            BinaryOp::Shl => format!("({l} << {r})"),
            BinaryOp::Shr => format!("({l} >> {r})"),
            BinaryOp::BitAnd => format!("({l} & {r})"),
            BinaryOp::BitOr => format!("({l} | {r})"),
            BinaryOp::BitXor => format!("({l} ^ {r})"),
            BinaryOp::Eq => format!("({l} === {r})"),
            BinaryOp::Ne => format!("({l} !== {r})"),
            BinaryOp::And => format!("({l} && {r})"),
            BinaryOp::Or => format!("({l} || {r})"),
            BinaryOp::Coalesce => format!("({l} ?? {r})"),
            BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge => {
                format!("({l} {} {r})", op.symbol())
            }
        }
    }

    fn binary(&mut self, e: &Expr, op: BinaryOp, lhs: &Expr, rhs: &Expr) -> String {
        let (lt, rt) = (self.ty(lhs), self.ty(rhs));
        match op {
            BinaryOp::And | BinaryOp::Or | BinaryOp::Coalesce => {
                let l = self.expr(lhs);
                let r = self.expr(rhs);
                format!("({l} {} {r})", op.symbol())
            }
            BinaryOp::Eq | BinaryOp::Ne => {
                let negate = op == BinaryOp::Ne;
                if lt == Type::Null || rt == Type::Null {
                    if lt == Type::Null && rt == Type::Null {
                        return (!negate).to_string();
                    }
                    let (other, oty) = if lt == Type::Null {
                        (rhs, &rt)
                    } else {
                        (lhs, &lt)
                    };
                    let v = self.expr(other);
                    return if oty.is_object() || is_dyn(oty) {
                        format!("({v} {} null)", if negate { "!=" } else { "==" })
                    } else {
                        // A value type is never null.
                        negate.to_string()
                    };
                }
                let l = self.expr(lhs);
                let r = self.expr(rhs);
                // A union may hold a value of another kind: such values are not equal (spec §2.6.2).
                if lt.is_value_union() || rt.is_value_union() {
                    format!("{}$rt.unionEq({l}, {r})", if negate { "!" } else { "" })
                } else if is_dyn(&lt) || is_dyn(&rt) {
                    format!("{}$rt.eq({l}, {r})", if negate { "!" } else { "" })
                } else if lt == Type::Date || rt == Type::Date {
                    format!("(+{l} {} +{r})", if negate { "!==" } else { "===" })
                } else {
                    format!("({l} {} {r})", if negate { "!==" } else { "===" })
                }
            }
            _ => {
                let result = self.ty(e);
                let l = self.expr(lhs);
                let r = self.expr(rhs);
                self.arith(op, l, &lt, r, &rt, &result)
            }
        }
    }

    fn arrow(&mut self, e: &Expr, a: &ast::ArrowFn) -> String {
        let locals = self
            .info()
            .lambda_params
            .get(&e.id)
            .cloned()
            .unwrap_or_default();
        let params: Vec<String> = a
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                locals
                    .get(i)
                    .map(|&l| self.local_name(l))
                    .unwrap_or_else(|| js_ident(&p.name.text))
            })
            .collect();
        let (saved_ret, saved_resume) = (self.ret_ty.take(), self.resume);
        self.resume = 0;
        let out = match &a.body {
            ast::ArrowBody::Expr(x) => {
                let v = self.expr(x);
                format!("(({}) => {v})", params.join(", "))
            }
            ast::ArrowBody::Block(b) => {
                let body = self.capture(|g| {
                    g.indent += 1;
                    g.stmts(&b.stmts);
                    g.indent -= 1;
                });
                format!("(({}) => {{\n{body}{}}})", params.join(", "), self.pad())
            }
        };
        self.ret_ty = saved_ret;
        self.resume = saved_resume;
        out
    }

    // ------------------------------------------------------------------ calls

    fn plain_args(&mut self, args: &[Expr]) -> String {
        let mut vals: Vec<String> = args.iter().map(|a| self.expr(a)).collect();
        while vals.last().is_some_and(|v| v == "undefined") {
            vals.pop();
        }
        vals.join(", ")
    }

    /// A `{ v }` cell through which the callee reads and writes the caller's variable.
    fn by_ref(&mut self, a: &Expr) -> String {
        if let ExprKind::Name(_) | ExprKind::This = &a.peel().kind {
            if let Some(Res::Local(l)) = self.info().res.get(&a.peel().id) {
                if self.local(*l).is_some_and(|d| d.byref) {
                    // Already a cell: pass it on.
                    return self.local_name(*l);
                }
            }
        }
        let place = self.place(a);
        let get = self.place_get(&place);
        let set = self.place_set(&place, "$v".into());
        format!("{{ get v() {{ return {get}; }}, set v($v) {{ {set}; }} }}")
    }

    /// Arguments of a call to the RAS function `f` (parameters from `skip` on).
    fn ras_args(&mut self, f: FnId, skip: usize, args: &[Expr]) -> Vec<String> {
        let prog = self.prog;
        let def = &prog.fns[f];
        let params = &def.params[skip.min(def.params.len())..];
        let mut out = Vec::new();
        for (i, a) in args.iter().enumerate() {
            let v = match params.get(i) {
                _ if matches!(a.kind, ExprKind::Undefined) => "undefined".to_string(),
                Some(p) if p.mode == PassMode::ByRef => self.by_ref(a),
                // Arrays and structs are passed by reference in the generated VBA: no copy here either.
                Some(p) if !p.rest && !p.ty.contains_param() => {
                    let from = self.ty(a);
                    if from == Type::Null {
                        if p.ty.is_object() || is_dyn(&p.ty) {
                            "null".into()
                        } else {
                            self.zero(&p.ty)
                        }
                    } else {
                        let v = self.expr(a);
                        self.coerce(v, &from, &p.ty)
                    }
                }
                _ => self.expr(a),
            };
            out.push(v);
        }
        while out.last().is_some_and(|v| v == "undefined") {
            out.pop();
        }
        out
    }

    /// Call of a host (COM / VBA) function. `recv` is the lowered receiver of a member call.
    fn host_call(
        &mut self,
        recv: Option<String>,
        f: FnId,
        args: &[Expr],
        qualified: bool,
    ) -> String {
        let prog = self.prog;
        let def = &prog.fns[f];
        let mut vals: Vec<String> = Vec::new();
        for (i, a) in args.iter().enumerate() {
            let v = if self.ty(a) == Type::Null
                && def.params.get(i).is_some_and(|p| p.ty.is_object())
            {
                // `Nothing`, not `Empty`.
                "$rt.Nothing".to_string()
            } else {
                self.expr(a)
            };
            vals.push(v);
        }
        while vals.last().is_some_and(|v| v == "undefined") {
            vals.pop();
        }
        let argv = vals.join(", ");
        let name = js_str(def.vba_name());
        if let Some(o) = recv {
            return format!("$rt.call({o}, {name}, [{argv}])");
        }
        if let (true, Some(q)) = (qualified, self.qualifier.clone()) {
            return format!("$rt.call({q}, {name}, [{argv}])");
        }
        match prog.modules[def.module].key.as_str() {
            "vba.d" => format!("$rt.vba.{}({argv})", def.vba_name()),
            "msforms.d" => format!("$rt.msforms.{}({argv})", def.vba_name()),
            "test.d" => format!("$rt.test.{}({argv})", def.vba_name()),
            "excel.d" | "office.d" => {
                format!("$rt.call($rt.global(\"Application\"), {name}, [{argv}])")
            }
            // Declared in the project's own type definitions: a procedure of the workbook.
            _ => format!("$rt.runMacro({name}, [{argv}])"),
        }
    }

    pub(crate) fn call(&mut self, e: &Expr) -> String {
        let ExprKind::Call { callee, args, .. } = &e.kind else {
            return "undefined".into();
        };
        let args = self.info().call_args(e.id, args);
        let Some(ci) = self.info().calls.get(&e.id) else {
            return "undefined".into();
        };
        let c = callee.peel();
        let recv = match &c.kind {
            ExprKind::Member { object, .. } => Some(&**object),
            _ => None,
        };
        match ci.target.clone() {
            CallTarget::Intrinsic(i) => self.intrinsic(e, i, recv, args),
            CallTarget::Fn(f) => {
                if let FnKind::Declare { lib, alias } = &self.prog.fns[f].kind {
                    let name = alias.as_deref().unwrap_or(&self.prog.fns[f].name);
                    let emulated = lib.to_lowercase().trim_end_matches(".dll") == "kernel32"
                        && matches!(
                            name,
                            "Sleep"
                                | "GetTickCount"
                                | "GetTickCount64"
                                | "GetCurrentProcessId"
                                | "QueryPerformanceCounter"
                                | "QueryPerformanceFrequency"
                        );
                    if !emulated {
                        let _ = self.unsupported(e.span, "`declare` 関数（DLL 呼び出し）");
                    }
                }
                let target = self.fn_ref(f);
                let argv = self.ras_args(f, 0, args).join(", ");
                format!("{target}({argv})")
            }
            CallTarget::Ext(f) => {
                let Some(r) = recv else {
                    return "undefined".into();
                };
                let byref = self.prog.fns[f]
                    .params
                    .first()
                    .is_some_and(|p| p.mode == PassMode::ByRef);
                let rv = if byref { self.by_ref(r) } else { self.expr(r) };
                let target = self.fn_ref(f);
                let mut argv = vec![rv];
                argv.extend(self.ras_args(f, 1, args));
                format!("{target}({})", argv.join(", "))
            }
            CallTarget::Method(f) => {
                let Some(r) = recv else {
                    return "undefined".into();
                };
                // `Sheet1.clearInput()`: an `export function` of the sheet's file (spec §12.2).
                if matches!(self.prog.fns[f].kind, FnKind::Free)
                    && self.prog.modules[self.prog.fns[f].module]
                        .document()
                        .is_some()
                {
                    let target = self.fn_ref(f);
                    let argv = self.ras_args(f, 0, args).join(", ");
                    return format!("{target}({argv})");
                }
                let rt = self.ty(r);
                let o = self.expr(r);
                let argv = self.ras_args(f, 0, args).join(", ");
                if matches!(rt, Type::Union(_)) && self.is_late_bound(&rt) {
                    return format!(
                        "$rt.call({o}, {}, [{argv}])",
                        js_str(&self.prog.fns[f].name)
                    );
                }
                format!("{o}.{}({argv})", self.prog.fns[f].name)
            }
            CallTarget::Host(f) => {
                let is_member = matches!(self.prog.fns[f].kind, FnKind::Host { class: Some(_) });
                let recv = recv.filter(|r| is_member && !matches!(self.ty(r), Type::Ns(_)));
                let Some(r) = recv else {
                    let qualified = self.info().auto_qualified.contains(&c.id);
                    return self.host_call(None, f, args, qualified);
                };
                let o = self.expr(r);
                // `ws.Range(Cells(1, 1), Cells(2, 2))`: the inner references follow the receiver (spec §13.4).
                let info = self.info();
                let qualifies = self.prog.fns[f].vba_name() == "Range"
                    && args.iter().any(|a| {
                        crate::astutil::any_expr_in(a, &mut |x| info.auto_qualified.contains(&x.id))
                    });
                if qualifies {
                    let q = self.tmp("q");
                    let outer = self.qualifier.replace(q.clone());
                    let inner = self.host_call(Some(q.clone()), f, args, false);
                    self.qualifier = outer;
                    format!("(({q}) => {inner})({o})")
                } else {
                    self.host_call(Some(o), f, args, false)
                }
            }
            CallTarget::DefaultMember(_) => {
                let o = self.expr(callee);
                let argv = self.plain_args(args);
                format!("$rt.item({o}, [{argv}])")
            }
            CallTarget::Closure => {
                let o = self.expr(callee);
                let argv = self.plain_args(args);
                format!("{o}({argv})")
            }
            CallTarget::Dynamic | CallTarget::Ctor(_) => {
                let argv = self.plain_args(args);
                match &c.kind {
                    ExprKind::Super => format!("super({argv})"),
                    ExprKind::Member { object, name, .. }
                        if matches!(object.kind, ExprKind::Super) =>
                    {
                        format!("super.{}({argv})", name.text)
                    }
                    ExprKind::Member { object, name, .. } => {
                        let o = self.expr(object);
                        format!("$rt.call({o}, {}, [{argv}])", js_str(&name.text))
                    }
                    _ => {
                        let o = self.expr(callee);
                        format!("$rt.invoke({o}, [{argv}])")
                    }
                }
            }
        }
    }

    fn new_expr(&mut self, e: &Expr, ty: &ast::TypeRef, args: &[Expr]) -> String {
        if let ast::TypeRef::Named { path, .. } = ty {
            if path.len() == 1 {
                match path[0].text.as_str() {
                    "Error" => return format!("$rt.makeError({})", self.plain_args(args)),
                    "Map" if matches!(self.ty(e), Type::Map(..)) => return "new Map()".into(),
                    "Array" => {
                        // `new Array<int>(rows, cols)` is `int[][]`: each cell starts as an `int`.
                        let elem = match args.len() {
                            2 => self.ty(e).elem_of().elem_of(),
                            _ => self.ty(e).elem_of(),
                        };
                        let z = self.zero(&elem);
                        let sizes = self.plain_args(args);
                        return format!("$rt.newArray(() => {z}, [{sizes}])");
                    }
                    _ => {}
                }
            }
        }
        match self.info().calls.get(&e.id).map(|c| c.target.clone()) {
            Some(CallTarget::Ctor(c)) => {
                let def = &self.prog.classes[c];
                // A user form (spec §14.11.9).
                if let Some(m) = self.prog.form_module_of(c) {
                    let argv = match self.prog.modules[m].form.as_ref().and_then(|f| f.ctor) {
                        Some(ctor) => self.ras_args(ctor, 0, args).join(", "),
                        None => String::new(),
                    };
                    return format!(
                        "$rt.newForm({}, [{argv}])",
                        js_str(&self.prog.modules[m].vba_name)
                    );
                }
                if def.kind == ClassKind::Host {
                    return format!(
                        "$rt.newHost({}, {})",
                        js_str(def.namespace.as_deref().unwrap_or("")),
                        js_str(&def.name)
                    );
                }
                let argv = match def.ctor {
                    Some(ctor) => self.ras_args(ctor, 0, args).join(", "),
                    None => String::new(),
                };
                let created = format!("new {}({argv})", self.class_ref(c));
                let targs = match self.ty(e) {
                    Type::Class(_, a) => a,
                    _ => Vec::new(),
                };
                let zeros: Vec<String> = def
                    .type_params
                    .iter()
                    .zip(targs.iter())
                    .filter(|(_, t)| matches!(t, Type::Struct(_)))
                    .map(|(p, t)| format!("{}: () => {}", js_ident(&p.name), self.zero(t)))
                    .collect();
                if zeros.is_empty() {
                    created
                } else {
                    format!(
                        "Object.assign({created}, {{ $zero: {{ {} }} }})",
                        zeros.join(", ")
                    )
                }
            }
            _ => "null".into(),
        }
    }

    // ------------------------------------------------------------------ intrinsics

    /// `target.on("<Event>", handler)` (spec §12.2.3): the run time keeps the handler and the
    /// forms and the stand-in for Excel call it when they raise the event of `target`.
    fn event_on(&mut self, e: &Expr, target: &str, args: &[Expr]) -> String {
        let Some(bind) = self.info().event_binds.get(&e.id).cloned() else {
            return "null".into();
        };
        match bind.kind {
            EventBindKind::OnAction { each } => match bind.handler_fn {
                Some(f) => {
                    let d = &self.prog.fns[f];
                    let macro_name = format!("{}.{}", self.prog.modules[d.module].vba_name, d.name);
                    format!(
                        "$rt.{}({target}, {})",
                        if each { "onActionEach" } else { "onAction" },
                        js_str(&macro_name)
                    )
                }
                None => "null".into(),
            },
            EventBindKind::Listener {
                classes,
                dynamic,
                each,
                event,
                params,
                ..
            } => {
                let handler = match (bind.handler_fn, args.get(1)) {
                    (Some(f), _) => self.fn_ref(f),
                    (None, Some(a)) => self.expr(a),
                    (None, None) => "null".into(),
                };
                let cancel = params
                    .iter()
                    .position(|p| p.2 && p.1 == "Boolean")
                    .map(|i| i as i64)
                    .unwrap_or(-1);
                // The form whose code attaches the handler: it goes when the form is unloaded.
                let owner = if self.prog.modules[self.module].is_form() {
                    "$me"
                } else {
                    "null"
                };
                if dynamic || each {
                    // The members (the target) of the classes that have the event.
                    let kinds: Vec<String> = classes
                        .iter()
                        .map(|c| js_str(&self.prog.classes[*c].name))
                        .collect();
                    return format!(
                        "$rt.{}({target}, [{}], {}, {handler}, {owner}, {cancel})",
                        if each { "listenEach" } else { "listenAny" },
                        kinds.join(", "),
                        js_str(&event)
                    );
                }
                format!(
                    "$rt.listen({target}, {}, {handler}, {owner}, {cancel})",
                    js_str(&event)
                )
            }
        }
    }

    fn intrinsic(&mut self, e: &Expr, intr: Intr, recv: Option<&Expr>, args: &[Expr]) -> String {
        use Intr::*;
        let recv_ty = recv.map(|r| self.ty(r)).unwrap_or(Type::Error);
        let r = match recv {
            Some(r) if !matches!(recv_ty, Type::Ns(_)) => self.expr(r),
            _ => String::new(),
        };
        // Array-style methods of a host collection work on a copy of its members (spec §4.4).
        let (r, recv_ty) = match self.prog.collection_elem(&recv_ty) {
            Some(elem)
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
                        | CollToArray
                ) =>
            {
                (format!("Array.from($rt.iter({r}))"), Type::array(elem))
            }
            _ => (r, recv_ty),
        };
        if intr == EventOn {
            return self.event_on(e, &r, args);
        }
        if intr.is_file_io() {
            return self.file_io(intr, args);
        }
        let a: Vec<String> = args.iter().map(|x| self.expr(x)).collect();
        let argv = a.join(", ");
        let arg =
            |i: usize, default: &str| a.get(i).cloned().unwrap_or_else(|| default.to_string());
        let elem = recv_ty.elem_of();
        match intr {
            EventOn => unreachable!(),
            FileOpen | FileClose | FileReset | FilePrint | FilePrintLine | FileWrite
            | FileWriteLine | FileInput | FileLineInput | FileInputString | FileGet | FilePut
            | FileSeek | FileLock | FileUnlock | FileWidth | FileRename => unreachable!(),
            ArrLength | ArrRows | StrLength => format!("{r}.length"),
            ArrCols => format!("$rt.cols({r})"),
            ArrTranspose => format!("$rt.transpose({r})"),
            MapSize => format!("{r}.size"),
            ArrPush => {
                let v = match args.first() {
                    Some(x) => self.coerce_store(x, &elem),
                    None => "undefined".into(),
                };
                format!("{r}.push({v})")
            }
            ArrPop => {
                let z = self.zero(&elem);
                format!("$rt.pop({r}, {z})")
            }
            CollToArray => r,
            ArrMap => format!("{r}.map({argv})"),
            ArrFilter => format!("{r}.filter({argv})"),
            ArrForEach => format!("{r}.forEach({argv})"),
            ArrReduce => format!("{r}.reduce({argv})"),
            ArrSome => format!("{r}.some({argv})"),
            ArrEvery => format!("{r}.every({argv})"),
            ArrFind => {
                let z = self.zero(&elem);
                format!("$rt.find({r}, {}, {z})", arg(0, "null"))
            }
            ArrFindIndex => format!("{r}.findIndex({argv})"),
            ArrIndexOf => format!("{r}.indexOf({argv})"),
            ArrIncludes => format!("{r}.includes({argv})"),
            ArrJoin => format!("$rt.join({r}, {})", arg(0, "\",\"")),
            ArrSlice => format!("{r}.slice({argv})"),
            // Struct elements are values: the new array holds copies (as the VBA build does).
            ArrStep | ArrToReversed => {
                let v = if intr == ArrStep {
                    format!("$rt.step({r}, {argv})")
                } else {
                    format!("{r}.slice().reverse()")
                };
                if matches!(elem, Type::Struct(_)) {
                    format!("{v}.map($rt.clone)")
                } else {
                    v
                }
            }
            ArrSort => format!(
                "$rt.sort({r}{}{argv})",
                if a.is_empty() { "" } else { ", " }
            ),
            ArrReverse => format!("{r}.reverse()"),
            ArrConcat => format!("{r}.concat({argv})"),
            ArrClear => format!("{r}.length = 0"),
            // VBA's Trim only removes spaces.
            StrTrim => format!("$rt.trim({r}, 3)"),
            StrTrimStart => format!("$rt.trim({r}, 1)"),
            StrTrimEnd => format!("$rt.trim({r}, 2)"),
            StrUpper => format!("{r}.toUpperCase()"),
            StrLower => format!("{r}.toLowerCase()"),
            StrToString => r,
            StrSubstring => format!("{r}.substring({argv})"),
            StrSubstr => format!("{r}.substr({argv})"),
            StrSlice => format!("{r}.slice({argv})"),
            StrIndexOf => format!("{r}.indexOf({argv})"),
            StrLastIndexOf => format!("{r}.lastIndexOf({argv})"),
            StrIncludes => format!("{r}.includes({argv})"),
            StrStartsWith => format!("{r}.startsWith({argv})"),
            StrEndsWith => format!("{r}.endsWith({argv})"),
            StrPadStart => format!("{r}.padStart({argv})"),
            StrPadEnd => format!("{r}.padEnd({argv})"),
            StrRepeat => format!("$rt.repeat({r}, {argv})"),
            StrReplace => format!("$rt.replace({r}, {argv}, false)"),
            StrReplaceAll => format!("$rt.replace({r}, {argv}, true)"),
            StrSplit => format!("$rt.split({r}, {argv})"),
            StrCharAt => format!("{r}.charAt({argv})"),
            StrCharCodeAt => format!("$rt.charCode({r}, {argv})"),
            NumToString | BoolToString | DateToString => self.to_string(r, &recv_ty),
            NumToFixed => format!("$rt.toFixed({r}, {})", arg(0, "0")),
            MapGet => {
                let z = match &recv_ty {
                    Type::Map(_, v) => self.zero(v),
                    _ => "null".into(),
                };
                format!("$rt.mapGet({r}, {}, {z})", arg(0, "null"))
            }
            MapSet => {
                let v = match (&recv_ty, args.get(1)) {
                    (Type::Map(_, vt), Some(x)) => self.coerce_store(x, vt),
                    _ => arg(1, "null"),
                };
                format!("$rt.mapSet({r}, {}, {v})", arg(0, "null"))
            }
            MapHas => format!("{r}.has({argv})"),
            MapDelete => format!("{r}.delete({argv})"),
            MapKeys => format!("[...{r}.keys()]"),
            MapValues => format!("[...{r}.values()]"),
            MapClear => format!("{r}.clear()"),
            MathAbs => format!("Math.abs({argv})"),
            MathFloor => format!("Math.floor({argv})"),
            MathCeil => format!("Math.ceil({argv})"),
            MathTrunc => format!("Math.trunc({argv})"),
            MathRound => format!("$rt.round({argv})"),
            MathBankersRound => format!("$rt.bankersRound({argv})"),
            MathSqrt => format!("$rt.sqrt({argv})"),
            MathPow => format!("$rt.pow({argv})"),
            MathMin => format!("Math.min({argv})"),
            MathMax => format!("Math.max({argv})"),
            MathSin => format!("Math.sin({argv})"),
            MathCos => format!("Math.cos({argv})"),
            MathTan => format!("Math.tan({argv})"),
            MathAtan => format!("Math.atan({argv})"),
            MathLog => format!("$rt.log({argv})"),
            MathExp => format!("Math.exp({argv})"),
            MathRandom => "Math.random()".into(),
            MathSign => format!("Math.sign({argv})"),
            MathPi => "Math.PI".into(),
            ParseInt => format!("$rt.toInt($rt.val({argv}))"),
            ParseFloat => format!("$rt.val({argv})"),
            ToStr => {
                let at = args.first().map(|x| self.ty(x)).unwrap_or(Type::Any);
                self.to_string(arg(0, "\"\""), &at)
            }
            ConsoleLog => {
                let parts: Vec<String> = args
                    .iter()
                    .zip(a.iter())
                    .map(|(x, v)| {
                        let at = self.ty(x);
                        self.to_string(v.clone(), &at)
                    })
                    .collect();
                format!("$rt.print({})", parts.join(", "))
            }
            IsNullish => "false".into(),
            FastMode => crate::lower::expr::fast_mode_mask(e).to_string(),
            SuspendScreen => "$rt.suspend(1)".into(),
            ResumeScreen => "$rt.resume(1)".into(),
            SuspendEvents => "$rt.suspend(2)".into(),
            ResumeEvents => "$rt.resume(2)".into(),
            SuspendCalc => "$rt.suspend(4)".into(),
            ResumeCalc => "$rt.resume(4)".into(),
            SuspendAlerts => "$rt.suspend(8)".into(),
            ResumeAlerts => "$rt.resume(8)".into(),
            SuspendAll => "$rt.suspend(63)".into(),
            ResumeAll => "$rt.resume(63)".into(),
            Status => format!("$rt.status({})", arg(0, "\"\"")),
            Sleep => format!("$rt.sleep({argv})"),
            WaitUntil => format!("$rt.waitUntil({argv})"),
            YieldNow => "$rt.yieldNow()".into(),
            PlatformIs64 => "$rt.is64()".into(),
            PlatformIsVba7 => "true".into(),
            FormShowModal => {
                let z = match self.ty(e) {
                    Type::Void => "undefined".to_string(),
                    t => self.zero(&t),
                };
                format!("{r}.$showModal({z})")
            }
            FormClose => {
                let result = match recv.map(|x| self.ty(x)) {
                    Some(Type::Class(c, _)) => self
                        .prog
                        .form_module_of(c)
                        .and_then(|m| self.prog.modules[m].form.as_ref())
                        .and_then(|f| f.result.clone())
                        .unwrap_or(Type::Void),
                    _ => Type::Void,
                };
                let v = match (args.first(), &result) {
                    (_, Type::Void) => "undefined".to_string(),
                    (Some(a), t) => {
                        let t = t.clone();
                        self.coerce_store(a, &t)
                    }
                    (None, t) => self.zero(t),
                };
                format!("{r}.$close({v})")
            }
            RangeToArray => format!("$rt.read2D({r})"),
            SheetUsedTable => format!("$rt.read2D($rt.get({r}, \"UsedRange\"))"),
            RangeToColumn => format!("$rt.toColumn({r})"),
            RangeToRow => format!("$rt.toRow({r})"),
            RangeExtend => {
                let side = |i: usize| match args.get(i) {
                    Some(x) if !matches!(x.kind, ExprKind::Undefined) => a[i].clone(),
                    _ => "0".to_string(),
                };
                format!(
                    "$rt.extend({r}, {}, {}, {}, {})",
                    side(0),
                    side(1),
                    side(2),
                    side(3)
                )
            }
            RangeFromArray => format!("$rt.fromArray({r}, {argv})"),
            ExcelMatch => format!("$rt.match({argv})"),
            ExcelErrorToNull => format!("$rt.errorToNull({argv})"),
        }
    }
}

trait ElemOf {
    fn elem_of(&self) -> Type;
}

impl ElemOf for Type {
    /// Element type of an array (one level: the element of `T[][]` is `T[]`).
    fn elem_of(&self) -> Type {
        match self {
            Type::Array(inner) => (**inner).clone(),
            _ => Type::Any,
        }
    }
}

/// How the run time converts what it reads into a variable of type `t` (`Input #`).
fn file_kind(t: &Type) -> &'static str {
    match t {
        Type::Str => "s",
        Type::Bool => "b",
        Type::Date => "d",
        Type::Byte => "u8",
        Type::Short => "i16",
        Type::Int | Type::Enum(_) => "i32",
        Type::Long | Type::Ptr => "i64",
        Type::Float => "f32",
        Type::Double | Type::Decimal => "f64",
        Type::Currency => "cur",
        _ => "v",
    }
}

/// The layout of a value written by `Put #` / read by `Get #`: its kind, `[kind]` for an array.
fn binary_kind(t: &Type) -> String {
    match t {
        Type::Array(e) => format!("[{}]", file_kind(e)),
        other => file_kind(other).to_string(),
    }
}

impl Gen<'_> {
    /// The statements of the files opened by number (spec §15.3.2), on the run time's file table.
    fn file_io(&mut self, intr: Intr, args: &[Expr]) -> String {
        use Intr::*;
        match intr {
            FilePrint | FilePrintLine => {
                let n = self.expr(&args[0]);
                let parts: Vec<String> = args[1..]
                    .iter()
                    .map(|x| {
                        let at = self.ty(x);
                        let v = self.expr(x);
                        self.to_string(v, &at)
                    })
                    .collect();
                format!(
                    "$rt.filePrint({n}, [{}], {})",
                    parts.join(", "),
                    intr == FilePrintLine
                )
            }
            FileWrite | FileWriteLine => {
                let a: Vec<String> = args.iter().map(|x| self.expr(x)).collect();
                format!(
                    "$rt.fileWrite({}, [{}], {})",
                    a[0],
                    a[1..].join(", "),
                    intr == FileWriteLine
                )
            }
            FileInput => {
                let n = self.expr(&args[0]);
                let targets: Vec<String> = args[1..]
                    .iter()
                    .map(|x| {
                        let kind = file_kind(&self.ty(x));
                        let cell = self.by_ref(x);
                        format!("[{cell}, \"{kind}\"]")
                    })
                    .collect();
                format!("$rt.fileInput({n}, [{}])", targets.join(", "))
            }
            FileGet | FilePut => {
                let n = self.expr(&args[0]);
                let kind = binary_kind(&self.ty(&args[1]));
                let v = if intr == FileGet && !self.ty(&args[1]).is_array() {
                    self.by_ref(&args[1])
                } else {
                    self.expr(&args[1])
                };
                let rec = match args.get(2) {
                    Some(r) => self.expr(r),
                    None => "undefined".into(),
                };
                let f = if intr == FileGet {
                    "fileGet"
                } else {
                    "filePut"
                };
                format!("$rt.{f}({n}, {v}, \"{kind}\", {rec})")
            }
            _ => {
                let a: Vec<String> = args.iter().map(|x| self.expr(x)).collect();
                let f = match intr {
                    FileOpen => "fileOpen",
                    FileClose => "fileClose",
                    FileReset => "fileReset",
                    FileLineInput => "fileLineInput",
                    FileInputString => "fileInputString",
                    FileSeek => "fileSeek",
                    FileLock => "fileLock",
                    FileUnlock => "fileUnlock",
                    FileWidth => "fileWidth",
                    _ => "fileRename",
                };
                format!("$rt.{f}({})", a.join(", "))
            }
        }
    }
}
