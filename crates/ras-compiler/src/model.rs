//! Semantic model: types, declarations and the per-module analysis tables.

use std::collections::HashMap;
use std::path::PathBuf;
use std::rc::Rc;

use ras_syntax::ast::{self, ExprId, InlineMode, PassMode, Span, Visibility};

use crate::diag::Diagnostic;

pub type ModuleId = usize;
pub type FnId = usize;
pub type ClassId = usize;
pub type EnumId = usize;
pub type StructId = usize;
pub type VarId = usize;
pub type LocalId = usize;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[derive(Default)]
pub enum Type {
    /// Poison type: produced after an error, compatible with everything.
    Error,
    #[default]
    Void,
    /// The type of `null` / `undefined`.
    Null,
    Byte,
    Short,
    Int,
    Long,
    Float,
    Double,
    Currency,
    Decimal,
    Ptr,
    Date,
    Str,
    Bool,
    Any,
    Object,
    /// `declare`-only string marshalling types (spec §11.2).
    WStr,
    AStr,
    /// `T[]`. `Array(Array(T))` is a rectangular 2-D array (spec §9.2.1).
    Array(Box<Type>),
    /// RAS class, RAS interface or host (COM) interface, with type arguments.
    Class(ClassId, Vec<Type>),
    Enum(EnumId),
    Struct(StructId),
    Map(Box<Type>, Box<Type>),
    Func(Box<FuncType>),
    /// A type parameter of the enclosing generic declaration.
    Param(String),
    /// `A | B`: two or more distinct member types, flattened and in canonical order
    /// (build with [`Type::union`]). Stored in an `Object` when every member is an object
    /// type, otherwise in a `Variant`.
    Union(Vec<Type>),
    /// Not a value: a module namespace, host namespace, class or enum name, or built-in namespace.
    Ns(NsRef),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum NsRef {
    Module(ModuleId),
    Host(String),
    Class(ClassId),
    Enum(EnumId),
    Builtin(&'static str),
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct FuncType {
    pub params: Vec<Type>,
    pub ret: Type,
}


impl Type {
    pub fn array(elem: Type) -> Type {
        Type::Array(Box::new(elem))
    }

    pub fn is_error(&self) -> bool {
        matches!(self, Type::Error)
    }

    pub fn is_integer(&self) -> bool {
        matches!(self, Type::Byte | Type::Short | Type::Int | Type::Long)
    }

    pub fn is_float(&self) -> bool {
        matches!(self, Type::Float | Type::Double)
    }

    pub fn is_numeric(&self) -> bool {
        self.is_integer() || self.is_float() || matches!(self, Type::Currency | Type::Decimal | Type::Enum(_) | Type::Ptr)
    }

    /// Position in the C#-style promotion order `int < long < float < double`.
    pub fn rank(&self) -> Option<u8> {
        Some(match self {
            Type::Byte | Type::Short | Type::Int | Type::Enum(_) => 1,
            Type::Long => 2,
            Type::Float => 3,
            Type::Double => 4,
            _ => return None,
        })
    }

    /// Types whose values are VBA object references (assigned with `Set`).
    pub fn is_object(&self) -> bool {
        match self {
            Type::Union(ms) => ms.iter().all(|m| m.is_object()),
            _ => matches!(self, Type::Object | Type::Class(..) | Type::Map(..) | Type::Func(_) | Type::Null),
        }
    }

    /// The members of a union type.
    pub fn union_members(&self) -> Option<&[Type]> {
        match self {
            Type::Union(ms) => Some(ms),
            _ => None,
        }
    }

    /// A union stored in a `Variant` (it has a member that is not an object type).
    pub fn is_value_union(&self) -> bool {
        matches!(self, Type::Union(ms) if !ms.iter().all(|m| m.is_object()))
    }

    /// A union stored in a `Variant` that may also hold an object reference (`string | Excel.Range`).
    pub fn is_mixed_union(&self) -> bool {
        matches!(self, Type::Union(ms) if ms.iter().any(|m| m.is_object()) && !ms.iter().all(|m| m.is_object()))
    }

    /// Values whose kind is only known at run time: they live in a `Variant` and may hold an object.
    pub fn is_dynamic(&self) -> bool {
        matches!(self, Type::Any | Type::Error | Type::Param(_)) || self.is_mixed_union()
    }

    /// `A | B | ...`: flattens nested unions, drops duplicates and sorts the members into a
    /// canonical order. `any` absorbs everything, `object` absorbs the other object types,
    /// and a single remaining member is returned as it is.
    pub fn union(members: impl IntoIterator<Item = Type>) -> Type {
        let mut flat: Vec<Type> = Vec::new();
        for m in members {
            match m {
                Type::Union(inner) => flat.extend(inner),
                other => flat.push(other),
            }
        }
        if flat.iter().any(|m| m.is_error()) {
            return Type::Error;
        }
        if flat.contains(&Type::Any) {
            return Type::Any;
        }
        flat.retain(|m| *m != Type::Null);
        if flat.contains(&Type::Object) {
            flat.retain(|m| *m == Type::Object || !m.is_object());
        }
        let mut out: Vec<Type> = Vec::new();
        for m in flat {
            if !out.contains(&m) {
                out.push(m);
            }
        }
        out.sort_by_cached_key(Type::order_key);
        match out.len() {
            0 => Type::Null,
            1 => out.pop().unwrap(),
            _ => Type::Union(out),
        }
    }

    /// Sort key of the canonical member order of a union: value types first, then object types.
    fn order_key(&self) -> (u8, usize, String) {
        let rank = match self {
            Type::Str => 0,
            Type::Bool => 1,
            Type::Byte => 2,
            Type::Short => 3,
            Type::Int => 4,
            Type::Long => 5,
            Type::Float => 6,
            Type::Double => 7,
            Type::Currency => 8,
            Type::Decimal => 9,
            Type::Ptr => 10,
            Type::Date => 11,
            Type::Enum(_) => 12,
            Type::Struct(_) => 13,
            Type::Array(_) => 14,
            Type::Class(..) => 15,
            Type::Map(..) => 16,
            Type::Func(_) => 17,
            Type::Object => 18,
            _ => 19,
        };
        let id = match self {
            Type::Enum(i) | Type::Struct(i) | Type::Class(i, _) => *i,
            _ => 0,
        };
        (rank, id, format!("{self:?}"))
    }

    pub fn is_array(&self) -> bool {
        matches!(self, Type::Array(_))
    }

    /// Number of array dimensions (0 for non-arrays; 2 for `T[][]`).
    pub fn dims(&self) -> usize {
        match self {
            Type::Array(inner) => {
                if matches!(**inner, Type::Array(_)) {
                    2
                } else {
                    1
                }
            }
            _ => 0,
        }
    }

    /// Innermost element type of a 1-D or 2-D array.
    pub fn elem(&self) -> Option<&Type> {
        match self {
            Type::Array(inner) => match &**inner {
                Type::Array(inner2) => Some(inner2),
                other => Some(other),
            },
            _ => None,
        }
    }

    /// A struct, or an array of structs: values that cannot be stored in a `Variant`.
    pub fn has_struct(&self) -> bool {
        match self {
            Type::Struct(_) => true,
            Type::Array(t) => t.has_struct(),
            Type::Union(ms) => ms.iter().any(|m| m.has_struct()),
            _ => false,
        }
    }

    pub fn is_value(&self) -> bool {
        !matches!(self, Type::Ns(_) | Type::Void)
    }

    pub fn contains_param(&self) -> bool {
        match self {
            Type::Param(_) => true,
            Type::Array(t) => t.contains_param(),
            Type::Class(_, args) => args.iter().any(|a| a.contains_param()),
            Type::Map(k, v) => k.contains_param() || v.contains_param(),
            Type::Func(f) => f.params.iter().any(|p| p.contains_param()) || f.ret.contains_param(),
            Type::Union(ms) => ms.iter().any(|m| m.contains_param()),
            _ => false,
        }
    }

    pub fn subst(&self, env: &HashMap<String, Type>) -> Type {
        if env.is_empty() {
            return self.clone();
        }
        match self {
            Type::Param(name) => env.get(name).cloned().unwrap_or_else(|| self.clone()),
            Type::Array(t) => Type::Array(Box::new(t.subst(env))),
            Type::Class(id, args) => Type::Class(*id, args.iter().map(|a| a.subst(env)).collect()),
            Type::Map(k, v) => Type::Map(Box::new(k.subst(env)), Box::new(v.subst(env))),
            Type::Func(f) => Type::Func(Box::new(FuncType {
                params: f.params.iter().map(|p| p.subst(env)).collect(),
                ret: f.ret.subst(env),
            })),
            Type::Union(ms) => Type::union(ms.iter().map(|m| m.subst(env))),
            other => other.clone(),
        }
    }
}

/// Built-in type-parameter constraints (spec §2.4.1) or an interface type.
#[derive(Clone, Debug, PartialEq)]
pub enum Constraint {
    None,
    Comparable,
    Numeric,
    ObjectType,
    Type(Type),
}

#[derive(Clone, Debug)]
pub struct TypeParamDef {
    pub name: String,
    pub constraint: Constraint,
    /// `T = Default` of a class or interface: the type argument when it is omitted.
    pub default: Option<Type>,
}

#[derive(Clone, Debug)]
pub struct ParamDef {
    pub name: String,
    pub ty: Type,
    pub mode: PassMode,
    pub optional: bool,
    pub default: Option<Rc<ast::Expr>>,
    pub rest: bool,
    pub span: Span,
    /// Local slot of this parameter in `FnDef::locals`.
    pub local: LocalId,
}

#[derive(Clone, Debug, PartialEq)]
pub enum FnKind {
    Free,
    Method { class: ClassId, is_static: bool },
    Ctor { class: ClassId },
    Getter { class: ClassId, is_static: bool },
    Setter { class: ClassId, is_static: bool },
    /// `declare module "lib" { function ... }`
    Declare { lib: String, alias: Option<String> },
    /// Member of a host (COM) interface or a declared global host function.
    Host { class: Option<ClassId> },
    /// Interface method signature (RAS interface).
    Abstract { class: ClassId },
    /// The synthesized `Main` holding a module's top-level statements.
    Main,
    /// The synthesized initializer of a module's top-level variables.
    ModInit,
}

/// Decorator-derived attributes of a function.
#[derive(Clone, Debug, Default)]
pub struct FnAttrs {
    pub macro_entry: bool,
    pub on_events: Vec<OnEvent>,
    pub init: bool,
    pub udf: bool,
    pub fast: bool,
    pub silent: bool,
    pub reentrant: bool,
    pub nocancel: bool,
    pub nodoevents: bool,
    pub doevents: bool,
    pub on_error: Option<String>,
    pub ribbon: bool,
    pub fallback: Option<String>,
    /// Declared in a form's code-behind: a member of the form's code module.
    pub form_member: bool,
    /// Declared in the file of a sheet or of `ThisWorkbook` (spec §12.2): a member of that
    /// document module.
    pub doc_member: bool,
}

#[derive(Clone, Debug)]
pub struct OnEvent {
    /// `Workbook` or a sheet code name.
    pub target: String,
    /// Event name (`Open`, `Change`, `SheetChange` …).
    pub event: String,
    pub reentrant: bool,
    pub span: Span,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum LocalKind {
    Param,
    Let,
    Const,
    Using,
    Catch,
    /// Loop variable of `for...of`.
    Iter,
    /// Parameter of an arrow function.
    LambdaParam,
    /// The receiver of an extension method (`this`).
    This,
}

#[derive(Clone, Debug)]
pub struct LocalDef {
    pub name: String,
    pub ty: Type,
    pub kind: LocalKind,
    pub byref: bool,
    /// Assigned anywhere after its declaration.
    pub reassigned: bool,
    /// Referenced from inside an arrow function declared in an outer scope.
    pub captured: bool,
    pub span: Span,
    pub used: bool,
}

#[derive(Clone, Debug)]
pub struct FnDef {
    pub id: FnId,
    pub module: ModuleId,
    /// Name in RAS source. For members and globals of the bundled type definitions this is
    /// the camelCase spelling (`cells`); `host_name` keeps the type library's (`Cells`).
    pub name: String,
    pub host_name: Option<String>,
    pub kind: FnKind,
    pub exported: bool,
    pub vis: Visibility,
    pub inline: InlineMode,
    pub type_params: Vec<TypeParamDef>,
    pub params: Vec<ParamDef>,
    /// True when `params[0]` is the `this` receiver of an extension method.
    pub is_extension: bool,
    /// `None` until inferred (functions without a return type annotation).
    pub ret: Option<Type>,
    pub attrs: FnAttrs,
    pub doc: Option<String>,
    pub name_span: Span,
    pub decl: Option<Rc<ast::FnDecl>>,
    pub locals: Vec<LocalDef>,
    /// Functions called directly from the body (for the call graph).
    pub calls: Vec<FnId>,
    /// Module-level variables referenced from the body.
    pub var_refs: Vec<VarId>,
    pub checked: bool,
    pub has_loops: bool,
    /// Body contains constructs that block inlining (spec §8.2).
    pub inline_blockers: Vec<&'static str>,
    pub cost: u32,
}

impl FnDef {
    pub fn ret_ty(&self) -> Type {
        self.ret.clone().unwrap_or(Type::Void)
    }

    /// The name to emit: the type library's spelling for host functions.
    pub fn vba_name(&self) -> &str {
        self.host_name.as_deref().unwrap_or(&self.name)
    }

    /// The VBA identifier of the procedure. A function of a module must not hide a function of
    /// the VBA library (`names::module_ident`); a member of a class hides nothing outside it.
    pub fn proc_ident(&self) -> String {
        match self.kind {
            FnKind::Free => crate::names::module_ident(&self.name),
            _ => crate::names::member_ident(&self.name),
        }
    }

    pub fn is_generic(&self) -> bool {
        !self.type_params.is_empty()
    }

    pub fn required_params(&self) -> usize {
        self.params.iter().filter(|p| !p.optional && p.default.is_none() && !p.rest).count()
    }

    /// A macro of the workbook: shown in Excel's macro dialog and runnable with `Application.Run`.
    pub fn is_macro_entry(&self) -> bool {
        if self.is_generic() || !self.attrs.on_events.is_empty() || self.attrs.init || self.attrs.udf || self.attrs.form_member {
            return false;
        }
        match self.kind {
            FnKind::Main => true,
            FnKind::Free => self.attrs.macro_entry || (self.exported && self.params.is_empty() && self.ret_ty() == Type::Void),
            _ => false,
        }
    }

    /// A function the debugger can start (spec §14.10): one without required arguments that
    /// Excel does not call by itself. Only meaningful for functions with a body.
    pub fn is_debug_entry(&self) -> bool {
        matches!(self.kind, FnKind::Free | FnKind::Main) && self.required_params() == 0 && !self.is_extension && !self.is_generic() && self.attrs.on_events.is_empty()
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ClassKind {
    Class,
    Interface,
    /// Host (COM) interface from a `.d.ras` file.
    Host,
}

#[derive(Clone, Debug)]
pub struct FieldDef {
    /// Name in RAS source (camelCase for the bundled type definitions, see `FnDef::name`).
    pub name: String,
    pub host_name: Option<String>,
    pub ty: Type,
    pub vis: Visibility,
    pub is_static: bool,
    pub readonly: bool,
    pub init: Option<Rc<ast::Expr>>,
    pub doc: Option<String>,
    pub span: Span,
}

impl FieldDef {
    pub fn vba_name(&self) -> &str {
        self.host_name.as_deref().unwrap_or(&self.name)
    }
}

#[derive(Clone, Debug)]
pub struct ClassDef {
    pub id: ClassId,
    pub module: ModuleId,
    pub name: String,
    pub kind: ClassKind,
    /// Host namespace (`Excel`) for host types.
    pub namespace: Option<String>,
    pub exported: bool,
    pub type_params: Vec<TypeParamDef>,
    pub extends: Vec<Type>,
    pub implements: Vec<Type>,
    pub fields: Vec<FieldDef>,
    /// Methods, accessors and the constructor.
    pub methods: Vec<FnId>,
    pub ctor: Option<FnId>,
    /// Default member (`(index: any): T` call signature / indexer).
    pub call_sig: Option<FnId>,
    pub predeclared: bool,
    pub doc: Option<String>,
    pub name_span: Span,
}

#[derive(Clone, Debug)]
pub struct EnumDef {
    pub id: EnumId,
    pub module: ModuleId,
    pub name: String,
    pub host: bool,
    /// A host enum of a library the generated VBA does not reference (`Scripting`): its
    /// members are emitted as numbers.
    pub late_bound: bool,
    pub exported: bool,
    pub members: Vec<EnumMemberDef>,
    pub doc: Option<String>,
    pub name_span: Span,
}

#[derive(Clone, Debug)]
pub struct EnumMemberDef {
    pub name: String,
    pub value: i64,
    pub doc: Option<String>,
    pub span: Span,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StructKind {
    /// `struct Name { ... }`
    Named,
    /// Anonymous `{ a: A; b: B }`: identified by its fields (stored sorted by name).
    Object,
    /// `[A, B]`: fields `item0`, `item1`, …
    Tuple,
}

#[derive(Clone, Debug)]
pub struct StructDef {
    pub id: StructId,
    pub module: ModuleId,
    /// Name of the VBA `Type` (generated for anonymous structs).
    pub name: String,
    pub kind: StructKind,
    pub fields: Vec<(String, Type)>,
    /// Per field: declared with `?` (may be omitted in a literal).
    pub optional: Vec<bool>,
    pub doc: Option<String>,
    pub name_span: Span,
}

#[derive(Clone, Debug, PartialEq)]
pub enum ConstVal {
    Int(i128),
    Float(f64),
    Str(String),
    Bool(bool),
}

#[derive(Clone, Debug)]
pub struct VarDef {
    pub id: VarId,
    pub module: ModuleId,
    /// Name in RAS source (camelCase for the bundled host globals, see `FnDef::name`).
    pub name: String,
    pub host_name: Option<String>,
    pub ty: Type,
    pub is_const: bool,
    pub exported: bool,
    /// Host global from a `.d.ras` file (`Application`, `Sheet1`).
    pub declared: bool,
    pub init: Option<Rc<ast::Expr>>,
    /// Compile-time constant value of a `const` (emitted as VBA `Const`).
    pub const_value: Option<ConstVal>,
    pub doc: Option<String>,
    pub name_span: Span,
    /// Referenced from a function other than the module's `Main`.
    pub used_outside_main: bool,
    pub reassigned: bool,
}

impl VarDef {
    pub fn vba_name(&self) -> &str {
        self.host_name.as_deref().unwrap_or(&self.name)
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Sym {
    Fn(FnId),
    Class(ClassId),
    Enum(EnumId),
    EnumMember(EnumId, usize),
    Var(VarId),
    Struct(StructId),
    Module(ModuleId),
    Namespace(String),
    Alias(Type),
}

/// The host classes `new` and `CreateObject` create by their ProgID: (ProgID, namespace, name).
/// The generated VBA creates them with `CreateObject`, without a reference to their library.
pub const PROG_IDS: &[(&str, &str, &str)] = &[
    ("Scripting.FileSystemObject", "Scripting", "FileSystemObject"),
    ("Scripting.Dictionary", "Scripting", "Dictionary"),
    ("VBScript.RegExp", "VBScript", "RegExp"),
];

/// The ProgID of a host class of namespace `namespace`, when it has one.
pub fn prog_id(namespace: Option<&str>, name: &str) -> Option<&'static str> {
    PROG_IDS.iter().find(|(_, ns, n)| Some(*ns) == namespace && *n == name).map(|(id, ..)| *id)
}

/// Whether the types of a host namespace are those of a library the VBA project references
/// (the others are late-bound: `Object`, and numbers for their enums; spec §15.3).
pub fn referenced_namespace(namespace: Option<&str>) -> bool {
    matches!(namespace, None | Some("Excel" | "VBA" | "Office" | "MSForms"))
}

/// What a name / member / call expression resolved to.
#[derive(Clone, Debug, PartialEq)]
pub enum Res {
    Local(LocalId),
    Global(VarId),
    Fn(FnId),
    Class(ClassId),
    Enum(EnumId),
    EnumMember(EnumId, usize),
    Module(ModuleId),
    Namespace(String),
    Builtin(&'static str),
    /// Field of a RAS class or property of a host interface.
    Field(ClassId, usize),
    /// Method of a class / host interface referenced through member access.
    Method(FnId),
    Property { get: Option<FnId>, set: Option<FnId> },
    ExtMethod(FnId),
    Intrinsic(crate::intrinsics::Intr),
    /// Late-bound member of an `any` / `object` value.
    Dynamic,
}

#[derive(Clone, Debug, PartialEq)]
pub enum CallTarget {
    /// Free function, static method or `declare` function.
    Fn(FnId),
    /// Instance method; the receiver is the object of the callee member expression.
    Method(FnId),
    /// Extension method; the receiver becomes the first argument.
    Ext(FnId),
    /// Host member call or host global function.
    Host(FnId),
    /// Default member of a host value (`Worksheets("x")`).
    DefaultMember(FnId),
    Ctor(ClassId),
    Intrinsic(crate::intrinsics::Intr),
    /// Calling a function-typed value (closure object).
    Closure,
    Dynamic,
}

#[derive(Clone, Debug)]
pub struct CallInfo {
    pub target: CallTarget,
    /// Explicit or inferred type arguments of a generic callee.
    pub type_args: Vec<Type>,
}

/// Why an optimisation was or was not applied (spec §6 rule 13); shown as hints / CodeLens.
#[derive(Clone, Debug)]
pub struct OptNote {
    pub span: Span,
    pub applied: bool,
    pub kind: &'static str,
    pub message: String,
}

#[derive(Debug, Default, Clone)]
pub struct TypeInfo {
    pub expr_types: Vec<Type>,
    pub res: HashMap<ExprId, Res>,
    pub calls: HashMap<ExprId, CallInfo>,
    /// Declaration site (start offset of the name) → local slot of the enclosing function.
    pub decl_locals: HashMap<u32, LocalId>,
    /// Arrow function (by expression id) → its parameter locals.
    pub lambda_params: HashMap<ExprId, Vec<LocalId>>,
    /// Arrow function → captured outer locals.
    pub lambda_captures: HashMap<ExprId, Vec<LocalId>>,
    /// Top-level `let` / `const` executed inside `Main` / the module initializer (name offset → variable).
    pub decl_globals: HashMap<u32, VarId>,
    /// Unqualified `Cells` / `Range` calls inside `ws.Range(...)` that get the receiver prepended (spec §13.4).
    pub auto_qualified: std::collections::HashSet<ExprId>,
    /// Calls with named arguments (spec §3.3) → their arguments in parameter order. A skipped
    /// parameter is an `undefined` expression; the values keep their ids.
    pub positional_args: HashMap<ExprId, Rc<[ast::Expr]>>,
    /// `target.on("<Event>", handler)` calls (spec §12.2.3) → what they do.
    pub event_binds: HashMap<ExprId, EventBind>,
}

/// What a call of `target.on("<Event>", handler)` / `target.addEventListener(…)` does
/// (spec §12.2.3), as the checker found it.
#[derive(Clone, Debug)]
pub struct EventBind {
    pub kind: EventBindKind,
    /// The handler is the name of a function (not a value: an arrow function or a variable).
    pub handler_fn: Option<FnId>,
}

#[derive(Clone, Debug)]
pub enum EventBindKind {
    /// A listener object holds the target in a `WithEvents` variable of the host class
    /// `class` and calls the handler from the event procedure. `params` are those of the
    /// event (name, VBA type, by reference); the handler takes the first `argc` of them.
    Listener { class: ClassId, event: String, params: Vec<(String, String, bool)>, argc: usize },
    /// The target's `OnAction` property gets the name of the macro.
    OnAction,
}

impl TypeInfo {
    pub fn ty(&self, id: ExprId) -> Type {
        self.expr_types.get(id as usize).cloned().unwrap_or(Type::Error)
    }

    /// The arguments of call (or `new`) `call` in parameter order: `args` unless named
    /// arguments were reordered (spec §3.3).
    pub fn call_args<'a>(&'a self, call: ExprId, args: &'a [ast::Expr]) -> &'a [ast::Expr] {
        self.positional_args.get(&call).map(|v| &v[..]).unwrap_or(args)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    User,
    /// Bundled standard library module written in RAS.
    Std,
    /// `.d.ras` type definitions (no code is generated).
    TypeDef,
}

#[derive(Clone, Debug)]
pub struct ModuleData {
    pub id: ModuleId,
    pub path: PathBuf,
    /// Display path relative to the project root (used in source maps and messages).
    pub rel_path: String,
    /// Import key: path relative to `src/` without extension, `/`-separated.
    pub key: String,
    pub vba_name: String,
    pub kind: ModuleKind,
    pub text: Rc<str>,
    pub ast: Rc<ast::Module>,
    pub scope: HashMap<String, Sym>,
    pub exports: HashMap<String, Sym>,
    /// Extension methods visible in this module (own and imported).
    pub ext_methods: Vec<FnId>,
    pub info: TypeInfo,
    pub main_fn: Option<FnId>,
    /// Synthesized initializer for top-level variables (modules without top-level statements).
    pub init_fn: Option<FnId>,
    /// Functions, in declaration order.
    pub fns: Vec<FnId>,
    pub classes: Vec<ClassId>,
    pub enums: Vec<EnumId>,
    pub structs: Vec<StructId>,
    pub vars: Vec<VarId>,
    pub imports: Vec<ModuleId>,
    pub line_starts: Vec<u32>,
    /// Set for a module that is the code of an object of the workbook: a form's code-behind
    /// (`forms/<Name>.ras`, spec §14.11), or the file of a sheet or of `ThisWorkbook`
    /// (`objects/<CodeName>.ras`, spec §12.2).
    pub form: Option<FormModule>,
}

/// A variable or a class field whose object's events are handled (`@on("app.SheetChange")`,
/// spec §12.2.2): VBA's `WithEvents`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum EventSource {
    /// A module variable of a sheet's file, of `ThisWorkbook`'s or of a form's code-behind.
    Var(VarId),
    /// A field of a class (the handler is a method of the class).
    Field(ClassId, usize),
}

/// Which object of the workbook a module is the code of.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ObjectKind {
    /// A user form (`forms/<Name>.ras`).
    Form,
    /// A worksheet (`objects/<CodeName>.ras`).
    Sheet,
    /// The workbook (`objects/ThisWorkbook.ras`).
    Workbook,
}

/// A module that is the code of a user form, of a sheet or of the workbook.
#[derive(Clone, Debug)]
pub struct FormModule {
    pub kind: ObjectKind,
    /// Name and controls; `None` when `forms/<Name>.form.json` is missing or unreadable.
    /// For a sheet, its code name and the ActiveX controls on it (spec §12.2.1).
    pub decl: Option<crate::config::FormDecl>,
    /// The type of `this`: the form (sheet, workbook) with its controls as fields.
    pub class: Option<ClassId>,
}

impl Program {
    /// The variable or field the handler `f` names in `@on("<target>.<Event>")`, when the
    /// target is neither the object of the file (`Worksheet`, `Workbook`, `UserForm`) nor a
    /// control: the object it holds is listened to (spec §12.2.2). A variable of a standard
    /// module is found too (VBA cannot listen there: an error of the caller's).
    pub fn event_source(&self, f: FnId, target: &str) -> Option<EventSource> {
        let d = &self.fns[f];
        match d.kind {
            FnKind::Method { class, is_static: false } => self.classes[class].fields.iter().position(|x| !x.is_static && x.name.eq_ignore_ascii_case(target)).map(|i| EventSource::Field(class, i)),
            FnKind::Free => {
                let md = &self.modules[d.module];
                if let Some(form) = &md.form {
                    let reserved = ["UserForm", "Worksheet", "Workbook", "ThisWorkbook"].iter().any(|r| r.eq_ignore_ascii_case(target)) || md.vba_name.eq_ignore_ascii_case(target);
                    let control = form.decl.as_ref().is_some_and(|x| x.controls.iter().any(|c| c.name.eq_ignore_ascii_case(target)));
                    if reserved || control {
                        return None;
                    }
                }
                md.vars.iter().copied().find(|&v| !self.vars[v].declared && self.vars[v].module == d.module && self.vars[v].name.eq_ignore_ascii_case(target)).map(EventSource::Var)
            }
            _ => None,
        }
    }
}

impl ModuleData {
    /// A form's code-behind (`forms/<Name>.ras`).
    pub fn is_form(&self) -> bool {
        self.form.as_ref().is_some_and(|f| f.kind == ObjectKind::Form)
    }

    /// The file of a sheet or of `ThisWorkbook` (`objects/<CodeName>.ras`): a document module.
    pub fn document(&self) -> Option<&FormModule> {
        self.form.as_ref().filter(|f| f.kind != ObjectKind::Form)
    }

    /// 1-based line number of a byte offset.
    pub fn line_of(&self, offset: u32) -> u32 {
        match self.line_starts.binary_search(&offset) {
            Ok(i) => i as u32 + 1,
            Err(i) => i as u32,
        }
    }

    /// 0-based (line, column in UTF-16 code units) of a byte offset — LSP positions.
    pub fn line_col_utf16(&self, offset: u32) -> (u32, u32) {
        let line = self.line_of(offset) - 1;
        let start = self.line_starts[line as usize] as usize;
        let end = (offset as usize).min(self.text.len());
        let col: usize = self.text.get(start..end).map(|s| s.chars().map(|c| c.len_utf16()).sum()).unwrap_or(0);
        (line, col as u32)
    }

    /// Byte offset of a 0-based (line, UTF-16 column) position.
    pub fn offset_of(&self, line: u32, col_utf16: u32) -> u32 {
        let Some(&start) = self.line_starts.get(line as usize) else { return self.text.len() as u32 };
        let line_end = self.line_starts.get(line as usize + 1).copied().unwrap_or(self.text.len() as u32);
        let mut units = 0u32;
        let mut off = start;
        for c in self.text[start as usize..line_end as usize].chars() {
            if units >= col_utf16 || c == '\n' || c == '\r' {
                break;
            }
            units += c.len_utf16() as u32;
            off += c.len_utf8() as u32;
        }
        off
    }
}

pub fn line_starts(text: &str) -> Vec<u32> {
    let mut v = vec![0u32];
    for (i, b) in text.bytes().enumerate() {
        if b == b'\n' {
            v.push(i as u32 + 1);
        }
    }
    v
}

/// The analysed project: all declarations plus diagnostics.
#[derive(Debug, Default)]
pub struct Program {
    pub modules: Vec<ModuleData>,
    pub fns: Vec<FnDef>,
    pub classes: Vec<ClassDef>,
    pub enums: Vec<EnumDef>,
    pub structs: Vec<StructDef>,
    pub vars: Vec<VarDef>,
    /// Host namespaces (`Excel`, `VBA`, `Office`) → members.
    pub namespaces: HashMap<String, HashMap<String, Sym>>,
    /// Names visible everywhere (host globals, host enum members).
    pub globals: HashMap<String, Sym>,
    /// The API of tests (`describe`, `expect`, ...; spec §14.12.4): visible in test modules only.
    pub test_globals: HashMap<String, Sym>,
    /// Lower-cased type-library spelling of every bundled host global → its RAS name
    /// (`cells` → `cells`, `msgbox` → `msgBox`). Finds a name written in another casing.
    pub host_global_names: HashMap<String, String>,
    /// Lower-cased name of every member of the bundled host interfaces → the type library's
    /// spelling. Used to emit a late-bound member (`activeSheet.cells`) the way VBA writes it.
    pub host_member_spelling: HashMap<String, String>,
    pub diagnostics: Vec<Diagnostic>,
    /// Corrections that can be applied mechanically (`ras check --fix`).
    pub fixes: Vec<crate::diag::Fix>,
    pub opt_notes: Vec<(ModuleId, OptNote)>,
    pub config: crate::config::Config,
    /// Host namespace whose declarations are being resolved (unqualified type names look there first).
    #[doc(hidden)]
    pub cur_namespace: Option<String>,
}

impl ModuleData {
    /// A test file or a helper module of the tests (spec §14.12.2): not built into the workbook.
    pub fn is_test(&self) -> bool {
        self.kind == ModuleKind::User && self.key.starts_with(crate::collect::TEST_KEY_PREFIX)
    }
}

impl Program {
    /// A test file (matching `[test] include`), as opposed to a helper module.
    pub fn is_test_file(&self, m: ModuleId) -> bool {
        self.modules[m].is_test() && self.config.test_files.iter().any(|f| *f == self.modules[m].rel_path)
    }

    pub fn module_by_key(&self, key: &str) -> Option<ModuleId> {
        self.modules.iter().position(|m| m.key == key)
    }

    /// The bundled host class a ProgID creates (`Scripting.FileSystemObject`), in any case.
    pub fn class_of_prog_id(&self, id: &str) -> Option<ClassId> {
        let (_, ns, name) = PROG_IDS.iter().find(|(p, ..)| p.eq_ignore_ascii_case(id.trim()))?;
        match self.namespaces.get(*ns)?.get(*name)? {
            Sym::Class(c) => Some(*c),
            _ => None,
        }
    }

    pub fn class_display(&self, id: ClassId) -> String {
        let c = &self.classes[id];
        match &c.namespace {
            Some(ns) => format!("{ns}.{}", c.name),
            None => c.name.clone(),
        }
    }

    /// RAS-syntax rendering of a type, for messages and hovers.
    pub fn show(&self, t: &Type) -> String {
        match t {
            Type::Error => "<error>".into(),
            Type::Void => "void".into(),
            Type::Null => "null".into(),
            Type::Byte => "byte".into(),
            Type::Short => "short".into(),
            Type::Int => "int".into(),
            Type::Long => "long".into(),
            Type::Float => "float".into(),
            Type::Double => "double".into(),
            Type::Currency => "currency".into(),
            Type::Decimal => "decimal".into(),
            Type::Ptr => "ptr".into(),
            Type::Date => "date".into(),
            Type::Str => "string".into(),
            Type::Bool => "boolean".into(),
            Type::Any => "any".into(),
            Type::Object => "object".into(),
            Type::WStr => "wstr".into(),
            Type::AStr => "astr".into(),
            Type::Array(e) => match &**e {
                Type::Func(_) | Type::Union(_) => format!("({})[]", self.show(e)),
                _ => format!("{}[]", self.show(e)),
            },
            Type::Union(ms) => ms
                .iter()
                .map(|m| match m {
                    Type::Func(_) => format!("({})", self.show(m)),
                    _ => self.show(m),
                })
                .collect::<Vec<_>>()
                .join(" | "),
            Type::Class(id, args) => {
                let mut s = self.class_display(*id);
                if !args.is_empty() {
                    s.push('<');
                    s.push_str(&args.iter().map(|a| self.show(a)).collect::<Vec<_>>().join(", "));
                    s.push('>');
                }
                s
            }
            Type::Enum(id) => self.enums[*id].name.clone(),
            Type::Struct(id) => {
                let d = &self.structs[*id];
                match d.kind {
                    StructKind::Named => d.name.clone(),
                    StructKind::Tuple => format!("[{}]", d.fields.iter().map(|f| self.show(&f.1)).collect::<Vec<_>>().join(", ")),
                    StructKind::Object => {
                        let fs: Vec<String> = d.fields.iter().zip(&d.optional).map(|(f, o)| format!("{}{}: {}", f.0, if *o { "?" } else { "" }, self.show(&f.1))).collect();
                        format!("{{ {} }}", fs.join("; "))
                    }
                }
            }
            Type::Map(k, v) => format!("Map<{}, {}>", self.show(k), self.show(v)),
            Type::Func(f) => {
                let ps = f.params.iter().enumerate().map(|(i, p)| format!("a{}: {}", i + 1, self.show(p))).collect::<Vec<_>>();
                format!("({}) => {}", ps.join(", "), self.show(&f.ret))
            }
            Type::Param(n) => n.clone(),
            Type::Ns(NsRef::Module(m)) => format!("module {}", self.modules[*m].key),
            Type::Ns(NsRef::Host(n)) => format!("namespace {n}"),
            Type::Ns(NsRef::Class(c)) => format!("typeof {}", self.class_display(*c)),
            Type::Ns(NsRef::Enum(e)) => format!("enum {}", self.enums[*e].name),
            Type::Ns(NsRef::Builtin(n)) => (*n).to_string(),
        }
    }

    /// The anonymous struct with these fields (created in module `m` on first use).
    /// Object types ignore the order of their fields; tuples do not.
    pub fn intern_struct(&mut self, m: ModuleId, kind: StructKind, mut fields: Vec<(String, Type, bool)>) -> StructId {
        if kind == StructKind::Object {
            fields.sort_by(|a, b| a.0.cmp(&b.0));
        }
        let found = self.structs.iter().position(|d| {
            d.kind == kind && d.fields.len() == fields.len() && d.fields.iter().zip(&d.optional).zip(&fields).all(|((f, o), g)| f.0 == g.0 && f.1 == g.1 && *o == g.2)
        });
        if let Some(id) = found {
            return id;
        }
        let id = self.structs.len();
        let n = self.structs.iter().filter(|d| d.kind == kind).count() + 1;
        let name = format!("{}{n}", if kind == StructKind::Tuple { "RAS_Tup" } else { "RAS_Obj" });
        // Generated code needs a standard module to hold the `Type`: a user module.
        let home = if self.modules[m].kind == ModuleKind::User { m } else { self.modules.iter().position(|x| x.kind == ModuleKind::User).unwrap_or(m) };
        self.structs.push(StructDef {
            id,
            module: home,
            name,
            kind,
            optional: fields.iter().map(|f| f.2).collect(),
            fields: fields.into_iter().map(|f| (f.0, f.1)).collect(),
            doc: None,
            name_span: Span::default(),
        });
        self.modules[home].structs.push(id);
        id
    }

    /// Finds a member (field index or method) of a class, searching base interfaces.
    pub fn find_field(&self, class: ClassId, name: &str) -> Option<(ClassId, usize)> {
        let c = &self.classes[class];
        if let Some(i) = c.fields.iter().position(|f| f.name == name) {
            return Some((class, i));
        }
        for base in c.extends.iter() {
            if let Type::Class(b, _) = base {
                if let Some(r) = self.find_field(*b, name) {
                    return Some(r);
                }
            }
        }
        None
    }

    pub fn find_methods(&self, class: ClassId, name: &str) -> Vec<FnId> {
        let c = &self.classes[class];
        let mut out: Vec<FnId> = c.methods.iter().copied().filter(|&f| self.fns[f].name == name).collect();
        if out.is_empty() {
            for base in c.extends.iter() {
                if let Type::Class(b, _) = base {
                    out = self.find_methods(*b, name);
                    if !out.is_empty() {
                        break;
                    }
                }
            }
        }
        out
    }

    /// True when `class` is, extends or implements `target`.
    pub fn class_conforms(&self, class: ClassId, target: ClassId) -> bool {
        if class == target {
            return true;
        }
        let c = &self.classes[class];
        c.extends.iter().chain(c.implements.iter()).any(|t| matches!(t, Type::Class(b, _) if self.class_conforms(*b, target)))
    }

    /// Element type of a host collection: a host interface with a default member
    /// (`(index: any): T`) and a `Count` property, which is what `For Each` enumerates.
    pub fn collection_elem(&self, t: &Type) -> Option<Type> {
        let Type::Class(c, args) = t else { return None };
        let def = &self.classes[*c];
        if def.kind != ClassKind::Host {
            return None;
        }
        let sig = def.call_sig?;
        self.find_field(*c, "count").or_else(|| self.find_field(*c, "Count"))?;
        Some(self.fns[sig].ret_ty().subst(&self.class_env(*c, args)))
    }

    /// Type-parameter substitution environment for a class instantiation.
    pub fn class_env(&self, class: ClassId, args: &[Type]) -> HashMap<String, Type> {
        self.classes[class].type_params.iter().zip(args.iter()).map(|(p, a)| (p.name.clone(), a.clone())).collect()
    }
}
