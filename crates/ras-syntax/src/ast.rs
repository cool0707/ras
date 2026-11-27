//! Owned abstract syntax tree, lowered from the lossless CST (`lower.rs`).
//!
//! Every expression carries an `ExprId` that is unique within its file; the
//! type checker stores types and resolutions in side tables keyed by it.

pub use rowan::{TextRange, TextSize};

pub type Span = TextRange;
pub type ExprId = u32;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Ident {
    pub text: String,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct Module {
    pub items: Vec<Item>,
    pub expr_count: u32,
}

#[derive(Debug, Clone)]
pub struct Decorator {
    pub name: Ident,
    pub args: Vec<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum Item {
    Import(Import),
    Fn(FnDecl),
    Class(ClassDecl),
    Interface(InterfaceDecl),
    Enum(EnumDecl),
    TypeAlias(TypeAlias),
    Var(VarDecl),
    DeclareModule(DeclareModule),
    Struct(StructDecl),
    Namespace(NamespaceDecl),
    /// Top-level executable statement (script style; collected into `Sub Main`).
    Stmt(Stmt),
}

#[derive(Debug, Clone)]
pub struct Import {
    pub names: Vec<ImportName>,
    /// `import * as ns from "..."`
    pub namespace: Option<Ident>,
    pub source: String,
    pub source_span: Span,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct ImportName {
    pub name: Ident,
    pub alias: Option<Ident>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InlineMode {
    Default,
    Inline,
    NoInline,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PassMode {
    ByVal,
    ByRef,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
    Protected,
}

#[derive(Debug, Clone)]
pub struct FnDecl {
    pub decorators: Vec<Decorator>,
    pub doc: Option<String>,
    pub exported: bool,
    pub declared: bool,
    pub inline: InlineMode,
    pub name: Ident,
    pub type_params: Vec<TypeParam>,
    pub params: Vec<Param>,
    pub ret: Option<TypeRef>,
    pub body: Option<Block>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct TypeParam {
    pub name: Ident,
    pub constraint: Option<TypeRef>,
    /// `T = Default`: used when the type argument is omitted.
    pub default: Option<TypeRef>,
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: Ident,
    pub ty: Option<TypeRef>,
    pub default: Option<Expr>,
    pub mode: PassMode,
    /// `x?: T`
    pub optional: bool,
    /// `this: T` receiver of an extension method
    pub is_this: bool,
    /// `...rest: T[]`
    pub rest: bool,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum TypeRef {
    Named {
        path: Vec<Ident>,
        args: Vec<TypeRef>,
        span: Span,
    },
    Array {
        elem: Box<TypeRef>,
        span: Span,
    },
    Fn {
        params: Vec<Param>,
        ret: Box<TypeRef>,
        span: Span,
    },
    /// `T | null` / `T | undefined`
    Nullable {
        inner: Box<TypeRef>,
        span: Span,
    },
    /// `A | B | ...` with two or more non-null members; `nullable` when `| null` was also written.
    Union {
        elems: Vec<TypeRef>,
        nullable: bool,
        span: Span,
    },
    /// `[A, B]` — a tuple: an anonymous struct with positional fields.
    Tuple {
        elems: Vec<TypeRef>,
        span: Span,
    },
    /// `{ a: A; b: B }` — an anonymous struct.
    Object {
        fields: Vec<PropSig>,
        span: Span,
    },
    Error {
        span: Span,
    },
}

impl TypeRef {
    pub fn span(&self) -> Span {
        match self {
            TypeRef::Named { span, .. }
            | TypeRef::Array { span, .. }
            | TypeRef::Fn { span, .. }
            | TypeRef::Nullable { span, .. }
            | TypeRef::Union { span, .. }
            | TypeRef::Tuple { span, .. }
            | TypeRef::Object { span, .. }
            | TypeRef::Error { span } => *span,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub decorators: Vec<Decorator>,
    pub doc: Option<String>,
    pub exported: bool,
    pub declared: bool,
    pub name: Ident,
    pub type_params: Vec<TypeParam>,
    pub extends: Option<TypeRef>,
    pub implements: Vec<TypeRef>,
    pub members: Vec<ClassMember>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum ClassMember {
    Field(FieldDecl),
    Method(MethodDecl),
    Ctor(FnDecl),
    Accessor(AccessorDecl),
}

#[derive(Debug, Clone)]
pub struct FieldDecl {
    pub doc: Option<String>,
    pub vis: Visibility,
    pub is_static: bool,
    pub readonly: bool,
    pub name: Ident,
    pub ty: Option<TypeRef>,
    pub init: Option<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct MethodDecl {
    pub vis: Visibility,
    pub is_static: bool,
    pub func: FnDecl,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AccessorKind {
    Get,
    Set,
}

#[derive(Debug, Clone)]
pub struct AccessorDecl {
    pub kind: AccessorKind,
    pub vis: Visibility,
    pub is_static: bool,
    pub func: FnDecl,
}

#[derive(Debug, Clone)]
pub struct InterfaceDecl {
    pub doc: Option<String>,
    pub exported: bool,
    pub declared: bool,
    pub name: Ident,
    pub type_params: Vec<TypeParam>,
    pub extends: Vec<TypeRef>,
    pub members: Vec<InterfaceMember>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum InterfaceMember {
    Prop(PropSig),
    Method(FnDecl),
    /// `[index: K]: V` — the default (indexer) member of a host type.
    Index {
        param: Param,
        ty: TypeRef,
        span: Span,
    },
}

#[derive(Debug, Clone)]
pub struct PropSig {
    pub doc: Option<String>,
    pub name: Ident,
    pub ty: Option<TypeRef>,
    pub readonly: bool,
    pub optional: bool,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct EnumDecl {
    pub doc: Option<String>,
    pub exported: bool,
    pub declared: bool,
    pub name: Ident,
    pub members: Vec<EnumMember>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct EnumMember {
    pub doc: Option<String>,
    pub name: Ident,
    pub value: Option<Expr>,
}

#[derive(Debug, Clone)]
pub struct TypeAlias {
    pub doc: Option<String>,
    pub exported: bool,
    pub name: Ident,
    pub type_params: Vec<TypeParam>,
    pub ty: TypeRef,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct DeclareModule {
    pub lib: String,
    pub lib_span: Span,
    pub functions: Vec<FnDecl>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct StructDecl {
    pub doc: Option<String>,
    pub exported: bool,
    pub name: Ident,
    pub fields: Vec<PropSig>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct NamespaceDecl {
    pub name: Ident,
    pub items: Vec<Item>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VarKind {
    Let,
    Const,
    Using,
}

#[derive(Debug, Clone)]
pub struct VarDecl {
    pub doc: Option<String>,
    pub kind: VarKind,
    pub exported: bool,
    pub declared: bool,
    pub name: Ident,
    pub ty: Option<TypeRef>,
    pub init: Option<Expr>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct Block {
    pub stmts: Vec<Stmt>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct Stmt {
    pub kind: StmtKind,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum StmtKind {
    Var(VarDecl),
    Expr(Expr),
    Block(Block),
    If {
        cond: Expr,
        then: Box<Stmt>,
        els: Option<Box<Stmt>>,
    },
    For {
        decorators: Vec<Decorator>,
        init: Vec<Stmt>,
        cond: Option<Expr>,
        update: Vec<Expr>,
        body: Box<Stmt>,
    },
    ForOf {
        decorators: Vec<Decorator>,
        kind: VarKind,
        name: Ident,
        ty: Option<TypeRef>,
        iter: Expr,
        body: Box<Stmt>,
    },
    While {
        decorators: Vec<Decorator>,
        cond: Expr,
        body: Box<Stmt>,
    },
    DoWhile {
        decorators: Vec<Decorator>,
        body: Box<Stmt>,
        cond: Expr,
    },
    Switch {
        disc: Expr,
        cases: Vec<SwitchCase>,
    },
    Try(TryStmt),
    Throw(Expr),
    Return(Option<Expr>),
    Break(Option<Ident>),
    Continue(Option<Ident>),
    Labeled {
        label: Ident,
        body: Box<Stmt>,
    },
    /// `resume;` — only valid as the sole statement of a `catch` block (spec §5.2).
    Resume,
    /// `debugger;` — pauses the debug run (spec §14.10.9); nothing in the VBA build.
    Debugger,
    /// Nested function declaration (rejected by the checker, kept for diagnostics).
    Fn(Box<FnDecl>),
    Empty,
    Error,
}

#[derive(Debug, Clone)]
pub struct SwitchCase {
    /// `None` for `default:`
    pub test: Option<Expr>,
    pub body: Vec<Stmt>,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub struct TryStmt {
    pub block: Block,
    pub catch: Option<CatchClause>,
    pub finally: Option<Block>,
}

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub param: Option<Ident>,
    pub body: Block,
    pub span: Span,
}

impl CatchClause {
    /// `catch { resume; }` — the "On Error Resume Next" form.
    pub fn is_resume_block(&self) -> bool {
        self.body.stmts.len() == 1 && matches!(self.body.stmts[0].kind, StmtKind::Resume)
    }

    pub fn contains_resume(&self) -> bool {
        self.body
            .stmts
            .iter()
            .any(|s| matches!(s.kind, StmtKind::Resume))
    }
}

#[derive(Debug, Clone)]
pub struct Expr {
    pub id: ExprId,
    pub kind: ExprKind,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NumSuffix {
    None,
    /// `L` — long
    L,
    /// `f` — float
    F,
    /// `d` — double
    D,
    /// `m` — decimal
    M,
}

#[derive(Debug, Clone)]
pub enum ExprKind {
    /// Integer literal: value (parsed, may exceed i64 → clamped and flagged), suffix, is-hex.
    Int {
        value: i128,
        suffix: NumSuffix,
        hex: bool,
    },
    /// Real literal: the normalised digits (no suffix, no `_`).
    Float {
        text: String,
        suffix: NumSuffix,
    },
    Str(String),
    Template(Vec<TemplatePart>),
    Bool(bool),
    Null,
    Undefined,
    Name(Ident),
    This,
    Super,
    Paren(Box<Expr>),
    Array(Vec<Expr>),
    Object(Vec<ObjectProp>),
    Arrow(Box<ArrowFn>),
    Unary {
        op: UnaryOp,
        operand: Box<Expr>,
    },
    /// `x++` / `x--` (postfix) and `++x` / `--x` (prefix)
    IncDec {
        inc: bool,
        prefix: bool,
        operand: Box<Expr>,
    },
    Binary {
        op: BinaryOp,
        lhs: Box<Expr>,
        rhs: Box<Expr>,
    },
    Assign {
        op: Option<BinaryOp>,
        target: Box<Expr>,
        value: Box<Expr>,
    },
    Ternary {
        cond: Box<Expr>,
        then: Box<Expr>,
        els: Box<Expr>,
    },
    Call {
        callee: Box<Expr>,
        type_args: Vec<TypeRef>,
        args: Vec<Expr>,
    },
    New {
        ty: TypeRef,
        args: Vec<Expr>,
    },
    Member {
        object: Box<Expr>,
        name: Ident,
        optional: bool,
    },
    Index {
        object: Box<Expr>,
        index: Box<Expr>,
    },
    /// `^n` — the n-th position from the end; only directly inside an index or a range.
    FromEnd(Box<Expr>),
    /// `a..b`, `..b`, `a..`, `..` — only as an index (`xs[1..^1]`) or the source of `for...of`.
    Range {
        start: Option<Box<Expr>>,
        end: Option<Box<Expr>>,
    },
    As {
        expr: Box<Expr>,
        ty: TypeRef,
    },
    InstanceOf {
        expr: Box<Expr>,
        ty: TypeRef,
    },
    NonNull(Box<Expr>),
    /// `name: value` — only directly inside the argument list of a call or `new` (spec §3.3).
    /// `hole` is a spare id for the `undefined` that fills each skipped parameter.
    NamedArg {
        name: Ident,
        value: Box<Expr>,
        hole: ExprId,
    },
    Error,
}

#[derive(Debug, Clone)]
pub enum TemplatePart {
    Str(String),
    Expr(Expr),
}

#[derive(Debug, Clone)]
pub struct ObjectProp {
    pub key: Ident,
    pub value: Expr,
}

#[derive(Debug, Clone)]
pub struct ArrowFn {
    pub params: Vec<Param>,
    pub ret: Option<TypeRef>,
    pub body: ArrowBody,
    pub span: Span,
}

#[derive(Debug, Clone)]
pub enum ArrowBody {
    Expr(Expr),
    Block(Block),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UnaryOp {
    Not,
    Neg,
    Plus,
    BitNot,
    TypeOf,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Rem,
    Pow,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    And,
    Or,
    Coalesce,
    BitAnd,
    BitOr,
    BitXor,
    Shl,
    Shr,
}

impl BinaryOp {
    pub fn is_comparison(self) -> bool {
        matches!(
            self,
            BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Le | BinaryOp::Gt | BinaryOp::Ge
        )
    }

    pub fn is_arithmetic(self) -> bool {
        matches!(
            self,
            BinaryOp::Add
                | BinaryOp::Sub
                | BinaryOp::Mul
                | BinaryOp::Div
                | BinaryOp::Rem
                | BinaryOp::Pow
        )
    }

    pub fn symbol(self) -> &'static str {
        match self {
            BinaryOp::Add => "+",
            BinaryOp::Sub => "-",
            BinaryOp::Mul => "*",
            BinaryOp::Div => "/",
            BinaryOp::Rem => "%",
            BinaryOp::Pow => "**",
            BinaryOp::Eq => "==",
            BinaryOp::Ne => "!=",
            BinaryOp::Lt => "<",
            BinaryOp::Le => "<=",
            BinaryOp::Gt => ">",
            BinaryOp::Ge => ">=",
            BinaryOp::And => "&&",
            BinaryOp::Or => "||",
            BinaryOp::Coalesce => "??",
            BinaryOp::BitAnd => "&",
            BinaryOp::BitOr => "|",
            BinaryOp::BitXor => "^",
            BinaryOp::Shl => "<<",
            BinaryOp::Shr => ">>",
        }
    }
}

impl Expr {
    /// Strips parentheses and non-null assertions.
    pub fn peel(&self) -> &Expr {
        match &self.kind {
            ExprKind::Paren(inner) | ExprKind::NonNull(inner) => inner.peel(),
            _ => self,
        }
    }
}
