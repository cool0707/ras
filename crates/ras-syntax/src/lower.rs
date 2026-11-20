//! CST → AST lowering.

use rowan::{TextRange, TextSize};

use crate::ast::*;
use crate::kind::{SyntaxKind, SyntaxKind::*, SyntaxNode, SyntaxToken};
use crate::parser::{Parser, SyntaxError};

pub(crate) struct Lower {
    next_id: u32,
    /// Offset added to every span (non-zero while lowering a template fragment).
    base: u32,
    pub(crate) errors: Vec<SyntaxError>,
    /// Counter for the hidden variables of destructuring patterns.
    pattern_n: u32,
}

fn sig_tokens(n: &SyntaxNode) -> impl Iterator<Item = SyntaxToken> {
    n.children_with_tokens().filter_map(|e| e.into_token()).filter(|t| !t.kind().is_trivia())
}

fn child(n: &SyntaxNode, kind: SyntaxKind) -> Option<SyntaxNode> {
    n.children().find(|c| c.kind() == kind)
}

fn has_tok(n: &SyntaxNode, kind: SyntaxKind) -> bool {
    sig_tokens(n).any(|t| t.kind() == kind)
}

fn has_word(n: &SyntaxNode, word: &str) -> bool {
    sig_tokens(n).any(|t| t.kind() == IDENT && t.text() == word)
}

/// Extracts the JSDoc comment (`/** ... */`) directly preceding `n`.
pub fn doc_comment(n: &SyntaxNode) -> Option<String> {
    let mut tok = n.first_token()?.prev_token();
    while let Some(t) = tok {
        match t.kind() {
            WHITESPACE => tok = t.prev_token(),
            BLOCK_COMMENT if t.text().starts_with("/**") && t.text().len() >= 5 => {
                return Some(clean_doc(t.text()));
            }
            _ => return None,
        }
    }
    None
}

fn clean_doc(raw: &str) -> String {
    let inner = raw.trim_start_matches("/**").trim_end_matches("*/");
    let mut lines: Vec<&str> = Vec::new();
    for line in inner.lines() {
        let l = line.trim_start();
        let l = l.strip_prefix('*').unwrap_or(l);
        let l = l.strip_prefix(' ').unwrap_or(l);
        lines.push(l.trim_end());
    }
    while lines.first().is_some_and(|l| l.is_empty()) {
        lines.remove(0);
    }
    while lines.last().is_some_and(|l| l.is_empty()) {
        lines.pop();
    }
    lines.join("\n")
}

/// Decodes the body of a string / template chunk (handles `\n`, `\u{...}` …).
pub fn unescape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }
        match chars.next() {
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some('t') => out.push('\t'),
            Some('b') => out.push('\u{8}'),
            Some('f') => out.push('\u{c}'),
            Some('v') => out.push('\u{b}'),
            Some('0') => out.push('\0'),
            Some('x') => {
                let hex: String = chars.by_ref().take(2).collect();
                match u32::from_str_radix(&hex, 16).ok().and_then(char::from_u32) {
                    Some(ch) => out.push(ch),
                    None => out.push_str(&hex),
                }
            }
            Some('u') => {
                let hex: String = if chars.peek() == Some(&'{') {
                    chars.next();
                    let mut h = String::new();
                    for ch in chars.by_ref() {
                        if ch == '}' {
                            break;
                        }
                        h.push(ch);
                    }
                    h
                } else {
                    chars.by_ref().take(4).collect()
                };
                match u32::from_str_radix(&hex, 16).ok().and_then(char::from_u32) {
                    Some(ch) => out.push(ch),
                    None => out.push('\u{FFFD}'),
                }
            }
            Some('\r') => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
            }
            Some('\n') => {}
            Some(other) => out.push(other),
            None => {}
        }
    }
    out
}

impl Lower {
    pub(crate) fn new() -> Self {
        Lower { next_id: 0, base: 0, errors: Vec::new(), pattern_n: 0 }
    }

    fn span(&self, r: TextRange) -> Span {
        if self.base == 0 {
            r
        } else {
            r + TextSize::from(self.base)
        }
    }

    fn nspan(&self, n: &SyntaxNode) -> Span {
        self.span(n.text_range())
    }

    fn err(&mut self, span: Span, msg: impl Into<String>) {
        self.errors.push(SyntaxError { message: msg.into(), range: span });
    }

    fn id(&mut self) -> ExprId {
        let id = self.next_id;
        self.next_id += 1;
        id
    }

    fn ident_tok(&self, t: &SyntaxToken) -> Ident {
        Ident { text: t.text().to_string(), span: self.span(t.text_range()) }
    }

    fn name_of(&self, n: &SyntaxNode) -> Option<Ident> {
        let name = child(n, NAME)?;
        let tok = sig_tokens(&name).next()?;
        Some(self.ident_tok(&tok))
    }

    fn name_or_missing(&self, n: &SyntaxNode) -> Ident {
        self.name_of(n).unwrap_or_else(|| Ident { text: String::new(), span: TextRange::empty(self.nspan(n).start()) })
    }

    // ------------------------------------------------------------------ items

    pub(crate) fn module(&mut self, root: &SyntaxNode) -> Module {
        let mut items = Vec::new();
        for c in root.children() {
            self.item(&c, &mut items);
        }
        Module { items, expr_count: self.next_id }
    }

    fn item(&mut self, n: &SyntaxNode, out: &mut Vec<Item>) {
        match n.kind() {
            IMPORT_DECL => out.push(Item::Import(self.import(n))),
            FN_DECL => out.push(Item::Fn(self.fn_decl(n))),
            CLASS_DECL => out.push(Item::Class(self.class_decl(n))),
            INTERFACE_DECL => out.push(Item::Interface(self.interface_decl(n))),
            ENUM_DECL => out.push(Item::Enum(self.enum_decl(n))),
            TYPE_ALIAS => {
                let ty = self.first_type(n);
                out.push(Item::TypeAlias(TypeAlias {
                    doc: doc_comment(n),
                    exported: has_tok(n, EXPORT_KW),
                    name: self.name_or_missing(n),
                    type_params: self.type_params(n),
                    ty,
                    span: self.nspan(n),
                }))
            }
            DECLARE_MODULE => {
                let lib_tok = sig_tokens(n).find(|t| t.kind() == STRING_LIT);
                let (lib, lib_span) = match lib_tok {
                    Some(t) => (string_value(t.text()), self.span(t.text_range())),
                    None => (String::new(), self.nspan(n)),
                };
                let functions = n.children().filter(|c| c.kind() == FN_DECL).map(|c| self.fn_decl(&c)).collect();
                out.push(Item::DeclareModule(DeclareModule { lib, lib_span, functions, span: self.nspan(n) }))
            }
            STRUCT_DECL => {
                let fields = n.children().filter(|c| c.kind() == PROP_SIG).map(|c| self.prop_sig(&c)).collect();
                out.push(Item::Struct(StructDecl {
                    doc: doc_comment(n),
                    exported: has_tok(n, EXPORT_KW),
                    name: self.name_or_missing(n),
                    fields,
                    span: self.nspan(n),
                }))
            }
            NAMESPACE_DECL => {
                let mut items = Vec::new();
                for c in n.children() {
                    if c.kind() != NAME {
                        self.item(&c, &mut items);
                    }
                }
                out.push(Item::Namespace(NamespaceDecl { name: self.name_or_missing(n), items, span: self.nspan(n) }))
            }
            VAR_DECL => {
                for v in self.var_decls(n) {
                    out.push(Item::Var(v));
                }
            }
            ERROR => {}
            k if k.is_stmt() => {
                let mut stmts = Vec::new();
                self.stmts(n, &mut stmts);
                out.extend(stmts.into_iter().map(Item::Stmt));
            }
            _ => {}
        }
    }

    fn import(&mut self, n: &SyntaxNode) -> Import {
        let mut names = Vec::new();
        for spec in n.children().filter(|c| c.kind() == IMPORT_SPEC) {
            let mut idents = spec.children().filter(|c| c.kind() == NAME).filter_map(|c| sig_tokens(&c).next());
            if let Some(first) = idents.next() {
                let alias = idents.next().map(|t| self.ident_tok(&t));
                names.push(ImportName { name: self.ident_tok(&first), alias });
            }
        }
        let namespace = child(n, IMPORT_NAMESPACE).and_then(|ns| self.name_of(&ns));
        let src_tok = sig_tokens(n).find(|t| t.kind() == STRING_LIT);
        let (source, source_span) = match src_tok {
            Some(t) => (string_value(t.text()), self.span(t.text_range())),
            None => (String::new(), self.nspan(n)),
        };
        Import { names, namespace, source, source_span, span: self.nspan(n) }
    }

    fn decorators(&mut self, n: &SyntaxNode) -> Vec<Decorator> {
        let mut out = Vec::new();
        for d in n.children().filter(|c| c.kind() == DECORATOR) {
            let toks: Vec<SyntaxToken> = sig_tokens(&d).filter(|t| t.kind() != AT && t.kind() != DOT).collect();
            let Some(first) = toks.first() else { continue };
            let text = toks.iter().map(|t| t.text().to_string()).collect::<Vec<_>>().join(".");
            let span = self.span(TextRange::new(first.text_range().start(), toks.last().unwrap().text_range().end()));
            let args = child(&d, ARG_LIST).map(|a| self.exprs(&a)).unwrap_or_default();
            out.push(Decorator { name: Ident { text, span }, args, span: self.nspan(&d) });
        }
        out
    }

    fn type_params(&mut self, n: &SyntaxNode) -> Vec<TypeParam> {
        let Some(tps) = child(n, TYPE_PARAMS) else { return Vec::new() };
        tps.children()
            .filter(|c| c.kind() == TYPE_PARAM)
            .map(|tp| {
                let has_extends = has_tok(&tp, EXTENDS_KW);
                let has_default = has_tok(&tp, EQ);
                let mut types = tp.children().filter(|c| c.kind().is_type());
                let constraint = if has_extends { types.next().map(|t| self.ty(&t)) } else { None };
                let default = if has_default { types.next().map(|t| self.ty(&t)) } else { None };
                TypeParam { name: self.name_or_missing(&tp), constraint, default }
            })
            .collect()
    }

    fn fn_decl(&mut self, n: &SyntaxNode) -> FnDecl {
        let inline = if has_tok(n, NOINLINE_KW) {
            InlineMode::NoInline
        } else if has_tok(n, INLINE_KW) {
            InlineMode::Inline
        } else {
            InlineMode::Default
        };
        let name = match n.kind() {
            CTOR_DECL => {
                let tok = sig_tokens(n).find(|t| t.kind() == IDENT && t.text() == "constructor");
                match tok {
                    Some(t) => self.ident_tok(&t),
                    None => self.name_or_missing(n),
                }
            }
            _ => self.name_or_missing(n),
        };
        let params = child(n, PARAM_LIST).map(|p| self.params(&p)).unwrap_or_default();
        let ret = n.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t));
        let body = child(n, BLOCK).map(|b| self.block(&b));
        FnDecl {
            decorators: self.decorators(n),
            doc: doc_comment(n),
            exported: has_tok(n, EXPORT_KW),
            declared: has_word(n, "declare"),
            inline,
            name,
            type_params: self.type_params(n),
            params,
            ret,
            body,
            span: self.nspan(n),
        }
    }

    fn params(&mut self, plist: &SyntaxNode) -> Vec<Param> {
        plist.children().filter(|c| c.kind() == PARAM).map(|p| self.param(&p)).collect()
    }

    fn param(&mut self, p: &SyntaxNode) -> Param {
        let this_tok = sig_tokens(p).find(|t| t.kind() == THIS_KW);
        let name = match &this_tok {
            Some(t) => self.ident_tok(t),
            None => self.name_or_missing(p),
        };
        let ty = p.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t));
        let default = p.children().find(|c| c.kind().is_expr()).map(|e| self.expr(&e));
        Param {
            name,
            ty,
            default,
            mode: if has_tok(p, BYREF_KW) { PassMode::ByRef } else { PassMode::ByVal },
            optional: has_tok(p, QUESTION),
            is_this: this_tok.is_some(),
            rest: has_tok(p, DOT3),
            span: self.nspan(p),
        }
    }

    fn member_modifiers(&self, n: &SyntaxNode) -> (Visibility, bool, bool) {
        let mut vis = Visibility::Public;
        let mut is_static = false;
        let mut readonly = false;
        for t in sig_tokens(n) {
            if t.kind() != IDENT {
                continue;
            }
            match t.text() {
                "public" => vis = Visibility::Public,
                "private" => vis = Visibility::Private,
                "protected" => vis = Visibility::Protected,
                "static" => is_static = true,
                "readonly" => readonly = true,
                _ => {}
            }
        }
        (vis, is_static, readonly)
    }

    fn class_decl(&mut self, n: &SyntaxNode) -> ClassDecl {
        let extends = child(n, EXTENDS_CLAUSE).and_then(|c| c.children().find(|t| t.kind().is_type())).map(|t| self.ty(&t));
        let implements = child(n, IMPLEMENTS_CLAUSE)
            .map(|c| c.children().filter(|t| t.kind().is_type()).map(|t| self.ty(&t)).collect())
            .unwrap_or_default();
        let mut members = Vec::new();
        for m in n.children() {
            match m.kind() {
                FIELD_DECL => {
                    let (vis, is_static, readonly) = self.member_modifiers(&m);
                    members.push(ClassMember::Field(FieldDecl {
                        doc: doc_comment(&m),
                        vis,
                        is_static,
                        readonly,
                        name: self.name_or_missing(&m),
                        ty: m.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t)),
                        init: m.children().find(|c| c.kind().is_expr()).map(|e| self.expr(&e)),
                        span: self.nspan(&m),
                    }));
                }
                METHOD_DECL => {
                    let (vis, is_static, _) = self.member_modifiers(&m);
                    members.push(ClassMember::Method(MethodDecl { vis, is_static, func: self.fn_decl(&m) }));
                }
                CTOR_DECL => members.push(ClassMember::Ctor(self.fn_decl(&m))),
                ACCESSOR_DECL => {
                    let (vis, is_static, _) = self.member_modifiers(&m);
                    // The accessor keyword is the last bare identifier before the NAME node.
                    let kind = if sig_tokens(&m).any(|t| t.kind() == IDENT && t.text() == "set") {
                        AccessorKind::Set
                    } else {
                        AccessorKind::Get
                    };
                    members.push(ClassMember::Accessor(AccessorDecl { kind, vis, is_static, func: self.fn_decl(&m) }));
                }
                _ => {}
            }
        }
        ClassDecl {
            decorators: self.decorators(n),
            doc: doc_comment(n),
            exported: has_tok(n, EXPORT_KW),
            declared: has_word(n, "declare"),
            name: self.name_or_missing(n),
            type_params: self.type_params(n),
            extends,
            implements,
            members,
            span: self.nspan(n),
        }
    }

    fn prop_sig(&mut self, m: &SyntaxNode) -> PropSig {
        PropSig {
            doc: doc_comment(m),
            name: self.name_or_missing(m),
            ty: m.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t)),
            readonly: has_word(m, "readonly"),
            optional: has_tok(m, QUESTION),
            span: self.nspan(m),
        }
    }

    fn interface_decl(&mut self, n: &SyntaxNode) -> InterfaceDecl {
        let extends = child(n, EXTENDS_CLAUSE)
            .map(|c| c.children().filter(|t| t.kind().is_type()).map(|t| self.ty(&t)).collect())
            .unwrap_or_default();
        let mut members = Vec::new();
        for m in n.children() {
            match m.kind() {
                PROP_SIG => members.push(InterfaceMember::Prop(self.prop_sig(&m))),
                METHOD_SIG => members.push(InterfaceMember::Method(self.fn_decl(&m))),
                INDEX_SIG => {
                    let param = match child(&m, PARAM) {
                        Some(p) => self.param(&p),
                        None => continue,
                    };
                    let ty = self.first_type(&m);
                    members.push(InterfaceMember::Index { param, ty, span: self.nspan(&m) });
                }
                _ => {}
            }
        }
        InterfaceDecl {
            doc: doc_comment(n),
            exported: has_tok(n, EXPORT_KW),
            declared: has_word(n, "declare"),
            name: self.name_or_missing(n),
            type_params: self.type_params(n),
            extends,
            members,
            span: self.nspan(n),
        }
    }

    fn enum_decl(&mut self, n: &SyntaxNode) -> EnumDecl {
        let members = n
            .children()
            .filter(|c| c.kind() == ENUM_MEMBER)
            .map(|m| EnumMember {
                doc: doc_comment(&m),
                name: self.name_or_missing(&m),
                value: m.children().find(|c| c.kind().is_expr()).map(|e| self.expr(&e)),
            })
            .collect();
        EnumDecl {
            doc: doc_comment(n),
            exported: has_tok(n, EXPORT_KW),
            declared: has_word(n, "declare"),
            name: self.name_or_missing(n),
            members,
            span: self.nspan(n),
        }
    }

    fn var_decls(&mut self, n: &SyntaxNode) -> Vec<VarDecl> {
        let kind = if has_tok(n, CONST_KW) {
            VarKind::Const
        } else if sig_tokens(n).any(|t| t.kind() == IDENT && t.text() == "using") {
            VarKind::Using
        } else {
            VarKind::Let
        };
        let exported = has_tok(n, EXPORT_KW);
        let declared = has_word(n, "declare");
        let doc = doc_comment(n);
        let mut out = Vec::new();
        for d in n.children().filter(|c| c.kind() == VAR_DECLARATOR) {
            let ty = d.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t));
            let init = d.children().find(|c| c.kind().is_expr()).map(|e| self.expr(&e));
            let span = self.nspan(&d);
            if let Some(pat) = d.children().find(|c| matches!(c.kind(), OBJECT_PATTERN | ARRAY_PATTERN)) {
                if kind == VarKind::Using || declared {
                    self.err(self.nspan(&pat), "ここでは分割代入を使えません");
                }
                let Some(init) = init else {
                    self.err(self.nspan(&pat), "分割代入には初期化式が必要です");
                    continue;
                };
                // `const { a, b } = e;` → `const t = e; const a = t.a; const b = t.b;`
                // (a plain variable on the right is read directly, without the temporary).
                let source = match (&init.peel().kind, &ty) {
                    (ExprKind::Name(id), None) => id.clone(),
                    _ => {
                        let tmp = self.pattern_temp(&pat);
                        out.push(VarDecl { doc: None, kind: VarKind::Const, exported: false, declared: false, name: tmp.clone(), ty, init: Some(init), span });
                        tmp
                    }
                };
                for (name, init) in self.pattern_bindings(&pat, &source) {
                    out.push(VarDecl { doc: doc.clone(), kind, exported, declared: false, name, ty: None, init: Some(init), span });
                }
                continue;
            }
            out.push(VarDecl { doc: doc.clone(), kind, exported, declared, name: self.name_or_missing(&d), ty, init, span });
        }
        out
    }

    /// Hidden variable that holds the value being destructured. Its span is the opening
    /// bracket of the pattern, which no real identifier can share.
    fn pattern_temp(&mut self, pat: &SyntaxNode) -> Ident {
        self.pattern_n += 1;
        let start = self.nspan(pat).start();
        Ident { text: format!("ras$d{}", self.pattern_n), span: TextRange::at(start, TextSize::from(1)) }
    }

    /// The variables a pattern declares, each with the expression that reads it from `source`.
    fn pattern_bindings(&mut self, pat: &SyntaxNode, source: &Ident) -> Vec<(Ident, Expr)> {
        let object = pat.kind() == OBJECT_PATTERN;
        let mut out = Vec::new();
        for (i, el) in pat.children().filter(|c| c.kind() == PATTERN_ELEM).enumerate() {
            let names: Vec<Ident> = el.children().filter(|c| c.kind() == NAME).filter_map(|c| sig_tokens(&c).next()).map(|t| self.ident_tok(&t)).collect();
            let (Some(key), Some(binding)) = (names.first(), names.last()) else { continue };
            let span = self.nspan(&el);
            let object_expr = Box::new(Expr { id: self.id(), kind: ExprKind::Name(Ident { text: source.text.clone(), span }), span });
            let kind = if object {
                ExprKind::Member { object: object_expr, name: Ident { text: key.text.clone(), span: key.span }, optional: false }
            } else {
                let index = Box::new(Expr { id: self.id(), kind: ExprKind::Int { value: i as i128, suffix: NumSuffix::None, hex: false }, span });
                ExprKind::Index { object: object_expr, index }
            };
            out.push((binding.clone(), Expr { id: self.id(), kind, span }));
        }
        out
    }

    // ------------------------------------------------------------------ types

    fn first_type(&mut self, n: &SyntaxNode) -> TypeRef {
        match n.children().find(|c| c.kind().is_type()) {
            Some(t) => self.ty(&t),
            None => TypeRef::Error { span: TextRange::empty(self.nspan(n).end()) },
        }
    }

    fn ty(&mut self, n: &SyntaxNode) -> TypeRef {
        let span = self.nspan(n);
        match n.kind() {
            NAMED_TYPE => {
                let path: Vec<Ident> =
                    sig_tokens(n).filter(|t| t.kind() == IDENT || t.kind().is_keyword()).map(|t| self.ident_tok(&t)).collect();
                if path.is_empty() {
                    return TypeRef::Error { span };
                }
                let args = child(n, TYPE_ARGS)
                    .map(|a| a.children().filter(|c| c.kind().is_type()).map(|t| self.ty(&t)).collect())
                    .unwrap_or_default();
                TypeRef::Named { path, args, span }
            }
            ARRAY_TYPE => TypeRef::Array { elem: Box::new(self.first_type(n)), span },
            FN_TYPE => {
                let params = child(n, PARAM_LIST).map(|p| self.params(&p)).unwrap_or_default();
                TypeRef::Fn { params, ret: Box::new(self.first_type(n)), span }
            }
            PAREN_TYPE => self.first_type(n),
            TUPLE_TYPE => TypeRef::Tuple { elems: n.children().filter(|c| c.kind().is_type()).map(|t| self.ty(&t)).collect(), span },
            OBJECT_TYPE => TypeRef::Object { fields: n.children().filter(|c| c.kind() == PROP_SIG).map(|c| self.prop_sig(&c)).collect(), span },
            UNION_TYPE => {
                let mut nullable = false;
                let mut rest = Vec::new();
                for c in n.children().filter(|c| c.kind().is_type()) {
                    let is_null = c.kind() == NAMED_TYPE
                        && sig_tokens(&c).next().is_some_and(|t| matches!(t.kind(), NULL_KW | UNDEFINED_KW));
                    if is_null {
                        nullable = true;
                    } else {
                        rest.push(c);
                    }
                }
                match rest.len() {
                    // `null | undefined`
                    0 => self.first_type(n),
                    1 => {
                        let inner = self.ty(&rest[0]);
                        if nullable {
                            TypeRef::Nullable { inner: Box::new(inner), span }
                        } else {
                            inner
                        }
                    }
                    _ => TypeRef::Union { elems: rest.iter().map(|c| self.ty(c)).collect(), nullable, span },
                }
            }
            _ => TypeRef::Error { span },
        }
    }

    // ------------------------------------------------------------------ statements

    fn block(&mut self, n: &SyntaxNode) -> Block {
        let mut stmts = Vec::new();
        for c in n.children() {
            self.stmts(&c, &mut stmts);
        }
        Block { stmts, span: self.nspan(n) }
    }

    /// Lowers the body statement of `if` / loops: always exactly one statement.
    fn body_stmt(&mut self, parent: &SyntaxNode, skip: usize) -> Stmt {
        match parent.children().filter(|c| c.kind().is_stmt()).nth(skip) {
            Some(c) => {
                let mut out = Vec::new();
                self.stmts(&c, &mut out);
                if out.len() == 1 {
                    out.pop().unwrap()
                } else {
                    let span = self.nspan(&c);
                    Stmt { kind: StmtKind::Block(Block { stmts: out, span }), span }
                }
            }
            None => Stmt { kind: StmtKind::Error, span: TextRange::empty(self.nspan(parent).end()) },
        }
    }

    fn first_expr(&mut self, n: &SyntaxNode) -> Expr {
        match n.children().find(|c| c.kind().is_expr()) {
            Some(e) => self.expr(&e),
            None => self.error_expr(TextRange::empty(self.nspan(n).end())),
        }
    }

    fn error_expr(&mut self, span: Span) -> Expr {
        Expr { id: self.id(), kind: ExprKind::Error, span }
    }

    fn exprs(&mut self, n: &SyntaxNode) -> Vec<Expr> {
        n.children().filter(|c| c.kind().is_expr()).map(|e| self.expr(&e)).collect()
    }

    fn stmts(&mut self, n: &SyntaxNode, out: &mut Vec<Stmt>) {
        let span = self.nspan(n);
        let kind = match n.kind() {
            VAR_DECL => {
                for v in self.var_decls(n) {
                    let span = v.span;
                    out.push(Stmt { kind: StmtKind::Var(v), span });
                }
                return;
            }
            EXPR_STMT => match n.children().find(|c| c.kind().is_expr()) {
                Some(e) => StmtKind::Expr(self.expr(&e)),
                None => StmtKind::Error,
            },
            BLOCK => StmtKind::Block(self.block(n)),
            IF_STMT => {
                let cond = self.first_expr(n);
                let then = Box::new(self.body_stmt(n, 0));
                let els = child(n, ELSE_CLAUSE).map(|e| Box::new(self.body_stmt(&e, 0)));
                StmtKind::If { cond, then, els }
            }
            FOR_STMT => {
                let mut init = Vec::new();
                if let Some(fi) = child(n, FOR_INIT) {
                    for c in fi.children() {
                        if c.kind() == VAR_DECL {
                            self.stmts(&c, &mut init);
                        } else if c.kind().is_expr() {
                            let e = self.expr(&c);
                            let span = e.span;
                            init.push(Stmt { kind: StmtKind::Expr(e), span });
                        }
                    }
                }
                let cond = child(n, FOR_COND).and_then(|c| c.children().find(|e| e.kind().is_expr())).map(|e| self.expr(&e));
                let update = child(n, FOR_UPDATE).map(|c| self.exprs(&c)).unwrap_or_default();
                StmtKind::For { decorators: self.decorators(n), init, cond, update, body: Box::new(self.body_stmt(n, 0)) }
            }
            FOR_OF_STMT => {
                let var_kind = if has_tok(n, CONST_KW) { VarKind::Const } else { VarKind::Let };
                let ty = n.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t));
                let iter = self.first_expr(n);
                let mut body = self.body_stmt(n, 0);
                let name = match n.children().find(|c| matches!(c.kind(), OBJECT_PATTERN | ARRAY_PATTERN)) {
                    // `for (const { a, b } of xs) body` → `for (const t of xs) { const a = t.a; const b = t.b; body }`
                    Some(pat) => {
                        let tmp = self.pattern_temp(&pat);
                        let pspan = self.nspan(&pat);
                        let mut stmts: Vec<Stmt> = self
                            .pattern_bindings(&pat, &tmp)
                            .into_iter()
                            .map(|(name, init)| Stmt {
                                kind: StmtKind::Var(VarDecl { doc: None, kind: var_kind, exported: false, declared: false, name, ty: None, init: Some(init), span: pspan }),
                                span: pspan,
                            })
                            .collect();
                        let bspan = body.span;
                        stmts.push(body);
                        body = Stmt { kind: StmtKind::Block(Block { stmts, span: bspan }), span: bspan };
                        tmp
                    }
                    None => self.name_or_missing(n),
                };
                StmtKind::ForOf { decorators: self.decorators(n), kind: var_kind, name, ty, iter, body: Box::new(body) }
            }
            WHILE_STMT => StmtKind::While {
                decorators: self.decorators(n),
                cond: self.first_expr(n),
                body: Box::new(self.body_stmt(n, 0)),
            },
            DO_WHILE_STMT => StmtKind::DoWhile {
                decorators: self.decorators(n),
                body: Box::new(self.body_stmt(n, 0)),
                cond: self.first_expr(n),
            },
            SWITCH_STMT => {
                let disc = self.first_expr(n);
                let mut cases = Vec::new();
                for c in n.children().filter(|c| c.kind() == CASE_CLAUSE) {
                    let test = if has_tok(&c, CASE_KW) { Some(self.first_expr(&c)) } else { None };
                    let mut body = Vec::new();
                    for s in c.children().filter(|s| s.kind().is_stmt()) {
                        self.stmts(&s, &mut body);
                    }
                    cases.push(SwitchCase { test, body, span: self.nspan(&c) });
                }
                StmtKind::Switch { disc, cases }
            }
            TRY_STMT => {
                let block = match child(n, BLOCK) {
                    Some(b) => self.block(&b),
                    None => Block { stmts: Vec::new(), span },
                };
                let catch = child(n, CATCH_CLAUSE).map(|c| CatchClause {
                    param: self.name_of(&c),
                    body: match child(&c, BLOCK) {
                        Some(b) => self.block(&b),
                        None => Block { stmts: Vec::new(), span: self.nspan(&c) },
                    },
                    span: self.nspan(&c),
                });
                let finally = child(n, FINALLY_CLAUSE).and_then(|c| child(&c, BLOCK)).map(|b| self.block(&b));
                StmtKind::Try(TryStmt { block, catch, finally })
            }
            THROW_STMT => StmtKind::Throw(self.first_expr(n)),
            RETURN_STMT => StmtKind::Return(n.children().find(|c| c.kind().is_expr()).map(|e| self.expr(&e))),
            BREAK_STMT => StmtKind::Break(sig_tokens(n).find(|t| t.kind() == IDENT).map(|t| self.ident_tok(&t))),
            CONTINUE_STMT => StmtKind::Continue(sig_tokens(n).find(|t| t.kind() == IDENT).map(|t| self.ident_tok(&t))),
            LABELED_STMT => match sig_tokens(n).find(|t| t.kind() == IDENT) {
                Some(t) => StmtKind::Labeled { label: self.ident_tok(&t), body: Box::new(self.body_stmt(n, 0)) },
                None => StmtKind::Error,
            },
            RESUME_STMT => StmtKind::Resume,
            EMPTY_STMT => StmtKind::Empty,
            FN_DECL => StmtKind::Fn(Box::new(self.fn_decl(n))),
            _ => return,
        };
        out.push(Stmt { kind, span });
    }

    // ------------------------------------------------------------------ expressions

    pub(crate) fn expr(&mut self, n: &SyntaxNode) -> Expr {
        let span = self.nspan(n);
        let id = self.id();
        let kind = self.expr_kind(n, span);
        Expr { id, kind, span }
    }

    fn sub_exprs(&mut self, n: &SyntaxNode) -> Vec<SyntaxNode> {
        n.children().filter(|c| c.kind().is_expr()).collect()
    }

    fn boxed(&mut self, nodes: &[SyntaxNode], i: usize, parent_span: Span) -> Box<Expr> {
        Box::new(match nodes.get(i) {
            Some(n) => self.expr(n),
            None => self.error_expr(TextRange::empty(parent_span.end())),
        })
    }

    fn expr_kind(&mut self, n: &SyntaxNode, span: Span) -> ExprKind {
        match n.kind() {
            LITERAL => {
                let Some(tok) = sig_tokens(n).next() else { return ExprKind::Error };
                let text = tok.text();
                match tok.kind() {
                    INT_LIT => self.int_literal(text, span),
                    FLOAT_LIT => float_literal(text),
                    STRING_LIT => ExprKind::Str(string_value(text)),
                    TEMPLATE_LIT => self.template(text, span),
                    TRUE_KW => ExprKind::Bool(true),
                    FALSE_KW => ExprKind::Bool(false),
                    NULL_KW => ExprKind::Null,
                    UNDEFINED_KW => ExprKind::Undefined,
                    _ => ExprKind::Error,
                }
            }
            NAME_REF => match sig_tokens(n).next() {
                Some(t) => ExprKind::Name(self.ident_tok(&t)),
                None => ExprKind::Error,
            },
            THIS_EXPR => ExprKind::This,
            SUPER_EXPR => ExprKind::Super,
            PAREN_EXPR => {
                let subs = self.sub_exprs(n);
                ExprKind::Paren(self.boxed(&subs, 0, span))
            }
            ARRAY_EXPR => ExprKind::Array(self.exprs(n)),
            OBJECT_EXPR => {
                let mut props = Vec::new();
                for p in n.children().filter(|c| c.kind() == OBJECT_PROP) {
                    let Some(key_tok) = sig_tokens(&p).next() else { continue };
                    let key = if key_tok.kind() == STRING_LIT {
                        Ident { text: string_value(key_tok.text()), span: self.span(key_tok.text_range()) }
                    } else {
                        self.ident_tok(&key_tok)
                    };
                    let value = match p.children().find(|c| c.kind().is_expr()) {
                        Some(e) => self.expr(&e),
                        None => Expr { id: self.id(), kind: ExprKind::Name(key.clone()), span: key.span },
                    };
                    props.push(ObjectProp { key, value });
                }
                ExprKind::Object(props)
            }
            ARROW_EXPR => {
                let params = child(n, PARAM_LIST).map(|p| self.params(&p)).unwrap_or_default();
                let ret = n.children().find(|c| c.kind().is_type()).map(|t| self.ty(&t));
                let body = if let Some(b) = child(n, BLOCK) {
                    ArrowBody::Block(self.block(&b))
                } else {
                    ArrowBody::Expr(self.first_expr(n))
                };
                ExprKind::Arrow(Box::new(ArrowFn { params, ret, body, span }))
            }
            UNARY_EXPR => {
                let op = sig_tokens(n).next().map(|t| t.kind());
                let subs = self.sub_exprs(n);
                let operand = self.boxed(&subs, 0, span);
                match op {
                    Some(BANG) => ExprKind::Unary { op: UnaryOp::Not, operand },
                    Some(MINUS) => ExprKind::Unary { op: UnaryOp::Neg, operand },
                    Some(PLUS) => ExprKind::Unary { op: UnaryOp::Plus, operand },
                    Some(TILDE) => ExprKind::Unary { op: UnaryOp::BitNot, operand },
                    Some(TYPEOF_KW) => ExprKind::Unary { op: UnaryOp::TypeOf, operand },
                    Some(PLUS2) => ExprKind::IncDec { inc: true, prefix: true, operand },
                    Some(MINUS2) => ExprKind::IncDec { inc: false, prefix: true, operand },
                    _ => ExprKind::Error,
                }
            }
            POSTFIX_EXPR => {
                let inc = has_tok(n, PLUS2);
                let subs = self.sub_exprs(n);
                ExprKind::IncDec { inc, prefix: false, operand: self.boxed(&subs, 0, span) }
            }
            NON_NULL_EXPR => {
                let subs = self.sub_exprs(n);
                ExprKind::NonNull(self.boxed(&subs, 0, span))
            }
            BINARY_EXPR => {
                let op: String = sig_tokens(n).map(|t| t.text().to_string()).collect();
                let subs = self.sub_exprs(n);
                if op == "instanceof" {
                    let expr = self.boxed(&subs, 0, span);
                    return ExprKind::InstanceOf { expr, ty: self.first_type(n) };
                }
                let lhs = self.boxed(&subs, 0, span);
                let rhs = self.boxed(&subs, 1, span);
                let op = match op.as_str() {
                    "+" => BinaryOp::Add,
                    "-" => BinaryOp::Sub,
                    "*" => BinaryOp::Mul,
                    "/" => BinaryOp::Div,
                    "%" => BinaryOp::Rem,
                    "**" => BinaryOp::Pow,
                    "==" | "===" => BinaryOp::Eq,
                    "!=" | "!==" => BinaryOp::Ne,
                    "<" => BinaryOp::Lt,
                    "<=" => BinaryOp::Le,
                    ">" => BinaryOp::Gt,
                    ">=" => BinaryOp::Ge,
                    "&&" => BinaryOp::And,
                    "||" => BinaryOp::Or,
                    "??" => BinaryOp::Coalesce,
                    "&" => BinaryOp::BitAnd,
                    "|" => BinaryOp::BitOr,
                    "^" => BinaryOp::BitXor,
                    "<<" => BinaryOp::Shl,
                    ">>" => BinaryOp::Shr,
                    other => {
                        self.err(span, format!("演算子 `{other}` には対応していません"));
                        return ExprKind::Error;
                    }
                };
                ExprKind::Binary { op, lhs, rhs }
            }
            ASSIGN_EXPR => {
                let op: String = sig_tokens(n).map(|t| t.text().to_string()).collect();
                let subs = self.sub_exprs(n);
                let target = self.boxed(&subs, 0, span);
                let value = self.boxed(&subs, 1, span);
                let op = match op.as_str() {
                    "=" => None,
                    "+=" => Some(BinaryOp::Add),
                    "-=" => Some(BinaryOp::Sub),
                    "*=" => Some(BinaryOp::Mul),
                    "/=" => Some(BinaryOp::Div),
                    "%=" => Some(BinaryOp::Rem),
                    "&=" => Some(BinaryOp::BitAnd),
                    "|=" => Some(BinaryOp::BitOr),
                    "^=" => Some(BinaryOp::BitXor),
                    "<<=" => Some(BinaryOp::Shl),
                    ">>=" => Some(BinaryOp::Shr),
                    _ => None,
                };
                ExprKind::Assign { op, target, value }
            }
            TERNARY_EXPR => {
                let subs = self.sub_exprs(n);
                ExprKind::Ternary {
                    cond: self.boxed(&subs, 0, span),
                    then: self.boxed(&subs, 1, span),
                    els: self.boxed(&subs, 2, span),
                }
            }
            CALL_EXPR => {
                let subs = self.sub_exprs(n);
                let callee = self.boxed(&subs, 0, span);
                let type_args = child(n, TYPE_ARGS)
                    .map(|a| a.children().filter(|c| c.kind().is_type()).map(|t| self.ty(&t)).collect())
                    .unwrap_or_default();
                let args = child(n, ARG_LIST).map(|a| self.exprs(&a)).unwrap_or_default();
                ExprKind::Call { callee, type_args, args }
            }
            NAMED_ARG => {
                let name = self.name_or_missing(n);
                let subs = self.sub_exprs(n);
                let value = self.boxed(&subs, 0, span);
                ExprKind::NamedArg { name, value, hole: self.id() }
            }
            NEW_EXPR => {
                let ty = self.first_type(n);
                let args = child(n, ARG_LIST).map(|a| self.exprs(&a)).unwrap_or_default();
                ExprKind::New { ty, args }
            }
            MEMBER_EXPR => {
                let subs = self.sub_exprs(n);
                let object = self.boxed(&subs, 0, span);
                let optional = has_tok(n, QUESTION_DOT);
                match sig_tokens(n).filter(|t| t.kind() == IDENT || t.kind().is_keyword()).last() {
                    Some(t) => ExprKind::Member { object, name: self.ident_tok(&t), optional },
                    None => ExprKind::Member {
                        object,
                        name: Ident { text: String::new(), span: TextRange::empty(span.end()) },
                        optional,
                    },
                }
            }
            INDEX_EXPR => {
                let subs = self.sub_exprs(n);
                ExprKind::Index { object: self.boxed(&subs, 0, span), index: self.boxed(&subs, 1, span) }
            }
            AS_EXPR => {
                let subs = self.sub_exprs(n);
                ExprKind::As { expr: self.boxed(&subs, 0, span), ty: self.first_type(n) }
            }
            _ => ExprKind::Error,
        }
    }

    fn int_literal(&mut self, text: &str, span: Span) -> ExprKind {
        let mut digits: String = text.chars().filter(|c| *c != '_').collect();
        let mut suffix = NumSuffix::None;
        let is_hex = digits.starts_with("0x") || digits.starts_with("0X");
        let is_bin = digits.starts_with("0b") || digits.starts_with("0B");
        if digits.ends_with(['l', 'L']) {
            suffix = NumSuffix::L;
            digits.pop();
        }
        let parsed = if is_hex {
            i128::from_str_radix(&digits[2..], 16)
        } else if is_bin {
            i128::from_str_radix(&digits[2..], 2)
        } else {
            digits.parse::<i128>()
        };
        match parsed {
            Ok(value) => ExprKind::Int { value, suffix, hex: is_hex || is_bin },
            Err(_) => {
                self.err(span, "整数リテラルが大きすぎます");
                ExprKind::Int { value: 0, suffix, hex: is_hex }
            }
        }
    }

    fn template(&mut self, raw: &str, span: Span) -> ExprKind {
        // raw includes the back-ticks
        let body = &raw[1..raw.len() - 1];
        let body_start = u32::from(span.start()) + 1;
        let mut parts = Vec::new();
        let mut chunk = String::new();
        let bytes = body.as_bytes();
        let mut i = 0;
        while i < bytes.len() {
            if bytes[i] == b'\\' && i + 1 < bytes.len() {
                // keep the escape for `unescape`
                let ch_len = body[i + 1..].chars().next().map_or(1, |c| c.len_utf8());
                chunk.push_str(&body[i..i + 1 + ch_len]);
                i += 1 + ch_len;
                continue;
            }
            if bytes[i] == b'$' && i + 1 < bytes.len() && bytes[i + 1] == b'{' {
                // find the matching brace
                let mut depth = 1;
                let mut j = i + 2;
                while j < bytes.len() && depth > 0 {
                    match bytes[j] {
                        b'{' => depth += 1,
                        b'}' => depth -= 1,
                        _ => {}
                    }
                    j += 1;
                }
                if depth != 0 {
                    self.err(span, "テンプレート文字列の `${` が閉じられていません");
                    break;
                }
                if !chunk.is_empty() {
                    parts.push(TemplatePart::Str(unescape(&chunk)));
                    chunk.clear();
                }
                let inner = &body[i + 2..j - 1];
                let inner_base = body_start + (i as u32) + 2;
                parts.push(TemplatePart::Expr(self.fragment(inner, inner_base)));
                i = j;
                continue;
            }
            let ch_len = body[i..].chars().next().map_or(1, |c| c.len_utf8());
            chunk.push_str(&body[i..i + ch_len]);
            i += ch_len;
        }
        if !chunk.is_empty() {
            parts.push(TemplatePart::Str(unescape(&chunk)));
        }
        ExprKind::Template(parts)
    }

    /// Parses and lowers an expression embedded in a template literal.
    fn fragment(&mut self, text: &str, base: u32) -> Expr {
        let mut p = Parser::new(text);
        p.expr_file();
        let (green, errors) = p.finish_tree();
        for e in errors {
            self.errors.push(SyntaxError { message: e.message, range: e.range + TextSize::from(base) });
        }
        let root = SyntaxNode::new_root(green);
        let saved = self.base;
        self.base = base;
        let expr = match root.children().find(|c| c.kind().is_expr()) {
            Some(e) => self.expr(&e),
            None => self.error_expr(TextRange::empty(TextSize::from(0))),
        };
        self.base = saved;
        expr
    }
}

/// Value of a quoted string literal token (quotes removed, escapes decoded).
pub fn string_value(tok_text: &str) -> String {
    if tok_text.len() >= 2 {
        unescape(&tok_text[1..tok_text.len() - 1])
    } else {
        String::new()
    }
}

fn float_literal(text: &str) -> ExprKind {
    let mut digits: String = text.chars().filter(|c| *c != '_').collect();
    let suffix = match digits.chars().last() {
        Some('f' | 'F') => NumSuffix::F,
        Some('d' | 'D') => NumSuffix::D,
        Some('m' | 'M') => NumSuffix::M,
        _ => NumSuffix::None,
    };
    if suffix != NumSuffix::None {
        digits.pop();
    }
    ExprKind::Float { text: digits, suffix }
}
