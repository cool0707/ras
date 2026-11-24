//! Hand-written recursive-descent parser producing a lossless `rowan` tree.
//!
//! The parser never fails: on malformed input it records a `SyntaxError`,
//! inserts `ERROR` nodes and keeps going, so that the LSP always has a tree.

use rowan::{Checkpoint, GreenNode, GreenNodeBuilder, TextRange, TextSize};

use crate::kind::{lex, SyntaxKind, SyntaxKind::*, Token};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SyntaxError {
    pub message: String,
    pub range: TextRange,
}

pub(crate) struct Parser<'a> {
    text: &'a str,
    toks: Vec<Token>,
    /// Indices into `toks` of the significant (non-trivia) tokens.
    sig: Vec<usize>,
    spos: usize,
    /// Next raw token (including trivia) to be emitted into the tree.
    raw: usize,
    builder: GreenNodeBuilder<'static>,
    pub(crate) errors: Vec<SyntaxError>,
}

const MEMBER_MODIFIERS: &[&str] = &[
    "public",
    "private",
    "protected",
    "static",
    "readonly",
    "abstract",
    "override",
];

impl<'a> Parser<'a> {
    pub(crate) fn new(text: &'a str) -> Self {
        let toks = lex(text);
        let sig = toks
            .iter()
            .enumerate()
            .filter(|(_, t)| !t.kind.is_trivia())
            .map(|(i, _)| i)
            .collect();
        Parser {
            text,
            toks,
            sig,
            spos: 0,
            raw: 0,
            builder: GreenNodeBuilder::new(),
            errors: Vec::new(),
        }
    }

    pub(crate) fn finish_tree(self) -> (GreenNode, Vec<SyntaxError>) {
        (self.builder.finish(), self.errors)
    }

    // ------------------------------------------------------------------
    // token access
    // ------------------------------------------------------------------

    fn nth(&self, n: usize) -> SyntaxKind {
        self.sig
            .get(self.spos + n)
            .map(|&i| self.toks[i].kind)
            .unwrap_or(EOF)
    }

    fn nth_text(&self, n: usize) -> &'a str {
        match self.sig.get(self.spos + n) {
            Some(&i) => {
                let t = self.toks[i];
                &self.text[t.start as usize..t.end as usize]
            }
            None => "",
        }
    }

    fn at(&self, kind: SyntaxKind) -> bool {
        self.nth(0) == kind
    }

    fn nth_word(&self, n: usize, word: &str) -> bool {
        self.nth(n) == IDENT && self.nth_text(n) == word
    }

    fn at_word(&self, word: &str) -> bool {
        self.nth_word(0, word)
    }

    fn nth_is_name(&self, n: usize) -> bool {
        let k = self.nth(n);
        k == IDENT || k.is_keyword()
    }

    /// True when significant tokens `n` and `n + 1` touch (no trivia between).
    fn adjacent(&self, n: usize) -> bool {
        match (self.sig.get(self.spos + n), self.sig.get(self.spos + n + 1)) {
            (Some(&a), Some(&b)) => a + 1 == b,
            _ => false,
        }
    }

    /// True when a line break separates significant token `n` from the one before it.
    fn newline_before(&self, n: usize) -> bool {
        let Some(&cur) = self.sig.get(self.spos + n) else {
            return true;
        };
        let prev = if self.spos + n == 0 {
            0
        } else {
            self.sig[self.spos + n - 1] + 1
        };
        self.toks[prev..cur]
            .iter()
            .any(|t| self.text[t.start as usize..t.end as usize].contains('\n'))
    }

    fn eat_trivia(&mut self) {
        let end = self.sig.get(self.spos).copied().unwrap_or(self.toks.len());
        while self.raw < end {
            let t = self.toks[self.raw];
            self.builder
                .token(t.kind.into(), &self.text[t.start as usize..t.end as usize]);
            self.raw += 1;
        }
    }

    fn bump(&mut self) {
        self.eat_trivia();
        if let Some(&i) = self.sig.get(self.spos) {
            let t = self.toks[i];
            self.builder
                .token(t.kind.into(), &self.text[t.start as usize..t.end as usize]);
            self.raw = i + 1;
            self.spos += 1;
        }
    }

    fn start(&mut self, kind: SyntaxKind) {
        self.eat_trivia();
        self.builder.start_node(kind.into());
    }

    fn finish(&mut self) {
        self.builder.finish_node();
    }

    fn checkpoint(&mut self) -> Checkpoint {
        self.eat_trivia();
        self.builder.checkpoint()
    }

    fn start_at(&mut self, cp: Checkpoint, kind: SyntaxKind) {
        self.builder.start_node_at(cp, kind.into());
    }

    fn eat(&mut self, kind: SyntaxKind) -> bool {
        if self.at(kind) {
            self.bump();
            true
        } else {
            false
        }
    }

    fn current_range(&self) -> TextRange {
        match self.sig.get(self.spos) {
            Some(&i) => {
                let t = self.toks[i];
                TextRange::new(TextSize::from(t.start), TextSize::from(t.end))
            }
            None => {
                // Point just after the last significant token.
                let end = if self.spos == 0 {
                    0
                } else {
                    self.toks[self.sig[self.spos - 1]].end
                };
                TextRange::empty(TextSize::from(end))
            }
        }
    }

    fn error(&mut self, message: impl Into<String>) {
        let range = self.current_range();
        if self
            .errors
            .last()
            .is_some_and(|e| e.range.start() == range.start())
        {
            return;
        }
        self.errors.push(SyntaxError {
            message: message.into(),
            range,
        });
    }

    fn expect(&mut self, kind: SyntaxKind, what: &str) -> bool {
        if self.eat(kind) {
            true
        } else {
            self.error(format!("{what} が必要です"));
            false
        }
    }

    /// Consumes the current token wrapped in an `ERROR` node (guarantees progress).
    fn bump_error(&mut self) {
        if self.at(EOF) {
            return;
        }
        self.start(ERROR);
        self.bump();
        self.finish();
    }

    fn semi(&mut self) {
        if self.eat(SEMI) {
            return;
        }
        if self.at(R_BRACE) || self.at(EOF) || self.newline_before(0) {
            return;
        }
        self.error("`;` が必要です");
    }

    // ------------------------------------------------------------------
    // lookahead helpers (no tree building)
    // ------------------------------------------------------------------

    /// `n` points at an opening bracket; returns the index just after its match.
    fn skip_balanced(&self, mut n: usize) -> Option<usize> {
        let mut depth = 0usize;
        loop {
            match self.nth(n) {
                L_PAREN | L_BRACK | L_BRACE => depth += 1,
                R_PAREN | R_BRACK | R_BRACE => {
                    depth = depth.checked_sub(1)?;
                    if depth == 0 {
                        return Some(n + 1);
                    }
                }
                EOF => return None,
                _ => {}
            }
            n += 1;
        }
    }

    fn skip_decorators(&self, mut n: usize) -> usize {
        while self.nth(n) == AT {
            n += 1;
            if self.nth_is_name(n) {
                n += 1;
            }
            while self.nth(n) == DOT && self.nth(n + 1) == IDENT {
                n += 2;
            }
            if self.nth(n) == L_PAREN {
                match self.skip_balanced(n) {
                    Some(m) => n = m,
                    None => return n,
                }
            }
        }
        n
    }

    fn skip_type(&self, mut n: usize) -> Option<usize> {
        match self.nth(n) {
            L_PAREN => {
                n = self.skip_balanced(n)?;
                if self.nth(n) == FAT_ARROW {
                    n = self.skip_type(n + 1)?;
                }
            }
            IDENT | VOID_KW | NULL_KW | UNDEFINED_KW => {
                n += 1;
                while self.nth(n) == DOT && self.nth_is_name(n + 1) {
                    n += 2;
                }
                if self.nth(n) == LT {
                    n = self.skip_type_args(n)?;
                }
            }
            // tuple `[A, B]` / object type `{ a: A; b: B }`
            L_BRACK | L_BRACE => n = self.skip_balanced(n)?,
            _ => return None,
        }
        while self.nth(n) == L_BRACK && self.nth(n + 1) == R_BRACK {
            n += 2;
        }
        if self.nth(n) == PIPE {
            return self.skip_type(n + 1);
        }
        Some(n)
    }

    /// `n` points at `<`; returns the index just after the matching `>`.
    fn skip_type_args(&self, mut n: usize) -> Option<usize> {
        n += 1;
        loop {
            n = self.skip_type(n)?;
            match self.nth(n) {
                COMMA => n += 1,
                GT => return Some(n + 1),
                _ => return None,
            }
        }
    }

    // ------------------------------------------------------------------
    // items
    // ------------------------------------------------------------------

    pub(crate) fn source_file(&mut self) {
        self.builder.start_node(SOURCE_FILE.into());
        while !self.at(EOF) {
            self.item();
        }
        self.eat_trivia();
        self.builder.finish_node();
    }

    /// Parses a single expression as a whole file (used for template literals).
    pub(crate) fn expr_file(&mut self) {
        self.builder.start_node(SOURCE_FILE.into());
        self.expr();
        while !self.at(EOF) {
            self.error("式の終わりが必要です");
            self.bump_error();
        }
        self.eat_trivia();
        self.builder.finish_node();
    }

    fn item(&mut self) {
        let k0 = self.skip_decorators(0);
        let mut k = k0;
        if self.nth(k) == EXPORT_KW {
            k += 1;
        }
        let mut declared = false;
        if self.nth_word(k, "declare") && self.is_declarable(k + 1) {
            declared = true;
            k += 1;
        }
        match self.nth(k) {
            IMPORT_KW if k == 0 => self.import_decl(),
            FUNCTION_KW | INLINE_KW | NOINLINE_KW => self.fn_decl(),
            CLASS_KW => self.class_decl(),
            INTERFACE_KW => self.interface_decl(),
            ENUM_KW => self.enum_decl(),
            CONST_KW if self.nth(k + 1) == ENUM_KW => self.enum_decl(),
            LET_KW | CONST_KW | VAR_KW => self.var_decl(true),
            IDENT => {
                let word = self.nth_text(k);
                if word == "abstract" && self.nth(k + 1) == CLASS_KW {
                    self.class_decl();
                } else if word == "type"
                    && self.nth(k + 1) == IDENT
                    && matches!(self.nth(k + 2), EQ | LT)
                {
                    self.type_alias();
                } else if word == "module" && declared && self.nth(k + 1) == STRING_LIT {
                    self.declare_module();
                } else if word == "struct"
                    && self.nth(k + 1) == IDENT
                    && (declared || self.nth(k + 2) == L_BRACE)
                {
                    self.struct_decl();
                } else if word == "namespace"
                    && self.nth(k + 1) == IDENT
                    && self.nth(k + 2) == L_BRACE
                {
                    self.namespace_decl();
                } else if k > k0 {
                    self.bad_modifiers(k - k0, k0);
                } else {
                    self.stmt();
                }
            }
            _ if k > k0 => self.bad_modifiers(k - k0, k0),
            _ => self.stmt(),
        }
    }

    fn is_declarable(&self, n: usize) -> bool {
        match self.nth(n) {
            FUNCTION_KW | CLASS_KW | INTERFACE_KW | ENUM_KW | LET_KW | CONST_KW | VAR_KW
            | INLINE_KW | NOINLINE_KW => true,
            IDENT => matches!(
                self.nth_text(n),
                "module" | "struct" | "namespace" | "abstract" | "type"
            ),
            _ => false,
        }
    }

    /// `export` / `declare` followed by something that is not a declaration.
    fn bad_modifiers(&mut self, count: usize, decorators: usize) {
        self.start(ERROR);
        if decorators > 0 {
            self.decorators();
        }
        self.error("`export` / `declare` の後には宣言が必要です");
        for _ in 0..count {
            self.bump();
        }
        self.finish();
    }

    fn decorators(&mut self) {
        while self.at(AT) {
            self.start(DECORATOR);
            self.bump();
            if self.nth_is_name(0) {
                self.bump();
            } else {
                self.error("デコレータ名が必要です");
            }
            while self.at(DOT) && self.nth(1) == IDENT {
                self.bump();
                self.bump();
            }
            if self.at(L_PAREN) && !self.newline_before(0) {
                self.arg_list();
            }
            self.finish();
        }
    }

    fn name(&mut self) {
        if self.at(IDENT) {
            self.start(NAME);
            self.bump();
            self.finish();
        } else if self.nth(0).is_keyword() {
            self.error(format!(
                "`{}` は予約語のため名前に使えません",
                self.nth_text(0)
            ));
            self.start(NAME);
            self.bump();
            self.finish();
        } else {
            self.error("識別子が必要です");
        }
    }

    /// Member names may be keywords (`x.default`, COM members …).
    fn member_name(&mut self) {
        if self.nth_is_name(0) {
            self.start(NAME);
            self.bump();
            self.finish();
        } else {
            self.error("メンバー名が必要です");
        }
    }

    fn import_decl(&mut self) {
        self.start(IMPORT_DECL);
        self.bump(); // import
        if self.at(STAR) {
            self.start(IMPORT_NAMESPACE);
            self.bump();
            if self.at_word("as") {
                self.bump();
            } else {
                self.error("`as` が必要です");
            }
            self.name();
            self.finish();
        } else if self.at(L_BRACE) {
            self.bump();
            while !self.at(R_BRACE) && !self.at(EOF) {
                let before = self.spos;
                self.start(IMPORT_SPEC);
                self.name();
                if self.at_word("as") {
                    self.bump();
                    self.name();
                }
                self.finish();
                if !self.eat(COMMA) || self.spos == before {
                    break;
                }
            }
            self.expect(R_BRACE, "`}`");
        } else if self.at(IDENT) && !self.at_word("from") {
            self.error("default import には対応していません。`import { 名前 } from \"...\"` を使ってください");
            self.bump_error();
        }
        if self.at_word("from") {
            self.bump();
        } else if !self.at(STRING_LIT) {
            self.error("`from` が必要です");
        }
        self.expect(STRING_LIT, "モジュール名の文字列");
        self.semi();
        self.finish();
    }

    fn decl_modifiers(&mut self) {
        self.eat(EXPORT_KW);
        if self.at_word("declare") {
            self.bump();
        }
    }

    fn fn_decl(&mut self) {
        self.start(FN_DECL);
        self.decorators();
        self.decl_modifiers();
        while matches!(self.nth(0), INLINE_KW | NOINLINE_KW) {
            self.bump();
        }
        self.expect(FUNCTION_KW, "`function`");
        self.name();
        if self.at(LT) {
            self.type_params();
        }
        self.param_list();
        if self.eat(COLON) {
            self.ty();
        }
        if self.at(L_BRACE) {
            self.block();
        } else {
            self.semi();
        }
        self.finish();
    }

    fn type_params(&mut self) {
        self.start(TYPE_PARAMS);
        self.bump(); // <
        loop {
            let before = self.spos;
            self.start(TYPE_PARAM);
            self.name();
            if self.eat(EXTENDS_KW) {
                self.ty();
            }
            if self.eat(EQ) {
                self.ty();
            }
            self.finish();
            if !self.eat(COMMA) || self.spos == before {
                break;
            }
        }
        self.expect(GT, "`>`");
        self.finish();
    }

    fn type_args(&mut self) {
        self.start(TYPE_ARGS);
        self.bump(); // <
        loop {
            let before = self.spos;
            self.ty();
            if !self.eat(COMMA) || self.spos == before {
                break;
            }
        }
        self.expect(GT, "`>`");
        self.finish();
    }

    fn param_list(&mut self) {
        self.start(PARAM_LIST);
        if self.expect(L_PAREN, "`(`") {
            while !self.at(R_PAREN) && !self.at(EOF) {
                let before = self.spos;
                self.param();
                if self.spos == before {
                    self.error("引数が必要です");
                    if self.at(L_BRACE) {
                        break;
                    }
                    self.bump_error();
                    continue;
                }
                if !self.eat(COMMA) {
                    break;
                }
            }
            self.expect(R_PAREN, "`)`");
        }
        self.finish();
    }

    fn param(&mut self) {
        self.start(PARAM);
        self.decorators();
        if matches!(self.nth(0), BYREF_KW | BYVAL_KW) {
            self.bump();
        }
        self.eat(DOT3);
        if self.at(THIS_KW) {
            self.bump();
        } else if self.at(IDENT) {
            self.name();
        } else {
            self.finish();
            return;
        }
        self.eat(QUESTION);
        if self.eat(COLON) {
            self.ty();
        }
        if self.eat(EQ) {
            self.expr();
        }
        self.finish();
    }

    fn class_decl(&mut self) {
        self.start(CLASS_DECL);
        self.decorators();
        self.decl_modifiers();
        if self.at_word("abstract") {
            self.bump();
        }
        self.expect(CLASS_KW, "`class`");
        self.name();
        if self.at(LT) {
            self.type_params();
        }
        if self.at(EXTENDS_KW) {
            self.start(EXTENDS_CLAUSE);
            self.bump();
            self.ty();
            self.finish();
        }
        if self.at_word("implements") {
            self.start(IMPLEMENTS_CLAUSE);
            self.bump();
            loop {
                let before = self.spos;
                self.ty();
                if !self.eat(COMMA) || self.spos == before {
                    break;
                }
            }
            self.finish();
        }
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                self.class_member();
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    fn class_member(&mut self) {
        if self.eat(SEMI) {
            return;
        }
        let k0 = self.skip_decorators(0);
        let mut k = k0;
        loop {
            let is_mod = match self.nth(k) {
                INLINE_KW | NOINLINE_KW => true,
                IDENT => MEMBER_MODIFIERS.contains(&self.nth_text(k)),
                _ => false,
            };
            // A modifier word directly followed by `(`, `:`, `=`, `;`, `?`, `<` is a member *name*.
            if is_mod
                && !matches!(
                    self.nth(k + 1),
                    L_PAREN | COLON | EQ | SEMI | QUESTION | LT | R_BRACE
                )
            {
                k += 1;
            } else {
                break;
            }
        }
        let kind = if self.nth_word(k, "constructor") && self.nth(k + 1) == L_PAREN {
            CTOR_DECL
        } else if (self.nth_word(k, "get") || self.nth_word(k, "set"))
            && self.nth_is_name(k + 1)
            && self.nth(k + 2) == L_PAREN
        {
            ACCESSOR_DECL
        } else if self.nth_is_name(k) && matches!(self.nth(k + 1), L_PAREN | LT) {
            METHOD_DECL
        } else if self.nth_is_name(k) {
            FIELD_DECL
        } else {
            self.error("クラスメンバーが必要です");
            self.bump_error();
            return;
        };
        self.start(kind);
        self.decorators();
        for _ in k0..k {
            self.bump();
        }
        match kind {
            CTOR_DECL => {
                self.bump(); // constructor
                self.param_list();
                self.block_or_semi();
            }
            ACCESSOR_DECL => {
                self.bump(); // get / set
                self.member_name();
                self.param_list();
                if self.eat(COLON) {
                    self.ty();
                }
                self.block_or_semi();
            }
            METHOD_DECL => {
                self.member_name();
                if self.at(LT) {
                    self.type_params();
                }
                self.param_list();
                if self.eat(COLON) {
                    self.ty();
                }
                self.block_or_semi();
            }
            _ => {
                self.member_name();
                if !self.eat(QUESTION) {
                    self.eat(BANG);
                }
                if self.eat(COLON) {
                    self.ty();
                }
                if self.eat(EQ) {
                    self.expr();
                }
                self.semi();
            }
        }
        self.finish();
    }

    fn block_or_semi(&mut self) {
        if self.at(L_BRACE) {
            self.block();
        } else {
            self.semi();
        }
    }

    fn interface_decl(&mut self) {
        self.start(INTERFACE_DECL);
        self.decorators();
        self.decl_modifiers();
        self.expect(INTERFACE_KW, "`interface`");
        self.name();
        if self.at(LT) {
            self.type_params();
        }
        if self.at(EXTENDS_KW) {
            self.start(EXTENDS_CLAUSE);
            self.bump();
            loop {
                let before = self.spos;
                self.ty();
                if !self.eat(COMMA) || self.spos == before {
                    break;
                }
            }
            self.finish();
        }
        self.sig_body();
        self.finish();
    }

    fn struct_decl(&mut self) {
        self.start(STRUCT_DECL);
        self.decorators();
        self.decl_modifiers();
        self.bump(); // struct
        self.name();
        self.sig_body();
        self.finish();
    }

    fn sig_body(&mut self) {
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                self.interface_member();
            }
            self.expect(R_BRACE, "`}`");
        }
    }

    fn interface_member(&mut self) {
        if self.eat(SEMI) || self.eat(COMMA) {
            return;
        }
        if self.at(L_BRACK) {
            self.start(INDEX_SIG);
            self.bump();
            self.start(PARAM);
            self.name();
            self.eat(QUESTION);
            if self.eat(COLON) {
                self.ty();
            }
            self.finish();
            self.expect(R_BRACK, "`]`");
            if self.eat(COLON) {
                self.ty();
            }
            self.finish();
            return;
        }
        if self.at(L_PAREN) {
            // call signature `(row: any, col?: any): Range` — the default member of a host type
            self.start(METHOD_SIG);
            self.param_list();
            if self.eat(COLON) {
                self.ty();
            }
            self.finish();
            return;
        }
        let readonly = self.at_word("readonly") && self.nth_is_name(1);
        let k = usize::from(readonly);
        if !self.nth_is_name(k) {
            self.error("メンバーの宣言が必要です");
            self.bump_error();
            return;
        }
        let mut j = k + 1;
        if self.nth(j) == QUESTION {
            j += 1;
        }
        if matches!(self.nth(j), L_PAREN | LT) {
            self.start(METHOD_SIG);
            if readonly {
                self.bump();
            }
            self.member_name();
            self.eat(QUESTION);
            if self.at(LT) {
                self.type_params();
            }
            self.param_list();
            if self.eat(COLON) {
                self.ty();
            }
            self.finish();
        } else {
            self.start(PROP_SIG);
            if readonly {
                self.bump();
            }
            self.member_name();
            self.eat(QUESTION);
            if self.eat(COLON) {
                self.ty();
            }
            self.finish();
        }
    }

    fn enum_decl(&mut self) {
        self.start(ENUM_DECL);
        self.decorators();
        self.decl_modifiers();
        self.eat(CONST_KW);
        self.expect(ENUM_KW, "`enum`");
        self.name();
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                let before = self.spos;
                self.start(ENUM_MEMBER);
                self.member_name();
                if self.eat(EQ) {
                    self.expr();
                }
                self.finish();
                if !self.eat(COMMA) || self.spos == before {
                    break;
                }
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    fn type_alias(&mut self) {
        self.start(TYPE_ALIAS);
        self.decl_modifiers();
        self.bump(); // type
        self.name();
        if self.at(LT) {
            self.type_params();
        }
        self.expect(EQ, "`=`");
        self.ty();
        self.semi();
        self.finish();
    }

    fn declare_module(&mut self) {
        self.start(DECLARE_MODULE);
        self.decl_modifiers();
        self.bump(); // module
        self.bump(); // "lib"
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                let k = self.skip_decorators(0);
                let k = if self.nth(k) == EXPORT_KW { k + 1 } else { k };
                if self.nth(k) == FUNCTION_KW {
                    self.fn_decl();
                } else if self.eat(SEMI) {
                } else {
                    self.error("`declare module` の中には `function` 宣言のみ書けます");
                    self.bump_error();
                }
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    fn namespace_decl(&mut self) {
        self.start(NAMESPACE_DECL);
        self.decl_modifiers();
        self.bump(); // namespace
        self.name();
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                self.item();
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    fn var_decl(&mut self, with_semi: bool) {
        self.start(VAR_DECL);
        self.decl_modifiers();
        self.bump(); // let / const / var / using
        loop {
            let before = self.spos;
            self.start(VAR_DECLARATOR);
            if matches!(self.nth(0), L_BRACE | L_BRACK) {
                self.pattern();
            } else {
                self.name();
            }
            self.eat(BANG);
            if self.eat(COLON) {
                self.ty();
            }
            if self.eat(EQ) {
                self.expr();
            }
            self.finish();
            if !self.eat(COMMA) || self.spos == before {
                break;
            }
        }
        if with_semi {
            self.semi();
        }
        self.finish();
    }

    /// Destructuring pattern of a declaration: `{ a, b: alias }` or `[a, , c]`.
    fn pattern(&mut self) {
        let object = self.at(L_BRACE);
        let close = if object { R_BRACE } else { R_BRACK };
        self.start(if object {
            OBJECT_PATTERN
        } else {
            ARRAY_PATTERN
        });
        self.bump();
        while !self.at(close) && !self.at(EOF) {
            let before = self.spos;
            self.start(PATTERN_ELEM);
            if self.at(COMMA) && !object {
                // hole: `[a, , c]`
            } else if self.nth_is_name(0) {
                if object && self.nth(1) == COLON {
                    self.member_name();
                    self.bump();
                }
                if self.at(IDENT) {
                    self.name();
                } else {
                    self.error(
                        "変数名が必要です（分割代入の入れ子・既定値・`...` には対応していません）",
                    );
                }
            } else {
                self.error(
                    "変数名が必要です（分割代入の入れ子・既定値・`...` には対応していません）",
                );
            }
            self.finish();
            if !self.eat(COMMA) {
                break;
            }
            if self.spos == before {
                break;
            }
        }
        self.expect(close, if object { "`}`" } else { "`]`" });
        self.finish();
    }

    // ------------------------------------------------------------------
    // types
    // ------------------------------------------------------------------

    fn ty(&mut self) {
        let cp = self.checkpoint();
        self.ty_postfix();
        if self.at(PIPE) {
            self.start_at(cp, UNION_TYPE);
            while self.eat(PIPE) {
                self.ty_postfix();
            }
            self.finish();
        }
    }

    fn ty_postfix(&mut self) {
        let cp = self.checkpoint();
        self.ty_primary();
        while self.at(L_BRACK) && self.nth(1) == R_BRACK {
            self.start_at(cp, ARRAY_TYPE);
            self.bump();
            self.bump();
            self.finish();
        }
    }

    fn ty_primary(&mut self) {
        match self.nth(0) {
            L_PAREN => {
                let is_fn = matches!(self.skip_balanced(0), Some(n) if self.nth(n) == FAT_ARROW);
                if is_fn {
                    self.start(FN_TYPE);
                    self.param_list();
                    self.expect(FAT_ARROW, "`=>`");
                    self.ty();
                    self.finish();
                } else {
                    self.start(PAREN_TYPE);
                    self.bump();
                    self.ty();
                    self.expect(R_PAREN, "`)`");
                    self.finish();
                }
            }
            IDENT | VOID_KW | NULL_KW | UNDEFINED_KW => {
                self.start(NAMED_TYPE);
                self.bump();
                while self.at(DOT) && self.nth_is_name(1) {
                    self.bump();
                    self.bump();
                }
                if self.at(LT) {
                    self.type_args();
                }
                self.finish();
            }
            L_BRACK => {
                self.start(TUPLE_TYPE);
                self.bump();
                while !self.at(R_BRACK) && !self.at(EOF) {
                    let before = self.spos;
                    self.ty();
                    if !self.eat(COMMA) || self.spos == before {
                        break;
                    }
                }
                self.expect(R_BRACK, "`]`");
                self.finish();
            }
            L_BRACE => {
                self.start(OBJECT_TYPE);
                self.bump();
                while !self.at(R_BRACE) && !self.at(EOF) {
                    if self.eat(SEMI) || self.eat(COMMA) {
                        continue;
                    }
                    if !self.nth_is_name(0) {
                        self.error("フィールド名が必要です");
                        break;
                    }
                    self.start(PROP_SIG);
                    self.member_name();
                    self.eat(QUESTION);
                    if self.eat(COLON) {
                        self.ty();
                    }
                    self.finish();
                }
                self.expect(R_BRACE, "`}`");
                self.finish();
            }
            _ => self.error("型が必要です"),
        }
    }

    // ------------------------------------------------------------------
    // statements
    // ------------------------------------------------------------------

    fn block(&mut self) {
        self.start(BLOCK);
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                self.stmt();
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    fn stmt(&mut self) {
        let k = self.skip_decorators(0);
        match self.nth(k) {
            FOR_KW => return self.for_stmt(k),
            WHILE_KW => {
                self.start(WHILE_STMT);
                self.decorators();
                self.bump();
                self.paren_cond();
                self.stmt();
                self.finish();
                return;
            }
            DO_KW => {
                self.start(DO_WHILE_STMT);
                self.decorators();
                self.bump();
                self.stmt();
                self.expect(WHILE_KW, "`while`");
                self.paren_cond();
                self.eat(SEMI);
                self.finish();
                return;
            }
            FUNCTION_KW | INLINE_KW | NOINLINE_KW => return self.fn_decl(),
            _ if k > 0 => {
                self.error("デコレータは関数・クラス・ループの前にのみ書けます");
                self.start(ERROR);
                self.decorators();
                self.finish();
                return;
            }
            _ => {}
        }
        match self.nth(0) {
            L_BRACE => self.block(),
            LET_KW | CONST_KW | VAR_KW => self.var_decl(true),
            IDENT if self.at_word("using") && self.nth(1) == IDENT && !self.newline_before(1) => {
                self.var_decl(true)
            }
            IDENT if self.nth(1) == COLON => {
                self.start(LABELED_STMT);
                self.bump();
                self.bump();
                self.stmt();
                self.finish();
            }
            IF_KW => {
                self.start(IF_STMT);
                self.bump();
                self.paren_cond();
                self.stmt();
                if self.at(ELSE_KW) {
                    self.start(ELSE_CLAUSE);
                    self.bump();
                    self.stmt();
                    self.finish();
                }
                self.finish();
            }
            SWITCH_KW => self.switch_stmt(),
            TRY_KW => {
                self.start(TRY_STMT);
                self.bump();
                self.block();
                if self.at(CATCH_KW) {
                    self.start(CATCH_CLAUSE);
                    self.bump();
                    if self.eat(L_PAREN) {
                        self.name();
                        if self.eat(COLON) {
                            self.ty();
                        }
                        self.expect(R_PAREN, "`)`");
                    }
                    self.block();
                    self.finish();
                }
                if self.at(FINALLY_KW) {
                    self.start(FINALLY_CLAUSE);
                    self.bump();
                    self.block();
                    self.finish();
                }
                self.finish();
            }
            THROW_KW => {
                self.start(THROW_STMT);
                self.bump();
                self.expr();
                self.semi();
                self.finish();
            }
            RETURN_KW => {
                self.start(RETURN_STMT);
                self.bump();
                if !self.at(SEMI) && !self.at(R_BRACE) && !self.at(EOF) && !self.newline_before(0) {
                    self.expr();
                }
                self.semi();
                self.finish();
            }
            BREAK_KW | CONTINUE_KW => {
                let kind = if self.at(BREAK_KW) {
                    BREAK_STMT
                } else {
                    CONTINUE_STMT
                };
                self.start(kind);
                self.bump();
                if self.at(IDENT) && !self.newline_before(0) {
                    self.bump();
                }
                self.semi();
                self.finish();
            }
            RESUME_KW => {
                self.start(RESUME_STMT);
                self.bump();
                self.semi();
                self.finish();
            }
            SEMI => {
                self.start(EMPTY_STMT);
                self.bump();
                self.finish();
            }
            CLASS_KW | INTERFACE_KW | ENUM_KW | IMPORT_KW | EXPORT_KW => {
                self.error("この宣言はモジュールの最上位にのみ書けます");
                match self.nth(0) {
                    CLASS_KW => self.class_decl(),
                    INTERFACE_KW => self.interface_decl(),
                    ENUM_KW => self.enum_decl(),
                    IMPORT_KW => self.import_decl(),
                    _ => self.bump_error(),
                }
            }
            _ => self.expr_stmt(),
        }
    }

    fn paren_cond(&mut self) {
        self.expect(L_PAREN, "`(`");
        self.expr();
        self.expect(R_PAREN, "`)`");
    }

    fn expr_stmt(&mut self) {
        let before = self.spos;
        self.start(EXPR_STMT);
        self.expr();
        if self.spos == before {
            self.bump_error();
        } else {
            self.semi();
        }
        self.finish();
    }

    fn for_stmt(&mut self, k: usize) {
        // `for (let x of xs)` / `for (const x: T of xs)` / `for (let k in obj)`
        let mut is_of = false;
        let binding_end = match self.nth(k + 3) {
            IDENT => Some(k + 4),
            L_BRACE | L_BRACK => self.skip_balanced(k + 3),
            _ => None,
        };
        if let (true, Some(after_name)) = (
            self.nth(k + 1) == L_PAREN && matches!(self.nth(k + 2), LET_KW | CONST_KW | VAR_KW),
            binding_end,
        ) {
            let is_iter_kw = |p: &Self, n: usize| p.nth_word(n, "of") || p.nth(n) == IN_KW;
            if is_iter_kw(self, after_name) {
                is_of = true;
            } else if self.nth(after_name) == COLON {
                if let Some(m) = self.skip_type(after_name + 1) {
                    is_of = is_iter_kw(self, m);
                }
            }
        }
        if is_of {
            self.start(FOR_OF_STMT);
            self.decorators();
            self.bump(); // for
            self.bump(); // (
            self.bump(); // let / const
            if matches!(self.nth(0), L_BRACE | L_BRACK) {
                self.pattern();
            } else {
                self.name();
            }
            if self.eat(COLON) {
                self.ty();
            }
            if self.at(IN_KW) {
                self.error("`for...in` には対応していません。`for...of` を使ってください");
            }
            self.bump(); // of / in
            self.expr();
            self.expect(R_PAREN, "`)`");
            self.stmt();
            self.finish();
            return;
        }
        self.start(FOR_STMT);
        self.decorators();
        self.bump(); // for
        self.expect(L_PAREN, "`(`");
        self.start(FOR_INIT);
        if matches!(self.nth(0), LET_KW | CONST_KW | VAR_KW) {
            self.var_decl(false);
        } else if !self.at(SEMI) {
            self.expr_list(SEMI);
        }
        self.finish();
        self.expect(SEMI, "`;`");
        self.start(FOR_COND);
        if !self.at(SEMI) {
            self.expr();
        }
        self.finish();
        self.expect(SEMI, "`;`");
        self.start(FOR_UPDATE);
        if !self.at(R_PAREN) {
            self.expr_list(R_PAREN);
        }
        self.finish();
        self.expect(R_PAREN, "`)`");
        self.stmt();
        self.finish();
    }

    fn expr_list(&mut self, end: SyntaxKind) {
        while !self.at(end) && !self.at(EOF) {
            let before = self.spos;
            self.expr();
            if !self.eat(COMMA) || self.spos == before {
                break;
            }
        }
    }

    fn switch_stmt(&mut self) {
        self.start(SWITCH_STMT);
        self.bump();
        self.paren_cond();
        if self.expect(L_BRACE, "`{`") {
            while !self.at(R_BRACE) && !self.at(EOF) {
                if !matches!(self.nth(0), CASE_KW | DEFAULT_KW) {
                    self.error("`case` または `default` が必要です");
                    self.bump_error();
                    continue;
                }
                self.start(CASE_CLAUSE);
                if self.eat(CASE_KW) {
                    self.expr();
                } else {
                    self.bump();
                }
                self.expect(COLON, "`:`");
                while !matches!(self.nth(0), CASE_KW | DEFAULT_KW | R_BRACE | EOF) {
                    self.stmt();
                }
                self.finish();
            }
            self.expect(R_BRACE, "`}`");
        }
        self.finish();
    }

    // ------------------------------------------------------------------
    // expressions
    // ------------------------------------------------------------------

    fn expr(&mut self) {
        if self.at_arrow_start() {
            return self.arrow();
        }
        let cp = self.checkpoint();
        self.expr_ternary();
        let n = self.assign_op_len();
        if n > 0 {
            self.start_at(cp, ASSIGN_EXPR);
            for _ in 0..n {
                self.bump();
            }
            self.expr();
            self.finish();
        }
    }

    fn assign_op_len(&self) -> usize {
        match self.nth(0) {
            EQ | PLUSEQ | MINUSEQ | STAREQ | SLASHEQ | PERCENTEQ | AMPEQ | PIPEEQ | CARETEQ
            | SHLEQ => 1,
            // `>>=`
            GT if self.nth(1) == GT
                && self.nth(2) == EQ
                && self.adjacent(0)
                && self.adjacent(1) =>
            {
                3
            }
            _ => 0,
        }
    }

    fn at_arrow_start(&self) -> bool {
        match self.nth(0) {
            IDENT => self.nth(1) == FAT_ARROW,
            L_PAREN => match self.skip_balanced(0) {
                Some(n) => {
                    self.nth(n) == FAT_ARROW
                        || (self.nth(n) == COLON
                            && matches!(self.skip_type(n + 1), Some(m) if self.nth(m) == FAT_ARROW))
                }
                None => false,
            },
            _ => false,
        }
    }

    fn arrow(&mut self) {
        self.start(ARROW_EXPR);
        if self.at(IDENT) {
            self.start(PARAM_LIST);
            self.start(PARAM);
            self.name();
            self.finish();
            self.finish();
        } else {
            self.param_list();
            if self.eat(COLON) {
                self.ty();
            }
        }
        self.expect(FAT_ARROW, "`=>`");
        if self.at(L_BRACE) {
            self.block();
        } else {
            self.expr();
        }
        self.finish();
    }

    fn expr_ternary(&mut self) {
        let cp = self.checkpoint();
        self.expr_bp(1);
        if self.at(QUESTION) {
            self.start_at(cp, TERNARY_EXPR);
            self.bump();
            self.expr();
            self.expect(COLON, "`:`");
            self.expr();
            self.finish();
        }
    }

    /// Returns (binding power, token count, right-assoc).
    fn peek_binop(&self) -> Option<(u8, usize, bool)> {
        Some(match self.nth(0) {
            QUESTION2 => (1, 1, false),
            PIPE2 => (2, 1, false),
            AMP2 => (3, 1, false),
            PIPE => (4, 1, false),
            CARET => (5, 1, false),
            AMP => (6, 1, false),
            EQ2 | EQ3 | NEQ | NEQ2 => (7, 1, false),
            LT | LTEQ | INSTANCEOF_KW | IN_KW => (8, 1, false),
            GT => {
                if self.nth(1) == GT && self.adjacent(0) {
                    if self.nth(2) == EQ && self.adjacent(1) {
                        return None; // `>>=`
                    }
                    (9, 2, false)
                } else if self.nth(1) == EQ && self.adjacent(0) {
                    (8, 2, false)
                } else {
                    (8, 1, false)
                }
            }
            SHL => (9, 1, false),
            PLUS | MINUS => (10, 1, false),
            STAR | SLASH | PERCENT => (11, 1, false),
            STAR2 => (12, 1, true),
            _ => return None,
        })
    }

    fn expr_bp(&mut self, min_bp: u8) {
        let cp = self.checkpoint();
        self.expr_unary();
        loop {
            if self.at_word("as") && min_bp <= 8 && !self.newline_before(0) {
                self.start_at(cp, AS_EXPR);
                self.bump();
                self.ty();
                self.finish();
                continue;
            }
            let Some((bp, n, right)) = self.peek_binop() else {
                break;
            };
            if bp < min_bp {
                break;
            }
            let is_instanceof = self.at(INSTANCEOF_KW);
            self.start_at(cp, BINARY_EXPR);
            for _ in 0..n {
                self.bump();
            }
            if is_instanceof {
                self.ty();
            } else {
                self.expr_bp(if right { bp } else { bp + 1 });
            }
            self.finish();
        }
    }

    fn expr_unary(&mut self) {
        match self.nth(0) {
            BANG | MINUS | PLUS | TILDE | TYPEOF_KW | PLUS2 | MINUS2 => {
                self.start(UNARY_EXPR);
                self.bump();
                self.expr_unary();
                self.finish();
            }
            _ => self.expr_postfix(),
        }
    }

    fn expr_postfix(&mut self) {
        let cp = self.checkpoint();
        self.expr_primary();
        loop {
            match self.nth(0) {
                DOT | QUESTION_DOT => {
                    self.start_at(cp, MEMBER_EXPR);
                    self.bump();
                    if self.nth_is_name(0) {
                        self.bump();
                    } else {
                        self.error("メンバー名が必要です");
                    }
                    self.finish();
                }
                L_BRACK => {
                    self.start_at(cp, INDEX_EXPR);
                    self.bump();
                    self.expr();
                    self.expect(R_BRACK, "`]`");
                    self.finish();
                }
                L_PAREN => {
                    self.start_at(cp, CALL_EXPR);
                    self.arg_list();
                    self.finish();
                }
                LT => match self.skip_type_args(0) {
                    Some(n) if self.nth(n) == L_PAREN => {
                        self.start_at(cp, CALL_EXPR);
                        self.type_args();
                        self.arg_list();
                        self.finish();
                    }
                    _ => break,
                },
                BANG if !self.newline_before(0)
                    && matches!(
                        self.nth(1),
                        DOT | R_PAREN | R_BRACK | SEMI | COMMA | L_BRACK | L_PAREN | EOF | R_BRACE
                    ) =>
                {
                    self.start_at(cp, NON_NULL_EXPR);
                    self.bump();
                    self.finish();
                }
                // `` fx`…` `` (spec §2.7): the tag and the template on the same line.
                TEMPLATE_LIT if !self.newline_before(0) => {
                    self.start_at(cp, TAGGED_TEMPLATE);
                    self.start(LITERAL);
                    self.bump();
                    self.finish();
                    self.finish();
                }
                PLUS2 | MINUS2 if !self.newline_before(0) => {
                    self.start_at(cp, POSTFIX_EXPR);
                    self.bump();
                    self.finish();
                }
                _ => break,
            }
        }
    }

    fn arg_list(&mut self) {
        self.start(ARG_LIST);
        self.bump(); // (
        while !self.at(R_PAREN) && !self.at(EOF) {
            let before = self.spos;
            if self.at(DOT3) {
                self.error("スプレッド引数には対応していません");
                self.bump();
            }
            if self.nth_is_name(0) && self.nth(1) == COLON {
                // Named argument `name: value` (spec §3.3).
                self.start(NAMED_ARG);
                self.start(NAME);
                self.bump();
                self.finish();
                self.bump(); // :
                self.expr();
                self.finish();
            } else {
                self.expr();
            }
            if !self.eat(COMMA) || self.spos == before {
                break;
            }
        }
        self.expect(R_PAREN, "`)`");
        self.finish();
    }

    fn expr_primary(&mut self) {
        match self.nth(0) {
            INT_LIT | FLOAT_LIT | STRING_LIT | TEMPLATE_LIT | TRUE_KW | FALSE_KW | NULL_KW
            | UNDEFINED_KW => {
                self.start(LITERAL);
                self.bump();
                self.finish();
            }
            IDENT => {
                self.start(NAME_REF);
                self.bump();
                self.finish();
            }
            THIS_KW => {
                self.start(THIS_EXPR);
                self.bump();
                self.finish();
            }
            SUPER_KW => {
                self.start(SUPER_EXPR);
                self.bump();
                self.finish();
            }
            L_PAREN => {
                self.start(PAREN_EXPR);
                self.bump();
                self.expr();
                self.expect(R_PAREN, "`)`");
                self.finish();
            }
            L_BRACK => {
                self.start(ARRAY_EXPR);
                self.bump();
                self.expr_list(R_BRACK);
                self.expect(R_BRACK, "`]`");
                self.finish();
            }
            L_BRACE => {
                self.start(OBJECT_EXPR);
                self.bump();
                while !self.at(R_BRACE) && !self.at(EOF) {
                    if !(self.nth_is_name(0) || self.at(STRING_LIT)) {
                        self.error("プロパティ名が必要です");
                        break;
                    }
                    self.start(OBJECT_PROP);
                    self.bump();
                    if self.eat(COLON) {
                        self.expr();
                    }
                    self.finish();
                    if !self.eat(COMMA) {
                        break;
                    }
                }
                self.expect(R_BRACE, "`}`");
                self.finish();
            }
            NEW_KW => {
                self.start(NEW_EXPR);
                self.bump();
                self.start(NAMED_TYPE);
                if self.at(IDENT) {
                    self.bump();
                } else {
                    self.error("クラス名が必要です");
                }
                while self.at(DOT) && self.nth(1) == IDENT {
                    self.bump();
                    self.bump();
                }
                if self.at(LT) && self.skip_type_args(0).is_some() {
                    self.type_args();
                }
                self.finish();
                if self.at(L_PAREN) {
                    self.arg_list();
                }
                self.finish();
            }
            FUNCTION_KW => {
                self.error(
                    "`function` 式には対応していません。アロー関数 `(x) => ...` を使ってください",
                );
                self.start(ERROR);
                self.finish();
            }
            ERROR_TOKEN => {
                self.error("解釈できない文字です");
                self.bump_error();
            }
            _ => {
                self.error("式が必要です");
                self.start(ERROR);
                self.finish();
            }
        }
    }
}
