//! Syntax kinds (tokens and tree nodes) and the `logos` lexer.

use logos::Logos;

/// Every token kind and node kind of the RAS concrete syntax tree.
#[allow(non_camel_case_types, clippy::upper_case_acronyms)]
#[derive(Logos, Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(u16)]
pub enum SyntaxKind {
    // ---- trivia ----
    #[regex(r"[ \t\r\n\u{00A0}\u{FEFF}\u{3000}]+")]
    WHITESPACE,
    #[regex(r"//[^\n]*", allow_greedy = true)]
    LINE_COMMENT,
    #[regex(r"/\*([^*]|\*+[^*/])*\*+/")]
    BLOCK_COMMENT,

    // ---- literals ----
    #[regex(r"[0-9][0-9_]*[lL]?")]
    #[regex(r"0[xX][0-9a-fA-F_]+[lL]?")]
    #[regex(r"0[bB][01_]+[lL]?")]
    INT_LIT,
    #[regex(r"[0-9][0-9_]*\.[0-9][0-9_]*([eE][+-]?[0-9]+)?[fFdDmM]?")]
    #[regex(r"[0-9][0-9_]*[eE][+-]?[0-9]+[fFdDmM]?")]
    #[regex(r"[0-9][0-9_]*[fFdDmM]")]
    FLOAT_LIT,
    #[regex(r#""([^"\\\n]|\\[^\n])*""#)]
    #[regex(r#"'([^'\\\n]|\\[^\n])*'"#)]
    STRING_LIT,
    #[regex(r"`([^`\\]|\\[\s\S])*`")]
    TEMPLATE_LIT,

    #[regex(r"[\p{XID_Start}_$][\p{XID_Continue}$]*")]
    IDENT,

    // ---- keywords (reserved) ----
    #[token("let")]
    LET_KW,
    #[token("const")]
    CONST_KW,
    #[token("function")]
    FUNCTION_KW,
    #[token("return")]
    RETURN_KW,
    #[token("if")]
    IF_KW,
    #[token("else")]
    ELSE_KW,
    #[token("for")]
    FOR_KW,
    #[token("while")]
    WHILE_KW,
    #[token("do")]
    DO_KW,
    #[token("break")]
    BREAK_KW,
    #[token("continue")]
    CONTINUE_KW,
    #[token("switch")]
    SWITCH_KW,
    #[token("case")]
    CASE_KW,
    #[token("default")]
    DEFAULT_KW,
    #[token("try")]
    TRY_KW,
    #[token("catch")]
    CATCH_KW,
    #[token("finally")]
    FINALLY_KW,
    #[token("throw")]
    THROW_KW,
    #[token("new")]
    NEW_KW,
    #[token("class")]
    CLASS_KW,
    #[token("extends")]
    EXTENDS_KW,
    #[token("interface")]
    INTERFACE_KW,
    #[token("enum")]
    ENUM_KW,
    #[token("import")]
    IMPORT_KW,
    #[token("export")]
    EXPORT_KW,
    #[token("this")]
    THIS_KW,
    #[token("super")]
    SUPER_KW,
    #[token("null")]
    NULL_KW,
    #[token("undefined")]
    UNDEFINED_KW,
    #[token("true")]
    TRUE_KW,
    #[token("false")]
    FALSE_KW,
    #[token("void")]
    VOID_KW,
    #[token("typeof")]
    TYPEOF_KW,
    #[token("instanceof")]
    INSTANCEOF_KW,
    #[token("in")]
    IN_KW,
    #[token("var")]
    VAR_KW,
    // RAS additions (spec §1)
    #[token("byref")]
    BYREF_KW,
    #[token("byval")]
    BYVAL_KW,
    #[token("inline")]
    INLINE_KW,
    #[token("noinline")]
    NOINLINE_KW,
    #[token("resume")]
    RESUME_KW,

    // ---- punctuation ----
    #[token("(")]
    L_PAREN,
    #[token(")")]
    R_PAREN,
    #[token("{")]
    L_BRACE,
    #[token("}")]
    R_BRACE,
    #[token("[")]
    L_BRACK,
    #[token("]")]
    R_BRACK,
    #[token(";")]
    SEMI,
    #[token(",")]
    COMMA,
    #[token(".")]
    DOT,
    #[token("...")]
    DOT3,
    #[token("?.")]
    QUESTION_DOT,
    #[token(":")]
    COLON,
    #[token("?")]
    QUESTION,
    #[token("=>")]
    FAT_ARROW,
    #[token("=")]
    EQ,
    #[token("==")]
    EQ2,
    #[token("===")]
    EQ3,
    #[token("!=")]
    NEQ,
    #[token("!==")]
    NEQ2,
    #[token("<")]
    LT,
    /// `>` is always lexed alone; `>=`, `>>`, `>>=` are recombined by the parser
    /// so that `Map<K, V[]>` and `Stack<Map<K, V>>` parse without re-lexing.
    #[token(">")]
    GT,
    #[token("<=")]
    LTEQ,
    #[token("+")]
    PLUS,
    #[token("-")]
    MINUS,
    #[token("*")]
    STAR,
    #[token("/")]
    SLASH,
    #[token("%")]
    PERCENT,
    #[token("**")]
    STAR2,
    #[token("++")]
    PLUS2,
    #[token("--")]
    MINUS2,
    #[token("+=")]
    PLUSEQ,
    #[token("-=")]
    MINUSEQ,
    #[token("*=")]
    STAREQ,
    #[token("/=")]
    SLASHEQ,
    #[token("%=")]
    PERCENTEQ,
    #[token("&=")]
    AMPEQ,
    #[token("|=")]
    PIPEEQ,
    #[token("^=")]
    CARETEQ,
    #[token("<<=")]
    SHLEQ,
    #[token("&&")]
    AMP2,
    #[token("||")]
    PIPE2,
    #[token("??")]
    QUESTION2,
    #[token("!")]
    BANG,
    #[token("&")]
    AMP,
    #[token("|")]
    PIPE,
    #[token("^")]
    CARET,
    #[token("~")]
    TILDE,
    #[token("<<")]
    SHL,
    #[token("@")]
    AT,

    ERROR_TOKEN,
    EOF,

    // ---- nodes ----
    SOURCE_FILE,
    ERROR,
    NAME,
    DECORATOR,
    // items
    IMPORT_DECL,
    IMPORT_SPEC,
    IMPORT_NAMESPACE,
    FN_DECL,
    TYPE_PARAMS,
    TYPE_PARAM,
    TYPE_ARGS,
    PARAM_LIST,
    PARAM,
    CLASS_DECL,
    EXTENDS_CLAUSE,
    IMPLEMENTS_CLAUSE,
    FIELD_DECL,
    METHOD_DECL,
    CTOR_DECL,
    ACCESSOR_DECL,
    INTERFACE_DECL,
    PROP_SIG,
    METHOD_SIG,
    INDEX_SIG,
    ENUM_DECL,
    ENUM_MEMBER,
    TYPE_ALIAS,
    DECLARE_MODULE,
    STRUCT_DECL,
    NAMESPACE_DECL,
    VAR_DECL,
    VAR_DECLARATOR,
    OBJECT_PATTERN,
    ARRAY_PATTERN,
    PATTERN_ELEM,
    // statements
    BLOCK,
    EXPR_STMT,
    IF_STMT,
    ELSE_CLAUSE,
    FOR_STMT,
    FOR_INIT,
    FOR_COND,
    FOR_UPDATE,
    FOR_OF_STMT,
    WHILE_STMT,
    DO_WHILE_STMT,
    SWITCH_STMT,
    CASE_CLAUSE,
    TRY_STMT,
    CATCH_CLAUSE,
    FINALLY_CLAUSE,
    THROW_STMT,
    RETURN_STMT,
    BREAK_STMT,
    CONTINUE_STMT,
    LABELED_STMT,
    RESUME_STMT,
    EMPTY_STMT,
    // expressions
    LITERAL,
    NAME_REF,
    THIS_EXPR,
    SUPER_EXPR,
    PAREN_EXPR,
    ARRAY_EXPR,
    OBJECT_EXPR,
    OBJECT_PROP,
    ARROW_EXPR,
    UNARY_EXPR,
    POSTFIX_EXPR,
    NON_NULL_EXPR,
    BINARY_EXPR,
    ASSIGN_EXPR,
    TERNARY_EXPR,
    CALL_EXPR,
    NEW_EXPR,
    MEMBER_EXPR,
    INDEX_EXPR,
    AS_EXPR,
    ARG_LIST,
    // types
    NAMED_TYPE,
    ARRAY_TYPE,
    FN_TYPE,
    PAREN_TYPE,
    UNION_TYPE,
    TUPLE_TYPE,
    OBJECT_TYPE,

    #[doc(hidden)]
    __LAST,
}

use SyntaxKind::*;

impl SyntaxKind {
    pub fn is_trivia(self) -> bool {
        matches!(self, WHITESPACE | LINE_COMMENT | BLOCK_COMMENT)
    }

    pub fn is_keyword(self) -> bool {
        (self as u16) >= (LET_KW as u16) && (self as u16) <= (RESUME_KW as u16)
    }

    pub fn is_expr(self) -> bool {
        matches!(
            self,
            LITERAL
                | NAME_REF
                | THIS_EXPR
                | SUPER_EXPR
                | PAREN_EXPR
                | ARRAY_EXPR
                | OBJECT_EXPR
                | ARROW_EXPR
                | UNARY_EXPR
                | POSTFIX_EXPR
                | NON_NULL_EXPR
                | BINARY_EXPR
                | ASSIGN_EXPR
                | TERNARY_EXPR
                | CALL_EXPR
                | NEW_EXPR
                | MEMBER_EXPR
                | INDEX_EXPR
                | AS_EXPR
                | ERROR
        )
    }

    pub fn is_type(self) -> bool {
        matches!(self, NAMED_TYPE | ARRAY_TYPE | FN_TYPE | PAREN_TYPE | UNION_TYPE | TUPLE_TYPE | OBJECT_TYPE)
    }

    pub fn is_stmt(self) -> bool {
        matches!(
            self,
            BLOCK
                | EXPR_STMT
                | IF_STMT
                | FOR_STMT
                | FOR_OF_STMT
                | WHILE_STMT
                | DO_WHILE_STMT
                | SWITCH_STMT
                | TRY_STMT
                | THROW_STMT
                | RETURN_STMT
                | BREAK_STMT
                | CONTINUE_STMT
                | LABELED_STMT
                | RESUME_STMT
                | EMPTY_STMT
                | VAR_DECL
                | FN_DECL
        )
    }

    pub fn is_item(self) -> bool {
        matches!(
            self,
            IMPORT_DECL
                | FN_DECL
                | CLASS_DECL
                | INTERFACE_DECL
                | ENUM_DECL
                | TYPE_ALIAS
                | DECLARE_MODULE
                | STRUCT_DECL
                | NAMESPACE_DECL
                | VAR_DECL
        )
    }
}

impl From<SyntaxKind> for rowan::SyntaxKind {
    fn from(kind: SyntaxKind) -> Self {
        Self(kind as u16)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum RasLanguage {}

impl rowan::Language for RasLanguage {
    type Kind = SyntaxKind;

    fn kind_from_raw(raw: rowan::SyntaxKind) -> SyntaxKind {
        assert!(raw.0 < __LAST as u16);
        // SAFETY: `SyntaxKind` is `repr(u16)` with contiguous discriminants and the
        // value was bounds-checked against `__LAST` above.
        unsafe { std::mem::transmute::<u16, SyntaxKind>(raw.0) }
    }

    fn kind_to_raw(kind: SyntaxKind) -> rowan::SyntaxKind {
        kind.into()
    }
}

pub type SyntaxNode = rowan::SyntaxNode<RasLanguage>;
pub type SyntaxToken = rowan::SyntaxToken<RasLanguage>;
pub type SyntaxElement = rowan::SyntaxElement<RasLanguage>;

/// One lexed token: its kind and the byte range in the source.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Token {
    pub kind: SyntaxKind,
    pub start: u32,
    pub end: u32,
}

/// Lexes the whole input. Never fails: unknown input becomes `ERROR_TOKEN`s.
/// The concatenation of all token texts is exactly `text` (lossless).
pub fn lex(text: &str) -> Vec<Token> {
    let mut lexer = SyntaxKind::lexer(text);
    let mut out: Vec<Token> = Vec::new();
    while let Some(res) = lexer.next() {
        let span = lexer.span();
        let kind = res.unwrap_or(ERROR_TOKEN);
        // Merge adjacent error bytes into one token.
        if kind == ERROR_TOKEN {
            if let Some(last) = out.last_mut() {
                if last.kind == ERROR_TOKEN && last.end as usize == span.start {
                    last.end = span.end as u32;
                    continue;
                }
            }
        }
        out.push(Token { kind, start: span.start as u32, end: span.end as u32 });
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn kinds(src: &str) -> Vec<SyntaxKind> {
        lex(src).into_iter().map(|t| t.kind).filter(|k| !k.is_trivia()).collect()
    }

    #[test]
    fn numbers() {
        assert_eq!(kinds("123 123L 0xFF 1.5 1e3 2d 1.5f 1.5m"), vec![
            INT_LIT, INT_LIT, INT_LIT, FLOAT_LIT, FLOAT_LIT, FLOAT_LIT, FLOAT_LIT, FLOAT_LIT
        ]);
    }

    #[test]
    fn keywords_and_idents() {
        assert_eq!(kinds("byref x inlined 変数 resume"), vec![BYREF_KW, IDENT, IDENT, IDENT, RESUME_KW]);
    }

    #[test]
    fn lossless() {
        let src = "let s = \"a\\\"b\"; // c\n/* d */ `t${x}` § @x";
        let toks = lex(src);
        let joined: String = toks.iter().map(|t| &src[t.start as usize..t.end as usize]).collect();
        assert_eq!(joined, src);
    }

    #[test]
    fn member_after_index() {
        assert_eq!(kinds("a[0].length"), vec![IDENT, L_BRACK, INT_LIT, R_BRACK, DOT, IDENT]);
    }
}
