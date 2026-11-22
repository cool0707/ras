//! RAS syntax: `logos` lexer, hand-written recursive-descent parser building a
//! lossless `rowan` tree (CST), and an owned AST lowered from it.

pub mod ast;
mod format;
mod kind;
mod lower;
mod parser;

pub use format::format;
pub use kind::{lex, RasLanguage, SyntaxElement, SyntaxKind, SyntaxNode, SyntaxToken, Token};
pub use lower::{doc_comment, string_value, unescape};
pub use parser::SyntaxError;
pub use rowan::{GreenNode, TextRange, TextSize};

/// Result of parsing one source file.
#[derive(Debug, Clone)]
pub struct Parse {
    pub green: GreenNode,
    pub errors: Vec<SyntaxError>,
    pub ast: ast::Module,
}

impl Parse {
    /// Root of the lossless concrete syntax tree.
    pub fn syntax(&self) -> SyntaxNode {
        SyntaxNode::new_root(self.green.clone())
    }
}

/// Parses a `.ras` / `.d.ras` source file. Never fails; syntax problems are
/// reported in `Parse::errors` and the tree always covers the whole input.
pub fn parse(text: &str) -> Parse {
    let mut p = parser::Parser::new(text);
    p.source_file();
    let (green, mut errors) = p.finish_tree();
    let root = SyntaxNode::new_root(green.clone());
    let mut lower = lower::Lower::new();
    let ast = lower.module(&root);
    errors.extend(lower.errors);
    errors.sort_by_key(|e| e.range.start());
    Parse { green, errors, ast }
}

#[cfg(test)]
mod tests {
    use super::ast::*;
    use super::*;

    fn ok(src: &str) -> Module {
        let p = parse(src);
        assert!(p.errors.is_empty(), "unexpected errors: {:?}", p.errors);
        assert_eq!(p.syntax().text().to_string(), src, "tree must be lossless");
        p.ast
    }

    #[test]
    fn function_with_byref() {
        let m = ok("function calculate(x: double, byref outResult: double): boolean {\n    outResult = x * 2;\n    return true;\n}\n");
        let Item::Fn(f) = &m.items[0] else { panic!() };
        assert_eq!(f.name.text, "calculate");
        assert_eq!(f.params.len(), 2);
        assert_eq!(f.params[1].mode, PassMode::ByRef);
        assert_eq!(f.body.as_ref().unwrap().stmts.len(), 2);
    }

    #[test]
    fn extension_method_and_generics() {
        let m = ok("export function maxOf<T extends Comparable>(this: T[]): T {\n let m = this[0];\n for (const v of this) if (v > m) m = v;\n return m;\n}");
        let Item::Fn(f) = &m.items[0] else { panic!() };
        assert!(f.exported);
        assert!(f.params[0].is_this);
        assert_eq!(f.type_params.len(), 1);
        assert!(f.type_params[0].constraint.is_some());
    }

    #[test]
    fn class_members() {
        let m = ok("class Stack<T> {\n private items: T[] = [];\n push(x: T): void { this.items.push(x); }\n pop(): T { return this.items.pop(); }\n get size(): int { return this.items.length; }\n constructor(n: int) {}\n}");
        let Item::Class(c) = &m.items[0] else {
            panic!()
        };
        assert_eq!(c.members.len(), 5);
        assert!(matches!(&c.members[0], ClassMember::Field(f) if f.vis == Visibility::Private));
        assert!(matches!(&c.members[3], ClassMember::Accessor(a) if a.kind == AccessorKind::Get));
        assert!(matches!(&c.members[4], ClassMember::Ctor(_)));
    }

    #[test]
    fn generic_call_vs_comparison() {
        let m = ok("const a = f<int>(1);\nconst b = x < y;\nconst s = new Stack<Excel.Range>();\nconst c = a < b && c > (d);");
        let Item::Var(v) = &m.items[0] else { panic!() };
        assert!(
            matches!(&v.init.as_ref().unwrap().kind, ExprKind::Call { type_args, .. } if type_args.len() == 1)
        );
        let Item::Var(v) = &m.items[1] else { panic!() };
        assert!(matches!(
            &v.init.as_ref().unwrap().kind,
            ExprKind::Binary {
                op: BinaryOp::Lt,
                ..
            }
        ));
        let Item::Var(v) = &m.items[3] else { panic!() };
        assert!(matches!(
            &v.init.as_ref().unwrap().kind,
            ExprKind::Binary {
                op: BinaryOp::And,
                ..
            }
        ));
    }

    #[test]
    fn shift_and_ge_recombine() {
        let m = ok("x = a >= b; y = a >> 2; z >>= 1; let q: Map<string, int[]> = w;");
        let Item::Stmt(s) = &m.items[0] else { panic!() };
        let StmtKind::Expr(e) = &s.kind else { panic!() };
        let ExprKind::Assign { value, .. } = &e.kind else {
            panic!()
        };
        assert!(matches!(
            value.kind,
            ExprKind::Binary {
                op: BinaryOp::Ge,
                ..
            }
        ));
        let Item::Stmt(s) = &m.items[2] else { panic!() };
        let StmtKind::Expr(e) = &s.kind else { panic!() };
        assert!(matches!(
            e.kind,
            ExprKind::Assign {
                op: Some(BinaryOp::Shr),
                ..
            }
        ));
    }

    #[test]
    fn arrows_and_ternary() {
        let m = ok("let r = arr.filter(x => x > 0).map((x: double): double => { return x * 2; });\nlet t = c ? (a) : b;");
        let Item::Var(v) = &m.items[1] else { panic!() };
        assert!(matches!(
            v.init.as_ref().unwrap().kind,
            ExprKind::Ternary { .. }
        ));
    }

    #[test]
    fn try_resume_and_using() {
        let m = ok("function f() {\n try { ws.Activate(); } catch { resume; }\n using _ = fastMode({ calc: false });\n try { a(); } catch (e) { throw e; } finally { b(); }\n}");
        let Item::Fn(f) = &m.items[0] else { panic!() };
        let body = &f.body.as_ref().unwrap().stmts;
        let StmtKind::Try(t) = &body[0].kind else {
            panic!()
        };
        assert!(t.catch.as_ref().unwrap().is_resume_block());
        assert!(matches!(&body[1].kind, StmtKind::Var(v) if v.kind == VarKind::Using));
        let StmtKind::Try(t) = &body[2].kind else {
            panic!()
        };
        assert!(t.finally.is_some());
    }

    #[test]
    fn decorators_and_declare() {
        let m = ok("@on(\"Sheet1.Change\")\nfunction h(target: Excel.Range) {}\n@fast\nexport function r() { @bulk for (let i = 0; i < 3; i++) {} }\ndeclare module \"user32\" {\n function GetForegroundWindow(): ptr;\n @alias(\"GetWindowTextW\")\n function GetWindowText(hWnd: ptr, lpString: wstr, nMaxCount: int): int;\n}");
        let Item::Fn(f) = &m.items[0] else { panic!() };
        assert_eq!(f.decorators[0].name.text, "on");
        assert_eq!(f.decorators[0].args.len(), 1);
        let Item::DeclareModule(d) = &m.items[2] else {
            panic!()
        };
        assert_eq!(d.lib, "user32");
        assert_eq!(d.functions.len(), 2);
        assert_eq!(d.functions[1].decorators[0].name.text, "alias");
    }

    #[test]
    fn literals() {
        let m = ok("const a = 123L; const b = 0xFF; const c = 1.5f; const d = 1.5m; const e = `x=${x + 1}!`; const s = 'a\\n\\u{3042}';");
        let get = |i: usize| match &m.items[i] {
            Item::Var(v) => v.init.clone().unwrap().kind,
            _ => panic!(),
        };
        assert!(matches!(
            get(0),
            ExprKind::Int {
                value: 123,
                suffix: NumSuffix::L,
                ..
            }
        ));
        assert!(matches!(
            get(1),
            ExprKind::Int {
                value: 255,
                hex: true,
                ..
            }
        ));
        assert!(matches!(
            get(2),
            ExprKind::Float {
                suffix: NumSuffix::F,
                ..
            }
        ));
        assert!(matches!(
            get(3),
            ExprKind::Float {
                suffix: NumSuffix::M,
                ..
            }
        ));
        let ExprKind::Template(parts) = get(4) else {
            panic!()
        };
        assert_eq!(parts.len(), 3);
        assert!(matches!(get(5), ExprKind::Str(s) if s == "a\nあ"));
    }

    #[test]
    fn jsdoc_is_attached() {
        let m = ok("/**\n * 指定列の最終行を返す。\n * @param col 列番号\n */\nexport function lastRow(this: Excel.Worksheet, col: int = 1): int { return 1; }");
        let Item::Fn(f) = &m.items[0] else { panic!() };
        assert_eq!(
            f.doc.as_deref(),
            Some("指定列の最終行を返す。\n@param col 列番号")
        );
    }

    #[test]
    fn host_typedefs() {
        let m = ok("declare namespace Excel {\n interface Range {\n  Value: any;\n  readonly Row: int;\n  Cells(row?: any, col?: any): Range;\n  [index: any]: Range;\n }\n enum XlDirection { xlDown = -4121, xlUp = -4162 }\n}\ndeclare const Application: Excel.Application;");
        let Item::Namespace(ns) = &m.items[0] else {
            panic!()
        };
        assert_eq!(ns.items.len(), 2);
        let Item::Interface(i) = &ns.items[0] else {
            panic!()
        };
        assert_eq!(i.members.len(), 4);
        let Item::Var(v) = &m.items[1] else { panic!() };
        assert!(v.declared);
    }

    #[test]
    fn error_recovery_keeps_tree_lossless() {
        for src in [
            "function f( { let x = ; }",
            "let = 3 +;",
            "class { foo(: int }",
            "if (x { y( }",
            "@@@",
            "a ? b",
            "`${`",
            "for (;;",
            "x = (1, 2",
        ] {
            let p = parse(src);
            assert!(!p.errors.is_empty(), "expected errors for {src:?}");
            assert_eq!(p.syntax().text().to_string(), src);
        }
    }
}

#[cfg(test)]
mod prop {
    use proptest::prelude::*;

    proptest! {
        #![proptest_config(ProptestConfig::with_cases(400))]

        /// The parser must terminate, never panic and stay lossless on arbitrary input.
        #[test]
        fn parser_is_total_and_lossless(src in "[ -~\\n\\t]{0,200}") {
            let p = super::parse(&src);
            prop_assert_eq!(p.syntax().text().to_string(), src);
        }

        #[test]
        fn parser_survives_token_soup(toks in proptest::collection::vec(
            prop_oneof![
                Just("function"), Just("class"), Just("("), Just(")"), Just("{"), Just("}"), Just("["), Just("]"),
                Just("<"), Just(">"), Just("="), Just("=>"), Just(","), Just(";"), Just(":"), Just("?"), Just("."),
                Just("x"), Just("1"), Just("\"s\""), Just("let"), Just("for"), Just("of"), Just("@d"), Just("export"),
                Just("declare"), Just("if"), Just("else"), Just("try"), Just("catch"), Just("new"), Just("as"),
                Just("+"), Just("++"), Just("!"), Just("switch"), Just("case"), Just("interface"), Just("enum"),
                Just("import"), Just("from"), Just("this"), Just("byref"), Just("return"), Just("`a${b}`"),
            ], 0..60)) {
            let src = toks.join(" ");
            let p = super::parse(&src);
            prop_assert_eq!(p.syntax().text().to_string(), src);
        }
    }
}
