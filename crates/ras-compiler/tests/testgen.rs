//! The skeleton of a test for a function (`testgen`, spec §14.12.10): the generated test files
//! compile without errors, and generating into an existing file adds to it.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::testgen;
use ras_compiler::{Config, ModuleKind, Program, SourceInput};

fn source(path: &str, text: &str) -> SourceInput {
    let test = path.starts_with("test/");
    let key = if test {
        format!(
            "{}{}",
            ras_compiler::collect::TEST_KEY_PREFIX,
            path.strip_suffix(".ras").unwrap()
        )
    } else {
        path["src/".len()..]
            .strip_suffix(".ras")
            .unwrap()
            .to_string()
    };
    SourceInput {
        path: PathBuf::from(path),
        rel_path: path.to_string(),
        key,
        kind: ModuleKind::User,
        text: text.to_string(),
    }
}

fn analyze(files: &[(&str, &str)]) -> Program {
    let mut cfg = Config::default();
    cfg.sheet_code_names = vec!["Sheet1".into()];
    for (p, _) in files {
        if cfg.test.is_test_file(p) {
            cfg.test_files.push(p.to_string());
        }
    }
    ras_compiler::analyze(files.iter().map(|(p, t)| source(p, t)).collect(), cfg)
}

fn errors(prog: &Program) -> Vec<String> {
    prog.diagnostics
        .iter()
        .filter(|d| d.is_error())
        .map(|d| {
            let path = d.module.map(|m| prog.modules[m].rel_path.clone());
            format!("{path:?} {}: {}", d.code, d.message)
        })
        .collect()
}

fn fn_named(prog: &Program, name: &str) -> ras_compiler::model::FnId {
    prog.fns
        .iter()
        .position(|f| f.name == name && prog.modules[f.module].rel_path.starts_with("src/"))
        .unwrap_or_else(|| panic!("no function {name}"))
}

/// Generates the tests of `names` one after the other into the test file, checking that the
/// project compiles after each, and returns the test file.
fn generate_all(src_path: &str, src: &str, names: &[&str]) -> String {
    let mut test: Option<String> = None;
    let mut test_path = String::new();
    for name in names {
        let mut files = vec![(src_path, src)];
        if let Some(t) = &test {
            files.push((&test_path, t));
        }
        let prog = analyze(&files);
        assert!(errors(&prog).is_empty(), "{:#?}", errors(&prog));
        let f = fn_named(&prog, name);
        let g = testgen::generate(&prog, f, test.as_deref()).expect(name);
        assert_eq!(g.created, test.is_none());
        assert_eq!(
            testgen::apply(test.as_deref().unwrap_or(""), &g.edits),
            g.text
        );
        let line = g.text.lines().nth(g.line as usize).unwrap_or_default();
        assert!(
            line.trim_start().starts_with("test("),
            "line {}: {line}\n{}",
            g.line,
            g.text
        );
        test_path = g.path;
        test = Some(g.text);
    }
    let text = test.unwrap();
    let prog = analyze(&[(src_path, src), (&test_path, &text)]);
    assert!(errors(&prog).is_empty(), "{:#?}\n{text}", errors(&prog));
    text
}

const SRC: &str = r#"
export enum Kind { A, B }
export struct Point { x: int; y: double; label?: string }
export class Counter { n: int = 0; }

export function sumTo(n: int): int {
    let s = 0;
    for (let i = 1; i <= n; i++) s += i;
    return s;
}
export function half(v: double): double { return v / 2; }
export function greet(name: string, polite: boolean = true): string { return name; }
export function fill() { Sheet1.range("A1").value = 1; }
export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}
export function isBlank(this: string): boolean { return this.trim().length == 0; }
export function kindOf(k: Kind): Kind { return k; }
export function move(p: Point): Point { return p; }
export function names(r: Excel.Range): string[] { return []; }
export function makeCounter(): Counter { return new Counter(); }
export function apply(xs: int[], f: (x: int) => int): int[] { return xs.map(f); }
export function cell(r: Excel.Range): cellValue { return r.value; }
export function today(d: date): date { return d; }
function hidden(): int { return 1; }
"#;

#[test]
fn generated_tests_compile() {
    let names = [
        "sumTo",
        "half",
        "greet",
        "fill",
        "lastRow",
        "isBlank",
        "kindOf",
        "move",
        "names",
        "makeCounter",
        "apply",
        "cell",
        "today",
    ];
    let text = generate_all("src/util.ras", SRC, &names);
    assert!(text.starts_with("import { sumTo, half, greet, fill, lastRow, isBlank, kindOf, Kind, move, Point, names, makeCounter, apply, cell, today } from \"src/util\";\n"), "{text}");
    for n in names {
        assert!(
            text.contains(&format!("describe(\"{n}\", () => {{")),
            "{n}\n{text}"
        );
    }
    // An extension method is called on a receiver; the optional arguments are left out.
    assert!(
        text.contains(
            "const target: Excel.Worksheet = Sheet1;\n\n        const actual = target.lastRow();"
        ),
        "{text}"
    );
    assert!(text.contains("const actual = greet(name);"), "{text}");
    assert!(
        text.contains("const p: Point = { x: 0, y: 0.0 };"),
        "{text}"
    );
}

#[test]
fn only_exported_functions_of_src() {
    let prog = analyze(&[("src/util.ras", SRC)]);
    let f = fn_named(&prog, "hidden");
    assert!(!testgen::can_generate(&prog, f));
    let m = prog.fns[f].module;
    let names: Vec<&str> = testgen::targets(&prog, m)
        .iter()
        .map(|&f| prog.fns[f].name.as_str())
        .collect();
    assert!(names.contains(&"sumTo") && !names.contains(&"hidden"));
}

#[test]
fn new_file() {
    let prog = analyze(&[(
        "src/a/calc.ras",
        "export function sumTo(n: int): int { return n; }\n",
    )]);
    let g = testgen::generate(&prog, fn_named(&prog, "sumTo"), None).unwrap();
    assert_eq!(g.path, "test/a/calc.test.ras");
    assert_eq!(
        g.text,
        r#"import { sumTo } from "src/a/calc";

describe("sumTo", () => {
    test("sumTo の戻り値", () => {
        // 準備: 引数の値を書き換える
        const n: int = 0;

        const actual = sumTo(n);

        // 確認: 期待する値に書き換える
        expect(actual).toBe(0);
    });
    test.todo("境界の値（0・空の文字列・最大の値など）");
});
"#
    );
    assert_eq!(g.line, 3);
}

#[test]
fn existing_file_gets_the_import_and_the_test_in_its_describe() {
    let src = "export function sumTo(n: int): int { return n; }\nexport function fill() {}\n";
    let existing = r#"// テスト
import { sumTo } from "../src/main";

describe("sumTo", () => {
    test("0 なら 0", () => {
        expect(sumTo(0)).toBe(0);
    });
});
"#;
    let prog = analyze(&[("src/main.ras", src), ("test/main.test.ras", existing)]);
    let g = testgen::generate(&prog, fn_named(&prog, "sumTo"), Some(existing)).unwrap();
    assert!(!g.created);
    // Already imported; the test goes at the end of `describe("sumTo")`.
    assert_eq!(g.edits.len(), 1);
    assert!(
        g.text
            .contains("    });\n\n    test(\"sumTo の戻り値\", () => {"),
        "{}",
        g.text
    );
    assert!(
        g.text
            .ends_with("    test.todo(\"境界の値（0・空の文字列・最大の値など）\");\n});\n"),
        "{}",
        g.text
    );
    assert_eq!(
        g.text.lines().nth(g.line as usize),
        Some("    test(\"sumTo の戻り値\", () => {")
    );

    let g = testgen::generate(&prog, fn_named(&prog, "fill"), Some(existing)).unwrap();
    assert!(
        g.text
            .starts_with("// テスト\nimport { sumTo, fill } from \"../src/main\";\n"),
        "{}",
        g.text
    );
    assert!(g.text.ends_with("});\n\ndescribe(\"fill\", () => {\n    test(\"fill を実行する\", () => {\n        fill();\n\n        // 確認: 実行した後の状態（セルの値など）を確かめる。例:\n        // expect(Sheet1.range(\"A1\").value).toBe(\"期待する値\");\n    });\n});\n"), "{}", g.text);
    assert_eq!(
        g.text.lines().nth(g.line as usize),
        Some("    test(\"fill を実行する\", () => {")
    );
}

#[test]
fn existing_file_without_the_import() {
    let src = "export function sumTo(n: int): int { return n; }\n";
    let existing = "describe(\"other\", () => {});\n";
    let prog = analyze(&[("src/main.ras", src)]);
    let g = testgen::generate(&prog, fn_named(&prog, "sumTo"), Some(existing)).unwrap();
    assert!(
        g.text
            .starts_with("import { sumTo } from \"src/main\";\n\ndescribe(\"other\""),
        "{}",
        g.text
    );
    assert_eq!(
        g.text.lines().nth(g.line as usize),
        Some("    test(\"sumTo の戻り値\", () => {")
    );
}
