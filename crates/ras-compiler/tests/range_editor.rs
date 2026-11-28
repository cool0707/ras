//! The editor queries of the Range members RAS adds (spec §9.4.1, §9.4.2): hover, parameter
//! information and completion.

#![allow(clippy::field_reassign_with_default)]

use std::path::PathBuf;

use ras_compiler::{analyze, query, Config, ModuleKind, Program, SourceInput};

fn analyze_one(src: &str) -> (Program, usize) {
    let input = SourceInput {
        path: PathBuf::from("src/m.ras"),
        rel_path: "src/m.ras".into(),
        key: "m".into(),
        kind: ModuleKind::User,
        text: src.into(),
    };
    let prog = analyze(vec![input], Config::default());
    let m = prog.modules.iter().position(|md| md.key == "m").unwrap();
    (prog, m)
}

/// The offset just after the first `marker` in `src`.
fn after(src: &str, marker: &str) -> u32 {
    (src.find(marker).unwrap_or_else(|| panic!("no {marker}")) + marker.len()) as u32
}

fn hover(src: &str, marker: &str) -> String {
    let (prog, m) = analyze_one(src);
    query::hover(&prog, m, after(src, marker))
        .map(|h| h.markdown)
        .unwrap_or_default()
}

/// The labels of the signatures, the active one and its active parameter.
fn help(src: &str, marker: &str) -> (Vec<String>, u32, Option<u32>) {
    let (prog, m) = analyze_one(src);
    let h = query::signature_help(&prog, m, after(src, marker)).expect("signature help");
    let active = h.signatures[h.active_signature as usize].active;
    (
        h.signatures.iter().map(|s| s.label.clone()).collect(),
        h.active_signature,
        active,
    )
}

const SRC: &str = r#"
export function f(r: Excel.Range, n: int, ws: Excel.Worksheet) {
    const a = r.extend(1, 2);
    const b = ws.range(2, 4, 8, 9);
    const c = ws.range("A1");
}
"#;

#[test]
fn hover_shows_the_added_signatures() {
    let h = hover(SRC, "r.exte");
    assert!(
        h.contains("(組み込み) extend(up?: int, down?: int, left?: int, right?: int): Excel.Range"),
        "{h}"
    );
    assert!(h.contains("`down` — 下に足す行数"), "{h}");
    for marker in ["b = ws.ran", "c = ws.ran"] {
        let h = hover(SRC, marker);
        assert!(
            h.contains("range(cell1: any, cell2?: any): Excel.Range\n"),
            "{h}"
        );
        assert!(
            h.contains("range(top: int, left: int, bottom: int, right: int): Excel.Range\n// VBA: Worksheet.Range"),
            "{h}"
        );
        assert!(h.contains("D2:I8"), "{h}");
    }
}

#[test]
fn parameter_information_has_the_overloads() {
    let (labels, active, param) = help(SRC, "r.extend(1, ");
    assert_eq!(
        labels,
        vec!["extend(up?: int, down?: int, left?: int, right?: int): Excel.Range"]
    );
    assert_eq!((active, param), (0, Some(1)));
    // The sheet's `range`: the host form and the form with numbers, chosen by the arguments.
    let (labels, active, param) = help(SRC, "ws.range(2, 4, ");
    assert_eq!(
        labels,
        vec![
            "range(cell1: any, cell2?: any): Excel.Range",
            "range(top: int, left: int, bottom: int, right: int): Excel.Range",
        ]
    );
    assert_eq!((active, param), (1, Some(2)));
    let (_, active, param) = help(SRC, "c = ws.range(");
    assert_eq!((active, param), (0, Some(0)));
    // Named arguments of `extend`.
    let src = "function g(r: Excel.Range) { const x = r.extend(down: 1, left: 2); }\n";
    let (_, _, param) = help(src, "left: ");
    assert_eq!(param, Some(2));
}

#[test]
fn completion_lists_extend_and_its_argument_names() {
    let src = "function g(r: Excel.Range) {\n    const x = r.extend(1, ";
    let (prog, m) = analyze_one(&format!("{src}\n}}\n"));
    let names: Vec<String> = query::named_arg_completions(&prog, m, src.len() as u32)
        .into_iter()
        .map(|c| c.label)
        .collect();
    assert_eq!(names, vec!["down:", "left:", "right:"]);
    let items = |src: &str| {
        let (prog, m) = analyze_one(&format!("{src}\n}}\n"));
        query::completions(&prog, m, src.len() as u32)
    };
    let list = items("function g(r: Excel.Range) {\n    r.");
    let ext = list
        .iter()
        .find(|c| c.label == "extend")
        .unwrap_or_else(|| panic!("{:?}", list.iter().map(|c| &c.label).collect::<Vec<_>>()));
    assert_eq!(
        ext.detail,
        "extend(up?: int, down?: int, left?: int, right?: int): Excel.Range"
    );
    let list = items("function g(ws: Excel.Worksheet) {\n    ws.");
    let range = list.iter().find(|c| c.label == "range").expect("range");
    assert!(
        range.detail.ends_with("（+1 個のオーバーロード）"),
        "{}",
        range.detail
    );
}

#[test]
fn last_cell_is_shown_as_a_property() {
    let src = "function g(ws: Excel.Worksheet) { const c = ws.usedRange.lastCell; }\n";
    let h = hover(src, "usedRange.last");
    assert!(h.contains("(組み込み) lastCell: Excel.Range"), "{h}");
    assert!(h.contains("右下のセル"), "{h}");
    let (prog, m) = analyze_one("function g(r: Excel.Range) {\n    r.\n}\n");
    let items = query::completions(
        &prog,
        m,
        "function g(r: Excel.Range) {\n    r.".len() as u32,
    );
    let c = items
        .iter()
        .find(|c| c.label == "lastCell")
        .expect("lastCell");
    assert_eq!(c.detail, "lastCell: Excel.Range");
}
