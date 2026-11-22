//! The code module of a form written from `forms/`: the marker line, the
//! code-behind and the generated code that fills ComboBox / ListBox lists.
//!
//! MS-Forms does not keep list items in the designer storage (Excel drops
//! them when the workbook is saved), so `list` and `selectedIndex` of a
//! FormSpec become VBA that runs in `UserForm_Initialize`:
//!
//! ```text
//! ' RAS-FORM forms/frmInput.form.json design=0123456789abcdef code=fedcba9876543210
//! (code-behind; a call of RAS_FormInit is added to its UserForm_Initialize)
//! '--- RAS BEGIN ---
//! Private Sub RAS_FormInit()
//!     With Me.cboKind
//!         .AddItem "A"
//!         .ListIndex = 0
//!     End With
//! End Sub
//! '--- RAS END ---
//! ```

use ras_ovba::{RAS_BEGIN_MARKER, RAS_FORM_MARKER};

use crate::oforms::Designer;
use crate::spec::FormSpec;

/// Name of the generated procedure that fills the lists.
pub const INIT_PROCEDURE: &str = "RAS_FormInit";
/// The call added to a hand-written `UserForm_Initialize`.
const INIT_CALL: &str = "RAS_FormInit '--- RAS ---";
/// Longest statement written on one line (the VBE accepts 1023 characters).
const MAX_LINE: usize = 1000;

/// 64-bit FNV-1a over byte strings, each followed by a separator.
struct Fnv(u64);

impl Fnv {
    fn new() -> Fnv {
        Fnv(0xcbf2_9ce4_8422_2325)
    }

    fn feed(&mut self, bytes: &[u8]) {
        for byte in bytes {
            self.0 ^= *byte as u64;
            self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
        }
        // Separator, so that ("ab", "c") and ("a", "bc") differ.
        self.0 ^= 0xff;
        self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
    }
}

/// A hash of a designer storage. It is recorded in the marker line
/// (`design=`), so that a sync into an open workbook can tell whether the
/// design changed without reading the designer back.
pub fn design_hash(designer: &Designer) -> u64 {
    let mut hash = Fnv::new();
    for (path, data) in &designer.streams {
        hash.feed(path.as_bytes());
        hash.feed(data);
    }
    for (path, clsid) in &designer.storages {
        hash.feed(path.as_bytes());
        hash.feed(clsid);
    }
    hash.0
}

/// Puts the marker line in front of the code. The marker also records a
/// hash of the code (`code=`): the VBA editor reformats what it stores, so
/// the stored text cannot be compared with the text written.
fn with_marker(rel_path: &str, designer: &Designer, code: &str) -> String {
    let mut hash = Fnv::new();
    hash.feed(code.as_bytes());
    format!(
        "{RAS_FORM_MARKER}{rel_path} design={:016x} code={:016x}\n{code}",
        design_hash(designer),
        hash.0
    )
}

/// A VBA expression for `text`. Characters a string literal cannot hold
/// (control characters, characters outside the project code page) are
/// written with `ChrW`.
fn vba_string(text: &str, codepage: u16) -> String {
    let mut parts: Vec<String> = Vec::new();
    let mut literal = String::new();
    let flush = |literal: &mut String, parts: &mut Vec<String>| {
        if !literal.is_empty() {
            parts.push(format!("\"{literal}\""));
            literal.clear();
        }
    };
    for ch in text.chars() {
        let plain = !ch.is_control()
            && (ch.is_ascii()
                || ras_ovba::encode_text(codepage, ch.encode_utf8(&mut [0; 4])).is_ok());
        if plain {
            if ch == '"' {
                literal.push('"');
            }
            literal.push(ch);
            continue;
        }
        flush(&mut literal, &mut parts);
        let mut units = [0u16; 2];
        for unit in ch.encode_utf16(&mut units) {
            parts.push(format!("ChrW({unit})"));
        }
    }
    flush(&mut literal, &mut parts);
    if parts.is_empty() {
        return "\"\"".into();
    }
    parts.join(" & ")
}

/// The generated procedure that fills the lists of `spec`; `None` when no
/// control has items or a selection.
pub fn init_procedure(spec: &FormSpec, codepage: u16) -> Result<Option<String>, String> {
    let mut body = String::new();
    for control in &spec.controls {
        let kind = control.control_type.trim();
        if !kind.eq_ignore_ascii_case("ComboBox") && !kind.eq_ignore_ascii_case("ListBox") {
            continue;
        }
        let items = control.list.as_deref().unwrap_or(&[]);
        let selected = control.selected_index.filter(|i| *i >= 0);
        if items.is_empty() && selected.is_none() {
            continue;
        }
        body.push_str(&format!("    With Me.{}\n", control.name));
        for item in items {
            // AddItem gives the first column; List(row, column) the others.
            let columns = item.columns();
            let mut lines = vec![format!(
                "        .AddItem {}\n",
                vba_string(columns.first().map_or("", String::as_str), codepage)
            )];
            for (column, text) in columns.iter().enumerate().skip(1) {
                lines.push(format!(
                    "        .List(.ListCount - 1, {column}) = {}\n",
                    vba_string(text, codepage)
                ));
            }
            for line in lines {
                if line.chars().count() > MAX_LINE {
                    return Err(format!(
                        "{} の list の項目が長すぎます（1 項目・1 列は VBA の 1 行に収まる長さにしてください）",
                        control.name
                    ));
                }
                body.push_str(&line);
            }
        }
        if let Some(index) = selected {
            body.push_str(&format!("        .ListIndex = {index}\n"));
        }
        body.push_str("    End With\n");
    }
    if body.is_empty() {
        return Ok(None);
    }
    Ok(Some(format!(
        "Private Sub {INIT_PROCEDURE}()\n{body}End Sub\n"
    )))
}

/// The code part of a line: without a trailing comment, string literals
/// blanked out.
fn code_part(line: &str) -> String {
    let mut out = String::with_capacity(line.len());
    let mut in_string = false;
    for ch in line.chars() {
        match ch {
            '"' => in_string = !in_string,
            '\'' if !in_string => break,
            _ if in_string => {}
            _ => out.push(ch),
        }
    }
    out
}

fn is_identifier_char(ch: char) -> bool {
    ch.is_alphanumeric() || ch == '_'
}

/// Whether `code` mentions the identifier `name` outside comments and strings.
fn mentions(code: &str, name: &str) -> bool {
    let name = name.to_lowercase();
    code.lines().any(|line| {
        let text = code_part(line).to_lowercase();
        let mut from = 0;
        while let Some(at) = text[from..].find(&name) {
            let start = from + at;
            let end = start + name.len();
            let before = text[..start].chars().next_back();
            let after = text[end..].chars().next();
            if !before.is_some_and(is_identifier_char) && !after.is_some_and(is_identifier_char) {
                return true;
            }
            from = end;
        }
        false
    })
}

/// Index of the last line of the declaration `Sub <name>(...)` in `lines`
/// (the declaration may continue over several lines with ` _`).
fn sub_declaration_end(lines: &[&str], name: &str) -> Option<usize> {
    for (index, line) in lines.iter().enumerate() {
        let text = code_part(line);
        let mut words = text.split_whitespace().peekable();
        while words.peek().is_some_and(|w| {
            ["private", "public", "friend", "static"].contains(&w.to_lowercase().as_str())
        }) {
            words.next();
        }
        if !words.next().is_some_and(|w| w.eq_ignore_ascii_case("Sub")) {
            continue;
        }
        let declared = words.next().unwrap_or("");
        let declared = declared.split('(').next().unwrap_or("");
        if !declared.eq_ignore_ascii_case(name) {
            continue;
        }
        let mut end = index;
        while end + 1 < lines.len() && code_part(lines[end]).trim_end().ends_with(" _") {
            end += 1;
        }
        return Some(end);
    }
    None
}

/// The whole code-pane text of a form written from `forms/`.
///
/// * `rel_path` — the `.form.json` file relative to the project root.
/// * `code` — the code of the form: what the compiler generated from
///   `forms/<Name>.ras`.
pub fn compose(
    rel_path: &str,
    designer: &Designer,
    spec: &FormSpec,
    code: &str,
    codepage: u16,
) -> Result<String, String> {
    let code = code.replace("\r\n", "\n").replace('\r', "\n");
    if code.lines().any(|l| l.trim() == RAS_BEGIN_MARKER) {
        return Err(format!(
            "コードビハインドに `{RAS_BEGIN_MARKER}` の行があります。この行は RAS が生成するコードの目印なので、書かないでください"
        ));
    }
    let Some(init) = init_procedure(spec, codepage)? else {
        return Ok(with_marker(rel_path, designer, &code));
    };
    if ras_ovba::procedure_names(&code)
        .iter()
        .any(|n| n.eq_ignore_ascii_case(INIT_PROCEDURE))
    {
        return Err(format!(
            "コードビハインドに `{INIT_PROCEDURE}` があります。この名前は list の項目を設定するために RAS が生成するプロシージャが使います"
        ));
    }
    let lines: Vec<&str> = code.lines().collect();
    let mut generated = init;
    match sub_declaration_end(&lines, "UserForm_Initialize") {
        // The hand-written handler fills the lists first, unless it says when.
        Some(end) if !mentions(&code, INIT_PROCEDURE) => {
            let mut body: String = lines[..=end].iter().map(|l| format!("{l}\n")).collect();
            body.push_str(&format!("    {INIT_CALL}\n"));
            body.extend(lines[end + 1..].iter().map(|l| format!("{l}\n")));
            let code = ras_ovba::merge_managed_region(&body, &generated);
            return Ok(with_marker(rel_path, designer, &code));
        }
        None if !mentions(&code, INIT_PROCEDURE) => generated.push_str(&format!(
            "\nPrivate Sub UserForm_Initialize()\n    {INIT_PROCEDURE}\nEnd Sub\n"
        )),
        // The code-behind calls the procedure itself.
        _ => {}
    }
    let mut body = code;
    if !body.is_empty() && !body.ends_with('\n') {
        body.push('\n');
    }
    // A module without code-behind would otherwise get CR LF line ends.
    let code = ras_ovba::merge_managed_region(&body, &generated).replace("\r\n", "\n");
    Ok(with_marker(rel_path, designer, &code))
}

/// The code-behind of a form module: `source` without its `Attribute`
/// header and without what [`compose`] added. For a form made in the VBE
/// this is simply its code-pane text. Lines end with LF.
pub fn code_behind(source: &str) -> String {
    let text = source.replace("\r\n", "\n").replace('\r', "\n");
    let mut lines: Vec<&str> = text
        .lines()
        .skip_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION "))
        .collect();
    if lines
        .first()
        .is_some_and(|l| l.starts_with(RAS_FORM_MARKER))
    {
        lines.remove(0);
    }
    let joined: String = lines.iter().map(|l| format!("{l}\n")).collect();
    let without_region = ras_ovba::merge_managed_region(&joined, "");
    let mut out: String = without_region
        .lines()
        .filter(|l| !l.trim().eq_ignore_ascii_case(INIT_CALL))
        .map(|l| format!("{l}\n"))
        .collect();
    while out.ends_with("\n\n") {
        out.pop();
    }
    if out.trim().is_empty() {
        out.clear();
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn spec(controls: &str) -> FormSpec {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{controls}]}}"#
        );
        let parsed = crate::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        parsed.spec.unwrap()
    }

    fn compose_for(spec: &FormSpec, code: &str) -> String {
        let designer = crate::compile(spec, 932).unwrap();
        compose("forms/F.form.json", &designer, spec, code, 932).unwrap()
    }

    const COMBO: &str =
        r#"{"id":"c","name":"cboKind","type":"ComboBox","list":["A","B \"x\""],"selectedIndex":1}"#;

    #[test]
    fn strings_become_vba_expressions() {
        assert_eq!(vba_string("", 932), "\"\"");
        assert_eq!(vba_string("a\"b", 932), "\"a\"\"b\"");
        assert_eq!(vba_string("日本語", 932), "\"日本語\"");
        assert_eq!(vba_string("a\tb", 932), "\"a\" & ChrW(9) & \"b\"");
        assert_eq!(vba_string("\r\n", 932), "ChrW(13) & ChrW(10)");
        // Not in code page 932: written by code point (a surrogate pair for U+1F600).
        assert_eq!(vba_string("é", 932), "ChrW(233)");
        assert_eq!(vba_string("😀", 932), "ChrW(55357) & ChrW(56832)");
        assert_eq!(vba_string("é", 1252), "\"é\"");
    }

    #[test]
    fn a_form_without_lists_gets_only_the_marker() {
        let spec = spec(r#"{"id":"t","name":"txt","type":"TextBox"}"#);
        let code = compose_for(&spec, "Private Sub txt_Change()\nEnd Sub\n");
        let mut lines = code.lines();
        let marker = lines.next().unwrap();
        assert!(
            marker.starts_with("' RAS-FORM forms/F.form.json design="),
            "{marker}"
        );
        let fields: Vec<&str> = marker.split(' ').skip(3).collect();
        assert_eq!(fields.len(), 2, "{marker}");
        assert!(
            fields[0].starts_with("design=") && fields[0].len() == 7 + 16,
            "{marker}"
        );
        assert!(
            fields[1].starts_with("code=") && fields[1].len() == 5 + 16,
            "{marker}"
        );
        // The code hash follows the code, the design hash does not.
        let other = compose_for(&spec, "Private Sub txt_Change()\n    Beep\nEnd Sub\n");
        let other: Vec<&str> = other.lines().next().unwrap().split(' ').skip(3).collect();
        assert_eq!(other[0], fields[0]);
        assert_ne!(other[1], fields[1]);
        assert_eq!(
            lines.collect::<Vec<_>>(),
            ["Private Sub txt_Change()", "End Sub"]
        );
        assert_eq!(code_behind(&code), "Private Sub txt_Change()\nEnd Sub\n");
    }

    #[test]
    fn items_of_several_columns_fill_the_other_columns_by_list() {
        let spec = spec(
            r#"{"id":"l","name":"lstStaff","type":"ListBox","columnCount":3,"list":[["001","山田","営業"],"002",["003",""]]}"#,
        );
        let code = init_procedure(&spec, 932).unwrap().unwrap();
        assert_eq!(
            code,
            "Private Sub RAS_FormInit()\n    With Me.lstStaff\n        .AddItem \"001\"\n        .List(.ListCount - 1, 1) = \"山田\"\n        .List(.ListCount - 1, 2) = \"営業\"\n        .AddItem \"002\"\n        .AddItem \"003\"\n        .List(.ListCount - 1, 1) = \"\"\n    End With\nEnd Sub\n"
        );
        // 1 to 10 columns.
        for bad in [
            r#"[]"#,
            r#"["1","2","3","4","5","6","7","8","9","10","11"]"#,
        ] {
            let text = format!(
                r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{{"id":"l","name":"lst","type":"ListBox","list":[{bad}]}}]}}"#
            );
            assert!(crate::parse(&text).has_errors(), "{bad}");
        }
    }

    #[test]
    fn lists_are_filled_by_a_generated_initialize() {
        let spec = spec(COMBO);
        let code = compose_for(&spec, "Option Explicit\n");
        let expected = "Option Explicit\n'--- RAS BEGIN ---\nPrivate Sub RAS_FormInit()\n    With Me.cboKind\n        .AddItem \"A\"\n        .AddItem \"B \"\"x\"\"\"\n        .ListIndex = 1\n    End With\nEnd Sub\n\nPrivate Sub UserForm_Initialize()\n    RAS_FormInit\nEnd Sub\n'--- RAS END ---\n";
        assert_eq!(code.split_once('\n').unwrap().1, expected);
        assert_eq!(code_behind(&code), "Option Explicit\n");
        // No code-behind at all.
        let bare = compose_for(&spec, "");
        assert!(bare.contains("Private Sub UserForm_Initialize()"));
        assert_eq!(code_behind(&bare), "");
    }

    #[test]
    fn a_hand_written_initialize_calls_the_generated_procedure() {
        let spec = spec(COMBO);
        let hand = "Option Explicit\n\nPrivate Sub UserForm_Initialize()\n    Me.Caption = \"x\"\nEnd Sub\n";
        let code = compose_for(&spec, hand);
        assert!(
            code.contains("Private Sub UserForm_Initialize()\n    RAS_FormInit '--- RAS ---\n    Me.Caption = \"x\"\nEnd Sub\n"),
            "{code}"
        );
        assert_eq!(code.matches("Sub UserForm_Initialize").count(), 1);
        assert!(code.contains("Private Sub RAS_FormInit()"));
        assert_eq!(code_behind(&code), hand);

        // The handler decides when the lists are filled.
        let own = "Private Sub UserForm_Initialize()\n    Prepare\n    RAS_FormInit\nEnd Sub\n";
        let code = compose_for(&spec, own);
        assert_eq!(code.matches("RAS_FormInit").count(), 2, "{code}");
        assert_eq!(code_behind(&code), own);

        // Called from somewhere else: no handler is generated either.
        let elsewhere = "Public Sub Reload()\n    RAS_FormInit\nEnd Sub\n";
        assert!(!compose_for(&spec, elsewhere).contains("UserForm_Initialize"));

        // A mention in a comment or a string is not a call.
        let comment =
            "Private Sub UserForm_Initialize() ' RAS_FormInit\n    x = \"RAS_FormInit\"\nEnd Sub\n";
        assert!(compose_for(&spec, comment).contains("RAS_FormInit '--- RAS ---"));
    }

    #[test]
    fn names_reserved_for_generated_code_are_refused() {
        let spec = spec(COMBO);
        let designer = crate::compile(&spec, 932).unwrap();
        let own = "Private Sub RAS_FormInit()\nEnd Sub\n";
        assert!(compose("forms/F.form.json", &designer, &spec, own, 932).is_err());
        let marker = "'--- RAS BEGIN ---\n";
        assert!(compose("forms/F.form.json", &designer, &spec, marker, 932).is_err());
    }

    #[test]
    fn the_hash_follows_the_design() {
        let a = crate::compile(&spec(COMBO), 932).unwrap();
        let moved = COMBO.replace("\"name\":\"cboKind\"", "\"name\":\"cboKind\",\"left\":30");
        let b = crate::compile(&spec(&moved), 932).unwrap();
        assert_eq!(design_hash(&a), design_hash(&a.clone()));
        assert_ne!(design_hash(&a), design_hash(&b));
        // List items are not part of the design.
        let items = COMBO.replace("\"A\"", "\"Z\"");
        assert_eq!(
            design_hash(&a),
            design_hash(&crate::compile(&spec(&items), 932).unwrap())
        );
    }

    #[test]
    fn code_behind_of_a_form_made_in_the_vbe() {
        let source = "Attribute VB_Name = \"F\"\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\n\r\nSub A()\r\nEnd Sub\r\n\r\n\r\n";
        assert_eq!(code_behind(source), "Option Explicit\n\nSub A()\nEnd Sub\n");
        assert_eq!(code_behind("Attribute VB_Name = \"F\"\r\n"), "");
    }
}
