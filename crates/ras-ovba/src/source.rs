//! Module source text helpers: line-ending normalisation, the leading
//! `Attribute VB_...` header and the RAS-managed region of document modules.

use crate::project::split_lines;
use crate::ModuleKind;

/// First line of the RAS-managed region inside a document module.
pub const RAS_BEGIN_MARKER: &str = "'--- RAS BEGIN ---";
/// Last line of the RAS-managed region inside a document module.
pub const RAS_END_MARKER: &str = "'--- RAS END ---";

/// `VB_Base` of the workbook document module (`ThisWorkbook`).
pub(crate) const VB_BASE_WORKBOOK: &str = "\"0{00020819-0000-0000-C000-000000000046}\"";
/// `VB_Base` of a worksheet document module.
pub(crate) const VB_BASE_WORKSHEET: &str = "\"0{00020820-0000-0000-C000-000000000046}\"";
/// `VB_Base` of a chart sheet document module.
pub(crate) const VB_BASE_CHART: &str = "\"0{00020821-0000-0000-C000-000000000046}\"";
/// `VB_Base` of a class module. Excel reports "Invalid procedure call or
/// argument" for the whole project when a class module stream lacks it.
pub(crate) const VB_BASE_CLASS: &str = "\"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"";

/// Converts every line terminator (CR LF, LF, CR) to CR LF and makes sure
/// non-empty text ends with one.
pub(crate) fn normalize_newlines(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 16);
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '\r' => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
                out.push_str("\r\n");
            }
            '\n' => out.push_str("\r\n"),
            other => out.push(other),
        }
    }
    if !out.is_empty() && !out.ends_with("\r\n") {
        out.push_str("\r\n");
    }
    out
}

/// Compares two module sources ignoring line-terminator style and trailing
/// blank lines / trailing whitespace at the very end.
pub(crate) fn sources_equivalent(a: &str, b: &str) -> bool {
    fn canon(s: &str) -> String {
        let n = normalize_newlines(s);
        n.trim_end().to_string()
    }
    a == b || canon(a) == canon(b)
}

/// Parses `Attribute VB_Xxx = value` (a module-level attribute line).
/// Returns `(key, value)`; member attributes (`Attribute Foo.VB_...`) are
/// not module-level and return `None`.
fn parse_attribute(line: &str) -> Option<(&str, &str)> {
    let rest = line.trim_start();
    let keyword = rest.get(..9)?;
    if !keyword.eq_ignore_ascii_case("attribute") {
        return None;
    }
    let rest = &rest[9..];
    if !rest.starts_with([' ', '\t']) {
        return None;
    }
    let (key, value) = rest.split_once('=')?;
    let key = key.trim();
    let is_vb = key.get(..3).is_some_and(|p| p.eq_ignore_ascii_case("VB_"));
    if !is_vb || key.contains(['.', ' ', '\t']) {
        return None;
    }
    Some((key, value.trim()))
}

/// The leading attribute lines of a (CR LF normalised) module source and
/// the byte offset where the rest of the source starts.
fn leading_attributes(source: &str) -> (Vec<(String, String)>, usize) {
    let mut attributes = Vec::new();
    let mut offset = 0usize;
    for line in source.split_inclusive("\r\n") {
        match parse_attribute(line.trim_end_matches(['\r', '\n'])) {
            Some((key, value)) => {
                attributes.push((key.to_string(), value.to_string()));
                offset += line.len();
            }
            None => break,
        }
    }
    (attributes, offset)
}

/// Value of the leading attribute `key` in `source`, if present.
pub(crate) fn attribute_value(source: &str, key: &str) -> Option<String> {
    let normalized = normalize_newlines(source);
    let (attributes, _) = leading_attributes(&normalized);
    attributes.into_iter().find(|(k, _)| k.eq_ignore_ascii_case(key)).map(|(_, v)| v)
}

/// Removes the preamble that `File > Export` puts in front of the
/// attributes of class modules and forms and that is not part of the stored
/// module text:
///
/// ```text
/// VERSION 1.0 CLASS
/// BEGIN
///   MultiUse = -1  'True
/// END
/// ```
fn strip_export_preamble(source: &str) -> &str {
    let mut lines = source.split_inclusive("\r\n");
    let Some(first) = lines.next() else {
        return source;
    };
    let first_trim = first.trim();
    if !first_trim.get(..8).is_some_and(|p| p.eq_ignore_ascii_case("VERSION ")) {
        return source;
    }
    let mut consumed = first.len();
    let mut depth = 0usize;
    let mut saw_block = false;
    for line in lines {
        let t = line.trim();
        let word = t.split_whitespace().next().unwrap_or("");
        if word.eq_ignore_ascii_case("begin") {
            depth += 1;
            saw_block = true;
            consumed += line.len();
        } else if depth > 0 {
            consumed += line.len();
            if t.eq_ignore_ascii_case("end") {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
        } else {
            break;
        }
    }
    if saw_block && depth != 0 {
        // Unbalanced block: do not guess.
        return source;
    }
    &source[consumed..]
}

/// The attributes every module of `kind` starts with, in the order Excel
/// writes them, with their default values. `VB_Name` is always first and its
/// value comes from the module name.
fn required_attributes(kind: ModuleKind, doc_base: &str) -> Vec<(&'static str, String)> {
    let f = |s: &str| s.to_string();
    match kind {
        ModuleKind::Standard => vec![("VB_Name", String::new())],
        ModuleKind::Class => vec![
            ("VB_Name", String::new()),
            ("VB_Base", f(VB_BASE_CLASS)),
            ("VB_GlobalNameSpace", f("False")),
            ("VB_Creatable", f("False")),
            ("VB_PredeclaredId", f("False")),
            ("VB_Exposed", f("False")),
            ("VB_TemplateDerived", f("False")),
            ("VB_Customizable", f("False")),
        ],
        ModuleKind::Document => vec![
            ("VB_Name", String::new()),
            ("VB_Base", f(doc_base)),
            ("VB_GlobalNameSpace", f("False")),
            ("VB_Creatable", f("False")),
            ("VB_PredeclaredId", f("True")),
            ("VB_Exposed", f("True")),
            ("VB_TemplateDerived", f("False")),
            ("VB_Customizable", f("True")),
        ],
        ModuleKind::Form => vec![
            ("VB_Name", String::new()),
            ("VB_Base", f(doc_base)),
            ("VB_GlobalNameSpace", f("False")),
            ("VB_Creatable", f("False")),
            ("VB_PredeclaredId", f("True")),
            ("VB_Exposed", f("False")),
            ("VB_TemplateDerived", f("False")),
            ("VB_Customizable", f("False")),
        ],
    }
}

/// Produces the text that is stored in a module stream.
///
/// * line terminators become CR LF,
/// * an export preamble (`VERSION 1.0 CLASS` / `BEGIN` ... `END`) is dropped,
/// * `Attribute VB_Name` is forced to `name`,
/// * missing header attributes are added: for standard and class modules
///   with the defaults Excel uses, for document modules and forms with the
///   value of the module being replaced (`previous`) when there is one. When
///   something is missing the required attributes are written first, in
///   Excel's order, followed by any additional attributes of the caller;
///   a complete header is kept exactly as given.
/// * A document module or form whose new text has no attribute header at
///   all inherits the complete header of `previous`.
pub(crate) fn prepare_module_source(
    kind: ModuleKind,
    name: &str,
    source: &str,
    previous: Option<&str>,
    doc_base: &str,
) -> String {
    let normalized = normalize_newlines(source);
    let stripped = strip_export_preamble(&normalized);
    let (given, body_start) = leading_attributes(stripped);
    let body = &stripped[body_start..];

    let previous_attributes = previous
        .map(|p| leading_attributes(&normalize_newlines(p)).0)
        .unwrap_or_default();
    let find = |list: &[(String, String)], key: &str| -> Option<String> {
        list.iter().find(|(k, _)| k.eq_ignore_ascii_case(key)).map(|(_, v)| v.clone())
    };
    let name_value = format!("\"{name}\"");
    let inherits = matches!(kind, ModuleKind::Document | ModuleKind::Form);

    let mut header: Vec<(String, String)> = Vec::new();
    if given.is_empty() && inherits && !previous_attributes.is_empty() {
        header = previous_attributes.clone();
    } else {
        let required = required_attributes(kind, doc_base);
        let complete = required.iter().all(|(key, _)| find(&given, key).is_some());
        if complete {
            header = given.clone();
        } else {
            for (key, default) in &required {
                let value = find(&given, key)
                    .or_else(|| if inherits { find(&previous_attributes, key) } else { None })
                    .unwrap_or_else(|| default.clone());
                header.push((key.to_string(), value));
            }
            for (key, value) in &given {
                if !required.iter().any(|(k, _)| k.eq_ignore_ascii_case(key)) {
                    header.push((key.clone(), value.clone()));
                }
            }
        }
    }
    let mut has_name = false;
    for (key, value) in header.iter_mut() {
        if key.eq_ignore_ascii_case("VB_Name") {
            *value = name_value.clone();
            has_name = true;
        }
    }
    if !has_name {
        header.insert(0, ("VB_Name".to_string(), name_value));
    }

    let mut out = String::with_capacity(stripped.len() + 256);
    for (key, value) in &header {
        out.push_str("Attribute ");
        out.push_str(key);
        out.push_str(" = ");
        out.push_str(value);
        out.push_str("\r\n");
    }
    out.push_str(body);
    out
}

/// Source of a freshly created, empty document module.
pub(crate) fn default_document_source(name: &str, doc_base: &str) -> String {
    prepare_module_source(ModuleKind::Document, name, "", None, doc_base)
}

fn newline_style(text: &str) -> &'static str {
    if text.contains("\r\n") {
        "\r\n"
    } else if text.contains('\n') {
        "\n"
    } else if text.contains('\r') {
        "\r"
    } else {
        "\r\n"
    }
}

/// Finds the marker lines: `(begin, Some(end))` for a complete region,
/// `(begin, None)` when the end marker is missing.
fn find_region(lines: &[&str]) -> Option<(usize, Option<usize>)> {
    let begin = lines.iter().position(|l| l.trim() == RAS_BEGIN_MARKER)?;
    let end = lines[begin + 1..]
        .iter()
        .position(|l| l.trim() == RAS_END_MARKER)
        .map(|p| begin + 1 + p);
    Some((begin, end))
}

/// Replaces the RAS-managed region of a document module, or appends one.
///
/// The region is the block between the marker lines `'--- RAS BEGIN ---`
/// and `'--- RAS END ---` (both inclusive). Everything outside the region —
/// the `Attribute` header and hand-written code — is preserved.
///
/// * If `existing_source` has a region, it is replaced in place.
/// * Otherwise the region is appended at the end.
/// * If `managed_code` is empty (or only whitespace) the region, including
///   its markers, is removed.
/// * A begin marker without an end marker is treated as a region that
///   consists of that single line, so no hand-written code is ever dropped.
///
/// The result uses the line terminator of `existing_source` (CR LF when it
/// has none) and ends with a line terminator whenever it is not empty. The
/// function is idempotent: merging the same code twice changes nothing.
pub fn merge_managed_region(existing_source: &str, managed_code: &str) -> String {
    let nl = newline_style(existing_source);
    let lines = split_lines(existing_source);
    let mut managed: Vec<&str> = split_lines(managed_code);
    while managed.last().is_some_and(|l| l.trim().is_empty()) {
        managed.pop();
    }
    let remove = managed.is_empty();
    let region = find_region(&lines);
    if remove && region.is_none() {
        return existing_source.to_string();
    }

    let mut block: Vec<&str> = Vec::new();
    if !remove {
        block.push(RAS_BEGIN_MARKER);
        block.extend(managed.iter().copied());
        block.push(RAS_END_MARKER);
    }
    let (before, after): (&[&str], &[&str]) = match region {
        Some((begin, Some(end))) => (&lines[..begin], &lines[end + 1..]),
        Some((begin, None)) => (&lines[..begin], &lines[begin + 1..]),
        None => (&lines[..], &[]),
    };

    let mut out = String::with_capacity(existing_source.len() + managed_code.len() + 64);
    for line in before.iter().chain(block.iter()).chain(after.iter()) {
        out.push_str(line);
        out.push_str(nl);
    }
    out
}

/// Names of the procedures declared in `code` (`Sub`, `Function`,
/// `Property`), in order of appearance. Comments are not looked into; a
/// declaration whose name sits on a continuation line is not found.
pub(crate) fn procedure_names(code: &str) -> Vec<String> {
    let mut names = Vec::new();
    for line in split_lines(code) {
        let mut words = line.split_whitespace().peekable();
        while words.peek().is_some_and(|w| {
            ["public", "private", "friend", "static"].iter().any(|m| w.eq_ignore_ascii_case(m))
        }) {
            words.next();
        }
        let Some(keyword) = words.next() else { continue };
        if keyword.eq_ignore_ascii_case("property") {
            let accessor = words.next().unwrap_or("");
            if !["get", "let", "set"].iter().any(|a| accessor.eq_ignore_ascii_case(a)) {
                continue;
            }
        } else if !keyword.eq_ignore_ascii_case("sub") && !keyword.eq_ignore_ascii_case("function") {
            continue;
        }
        let name = words.next().unwrap_or("").split('(').next().unwrap_or("");
        if !name.is_empty() {
            names.push(name.to_string());
        }
    }
    names
}

/// Returns the code inside the RAS-managed region (without the marker
/// lines and without a trailing line terminator), or `None` when the source
/// has no complete region.
pub fn extract_managed_region(source: &str) -> Option<String> {
    let nl = newline_style(source);
    let lines = split_lines(source);
    match find_region(&lines)? {
        (begin, Some(end)) => Some(lines[begin + 1..end].join(nl)),
        (_, None) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn procedures_are_found() {
        let code = "Option Explicit\r\n' Sub Commented()\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n  public static function F(x) As Long\r\nEnd Function\r\nFriend Property Get P() As Long\r\nEnd Property\r\nProperty Let P(v As Long)\r\nEnd Property\r\nSub NoParens\r\nExit Sub\r\nDim Subtotal As Long\r\n";
        assert_eq!(procedure_names(code), ["Workbook_Open", "F", "P", "P", "NoParens"]);
    }

    #[test]
    fn newline_normalisation() {
        assert_eq!(normalize_newlines("a\nb\r\nc\rd"), "a\r\nb\r\nc\r\nd\r\n");
        assert_eq!(normalize_newlines(""), "");
        assert_eq!(normalize_newlines("x\r\n"), "x\r\n");
        assert!(sources_equivalent("a\nb", "a\r\nb\r\n\r\n"));
        assert!(!sources_equivalent("a\nb", "a\nc"));
    }

    #[test]
    fn standard_module_gets_its_name() {
        let s = prepare_module_source(ModuleKind::Standard, "Mod", "Sub A()\nEnd Sub", None, "");
        assert_eq!(s, "Attribute VB_Name = \"Mod\"\r\nSub A()\r\nEnd Sub\r\n");
        // A wrong name is corrected, the rest of the header is kept.
        let s = prepare_module_source(
            ModuleKind::Standard,
            "Mod",
            "Attribute VB_Name = \"Other\"\r\nAttribute VB_Description = \"d\"\r\nOption Explicit\r\n",
            None,
            "",
        );
        assert_eq!(
            s,
            "Attribute VB_Name = \"Mod\"\r\nAttribute VB_Description = \"d\"\r\nOption Explicit\r\n"
        );
    }

    #[test]
    fn class_module_defaults() {
        let base = "Attribute VB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\n";
        let tail = "Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n";
        let s = prepare_module_source(ModuleKind::Class, "C", "Public X As Long\n", None, "");
        assert_eq!(
            s,
            format!("Attribute VB_Name = \"C\"\r\n{base}Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n{tail}Public X As Long\r\n")
        );
        // Given values win; missing ones are filled in.
        let s = prepare_module_source(
            ModuleKind::Class,
            "C",
            "Attribute VB_PredeclaredId = True\nAttribute VB_Description = \"x\"\nOption Explicit\n",
            None,
            "",
        );
        assert_eq!(
            s,
            format!("Attribute VB_Name = \"C\"\r\n{base}Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\n{tail}Attribute VB_Description = \"x\"\r\nOption Explicit\r\n")
        );
        // The header of an exported .cls has no VB_Base: it is completed in Excel's order.
        let exported = "Attribute VB_Name = \"C\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\n";
        assert_eq!(prepare_module_source(ModuleKind::Class, "C", exported, None, ""), format!("Attribute VB_Name = \"C\"\r\n{base}Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n{tail}Option Explicit\r\n"));
        // A complete header is kept verbatim (order included).
        let full = format!("Attribute VB_Name = \"C\"\r\nAttribute VB_Exposed = True\r\n{base}Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\n{tail}Option Explicit\r\n");
        assert_eq!(prepare_module_source(ModuleKind::Class, "C", &full, None, ""), full);
    }

    #[test]
    fn export_preamble_is_dropped() {
        let cls = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"C\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\n";
        let s = prepare_module_source(ModuleKind::Class, "C", cls, None, "");
        assert!(s.starts_with("Attribute VB_Name = \"C\"\r\n"));
        assert!(s.ends_with("Attribute VB_Customizable = False\r\nOption Explicit\r\n"));
        assert!(!s.contains("VERSION"));

        let frm = "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F \r\n   Caption = \"F\"\r\n   OleObjectBlob = \"F.frx\":0000\r\nEnd\r\nAttribute VB_Name = \"F\"\r\nSub X()\r\nEnd Sub\r\n";
        let prev = "Attribute VB_Name = \"F\"\r\nAttribute VB_Base = \"0{A}{B}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n";
        let s = prepare_module_source(ModuleKind::Form, "F", frm, Some(prev), "");
        assert_eq!(s, format!("{prev}Sub X()\r\nEnd Sub\r\n"));
    }

    #[test]
    fn document_header() {
        let s = default_document_source("ThisWorkbook", VB_BASE_WORKBOOK);
        assert_eq!(
            s,
            "Attribute VB_Name = \"ThisWorkbook\"\r\nAttribute VB_Base = \"0{00020819-0000-0000-C000-000000000046}\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n"
        );
        // Code without a header inherits the previous header completely.
        let prev = format!("{s}Attribute VB_Control = \"Btn, 1, 0, MSForms, CommandButton\"\r\nOld\r\n");
        let s2 = prepare_module_source(ModuleKind::Document, "ThisWorkbook", "New\n", Some(&prev), VB_BASE_WORKSHEET);
        assert_eq!(s2, format!("{s}Attribute VB_Control = \"Btn, 1, 0, MSForms, CommandButton\"\r\nNew\r\n"));
        assert_eq!(attribute_value(&s2, "vb_base").as_deref(), Some(VB_BASE_WORKBOOK));
    }

    #[test]
    fn member_attributes_are_not_header() {
        let src = "Attribute VB_Name = \"M\"\r\nSub A()\r\nAttribute A.VB_Description = \"x\"\r\nEnd Sub\r\n";
        assert_eq!(prepare_module_source(ModuleKind::Standard, "M", src, None, ""), src);
        assert!(parse_attribute("Attribute A.VB_Description = \"x\"").is_none());
        assert!(parse_attribute("AttributeVB_Name = 1").is_none());
        assert!(parse_attribute("' Attribute VB_Name = 1").is_none());
    }

    #[test]
    fn merge_appends_replaces_and_removes() {
        let base = "Attribute VB_Name = \"Sheet1\"\r\nPrivate Sub Hand()\r\nEnd Sub\r\n";
        let merged = merge_managed_region(base, "Sub Gen()\nEnd Sub\n");
        assert_eq!(
            merged,
            "Attribute VB_Name = \"Sheet1\"\r\nPrivate Sub Hand()\r\nEnd Sub\r\n'--- RAS BEGIN ---\r\nSub Gen()\r\nEnd Sub\r\n'--- RAS END ---\r\n"
        );
        assert_eq!(extract_managed_region(&merged).as_deref(), Some("Sub Gen()\r\nEnd Sub"));
        // Idempotent.
        assert_eq!(merge_managed_region(&merged, "Sub Gen()\r\nEnd Sub"), merged);
        // Replace in place, keeping code after the region.
        let with_tail = format!("{merged}' tail\r\n");
        let replaced = merge_managed_region(&with_tail, "X = 1");
        assert_eq!(
            replaced,
            "Attribute VB_Name = \"Sheet1\"\r\nPrivate Sub Hand()\r\nEnd Sub\r\n'--- RAS BEGIN ---\r\nX = 1\r\n'--- RAS END ---\r\n' tail\r\n"
        );
        // Empty code removes the region and restores the original text.
        assert_eq!(merge_managed_region(&merged, ""), base);
        assert_eq!(merge_managed_region(&merged, " \r\n"), base);
        assert_eq!(merge_managed_region(base, ""), base);
        assert_eq!(extract_managed_region(base), None);
    }

    #[test]
    fn merge_edge_cases() {
        assert_eq!(merge_managed_region("", "A"), "'--- RAS BEGIN ---\r\nA\r\n'--- RAS END ---\r\n");
        // LF-only sources stay LF-only.
        assert_eq!(
            merge_managed_region("x\ny", "A\r\nB"),
            "x\ny\n'--- RAS BEGIN ---\nA\nB\n'--- RAS END ---\n"
        );
        // Indented markers are recognised.
        let src = "a\r\n   '--- RAS BEGIN ---  \r\nold\r\n\t'--- RAS END ---\r\nb\r\n";
        assert_eq!(merge_managed_region(src, "new"), "a\r\n'--- RAS BEGIN ---\r\nnew\r\n'--- RAS END ---\r\nb\r\n");
        // Unterminated region: only the marker line is replaced.
        let broken = "a\r\n'--- RAS BEGIN ---\r\nhand written\r\n";
        assert_eq!(extract_managed_region(broken), None);
        assert_eq!(
            merge_managed_region(broken, "new"),
            "a\r\n'--- RAS BEGIN ---\r\nnew\r\n'--- RAS END ---\r\nhand written\r\n"
        );
        // An empty region is extractable.
        assert_eq!(
            extract_managed_region("'--- RAS BEGIN ---\r\n'--- RAS END ---\r\n").as_deref(),
            Some("")
        );
    }
}
