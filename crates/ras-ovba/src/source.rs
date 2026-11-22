//! The code of document modules as RAS writes it: the whole code after the
//! `Attribute` header ([`replace_document_code`]), and the RAS-managed region
//! between marker lines. The module text itself (line terminators, the
//! `Attribute` header) is [`umya_spreadsheet::vba::source`].

use umya_spreadsheet::vba::source::split_lines;

/// First line of the RAS-managed region inside a document module.
pub const RAS_BEGIN_MARKER: &str = "'--- RAS BEGIN ---";
/// Last line of the RAS-managed region inside a document module.
pub const RAS_END_MARKER: &str = "'--- RAS END ---";

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
pub fn procedure_names(code: &str) -> Vec<String> {
    let mut names = Vec::new();
    for line in split_lines(code) {
        let mut words = line.split_whitespace().peekable();
        while words.peek().is_some_and(|w| {
            ["public", "private", "friend", "static"]
                .iter()
                .any(|m| w.eq_ignore_ascii_case(m))
        }) {
            words.next();
        }
        let Some(keyword) = words.next() else {
            continue;
        };
        if keyword.eq_ignore_ascii_case("property") {
            let accessor = words.next().unwrap_or("");
            if !["get", "let", "set"]
                .iter()
                .any(|a| accessor.eq_ignore_ascii_case(a))
            {
                continue;
            }
        } else if !keyword.eq_ignore_ascii_case("sub") && !keyword.eq_ignore_ascii_case("function")
        {
            continue;
        }
        let name = words.next().unwrap_or("").split('(').next().unwrap_or("");
        if !name.is_empty() {
            names.push(name.to_string());
        }
    }
    names
}

/// The source of a document module with all of its code replaced by `code`
/// (RAS specification §12.2): the leading `Attribute` lines of
/// `existing_source` stay, everything after them goes. The result uses the
/// line terminator of `existing_source` (CR LF when it has none).
pub fn replace_document_code(existing_source: &str, code: &str) -> String {
    let nl = newline_style(existing_source);
    let mut out = String::with_capacity(existing_source.len() + code.len());
    for line in split_lines(existing_source)
        .into_iter()
        .take_while(|l| l.starts_with("Attribute ") || l.starts_with("VERSION "))
    {
        out.push_str(line);
        out.push_str(nl);
    }
    let mut lines: Vec<&str> = split_lines(code);
    while lines.last().is_some_and(|l| l.trim().is_empty()) {
        lines.pop();
    }
    for line in lines {
        out.push_str(line);
        out.push_str(nl);
    }
    out
}

/// The source of a document module without its hand-written code: the
/// `Attribute` header, the `Option` statements and the RAS-managed region
/// are kept, everything else goes. Used when the hand-written code was
/// taken over into the project (converted to RAS).
pub fn strip_hand_written(source: &str) -> String {
    let nl = newline_style(source);
    let lines = split_lines(source);
    let region = find_region(&lines);
    let mut out = String::with_capacity(source.len());
    for (i, line) in lines.iter().enumerate() {
        let in_region = match region {
            Some((begin, Some(end))) => i >= begin && i <= end,
            Some((begin, None)) => i >= begin,
            None => false,
        };
        let trimmed = line.trim_start().to_ascii_lowercase();
        if in_region
            || line.starts_with("Attribute ")
            || line.starts_with("VERSION ")
            || trimmed.starts_with("option ")
        {
            out.push_str(line);
            out.push_str(nl);
        }
    }
    out
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
    fn hand_written_code_is_stripped() {
        let source = "Attribute VB_Name = \"Sheet1\"\r\nOption Explicit\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n'--- RAS BEGIN ---\r\nPrivate Sub Worksheet_Activate()\r\nEnd Sub\r\n'--- RAS END ---\r\nSub Tail()\r\nEnd Sub\r\n";
        assert_eq!(
            strip_hand_written(source),
            "Attribute VB_Name = \"Sheet1\"\r\nOption Explicit\r\n'--- RAS BEGIN ---\r\nPrivate Sub Worksheet_Activate()\r\nEnd Sub\r\n'--- RAS END ---\r\n"
        );
        assert_eq!(strip_hand_written("Sub A()\nEnd Sub\n"), "");
    }

    #[test]
    fn procedures_are_found() {
        let code = "Option Explicit\r\n' Sub Commented()\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n  public static function F(x) As Long\r\nEnd Function\r\nFriend Property Get P() As Long\r\nEnd Property\r\nProperty Let P(v As Long)\r\nEnd Property\r\nSub NoParens\r\nExit Sub\r\nDim Subtotal As Long\r\n";
        assert_eq!(
            procedure_names(code),
            ["Workbook_Open", "F", "P", "P", "NoParens"]
        );
    }

    #[test]
    fn merge_appends_replaces_and_removes() {
        let base = "Attribute VB_Name = \"Sheet1\"\r\nPrivate Sub Hand()\r\nEnd Sub\r\n";
        let merged = merge_managed_region(base, "Sub Gen()\nEnd Sub\n");
        assert_eq!(
            merged,
            "Attribute VB_Name = \"Sheet1\"\r\nPrivate Sub Hand()\r\nEnd Sub\r\n'--- RAS BEGIN ---\r\nSub Gen()\r\nEnd Sub\r\n'--- RAS END ---\r\n"
        );
        assert_eq!(
            extract_managed_region(&merged).as_deref(),
            Some("Sub Gen()\r\nEnd Sub")
        );
        // Idempotent.
        assert_eq!(
            merge_managed_region(&merged, "Sub Gen()\r\nEnd Sub"),
            merged
        );
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
        assert_eq!(
            merge_managed_region("", "A"),
            "'--- RAS BEGIN ---\r\nA\r\n'--- RAS END ---\r\n"
        );
        // LF-only sources stay LF-only.
        assert_eq!(
            merge_managed_region("x\ny", "A\r\nB"),
            "x\ny\n'--- RAS BEGIN ---\nA\nB\n'--- RAS END ---\n"
        );
        // Indented markers are recognised.
        let src = "a\r\n   '--- RAS BEGIN ---  \r\nold\r\n\t'--- RAS END ---\r\nb\r\n";
        assert_eq!(
            merge_managed_region(src, "new"),
            "a\r\n'--- RAS BEGIN ---\r\nnew\r\n'--- RAS END ---\r\nb\r\n"
        );
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
