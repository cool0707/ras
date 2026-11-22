//! Source formatter (`ras fmt`, LSP formatting).
//!
//! A conservative, token-based formatter: it re-indents by bracket depth,
//! strips trailing whitespace, collapses runs of blank lines and normalises
//! line endings. Spacing inside a line is left as written.

use crate::kind::{lex, SyntaxKind, SyntaxKind::*};

const INDENT: &str = "    ";

/// Formats `text`. Returns `None` when the file has syntax errors (it is then left alone).
pub fn format(text: &str) -> Option<String> {
    if !crate::parse(text).errors.is_empty() {
        return None;
    }
    let toks = lex(text);
    let mut out = String::with_capacity(text.len());
    let mut depth: usize = 0;
    let mut at_line_start = true;
    let mut blank_run = 0usize;
    // `case` / `default` labels are indented one level less than their bodies.
    let mut switch_depths: Vec<usize> = Vec::new();
    let mut pending_switch = false;
    for (i, t) in toks.iter().enumerate() {
        let s = &text[t.start as usize..t.end as usize];
        if t.kind == WHITESPACE {
            let newlines = s.matches('\n').count();
            if newlines == 0 {
                if !at_line_start {
                    // Keep inner spacing, but never trailing whitespace.
                    let next_is_newline = toks.get(i + 1).is_none();
                    if !next_is_newline {
                        out.push_str(&s.replace('\t', INDENT));
                    }
                }
                continue;
            }
            // strip the spaces we may just have copied before this line break
            while out.ends_with(' ') || out.ends_with('\t') {
                out.pop();
            }
            let keep = if out.is_empty() { 0 } else { newlines.min(2) };
            for _ in 0..keep.saturating_sub(blank_run) {
                out.push('\n');
            }
            blank_run = keep.saturating_sub(1);
            at_line_start = true;
            continue;
        }
        blank_run = 0;
        let closer = matches!(t.kind, R_BRACE | R_PAREN | R_BRACK);
        if closer {
            depth = depth.saturating_sub(1);
            if t.kind == R_BRACE && switch_depths.last() == Some(&depth) {
                switch_depths.pop();
            }
        }
        if at_line_start {
            let mut level = depth;
            let in_switch = switch_depths.last().is_some_and(|d| *d + 1 == depth);
            if in_switch && !matches!(t.kind, CASE_KW | DEFAULT_KW) {
                level += 1;
            }
            for _ in 0..level {
                out.push_str(INDENT);
            }
            at_line_start = false;
        }
        match t.kind {
            BLOCK_COMMENT => {
                // Re-indent the continuation lines of a JSDoc / block comment.
                let level = out
                    .rsplit('\n')
                    .next()
                    .map_or(0, |l| l.len() - l.trim_start().len());
                let mut lines = s.lines();
                out.push_str(lines.next().unwrap_or(""));
                for l in lines {
                    out.push('\n');
                    let body = l.trim_start();
                    if body.starts_with('*') {
                        out.push_str(&" ".repeat(level + 1));
                        out.push_str(body.trim_end());
                    } else {
                        out.push_str(l.trim_end());
                    }
                }
            }
            LINE_COMMENT => out.push_str(s.trim_end()),
            _ => out.push_str(s),
        }
        match t.kind {
            SWITCH_KW => pending_switch = true,
            L_BRACE => {
                if pending_switch {
                    switch_depths.push(depth);
                    pending_switch = false;
                }
                depth += 1;
            }
            L_PAREN | L_BRACK => depth += 1,
            _ => {}
        }
        let _: SyntaxKind = t.kind;
    }
    while out.ends_with([' ', '\t', '\n']) {
        out.pop();
    }
    out.push('\n');
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reindents_and_trims() {
        let src = "function f(x: int): int {   \n  if (x > 0) {\n\t\treturn x;\n}\n\n\n\n      return 0;\n}";
        let want = "function f(x: int): int {\n    if (x > 0) {\n        return x;\n    }\n\n    return 0;\n}\n";
        assert_eq!(format(src).as_deref(), Some(want));
    }

    #[test]
    fn switch_and_comments() {
        let src = "function f(x: int) {\nswitch (x) {\ncase 1:\ng();\nbreak;\ndefault:\nh();\n}\n}\n/**\n   * doc\n */\nfunction g() {}\n";
        let want = "function f(x: int) {\n    switch (x) {\n        case 1:\n            g();\n            break;\n        default:\n            h();\n    }\n}\n/**\n * doc\n */\nfunction g() {}\n";
        assert_eq!(format(src).as_deref(), Some(want));
    }

    #[test]
    fn idempotent_and_refuses_broken_code() {
        let src = "const a = [\n1,\n2,\n];\nlet s = `multi\n  line`;\n";
        let once = format(src).unwrap();
        assert_eq!(format(&once).unwrap(), once);
        assert!(format("function (").is_none());
    }
}
