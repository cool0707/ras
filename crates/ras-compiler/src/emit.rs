//! Final assembly: prints the VBA modules, adds the tree-shaken runtime,
//! enforces code-page and VBA structural limits, and builds source maps.

use std::collections::HashSet;

use crate::diag::{Diagnostic, Severity};
use crate::lower::Lowered;
use crate::model::Program;
use crate::runtime;
use crate::vba::{file_header, print_module, Mapping, PrintOptions, VModule, VModuleKind};

/// Can every character of `s` be stored in the given Windows code page?
pub fn encodable(s: &str, codepage: u16) -> bool {
    if s.is_ascii() {
        return true;
    }
    let enc = match codepage {
        932 => encoding_rs::SHIFT_JIS,
        936 => encoding_rs::GBK,
        949 => encoding_rs::EUC_KR,
        950 => encoding_rs::BIG5,
        874 => encoding_rs::WINDOWS_874,
        1250 => encoding_rs::WINDOWS_1250,
        1251 => encoding_rs::WINDOWS_1251,
        1252 => encoding_rs::WINDOWS_1252,
        1253 => encoding_rs::WINDOWS_1253,
        1254 => encoding_rs::WINDOWS_1254,
        1255 => encoding_rs::WINDOWS_1255,
        1256 => encoding_rs::WINDOWS_1256,
        1257 => encoding_rs::WINDOWS_1257,
        1258 => encoding_rs::WINDOWS_1258,
        _ => return false,
    };
    let (_, _, unmappable) = enc.encode(s);
    !unmappable
}

#[derive(Debug, Clone)]
pub struct OutputFile {
    /// VBA module name.
    pub name: String,
    pub kind: VModuleKind,
    /// `Report.bas`, `Stack__Range.cls`, `ThisWorkbook.doccls`.
    pub file_name: String,
    /// Code-pane text (what the VBE shows), CRLF line endings.
    pub code: String,
    /// `code` preceded by the exported-file header (`Attribute VB_Name` …).
    pub full_text: String,
    pub sources: Vec<String>,
    /// (VBA line in `code`, index into `sources`, RAS line), 1-based.
    pub mappings: Vec<Mapping>,
    /// CRC-32 of the generated code, also recorded in its first line (spec §14.7).
    pub hash: u32,
}

impl OutputFile {
    /// Source map as JSON (`build/*.rasmap`, spec §14.8).
    pub fn rasmap(&self) -> String {
        let mappings: Vec<[u32; 3]> = self.mappings.iter().map(|m| [m.0, m.1, m.2]).collect();
        serde_json::json!({
            "version": 1,
            "module": self.name,
            "file": self.file_name,
            "sources": self.sources,
            "mappings": mappings,
        })
        .to_string()
    }

    /// Module text as stored in a `vbaProject.bin` module stream: `Attribute` lines followed by the code
    /// (the `VERSION 1.0 CLASS` block of exported `.cls` files is not part of the stream).
    pub fn stream_text(&self) -> String {
        match self.full_text.find("Attribute VB_Name") {
            Some(i) => self.full_text[i..].to_string(),
            None => self.code.clone(),
        }
    }

    /// VBA line (1-based, in `code`) generated from the given RAS line, if any.
    pub fn vba_line_for(&self, source: &str, ras_line: u32) -> Option<u32> {
        let si = self.sources.iter().position(|s| s == source)? as u32;
        self.mappings
            .iter()
            .filter(|m| m.1 == si && m.2 >= ras_line)
            .min_by_key(|m| (m.2, m.0))
            .map(|m| m.0)
    }
}

pub struct EmitOptions {
    pub erl: bool,
    pub attributes: bool,
}

const MAX_LINE: usize = 1000;
const MAX_CONTINUATIONS: usize = 24;
const MAX_PROC_BYTES: usize = 60_000;

/// Splits an over-long line at `, ` / ` & ` outside string literals using `_` continuations (spec §13.8).
fn split_long_line(line: &str) -> Option<Vec<String>> {
    if line.len() <= MAX_LINE {
        return Some(vec![line.to_string()]);
    }
    let indent: String = line.chars().take_while(|c| *c == ' ').collect();
    let mut out: Vec<String> = Vec::new();
    let mut rest = line;
    while rest.len() > MAX_LINE {
        // Last split point before the limit that is outside a string literal.
        let mut in_str = false;
        let mut best: Option<usize> = None;
        for (i, ch) in rest.char_indices() {
            if i > MAX_LINE - 10 {
                break;
            }
            if ch == '"' {
                in_str = !in_str;
            } else if !in_str
                && i > indent.len() + 8
                && (rest[i..].starts_with(", ") || rest[i..].starts_with(" & "))
            {
                best = Some(i + if rest[i..].starts_with(", ") { 1 } else { 2 });
            }
        }
        let at = best?;
        out.push(format!("{} _", &rest[..at]));
        rest = &rest[at..];
        if out.len() > MAX_CONTINUATIONS {
            return None;
        }
    }
    let tail = format!("{indent}    {}", rest.trim_start());
    // Continuation lines are indented one extra level.
    for l in out.iter_mut().skip(1) {
        *l = format!("{indent}    {}", l.trim_start());
    }
    out.push(tail);
    Some(out)
}

/// Replaces characters the code page cannot store inside comments by `\u{...}`;
/// returns the offending character when one appears in code (spec §13.7).
fn sanitize_line(line: &str, codepage: u16) -> (String, Option<char>) {
    if encodable(line, codepage) {
        return (line.to_string(), None);
    }
    let mut out = String::with_capacity(line.len());
    let mut in_str = false;
    let mut in_comment = false;
    let mut bad = None;
    for ch in line.chars() {
        if !in_comment {
            if ch == '"' {
                in_str = !in_str;
            } else if ch == '\'' && !in_str {
                in_comment = true;
            }
        }
        if ch.is_ascii() || encodable(ch.encode_utf8(&mut [0; 4]), codepage) {
            out.push(ch);
        } else if in_comment {
            out.push_str(&format!("\\u{{{:X}}}", ch as u32));
        } else {
            bad.get_or_insert(ch);
            out.push('?');
        }
    }
    (out, bad)
}

fn file_name(m: &VModule) -> String {
    match m.kind {
        VModuleKind::Standard => format!("{}.bas", m.name),
        VModuleKind::Class => format!("{}.cls", m.name),
        VModuleKind::Document => format!("{}.doccls", m.name),
        VModuleKind::Form => format!("{}.frmcode", m.name),
    }
}

fn finish(
    m: &VModule,
    code: String,
    mappings: Vec<Mapping>,
    codepage: u16,
    diags: &mut Vec<Diagnostic>,
) -> OutputFile {
    let mut lines: Vec<String> = Vec::new();
    let mut new_mappings: Vec<Mapping> = Vec::new();
    let mut proc_start: Option<(usize, String)> = None;
    let mut proc_bytes = 0usize;
    let error = |diags: &mut Vec<Diagnostic>, code: &'static str, message: String| {
        diags.push(Diagnostic {
            severity: Severity::Error,
            code,
            message,
            module: None,
            span: Default::default(),
        });
    };
    let mut reported_char = false;
    // Forms carry their own marker line (spec §14.11.3).
    let no_hash = matches!(m.kind, VModuleKind::Form);
    let offset = if no_hash { 0 } else { 1 };
    for (i, raw) in code.lines().enumerate() {
        let (line, bad) = sanitize_line(raw, codepage);
        if let (Some(ch), false) = (bad, reported_char) {
            reported_char = true;
            error(diags, "E0901", format!("モジュール `{}`: 識別子に文字 `{ch}`（U+{:04X}）が含まれていますが、コードページ {codepage} では表現できません", m.name, ch as u32));
        }
        let first_out = lines.len() as u32 + 1 + offset;
        for mp in mappings.iter().filter(|mp| mp.0 == i as u32 + 1) {
            new_mappings.push((first_out, mp.1, mp.2));
        }
        match split_long_line(&line) {
            Some(parts) => lines.extend(parts),
            None => {
                error(diags, "E0902", format!("モジュール `{}`: 1 行が長すぎて分割できません（VBA の行継続は 24 個まで）。式を複数の文に分けてください", m.name));
                lines.push(line.clone());
            }
        }
        // 64KB-per-procedure limit (spec §13.8)
        let t = raw.trim_start();
        let is_header = [
            "Public Sub ",
            "Private Sub ",
            "Public Function ",
            "Private Function ",
            "Public Property ",
            "Private Property ",
        ]
        .iter()
        .any(|p| t.starts_with(p));
        if is_header {
            proc_start = Some((i, t.split('(').next().unwrap_or(t).to_string()));
            proc_bytes = 0;
        }
        proc_bytes += raw.len() + 2;
        if t.starts_with("End Sub")
            || t.starts_with("End Function")
            || t.starts_with("End Property")
        {
            if let Some((_, head)) = proc_start.take() {
                if proc_bytes > MAX_PROC_BYTES {
                    error(diags, "E0903", format!("モジュール `{}`: プロシージャ `{head}` が大きすぎます（約 {proc_bytes} バイト。VBA の上限は 64KB）。関数を分割してください", m.name));
                }
            }
        }
    }
    let body: String = lines.iter().map(|l| format!("{l}\r\n")).collect();
    let hash = drift_hash(&body);
    let code = if no_hash {
        body
    } else {
        format!("' RAS-GENERATED hash={hash:08X}\r\n{body}")
    };
    let full_text = format!("{}{}", file_header(m), code);
    OutputFile {
        name: m.name.clone(),
        kind: m.kind.clone(),
        file_name: file_name(m),
        code,
        full_text,
        sources: m.sources.clone(),
        mappings: new_mappings,
        hash,
    }
}

/// Reads the hash recorded in the first line of generated code (`' RAS-GENERATED hash=XXXXXXXX`).
pub fn recorded_hash(code: &str) -> Option<u32> {
    let first = code
        .lines()
        .find(|l| !l.starts_with("Attribute ") && !l.trim().is_empty())?;
    let hex = first.trim().strip_prefix("' RAS-GENERATED hash=")?;
    u32::from_str_radix(hex.trim(), 16).ok()
}

/// True when `code` (code-pane text of a generated module) was edited after generation (drift, spec §14.7).
pub fn has_drifted(code: &str) -> Option<bool> {
    let recorded = recorded_hash(code)?;
    let normalized: String = code
        .replace("\r\n", "\n")
        .lines()
        .skip_while(|l| !l.starts_with("' RAS-GENERATED"))
        .skip(1)
        .map(|l| format!("{l}\r\n"))
        .collect();
    Some(drift_hash(&normalized) != recorded)
}

/// True when two code-pane texts are the same code, ignoring line terminators, trailing blank
/// lines and what the VBE rewrites by itself (see [`drift_hash`]).
pub fn same_code(a: &str, b: &str) -> bool {
    fn normalized(code: &str) -> String {
        let text = code.replace("\r\n", "\n").replace('\r', "\n");
        text.trim_end_matches('\n')
            .split('\n')
            .map(|l| format!("{l}\r\n"))
            .collect()
    }
    a == b || drift_hash(&normalized(a)) == drift_hash(&normalized(b))
}

/// Hash of module code that ignores what the VBE rewrites by itself when a book is saved:
/// the casing of identifiers (unified across the project, e.g. `Err.Source` -> `Err.source`
/// once a class declares `source`) and the spelling of numeric literals (`0.1#` -> `0.1`,
/// `1.0E+20` -> `1E+20`). String literals and comments are hashed as written.
fn drift_hash(code: &str) -> u32 {
    let mut out = String::with_capacity(code.len());
    for line in code.split("\r\n") {
        let chars: Vec<char> = line.chars().collect();
        let mut i = 0;
        let mut in_string = false;
        while i < chars.len() {
            let c = chars[i];
            if in_string {
                out.push(c);
                in_string = c != '"';
                i += 1;
            } else if c == '"' {
                out.push(c);
                in_string = true;
                i += 1;
            } else if c == '\'' {
                out.extend(&chars[i..]);
                break;
            } else if c.is_ascii_digit()
                && !(i > 0 && (chars[i - 1].is_alphanumeric() || chars[i - 1] == '_'))
            {
                let start = i;
                while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
                    i += 1;
                }
                if i < chars.len()
                    && matches!(chars[i], 'E' | 'e' | 'D' | 'd')
                    && chars
                        .get(i + 1)
                        .is_some_and(|n| n.is_ascii_digit() || *n == '+' || *n == '-')
                {
                    i += 2;
                    while i < chars.len() && chars[i].is_ascii_digit() {
                        i += 1;
                    }
                }
                let text: String = chars[start..i]
                    .iter()
                    .collect::<String>()
                    .replace(['D', 'd'], "E");
                if i < chars.len() && matches!(chars[i], '#' | '!' | '@' | '&' | '%' | '^') {
                    i += 1;
                }
                match text.parse::<f64>() {
                    Ok(v) => out.push_str(&v.to_string()),
                    Err(_) => out.push_str(&text),
                }
            } else {
                out.push(c.to_ascii_lowercase());
                i += 1;
            }
        }
        out.push('\n');
    }
    crc32fast::hash(out.as_bytes())
}

/// A standard module with nothing in it: its RAS files hold only comments, blank lines,
/// `import`s or type declarations. It is not written into the workbook (spec §12.1).
pub fn is_empty_module(m: &VModule) -> bool {
    m.kind == VModuleKind::Standard && m.decls.is_empty() && m.procs.is_empty()
}

pub fn finalize(
    prog: &Program,
    lowered: &Lowered,
    opts: &EmitOptions,
) -> (Vec<OutputFile>, Vec<Diagnostic>) {
    let cfg = &prog.config;
    let popts = PrintOptions {
        erl: opts.erl,
        attributes: opts.attributes,
        vba6: cfg.target.vba6,
    };
    let mut diags = Vec::new();
    let mut files = Vec::new();
    let mut used: HashSet<String> = HashSet::new();
    for m in lowered.modules.iter().filter(|m| !is_empty_module(m)) {
        let printed = print_module(m, &popts);
        for id in runtime::identifiers(&printed.text) {
            used.insert(id.to_string());
        }
        files.push(finish(
            m,
            printed.text,
            printed.mappings,
            cfg.target.codepage,
            &mut diags,
        ));
    }
    // "Run in Excel" of a form goes through the runtime (`Application.Run`).
    if lowered.modules.iter().any(|m| m.kind == VModuleKind::Form) {
        used.insert("RAS_ShowForm".into());
    }
    let has_glue = lowered.modules.iter().any(|m| m.name == "RAS_Glue");
    let rt = runtime::build(&used, cfg, has_glue);
    if !rt.module.is_empty() {
        let mut m = VModule::new("RAS_Runtime", VModuleKind::Standard);
        m.private_module = true;
        let text = format!(
            "Option Explicit\r\nOption Compare Binary\r\nOption Private Module\r\n\r\n{}",
            rt.module.replace('\n', "\r\n")
        );
        files.push(finish(
            &m,
            text,
            Vec::new(),
            cfg.target.codepage,
            &mut diags,
        ));
    }
    if let Some(cls) = rt.error_class {
        let m = VModule::new("RAS_Error", VModuleKind::Class);
        let text = format!("Option Explicit\r\n\r\n{}", cls.replace('\n', "\r\n"));
        files.push(finish(
            &m,
            text,
            Vec::new(),
            cfg.target.codepage,
            &mut diags,
        ));
    }
    files.sort_by(|a, b| a.file_name.cmp(&b.file_name));
    (files, diags)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn code_page_checks() {
        assert!(encodable("日本語 abc", 932));
        assert!(!encodable("😀", 932));
        assert!(!encodable("日本語", 1252));
        assert!(encodable("é", 1252));
    }

    #[test]
    fn comments_are_escaped_and_code_is_flagged() {
        let (l, bad) = sanitize_line("    x = 1& ' メモ 😀", 932);
        assert_eq!(l, "    x = 1& ' メモ \\u{1F600}");
        assert!(bad.is_none());
        let (_, bad) = sanitize_line("    😀 = 1&", 932);
        assert_eq!(bad, Some('😀'));
    }

    #[test]
    fn long_lines_are_continued() {
        let args: Vec<String> = (0..300).map(|i| format!("arg{i}")).collect();
        let line = format!("    Call f({})", args.join(", "));
        let parts = split_long_line(&line).unwrap();
        assert!(parts.len() > 1);
        assert!(parts.iter().all(|p| p.len() <= MAX_LINE + 2));
        assert!(parts[..parts.len() - 1].iter().all(|p| p.ends_with(" _")));
        let rejoined: String = parts
            .iter()
            .map(|p| p.trim().trim_end_matches(" _").to_string())
            .collect::<Vec<_>>()
            .join(" ");
        assert_eq!(rejoined.replace("  ", " "), line.trim());
    }

    #[test]
    fn drift_detection() {
        let body = "Option Explicit\r\nSub A()\r\nEnd Sub\r\n";
        let code = format!("' RAS-GENERATED hash={:08X}\r\n{body}", drift_hash(body));
        assert_eq!(has_drifted(&code), Some(false));
        assert_eq!(has_drifted(&code.replace("Sub A", "Sub B")), Some(true));
        assert_eq!(has_drifted("Sub A()\r\nEnd Sub\r\n"), None);
    }

    #[test]
    fn same_code_ignores_line_ends_and_vbe_rewrites() {
        assert!(same_code(
            "Sub A()\r\n    x = 0.1#\r\nEnd Sub\r\n",
            "Sub A()\n    X = 0.1\nEnd Sub"
        ));
        assert!(same_code("", "\r\n"));
        assert!(!same_code(
            "Sub A()\r\nEnd Sub\r\n",
            "Sub B()\r\nEnd Sub\r\n"
        ));
    }

    #[test]
    fn drift_ignores_what_the_vbe_rewrites_on_save() {
        // Generated text, and the same code as Excel stores it after a save without edits.
        let body = "Private Const R As Double = 0.1#\r\nSub A()\r\n    Err.Raise Err.Number, Err.Source, \"Msg\"\r\n    x = RAS_NumStr(0.1# + 0.2#) & RAS_NumStr(1.0E+20) & 7# & (v And &H3FFFFFFF)\r\nEnd Sub\r\n";
        let saved = "Private Const R As Double = 0.1\r\nSub A()\r\n    Err.Raise Err.Number, Err.source, \"Msg\"\r\n    x = RAS_NumStr(0.1 + 0.2) & RAS_NumStr(1E+20) & 7# & (v And &H3FFFFFFF)\r\nEnd Sub\r\n";
        let code = format!("' RAS-GENERATED hash={:08X}\r\n{saved}", drift_hash(body));
        assert_eq!(has_drifted(&code), Some(false));
        // Real edits are still found: values, string contents, comments.
        assert_eq!(has_drifted(&code.replace("0.1\r\n", "0.2\r\n")), Some(true));
        assert_eq!(has_drifted(&code.replace("\"Msg\"", "\"msg\"")), Some(true));
        assert_eq!(
            has_drifted(&code.replace("Sub A()\r\n", "Sub A() ' note\r\n")),
            Some(true)
        );
    }
}
