//! Editor support inside the formulas of `` fx`…` `` (spec §2.7): completion of Excel's
//! worksheet functions, their signatures and hover.
//!
//! The functions are those the stand-in for Excel works out (`stdlib/formula_functions.json`,
//! listed from its engine by a test of `ras-workbook`), as the grid completes them.
//! Everything works on the text, so it also works while the formula is being typed.

use std::sync::OnceLock;

use ras_syntax::{lex, SyntaxKind, FORMULA_TAG};

const FUNCTIONS_JSON: &str = include_str!("../stdlib/formula_functions.json");

/// A worksheet function: `SUM`, `SUM(number1, [number2], ...)`, what it does.
pub struct FormulaFn {
    pub name: String,
    pub signature: String,
    pub description: String,
}

/// Every function, by name.
pub fn functions() -> &'static [FormulaFn] {
    static LIST: OnceLock<Vec<FormulaFn>> = OnceLock::new();
    LIST.get_or_init(|| {
        let rows: Vec<(String, String, String)> =
            serde_json::from_str(FUNCTIONS_JSON).unwrap_or_default();
        rows.into_iter()
            .map(|(name, signature, description)| FormulaFn {
                name,
                signature,
                description,
            })
            .collect()
    })
}

/// The function named `name` (any case, with or without `_xlfn.`).
pub fn lookup(name: &str) -> Option<&'static FormulaFn> {
    let up = name.to_ascii_uppercase();
    let up = up.strip_prefix("_XLFN.").unwrap_or(&up);
    functions()
        .binary_search_by(|f| f.name.as_str().cmp(up))
        .ok()
        .map(|i| &functions()[i])
}

/// The formula text of the `` fx`…` `` that `offset` is in: `start` is just after the
/// opening back-tick, `end` at the closing one (the end of the line while it is not written
/// yet). `None` outside a formula and inside its `${…}`.
pub fn formula_at(text: &str, offset: usize) -> Option<(usize, usize)> {
    let toks = lex(text);
    let mut prev: Option<usize> = None;
    for (i, t) in toks.iter().enumerate() {
        if t.kind.is_trivia() {
            continue;
        }
        let (start, end) = (t.start as usize, t.end as usize);
        let opens = text[start..end].starts_with('`')
            && matches!(t.kind, SyntaxKind::TEMPLATE_LIT | SyntaxKind::ERROR_TOKEN);
        if opens && start < offset {
            let tagged = prev.is_some_and(|p| {
                let tag = &toks[p];
                tag.kind == SyntaxKind::IDENT
                    && &text[tag.start as usize..tag.end as usize] == FORMULA_TAG
                    && !text[tag.end as usize..start].contains('\n')
            });
            let end = if t.kind == SyntaxKind::TEMPLATE_LIT {
                end - 1
            } else {
                // Not closed yet: up to the end of the line.
                text[start..].find('\n').map_or(text.len(), |n| start + n)
            };
            if offset <= end {
                if !tagged || in_interpolation(&text[start + 1..offset]) {
                    return None;
                }
                return Some((start + 1, end));
            }
        }
        if start >= offset {
            return None;
        }
        prev = Some(i);
    }
    None
}

/// Whether the end of `body` (the text of a template after its back-tick) is inside `${…}`.
fn in_interpolation(body: &str) -> bool {
    let b = body.as_bytes();
    let mut depth = 0usize;
    let mut i = 0;
    while i < b.len() {
        match b[i] {
            b'\\' => i += 1,
            b'$' if depth == 0 && b.get(i + 1) == Some(&b'{') => {
                depth = 1;
                i += 1;
            }
            b'{' if depth > 0 => depth += 1,
            b'}' if depth > 0 => depth -= 1,
            _ => {}
        }
        i += 1;
    }
    depth > 0
}

/// A call still open before the cursor: the function's name and the argument being written.
struct OpenCall {
    name: String,
    arg: usize,
}

/// What the formula before the cursor is: in a string or not, the word being written (and
/// where it starts) and the calls still open, the innermost last.
struct Scan {
    in_string: bool,
    word: String,
    word_start: usize,
    calls: Vec<OpenCall>,
}

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.'
}

/// Scans `text[start..offset]`; a `${…}` counts as a value.
fn scan(text: &str, start: usize, offset: usize) -> Scan {
    let mut s = Scan {
        in_string: false,
        word: String::new(),
        word_start: start,
        calls: Vec::new(),
    };
    let body = &text[start..offset];
    let mut it = body.char_indices().peekable();
    while let Some((i, ch)) = it.next() {
        if s.in_string {
            if ch == '"' {
                s.in_string = false;
            }
            continue;
        }
        if ch == '\\' {
            it.next();
            s.word.clear();
            continue;
        }
        if ch == '$' && it.peek().is_some_and(|&(_, c)| c == '{') {
            let mut depth = 0usize;
            for (_, c) in it.by_ref() {
                match c {
                    '{' => depth += 1,
                    '}' => {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    }
                    _ => {}
                }
            }
            s.word.clear();
            continue;
        }
        if ch == '"' {
            s.in_string = true;
            s.word.clear();
            continue;
        }
        if is_word_char(ch) {
            if s.word.is_empty() {
                s.word_start = start + i;
            }
            s.word.push(ch);
            continue;
        }
        match ch {
            '(' => s.calls.push(OpenCall {
                name: s.word.to_ascii_uppercase(),
                arg: 0,
            }),
            ')' => {
                s.calls.pop();
            }
            ',' => {
                if let Some(c) = s.calls.last_mut() {
                    c.arg += 1;
                }
            }
            // Array constants: their commas are not arguments.
            '{' => s.calls.push(OpenCall {
                name: "{".into(),
                arg: 0,
            }),
            '}' if s.calls.last().is_some_and(|c| c.name == "{") => {
                s.calls.pop();
            }
            _ => {}
        }
        s.word.clear();
    }
    if s.word.is_empty() {
        s.word_start = offset;
    }
    s
}

/// Whether `w` is a cell (`A1`, `XFD1048576`), which is not completed.
fn is_cell(w: &str) -> bool {
    let letters = w.chars().take_while(|c| c.is_ascii_alphabetic()).count();
    (1..=3).contains(&letters)
        && w.len() > letters
        && w[letters..].chars().all(|c| c.is_ascii_digit())
}

/// A function offered for the word being written.
pub struct FormulaCompletion {
    pub function: &'static FormulaFn,
    /// The text the function's name replaces (the word written so far, with its `.`s).
    pub replace: (usize, usize),
}

/// The functions to offer at `offset`, or `None` when the cursor is not in a formula (an
/// empty list in a formula's string or reference: nothing of RAS is offered there).
pub fn completions(text: &str, offset: usize) -> Option<Vec<FormulaCompletion>> {
    let (start, _) = formula_at(text, offset)?;
    let s = scan(text, start, offset);
    let prev = text[..s.word_start].chars().next_back();
    let after = text[offset..].chars().next();
    let offer = !s.in_string
        && (s.word.is_empty() || s.word.starts_with(|c: char| c.is_alphabetic() || c == '_'))
        && !is_cell(&s.word)
        // a sheet's cell (`Sheet1!A1`) or an absolute reference (`$A$1`)
        && !matches!(prev, Some('!' | '$' | ':' | '\''))
        // not in the middle of a name
        && !after.is_some_and(|c| is_word_char(c) || c == '(');
    if !offer {
        return Some(Vec::new());
    }
    let up = s.word.to_ascii_uppercase();
    Some(
        functions()
            .iter()
            .filter(|f| f.name.starts_with(&up))
            .map(|function| FormulaCompletion {
                function,
                replace: (s.word_start, offset),
            })
            .collect(),
    )
}

/// The signature of the innermost function call open at `offset`: the function, the
/// parameters' ranges in its signature (bytes; the signatures are ASCII) and the one being
/// written. `None` outside a formula; `Some(None)` in a formula outside a known call.
#[allow(clippy::type_complexity)]
pub fn signature_at(
    text: &str,
    offset: usize,
) -> Option<Option<(&'static FormulaFn, Vec<(usize, usize)>, Option<usize>)>> {
    let (start, _) = formula_at(text, offset)?;
    let s = scan(text, start, offset);
    if s.in_string {
        return Some(None);
    }
    let Some(call) = s.calls.iter().rev().find(|c| c.name != "{") else {
        return Some(None);
    };
    let Some(f) = lookup(&call.name) else {
        return Some(None);
    };
    let params = param_ranges(&f.signature);
    let names = |i: usize| &f.signature[params[i].0..params[i].1];
    let named = (0..params.len()).filter(|&i| names(i) != "...").count();
    let repeats = (0..params.len()).any(|i| names(i) == "...");
    // Past the last named parameter of a function that takes more: the last named one.
    let active = if call.arg < named {
        Some(call.arg)
    } else if repeats && named > 0 {
        Some(named - 1)
    } else {
        None
    };
    Some(Some((f, params, active)))
}

/// The ranges of the parameters in `NAME(a, [b], ...)`, split at the commas outside brackets.
fn param_ranges(sig: &str) -> Vec<(usize, usize)> {
    let (Some(open), Some(close)) = (sig.find('('), sig.rfind(')')) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut depth = 0i32;
    let mut from = open + 1;
    let push = |a: usize, b: usize, out: &mut Vec<(usize, usize)>| {
        let part = &sig[a..b];
        let lead = part.len() - part.trim_start().len();
        let a = a + lead;
        let b = a + part.trim().len();
        if b > a {
            out.push((a, b));
        }
    };
    for (i, ch) in sig[open + 1..close].char_indices() {
        let i = open + 1 + i;
        match ch {
            '[' | '(' => depth += 1,
            ']' | ')' => depth -= 1,
            ',' if depth == 0 => {
                push(from, i, &mut out);
                from = i + 1;
            }
            _ => {}
        }
    }
    push(from, close, &mut out);
    out
}

/// The function whose name is at `offset` in a formula, and the name's range. `None` outside
/// a formula; `Some(None)` in a formula elsewhere.
#[allow(clippy::type_complexity)]
pub fn hover_at(text: &str, offset: usize) -> Option<Option<(&'static FormulaFn, (usize, usize))>> {
    let (start, end) = formula_at(text, offset)?;
    if scan(text, start, offset).in_string {
        return Some(None);
    }
    let a = text[start..offset]
        .char_indices()
        .rev()
        .take_while(|&(_, c)| is_word_char(c))
        .last()
        .map_or(offset, |(i, _)| start + i);
    let b = text[offset..end]
        .char_indices()
        .find(|&(_, c)| !is_word_char(c))
        .map_or(end, |(i, _)| offset + i);
    // A function is a name followed by `(` (a defined name or a cell is not).
    if a == b || !text[b..end].trim_start().starts_with('(') {
        return Some(None);
    }
    Some(lookup(&text[a..b]).map(|f| (f, (a, b))))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `src` with the cursor at `|`.
    fn at(src: &str) -> (String, usize) {
        let i = src.find('|').unwrap();
        (src.replacen('|', "", 1), i)
    }

    #[test]
    fn the_catalogue_is_sorted_and_has_the_common_functions() {
        let f = functions();
        assert!(f.len() > 400);
        assert!(f.windows(2).all(|w| w[0].name < w[1].name));
        assert_eq!(lookup("vlookup").unwrap().name, "VLOOKUP");
        assert_eq!(lookup("_xlfn.XLOOKUP").unwrap().name, "XLOOKUP");
        assert!(lookup("IF").is_some() && lookup("SUBTOTAL").is_some());
    }

    #[test]
    fn formula_is_found_only_in_fx_templates() {
        let (t, o) = at("const a = fx`SUM(|A1)`;");
        assert_eq!(formula_at(&t, o), Some((13, 20)));
        let (t, o) = at("const a = ws.fx`SU|`;");
        assert!(formula_at(&t, o).is_some());
        for src in [
            "const a = `SUM(|A1)`;",
            "const a = foo`SUM(|A1)`;",
            "const a = fx`SUM(${n|})`;",
            "const a = fx|`SUM()`;",
            "const a = fx`SUM()`|;",
        ] {
            let (t, o) = at(src);
            assert_eq!(formula_at(&t, o), None, "{src}");
        }
        // Not closed yet: up to the end of the line.
        let (t, o) = at("const a = fx`SUM(A1, VLO|\nconst b = 1;");
        assert!(formula_at(&t, o).is_some());
    }

    #[test]
    fn completes_function_names() {
        let names = |src: &str| -> Vec<String> {
            let (t, o) = at(src);
            completions(&t, o)
                .expect("in a formula")
                .iter()
                .map(|c| c.function.name.clone())
                .collect()
        };
        let got = names("fx`1 + vlo|`");
        assert_eq!(got, ["VLOOKUP"]);
        assert!(names("fx`STDEV.|`").contains(&"STDEV.S".to_string()));
        assert!(names("fx`SUM(A1, |`").len() > 400);
        // Not in strings, cells, sheets' references or the middle of a name.
        assert!(names(r#"fx`COUNTIF(A:A, "su|")`"#).is_empty());
        assert!(names("fx`A1|`").is_empty());
        assert!(names("fx`Sheet1!B|`").is_empty());
        assert!(names("fx`$A|$1`").is_empty());
        assert!(names("fx`SU|M(1)`").is_empty());
        // The word replaced, with its `.`.
        let (t, o) = at("fx`STDEV.|`");
        let c = completions(&t, o).unwrap();
        assert_eq!(c[0].replace, (3, o));
        // Outside a formula: nothing from here.
        let (t, o) = at("const su| = 1;");
        assert!(completions(&t, o).is_none());
    }

    #[test]
    fn signature_of_the_call_being_written() {
        let sig = |src: &str| {
            let (t, o) = at(src);
            signature_at(&t, o)
                .expect("in a formula")
                .map(|(f, p, a)| (f.name.clone(), p.len(), a))
        };
        assert_eq!(
            sig("fx`VLOOKUP(A1, B:C, |`"),
            Some(("VLOOKUP".into(), 4, Some(2)))
        );
        // Commas in strings, array constants and inner calls are not this call's.
        assert_eq!(
            sig(r#"fx`IF(A1=",", SUM({1,2}, 3), |`"#),
            Some(("IF".into(), 3, Some(2)))
        );
        // Past the named parameters of a function that takes more: the last named one.
        assert_eq!(sig("fx`SUM(1, 2, 3, |`"), Some(("SUM".into(), 3, Some(1))));
        assert_eq!(sig("fx`SUM(1) + |`"), None);
        assert_eq!(sig("fx`SUM(${n}, |`"), Some(("SUM".into(), 3, Some(1))));
    }

    #[test]
    fn hover_on_a_function_name() {
        let (t, o) = at("fx`1 + VLOO|KUP(A1, B:C, 2)`");
        let (f, (a, b)) = hover_at(&t, o).unwrap().unwrap();
        assert_eq!(f.name, "VLOOKUP");
        assert_eq!(&t[a..b], "VLOOKUP");
        let (t, o) = at("fx`A|1 + 1`");
        assert!(hover_at(&t, o).unwrap().is_none());
    }

    #[test]
    fn param_ranges_split_at_top_level_commas() {
        let s = "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])";
        let r = param_ranges(s);
        let names: Vec<&str> = r.iter().map(|&(a, b)| &s[a..b]).collect();
        assert_eq!(
            names,
            [
                "lookup_value",
                "table_array",
                "col_index_num",
                "[range_lookup]"
            ]
        );
    }
}
