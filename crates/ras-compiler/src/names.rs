//! VBA identifier rules: reserved words, built-in names and sanitising (spec §13.6).

/// VBA keywords and statements that can never be used as identifiers.
const RESERVED: &[&str] = &[
    "addressof", "and", "any", "as", "attribute", "boolean", "byref", "byte", "byval",
    "call", "case", "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng", "clnglng", "clngptr", "close", "const", "csng", "cstr", "currency", "cvar", "cverr", "date", "debug", "decimal", "declare", "defbool", "defbyte", "defcur",
    "defdate", "defdbl", "defint", "deflng", "defobj", "defsng", "defstr", "defvar", "dim", "do", "doevents", "double", "each",
    "else", "elseif", "empty", "end", "endif", "enum", "eqv", "erase", "error", "event", "exit", "false", "for",
    "friend", "function", "get", "global", "gosub", "goto", "if", "imp", "implements", "in", "input", "integer", "is", "lbound",
    "len", "let", "like", "lock", "long", "longlong", "longptr", "loop", "lset", "me", "mid", "mod", "new", "next", "not", "nothing", "null", "object", "on", "open", "option", "optional", "or", "paramarray",
    "preserve", "print", "private", "property", "ptrsafe", "public", "put", "raiseevent", "redim", "rem",
    "resume", "return", "rset", "seek", "select", "set", "shared", "single", "spc", "static", "stop", "string", "sub",
    "tab", "then", "to", "true", "type", "typeof", "ubound", "unlock", "until", "variant", "wend", "while", "with",
    "withevents", "write", "xor",
];

/// Built-in functions, objects and common member names. Using them as user
/// identifiers either shadows the built-in or makes the VBE rewrite the
/// casing of every member access, so they are renamed (`value` → `value_`).
const BUILTIN: &[&str] = &[
    "abs", "array", "asc", "ascw", "atn", "chr", "chrw", "cos", "curdir", "day", "dir", "environ", "eof", "erl", "err", "exp",
    "fix", "format", "hex", "hour", "iif", "instr", "instrrev", "int", "isarray", "isdate", "isempty", "iserror", "ismissing",
    "isnull", "isnumeric", "isobject", "join", "lcase", "left", "log", "ltrim", "minute", "month", "msgbox", "inputbox", "name",
    "now", "oct", "replace", "right", "rnd", "round", "rtrim", "second", "sgn", "shell", "sin", "space", "split", "sqr", "str",
    "strcomp", "strconv", "tan", "time", "timer", "trim", "typename", "ucase", "val", "vartype", "weekday", "year", "application",
    "range", "cells", "rows", "columns", "sheets", "worksheets", "workbooks", "activesheet", "activeworkbook", "activecell",
    "thisworkbook", "selection", "worksheetfunction", "value", "value2", "count", "item", "index", "row", "column", "text",
    "formula", "address", "font", "interior", "width", "height", "visible", "caption", "parent", "offset", "resize", "areas",
    "collection", "dictionary", "evaluate", "run", "union", "intersect", "calculate", "delete", "insert", "copy", "paste",
    "find", "sort", "clear", "add", "remove", "exists", "keys", "items", "quit", "wait", "version", "path", "cursor", "names",
    "windows", "charts", "shapes", "hyperlinks", "comments", "top", "color", "size", "bold", "italic", "number", "source",
    "description", "error", "raise", "print", "assert", "kill", "mkdir", "rmdir", "filecopy", "filelen", "freefile", "beep", "command",
    "choose", "switch", "partition", "rgb", "qbcolor", "strptr", "objptr", "varptr", "main",
];

pub fn is_reserved(name: &str) -> bool {
    RESERVED.binary_search(&name.to_ascii_lowercase().as_str()).is_ok()
}

pub fn is_builtin(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    BUILTIN.contains(&lower.as_str())
}

/// Makes a syntactically valid VBA identifier (no leading `_` / digit, no `$`, at most 255 characters).
pub fn valid_ident(name: &str) -> String {
    let mut out: String = name.chars().map(|c| if c == '$' { '_' } else { c }).collect();
    if out.is_empty() {
        out.push_str("ras_v");
    }
    let first = out.chars().next().unwrap();
    if first == '_' || first.is_ascii_digit() {
        out.insert(0, 'u');
    }
    if out.chars().count() > 200 {
        out = out.chars().take(200).collect();
    }
    out
}

/// Identifier for a procedure / member / module-level name.
/// Reserved words get a trailing underscore; casing is kept.
pub fn member_ident(name: &str) -> String {
    let base = valid_ident(name);
    if is_reserved(&base) {
        format!("{base}_")
    } else {
        base
    }
}

/// Identifier for a local variable: reserved words and built-in names are both avoided.
pub fn local_ident(name: &str) -> String {
    let base = valid_ident(name);
    if is_reserved(&base) || (is_builtin(&base) && base != "main") {
        format!("{base}_")
    } else {
        base
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reserved_is_sorted() {
        let mut sorted = RESERVED.to_vec();
        sorted.sort_unstable();
        assert_eq!(sorted, RESERVED, "RESERVED must stay sorted for binary search");
    }

    #[test]
    fn renames() {
        assert_eq!(local_ident("value"), "value_");
        assert_eq!(local_ident("Date"), "Date_");
        assert_eq!(local_ident("total"), "total");
        assert_eq!(local_ident("_"), "u_");
        assert_eq!(local_ident("$x"), "u_x");
        assert_eq!(member_ident("count"), "count");
        assert_eq!(member_ident("type"), "type_");
        assert_eq!(member_ident("isBlank"), "isBlank");
    }
}
