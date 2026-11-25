//! VBA identifier rules: reserved words, built-in names and sanitising (spec §13.6).

/// VBA keywords and statements that can never be used as identifiers.
/// `Circle`, `PSet` and `Scale` (graphics methods with their own syntax) are a syntax error
/// as a type or variable name in Excel's VBA, although the language reference does not list them.
const RESERVED: &[&str] = &[
    "addressof",
    "and",
    "any",
    "as",
    "attribute",
    "boolean",
    "byref",
    "byte",
    "byval",
    "call",
    "case",
    "cbool",
    "cbyte",
    "ccur",
    "cdate",
    "cdbl",
    "cdec",
    "cint",
    "circle",
    "clng",
    "clnglng",
    "clngptr",
    "close",
    "const",
    "csng",
    "cstr",
    "currency",
    "cvar",
    "cverr",
    "date",
    "debug",
    "decimal",
    "declare",
    "defbool",
    "defbyte",
    "defcur",
    "defdate",
    "defdbl",
    "defint",
    "deflng",
    "defobj",
    "defsng",
    "defstr",
    "defvar",
    "dim",
    "do",
    "doevents",
    "double",
    "each",
    "else",
    "elseif",
    "empty",
    "end",
    "endif",
    "enum",
    "eqv",
    "erase",
    "error",
    "event",
    "exit",
    "false",
    "for",
    "friend",
    "function",
    "get",
    "global",
    "gosub",
    "goto",
    "if",
    "imp",
    "implements",
    "in",
    "input",
    "integer",
    "is",
    "lbound",
    "len",
    "let",
    "like",
    "lock",
    "long",
    "longlong",
    "longptr",
    "loop",
    "lset",
    "me",
    "mid",
    "mod",
    "new",
    "next",
    "not",
    "nothing",
    "null",
    "object",
    "on",
    "open",
    "option",
    "optional",
    "or",
    "paramarray",
    "preserve",
    "print",
    "private",
    "property",
    "pset",
    "ptrsafe",
    "public",
    "put",
    "raiseevent",
    "redim",
    "rem",
    "resume",
    "return",
    "rset",
    "scale",
    "seek",
    "select",
    "set",
    "shared",
    "single",
    "spc",
    "static",
    "stop",
    "string",
    "sub",
    "tab",
    "then",
    "to",
    "true",
    "type",
    "typeof",
    "ubound",
    "unlock",
    "until",
    "variant",
    "wend",
    "while",
    "with",
    "withevents",
    "write",
    "xor",
];

/// Built-in functions, objects and common member names. Using them as user
/// identifiers either shadows the built-in or makes the VBE rewrite the
/// casing of every member access, so they are renamed (`value` → `value_`).
const BUILTIN: &[&str] = &[
    "abs",
    "array",
    "asc",
    "ascw",
    "atn",
    "chr",
    "chrw",
    "cos",
    "curdir",
    "day",
    "dir",
    "environ",
    "eof",
    "erl",
    "err",
    "exp",
    "fix",
    "format",
    "hex",
    "hour",
    "iif",
    "instr",
    "instrrev",
    "int",
    "isarray",
    "isdate",
    "isempty",
    "iserror",
    "ismissing",
    "isnull",
    "isnumeric",
    "isobject",
    "join",
    "lcase",
    "left",
    "log",
    "ltrim",
    "minute",
    "month",
    "msgbox",
    "inputbox",
    "name",
    "now",
    "oct",
    "replace",
    "right",
    "rnd",
    "round",
    "rtrim",
    "second",
    "sgn",
    "shell",
    "sin",
    "space",
    "split",
    "sqr",
    "str",
    "strcomp",
    "strconv",
    "tan",
    "time",
    "timer",
    "trim",
    "typename",
    "ucase",
    "val",
    "vartype",
    "weekday",
    "year",
    "application",
    "range",
    "cells",
    "rows",
    "columns",
    "sheets",
    "worksheets",
    "workbooks",
    "activesheet",
    "activeworkbook",
    "activecell",
    "thisworkbook",
    "selection",
    "worksheetfunction",
    "value",
    "value2",
    "count",
    "item",
    "index",
    "row",
    "column",
    "text",
    "formula",
    "address",
    "font",
    "interior",
    "width",
    "height",
    "visible",
    "caption",
    "parent",
    "offset",
    "resize",
    "areas",
    "collection",
    "dictionary",
    "evaluate",
    "run",
    "union",
    "intersect",
    "calculate",
    "delete",
    "insert",
    "copy",
    "paste",
    "find",
    "sort",
    "clear",
    "add",
    "remove",
    "exists",
    "keys",
    "items",
    "quit",
    "wait",
    "version",
    "path",
    "cursor",
    "names",
    "windows",
    "charts",
    "shapes",
    "hyperlinks",
    "comments",
    "top",
    "color",
    "size",
    "bold",
    "italic",
    "number",
    "source",
    "description",
    "error",
    "raise",
    "print",
    "assert",
    "kill",
    "mkdir",
    "rmdir",
    "filecopy",
    "filelen",
    "freefile",
    "beep",
    "command",
    "choose",
    "switch",
    "partition",
    "rgb",
    "qbcolor",
    "strptr",
    "objptr",
    "varptr",
    "main",
];

/// Members of `IUnknown` / `IDispatch`. A class module that declares one of them does not
/// compile ("このメンバー識別子は既にオブジェクト モジュールの中に存在しています"), and with Excel
/// opening the book without a prompt the procedures that use the class just cannot be run.
const DISPATCH_MEMBERS: &[&str] = &[
    "addref",
    "getidsofnames",
    "gettypeinfo",
    "gettypeinfocount",
    "invoke",
    "queryinterface",
    "release",
];

/// RAS keywords: a host global whose camelCase form is one of them keeps its VBA spelling (`Switch`).
const RAS_KEYWORDS: &[&str] = &[
    "let",
    "const",
    "function",
    "return",
    "if",
    "else",
    "for",
    "while",
    "do",
    "break",
    "continue",
    "switch",
    "case",
    "default",
    "try",
    "catch",
    "finally",
    "throw",
    "new",
    "class",
    "extends",
    "interface",
    "enum",
    "import",
    "export",
    "this",
    "super",
    "null",
    "undefined",
    "true",
    "false",
    "void",
    "typeof",
    "instanceof",
    "in",
    "var",
    "debugger",
    "byref",
    "byval",
    "inline",
    "noinline",
    "resume",
];

pub fn is_ras_keyword(name: &str) -> bool {
    RAS_KEYWORDS.contains(&name)
}

/// The JavaScript-style (camelCase) spelling of a host member or global, which the type
/// libraries spell in PascalCase: `Value` → `value`, `ScreenUpdating` → `screenUpdating`.
/// A leading run of capitals is one word, up to the capital that starts the next word:
/// `VBProject` → `vbProject`, `RGB` → `rgb`, `UBound` → `uBound`, `FormulaR1C1` → `formulaR1C1`.
/// Names that do not start with a capital letter (`xlUp`, `_NewEnum`) are returned unchanged.
pub fn js_name(host: &str) -> String {
    let chars: Vec<char> = host.chars().collect();
    let run = chars.iter().take_while(|c| c.is_ascii_uppercase()).count();
    if run == 0 {
        return host.to_string();
    }
    // `VBProject`: the last capital of the run belongs to the next word when a lowercase letter follows.
    let lower = if run > 1 && chars.get(run).is_some_and(|c| c.is_ascii_lowercase()) {
        run - 1
    } else {
        run
    };
    chars
        .iter()
        .enumerate()
        .map(|(i, c)| {
            if i < lower {
                c.to_ascii_lowercase()
            } else {
                *c
            }
        })
        .collect()
}

pub fn is_reserved(name: &str) -> bool {
    RESERVED
        .binary_search(&name.to_ascii_lowercase().as_str())
        .is_ok()
}

pub fn is_builtin(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    BUILTIN.contains(&lower.as_str())
}

/// Makes a syntactically valid VBA identifier (no leading `_` / digit, no `$`, at most 255 characters).
pub fn valid_ident(name: &str) -> String {
    let mut out: String = name
        .chars()
        .map(|c| if c == '$' { '_' } else { c })
        .collect();
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
/// Reserved words and `IDispatch` member names get a trailing underscore; casing is kept.
pub fn member_ident(name: &str) -> String {
    let base = valid_ident(name);
    if is_reserved(&base) || DISPATCH_MEMBERS.contains(&base.to_ascii_lowercase().as_str()) {
        format!("{base}_")
    } else {
        base
    }
}

/// Functions and statements of the VBA library that are not reserved words. A module-level
/// procedure or variable of such a name hides the library's for the whole project (in its own
/// module when private): `Left$(s, 1)` in `RAS_Runtime` then does not compile ("型宣言文字と
/// 宣言されたデータ型が一致しません"), and Excel stops in the VBA editor when a macro is run.
const VBA_GLOBALS: &[&str] = &[
    "abs",
    "appactivate",
    "array",
    "asc",
    "ascb",
    "ascw",
    "atn",
    "beep",
    "callbyname",
    "chdir",
    "chdrive",
    "choose",
    "chr",
    "chrb",
    "chrw",
    "command",
    "cos",
    "createobject",
    "curdir",
    "dateadd",
    "datediff",
    "datepart",
    "dateserial",
    "datevalue",
    "day",
    "ddb",
    "deletesetting",
    "dir",
    "environ",
    "eof",
    "erl",
    "err",
    "exp",
    "fileattr",
    "filecopy",
    "filedatetime",
    "filelen",
    "filter",
    "fix",
    "format",
    "formatcurrency",
    "formatdatetime",
    "formatnumber",
    "formatpercent",
    "freefile",
    "fv",
    "getallsettings",
    "getattr",
    "getobject",
    "getsetting",
    "hex",
    "hour",
    "iif",
    "imestatus",
    "inputbox",
    "instr",
    "instrb",
    "instrrev",
    "int",
    "ipmt",
    "irr",
    "isarray",
    "isdate",
    "isempty",
    "iserror",
    "ismissing",
    "isnull",
    "isnumeric",
    "isobject",
    "join",
    "kill",
    "lcase",
    "left",
    "leftb",
    "lenb",
    "load",
    "loc",
    "lof",
    "log",
    "ltrim",
    "midb",
    "minute",
    "mirr",
    "mkdir",
    "month",
    "monthname",
    "msgbox",
    "name",
    "now",
    "nper",
    "npv",
    "objptr",
    "oct",
    "partition",
    "pmt",
    "ppmt",
    "pv",
    "qbcolor",
    "randomize",
    "rate",
    "replace",
    "reset",
    "rgb",
    "right",
    "rightb",
    "rmdir",
    "rnd",
    "round",
    "rtrim",
    "savesetting",
    "second",
    "sendkeys",
    "setattr",
    "sgn",
    "shell",
    "sin",
    "sln",
    "space",
    "split",
    "sqr",
    "str",
    "strcomp",
    "strconv",
    "strptr",
    "strreverse",
    "switch",
    "syd",
    "tan",
    "time",
    "timer",
    "timeserial",
    "timevalue",
    "trim",
    "typename",
    "ucase",
    "unload",
    "val",
    "varptr",
    "vartype",
    "weekday",
    "weekdayname",
    "year",
];

/// Identifier for a module-level procedure or variable: as `member_ident`, and a name of the
/// VBA library gets a trailing underscore too (`left` → `left_`), so that it hides nothing.
pub fn module_ident(name: &str) -> String {
    let base = member_ident(name);
    if VBA_GLOBALS.contains(&base.to_ascii_lowercase().as_str()) {
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
        assert_eq!(
            sorted, RESERVED,
            "RESERVED must stay sorted for binary search"
        );
    }

    #[test]
    fn js_names() {
        for (host, js) in [
            ("Value", "value"),
            ("ScreenUpdating", "screenUpdating"),
            ("VBProject", "vbProject"),
            ("RGB", "rgb"),
            ("UBound", "uBound"),
            ("CStr", "cStr"),
            ("FormulaR1C1", "formulaR1C1"),
            ("OLEObjects", "oleObjects"),
            ("ID", "id"),
            ("A1", "a1"),
            ("xlUp", "xlUp"),
            ("_NewEnum", "_NewEnum"),
            ("message", "message"),
        ] {
            assert_eq!(js_name(host), js, "{host}");
        }
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
        assert_eq!(member_ident("Circle"), "Circle_");
        assert_eq!(member_ident("release"), "release_");
        assert_eq!(local_ident("release"), "release");
        assert_eq!(member_ident("isBlank"), "isBlank");
        assert_eq!(module_ident("left"), "left_");
        assert_eq!(module_ident("Trim"), "Trim_");
        assert_eq!(module_ident("type"), "type_");
        assert_eq!(module_ident("run"), "run");
    }
}
