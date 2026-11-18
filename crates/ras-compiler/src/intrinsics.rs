//! Built-in members and functions implemented directly by the compiler
//! (array / string / Map methods, `Math`, application-state API …).

use crate::model::{ClassKind, Program, Type};

macro_rules! intrinsics {
    ($($name:ident),* $(,)?) => {
        #[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
        pub enum Intr { $($name),* }
    };
}

intrinsics! {
    // arrays
    ArrLength, ArrRows, ArrCols, ArrPush, ArrPop, ArrMap, ArrFilter, ArrForEach, ArrReduce, ArrSome, ArrEvery,
    ArrFind, ArrFindIndex, ArrIndexOf, ArrIncludes, ArrJoin, ArrSlice, ArrSort, ArrReverse, ArrConcat, ArrClear,
    // strings
    StrLength, StrTrim, StrTrimStart, StrTrimEnd, StrUpper, StrLower, StrSubstring, StrSubstr, StrSlice,
    StrIndexOf, StrLastIndexOf, StrIncludes, StrStartsWith, StrEndsWith, StrPadStart, StrPadEnd, StrRepeat,
    StrReplace, StrReplaceAll, StrSplit, StrCharAt, StrCharCodeAt, StrToString,
    // numbers / any
    NumToString, NumToFixed, BoolToString, DateToString,
    // Map<K, V>
    MapGet, MapSet, MapHas, MapDelete, MapSize, MapKeys, MapValues, MapClear,
    // Math
    MathAbs, MathFloor, MathCeil, MathRound, MathBankersRound, MathTrunc, MathSqrt, MathPow, MathMin, MathMax,
    MathSin, MathCos, MathTan, MathAtan, MathLog, MathExp, MathRandom, MathSign, MathPi,
    // conversion / misc globals
    ParseInt, ParseFloat, ToStr, ConsoleLog, IsNullish,
    // application state (spec §10)
    FastMode, SuspendScreen, ResumeScreen, SuspendEvents, ResumeEvents, SuspendCalc, ResumeCalc,
    SuspendAlerts, ResumeAlerts, SuspendAll, ResumeAll, Status, Sleep, WaitUntil, YieldNow,
    // platform
    PlatformIs64, PlatformIsVba7,
    // Range helpers (spec §9.4)
    RangeToArray, RangeFromArray, RangeToColumn, RangeToRow, SheetUsedTable, ExcelMatch,
}

pub fn array_member(name: &str, dims: usize) -> Option<Intr> {
    use Intr::*;
    if dims == 2 {
        return Some(match name {
            "length" | "rows" => ArrRows,
            "cols" => ArrCols,
            _ => return None,
        });
    }
    Some(match name {
        "length" => ArrLength,
        "push" => ArrPush,
        "pop" => ArrPop,
        "map" => ArrMap,
        "filter" => ArrFilter,
        "forEach" => ArrForEach,
        "reduce" => ArrReduce,
        "some" => ArrSome,
        "every" => ArrEvery,
        "find" => ArrFind,
        "findIndex" => ArrFindIndex,
        "indexOf" => ArrIndexOf,
        "includes" => ArrIncludes,
        "join" => ArrJoin,
        "slice" => ArrSlice,
        "sort" => ArrSort,
        "reverse" => ArrReverse,
        "concat" => ArrConcat,
        "clear" => ArrClear,
        _ => return None,
    })
}

pub const ARRAY_MEMBERS: &[&str] = &[
    "length", "push", "pop", "map", "filter", "forEach", "reduce", "some", "every", "find", "findIndex", "indexOf",
    "includes", "join", "slice", "sort", "reverse", "concat", "clear",
];

pub fn string_member(name: &str) -> Option<Intr> {
    use Intr::*;
    Some(match name {
        "length" => StrLength,
        "trim" => StrTrim,
        "trimStart" | "trimLeft" => StrTrimStart,
        "trimEnd" | "trimRight" => StrTrimEnd,
        "toUpperCase" => StrUpper,
        "toLowerCase" => StrLower,
        "substring" => StrSubstring,
        "substr" => StrSubstr,
        "slice" => StrSlice,
        "indexOf" => StrIndexOf,
        "lastIndexOf" => StrLastIndexOf,
        "includes" => StrIncludes,
        "startsWith" => StrStartsWith,
        "endsWith" => StrEndsWith,
        "padStart" | "padLeft" => StrPadStart,
        "padEnd" | "padRight" => StrPadEnd,
        "repeat" => StrRepeat,
        "replace" => StrReplace,
        "replaceAll" => StrReplaceAll,
        "split" => StrSplit,
        "charAt" => StrCharAt,
        "charCodeAt" => StrCharCodeAt,
        "toString" => StrToString,
        _ => return None,
    })
}

pub const STRING_MEMBERS: &[&str] = &[
    "length", "trim", "trimStart", "trimEnd", "toUpperCase", "toLowerCase", "substring", "substr", "slice", "indexOf",
    "lastIndexOf", "includes", "startsWith", "endsWith", "padStart", "padEnd", "padLeft", "padRight", "repeat",
    "replace", "replaceAll", "split", "charAt", "charCodeAt", "toString",
];

pub fn map_member(name: &str) -> Option<Intr> {
    use Intr::*;
    Some(match name {
        "get" => MapGet,
        "set" => MapSet,
        "has" => MapHas,
        "delete" => MapDelete,
        "size" => MapSize,
        "keys" => MapKeys,
        "values" => MapValues,
        "clear" => MapClear,
        _ => return None,
    })
}

pub const MAP_MEMBERS: &[&str] = &["get", "set", "has", "delete", "size", "keys", "values", "clear"];

pub fn math_member(name: &str) -> Option<Intr> {
    use Intr::*;
    Some(match name {
        "abs" => MathAbs,
        "floor" => MathFloor,
        "ceil" => MathCeil,
        "round" => MathRound,
        "bankersRound" => MathBankersRound,
        "trunc" => MathTrunc,
        "sqrt" => MathSqrt,
        "pow" => MathPow,
        "min" => MathMin,
        "max" => MathMax,
        "sin" => MathSin,
        "cos" => MathCos,
        "tan" => MathTan,
        "atan" => MathAtan,
        "log" => MathLog,
        "exp" => MathExp,
        "random" => MathRandom,
        "sign" => MathSign,
        "PI" => MathPi,
        _ => return None,
    })
}

pub const MATH_MEMBERS: &[&str] = &[
    "abs", "floor", "ceil", "round", "bankersRound", "trunc", "sqrt", "pow", "min", "max", "sin", "cos", "tan", "atan",
    "log", "exp", "random", "sign", "PI",
];

/// Built-in global functions (resolved after user and host names).
pub fn global_fn(name: &str) -> Option<Intr> {
    use Intr::*;
    Some(match name {
        "parseInt" => ParseInt,
        "parseFloat" => ParseFloat,
        "fastMode" => FastMode,
        "suspendScreen" => SuspendScreen,
        "resumeScreen" => ResumeScreen,
        "suspendEvents" => SuspendEvents,
        "resumeEvents" => ResumeEvents,
        "suspendCalc" => SuspendCalc,
        "resumeCalc" => ResumeCalc,
        "suspendAlerts" => SuspendAlerts,
        "resumeAlerts" => ResumeAlerts,
        "suspendAll" => SuspendAll,
        "resumeAll" => ResumeAll,
        "status" => Status,
        "sleep" => Sleep,
        "waitUntil" => WaitUntil,
        "yieldNow" => YieldNow,
        _ => return None,
    })
}

pub const GLOBAL_FNS: &[&str] = &[
    "parseInt", "parseFloat", "fastMode", "suspendScreen", "resumeScreen", "suspendEvents", "resumeEvents",
    "suspendCalc", "resumeCalc", "suspendAlerts", "resumeAlerts", "suspendAll", "resumeAll", "status", "sleep",
    "waitUntil", "yieldNow",
];

/// Built-in namespaces usable as `Name.member`.
pub const BUILTIN_NAMESPACES: &[&str] = &["Math", "console", "Platform"];

pub fn builtin_ns_member(ns: &str, name: &str) -> Option<Intr> {
    match ns {
        "Math" => math_member(name),
        "console" => matches!(name, "log" | "debug" | "info" | "warn" | "error").then_some(Intr::ConsoleLog),
        "Platform" => match name {
            "is64bit" => Some(Intr::PlatformIs64),
            "isVBA7" => Some(Intr::PlatformIsVba7),
            _ => None,
        },
        _ => None,
    }
}

/// Members the compiler adds to host types (`Excel.Range.toArray()` …).
pub fn host_member(prog: &Program, recv: &Type, name: &str) -> Option<Intr> {
    let Type::Class(id, _) = recv else { return None };
    let c = &prog.classes[*id];
    if c.kind != ClassKind::Host || c.namespace.as_deref() != Some("Excel") {
        return None;
    }
    match (c.name.as_str(), name) {
        ("Range", "toArray") => Some(Intr::RangeToArray),
        ("Range", "fromArray") => Some(Intr::RangeFromArray),
        ("Range", "toColumn") => Some(Intr::RangeToColumn),
        ("Range", "toRow") => Some(Intr::RangeToRow),
        ("Worksheet", "usedTable") => Some(Intr::SheetUsedTable),
        _ => None,
    }
}

impl Intr {
    /// True for members read without a call (`arr.length`, `Math.PI`).
    pub fn is_property(self) -> bool {
        use Intr::*;
        matches!(self, ArrLength | ArrRows | ArrCols | StrLength | MapSize | MathPi | PlatformIs64 | PlatformIsVba7)
    }

    /// Array methods whose callback can be fused into a loop (spec §4.1).
    pub fn takes_lambda(self) -> bool {
        use Intr::*;
        matches!(self, ArrMap | ArrFilter | ArrForEach | ArrReduce | ArrSome | ArrEvery | ArrFind | ArrFindIndex | ArrSort)
    }
}
